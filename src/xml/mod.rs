//! Lossless XML tree: a mutable, namespace-aware DOM over quick-xml events.
//!
//! Fidelity contract: parsing a part and serializing it back is *semantically*
//! lossless. Every element, attribute, namespace declaration, comment, processing
//! instruction, CDATA section, and text node is preserved, in document order, along
//! with the XML declaration and any miscellaneous content surrounding the root
//! element. Byte-identical output is not promised — attribute quoting and empty-element
//! form may normalize — but a parse → serialize → parse cycle never changes meaning.
//!
//! # Design
//!
//! Nodes live in an arena ([`Vec`] inside the tree); a [`NodeId`] is a `Copy` index
//! into it. There are no `Rc`/`RefCell` cycles, so the tree is [`Send`] and cheap to
//! move between threads for batch processing.
//!
//! Names are stored as the qualified name exactly as written (`w:p`), so Word's
//! conventional prefixes round-trip. Attribute values are stored *unescaped* (the
//! semantic value) and re-escaped on serialize; `xmlns` / `xmlns:*` declarations are
//! ordinary attributes.

use quick_xml::Reader;
use quick_xml::events::{BytesRef, Event};

use crate::error::{Error, Result};

/// Index of a node in an [`XmlTree`]'s arena. Cheap to copy; stable for the life of
/// the tree (removed nodes stay in the arena, so ids are never reused or invalidated).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct NodeId(u32);

/// The kind of a node: an element or one of the leaf content types.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum NodeKind {
    /// An element, e.g. `<w:p …>…</w:p>`.
    Element(Element),
    /// A text node, stored unescaped.
    Text(String),
    /// A `<![CDATA[…]]>` section, stored with its literal contents.
    CData(String),
    /// A `<!-- … -->` comment, stored with its literal contents.
    Comment(String),
    /// A `<?target data?>` processing instruction, stored with its literal contents
    /// (target and data together, exactly as written).
    Pi(String),
}

/// An element: a qualified name, attributes in document order, and child nodes.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Element {
    /// Qualified name exactly as written in the source, e.g. `"w:p"`.
    name: String,
    /// Attributes in document order. Names are as written (including `xmlns` and
    /// `xmlns:*` declarations); values are stored unescaped and re-escaped on serialize.
    attrs: Vec<(String, String)>,
    /// Child node ids in document order.
    children: Vec<NodeId>,
}

/// A single arena node: its parent link and its content.
#[derive(Debug, Clone)]
struct Node {
    parent: Option<NodeId>,
    kind: NodeKind,
}

/// A parsed XML document as a mutable tree.
///
/// Holds the arena of nodes, the optional XML declaration, the root element, and any
/// miscellaneous nodes (comments, PIs, whitespace) that appear before or after the
/// root, in order.
#[derive(Debug, Clone)]
pub struct XmlTree {
    nodes: Vec<Node>,
    /// Raw declaration content between `<?` and `?>` (e.g. `xml version="1.0" …`), if any.
    decl: Option<String>,
    /// Miscellaneous nodes before the root element, in order.
    prolog: Vec<NodeId>,
    /// The root element.
    root: NodeId,
    /// Miscellaneous nodes after the root element, in order.
    epilog: Vec<NodeId>,
}

// Compile-time guarantee that the tree is `Send` (see the concurrency design): the
// arena carries no `Rc`/`RefCell`, so cross-thread batch processing is sound.
const _: () = {
    const fn assert_send<T: Send>() {}
    let _ = assert_send::<XmlTree>;
};

impl XmlTree {
    /// Parse XML bytes into a tree.
    ///
    /// Preserves the XML declaration, comments / PIs / whitespace on either side of the
    /// root element, CDATA sections, and all text verbatim (the reader is not configured
    /// to trim). Returns [`Error::Malformed`] if there is no root element or more than
    /// one, and [`Error::Xml`] on malformed markup.
    pub fn parse(bytes: &[u8]) -> Result<XmlTree> {
        let mut reader = Reader::from_reader(bytes);
        // Detect mismatched end tags so structurally broken XML is rejected. Text is
        // left untrimmed so whitespace-only nodes survive the round trip.
        reader.config_mut().check_end_names = true;

        let mut nodes: Vec<Node> = Vec::new();
        let mut decl: Option<String> = None;
        let mut prolog: Vec<NodeId> = Vec::new();
        let mut epilog: Vec<NodeId> = Vec::new();
        let mut root: Option<NodeId> = None;
        // Open elements, innermost last.
        let mut stack: Vec<NodeId> = Vec::new();
        // Accumulates a text run; entity references arrive as separate events and are
        // merged in before the next structural event flushes the buffer.
        let mut pending = String::new();

        loop {
            match reader.read_event()? {
                Event::Decl(e) => {
                    let raw = std::str::from_utf8(&e).map_err(|err| {
                        Error::Malformed(format!("XML declaration is not UTF-8: {err}"))
                    })?;
                    decl = Some(raw.to_owned());
                }
                Event::Start(e) => {
                    flush_text(
                        &mut nodes,
                        &mut prolog,
                        &mut epilog,
                        root,
                        &stack,
                        &mut pending,
                    );
                    let element = read_element(&e)?;
                    let id = push_node(&mut nodes, NodeKind::Element(element));
                    place_element(&mut nodes, &mut root, &stack, id)?;
                    stack.push(id);
                }
                Event::Empty(e) => {
                    flush_text(
                        &mut nodes,
                        &mut prolog,
                        &mut epilog,
                        root,
                        &stack,
                        &mut pending,
                    );
                    let element = read_element(&e)?;
                    let id = push_node(&mut nodes, NodeKind::Element(element));
                    place_element(&mut nodes, &mut root, &stack, id)?;
                }
                Event::End(_) => {
                    flush_text(
                        &mut nodes,
                        &mut prolog,
                        &mut epilog,
                        root,
                        &stack,
                        &mut pending,
                    );
                    // check_end_names guarantees this matches the open element.
                    stack.pop();
                }
                Event::Text(e) => {
                    let text = e.decode().map_err(|e| Error::Xml(e.into()))?;
                    pending.push_str(&text);
                }
                Event::GeneralRef(e) => {
                    pending.push_str(&resolve_entity(&e)?);
                }
                Event::CData(e) => {
                    flush_text(
                        &mut nodes,
                        &mut prolog,
                        &mut epilog,
                        root,
                        &stack,
                        &mut pending,
                    );
                    let text = e.decode().map_err(|e| Error::Xml(e.into()))?;
                    let id = push_node(&mut nodes, NodeKind::CData(text.into_owned()));
                    place_misc(&mut nodes, &mut prolog, &mut epilog, root, &stack, id);
                }
                Event::Comment(e) => {
                    flush_text(
                        &mut nodes,
                        &mut prolog,
                        &mut epilog,
                        root,
                        &stack,
                        &mut pending,
                    );
                    let text = e.decode().map_err(|e| Error::Xml(e.into()))?;
                    let id = push_node(&mut nodes, NodeKind::Comment(text.into_owned()));
                    place_misc(&mut nodes, &mut prolog, &mut epilog, root, &stack, id);
                }
                Event::PI(e) => {
                    flush_text(
                        &mut nodes,
                        &mut prolog,
                        &mut epilog,
                        root,
                        &stack,
                        &mut pending,
                    );
                    let text = std::str::from_utf8(&e).map_err(|err| {
                        Error::Malformed(format!("processing instruction is not UTF-8: {err}"))
                    })?;
                    let id = push_node(&mut nodes, NodeKind::Pi(text.to_owned()));
                    place_misc(&mut nodes, &mut prolog, &mut epilog, root, &stack, id);
                }
                // OOXML forbids DTDs, so a DOCTYPE never appears in a .docx part.
                Event::DocType(_) => {}
                Event::Eof => {
                    flush_text(
                        &mut nodes,
                        &mut prolog,
                        &mut epilog,
                        root,
                        &stack,
                        &mut pending,
                    );
                    break;
                }
            }
        }

        let root = root.ok_or_else(|| Error::Malformed("no root element".into()))?;
        Ok(XmlTree {
            nodes,
            decl,
            prolog,
            root,
            epilog,
        })
    }

    /// Serialize the tree back to XML bytes.
    ///
    /// Emits the declaration (if any), the prolog, the root element, and the epilog.
    /// Text escapes `&`, `<`, `>`; attribute values escape `&`, `<`, `>`, `"` and encode
    /// literal newline / tab / carriage-return as `&#10;` / `&#9;` / `&#13;` so XML
    /// attribute-value normalization does not corrupt them on re-parse. Elements with no
    /// children serialize self-closing (`<w:p/>`).
    pub fn serialize(&self) -> Vec<u8> {
        let mut out = Vec::new();
        if let Some(decl) = &self.decl {
            out.extend_from_slice(b"<?");
            out.extend_from_slice(decl.as_bytes());
            out.extend_from_slice(b"?>");
        }
        for &id in &self.prolog {
            self.write_node(id, &mut out);
        }
        self.write_node(self.root, &mut out);
        for &id in &self.epilog {
            self.write_node(id, &mut out);
        }
        out
    }

    fn write_node(&self, id: NodeId, out: &mut Vec<u8>) {
        match &self.node(id).kind {
            NodeKind::Element(el) => {
                out.push(b'<');
                out.extend_from_slice(el.name.as_bytes());
                for (name, value) in &el.attrs {
                    out.push(b' ');
                    out.extend_from_slice(name.as_bytes());
                    out.extend_from_slice(b"=\"");
                    escape_attr(value, out);
                    out.push(b'"');
                }
                if el.children.is_empty() {
                    out.extend_from_slice(b"/>");
                } else {
                    out.push(b'>');
                    for &child in &el.children {
                        self.write_node(child, out);
                    }
                    out.extend_from_slice(b"</");
                    out.extend_from_slice(el.name.as_bytes());
                    out.push(b'>');
                }
            }
            NodeKind::Text(s) => escape_text(s, out),
            NodeKind::CData(s) => {
                out.extend_from_slice(b"<![CDATA[");
                out.extend_from_slice(s.as_bytes());
                out.extend_from_slice(b"]]>");
            }
            NodeKind::Comment(s) => {
                out.extend_from_slice(b"<!--");
                out.extend_from_slice(s.as_bytes());
                out.extend_from_slice(b"-->");
            }
            NodeKind::Pi(s) => {
                out.extend_from_slice(b"<?");
                out.extend_from_slice(s.as_bytes());
                out.extend_from_slice(b"?>");
            }
        }
    }

    // --- Navigation -------------------------------------------------------------

    /// The root element's id.
    pub fn root(&self) -> NodeId {
        self.root
    }

    /// The kind (content) of a node.
    pub fn kind(&self, id: NodeId) -> &NodeKind {
        &self.node(id).kind
    }

    /// The parent of a node, or `None` for the root and detached / top-level nodes.
    pub fn parent(&self, id: NodeId) -> Option<NodeId> {
        self.node(id).parent
    }

    /// A node's children in document order; an empty slice for non-element nodes.
    pub fn children(&self, id: NodeId) -> &[NodeId] {
        match &self.node(id).kind {
            NodeKind::Element(el) => &el.children,
            _ => &[],
        }
    }

    /// A node's qualified name if it is an element, else `None`.
    pub fn name(&self, id: NodeId) -> Option<&str> {
        match &self.node(id).kind {
            NodeKind::Element(el) => Some(&el.name),
            _ => None,
        }
    }

    /// The unescaped value of an element's attribute by qualified name.
    pub fn attr(&self, id: NodeId, name: &str) -> Option<&str> {
        match &self.node(id).kind {
            NodeKind::Element(el) => el
                .attrs
                .iter()
                .find(|(k, _)| k == name)
                .map(|(_, v)| v.as_str()),
            _ => None,
        }
    }

    /// All of an element's attributes in document order; an empty slice otherwise.
    pub fn attrs(&self, id: NodeId) -> &[(String, String)] {
        match &self.node(id).kind {
            NodeKind::Element(el) => &el.attrs,
            _ => &[],
        }
    }

    /// The text of a [`NodeKind::Text`] or [`NodeKind::CData`] node, else `None`.
    pub fn text(&self, id: NodeId) -> Option<&str> {
        match &self.node(id).kind {
            NodeKind::Text(s) | NodeKind::CData(s) => Some(s),
            _ => None,
        }
    }

    /// Concatenated text of all descendant text and CDATA nodes, in document order —
    /// the DOM `textContent` of the node.
    pub fn text_content(&self, id: NodeId) -> String {
        let mut out = String::new();
        for d in self.descendants(id) {
            if let NodeKind::Text(s) | NodeKind::CData(s) = &self.node(d).kind {
                out.push_str(s);
            }
        }
        out
    }

    /// The node and all its descendants in document order (pre-order, self first).
    pub fn descendants(&self, id: NodeId) -> impl Iterator<Item = NodeId> + '_ {
        Descendants {
            tree: self,
            stack: vec![id],
        }
    }

    /// An element's direct children that are elements with the given qualified name.
    pub fn children_named<'a>(
        &'a self,
        id: NodeId,
        name: &'a str,
    ) -> impl Iterator<Item = NodeId> + 'a {
        self.children(id)
            .iter()
            .copied()
            .filter(move |&c| self.name(c) == Some(name))
    }

    /// Resolve a namespace prefix to its URI by walking from `id` up through its
    /// ancestors, returning the first matching `xmlns:prefix` (or `xmlns` for the
    /// default namespace when `prefix` is `None`) declaration in scope.
    pub fn namespace_uri(&self, id: NodeId, prefix: Option<&str>) -> Option<&str> {
        let target = match prefix {
            Some(p) => format!("xmlns:{p}"),
            None => "xmlns".to_string(),
        };
        let mut cur = Some(id);
        while let Some(c) = cur {
            if let Some(v) = self.attr(c, &target) {
                return Some(v);
            }
            cur = self.parent(c);
        }
        None
    }

    // --- Mutation ---------------------------------------------------------------

    /// Create a detached element node with the given qualified name.
    pub fn create_element(&mut self, name: impl Into<String>) -> NodeId {
        push_node(
            &mut self.nodes,
            NodeKind::Element(Element {
                name: name.into(),
                attrs: Vec::new(),
                children: Vec::new(),
            }),
        )
    }

    /// Create a detached text node with the given (unescaped) text.
    pub fn create_text(&mut self, text: impl Into<String>) -> NodeId {
        push_node(&mut self.nodes, NodeKind::Text(text.into()))
    }

    /// Set an attribute on an element, replacing the value in place if the attribute
    /// already exists (keeping its position) or appending it otherwise.
    ///
    /// # Panics
    /// Panics if `id` is not an element.
    pub fn set_attr(&mut self, id: NodeId, name: impl Into<String>, value: impl Into<String>) {
        let name = name.into();
        let value = value.into();
        match &mut self.node_mut(id).kind {
            NodeKind::Element(el) => {
                if let Some(slot) = el.attrs.iter_mut().find(|(k, _)| *k == name) {
                    slot.1 = value;
                } else {
                    el.attrs.push((name, value));
                }
            }
            _ => panic!("set_attr on a non-element node"),
        }
    }

    /// Remove an attribute from an element by name. No-op if it is absent or `id` is
    /// not an element.
    pub fn remove_attr(&mut self, id: NodeId, name: &str) {
        if let NodeKind::Element(el) = &mut self.node_mut(id).kind {
            el.attrs.retain(|(k, _)| k != name);
        }
    }

    /// Append `child` as the last child of `parent`.
    ///
    /// # Panics
    /// Panics if `parent` is not an element or `child` already has a parent (detach it
    /// first with [`remove_from_parent`](Self::remove_from_parent)).
    pub fn append_child(&mut self, parent: NodeId, child: NodeId) {
        self.attach(parent, child, None);
    }

    /// Insert `child` at `index` among `parent`'s children.
    ///
    /// # Panics
    /// Panics if `parent` is not an element, `child` already has a parent, or `index`
    /// is greater than the current child count.
    pub fn insert_child(&mut self, parent: NodeId, index: usize, child: NodeId) {
        self.attach(parent, child, Some(index));
    }

    fn attach(&mut self, parent: NodeId, child: NodeId, index: Option<usize>) {
        assert!(
            self.node(child).parent.is_none(),
            "attaching a node that already has a parent; detach it first"
        );
        match &mut self.node_mut(parent).kind {
            NodeKind::Element(el) => match index {
                Some(i) => el.children.insert(i, child),
                None => el.children.push(child),
            },
            _ => panic!("append/insert child under a non-element node"),
        }
        self.node_mut(child).parent = Some(parent);
    }

    /// Detach a node from its parent element. The node stays in the arena (its id
    /// remains valid) and can be re-attached elsewhere. No-op if the node has no parent.
    pub fn remove_from_parent(&mut self, id: NodeId) {
        if let Some(parent) = self.node(id).parent {
            if let NodeKind::Element(el) = &mut self.node_mut(parent).kind {
                el.children.retain(|&c| c != id);
            }
            self.node_mut(id).parent = None;
        }
    }

    /// Replace the contents of a text node.
    ///
    /// # Panics
    /// Panics if `id` is not a [`NodeKind::Text`] node.
    pub fn set_text(&mut self, id: NodeId, text: impl Into<String>) {
        match &mut self.node_mut(id).kind {
            NodeKind::Text(s) => *s = text.into(),
            _ => panic!("set_text on a non-text node"),
        }
    }

    // --- Internal ---------------------------------------------------------------

    fn node(&self, id: NodeId) -> &Node {
        &self.nodes[id.0 as usize]
    }

    fn node_mut(&mut self, id: NodeId) -> &mut Node {
        &mut self.nodes[id.0 as usize]
    }
}

/// Pre-order descendant iterator (self first), borrowing the tree.
struct Descendants<'a> {
    tree: &'a XmlTree,
    stack: Vec<NodeId>,
}

impl Iterator for Descendants<'_> {
    type Item = NodeId;

    fn next(&mut self) -> Option<NodeId> {
        let id = self.stack.pop()?;
        // Push children in reverse so they are visited left-to-right.
        for &child in self.tree.children(id).iter().rev() {
            self.stack.push(child);
        }
        Some(id)
    }
}

/// Append a node to the arena and return its id (parent unset).
fn push_node(nodes: &mut Vec<Node>, kind: NodeKind) -> NodeId {
    let id = NodeId(nodes.len() as u32);
    nodes.push(Node { parent: None, kind });
    id
}

/// Build an [`Element`] from a start/empty tag, decoding the name and unescaping
/// attribute values (xmlns declarations included, in document order).
fn read_element(e: &quick_xml::events::BytesStart<'_>) -> Result<Element> {
    let name = std::str::from_utf8(e.name().as_ref())
        .map_err(|err| Error::Malformed(format!("element name is not UTF-8: {err}")))?
        .to_owned();
    let mut attrs = Vec::new();
    for attr in e.attributes() {
        let attr = attr.map_err(|err| Error::Xml(err.into()))?;
        let key = std::str::from_utf8(attr.key.as_ref())
            .map_err(|err| Error::Malformed(format!("attribute name is not UTF-8: {err}")))?
            .to_owned();
        let value = attr
            .normalized_value(quick_xml::XmlVersion::Implicit1_0)?
            .into_owned();
        attrs.push((key, value));
    }
    Ok(Element {
        name,
        attrs,
        children: Vec::new(),
    })
}

/// Resolve an entity reference event to its textual value (numeric char refs and the
/// five predefined entities).
fn resolve_entity(r: &BytesRef<'_>) -> Result<String> {
    if let Some(c) = r.resolve_char_ref()? {
        return Ok(c.to_string());
    }
    let name = r.decode().map_err(|e| Error::Xml(e.into()))?;
    let replacement = match name.as_ref() {
        "amp" => "&",
        "lt" => "<",
        "gt" => ">",
        "quot" => "\"",
        "apos" => "'",
        other => return Err(Error::Malformed(format!("unknown entity &{other};"))),
    };
    Ok(replacement.to_string())
}

/// Attach the root element, or a child under the innermost open element. Errors if a
/// second root element appears at the top level.
fn place_element(
    nodes: &mut [Node],
    root: &mut Option<NodeId>,
    stack: &[NodeId],
    id: NodeId,
) -> Result<()> {
    if let Some(&parent) = stack.last() {
        nodes[id.0 as usize].parent = Some(parent);
        if let NodeKind::Element(el) = &mut nodes[parent.0 as usize].kind {
            el.children.push(id);
        }
    } else if root.is_none() {
        *root = Some(id);
    } else {
        return Err(Error::Malformed("multiple root elements".into()));
    }
    Ok(())
}

/// Place a leaf node (text, CDATA, comment, PI): under the innermost open element, or
/// in the prolog / epilog when at the top level.
fn place_misc(
    nodes: &mut [Node],
    prolog: &mut Vec<NodeId>,
    epilog: &mut Vec<NodeId>,
    root: Option<NodeId>,
    stack: &[NodeId],
    id: NodeId,
) {
    if let Some(&parent) = stack.last() {
        nodes[id.0 as usize].parent = Some(parent);
        if let NodeKind::Element(el) = &mut nodes[parent.0 as usize].kind {
            el.children.push(id);
        }
    } else if root.is_none() {
        prolog.push(id);
    } else {
        epilog.push(id);
    }
}

/// Flush a pending text run into a text node, if non-empty.
fn flush_text(
    nodes: &mut Vec<Node>,
    prolog: &mut Vec<NodeId>,
    epilog: &mut Vec<NodeId>,
    root: Option<NodeId>,
    stack: &[NodeId],
    pending: &mut String,
) {
    if pending.is_empty() {
        return;
    }
    let text = std::mem::take(pending);
    let id = push_node(nodes, NodeKind::Text(text));
    place_misc(nodes, prolog, epilog, root, stack, id);
}

/// Escape text content: `&`, `<`, `>`.
fn escape_text(s: &str, out: &mut Vec<u8>) {
    for &b in s.as_bytes() {
        match b {
            b'&' => out.extend_from_slice(b"&amp;"),
            b'<' => out.extend_from_slice(b"&lt;"),
            b'>' => out.extend_from_slice(b"&gt;"),
            _ => out.push(b),
        }
    }
}

/// Escape an attribute value: `&`, `<`, `>`, `"`, plus newline / tab / carriage-return
/// as numeric refs so attribute-value normalization cannot rewrite them on re-parse.
fn escape_attr(s: &str, out: &mut Vec<u8>) {
    for &b in s.as_bytes() {
        match b {
            b'&' => out.extend_from_slice(b"&amp;"),
            b'<' => out.extend_from_slice(b"&lt;"),
            b'>' => out.extend_from_slice(b"&gt;"),
            b'"' => out.extend_from_slice(b"&quot;"),
            b'\n' => out.extend_from_slice(b"&#10;"),
            b'\t' => out.extend_from_slice(b"&#9;"),
            b'\r' => out.extend_from_slice(b"&#13;"),
            _ => out.push(b),
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn parse_and_navigate() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8"?>
<w:p xmlns:w="http://w"><w:r><w:t>Hello</w:t></w:r><w:r><w:t>World</w:t></w:r></w:p>"#;
        let tree = XmlTree::parse(xml).unwrap();
        let root = tree.root();
        assert_eq!(tree.name(root), Some("w:p"));
        assert_eq!(tree.children(root).len(), 2);
        assert_eq!(tree.text_content(root), "HelloWorld");

        let runs: Vec<_> = tree.children_named(root, "w:r").collect();
        assert_eq!(runs.len(), 2);
        let first_t = tree.children(runs[0])[0];
        assert_eq!(tree.name(first_t), Some("w:t"));
        assert_eq!(tree.text_content(first_t), "Hello");
    }

    #[test]
    fn attribute_unescaping() {
        let xml = br#"<a x="&amp; &quot;q&quot;&#10;tab&#9;end"/>"#;
        let tree = XmlTree::parse(xml).unwrap();
        assert_eq!(tree.attr(tree.root(), "x"), Some("& \"q\"\ntab\tend"));
    }

    #[test]
    fn text_entities_merge() {
        let xml = br#"<a>1 &amp; 2 &lt; 3 &#65;</a>"#;
        let tree = XmlTree::parse(xml).unwrap();
        // Entity references split the text run at parse; they must merge back into one.
        assert_eq!(tree.text_content(tree.root()), "1 & 2 < 3 A");
        assert_eq!(tree.children(tree.root()).len(), 1);
    }

    #[test]
    fn cdata_is_preserved() {
        let xml = br#"<a><![CDATA[<not> & markup]]></a>"#;
        let tree = XmlTree::parse(xml).unwrap();
        let child = tree.children(tree.root())[0];
        assert!(matches!(tree.kind(child), NodeKind::CData(_)));
        assert_eq!(tree.text(child), Some("<not> & markup"));
        let out = tree.serialize();
        assert_eq!(&out, br#"<a><![CDATA[<not> & markup]]></a>"#);
    }

    #[test]
    fn prolog_and_epilog_preserved() {
        let xml = b"<?xml version=\"1.0\"?>\n<!-- before -->\n<root/>\n<!-- after -->";
        let tree = XmlTree::parse(xml).unwrap();
        assert_eq!(tree.serialize(), xml);
    }

    #[test]
    fn empty_element_serializes_self_closing() {
        let tree = XmlTree::parse(b"<a></a>").unwrap();
        assert_eq!(tree.serialize(), b"<a/>");
    }

    #[test]
    fn mutation_keeps_structure_consistent() {
        let mut tree = XmlTree::parse(b"<root><a/><b/></root>").unwrap();
        let root = tree.root();

        let c = tree.create_element("c");
        tree.append_child(root, c);
        assert_eq!(tree.children(root).len(), 3);
        assert_eq!(tree.parent(c), Some(root));

        let first = tree.children(root)[0];
        let z = tree.create_element("z");
        tree.insert_child(root, 0, z);
        assert_eq!(tree.name(tree.children(root)[0]), Some("z"));
        assert_eq!(tree.name(tree.children(root)[1]), Some("a"));

        tree.remove_from_parent(first);
        assert_eq!(tree.parent(first), None);
        assert!(!tree.children(root).contains(&first));

        // A detached node can be re-attached.
        tree.append_child(c, first);
        assert_eq!(tree.parent(first), Some(c));
    }

    #[test]
    fn set_and_remove_attr_keep_position() {
        let mut tree = XmlTree::parse(br#"<a p="1" q="2" r="3"/>"#).unwrap();
        let root = tree.root();
        tree.set_attr(root, "q", "changed");
        assert_eq!(tree.attr(root, "q"), Some("changed"));
        // Position preserved: still the middle attribute.
        assert_eq!(tree.attrs(root)[1], ("q".into(), "changed".into()));

        tree.remove_attr(root, "p");
        assert_eq!(tree.attr(root, "p"), None);
        assert_eq!(tree.attrs(root).len(), 2);

        tree.set_attr(root, "new", "v");
        assert_eq!(tree.attrs(root).last().unwrap().0, "new");
    }

    #[test]
    #[should_panic(expected = "already has a parent")]
    fn attaching_parented_node_panics() {
        let mut tree = XmlTree::parse(b"<root><a/></root>").unwrap();
        let root = tree.root();
        let a = tree.children(root)[0];
        let b = tree.create_element("b");
        tree.append_child(b, a); // panics: `a` already has a parent
    }

    #[test]
    fn namespace_resolution_with_nesting() {
        let xml = br#"<w:document xmlns:w="http://word" xmlns="http://default">
            <w:body><v:shape xmlns:v="http://vml"><inner xmlns="http://redefined"/></v:shape></w:body>
        </w:document>"#;
        let tree = XmlTree::parse(xml).unwrap();
        let root = tree.root();

        // Prefix resolved at the declaring element.
        assert_eq!(tree.namespace_uri(root, Some("w")), Some("http://word"));
        // Default namespace at the root.
        assert_eq!(tree.namespace_uri(root, None), Some("http://default"));

        // Locate the deeply nested nodes.
        let shape = tree
            .descendants(root)
            .find(|&n| tree.name(n) == Some("v:shape"))
            .unwrap();
        let inner = tree
            .descendants(root)
            .find(|&n| tree.name(n) == Some("inner"))
            .unwrap();

        // `v` is only in scope from `v:shape` inward; `w` is inherited from the root.
        assert_eq!(tree.namespace_uri(shape, Some("v")), Some("http://vml"));
        assert_eq!(tree.namespace_uri(shape, Some("w")), Some("http://word"));
        assert_eq!(tree.namespace_uri(root, Some("v")), None);

        // A nested default-namespace redeclaration wins over an ancestor's.
        assert_eq!(tree.namespace_uri(inner, None), Some("http://redefined"));
        assert_eq!(tree.namespace_uri(shape, None), Some("http://default"));
    }
}

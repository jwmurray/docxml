//! The XML-tree fidelity guarantee: parse → serialize is semantically lossless for
//! every XML part in a real `.docx`, and a mutation touches only its target.

use docxml::opc::Package;
use docxml::xml::XmlTree;

use quick_xml::Reader;
use quick_xml::events::{BytesRef, Event};

const FIXTURE: &str = concat!(env!("CARGO_MANIFEST_DIR"), "/tests/fixtures/basic.docx");

/// A normalized XML token: the semantic content of an event stream with the
/// permitted normalizations applied (empty vs start+end, escaping, entity merging).
#[derive(Debug, PartialEq, Eq)]
enum Token {
    /// Declaration compared by (version, encoding, standalone), not raw bytes.
    Decl(Option<String>, Option<String>, Option<String>),
    /// Start tag: qualified name plus attributes (unescaped) in document order.
    Start(String, Vec<(String, String)>),
    End(String),
    /// A run of text/entities, unescaped and merged. Whitespace is significant.
    Text(String),
    CData(String),
    Comment(String),
    Pi(String),
}

/// Reduce an XML byte stream to its normalized token sequence.
fn tokenize(bytes: &[u8]) -> Vec<Token> {
    let mut reader = Reader::from_reader(bytes);
    reader.config_mut().check_end_names = true;

    let mut tokens = Vec::new();
    let mut pending = String::new();

    // Flush an accumulated text run as a single Text token (dropping empty runs, so a
    // start immediately followed by an end matches an empty element).
    fn flush(pending: &mut String, tokens: &mut Vec<Token>) {
        if !pending.is_empty() {
            tokens.push(Token::Text(std::mem::take(pending)));
        }
    }

    loop {
        match reader.read_event().unwrap() {
            Event::Decl(e) => {
                let version = e
                    .version()
                    .ok()
                    .map(|v| String::from_utf8_lossy(&v).into_owned());
                let encoding = e
                    .encoding()
                    .and_then(|r| r.ok())
                    .map(|v| String::from_utf8_lossy(&v).into_owned());
                let standalone = e
                    .standalone()
                    .and_then(|r| r.ok())
                    .map(|v| String::from_utf8_lossy(&v).into_owned());
                tokens.push(Token::Decl(version, encoding, standalone));
            }
            Event::Start(e) => {
                flush(&mut pending, &mut tokens);
                tokens.push(Token::Start(name_of(&e), attrs_of(&e)));
            }
            Event::Empty(e) => {
                flush(&mut pending, &mut tokens);
                // Empty element == start + end with no content between.
                let name = name_of(&e);
                tokens.push(Token::Start(name.clone(), attrs_of(&e)));
                tokens.push(Token::End(name));
            }
            Event::End(e) => {
                flush(&mut pending, &mut tokens);
                tokens.push(Token::End(
                    String::from_utf8_lossy(e.name().as_ref()).into_owned(),
                ));
            }
            Event::Text(e) => {
                pending.push_str(&e.decode().unwrap());
            }
            Event::GeneralRef(e) => {
                pending.push_str(&resolve(&e));
            }
            Event::CData(e) => {
                flush(&mut pending, &mut tokens);
                tokens.push(Token::CData(e.decode().unwrap().into_owned()));
            }
            Event::Comment(e) => {
                flush(&mut pending, &mut tokens);
                tokens.push(Token::Comment(e.decode().unwrap().into_owned()));
            }
            Event::PI(e) => {
                flush(&mut pending, &mut tokens);
                tokens.push(Token::Pi(String::from_utf8_lossy(&e).into_owned()));
            }
            Event::DocType(_) => {}
            Event::Eof => {
                flush(&mut pending, &mut tokens);
                break;
            }
        }
    }
    tokens
}

fn name_of(e: &quick_xml::events::BytesStart<'_>) -> String {
    String::from_utf8_lossy(e.name().as_ref()).into_owned()
}

fn attrs_of(e: &quick_xml::events::BytesStart<'_>) -> Vec<(String, String)> {
    e.attributes()
        .map(|a| {
            let a = a.unwrap();
            let key = String::from_utf8_lossy(a.key.as_ref()).into_owned();
            let value = a
                .normalized_value(quick_xml::XmlVersion::Implicit1_0)
                .unwrap()
                .into_owned();
            (key, value)
        })
        .collect()
}

fn resolve(r: &BytesRef<'_>) -> String {
    if let Some(c) = r.resolve_char_ref().unwrap() {
        return c.to_string();
    }
    match r.decode().unwrap().as_ref() {
        "amp" => "&",
        "lt" => "<",
        "gt" => ">",
        "quot" => "\"",
        "apos" => "'",
        other => panic!("unknown entity &{other};"),
    }
    .to_string()
}

/// Assert two XML byte streams carry the same meaning under the documented
/// normalizations. Attribute order is compared; whitespace-only text is significant.
fn assert_xml_semantically_equal(original: &[u8], serialized: &[u8]) {
    let a = tokenize(original);
    let b = tokenize(serialized);
    if a != b {
        // Surface the first divergence for a readable failure.
        for (i, (x, y)) in a.iter().zip(b.iter()).enumerate() {
            assert!(
                x == y,
                "token {i} differs:\n  original:   {x:?}\n  serialized: {y:?}"
            );
        }
        panic!(
            "token streams differ in length: original {} vs serialized {}",
            a.len(),
            b.len()
        );
    }
}

#[test]
fn every_xml_part_roundtrips_semantically() {
    let pkg = Package::open(FIXTURE).unwrap();
    let mut checked = 0;
    for part in pkg.parts() {
        if !(part.name.ends_with(".xml") || part.name.ends_with(".rels")) {
            continue;
        }
        let tree = XmlTree::parse(&part.data)
            .unwrap_or_else(|e| panic!("parse failed for {}: {e}", part.name));
        let serialized = tree.serialize();
        assert_xml_semantically_equal(&part.data, &serialized);

        // Re-parse the serialized output and compare again, guarding against a
        // serializer that emits well-formed-looking but unparseable bytes.
        let reparsed = XmlTree::parse(&serialized)
            .unwrap_or_else(|e| panic!("reparse failed for {}: {e}", part.name));
        assert_xml_semantically_equal(&part.data, &reparsed.serialize());
        checked += 1;
    }
    assert!(checked > 0, "no .xml/.rels parts found in fixture");
}

#[test]
fn mutation_lands_and_leaves_everything_else_intact() {
    let pkg = Package::open(FIXTURE).unwrap();
    let doc = pkg.part("word/document.xml").unwrap();
    let original = &doc.data;

    // Find a deep element carrying at least one attribute and change that attribute.
    let mut tree = XmlTree::parse(original).unwrap();
    let root = tree.root();
    let target = tree
        .descendants(root)
        .find(|&id| id != root && !tree.attrs(id).is_empty())
        .expect("an attributed element deep in the document");
    let attr_name = tree.attrs(target)[0].0.clone();
    let old_value = tree.attrs(target)[0].1.clone();
    let new_value = format!("{old_value}__mutated");
    tree.set_attr(target, attr_name.clone(), new_value.clone());

    let serialized = tree.serialize();
    let reparsed = XmlTree::parse(&serialized).unwrap();

    // The mutation landed.
    let reparsed_root = reparsed.root();
    let same_path: Vec<usize> = path_to(&tree, target);
    let reparsed_target = node_at_path(&reparsed, reparsed_root, &same_path);
    assert_eq!(
        reparsed.attr(reparsed_target, &attr_name),
        Some(new_value.as_str())
    );

    // Everything else is unchanged: restoring the old value reproduces the original
    // document semantically.
    let mut restored = XmlTree::parse(&serialized).unwrap();
    let restored_root = restored.root();
    let restored_target = node_at_path(&restored, restored_root, &same_path);
    restored.set_attr(restored_target, &attr_name, old_value);
    assert_xml_semantically_equal(original, &restored.serialize());
}

/// The child-index path from the root to `id`.
fn path_to(tree: &XmlTree, id: docxml::xml::NodeId) -> Vec<usize> {
    let mut path = Vec::new();
    let mut cur = id;
    while let Some(parent) = tree.parent(cur) {
        let idx = tree
            .children(parent)
            .iter()
            .position(|&c| c == cur)
            .unwrap();
        path.push(idx);
        cur = parent;
    }
    path.reverse();
    path
}

/// Follow a child-index path from `start`.
fn node_at_path(tree: &XmlTree, start: docxml::xml::NodeId, path: &[usize]) -> docxml::xml::NodeId {
    let mut cur = start;
    for &idx in path {
        cur = tree.children(cur)[idx];
    }
    cur
}

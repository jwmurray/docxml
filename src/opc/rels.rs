//! Parsing of OPC relationship parts (`*.rels`).

use quick_xml::Reader;
use quick_xml::events::Event;

use crate::error::{Error, Result};

/// One `<Relationship>` entry from a `.rels` part.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Relationship {
    pub id: String,
    pub rel_type: String,
    pub target: String,
    /// True when `TargetMode="External"` (hyperlinks etc.).
    pub external: bool,
}

/// Parse the `<Relationship>` entries out of a `.rels` part's bytes.
pub fn parse_relationships(data: &[u8]) -> Result<Vec<Relationship>> {
    let mut reader = Reader::from_reader(data);
    reader.config_mut().trim_text(true);
    let mut rels = Vec::new();
    loop {
        match reader.read_event()? {
            Event::Start(e) | Event::Empty(e) if e.local_name().as_ref() == b"Relationship" => {
                let mut id = None;
                let mut rel_type = None;
                let mut target = None;
                let mut external = false;
                for attr in e.attributes() {
                    let attr = attr.map_err(|e| Error::Xml(e.into()))?;
                    let value = attr
                        .normalized_value(quick_xml::XmlVersion::Implicit1_0)?
                        .into_owned();
                    match attr.key.as_ref() {
                        b"Id" => id = Some(value),
                        b"Type" => rel_type = Some(value),
                        b"Target" => target = Some(value),
                        b"TargetMode" => external = value == "External",
                        _ => {}
                    }
                }
                match (id, rel_type, target) {
                    (Some(id), Some(rel_type), Some(target)) => rels.push(Relationship {
                        id,
                        rel_type,
                        target,
                        external,
                    }),
                    _ => {
                        return Err(Error::InvalidPackage(
                            "Relationship missing Id, Type, or Target".into(),
                        ));
                    }
                }
            }
            Event::Eof => break,
            _ => {}
        }
    }
    Ok(rels)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn parses_package_rels() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com/" TargetMode="External"/>
</Relationships>"#;
        let rels = parse_relationships(xml).unwrap();
        assert_eq!(rels.len(), 2);
        assert_eq!(rels[0].id, "rId1");
        assert_eq!(rels[0].target, "word/document.xml");
        assert!(!rels[0].external);
        assert!(rels[1].external);
    }

    #[test]
    fn rejects_incomplete_relationship() {
        let xml = br#"<Relationships><Relationship Id="rId1"/></Relationships>"#;
        assert!(parse_relationships(xml).is_err());
    }
}

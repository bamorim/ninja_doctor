use std::fmt::Write;
use std::fs::File;
use zip::ZipArchive;
use std::io::{BufReader, Error as IOError};
use std::path::Path;
use minidom::{Element, Node};
use minidom::quick_xml::Reader;
use zip::read::ZipFile;
use zip::result::ZipError;

pub enum DocxError {
    IO(IOError),
    Zip(ZipError),
    XML(minidom::Error),
    Fmt(core::fmt::Error)
}

impl From<IOError> for DocxError {
    fn from(err: IOError) -> Self {
        DocxError::IO(err)
    }
}

impl From<ZipError> for DocxError {
    fn from(err: ZipError) -> Self {
        DocxError::Zip(err)
    }
}

impl From<minidom::Error> for DocxError {
    fn from(err: minidom::Error) -> Self {
        DocxError::XML(err)
    }
}

impl From<core::fmt::Error> for DocxError {
    fn from(err: core::fmt::Error) -> Self {
        DocxError::Fmt(err)
    }
}

fn read_xml_file(zip_file: ZipFile) -> Result<Element, minidom::Error> {
    let buf_reader = BufReader::with_capacity(128 * 1024, zip_file);
    let mut xml_reader = Reader::from_reader(buf_reader);
    return Element::from_reader(&mut xml_reader);
}

pub fn parse_file<P: AsRef<Path>>(path: P) -> Result<String, DocxError> {
    let mut zip = ZipArchive::new(File::open(path)?)?;

    let document = read_xml_file(zip.by_name("word/document.xml")?)?;
    let mut buffer = html_builder::Buffer::new();
    parse_element(&document, &mut buffer)?;

    let result = buffer.finish();

    return Ok(result);
}

fn parse_node(node: &Node, target: &mut html_builder::Node) -> Result<(), DocxError> {
    match node {
        Node::Element(element) => parse_element(element, target)?,
        Node::Text(text) =>  target.write_str(text)?
    }

    Ok(())
}

fn parse_element(element: &Element, target: &mut html_builder::Node) -> Result<(), DocxError> {
    for child in element.nodes() {
        parse_node(child, target)?;
    }

    Ok(())
}

#[cfg(test)]
mod tests {
    use std::path::PathBuf;
    use super::*;

    #[test]
    fn test_simple() {
        let parsed = parse_file(fixture("simple.docx"));
        assert!(matches!(parsed, Ok(_)));
        if let Ok(result) = parsed {
            assert!(result.contains("Hello"));
        }
    }

    fn fixture(name: &str) -> PathBuf {
        let mut path = PathBuf::from(env!("CARGO_MANIFEST_DIR"));
        path.push("test_fixtures");
        path.push(name);
        return path;
    }
}
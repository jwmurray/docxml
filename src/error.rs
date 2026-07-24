/// Errors returned by docxml operations.
#[derive(Debug, thiserror::Error)]
pub enum Error {
    #[error("I/O error: {0}")]
    Io(#[from] std::io::Error),

    #[error("zip error: {0}")]
    Zip(#[from] zip::result::ZipError),

    #[error("XML error: {0}")]
    Xml(#[from] quick_xml::Error),

    #[error("not a valid OPC package: {0}")]
    InvalidPackage(String),

    #[error("malformed XML: {0}")]
    Malformed(String),

    #[error("image error: {0}")]
    Image(String),

    #[error("invalid merge: {0}")]
    InvalidMerge(String),
}

pub type Result<T> = std::result::Result<T, Error>;

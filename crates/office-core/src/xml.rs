use quick_xml::Writer;
use quick_xml::events::{BytesDecl, BytesEnd, BytesStart, BytesText, Event};
use std::io::Write;

/// Helper for building XML documents with quick-xml.
pub struct XmlBuilder<W: Write> {
    writer: Writer<W>,
}

impl<W: Write> XmlBuilder<W> {
    pub fn new(inner: W) -> Self {
        Self {
            writer: Writer::new_with_indent(inner, b' ', 2),
        }
    }

    pub fn write_declaration(&mut self) -> quick_xml::Result<()> {
        self.writer
            .write_event(Event::Decl(BytesDecl::new("1.0", Some("UTF-8"), Some("yes"))))?;
        Ok(())
    }

    pub fn start_element(&mut self, tag: &str) -> quick_xml::Result<()> {
        let elem = BytesStart::new(tag);
        self.writer.write_event(Event::Start(elem))?;
        Ok(())
    }

    pub fn start_element_with_attrs(
        &mut self,
        tag: &str,
        attrs: &[(&str, &str)],
    ) -> quick_xml::Result<()> {
        let mut elem = BytesStart::new(tag);
        for (key, val) in attrs {
            elem.push_attribute((*key, *val));
        }
        self.writer.write_event(Event::Start(elem))?;
        Ok(())
    }

    pub fn end_element(&mut self, tag: &str) -> quick_xml::Result<()> {
        self.writer
            .write_event(Event::End(BytesEnd::new(tag)))?;
        Ok(())
    }

    pub fn empty_element_with_attrs(
        &mut self,
        tag: &str,
        attrs: &[(&str, &str)],
    ) -> quick_xml::Result<()> {
        let mut elem = BytesStart::new(tag);
        for (key, val) in attrs {
            elem.push_attribute((*key, *val));
        }
        self.writer.write_event(Event::Empty(elem))?;
        Ok(())
    }

    pub fn write_text(&mut self, text: &str) -> quick_xml::Result<()> {
        self.writer
            .write_event(Event::Text(BytesText::new(text)))?;
        Ok(())
    }

    pub fn write_text_element(&mut self, tag: &str, text: &str) -> quick_xml::Result<()> {
        self.start_element(tag)?;
        self.write_text(text)?;
        self.end_element(tag)?;
        Ok(())
    }

    pub fn write_text_element_with_attrs(
        &mut self,
        tag: &str,
        attrs: &[(&str, &str)],
        text: &str,
    ) -> quick_xml::Result<()> {
        self.start_element_with_attrs(tag, attrs)?;
        self.write_text(text)?;
        self.end_element(tag)?;
        Ok(())
    }

    pub fn into_inner(self) -> W {
        self.writer.into_inner()
    }
}

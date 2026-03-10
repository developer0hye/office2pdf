use std::cell::Cell;
use std::collections::HashSet;

pub(in super::super) struct BidiContext {
    bidi_indices: HashSet<usize>,
    cursor: Cell<usize>,
}

impl BidiContext {
    pub(in super::super) fn from_xml(xml: Option<&str>) -> Self {
        let bidi_indices = xml.map(Self::scan).unwrap_or_default();
        Self {
            bidi_indices,
            cursor: Cell::new(0),
        }
    }

    pub(in super::super) fn next_is_bidi(&self) -> bool {
        let index = self.cursor.get();
        self.cursor.set(index + 1);
        self.bidi_indices.contains(&index)
    }

    fn scan(xml: &str) -> HashSet<usize> {
        let mut reader = quick_xml::Reader::from_str(xml);
        let mut buffer: Vec<u8> = Vec::new();
        let mut result: HashSet<usize> = HashSet::new();
        let mut paragraph_index: usize = 0;
        let mut in_paragraph_properties = false;
        let mut in_body = false;

        loop {
            match reader.read_event_into(&mut buffer) {
                Ok(quick_xml::events::Event::Start(ref element))
                | Ok(quick_xml::events::Event::Empty(ref element)) => {
                    match element.local_name().as_ref() {
                        b"body" => in_body = true,
                        b"pPr" if in_body => in_paragraph_properties = true,
                        b"bidi" if in_paragraph_properties => {
                            result.insert(paragraph_index);
                        }
                        _ => {}
                    }
                }
                Ok(quick_xml::events::Event::End(ref element)) => {
                    match element.local_name().as_ref() {
                        b"body" => in_body = false,
                        b"p" if in_body => {
                            paragraph_index += 1;
                            in_paragraph_properties = false;
                        }
                        b"pPr" => in_paragraph_properties = false,
                        _ => {}
                    }
                }
                Ok(quick_xml::events::Event::Eof) => break,
                Err(_) => break,
                _ => {}
            }
            buffer.clear();
        }

        result
    }
}

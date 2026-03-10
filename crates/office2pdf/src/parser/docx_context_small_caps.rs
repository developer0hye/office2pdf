use std::cell::Cell;

pub(in super::super) struct SmallCapsContext {
    flags: Vec<bool>,
    cursor: Cell<usize>,
}

impl SmallCapsContext {
    pub(in super::super) fn from_xml(xml: Option<&str>) -> Self {
        let flags = xml.map(Self::scan).unwrap_or_default();
        Self {
            flags,
            cursor: Cell::new(0),
        }
    }

    pub(in super::super) fn next_is_small_caps(&self) -> bool {
        let index = self.cursor.get();
        self.cursor.set(index + 1);
        self.flags.get(index).copied().unwrap_or(false)
    }

    fn scan(xml: &str) -> Vec<bool> {
        let mut reader = quick_xml::Reader::from_str(xml);
        let mut buffer: Vec<u8> = Vec::new();
        let mut result: Vec<bool> = Vec::new();
        let mut in_body = false;
        let mut in_run = false;
        let mut in_run_properties = false;
        let mut current_has_small_caps = false;

        loop {
            match reader.read_event_into(&mut buffer) {
                Ok(quick_xml::events::Event::Start(ref element))
                | Ok(quick_xml::events::Event::Empty(ref element)) => {
                    match element.local_name().as_ref() {
                        b"body" => in_body = true,
                        b"r" if in_body => {
                            in_run = true;
                            current_has_small_caps = false;
                        }
                        b"rPr" if in_run => in_run_properties = true,
                        b"smallCaps" if in_run_properties => {
                            let is_disabled = element.attributes().flatten().any(|attribute| {
                                attribute.key.local_name().as_ref() == b"val"
                                    && matches!(attribute.value.as_ref(), b"false" | b"0")
                            });
                            if !is_disabled {
                                current_has_small_caps = true;
                            }
                        }
                        _ => {}
                    }
                }
                Ok(quick_xml::events::Event::End(ref element)) => {
                    match element.local_name().as_ref() {
                        b"body" => in_body = false,
                        b"r" if in_body => {
                            result.push(current_has_small_caps);
                            in_run = false;
                            in_run_properties = false;
                            current_has_small_caps = false;
                        }
                        b"rPr" => in_run_properties = false,
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

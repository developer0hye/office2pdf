use std::cell::Cell;

#[derive(Debug, Clone, Copy, Default)]
pub(in super::super) struct TableHeaderInfo {
    pub(in super::super) repeat_rows: usize,
}

pub(in super::super) struct TableHeaderContext {
    headers: Vec<TableHeaderInfo>,
    cursor: Cell<usize>,
}

impl TableHeaderContext {
    pub(in super::super) fn from_xml(xml: Option<&str>) -> Self {
        Self {
            headers: xml.map(scan_table_headers).unwrap_or_default(),
            cursor: Cell::new(0),
        }
    }

    pub(in super::super) fn consume_next(&self) -> TableHeaderInfo {
        let index = self.cursor.get();
        self.cursor.set(index + 1);
        self.headers.get(index).copied().unwrap_or_default()
    }
}

struct TableHeaderScanState {
    table_index: usize,
    repeat_rows: usize,
    in_row: bool,
    current_row_is_header: bool,
    saw_body_row: bool,
}

#[cfg(test)]
pub(in super::super) fn scan_table_headers(xml: &str) -> Vec<TableHeaderInfo> {
    scan_table_headers_impl(xml)
}

#[cfg(not(test))]
pub(in super::super) fn scan_table_headers(xml: &str) -> Vec<TableHeaderInfo> {
    scan_table_headers_impl(xml)
}

fn scan_table_headers_impl(xml: &str) -> Vec<TableHeaderInfo> {
    let mut reader = quick_xml::Reader::from_str(xml);
    let mut buffer: Vec<u8> = Vec::new();
    let mut headers: Vec<TableHeaderInfo> = Vec::new();
    let mut stack: Vec<TableHeaderScanState> = Vec::new();

    loop {
        match reader.read_event_into(&mut buffer) {
            Ok(quick_xml::events::Event::Start(ref element)) => match element.local_name().as_ref()
            {
                b"tbl" => {
                    headers.push(TableHeaderInfo::default());
                    stack.push(TableHeaderScanState {
                        table_index: headers.len() - 1,
                        repeat_rows: 0,
                        in_row: false,
                        current_row_is_header: false,
                        saw_body_row: false,
                    });
                }
                b"tr" => {
                    if let Some(state) = stack.last_mut() {
                        state.in_row = true;
                        state.current_row_is_header = false;
                    }
                }
                b"tblHeader" => {
                    if let Some(state) = stack.last_mut()
                        && state.in_row
                        && on_off_element_is_enabled(element)
                    {
                        state.current_row_is_header = true;
                    }
                }
                _ => {}
            },
            Ok(quick_xml::events::Event::Empty(ref element)) => match element.local_name().as_ref()
            {
                b"tbl" => headers.push(TableHeaderInfo::default()),
                b"tr" => {
                    if let Some(state) = stack.last_mut() {
                        state.in_row = true;
                        state.current_row_is_header = false;
                        finalize_table_header_row(state);
                    }
                }
                b"tblHeader" => {
                    if let Some(state) = stack.last_mut()
                        && state.in_row
                        && on_off_element_is_enabled(element)
                    {
                        state.current_row_is_header = true;
                    }
                }
                _ => {}
            },
            Ok(quick_xml::events::Event::End(ref element)) => match element.local_name().as_ref() {
                b"tr" => {
                    if let Some(state) = stack.last_mut() {
                        finalize_table_header_row(state);
                    }
                }
                b"tbl" => {
                    if let Some(state) = stack.pop() {
                        headers[state.table_index].repeat_rows = state.repeat_rows;
                    }
                }
                _ => {}
            },
            Ok(quick_xml::events::Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
        buffer.clear();
    }

    headers
}

fn finalize_table_header_row(state: &mut TableHeaderScanState) {
    if !state.in_row {
        return;
    }

    if !state.saw_body_row && state.current_row_is_header {
        state.repeat_rows += 1;
    } else {
        state.saw_body_row = true;
    }

    state.in_row = false;
    state.current_row_is_header = false;
}

fn on_off_element_is_enabled(element: &quick_xml::events::BytesStart<'_>) -> bool {
    for attribute in element.attributes().flatten() {
        if attribute.key.local_name().as_ref() != b"val" {
            continue;
        }

        let value = attribute.value.as_ref();
        if value.eq_ignore_ascii_case(b"0")
            || value.eq_ignore_ascii_case(b"false")
            || value.eq_ignore_ascii_case(b"off")
        {
            return false;
        }
    }

    true
}

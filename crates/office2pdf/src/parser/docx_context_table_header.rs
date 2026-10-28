use std::cell::Cell;

#[derive(Debug, Clone, Default)]
pub(in super::super) struct TableHeaderInfo {
    pub(in super::super) repeat_rows: usize,
    pub(in super::super) is_visual_rtl: bool,
    /// Each `w:tr`'s `w:cantSplit`, in document order. docx-rs drops the
    /// element, so the flag is read from the story XML here (issue #2217).
    pub(in super::super) cant_split_rows: Vec<bool>,
    /// The `w:tblpPr` `w:topFromText` and `w:bottomFromText`, in twips.
    /// docx-rs reads the table position's other attributes but not these two
    /// (issue #2229).
    pub(in super::super) vertical_from_text_twips: (Option<f64>, Option<f64>),
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

    /// The `w:tblpPr` vertical distances of the table [`Self::consume_next`]
    /// yields next, without consuming it.
    pub(in super::super) fn next_vertical_from_text_twips(&self) -> (Option<f64>, Option<f64>) {
        self.headers
            .get(self.cursor.get())
            .map(|header| header.vertical_from_text_twips)
            .unwrap_or_default()
    }

    pub(in super::super) fn consume_next(&self) -> TableHeaderInfo {
        let index = self.cursor.get();
        self.cursor.set(index + 1);
        self.headers.get(index).cloned().unwrap_or_default()
    }
}

struct TableHeaderScanState {
    table_index: usize,
    repeat_rows: usize,
    in_row: bool,
    current_row_is_header: bool,
    saw_body_row: bool,
    in_table_properties: bool,
    is_visual_rtl: bool,
    cant_split_rows: Vec<bool>,
    vertical_from_text_twips: (Option<f64>, Option<f64>),
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
                        in_table_properties: false,
                        is_visual_rtl: false,
                        cant_split_rows: Vec::new(),
                        vertical_from_text_twips: (None, None),
                    });
                }
                b"tblPr" => {
                    if let Some(state) = stack.last_mut() {
                        state.in_table_properties = true;
                    }
                }
                b"tr" => {
                    if let Some(state) = stack.last_mut() {
                        state.in_row = true;
                        state.current_row_is_header = false;
                        state.cant_split_rows.push(false);
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
                b"cantSplit" => {
                    if let Some(state) = stack.last_mut()
                        && state.in_row
                        && on_off_element_is_enabled(element)
                        && let Some(cant_split) = state.cant_split_rows.last_mut()
                    {
                        *cant_split = true;
                    }
                }
                b"bidiVisual" => {
                    if let Some(state) = stack.last_mut()
                        && state.in_table_properties
                        && on_off_element_is_enabled(element)
                    {
                        state.is_visual_rtl = true;
                    }
                }
                b"tblpPr" => {
                    if let Some(state) = stack.last_mut()
                        && state.in_table_properties
                    {
                        state.vertical_from_text_twips = vertical_from_text_twips(element);
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
                        state.cant_split_rows.push(false);
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
                b"cantSplit" => {
                    if let Some(state) = stack.last_mut()
                        && state.in_row
                        && on_off_element_is_enabled(element)
                        && let Some(cant_split) = state.cant_split_rows.last_mut()
                    {
                        *cant_split = true;
                    }
                }
                b"bidiVisual" => {
                    if let Some(state) = stack.last_mut()
                        && state.in_table_properties
                        && on_off_element_is_enabled(element)
                    {
                        state.is_visual_rtl = true;
                    }
                }
                b"tblpPr" => {
                    if let Some(state) = stack.last_mut()
                        && state.in_table_properties
                    {
                        state.vertical_from_text_twips = vertical_from_text_twips(element);
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
                b"tblPr" => {
                    if let Some(state) = stack.last_mut() {
                        state.in_table_properties = false;
                    }
                }
                b"tbl" => {
                    if let Some(state) = stack.pop() {
                        headers[state.table_index].repeat_rows = state.repeat_rows;
                        headers[state.table_index].is_visual_rtl = state.is_visual_rtl;
                        headers[state.table_index].cant_split_rows = state.cant_split_rows;
                        headers[state.table_index].vertical_from_text_twips =
                            state.vertical_from_text_twips;
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

/// A `w:tblpPr`'s `w:topFromText` and `w:bottomFromText`, in twips.
fn vertical_from_text_twips(
    element: &quick_xml::events::BytesStart<'_>,
) -> (Option<f64>, Option<f64>) {
    let mut distances: (Option<f64>, Option<f64>) = (None, None);
    for attribute in element.attributes().flatten() {
        let slot: &mut Option<f64> = match attribute.key.local_name().as_ref() {
            b"topFromText" => &mut distances.0,
            b"bottomFromText" => &mut distances.1,
            _ => continue,
        };
        *slot = std::str::from_utf8(&attribute.value)
            .ok()
            .and_then(|value| value.trim().parse::<f64>().ok())
            .filter(|twips| twips.is_finite());
    }
    distances
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

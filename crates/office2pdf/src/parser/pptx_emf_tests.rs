use super::*;

fn fixture_records() -> Vec<Vec<u8>> {
    use std::io::Read;
    let fixture = include_bytes!("../../../../tests/fixtures/docx/docxide/emf_vector_table.docx");
    let mut archive = zip::ZipArchive::new(std::io::Cursor::new(fixture)).unwrap();
    let mut bytes = Vec::new();
    archive
        .by_name("word/media/table.emf")
        .unwrap()
        .read_to_end(&mut bytes)
        .unwrap();
    let mut records = Vec::new();
    let mut offset = 0;
    while offset < bytes.len() {
        let size = read_u32(&bytes, offset + 4).unwrap() as usize;
        records.push(bytes[offset..offset + size].to_vec());
        offset += size;
    }
    records
}

fn record(kind: u32, body: &[u8]) -> Vec<u8> {
    let mut data = kind.to_le_bytes().to_vec();
    data.extend_from_slice(&((body.len() + 8) as u32).to_le_bytes());
    data.extend_from_slice(body);
    data
}

#[test]
fn emf_text_preserves_font_and_color_across_saved_contexts() {
    let records = fixture_records();
    let text = records.iter().find(|r| read_u32(r, 0) == Some(84)).unwrap();
    let mut input = records[0..4].concat();
    input.extend(record(18, &1u32.to_le_bytes()));
    input.extend(record(37, &1u32.to_le_bytes()));
    input.extend(record(33, &[]));
    input.extend(record(37, &2u32.to_le_bytes()));
    input.extend(record(24, &255u32.to_le_bytes()));
    input.extend(text);
    input.extend(record(34, &(-1i32).to_le_bytes()));
    input.extend(text);
    input.extend(record(14, &[]));
    let svg = String::from_utf8(convert_emf_to_svg(&input).unwrap()).unwrap();
    let mut reader = quick_xml::Reader::from_str(&svg);
    let mut styles = Vec::new();
    loop {
        match reader.read_event().unwrap() {
            quick_xml::events::Event::Start(e) if e.local_name().as_ref() == b"text" => {
                let attributes: std::collections::HashMap<String, String> = e
                    .attributes()
                    .map(|a| {
                        let a = a.unwrap();
                        (
                            String::from_utf8(a.key.as_ref().to_vec()).unwrap(),
                            a.decoded_and_normalized_value(
                                quick_xml::XmlVersion::Implicit1_0,
                                reader.decoder(),
                            )
                            .unwrap()
                            .into_owned(),
                        )
                    })
                    .collect();
                styles.push((
                    attributes["font-weight"].clone(),
                    attributes["fill"].clone(),
                ));
            }
            quick_xml::events::Event::Eof => break,
            _ => {}
        }
    }
    assert_eq!(
        styles,
        [
            ("700".into(), "#ff0000".into()),
            ("400".into(), "#000000".into())
        ]
    );
}

#[test]
fn emf_office_font_declares_sans_serif_paint_fallbacks() {
    let mut records = fixture_records();
    for font in records.iter_mut().filter(|r| read_u32(r, 0) == Some(82)) {
        font[40..104].fill(0);
        for (index, unit) in "Calibri".encode_utf16().enumerate() {
            font[40 + index * 2..42 + index * 2].copy_from_slice(&unit.to_le_bytes());
        }
    }
    let svg = String::from_utf8(convert_emf_to_svg(&records.concat()).unwrap()).unwrap();
    let mut reader = quick_xml::Reader::from_str(&svg);
    let mut text_count = 0;
    loop {
        match reader.read_event().unwrap() {
            quick_xml::events::Event::Start(element)
                if element.local_name().as_ref() == b"text" =>
            {
                let family = element
                    .try_get_attribute("font-family")
                    .unwrap()
                    .unwrap()
                    .decoded_and_normalized_value(
                        quick_xml::XmlVersion::Implicit1_0,
                        reader.decoder(),
                    )
                    .unwrap()
                    .into_owned();
                assert!(
                    family.contains("Helvetica"),
                    "sans-serif paint fallback absent: {family}"
                );
                text_count += 1;
            }
            quick_xml::events::Event::Eof => break,
            _ => {}
        }
    }
    assert_eq!(text_count, 12);
}

#[test]
fn emf_text_rejects_out_of_record_strings_and_advance_arrays() {
    let records = fixture_records();
    let index = records
        .iter()
        .position(|r| read_u32(r, 0) == Some(84))
        .unwrap();
    for field in [44, 48, 72] {
        let mut malformed = records.clone();
        malformed[index][field..field + 4].copy_from_slice(&u32::MAX.to_le_bytes());
        assert!(
            convert_emf_to_svg(&malformed.concat()).is_none(),
            "field {field}"
        );
    }
}

#[test]
fn emf_surrogate_pair_uses_both_utf16_advances_without_splitting_the_character() {
    let records = fixture_records();
    let mut input = records[0..4].concat();
    input.extend(record(37, &1u32.to_le_bytes()));
    input.extend(record(18, &1u32.to_le_bytes()));
    let text = "A🌐B";
    let units: Vec<u16> = text.encode_utf16().collect();
    let mut body = vec![0u8; 68];
    body[28..32].copy_from_slice(&10i32.to_le_bytes());
    body[32..36].copy_from_slice(&20i32.to_le_bytes());
    body[36..40].copy_from_slice(&(units.len() as u32).to_le_bytes());
    body[40..44].copy_from_slice(&76u32.to_le_bytes());
    body[64..68].copy_from_slice(&84u32.to_le_bytes());
    for unit in units {
        body.extend(unit.to_le_bytes());
    }
    for advance in [5i32, 7, 8, 9] {
        body.extend(advance.to_le_bytes());
    }
    input.extend(record(84, &body));
    input.extend(record(14, &[]));
    let svg = String::from_utf8(convert_emf_to_svg(&input).unwrap()).unwrap();
    assert!(svg.contains("A🌐B"));
    assert!(svg.contains("x=\"10 15 30\""), "{svg}");
}

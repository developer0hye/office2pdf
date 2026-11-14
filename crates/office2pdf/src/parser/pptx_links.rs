//! Slide hyperlinks: where an `a:hlinkClick` leads (issue #2473).
//!
//! A click action names its target through the owning part's relationships,
//! so an `r:id` on a layout or master shape resolves against that layout's or
//! master's `.rels`, not the slide's. A slide jump names a slide part, and the
//! PDF knows only pages, so the deck's slide-to-page map turns it into a page
//! index. `hlinkshowjump` moves relative to the slide the action is shown on,
//! which for a layout or master shape is every slide that uses it.

use std::collections::HashMap;

use crate::ir::{Block, FixedElementKind, FixedPage, LinkTarget, Table};
use crate::parser::xml_util::{RelationshipEntry, parse_relationships, resolve_relative_path};

const HYPERLINK_RELATIONSHIP_SUFFIX: &str = "/hyperlink";
const SLIDE_RELATIONSHIP_SUFFIX: &str = "/slide";
const SLIDE_JUMP_ACTION: &str = "ppaction://hlinksldjump";
const SHOW_JUMP_ACTION_PREFIX: &str = "ppaction://hlinkshowjump?jump=";

/// The deck's printed pages, as link resolution needs them.
#[derive(Debug, Default)]
pub(super) struct DeckPages {
    /// Slide part path (`ppt/slides/slide3.xml`) to the index of the page it
    /// prints on. A hidden or out-of-range slide prints nothing and is absent.
    pub(super) page_by_slide: HashMap<String, usize>,
    /// How many pages the deck prints.
    pub(super) page_count: usize,
}

/// Resolves the click actions of one part (slide, layout or master) as shown
/// on one printed page.
pub(super) struct PartLinks<'a> {
    part_dir: String,
    relationships: HashMap<String, RelationshipEntry>,
    deck: &'a DeckPages,
    /// The page the part is being drawn on.
    page: usize,
}

impl<'a> PartLinks<'a> {
    /// Links for the part at `part_path`, whose relationships are `rels_xml`,
    /// drawn on `page`.
    pub(super) fn new(part_path: &str, rels_xml: &str, deck: &'a DeckPages, page: usize) -> Self {
        let part_dir: String = part_path
            .rsplit_once('/')
            .map_or_else(String::new, |(dir, _)| dir.to_string());
        let relationships: HashMap<String, RelationshipEntry> = parse_relationships(rels_xml)
            .into_iter()
            .map(|entry| (entry.id.clone(), entry))
            .collect();
        Self {
            part_dir,
            relationships,
            deck,
            page,
        }
    }

    /// Where an `a:hlinkClick` with these `r:id` and `action` attributes leads,
    /// or `None` for an action a PDF cannot carry (a macro, a program, a custom
    /// show, ending the show) or a target the deck does not print.
    pub(super) fn resolve(&self, rid: Option<&str>, action: Option<&str>) -> Option<LinkTarget> {
        let action: &str = action.unwrap_or("").trim();
        let relationship = rid
            .filter(|rid| !rid.is_empty())
            .and_then(|rid| self.relationships.get(rid));
        if action.is_empty() {
            let relationship = relationship?;
            let is_hyperlink: bool = relationship
                .rel_type
                .as_deref()
                .is_some_and(|rel_type| rel_type.ends_with(HYPERLINK_RELATIONSHIP_SUFFIX));
            let url: &str = relationship.target.trim();
            return (is_hyperlink && !url.is_empty()).then(|| LinkTarget::Url(url.to_string()));
        }
        if action == SLIDE_JUMP_ACTION {
            let relationship = relationship?;
            let is_slide: bool = relationship
                .rel_type
                .as_deref()
                .is_some_and(|rel_type| rel_type.ends_with(SLIDE_RELATIONSHIP_SUFFIX));
            if !is_slide {
                return None;
            }
            let slide_path: String = resolve_relative_path(&self.part_dir, &relationship.target);
            return self
                .deck
                .page_by_slide
                .get(&slide_path)
                .map(|page| LinkTarget::Page(*page));
        }
        let jump: &str = action.strip_prefix(SHOW_JUMP_ACTION_PREFIX)?;
        let last_page: usize = self.deck.page_count.checked_sub(1)?;
        let page: usize = match jump {
            "firstslide" => 0,
            "lastslide" => last_page,
            "nextslide" => self.page.checked_add(1).filter(|page| *page <= last_page)?,
            "previousslide" => self.page.checked_sub(1)?,
            _ => return None,
        };
        Some(LinkTarget::Page(page))
    }
}

/// Renumber every page link on `page` through `remap`, dropping the links
/// whose target no longer prints.
///
/// The deck's page map is built before the slides are parsed, so a slide that
/// then fails to parse shifts every later page down by one.
pub(super) fn remap_page_links(page: &mut FixedPage, remap: &dyn Fn(usize) -> Option<usize>) {
    page.links
        .retain_mut(|area| remap_target(&mut area.target, remap));
    for element in &mut page.elements {
        match &mut element.kind {
            FixedElementKind::TextBox(text_box) => remap_blocks(&mut text_box.content, remap),
            FixedElementKind::Table(table) => remap_table(table, remap),
            _ => {}
        }
    }
}

fn remap_blocks(blocks: &mut [Block], remap: &dyn Fn(usize) -> Option<usize>) {
    for block in blocks {
        match block {
            Block::Paragraph(paragraph) => remap_runs(&mut paragraph.runs, remap),
            Block::Table(table) => remap_table(table, remap),
            Block::List(list) => {
                for item in &mut list.items {
                    for paragraph in &mut item.content {
                        remap_runs(&mut paragraph.runs, remap);
                    }
                }
            }
            _ => {}
        }
    }
}

fn remap_table(table: &mut Table, remap: &dyn Fn(usize) -> Option<usize>) {
    for row in &mut table.rows {
        for cell in &mut row.cells {
            remap_blocks(&mut cell.content, remap);
        }
    }
}

fn remap_runs(runs: &mut [crate::ir::Run], remap: &dyn Fn(usize) -> Option<usize>) {
    for run in runs {
        if let Some(target) = run.href.as_mut()
            && !remap_target(target, remap)
        {
            run.href = None;
        }
    }
}

/// Renumber one target in place; `false` when it no longer prints.
fn remap_target(target: &mut LinkTarget, remap: &dyn Fn(usize) -> Option<usize>) -> bool {
    match target {
        LinkTarget::Url(_) => true,
        LinkTarget::Page(page) => match remap(*page) {
            Some(new_page) => {
                *page = new_page;
                true
            }
            None => false,
        },
    }
}

#[cfg(test)]
#[path = "pptx_links_tests.rs"]
mod tests;

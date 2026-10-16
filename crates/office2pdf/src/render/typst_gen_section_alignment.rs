use crate::ir::{Block, Page, PageVerticalAlignment};

/// Word counts the terminal paragraph gap in centered/bottom page content.
/// Typst trims block spacing at a region edge. A stronger zero weak spacer
/// cancels that outer gap; a strong spacer then preserves its source value.
/// Before-gaps remain weak so automatic page breaks still suppress them.
pub(super) fn paragraph_edge_spacing(block: &Block) -> Option<(f64, f64)> {
    let gaps = match block {
        Block::Paragraph(paragraph) => (paragraph.style.space_before, paragraph.style.space_after),
        Block::Caption(caption) => (
            caption.paragraph.style.space_before,
            caption.paragraph.style.space_after,
        ),
        Block::List(list) => (
            list.items.first()?.content.first()?.style.space_before,
            list.items.last()?.content.last()?.style.space_after,
        ),
        Block::Image(image) => {
            let spacing = image.paragraph_spacing.as_ref();
            (
                spacing.and_then(|s| s.before),
                spacing.and_then(|s| s.after),
            )
        }
        Block::InlineImages(images) => (
            images
                .first()?
                .paragraph_spacing
                .as_ref()
                .and_then(|s| s.before),
            images
                .last()?
                .paragraph_spacing
                .as_ref()
                .and_then(|s| s.after),
        ),
        _ => return None,
    };
    Some((
        gaps.0.unwrap_or(0.0).max(0.0),
        gaps.1.unwrap_or(0.0).max(0.0),
    ))
}

pub(super) fn leaves_flow_gap_unchanged(block: &Block) -> bool {
    matches!(
        block,
        Block::FloatingImage(_)
            | Block::FloatingTextBox(_)
            | Block::FloatingShape(_)
            | Block::FloatingTable(_)
    )
}

/// Resolve alignment on pages delimited by explicit breaks. Continuous
/// sections can share only the preceding section's final explicit segment.
/// Automatic pagination inside mixed segments remains a documented limitation.
pub(super) fn alignment_plan(pages: &[Page]) -> Vec<Vec<PageVerticalAlignment>> {
    let mut plan: Vec<Vec<PageVerticalAlignment>> = vec![Vec::new(); pages.len()];
    let mut group: Vec<(usize, usize)> = Vec::new();
    let mut common: Option<PageVerticalAlignment> = None;
    let flush = |plan: &mut Vec<Vec<PageVerticalAlignment>>,
                 group: &mut Vec<(usize, usize)>,
                 common: &mut Option<PageVerticalAlignment>| {
        for (section, segment) in group.drain(..) {
            plan[section][segment] = common.unwrap_or(PageVerticalAlignment::Top);
        }
        *common = None;
    };
    for (index, page) in pages.iter().enumerate() {
        let (Page::Flow(flow) | Page::FlowContinuous(flow)) = page else {
            flush(&mut plan, &mut group, &mut common);
            continue;
        };
        let alignment = match flow.vertical_alignment {
            Some(PageVerticalAlignment::Center) => PageVerticalAlignment::Center,
            Some(PageVerticalAlignment::Bottom) => PageVerticalAlignment::Bottom,
            _ => PageVerticalAlignment::Top,
        };
        let mut segments = 1 + flow
            .content
            .iter()
            .filter(|block| matches!(block, Block::PageBreak | Block::PageBreakBefore))
            .count();
        // These emitters suppress a trailing empty segment rather than a page.
        if (flow.columns.is_some() || matches!(page, Page::FlowContinuous(_)))
            && matches!(
                flow.content.last(),
                Some(Block::PageBreak | Block::PageBreakBefore)
            )
        {
            segments -= 1;
        }
        plan[index] = vec![alignment; segments];
        for segment in 0..segments {
            let joins_previous = segment == 0 && matches!(page, Page::FlowContinuous(_));
            if !joins_previous {
                flush(&mut plan, &mut group, &mut common);
            }
            common = Some(match common {
                Some(previous) if previous != alignment => PageVerticalAlignment::Top,
                Some(previous) => previous,
                None => alignment,
            });
            group.push((index, segment));
        }
    }
    flush(&mut plan, &mut group, &mut common);
    plan
}

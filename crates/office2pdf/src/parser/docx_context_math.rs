use std::collections::HashMap;

use crate::ir::MathEquation;
use crate::parser::omml;

pub(in super::super) struct MathContext {
    equations: HashMap<usize, Vec<MathEquation>>,
}

impl MathContext {
    pub(in super::super) fn empty() -> Self {
        Self {
            equations: HashMap::new(),
        }
    }

    pub(in super::super) fn take(&mut self, index: usize) -> Vec<MathEquation> {
        self.equations.remove(&index).unwrap_or_default()
    }
}

pub(in super::super) fn build_math_context_from_xml(doc_xml: Option<&str>) -> MathContext {
    let mut equations: HashMap<usize, Vec<MathEquation>> = HashMap::new();

    if let Some(xml) = doc_xml {
        let raw = omml::scan_math_equations(xml);
        for (index, content, display) in raw {
            equations
                .entry(index)
                .or_default()
                .push(MathEquation { content, display });
        }
    }

    MathContext { equations }
}

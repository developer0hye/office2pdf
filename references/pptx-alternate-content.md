# PPTX alternate-content selection

- [.NET XmlCompatibilityReader](https://source.dot.net/System.IO.Packaging/System/IO/Packaging/XmlCompatibilityReader.cs.html) resolves every `Requires` prefix to a namespace URI and selects the first understood choice, otherwise the fallback.
- [Apache POI XPathHelper](https://poi.apache.org/apidocs/6.0/org/apache/poi/ooxml/util/XPathHelper.html) traverses alternative representations when selecting DrawingML properties; its documented choice/fallback search reinforces that alternatives are not cumulative objects.

Apply selection before slide, layout, master and embedded-object scans. Retain
the selected XML wrappers to preserve inherited namespace bindings, whitespace
and entities. Nested alternatives have independent selection state. An empty
selected choice still suppresses fallback; an unsupported choice without a
fallback contributes nothing.

The supported URI list is deliberately limited to the drawing vocabularies
already consumed by the PPTX parser. This does not establish full Office Math
layout fidelity; its existing formatting differences are independent of branch
selection.

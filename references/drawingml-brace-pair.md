# DrawingML brace pair

- Source: the vendored ECMA-376 Annex D `bracePair` definition in `crates/office2pdf/src/parser/drawingml_presets/presetShapeDefinitions.xml`; provenance and license are recorded beside that file.
- The adjustment defaults to 8333, clamped to 0..25000; the circular radius is `min(width,height) * adjustment / 100000`.
- Each side has four quarter-circle arcs, two straight vertical segments, and an open stroke. Horizontal and vertical mirroring preserve the pair's symmetric outline.
- The fill connects both sides into one closed boundary. Its horizontal joining edges must not be stroked. The parser emits a separate fill element only when the vector shape actually has fill paint.
- Native PowerPoint 16.115.1 confirms 240 x 300 pt bounds with an approximately 20pt default radius for the public #2487 input. The same arc sampler as custom geometry is reused; no new runtime dependency is needed.

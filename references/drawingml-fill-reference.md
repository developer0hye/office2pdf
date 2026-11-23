# DrawingML fill references

- Source: [Microsoft Open XML SDK fillRef remarks](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.fillreference), quoting ISO/IEC 29500-1.
- `idx=0` and `idx=1000` mean no fill. A child color is a placeholder for a selected theme entry; it is not an independent fill.
- `1..999` selects a fill style; `1001` and above selects a background fill style.
- In the slide parser, suppress reference color and alpha for both no-fill sentinel indices while preserving explicit shape fills and strokes.
- Verified on the synthetic brace-pair input from #2487 using native PowerPoint 16.115.1. The unsupported brace outline is independent of reference-color suppression.

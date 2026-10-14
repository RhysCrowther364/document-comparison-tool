# RWS artwork provenance

These assets come from the user-supplied `RWS Office Guidance May 2025.pdf`.

- `rws-logo.png`: original primary full-colour combined logo with black wordmark, page 4, PDF image object 14 and alpha-mask object 13, 422 x 168. Its original artwork and aspect ratio are preserved. It appears on white with surrounding space.
- `rws-logomark.ico`: original isolated full-colour logomark, page 4, image object 16 and alpha-mask object 15. The whole mark is centred within transparent square padding at 16, 24, 32, 48, 64, 128 and 256 pixels, preserving its aspect ratio.
- `document-compare.svg` / `document-compare.png` / `document-compare.ico`: custom document-comparison symbol, introduced in 1.0.8. Two paper sheets and opposing arrows use Galaxy Purple, Hillier Pink and white. The company logo is a separate asset. `create-comparison-icon.py` renders the native vector geometry locally with Pillow (a source-artwork tool only), simplifying the 16/24 pixel frames for clarity. The ICO contains 16, 24, 32, 48, 64, 128 and 256 pixel frames. This is the current application/window/desktop icon; the earlier logomark asset remains as source provenance.

Galaxy Purple `#3E016F` and Hillier Pink `#E60054` in `BrandAppearance.cs` are rounded from the PDF's page 5 vector RGB fills. The printed guide gives CMYK/PMS values. Neutral white and grey surfaces follow the signage examples; the app is an adaptation of that physical-signage guide.

The PDF embeds Articulat CF. The app uses that family when installed and otherwise Arial. No fonts are extracted or distributed. No guide PDF, reference-page screenshots, OCR text or network asset dependency is shipped with the app.

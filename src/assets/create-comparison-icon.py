"""Render the app's simple vector icon locally; Pillow is a build-time tool only."""
from pathlib import Path
from PIL import Image, ImageDraw

ASSETS = Path(__file__).resolve().parent
PURPLE = "#3E016F"
PINK = "#E60054"

# SVG is the editable master. Shapes represent an archive sheet, a new sheet
# and two arrows comparing their contents. No company logo is modified.
SVG = '''<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 256 256">
  <title>Document Compare</title>
  <desc>Two documents with opposing comparison arrows in RWS purple and pink.</desc>
  <path d="M24 20H94L130 56V218H24Z" fill="#3E016F"/>
  <path d="M92 20V58H130Z" fill="#E60054"/>
  <path d="M92 48H192L238 94V242H92Z" fill="white"/>
  <path d="M102 58H186L226 98V230H102Z" fill="white" stroke="#3E016F" stroke-width="12" stroke-linecap="round" stroke-linejoin="round"/>
  <path d="M188 60V96H224" fill="none" stroke="#3E016F" stroke-width="12" stroke-linecap="round" stroke-linejoin="round"/>
  <path d="M46 86H76" fill="none" stroke-width="12" stroke-linecap="round" stroke="white"/>
  <path d="M124 204H192" fill="none" stroke-width="12" stroke-linecap="round" stroke="#E60054"/>
  <path d="M184 166H50" fill="none" stroke="white" stroke-width="32" stroke-linecap="round"/>
  <path d="M64 152L50 166L64 180" fill="none" stroke="white" stroke-width="30" stroke-linecap="round" stroke-linejoin="round"/>
  <path d="M56 126H190M176 112L190 126L176 140" fill="none" stroke="#E60054" stroke-width="16" stroke-linecap="round" stroke-linejoin="round"/>
  <path d="M184 166H50M64 152L50 166L64 180" fill="none" stroke="#3E016F" stroke-width="16" stroke-linecap="round" stroke-linejoin="round"/>
</svg>
'''

def render(size):
    scale = size * 4 / 256
    canvas = Image.new("RGBA", (size * 4, size * 4), (0, 0, 0, 0))
    draw = ImageDraw.Draw(canvas)
    def points(coords):
        return [(round(x * scale), round(y * scale)) for x, y in coords]
    def polygon(coords, fill):
        draw.polygon(points(coords), fill=fill)
    def line(coords, fill, width):
        p = points(coords)
        w = max(1, round(width * scale))
        draw.line(p, fill=fill, width=w, joint="curve")
        r = w / 2
        for x, y in p:
            draw.ellipse((x-r, y-r, x+r, y+r), fill=fill)

    # Broad paper silhouettes retain their shape on either light or dark desktops.
    polygon([(24,20),(94,20),(130,56),(130,218),(24,218)], PURPLE)
    if size >= 32:
        polygon([(92,20),(92,58),(130,58)], PINK)
    # White separator around the foreground sheet makes the overlap visible.
    polygon([(92,48),(192,48),(238,94),(238,242),(92,242)], "white")
    front = [(102,58),(186,58),(226,98),(226,230),(102,230)]
    polygon(front, "white")
    outline = 12 if size >= 32 else 18
    line(front + [front[0]], PURPLE, outline)
    if size >= 32:
        line([(188,60),(188,96),(224,96)], PURPLE, 12)
        line([(46,86),(76,86)], "white", 12)
        line([(124,204),(192,204)], PINK, 12)

    # A white lane separates the purple return arrow from the archive sheet.
    line([(184,166),(50,166)], "white", 32)
    line([(64,152),(50,166),(64,180)], "white", 30)
    arrow_width = 16 if size >= 32 else 20
    line([(56,126),(190,126)], PINK, arrow_width)
    line([(176,112),(190,126),(176,140)], PINK, arrow_width)
    line([(184,166),(50,166)], PURPLE, arrow_width)
    line([(64,152),(50,166),(64,180)], PURPLE, arrow_width)
    return canvas.resize((size, size), Image.Resampling.LANCZOS)

if __name__ == "__main__":
    ASSETS.joinpath("document-compare.svg").write_text(SVG, encoding="utf-8")
    sizes = [16, 24, 32, 48, 64, 128, 256]
    frames = [render(size) for size in sizes]
    frames[-1].save(ASSETS / "document-compare.png")
    frames[-1].save(ASSETS / "document-compare.ico", format="ICO",
                    sizes=[(size, size) for size in sizes], append_images=frames[:-1])
    # Review the actual frames on both backgrounds, including their native sizes.
    preview = Image.new("RGB", (660, 320), "white")
    for row, background in enumerate(["#FFFFFF", "#24202A"]):
        panel = Image.new("RGB", (660, 160), background)
        x = 24
        for frame in frames[:-1]:
            panel.paste(frame, (x, (160-frame.height)//2), frame)
            x += frame.width + 24
        preview.paste(panel, (0, row * 160))
    preview_path = ASSETS.parent.parent / ".build" / "comparison-icon-sizes.png"
    preview_path.parent.mkdir(parents=True, exist_ok=True)
    preview.save(preview_path)
    print("Created local SVG, PNG and multi-size ICO.")

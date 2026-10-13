# Japanese source text fonts

`NotoSansJP-Text-Regular.ttf` and `NotoSansJP-Text-Bold.ttf` are static instances
of the existing `NotoSansJP-Regular.ttf` variable font at `wght=400` and
`wght=700`. They retain the Noto Sans JP family and licence metadata. The licence
is included in `OFL-NotoSansJP.txt`.

These two faces let Canvas measure and compare real regular/bold Japanese
glyphs without depending on installed Office fonts or variable-font defaults.
The original variable font remains available to existing image text features.

To reproduce, use FontTools with the original variable font:

```python
from fontTools.ttLib import TTFont
from fontTools.varLib.instancer import instantiateVariableFont

for weight, label in [(400, "Regular"), (700, "Bold")]:
    font = TTFont("NotoSansJP-Regular.ttf")
    font = instantiateVariableFont(font, {"wght": weight}, inplace=True,
                                   updateFontNames=True)
    font.save(f"NotoSansJP-Text-{label}.ttf")
```

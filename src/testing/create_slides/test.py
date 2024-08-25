from pptx import Presentation
from pptx.dml.color import ColorFormat, RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.util import Inches

prs = Presentation()
title_slide_layout = prs.slide_layouts[0]

# extract slide text
# get number of slides

total = 10
i = 0
while i <= total:
    slide = prs.slides.add_slide(title_slide_layout)

    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = RGBColor(0, 255, 0)

    left = Inches(0)
    top = Inches(3)
    width = Inches(10)
    height = Inches(2.5)
    rect = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, left, top, width, height
    )
    
    # send rectangle to back    
    slide.shapes._spTree.remove(rect._element)
    slide.shapes._spTree.insert(2, rect._element)
    
    title = slide.shapes.title
    subtitle = slide.placeholders[1]
    title.text = "Hello, World! " + str(i)
    subtitle.text = "python-pptx was here!"

    i += 1

prs.save('test.pptx')

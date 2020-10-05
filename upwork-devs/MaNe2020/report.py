from pptx import Presentation
from pptx.util import Inches
from pptx.dml.color import RGBColor

prs = Presentation('GWtemplate.pptx')

title_slide_layout = prs.slide_layouts[0]

slide = prs.slides.add_slide(title_slide_layout)
title = slide.shapes.title
title.text = "Threat Intelligence Reports"

#subtitle.text = "python-pptx was here!"

# SLIDE 1
slide1 = prs.slides.add_slide(title_slide_layout)
title1 = slide1.shapes.title
title1.text = "Overall summary and overview"

# SLIDE 2
slide2 = prs.slides.add_slide(prs.slide_layouts[1])
title2 = slide2.shapes.title
title2.text = "Emails and files processed"
shapes = slide2.shapes
cols = 2
rows = 7
left = top = Inches(1.0)
width = Inches(8.0)
height = Inches(3.0)
table = shapes.add_table(rows, cols, left, top, width, height).table
table.columns[0].width = Inches(6.0)
table.columns[1].width = Inches(2.0)
# write column headings
table.cell(0, 0).text = 'File Status'
table.cell(0, 1).text = 'Total number of files'
# write body cells
table.cell(1, 0).text = 'Allowed - The original file is provided to the recipient'
table.cell(2, 0).text = 'Disallowed - No file is provided to the recipient'
table.cell(3, 0).text = 'Clean - The file was rebuilt but required no remediation or sanitization'
table.cell(4, 0).text = 'Remediated - The file required structural fixes for conformance in regeneration'
table.cell(5, 0).text = 'Sanitized - The file required removal of active content '
table.cell(6, 0).text = 'Held - The file was not sent to the recipient due to being Malware'
# write body cells
table.cell(1, 1).text = '1'
table.cell(2, 1).text = '2'
table.cell(3, 1).text = '3'
table.cell(4, 1).text = '4'
table.cell(5, 1).text = '5'
table.cell(6, 1).text = '6'


# SLIDE 3
slide3 = prs.slides.add_slide(prs.slide_layouts[1])
title3 = slide3.shapes.title
title3.text = "Outcomes by file type"
shapes = slide3.shapes
cols = 2
rows = 7
left = top = Inches(1.0)
width = Inches(8.0)
height = Inches(3.0)
table = shapes.add_table(rows, cols, left, top, width, height).table
table.columns[0].width = Inches(6.0)
table.columns[1].width = Inches(2.0)
# write column headings
table.cell(0, 0).text = 'File Status'
table.cell(0, 1).text = 'Total number of files'
# write body cells
table.cell(1, 0).text = 'Allowed - The original file is provided to the recipient'
table.cell(2, 0).text = 'Disallowed - No file is provided to the recipient'
table.cell(3, 0).text = 'Clean - The file was rebuilt but required no remediation or sanitization'
table.cell(4, 0).text = 'Remediated - The file required structural fixes for conformance in regeneration'
table.cell(5, 0).text = 'Sanitized - The file required removal of active content '
table.cell(6, 0).text = 'Held - The file was not sent to the recipient due to being Malware'
# write body cells
table.cell(1, 1).text = '1'
table.cell(2, 1).text = '2'
table.cell(3, 1).text = '3'
table.cell(4, 1).text = '4'
table.cell(5, 1).text = '5'
table.cell(6, 1).text = '6'

# SLIDE 4
slide4 = prs.slides.add_slide(prs.slide_layouts[1])
title4 = slide4.shapes.title
title4.text = "Modern and legacy Microsoft Office file types"
pic = slide4.shapes.add_picture("images/supported.png", left, top)

# SLIDE 5
slide5 = prs.slides.add_slide(prs.slide_layouts[1])
title5 = slide5.shapes.title
title5.text = "Unsupported file types count"
shapes = slide5.shapes
cols = 2
rows = 3
left = top = Inches(2.0)
width = Inches(2.0)
height = Inches(2.0)
table = shapes.add_table(rows, cols, left, top, width, height).table
table.columns[0].width = Inches(2.5)
table.columns[1].width = Inches(1.5)
# write column headings
table.cell(0, 0).text = 'Unsupported File Type'
table.cell(0, 1).text = 'Count'
# write body cells
table.cell(1, 0).text = 'A'
table.cell(2, 0).text = 'B'
# write body cells
table.cell(1, 1).text = '1'
table.cell(2, 1).text = '2'


# SLIDE 6
slide6 = prs.slides.add_slide(prs.slide_layouts[1])
title6 = slide6.shapes.title
title6.text = "Malware detection summary"
shapes = slide6.shapes
cols = 2
rows = 4
left = top = Inches(2.0)
width = Inches(2.0)
height = Inches(2.0)
table = shapes.add_table(rows, cols, left, top, width, height).table
table.columns[0].width = Inches(3.5)
table.columns[1].width = Inches(2.5)
table.first_row = False
# write column headings
table.cell(0, 0).text = 'Malware file types'
table.cell(0, 1).text = 'Type A, Type B, Type C'
# write body cells
table.cell(1, 0).text = 'Total Number of Malware files'
table.cell(1, 1).text = 'XYZ'
# write body cells
table.cell(2, 0).text = 'Number of Unique Malware files'
table.cell(2, 1).text = 'XYZ'
# write body cells
table.cell(3, 0).text = 'Percent of files allowed to enter organisation'
table.cell(3, 1).text = 'XY%'


# SLIDE 7
slide7 = prs.slides.add_slide(prs.slide_layouts[1])
title7 = slide7.shapes.title
title7.text = "Malicious file details"
left = top = width = height = Inches(0.5)
txBox = slide7.shapes.add_textbox(left, top, width, height)
tx = txBox.text_frame
subtitle = tx.add_paragraph()
run = subtitle.add_run()
run.text = 'Use SHA256 hash to look up the malicious file on: https://www.virustotal.com/'
font = run.font
font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
# subtitle.text = "Use SHA256 hash to look up the malicious file on: https://www.virustotal.com/"
shapes = slide7.shapes
cols = 2
rows = 3
left = top = Inches(2.0)
width = Inches(2.0)
height = Inches(2.0)
table = shapes.add_table(rows, cols, left, top, width, height).table
table.columns[0].width = Inches(3.5)
table.columns[1].width = Inches(2.5)
# write column headings
table.cell(0, 0).text = 'Malicious File Name'
table.cell(0, 1).text = 'SHA256 Hash'
# write body cells
table.cell(1, 0).text = 'A'
table.cell(2, 0).text = 'B'
# write body cells
table.cell(1, 1).text = '1'
table.cell(2, 1).text = '2'


prs.save('test.pptx')
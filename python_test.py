from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
import io

doc_buffer = io.BytesIO()  # Save the created document to a BytesIO buffer
# Add a heading
document = Document()
heading = document.add_heading('Dinosaurs', level=1)
font = heading.runs[0].font
font.size = Pt(24)
paragraph_format = heading.paragraph_format
paragraph_format.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

# Add the table with headers
table = document.add_table(rows=1, cols=4)
heading_cells = table.rows[0].cells
heading_cells[0].text = 'Weight'
heading_cells[1].text = 'Size'
heading_cells[2].text = 'Ferocity'
heading_cells[3].text = 'Colour'

# Add three rows to the table for dinosaur comparisons
for _ in range(3):
    row_cells = table.add_row().cells
    row_cells[0].text = 'Value'
    row_cells[1].text = 'Value'
    row_cells[2].text = 'Value'
    row_cells[3].text = 'Value'

# Add three paragraphs
for _ in range(3):
    document.add_paragraph('This is a paragraph about a dinosaur.')
    document.add_paragraph(
        'It provides details such as habitat and behaviors.')
    document.add_paragraph(
        'Finally, it discusses the dinosaur’s legacy and any interesting facts.')

document.save("hello.docx")  # save doc to this buffer variable
doc_buffer.seek(0)

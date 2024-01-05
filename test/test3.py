from docx import Document
from docx.shared import Pt
from docx.enum.section import WD_SECTION

def set_custom_page_borders(document):
    section = document.sections[0]  # Assuming you have only one section in the document
    section.start_type
    section.start_type
    section.footer_distance = Pt(1)  # Adjust the footer distance as needed
    section.left_margin = Pt(1)  # Adjust the left margin as needed
    section.right_margin = Pt(1)  # Adjust the right margin as needed

# Create a new Word document
doc = Document()

# Add some content to the document
doc.add_paragraph("This is a sample document with custom page borders.")

# Set custom page borders for the document
set_custom_page_borders(doc)

# Save the document
doc.save("custom_page_borders_example.docx")

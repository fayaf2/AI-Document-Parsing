from docx import Document
from docx.shared import Pt

def get_heading_font_sizes(docx_path):
    document = Document(docx_path)
    headings = []

    for paragraph in document.paragraphs:
        # Check if the paragraph style is a heading style
        if paragraph.style.name.startswith('Heading'):
            font_size = None
            for run in paragraph.runs:
                if run.font.size:
                    font_size = Pt(run.font.size).pt  # Convert to points
                    break  # Assuming all runs in a heading have the same size
            headings.append((paragraph.text, font_size))

    return headings

# Replace 'your_file.docx' with the path to your .docx file
headings = get_heading_font_sizes('ss.docx')
for text, size in headings:
    print(f'Heading: {text} | Font size: {size} pt')

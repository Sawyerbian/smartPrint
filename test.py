from docx import Document
from docx.shared import Inches


from docx import Document

# Load the Word document
doc = Document("demo.docx")

# Read and print all paragraphs
for para in doc.paragraphs:
    print(para.text)
doc.tables[0].cell(0,0).text = "upated value"
print(doc.tables[0].cell(0,0).text)
doc.save("demo.docx")
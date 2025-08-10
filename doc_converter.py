# pip install python-docx docxcompose

from docxcompose.composer import Composer
from docx import Document

file1 = "1.docx"
file2 = "2.docx"
output = "merged_python.docx"

doc1 = Document(file1)
doc2 = Document(file2)

composer = Composer(doc1)
composer.append(doc2)
composer.save(output)

print("Merge complete! Saved as", output)
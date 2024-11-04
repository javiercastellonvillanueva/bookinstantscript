import openai
from pdf2docx import Converter
import docx
from docx.shared import Pt
import pypandoc

# Define file paths
contract_pdf = "contract.pdf"
template_pdf = "template.pdf"
contract_docx = "contract.docx"
template_docx = "template.docx"
output_pdf = "contract_with_strikethroughs.pdf"

# Step 1: Convert both PDFs to DOCX
def convert_pdf_to_docx(pdf_path, docx_path):
    cv = Converter(pdf_path)
    cv.convert(docx_path, start=0, end=None)
    cv.close()

convert_pdf_to_docx(contract_pdf, contract_docx)
convert_pdf_to_docx(template_pdf, template_docx)

# Step 2: Use OpenAI's Vision API to analyze the DOCX files
openai.api_key = 'your_openai_api_key'

def analyze_documents_with_vision(contract_path, template_path):
    with open(contract_path, "rb") as contract_file, open(template_path, "rb") as template_file:
        # Sending files to OpenAI Vision API for comparison (pseudocode, as actual Vision API capabilities are limited)
        response = openai.Image.create_edit(
            images=[contract_file, template_file],
            instructions="Analyze the template file and apply matching strikethrough formatting to the contract document. Return this in a structured response that can lists sections and characters that need to be strikethrough."
        )
        return response['edited_document']

# Step 3: Apply the edits received from the Vision API response
def apply_edits_to_docx(contract_docx, edit_instructions):
    doc = docx.Document(contract_docx)
    
    # Assuming edit_instructions is a structured response (e.g., list of sections with strikethroughs)
    for instruction in edit_instructions:
        # Locate text based on instruction and apply strikethrough
        for paragraph in doc.paragraphs:
            if instruction['text'] in paragraph.text:
                for run in paragraph.runs:
                    if instruction['text'] in run.text:
                        run.font.strike = True

    # Save the edited document
    edited_docx = "edited_contract.docx"
    doc.save(edited_docx)
    return edited_docx

# Step 4: Convert the edited DOCX back to PDF
def convert_docx_to_pdf(docx_path, pdf_path):
    pypandoc.convert_file(docx_path, 'pdf', outputfile=pdf_path)

# Execution flow
try:
    # Analyze the documents with the Vision API and retrieve edit instructions
    edit_instructions = analyze_documents_with_vision(contract_docx, template_docx)
    
    # Apply the edits to the DOCX file
    edited_contract_docx = apply_edits_to_docx(contract_docx, edit_instructions)
    
    # Convert the edited DOCX to PDF
    convert_docx_to_pdf(edited_contract_docx, output_pdf)
    print("Process completed successfully. Check the output PDF:", output_pdf)
except Exception as e:
    print("An error occurred:", e)
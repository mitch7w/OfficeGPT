from flask import Flask, request, send_file
from flask_cors import CORS
from docx import Document
from openai import OpenAI
from dotenv import load_dotenv
import io
import os
import xlsxwriter
from pptx import Presentation

# setup
load_dotenv()
OPENAI_API_KEY = os.environ.get('OPENAI_API_KEY')
app = Flask(__name__)
CORS(app)  # Enable CORS for all routes
client = OpenAI()

def openai_call(document_instructions, system_instructions):
    completion = client.chat.completions.create(
  model="gpt-3.5-turbo",
  messages=[
    {"role": "system", "content": system_instructions},
    {"role": "user", "content": document_instructions}
  ]
)
    return completion.choices[0].message.content

# creates a docx file based on the OpenAI instructions
def create_docx(document_instructions):
        doc_buffer = io.BytesIO() # Save the created document to a BytesIO buffer
        document = Document()
        system_instructions = "You are a Word document expert and create word documents using only commands from the python-docx library. Please do not return any text - only the Python code. I have already imported the necessary libraries and setup the document with document = Document() as well as document.save() so do not include these commands in your response. So I do not need any functions or other code - only the python-docx commands for creating the specific items in user's document. Thank you."
        creation_commands = openai_call(document_instructions, system_instructions) # get python-docx commands from GPT
        
        # checking for prompt injection
        exec(creation_commands) # actually run the commands that create the Word doc elements
        
        document.save(doc_buffer) # save doc to this buffer variable
        doc_buffer.seek(0)
        return doc_buffer

def create_excel(document_instructions):
    excel_buffer = io.BytesIO() # Save the Excel workbook to a BytesIO buffer
    workbook = xlsxwriter.Workbook(excel_buffer)
    system_instructions = "You are an Excel document expert and create Excel documents using only commands from the XlsxWriter library. Please do not return any text - only the Python code. I have already imported the necessary libraries and setup the document with worksheet = workbook.add_worksheet() as well as workbook.close() so do not include these commands in your response. So I do not need any functions or other code - only the XlsxWriter commands for creating the specific items in user's Excel document. Thank you."
    worksheet = workbook.add_worksheet()
    creation_commands = openai_call(document_instructions, system_instructions) # get XlsxWriter commands from GPT

    # checking for prompt injection
    exec(creation_commands) # actually run the commands that create the Word doc elements
    
    workbook.close()
    excel_buffer.seek(0)
    return excel_buffer

def create_powerpoint(document_instructions):
    ppt_buffer = io.BytesIO() # Create an in-memory buffer
    prs = Presentation()
    system_instructions = "You are a PowerPoint document expert and create PowerPoint documents using only commands from the python-pptx library. Please do not return any text - only the Python code. I have already imported the necessary libraries and setup the document with prs = Presentation() as well as prs.save(ppt_buffer) so do not include these commands in your response. So I do not need any functions or other code - only the python-pptx commands for creating the specific items in user's PowerPoint document. Thank you."
    creation_commands = openai_call(document_instructions, system_instructions) # get python-pptx commands from GPT

    # checking for prompt injection
    exec(creation_commands) # actually run the commands that create the Word doc elements

    prs.save(ppt_buffer)
    ppt_buffer.seek(0)
    return ppt_buffer

@app.route('/create', methods=['GET', 'POST'])
def create_endpoint():
    document_instructions = request.args.get('document_instructions', default='', type=str)
    selected_product = request.args.get('selected_product', default='', type=str)

    download_name = None
    mimetype = None
    doc_buffer = None

    # create the document using the correct function and send it back to the front-end
    if(selected_product == "Word"):
        doc_buffer = create_docx(document_instructions) # Generate the Word document
        download_name='created.docx'
        mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
    elif(selected_product == "Excel"):
        doc_buffer = create_excel(document_instructions) # Generate the Word document
        download_name='created.xlsx'
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    elif(selected_product == "PowerPoint"):
        doc_buffer = create_powerpoint(document_instructions) # Generate the Word document
        download_name='created.pptx'
        mimetype='application/vnd.openxmlformats-officedocument.presentationml.presentation'

    return send_file(
        doc_buffer,
        as_attachment=True,
        download_name=download_name,
        mimetype=mimetype)

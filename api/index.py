from flask import Flask, send_file
from flask_cors import CORS
from docx import Document
from openai import OpenAI
from dotenv import load_dotenv
import io
import os

# setup
load_dotenv()
OPENAI_API_KEY = os.environ.get('OPENAI_API_KEY')
app = Flask(__name__)
CORS(app)  # Enable CORS for all routes
client = OpenAI()

def openai_call():
    completion = client.chat.completions.create(
  model="gpt-3.5-turbo",
  messages=[
    {"role": "system", "content": "You are a Word document expert and create word documents using only commands from the python-docx library. Please do not return any text - only the Python code. I have already imported the necessary libraries and setup the document with document = Document() as well as document.save() so do not include these commands in your response. So I do not need any functions or other code - only the python-docx commands for creating the specific items in user's document. Thank you."},
    {"role": "user", "content": "Please create a Word document that has a heading 'New Technologies' and then an ordered list below this of new AI technologies under development. Please also add a table containing some notable AI companies as well as their estimated valuation. Please actually do research and use real-world data in the text you insert based on your own knowledge."}
  ]
)
    return completion.choices[0].message.content

# creates a docx file based on the OpenAI instructions
def create_docx():
    document = Document()

    creation_commands = openai_call() # get python-docx commands from GPT
    print("creation_commands:", creation_commands)
    # checking for prompt injection
    exec(creation_commands) # actually run the commands that create the Word doc elements
    # Save the created document to a BytesIO buffer
    doc_buffer = io.BytesIO()
    document.save(doc_buffer) # save doc to this buffer variable
    doc_buffer.seek(0)
    return doc_buffer

@app.route('/create', methods=['GET', 'POST'])
def create_endpoint():
    doc_buffer = create_docx() # Generate the Word document
     # Send the Word document back to the client as a response
    return send_file(
        doc_buffer,
        as_attachment=True,
        download_name='created.docx',
        mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
    )
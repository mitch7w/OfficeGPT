# OfficeGPT
## Creates Word, Powerpoint and Excel files for you automatically - just say what you want and out pops a pretty document! 📃

<ul>
<li>Built with Flask and GPT 🤖</li>
<li>Front-end lives at https://github.com/mitch7w/office-gpt-frontend</li>
<li>Uses ![python-docx](https://github.com/python-openxml/python-docx), [python-pptx](https://github.com/scanny/python-pptx) and [XlsxWriter](https://github.com/jmcnamara/XlsxWriter) to actually create documents</li>
<li>GPT API called in the background translate user's requests into document commands</li>
<li>import statements don't always function so well so added a second call to GPT to check quality of code -> slows things down</li>
<li>A bit buggy but as a POC it's not bad - some other really good Powerpoint creators exist already like [this](https://github.com/otahina/PowerPoint-Generator-Python-Project)</li>
<li>To use, simply clone and create a .env file with your OPENAI_API_KEY set inside</li>
</ul>

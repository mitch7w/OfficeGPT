# OfficeGPT
## Creates Word, Powerpoint and Excel files for you automatically - just say what you want and out pops a pretty document! 📃


https://github.com/mitch7w/OfficeGPT/assets/58911571/749e9272-56dd-4723-9af3-3499c67ee7dc


* Built with Flask and GPT 🤖
* Front-end lives at https://github.com/mitch7w/office-gpt-frontend
* Uses [python-docx](https://github.com/python-openxml/python-docx), [python-pptx](https://github.com/scanny/python-pptx) and [XlsxWriter](https://github.com/jmcnamara/XlsxWriter) to actually create documents
* GPT API called in the background translate user's requests into document commands
* import statements don't always function so well so added a second call to GPT to check quality of code -> slows things down
* A bit buggy but as a POC it's not bad - some other really good Powerpoint creators exist already like [this](https://github.com/otahina/PowerPoint-Generator-Python-Project)
* To use, simply clone this and the front-end, create a .env file here in the back-end with your OPENAI_API_KEY set inside and then make sure the front-end and back-end are both running. I separated them with the intention of hosting them online but as a MVP I ultimately decided to just keep it here on Github.

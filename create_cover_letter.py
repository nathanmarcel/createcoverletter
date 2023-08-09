from docx import Document
from docx2pdf import convert

def replace_placeholder(doc, old_text, new_text):
    for p in doc.paragraphs:
        if old_text in p.text:
            inline = p.runs
            for i in range(len(inline)):
                if old_text in inline[i].text:
                    text = inline[i].text.replace(old_text, new_text)
                    inline[i].text = text

def main():
    company_name = input("Enter the company name: ")
    position = input("Enter the position: ")

    # TODO: add the name of your cover letter word file here
    doc = Document('YOURCOVERLETTERGOESHERE.docx')

    # TODO: in my cover letter, I have 'INSERT_COMPANY' and 'INSERT_POSITION' to be replaced for each application
    replace_placeholder(doc, 'INSERT_COMPANY', company_name)
    replace_placeholder(doc, 'INSERT_POSITION', position)

    file_name = 'CoverLetter_' + company_name + '_' + position
    word_file = file_name + '.docx'

    # Save the modified cover letter as a new Word document
    doc.save(word_file)
    convert(word_file)



if __name__ == "__main__":
    main()
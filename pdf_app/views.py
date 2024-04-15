from django.http import HttpResponse
from pypdf import PdfWriter, PdfReader

from pathlib import Path
from pdf2docx import Converter

from pypdf import PdfWriter

from .utils import find_font_size_of_text


from spire.pdf.common import *
from spire.pdf import *


import aspose.pdf as ap



import fitz

import pdfrw


pdf_writer = PdfWriter()

pdf_path = (
    Path.home()
    / "Desktop/Hrishi/Django_projects/django_code_optimization/pdf_app/anirudh.pdf"
    # / "practice_files"
    # / "Pride_and_Prejudice.pdf"
)

def generate_pdf(request):
    # # Create a new PDF
    # pdf = PdfWriter()

    # # Add a page
    # pdf.addPage()

    # # Set content
    # pdf.setFont("Arial", 12)
    # pdf.drawString(100, 750, "Hello, World!")

    # # Save the PDF to a buffer
    # pdf_buffer = HttpResponse(content_type='application/pdf')
    # pdf.write(pdf_buffer)

    # # Create a response
    # pdf_buffer['Content-Disposition'] = 'attachment; filename="example.pdf"'
    # return pdf_buffer

    #########################################################################

    # pdf_reader = PdfReader(pdf_path)
    # length = len(pdf_reader.pages)

    # print('========== length is : ', length)

    # meta_data = pdf_reader.metadata

    # print('========= meta data is : ', meta_data)

    # title = pdf_reader.metadata.title
    
    # print('========== title is : ', title)

    # first_page = pdf_reader.pages[0]
    # page_type = type(first_page)

    # print('========= first page is : ', first_page)

    # print('=========== page type is : ', page_type)

    # print("======= text is : ", first_page.extract_text())

    # # for page in pdf_reader.pages:
    # #     print("======= ", page.extract_text())


    # txt_file = Path.home() / "Desktop/Hrishi/Django_projects/django_code_optimization/pdf_app/test.txt"
    # content = [
    #     f"{pdf_reader.metadata.title}",
    #     f"Number of pages: {len(pdf_reader.pages)}"
    # ]

    # for page in pdf_reader.pages:
    #     content.append(page.extract_text())

    # txt_file.write_text("\n".join(content))

    # page = pdf_writer.add_blank_page(width=8.27 * 72, height=11.7 * 72)

    # type(page)

    # # pdf_writer.write("blank2.pdf")

    # input_pdf = PdfReader(pdf_path)

    # first_page = input_pdf.pages[0]

    # print('===== first page is : ', first_page)

    # output_pdf = PdfWriter()
    # output_pdf.add_page(first_page)

    # output_pdf.write("first_page.pdf")

    ###################################################################

    ############**************######################################################################

    # Create an object of the PdfDocument class
    # doc = PdfDocument()
    # # # Load a PDF file
    # # doc.LoadFromFile("first_page.pdf")
    # doc.LoadFromFile("test_doc.pdf")

    # # Iterate through the pages in the document
    # for i in range(doc.Pages.Count):
    #     # Get the current page
    #     page = doc.Pages[i]    
    #     # Create an object of the PdfTextReplace class and pass the page to the constructor of the class as a parameter
    #     replacer = PdfTextReplacer(page)
        
    #     # Replace All instances of a specific text with new text
    #     # replacer.ReplaceAllText("Signedly", "replaced text tttttttttt")
    #     replacer.ReplaceAllText("newsletter", "TTTTTT")

    # #     # Replace All instances of a specific text with new text and set text color
    # #     # replacer.ReplaceAllText("PDF Editor", "PDF Editor", Color.get_Yellow())

    # # Save the resulting file
    # doc.SaveToFile("ReplaceAllFoundText.pdf")

    # doc.Close()


    ############**************######################################################################
    

    ###############-----------------------------------------###############################################

    # text_to_add = request.POST.get('textToAdd', 'hellooooooooooo')
    # x_pos = float(request.POST.get('xPos', 7))
    # y_pos = float(request.POST.get('yPos', 20))
    # page_number = request.POST.get('page_number', 0)

    # pdf_path = "first_page.pdf"

    # output_path = "modified_example.pdf"

    # # Edit PDF using PyMuPDF
    # pdf_document = fitz.open(pdf_path)
    # page = pdf_document[page_number]  # For simplicity, assuming you're editing the first page

    # text_instances = page.search_for("Signedly")
    
    # # Print text and its coordinates
    # for text, rect in zip(page.get_text().splitlines(), text_instances):
    #     print(f"Text: {text}, Position: {rect}")
    

    
    # # Add text to the specified position
    # # page.insert_text((x_pos, y_pos), text_to_add)

    # text_box = fitz.Rect(x_pos, y_pos, x_pos + 200, y_pos + 20)  # Adjust dimensions as needed
    # text_color = (0, 0, 0)

    # # Set font parameters
    # # font_name = "Arial"  # Example font name
    # font_name = "Helvetica"
    # font_size = 120  # Example font size

    # page.insert_text(((7.0, 20.093)), text_to_add)

    # # page.insert_textbox(text_box, text_to_add, fontsize=font_size, fontname=font_name, color=text_color)


    # pdf_document.save(output_path)
    # pdf_document.close()

    ###############--------------------------------###############################################

    ##################@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@############################################

    # input_pdf_path = "first_page.pdf"
    # output_pdf_path = "new.pdf"

    # pdf_document = fitz.open(input_pdf_path)
    
    # # Create a new PDF document
    # new_pdf = fitz.open()
    
    # # Iterate over pages in the existing PDF
    # for page_number in range(pdf_document.page_count):
    #     # Get the page
    #     page = pdf_document[page_number]
        
    #     # Create a new page in the new PDF and insert the existing page
    #     new_page = new_pdf.new_page(width=page.rect.width, height=page.rect.height)
    #     new_page.insert_page(page_number, page)
    
    # # Save the new PDF document
    # new_pdf.save(output_pdf_path)
    
    # # Close the PDFs
    # pdf_document.close()
    # new_pdf.close()

    ##################@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@############################################

    import PyPDF2

    input_pdf_path = "first_page.pdf"
    output_pdf_path = "output.pdf"
    search_text = "Signedly"
    replace_text = "yldengiS"

    ###############################################################

    # # Create a PDF writer object
    # pdf_writer = PyPDF2.PdfWriter()

    # with open(input_pdf_path, "rb") as input_file:
    #     # Create a PDF reader object
    #     pdf_reader = PyPDF2.PdfReader(input_file)
        
        
        
    #     # Iterate over each page in the PDF
    #     for page_number in range(len(pdf_reader.pages)):
    #         # Get the current page
    #         page = pdf_reader.pages[page_number]
            
    #         # Search for the search text and replace it with the replace text
    #         text = page.extract_text()
    #         if search_text in text:
    #             text = text.replace(search_text, replace_text)
    #             # page.mergePage(PyPDF2.PageObject.create_text_object(text))
    #             # Create a new page object
    #             new_page = PyPDF2.PageObject.create_blank_page(width=page.mediaBox.getWidth(), height=page.mediaBox.getHeight())
            
    #         # Set the text content of the new page
    #         new_page.mergePage(page)
    #         new_page.addText(new_text_content)
    
            
    #         # Add the modified page to the PDF writer
    #         pdf_writer.add_page(page)
        
    #     # Write the modified PDF to the output file
    #     with open(output_pdf_path, "wb") as output_file:
    #         pdf_writer.write(output_file)

    ##########################################################################
    # **************************************************************************

    # it works but won't work with pdfs generated from pdf
    # from .solution5 import search_and_replace_text
    
    # search_and_replace_text(search_text, replace_text)
    
    # **************************************************************************
    ###########################################################################


    ############################################################################
    #222222222222222222222222222222222222222222222222222222222222222222222222222

    # input_pdf_path = "test_doc.pdf"
    # # input_pdf_path = "first_page.pdf"
    # # output_pdf_path = "output_redacted.pdf"
    # search_text = "Your goal is to showcase"
    # replace_text = "showcase to is your goal"

    # from .pymupdf_solution2 import replace_text_with_matching_font

    # replace_text_with_matching_font(input_pdf_path, output_pdf_path, search_text, replace_text)

    #222222222222222222222222222222222222222222222222222222222222222222222222222
    ############################################################################


    ############################################################################
    #333333333333333333333333333333333333333333333333333333333333333333333333333

    from pdf2docx import Converter
    from docx.shared import Pt

    def convert_pdf_to_word(input_pdf_path, output_docx_path):
        cv = Converter(input_pdf_path)
        cv.convert(output_docx_path)
        cv.close()

    # input_pdf_path = "test_doc.pdf"
    input_pdf_path = "table.pdf"
    output_docx_path = "output.docx"
    convert_pdf_to_word(input_pdf_path, output_docx_path)

    import subprocess

    def convert_docx_to_pdf(input_docx_path, output_pdf_path):
        """
        Convert a DOCX file to a PDF file using LibreOffice.

        Parameters:
        input_docx_path (str): The path to the input DOCX file.
        output_pdf_path (str): The path where the output PDF file will be saved.
        """
        # Use subprocess to call the libreoffice command for conversion
        command = ["libreoffice", "--headless", "--convert-to", "pdf", input_docx_path, "--outdir", os.path.dirname(output_pdf_path)]
        
        try:
            subprocess.run(command, check=True)
            print(f"Successfully converted {input_docx_path} to {output_pdf_path}")
        except subprocess.CalledProcessError as e:
            print(f"Failed to convert {input_docx_path} to {output_pdf_path}: {e}")


    from docx import Document

    # def edit_word_document(input_docx_path, output_docx_path, search_text, replace_text, font_size, is_bold=False):
    #     doc = Document(input_docx_path)


    #     word_to_bold = replace_text
    #     font_size = font_size
    #     is_bold = is_bold
    #     # Iterate through each paragraph in the document
    #     for paragraph in doc.paragraphs:
 
    #         if search_text in paragraph.text:
    #             paragraph.text = paragraph.text.replace(search_text, replace_text)

    #         # Create a list to hold new runs
    #         new_runs = []

    #         # Iterate through the runs in the paragraph
    #         for run in paragraph.runs:
    #             # Get the text of the run
    #             text = run.text
                
    #             # Start index for searching the word to bold
    #             start_index = 0
                
    #             # Loop through occurrences of the word to bold
    #             while True:
    #                 # Find the word in the current run text starting from the start index
    #                 start_index = text.find(word_to_bold, start_index)
                    
    #                 # If the word is not found, append the remaining text as a single run and break the loop
    #                 if start_index == -1:
    #                     new_runs.append((text, run.bold, run.font.size))
    #                     break
                    
    #                 # Append the text before the word as a new run
    #                 before_word = text[:start_index]
    #                 if before_word:
    #                     new_runs.append((before_word, run.bold, run.font.size))
                    
    #                 # Append the word to be bolded as a new run
    #                 word = text[start_index:start_index + len(word_to_bold)]
    #                 new_runs.append((word, is_bold, Pt(font_size)))
                    
    #                 # Update start index and remaining text
    #                 start_index += len(word_to_bold)
    #                 text = text[start_index:]
                
    #         # Clear the existing runs in the paragraph
    #         paragraph.clear()

    #         # Add the new runs to the paragraph with the correct formatting
    #         for run_text, bold, size in new_runs:
    #             new_run = paragraph.add_run(run_text)
    #             new_run.bold = bold
    #             new_run.font.size = size


    #     # Save the modified document
    #     doc.save(output_docx_path)

    from docx import Document
    from docx.shared import Pt

    def edit_word_document(input_docx_path, output_docx_path, search_text, replace_text, font_size, is_bold=False):
        doc = Document(input_docx_path)

        # Iterate through each paragraph in the document
        for paragraph in doc.paragraphs:
            # Create a list to hold new runs
            new_runs = []

            # Iterate through the runs in the paragraph
            for run in paragraph.runs:
                # Get the text of the run
                text = run.text
                # Start index for searching the word to bold
                start_index = 0

                # Loop through occurrences of the word to replace
                while True:
                    # Find the word in the current run text starting from the start index
                    start_index = text.find(search_text, start_index)

                    # If the word is not found, append the remaining text as a single run and break the loop
                    if start_index == -1:
                        # Append the remaining text as a new run with the current run's formatting
                        if text:
                            new_runs.append((text, run.bold, run.font.size))
                        break

                    # Append the text before the word as a new run with the current run's formatting
                    before_word = text[:start_index]
                    if before_word:
                        new_runs.append((before_word, run.bold, run.font.size))

                    # Append the replacement word as a new run with specified boldness and font size
                    replaced_word = replace_text
                    new_runs.append((replaced_word, is_bold, Pt(font_size)))

                    # Update start index and remaining text
                    start_index += len(search_text)
                    text = text[start_index:]

            # Clear the existing runs in the paragraph
            paragraph.clear()

            # Add the new runs to the paragraph with the correct formatting
            for run_text, bold, size in new_runs:
                new_run = paragraph.add_run(run_text)
                new_run.bold = bold
                new_run.font.size = size

        # Save the modified document
        doc.save(output_docx_path)



    

    input_docx_path = "output.docx"
    output_docx_path = "edited_output.docx"
    search_text = "Distance (cm)"
    replace_text = "Example 3: YASH"
    font_size, is_bold =  find_font_size_of_text(input_pdf_path, search_text)

    print('======= TESTING WHETHER BOLD OR NOT : ', is_bold)

    edit_word_document(input_docx_path, output_docx_path, search_text, replace_text, font_size, is_bold)

    # Usage
    input_docx_path = "edited_output.docx"
    output_pdf_path = "output.pdf"

    convert_docx_to_pdf(input_docx_path, output_pdf_path)


    print('..............test size is : ', font_size)
    

    #333333333333333333333333333333333333333333333333333333333333333333333333333
    ############################################################################
    

    return HttpResponse('hh')





def modify_pdf(request):
    # Read an existing PDF
    # with open("example.pdf", "rb") as existing_pdf_file:
    #     existing_pdf = PdfReader(existing_pdf_file)

    #     # Create a writer
    #     writer = PdfWriter()

    #     # Add pages from existing PDF
    #     for page_num in range(len(existing_pdf.pages)):
    #         writer.addPage(existing_pdf.pages[page_num])

    #     # Add a new page
    #     writer.addPage()

    #     # Set content
    #     with open("new_page_content.pdf", "rb") as new_page_content_file:
    #         new_page_content_pdf = PdfReader(new_page_content_file)
    #         page = writer.getPage(len(existing_pdf.pages))
    #         page.mergePage(new_page_content_pdf.pages[0])

    #     # Save the modified PDF to a buffer
    #     modified_pdf_buffer = HttpResponse(content_type='application/pdf')
    #     writer.write(modified_pdf_buffer)

    #     # Create a response
    #     modified_pdf_buffer['Content-Disposition'] = 'attachment; filename="modified_example.pdf"'
    #     return modified_pdf_buffer
    pass
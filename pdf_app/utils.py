import fitz  # PyMuPDF

def find_font_size_of_text(input_pdf_path, search_text):
    # Open the PDF document
    pdf_document = fitz.open(input_pdf_path)
    
    # Iterate through each page in the document
    for page_number in range(len(pdf_document)):
        page = pdf_document[page_number]
        
        # Search for the text and get its bounding boxes
        text_instances = page.search_for(search_text)
        
        # Iterate through each instance of the text
        for bbox in text_instances:
            # Get text spans in the bounding box
            text_spans = page.get_text("dict", clip=bbox)['blocks'][0]['lines'][0]['spans']
            
            print('fffffffffffffffffffffffffff text span is : ', text_spans)

            # Get the font size of the text to be replaced
            font_size = text_spans[0]['size']
            bold = text_spans[0]['flags'] & 16
            
            # Close the PDF document and return the font size
            pdf_document.close()
            return font_size, bool(bold)
    
    # Close the PDF document if no instances of the text are found
    pdf_document.close()
    
    # If no text is found, return None
    return None

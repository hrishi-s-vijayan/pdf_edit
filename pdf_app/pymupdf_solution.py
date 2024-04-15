import fitz  # PyMuPDF

def replace_text_with_matching_font(input_pdf_path, output_pdf_path, search_text, replace_text):
    # Open the PDF document
    pdf_document = fitz.open(input_pdf_path)
    
    # Iterate through each page
    for page_number in range(len(pdf_document)):
        page = pdf_document[page_number]
        
        # Search for the text and get its bounding boxes
        text_instances = page.search_for(search_text)
        
        # Iterate through each instance of the text
        for bbox in text_instances:
            # Calculate the bounding box for the replacement text
            text_x, text_y = bbox.tl.x, bbox.tl.y
            bbox_width = bbox.br.x - bbox.tl.x
            bbox_height = bbox.br.y - bbox.tl.y

            # Get text spans in the bounding box
            text_spans = page.get_text("dict", clip=bbox)['blocks'][0]['lines'][0]['spans']

            # Print text spans (you can remove or modify this part if needed)
            print('========= text spans are : ', text_spans)

            # Get the font size of the text to be replaced
            font_size = text_spans[0]['size']
            
            # Redact the text using the bounding box
            page.add_redact_annot(bbox)
            
            # Apply redactions to erase the original text
            page.apply_redactions()

            # Adjust the Y-coordinate if necessary
            adjusted_y = bbox.y0 + (bbox_height * 0.8)
            
            # Add replacement text in the same bounding box area with the same font size
            page.insert_text((bbox.x0, adjusted_y), replace_text, fontsize=font_size, color=(0, 0, 0))
    
    # Save the modified PDF with redactions and replacements
    pdf_document.save(output_pdf_path)
    pdf_document.close()

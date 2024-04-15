import PyPDF2

replacements = [('old string', 'new string')]
    pdf = PyPDF2.PdfReader(open("first_page.pdf", "rb"))

    writer = PyPDF2.PdfWriter()

    for page in pdf.pages:
        contents = page.get_contents().get_data()

        for (a, b) in replacements:
            contents = contents.replace(a.encode('utf-8'), b.encode('utf-8'))
        page.get_contents().set_data(contents)
        writer.addPage(page)
    
    with open('modified.pdf', 'wb') as f:
        writer.write(f)
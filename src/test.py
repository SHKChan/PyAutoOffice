import pdfplumber as plumber
pdf = plumber.open('files/P-50463-79.pdf')

im = pdf.pages[0].to_image(resolution=150)
im_png = im._repr_png_()
print('')
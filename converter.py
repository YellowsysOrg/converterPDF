import aspose.slides
import aspose.words
import os

class Converter:
    def __init__(self, license_path):
        self.license_path = license_path
        aspose.words.License().set_license(license_path)
        aspose.slides.License().set_license(license_path)
    def ConvertToPdf(self, input_path, output_path):
        extention = os.path.splitext(input_path)[1].lower()
        if extention == ".pptx" or extention == ".ppt":
            presentation = aspose.slides.Presentation(input_path)
            presentation.save(output_path, aspose.slides.export.SaveFormat.PDF)
        elif extention == ".docx" or extention == ".doc":
            document = aspose.words.Document(input_path)
            document.save(output_path, aspose.words.SaveFormat.PDF)


license_path = r"C:\Users\hanih\Documents\yellowsys\aspose\Aspose.TotalProductFamily.lic"
converter=Converter(license_path)
converter.ConvertToPdf("./examples/mybestdocument.docx", "./examples/mybestdocument.pdf" )
converter.ConvertToPdf(r"C:\Users\hanih\Downloads\jarvis ai privacy assistant.pptx", "./examples/assistant.pdf" )

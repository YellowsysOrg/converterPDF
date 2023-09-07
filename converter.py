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




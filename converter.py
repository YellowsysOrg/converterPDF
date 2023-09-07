from pythonnet import load
load("coreclr")
import clr
import os


# ASPOSE_DLL_DIRECTORY = r"C:\Users\hanih\Documents\yellowsys\aspose\publish"  # Replace with your directory path
# ASPOSE_DLL_DIRECTORY=r"C:\Users\hanih\Documents\yellowsys\aspose\publish"
ASPOSE_DLL_DIRECTORY=os.path.join(os.getcwd(),"publish")

dlls = [
    "Aspose.Words.dll",
    "Aspose.Cells.dll",
    "Aspose.Slides.dll",
    "Aspose.Diagram.dll",
    "Aspose.Tasks.dll",
    "Aspose.Note.dll",
    "Aspose.Imaging.dll",
    "Aspose.Pdf.dll",
    "Aspose.Email.dll"
]

for dll in dlls:
    clr.AddReference(os.path.join(ASPOSE_DLL_DIRECTORY, dll))


from Aspose.Words import Document as WordsDocument, License as WordsLicense, SaveFormat as WordsSaveFormat
from Aspose.Cells import Workbook, License as CellsLicense, SaveFormat as CellsSaveFormat
from Aspose.Slides import Presentation, License as SlidesLicense, Export
from Aspose.Diagram import Diagram, License as DiagramLicense, SaveFileFormat as DiagramSaveFileFormat
from Aspose.Tasks import Project, License as TasksLicense, Saving
from Aspose.Note import Document as NoteDocument, License as NoteLicense, SaveFormat as NoteSaveFormat
from Aspose.Imaging import Image, License as ImagingLicense, ImageOptions
from Aspose.Pdf import Document as PdfDocument, License as PdfLicense
from Aspose.Email import License as EmailLicense

class Converter:
    def __init__(self, license_path):
        self._license_path = license_path
        self.apply_license()

    def apply_license(self):
        # Words
        WordsLicense().SetLicense(self._license_path)

        # Slides
        SlidesLicense().SetLicense(self._license_path)

        # Cells
        CellsLicense().SetLicense(self._license_path)

        # Imaging
        ImagingLicense().SetLicense(self._license_path)

        # PDF
        PdfLicense().SetLicense(self._license_path)

        # Email
        EmailLicense().SetLicense(self._license_path)

        # Diagram
        DiagramLicense().SetLicense(self._license_path)

        # Tasks
        TasksLicense().SetLicense(self._license_path)

        # Note
        NoteLicense().SetLicense(self._license_path)

    def convert_to_pdf(self, input_path, output_path):
        if not input_path or not output_path:
            raise ValueError("Input or output path is null or empty.")

        if not os.path.exists(input_path):
            raise FileNotFoundError(f"File not found: {input_path}")

        file_extension = os.path.splitext(input_path)[1].lower()

        if file_extension in [".doc", ".docx"]:
            doc = WordsDocument(input_path)
            doc.Save(output_path, WordsSaveFormat.Pdf)

        elif file_extension in [".ppt", ".pptx"]:
            with Presentation(input_path) as pres:
                pres.Save(output_path, Export.SaveFormat.Pdf)

        elif file_extension in [".xls", ".xlsx"]:
            workbook = Workbook(input_path)
            workbook.Save(output_path, CellsSaveFormat.Pdf)

        elif file_extension in [".jpg", ".jpeg", ".png", ".bmp"]:
            with Image.Load(input_path) as image:
                image.Save(output_path, ImageOptions.PdfOptions())

        elif file_extension in [".pdf", ".epub"]:
            pdf_document = PdfDocument(input_path)
            pdf_document.Save(output_path)

        elif file_extension in [".vsd", ".vsdx"]:
            diagram = Diagram(input_path)
            diagram.Save(output_path, DiagramSaveFileFormat.Pdf)

        elif file_extension == ".mpp":
            project = Project(input_path)
            project.Save(output_path, Saving.SaveFileFormat.Pdf)

        elif file_extension == ".one":
            one_note = NoteDocument(input_path)
            one_note.Save(output_path, NoteSaveFormat.Pdf)

        else:
            raise ValueError(f"File format {file_extension} is not supported.")
        
converter=Converter(r"C:\Users\hanih\Documents\yellowsys\aspose\Aspose.TotalProductFamily.lic")
converter.convert_to_pdf("./examples/mybestdocument.docx", "./examples/mybestdocument6.pdf" )
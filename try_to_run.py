from converter import Converter
license_path = r"C:\Users\hanih\Documents\yellowsys\aspose\Aspose.TotalProductFamily.lic"
converter=Converter(license_path)
converter.ConvertToPdf("./examples/mybestdocument.docx", "./examples/mybestdocument.pdf" )
converter.ConvertToPdf(r"C:\Users\hanih\Downloads\jarvis ai privacy assistant.pptx", "./examples/assistant.pdf" )
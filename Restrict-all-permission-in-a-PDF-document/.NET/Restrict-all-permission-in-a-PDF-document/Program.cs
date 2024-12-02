using System.IO;
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.DocIORenderer;
using Syncfusion.Pdf;
using Syncfusion.Pdf.Security;

//Open the Word document file stream.
using (FileStream inputStream = new FileStream(Path.GetFullPath(@"../../../Data/Template.docx"), FileMode.Open, FileAccess.ReadWrite))
{
    //Load an existing Word document.
    using (WordDocument wordDocument = new WordDocument(inputStream, FormatType.Docx))
    {
        //Create an instance of DocIORenderer.
        using (DocIORenderer renderer = new DocIORenderer())
        {
            //Convert Word document into PDF document.
            using (PdfDocument pdfDocument = renderer.ConvertToPDF(wordDocument))
            {
                //Document security.
                PdfSecurity security = pdfDocument.Security;
                //Specific key size and encryption algorithm using 256-bit key in AES mode.
                security.KeySize = PdfEncryptionKeySize.Key256Bit;
                security.Algorithm = Syncfusion.Pdf.Security.PdfEncryptionAlgorithm.AES;
                security.OwnerPassword = "syncfusion";
                //It restrict printing and copying of PDF document.
                security.Permissions = ~(PdfPermissionsFlags.CopyContent | PdfPermissionsFlags.Print);
                //Save the PDF file to file system.    
                using (FileStream outputStream = new FileStream(Path.GetFullPath(@"../../../Output.pdf"), FileMode.Create, FileAccess.ReadWrite))
                {
                    pdfDocument.Save(outputStream);
                }
            }
        }
    }
}
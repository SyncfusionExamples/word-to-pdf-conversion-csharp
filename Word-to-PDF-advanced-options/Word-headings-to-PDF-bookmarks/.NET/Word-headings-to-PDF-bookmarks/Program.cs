using System.IO;
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.DocIORenderer;
using Syncfusion.Pdf;

//Open the Word document file stream.
using (FileStream inputStream = new FileStream(Path.GetFullPath(@"../../../Data/Template.docx"), FileMode.Open, FileAccess.ReadWrite))
{
    //Load an existing Word document.
    using (WordDocument wordDocument = new WordDocument(inputStream, FormatType.Docx))
    {
        //Create an instance of DocIORenderer.
        using (DocIORenderer renderer = new DocIORenderer())
        {
            //Set ExportBookmarks to preserve Word document headings as PDF bookmarks.
            renderer.Settings.ExportBookmarks = ExportBookmarkType.Headings;
            //Convert Word document into PDF document.
            using (PdfDocument pdfDocument = renderer.ConvertToPDF(wordDocument))
            {
                //Save the PDF file to file system.    
                using (FileStream outputStream = new FileStream(Path.GetFullPath(@"../../../Output.pdf"), FileMode.Create, FileAccess.ReadWrite))
                {
                    pdfDocument.Save(outputStream);
                }
            }
        }
    }
}
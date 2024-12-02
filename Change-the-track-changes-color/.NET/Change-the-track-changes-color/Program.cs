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
        //Set revision types to preserve track changes in  Word when converting to PDF.
        wordDocument.RevisionOptions.ShowMarkup = RevisionType.Deletions | RevisionType.Formatting | RevisionType.Insertions;
        //Set the color to be used for revision bars that identify document lines containing revised information.
        wordDocument.RevisionOptions.RevisionBarsColor = RevisionColor.Blue;
        //Set the color to be used for inserted content Insertion.
        wordDocument.RevisionOptions.InsertedTextColor = RevisionColor.ClassicBlue;
        //Set the color to be used for deleted content Deletion.
        wordDocument.RevisionOptions.DeletedTextColor = RevisionColor.ClassicRed;
        //Set the color to be used for content with changes of formatting properties.
        wordDocument.RevisionOptions.RevisedPropertiesColor = RevisionColor.DarkYellow;
        //Create an instance of DocIORenderer.
        using (DocIORenderer renderer = new DocIORenderer())
        {
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
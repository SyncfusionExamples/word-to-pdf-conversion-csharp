using System.IO;
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.DocIORenderer;
using Syncfusion.Drawing;
using Syncfusion.Pdf;

//Open the Word document file stream.
using (FileStream inputStream = new FileStream(Path.GetFullPath(@"../../../Data/Template.docx"), FileMode.Open, FileAccess.ReadWrite))
{
    //Load an existing Word document.
    using (WordDocument wordDocument = new WordDocument(inputStream, FormatType.Docx))
    {
        //Hook the font substitution event.
        wordDocument.FontSettings.SubstituteFont += FontSettings_SubstituteFont;
        //Create an instance of DocIORenderer.
        using (DocIORenderer renderer = new DocIORenderer())
        {
            //Convert Word document into PDF document.
            using (PdfDocument pdfDocument = renderer.ConvertToPDF(wordDocument))
            {
                //Unhook the font substitution event after converting to PDF.
                wordDocument.FontSettings.SubstituteFont -= FontSettings_SubstituteFont;
                //Save the PDF file to file system.    
                using (FileStream outputStream = new FileStream(Path.GetFullPath(@"../../../Output.pdf"), FileMode.Create, FileAccess.ReadWrite))
                {
                    pdfDocument.Save(outputStream);
                }
            }
        }
    }
}
/// <summary>
/// Sets the alternate font stream when a specified font is unavailable in the production environment.
/// </summary>
/// <param name="sender">FontSettings type of the word in which the specified font stream is used but unavailable in production environment. </param>
/// <param name="args">Retrieves the unavailable font name and receives the substitute font stream for conversion. </param>
static void FontSettings_SubstituteFont(object sender, SubstituteFontEventArgs args)
{
    //Set the alternate font when a specified font is not installed in the production environment.
    if (args.OriginalFontName == "Arial Unicode MS")
    {
        switch (args.FontStyle)
        {
            case FontStyle.Italic:
                args.AlternateFontStream = new FileStream(Path.GetFullPath(@"../../../Data/Arial_italic.TTF"), FileMode.Open, FileAccess.ReadWrite);
                break;
            case FontStyle.Bold:
                args.AlternateFontStream = new FileStream(Path.GetFullPath(@"../../../Data/Arial_bold.TTF"), FileMode.Open, FileAccess.ReadWrite);
                break;
            default:
                args.AlternateFontStream = new FileStream(Path.GetFullPath(@"../../../Data/Arial.TTF"), FileMode.Open, FileAccess.ReadWrite);
                break;
        }
    }
}

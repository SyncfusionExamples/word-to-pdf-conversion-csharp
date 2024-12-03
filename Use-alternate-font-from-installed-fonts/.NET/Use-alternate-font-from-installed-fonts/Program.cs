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
        //Hook the font substitution event to handle unavailable fonts.
        //This event will be triggered when a font used in the document is not found in the production environment.
        wordDocument.FontSettings.SubstituteFont += FontSettings_SubstituteFont;
        //Create an instance of DocIORenderer.
        using (DocIORenderer renderer = new DocIORenderer())
        {
            //Convert Word document into PDF document.
            using (PdfDocument pdfDocument = renderer.ConvertToPDF(wordDocument))
            {
                //Unhook the font substitution event after the conversion is complete.
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
/// Sets the alternate font when a specified font is unavailable in the production environment.
/// </summary>
/// <param name="sender">FontSettings type of the word in which the specified font is used but unavailable in production environment. </param>
/// <param name="args">Retrieves the unavailable font name and receives the substitute font name for conversion. </param>
static void FontSettings_SubstituteFont(object sender, SubstituteFontEventArgs args)
{
    //Check if the original font is "Arial Unicode MS" and substitute with "Arial".
    if (args.OriginalFontName == "Arial Unicode MS")
        args.AlternateFontName = "Arial";
    else
        //Subsitutue "Times New Roman" for any other missing fonts.
        args.AlternateFontName = "Times New Roman";
}
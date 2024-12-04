using Syncfusion.DocIO.DLS;
using Syncfusion.DocIORenderer;
using Syncfusion.Pdf;
using Syncfusion.Office;
using Syncfusion.DocIO;

//Open the Word document file stream.
using (FileStream inputStream = new FileStream(Path.GetFullPath(@"../../../Data/Template.docx"), FileMode.Open, FileAccess.ReadWrite))
{
    //Load an existing Word document file stream.
    using (WordDocument wordDocument = new WordDocument(inputStream, FormatType.Docx))
    {
        //Add fallback font for "Arabic" script type.
        wordDocument.FontSettings.FallbackFonts.Add(ScriptType.Arabic, "Arial, Times New Roman");
        //Add fallback font for "Hebrew" script type.
        wordDocument.FontSettings.FallbackFonts.Add(ScriptType.Hebrew, "Arial, Courier New");
        //Add fallback font for "Hindi" script type.
        wordDocument.FontSettings.FallbackFonts.Add(ScriptType.Hindi, "Mangal, Nirmala UI");
        //Add fallback font for "Chinese" script type.
        wordDocument.FontSettings.FallbackFonts.Add(ScriptType.Chinese, "DengXian, MingLiU");
        //Add fallback font for "Japanese" script type.
        wordDocument.FontSettings.FallbackFonts.Add(ScriptType.Japanese, "Yu Mincho, MS Mincho");
        //Add fallback font for "Thai" script type.
        wordDocument.FontSettings.FallbackFonts.Add(ScriptType.Thai, "Tahoma, Microsoft Sans Serif");
        //Add fallback font for "Korean" script type.
        wordDocument.FontSettings.FallbackFonts.Add(ScriptType.Korean, "Malgun Gothic, Batang");
        //Instantiation of DocIORenderer for Word to PDF conversion.
        using (DocIORenderer render = new DocIORenderer())
        {
            //Convert Word document into PDF document.
            using (PdfDocument pdfDocument = render.ConvertToPDF(wordDocument))
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

using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.DocIORenderer;
using Syncfusion.Office;
using Syncfusion.Pdf;

//Open the Word document file stream.
using (FileStream inputStream = new FileStream(Path.GetFullPath(@"../../../Data/Template.docx"), FileMode.Open, FileAccess.ReadWrite))
{
    //Load an existing Word document file stream.
    using (WordDocument wordDocument = new WordDocument(inputStream, FormatType.Docx))
    {
        //Add fallback font for "Arabic" specific unicode range.
        wordDocument.FontSettings.FallbackFonts.Add(new FallbackFont(0x0600, 0x06ff, "Arial"));
        //Add fallback font for "Hebrew" specific unicode range.
        wordDocument.FontSettings.FallbackFonts.Add(new FallbackFont(0x0590, 0x05ff, "Times New Roman"));
        //Add fallback font for "Hindi" specific unicode range.
        wordDocument.FontSettings.FallbackFonts.Add(new FallbackFont(0x0900, 0x097F, "Nirmala UI"));
        //Add fallback font for "Chinese" specific unicode range.
        wordDocument.FontSettings.FallbackFonts.Add(new FallbackFont(0x4E00, 0x9FFF, "DengXian"));
        //Add fallback font for "Japanese" specific unicode range.
        wordDocument.FontSettings.FallbackFonts.Add(new FallbackFont(0x3040, 0x309F, "MS Gothic"));
        //Add fallback font for "Thai" specific unicode range.
        wordDocument.FontSettings.FallbackFonts.Add(new FallbackFont(0x0E00, 0x0E7F, "Tahoma"));
        //Add fallback font for "Korean" specific unicode range.
        wordDocument.FontSettings.FallbackFonts.Add(new FallbackFont(0xAC00, 0xD7A3, "Malgun Gothic"));
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

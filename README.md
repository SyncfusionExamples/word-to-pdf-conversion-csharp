# Convert Word document to PDF with advanced options in C#

This repository contains examples that illustrates how to convert Word documents to PDFs with advanced options programmatically in C#. The [.NET Word Library](https://www.syncfusion.com/document-sdk/net-word-library) (DocIO) converts a Word document to PDF with just five lines of code and also it does not require Microsoft Word application to be installed in the machine. It preserves the original appearance of the Word document in the converted PDF.

<p align="center"> 
<img src="Images/Convert-Word-to-PDF.png" alt="Convert-Word-to-PDF-in-Word-library"/> 
</p>

## Key Features

- **Word to PDF conversions**
  - [Convert Word to PDF](https://github.com/SyncfusionExamples/word-to-pdf-conversion-csharp/blob/master/Word-to-PDF-conversions/Convert-Word-to-PDF) - Convert Word document to PDF.
  - [Accessible PDF document](https://github.com/SyncfusionExamples/word-to-pdf-conversion-csharp/blob/master/Word-to-PDF-conversions/Accessible-PDF-document) - Convert Word document to PDF/UA (Section 508 compliant).
  - [PDF conformance level](https://github.com/SyncfusionExamples/word-to-pdf-conversion-csharp/blob/master/Word-to-PDF-conversions/PDF-conformance-level) - Convert Word document to PDF/A with various PDF conformance levels for long-term archiving and standardization.

- **Embed fonts**
  - [Embed font subset](https://github.com/SyncfusionExamples/word-to-pdf-conversion-csharp/blob/master/Embed-fonts/Embed-font-subset) - Embed only the necessary font subsets to optimize file size.
  - [Embed complete font](https://github.com/SyncfusionExamples/word-to-pdf-conversion-csharp/blob/master/Embed-fonts/Embed-complete-font) - Embed fonts within the PDF for consistent display.

- **Word to PDF advanced options**
  - [Editable PDF form fields](https://github.com/SyncfusionExamples/word-to-pdf-conversion-csharp/blob/master/Word-to-PDF-advanced-options/Editable-PDF-form-fields) - Preserve Word document form fields as PDF forms, allowing the creation of editable PDFs.
  - [Word headings to PDF bookmarks](https://github.com/SyncfusionExamples/word-to-pdf-conversion-csharp/blob/master/Word-to-PDF-advanced-options/Word-headings-to-PDF-bookmarks) - Convert Word headings to PDF bookmarks, generating PDF documents with bookmarks based on paragraph styles and outline levels.
  - [Optimize identical images](https://github.com/SyncfusionExamples/word-to-pdf-conversion-csharp/blob/master/Word-to-PDF-advanced-options/Optimize-identical-images) - Optimize identical images to reduce PDF file size.
  - [Disable alternate chunks](https://github.com/SyncfusionExamples/word-to-pdf-conversion-csharp/blob/master/Word-to-PDF-advanced-options/Disable-alternate-chunks) - Include or exclude alternate chunks during Word to PDF conversion.
  - [Complex script text](https://github.com/SyncfusionExamples/word-to-pdf-conversion-csharp/blob/master/Word-to-PDF-advanced-options/Complex-script-text) - Convert complex script text accurately.
  - [Hyphenation in Word-to-PDF](https://github.com/SyncfusionExamples/word-to-pdf-conversion-csharp/blob/master/Word-to-PDF-advanced-options/Hyphenation-in-Word-to-PDF) - Use custom dictionaries for text hyphenation in the converted PDF.
  - [Comments in Word-to-PDF](https://github.com/SyncfusionExamples/word-to-pdf-conversion-csharp/blob/master/Word-to-PDF-advanced-options/Comments-in-Word-to-PDF) - Toggle between preserving or excluding comments during Word to PDF conversions.
  - [Restrict permission in PDF](https://github.com/SyncfusionExamples/word-to-pdf-conversion-csharp/blob/master/Word-to-PDF-advanced-options/Restrict-permission-in-PDF) - Restrict permissions in the converted PDF for added security.

- **Preserve track changes**
  - [Track-changes-in-Word-to-PDF](https://github.com/SyncfusionExamples/word-to-pdf-conversion-csharp/blob/master/Preserve-track-changes/Track-changes-in-Word-to-PDF) - Preserve revision marks of tracked changes in the converted PDF.
  - [Change track changes color](https://github.com/SyncfusionExamples/word-to-pdf-conversion-csharp/blob/master/Preserve-track-changes/Change-track-changes-color) - Customize the color of track changes marks during conversion.
  - [Show or hide revisions in balloons](https://github.com/SyncfusionExamples/word-to-pdf-conversion-csharp/blob/master/Preserve-track-changes/Show-or-hide-revisions-in-balloons) - Show or hide revisions in balloons during conversion.

- **Fallback fonts**
  - [Default fallback fonts](https://github.com/SyncfusionExamples/word-to-pdf-conversion-csharp/blob/master/Fallback-fonts/Default-fallback-fonts) - Initialize default fallback fonts for a smoother conversion
  - [Based on script type](https://github.com/SyncfusionExamples/word-to-pdf-conversion-csharp/blob/master/Fallback-fonts/Based-on-script-type) - Set fallback fonts based on script type for unsupported glyphs.
  - [For range of Unicode](https://github.com/SyncfusionExamples/word-to-pdf-conversion-csharp/blob/master/Fallback-fonts/For-range-of-Unicode) - Set fallback fonts for characters when glyphs are not available.

- **Font substitution**
  - [Alternate font from installed](https://github.com/SyncfusionExamples/word-to-pdf-conversion-csharp/blob/master/Font-substitution/Alternate-font-from-installed) - Use alternate installed fonts when the original fonts are not available during Word to PDF conversion.
  - [Alternate font without installing](https://github.com/SyncfusionExamples/word-to-pdf-conversion-csharp/blob/master/Font-substitution/Alternate-font-without-installing) - Use alternate fonts without requiring font installation.

## .NET Word Library

The Syncfusion&reg; DocIO is a [.NET Word Library](https://www.syncfusion.com/document-sdk/net-word-library) allows you to add advanced Word document processing functionalities to any .NET application and does not require Microsoft Word application to be installed in the machine. It is a non-UI component that provides a full-fledged document instance model similar to the Microsoft Office COM libraries to iterate with the document elements explicitly and perform necessary manipulation. 

*   Support to [create Word document](https://www.syncfusion.com/document-sdk/net-word-library/create-word-documents) from scratch.
*   Support to open (read), modify and save existing Word documents.
*   Advanced [Mail merge](https://www.syncfusion.com/document-sdk/net-word-library/mail-merge) support with different data sources.
*   Ability to create or edit Word 97-2003 and later version documents (DOCX), and convert them to commonly used file formats such as [RTF](https://help.syncfusion.com/document-processing/word/conversions/rtf-conversions), [WordML](https://help.syncfusion.com/document-processing/word/conversions/word-file-formats-conversions#word-processing-xml-xml), [TXT](https://www.syncfusion.com/document-sdk/net-word-library/text-conversions), [HTML](https://www.syncfusion.com/document-sdk/net-word-library/html-conversions) and vice versa.
*   Ability to export a Word document as an [Image](https://www.syncfusion.com/document-sdk/net-word-library/word-to-image-conversion) and [PDF](https://www.syncfusion.com/document-sdk/net-word-library/word-to-pdf-conversion)
*   Ability to [merge](https://www.syncfusion.com/document-sdk/net-word-library/merge-word-documents) and [split](https://www.syncfusion.com/document-sdk/net-word-library/split-word-documents) Word documents.
*   Support to [compare](https://www.syncfusion.com/document-sdk/net-word-library/compare-word-documents) two DOCX format documents.
*   Ability to create and manipulate [charts](https://help.syncfusion.com/document-processing/word/word-library/net/working-with-charts), [Shapes](https://help.syncfusion.com/document-processing/word/word-library/net/working-with-shapes), and [Group shape](https://help.syncfusion.com/document-processing/word/word-library/net/working-with-shapes#grouping-shapes) in DOCX and WordML format documents.
*   Ability to read and write [Built-In and Custom Document Properties](https://help.syncfusion.com/document-processing/word/word-library/net/working-with-word-document#working-with-word-document-properties).
*   Ability to [find and replace](https://www.syncfusion.com/document-sdk/net-word-library/find-and-replace) text with its original formatting.
*   Ability to insert [Bookmarks](https://www.syncfusion.com/document-sdk/net-word-library/bookmark-in-word-document) and navigate corresponding bookmarks to insert, replace, and delete content.
*   Support to insert and edit the [form fields](https://www.syncfusion.com/document-sdk/net-word-library/form-filling-in-word-document).
*   Support to protect the document to [restrict access](https://www.syncfusion.com/document-sdk/net-word-library/protect-word-documents) to the elements present within the document.
*   Ability to [encrypt and decrypt](https://help.syncfusion.com/document-processing/word/word-library/net/working-with-security) Word documents.
*   Support to [insert and extract OLE](https://help.syncfusion.com/document-processing/word/word-library/net/working-with-paragraph#working-with-ole-objects) objects.
*   Support to run the DocIO applications in multi-thread and its thread safe.

Compatible Microsoft Word Versions
----------------------------------

*   Microsoft Word 97-2003
*   Microsoft Word 2007
*   Microsoft Word 2010
*   Microsoft Word 2013
*   Microsoft Word 2016
*   Microsoft Word 2019
*   Microsoft 365

Supported File Formats
----------------------

*   Creates, reads, and edits popular text file formats like [DOC](https://help.syncfusion.com/document-processing/word/conversions/word-file-formats-conversions#doc-to-docx-and-docx-to-doc), DOT, [DOCM](https://help.syncfusion.com/document-processing/word/conversions/word-file-formats-conversions#macros-docm-dotm), DOTM, [DOCX](https://help.syncfusion.com/document-processing/word/conversions/word-file-formats-conversions#word-document-docx), [DOTX](https://help.syncfusion.com/document-processing/word/conversions/word-file-formats-conversions#word-template-dotx), [HTML](https://www.syncfusion.com/document-sdk/net-word-library/html-conversions), [RTF](https://help.syncfusion.com/document-processing/word/conversions/rtf-conversions), [TXT](https://www.syncfusion.com/document-sdk/net-word-library/text-conversions), [Markdown](https://help.syncfusion.com/document-processing/word/conversions/markdown-to-word-conversion) and [XML (WordML)](https://help.syncfusion.com/document-processing/word/conversions/word-file-formats-conversions#word-processing-xml-xml).
*   Converts Word documents also to [PDF](https://www.syncfusion.com/document-sdk/net-word-library/word-to-pdf-conversionhub-docio-examples), [Image](https://www.syncfusion.com/document-sdk/net-word-library/word-to-image-conversion), and [ODT](https://help.syncfusion.com/document-processing/word/conversions/word-to-odt-conversion) files.

## How to run the examples
- Download this project to a location in your disk.
- Open the solution file using Visual Studio.
- Rebuild the solution to install the required NuGet packages.
- Run the application.

## Resources

*   **Product page:** [Syncfusion&reg; Word Framework](https://www.syncfusion.com/document-sdk/net-word-library)
*   **Documentation:** [Convert Word document to PDF using Syncfusion&reg; Word Library](https://help.syncfusion.com/document-processing/word/conversions/word-to-pdf/net/word-to-pdf)
*   **Online demo:** [Syncfusion&reg; Word Library - Online demos](https://document.syncfusion.com/demos/word/wordtopdf#/bootstrap5)
*   **GitHub Examples:** [Syncfusion&reg; Word Library examples](https://github.com/SyncfusionExamples/DocIO-Examples)
*   **Blog:** [Syncfusion&reg; Word Library - Blog](https://www.syncfusion.com/blogs/category/docio?utm_source=github&utm_medium=listing&utm_campaign=github-docio-examples)
*   **Knowledge Base:** [Syncfusion&reg; Word Library - Knowledge Base](https://www.syncfusion.com/kb/aspnetcore/docio?utm_source=github&utm_medium=listing&utm_campaign=github-docio-examples)
*   **Ebooks:** [Syncfusion&reg; Word Library - Ebooks](https://www.syncfusion.com/succinctly-free-ebooks?utm_source=nuget&utm_medium=listing&utm_campaign=aspnetcore-docio-nuget)
*   **FAQ:** [Syncfusion&reg; Word Library - FAQ](https://www.syncfusion.com/faq/?utm_source=github&utm_medium=listing&utm_campaign=github-docio-examples)

## Support and feedback
For any other queries, reach our [Syncfusion&reg; support team](https://support.syncfusion.com/?utm_source=github&utm_medium=listing&utm_campaign=github-docio-examples) or post the queries through the [community forums](https://www.syncfusion.com/forums?utm_source=github&utm_medium=listing&utm_campaign=github-docio-examples).

Request new feature through [Syncfusion&reg; feedback portal](https://www.syncfusion.com/feedback?utm_source=github&utm_medium=listing&utm_campaign=github-docio-examples).

## License
This is a commercial product and requires a paid license for possession or use. Syncfusion's licensed software, including this component, is subject to the terms and conditions of [Syncfusion's EULA](https://www.syncfusion.com/license/studio/22.2.5/syncfusion_essential_studio_eula.pdf?utm_source=github&utm_medium=listing&utm_campaign=github-docio-examples). You can purchase a licnense [here](https://www.syncfusion.com/sales/products?utm_source=github&utm_medium=listing&utm_campaign=github-docio-examples) or start a free 30-day trial [here](https://www.syncfusion.com/account/manage-trials/start-trials?utm_source=github&utm_medium=listing&utm_campaign=github-docio-examples).

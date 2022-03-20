// See https://aka.ms/new-console-template for more information
using Aspose.Cells;
using Aspose.Slides;
using Aspose.Words;

try
{
    Console.WriteLine("是否写序列号？(y/n):");
    var s = Console.ReadKey(true);
    if (s.Key == ConsoleKey.Y)
    {
        Console.WriteLine("正在写入序列号...");
        new Aspose.Words.License().SetLicense(new MemoryStream(Convert.FromBase64String("PExpY2Vuc2U+CiAgPERhdGE+CiAgICA8TGljZW5zZWRUbz5TdXpob3UgQXVuYm94IFNvZnR3YXJlIENvLiwgTHRkLjwvTGljZW5zZWRUbz4KICAgIDxFbWFpbFRvPnNhbGVzQGF1bnRlYy5jb208L0VtYWlsVG8+CiAgICA8TGljZW5zZVR5cGU+RGV2ZWxvcGVyIE9FTTwvTGljZW5zZVR5cGU+CiAgICA8TGljZW5zZU5vdGU+TGltaXRlZCB0byAxIGRldmVsb3BlciwgdW5saW1pdGVkIHBoeXNpY2FsIGxvY2F0aW9uczwvTGljZW5zZU5vdGU+CiAgICA8T3JkZXJJRD4yMDA2MDIwMTI2MzM8L09yZGVySUQ+CiAgICA8VXNlcklEPjEzNDk3NjAwNjwvVXNlcklEPgogICAgPE9FTT5UaGlzIGlzIGEgcmVkaXN0cmlidXRhYmxlIGxpY2Vuc2U8L09FTT4KICAgIDxQcm9kdWN0cz4KICAgICAgPFByb2R1Y3Q+QXNwb3NlLlRvdGFsIGZvciAuTkVUPC9Qcm9kdWN0PgogICAgPC9Qcm9kdWN0cz4KICAgIDxFZGl0aW9uVHlwZT5FbnRlcnByaXNlPC9FZGl0aW9uVHlwZT4KICAgIDxTZXJpYWxOdW1iZXI+OTM2ZTVmZDEtODY2Mi00YWJmLTk1YmQtYzhkYzBmNTNhZmE2PC9TZXJpYWxOdW1iZXI+CiAgICA8U3Vic2NyaXB0aW9uRXhwaXJ5PjIwMjEwODI3PC9TdWJzY3JpcHRpb25FeHBpcnk+CiAgICA8TGljZW5zZVZlcnNpb24+My4wPC9MaWNlbnNlVmVyc2lvbj4KICAgIDxMaWNlbnNlSW5zdHJ1Y3Rpb25zPmh0dHBzOi8vcHVyY2hhc2UuYXNwb3NlLmNvbS9wb2xpY2llcy91c2UtbGljZW5zZTwvTGljZW5zZUluc3RydWN0aW9ucz4KICA8L0RhdGE+CiAgPFNpZ25hdHVyZT5wSkpjQndRdnYxV1NxZ1kyOHFJYUFKSysvTFFVWWRrQ2x5THE2RUNLU0xDQ3dMNkEwMkJFTnh5L3JzQ1V3UExXbjV2bTl0TDRQRXE1aFAzY2s0WnhEejFiK1JIWTBuQkh1SEhBY01TL1BSeEJES0NGbWg1QVFZRTlrT0FxSzM5NVBSWmJRSGowOUNGTElVUzBMdnRmVkp5cUhjblJvU3dPQnVqT1oyeDc4WFE9PC9TaWduYXR1cmU+CjwvTGljZW5zZT4=")));
    }
    Console.WriteLine("正在进行文档转换...");
    var excelFile = Path.Combine(System.AppDomain.CurrentDomain.BaseDirectory, "TestFiles\\test_excel.xlsx");
    var pptFile = Path.Combine(System.AppDomain.CurrentDomain.BaseDirectory, "TestFiles\\test_ppt.pptx");
    var wordFile = Path.Combine(System.AppDomain.CurrentDomain.BaseDirectory, "TestFiles\\test_word.docx");
    OfficeFileToPDF(excelFile);
    OfficeFileToPDF(pptFile);
    OfficeFileToPDF(wordFile);
    Console.WriteLine("success!");
}
catch (Exception e)
{
    Console.WriteLine(e.Message);
}
finally
{
    Console.ReadKey();
}

void OfficeFileToPDF(string fileName)
{
    var extension = Path.GetExtension(fileName);
    var pdfFileName = Path.ChangeExtension(fileName, ".pdf");

    if (extension.ToLower().Contains(".xls"))
    {
        Workbook book = new Workbook(fileName);
        var options = new PdfSaveOptions();
        options.OnePagePerSheet = true;
        book.Save(pdfFileName, options);
    }
    else if (extension.ToLower().Contains(".doc"))
    {
        var doc = new Document(fileName);
        doc.Save(pdfFileName);
    }
    else if (extension.ToLower().Contains(".ppt"))
    {
        var doc = new Presentation(fileName);
        doc.Save(pdfFileName, Aspose.Slides.Export.SaveFormat.Pdf);
    }
}
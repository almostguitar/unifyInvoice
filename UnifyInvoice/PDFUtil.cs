using iTextSharp.text;
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.converter;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace UnifyInvoice
{
    public class PDFUtil
    {



        public static byte[] getPDFbytes(string pathFile)
        {

            try
            {

                //string pdfFilePath = "c:/pdfdocuments/myfile.pdf";


                string pdfFilePath = pathFile;
                byte[] bytes = System.IO.File.ReadAllBytes(pdfFilePath);

                return bytes;

            }
            catch (Exception ex)
            {

                throw new Exception(ex.ToString());
            }
        }


        public static void CreateMergedPDF(string targetPDF, params string[] sourceDir)
        {
            List<byte[]> s = new List<byte[]>();

            foreach (var item in sourceDir)
            {
                s.Add(getPDFbytes(item));
            }

            var pdfbytes = MergeFiles(s);

            System.IO.File.WriteAllBytes(targetPDF, pdfbytes);

        }



        public static bool CreateMergedPDF(string targetpathPDF, List<byte[]> byteFiles, bool isFirstElemtIsDestiny, out string status)
        {
            status = "ok";

            try
            {
                byte[] pdfbytesFinal = null;

                if (isFirstElemtIsDestiny)
                {
                    pdfbytesFinal = MergeFiles(byteFiles, true);
                }
                else
                {
                    if (System.IO.File.Exists(targetpathPDF))
                    {
                        var destinyBytes = PDFUtil.getPDFbytes(targetpathPDF);

                        byteFiles = byteFiles.Prepend(destinyBytes).ToList();
                    }

                    pdfbytesFinal = MergeFiles(byteFiles, true);
                }

                System.IO.File.WriteAllBytes(targetpathPDF, pdfbytesFinal);

            }
            catch (Exception ex)
            {
                status = string.Format(@"Se produjo un error al intentar unir los bytes al archivo destino: {0} --> {1}", targetpathPDF, ex.ToString());
            }

            return status == "ok";
        }

        public static bool CreatePdf(string filepath, byte[] content)
        {
            try
            {


                PdfReader reader = new PdfReader(new MemoryStream(content));
                PdfStamper stamper = new PdfStamper(reader, new FileStream(filepath, FileMode.Create));

                return true;
            }
            catch (Exception ex)
            {

                return false;
            }
        }

        public static byte[] MergeFiles(List<byte[]> sourceFiles, bool isInGrayScale = false)
        {

            PdfContentToGrayscaleConverter grayscaleConverter = null;


            Document document = new Document();
            using (MemoryStream ms = new MemoryStream())
            {
                PdfCopy copy = new PdfCopy(document, ms);
                document.Open();

                // Iterate through all pdf documents
                for (int fileCounter = 0; fileCounter < sourceFiles.Count; fileCounter++)
                {
                    // Create pdf reader
                    PdfReader reader = new PdfReader(sourceFiles[fileCounter]);
                    int numberOfPages = reader.NumberOfPages;

                    grayscaleConverter = new PdfContentToGrayscaleConverter();

                    // Iterate through all pages
                    for (int currentPageIndex = 1; currentPageIndex <= numberOfPages; currentPageIndex++)
                    {
                        // >>> CONVERT CURRENT PAGE TO GRAYSCALE
                        if (isInGrayScale)
                        {
                            grayscaleConverter.Convert(reader, currentPageIndex);
                        }
                        // <<<<

                        PdfImportedPage importedPage = copy.GetImportedPage(reader, currentPageIndex);
                        PdfCopy.PageStamp pageStamp = copy.CreatePageStamp(importedPage);

                        pageStamp.AlterContents();

                        copy.AddPage(importedPage);
                    }

                    copy.FreeReader(reader);
                    reader.Close();
                }

                document.Close();
                return ms.GetBuffer();
            }
        }


        public static byte[] getTiffToPDFBytes(string filepath)
        {

            using (MemoryStream ms = new MemoryStream())
            {
                Document document = new Document(PageSize.LETTER, 0, 0, 0, 0);
                var writer = PdfWriter.GetInstance(document, ms);
                var bitmap = new System.Drawing.Bitmap(filepath);
                var pages = bitmap.GetFrameCount(System.Drawing.Imaging.FrameDimension.Page);

                document.Open();
                iTextSharp.text.pdf.PdfContentByte cb = writer.DirectContent;
                for (int i = 0; i < pages; ++i)
                {
                    bitmap.SelectActiveFrame(System.Drawing.Imaging.FrameDimension.Page, i);
                    iTextSharp.text.Image img = iTextSharp.text.Image.GetInstance(bitmap, System.Drawing.Imaging.ImageFormat.Bmp);
                    // scale the image to fit in the page 
                    //img.ScalePercent(72f / img.DpiX * 100);
                    img.ScaleAbsolute(document.PageSize.Width, document.PageSize.Height);
                    img.SetAbsolutePosition(0, 0);
                    cb.AddImage(img);
                    document.NewPage();
                }

                document.Close();

                return ms.GetBuffer();
            }

        }


    }
}

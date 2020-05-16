using Microsoft.Office.Interop.Excel;
using stdole;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace UnifyInvoice.BO
{



    [ComVisible(true)]
    public interface IInvoice
    {

        void UnifyInvoiceProcess();
    }


    public enum InvoiceFileType
    {
        UNDEFINED = -1,
        PDF = 0,
        TIFF = 1,
        NO_SUPPORT = 99,
    }

    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    public class Invoice : IInvoice
    {


        //public Invoice(string no_Cliente,
        // string nombre_CLiente,
        // string fecha_Factura,
        // string serie,
        // string factura,
        // string imagen_PDF,
        // string resultado,
        // string imagen_Fnal)
        //{ 

        //}

        public string No_Cliente { get; set; }
        public string Nombre_CLiente { get; set; }
        public string Fecha_Factura { get; set; }
        public string Serie { get; set; }
        public string Factura { get; set; }
        public string Ingreso_Copiado_MXN { get; set; }
        public string Imagen_PDF { get; set; }
        public string Resultado { get; set; }
        public string Imagen_Fnal { get; set; }

        public byte[] bytesFile { get; set; }

        public string status { get; set; }
        public int rowIndexSheet { get; set; }


       

        public void RecoveryBytesFile()
        {
            try
            {
                if (!string.IsNullOrEmpty(this.Imagen_PDF))
                {
                    if (System.IO.File.Exists(this.Imagen_PDF))
                    {
                        this.getFileType();
                        switch (this.FileType)
                        {
                            case InvoiceFileType.UNDEFINED:
                                break;
                            case InvoiceFileType.PDF:
                                this.bytesFile = PDFUtil.getPDFbytes(this.Imagen_PDF);
                                break;
                            case InvoiceFileType.TIFF:
                                this.bytesFile = PDFUtil.getTiffToPDFBytes(this.Imagen_PDF);
                                break;
                            case InvoiceFileType.NO_SUPPORT:
                                break;
                            default:
                                break;
                        }
                    }
                    else
                    {
                        this.Resultado = string.Format(@"archivo se encontro  el archivo {0}", this.Imagen_PDF);
                    }
                }
            }
            catch (Exception ex)
            {

                this.Resultado = "se profujo una excepcion al tratar de leer el archcivo " + this.Imagen_PDF + " " + ex.ToString();

                throw new Exception(this.Resultado);
            }
        }


        public void RecoveryBytesFile(string path)
        {
            if (!string.IsNullOrEmpty(path))
            {
                if (System.IO.Directory.Exists(path))
                {
                    //this.bytesFile = PDFUtil.getPDFbytes(path);
                    this.getFileType(path);

                    switch (this.FileType)
                    {
                        case InvoiceFileType.UNDEFINED:
                            break;
                        case InvoiceFileType.PDF:
                            this.bytesFile = PDFUtil.getPDFbytes(this.Imagen_PDF);
                            break;
                        case InvoiceFileType.TIFF:
                            this.bytesFile = PDFUtil.getTiffToPDFBytes(this.Imagen_PDF);
                            break;
                        case InvoiceFileType.NO_SUPPORT:
                            break;
                        default:
                            break;
                    }
                }
            }


        }


        public InvoiceFileType FileType { get; set; }


        public InvoiceFileType getFileType()
        {
            return getFileType(this.Imagen_PDF);
        }
        public InvoiceFileType getFileType(string fileName)
        {

            if (!string.IsNullOrEmpty(fileName))
            {
                if (this.Imagen_PDF.ToLower().EndsWith(".pdf"))
                {
                    this.FileType = InvoiceFileType.PDF;
                }
                else if (this.Imagen_PDF.ToLower().EndsWith(".tiff") || this.Imagen_PDF.ToLower().EndsWith(".tif"))
                {
                    this.FileType = InvoiceFileType.TIFF;
                }
                else
                {
                    this.FileType = InvoiceFileType.NO_SUPPORT;
                }
            }
            else
            {
                this.FileType = InvoiceFileType.UNDEFINED;
            }

            return this.FileType;
        }

        public List<Invoice> MergeInvoidePDF(List<Invoice> invoices, Invoice destinyInvoicePDFFile)
        {
            string status = "";
            List<Invoice> lstInvoiceResult = null;

            invoices.ForEach((p) =>
            {
                p.RecoveryBytesFile();
            });

            lstInvoiceResult = invoices.Where(p => p.bytesFile != null && p.bytesFile.Length > 1).ToList();

            var ss = lstInvoiceResult.Select(x => x.bytesFile).ToList();

            if (ss.Count > 0)
            {
                PDFUtil.CreateMergedPDF(destinyInvoicePDFFile.Imagen_Fnal, ss, false, out status);

                lstInvoiceResult.ForEach(p => p.status = status);
            }

            return lstInvoiceResult;
        }

        internal List<Invoice> ProceessInvoices(List<Invoice> invoices)
        {
            List<Invoice> lstInvoiceProcess = new List<Invoice>();

            var groupInvoice = (from c in invoices
                                group c by new { c.Factura } into d
                                select new
                                {
                                    d.Key.Factura
                                }).ToList();



            foreach (var item in groupInvoice)
            {
                var _invoiceList = invoices.Where(p => p.Factura == item.Factura).ToList();

                var _finalInvoice = _invoiceList.FirstOrDefault(p => !string.IsNullOrEmpty(p.Imagen_Fnal) && !string.IsNullOrEmpty(p.Imagen_PDF));

                //_finalInvoice.RecoveryBytesFile(_finalInvoice.Imagen_Fnal);

                if (_finalInvoice != null)
                {
                    lstInvoiceProcess.AddRange(MergeInvoidePDF(_invoiceList, _finalInvoice));
                }
            }

            return lstInvoiceProcess;
        }



        public void UnifyInvoiceProcess()
        {

            Microsoft.Office.Interop.Excel.Worksheet activeWorksheet = Globals.ThisAddIn.Application.ActiveSheet as Microsoft.Office.Interop.Excel.Worksheet;

            if (activeWorksheet != null)
            {
                UnifyInvoice(ref activeWorksheet);
            }

        }


        public void UnifyInvoice(ref Worksheet currentWorlsheet)
        {
            int maxRowEmpty = 3;
            try
            {
                Invoice invoice = new Invoice();
                List<Invoice> invoices = new List<Invoice>();


                int i = 0;
                int j = 2;
                while (i < 3)
                {

                    var value_ = Convert.ToString(currentWorlsheet.Range[$"A{j}"].Value);

                    if (value_ == null || value_ == "")
                    {
                        i++;
                    }
                    else
                    {
                        i = 0;
                    }

                    if (value_ != null && value_ != "" && i < maxRowEmpty)
                    {
                        if (Convert.ToString(currentWorlsheet.Range[$"J{j}"].Value) != "ok")
                        {
                            Invoice invoice_ = new Invoice()
                            {
                                rowIndexSheet = j,
                                No_Cliente = Convert.ToString(currentWorlsheet.Range[$"A{j}"].Value),
                                Nombre_CLiente = Convert.ToString(currentWorlsheet.Range[$"B{j}"].Value),
                                Fecha_Factura = Convert.ToString(currentWorlsheet.Range[$"C{j}"].Value),
                                Serie = Convert.ToString(currentWorlsheet.Range[$"D{j}"].Value),
                                Factura = Convert.ToString(currentWorlsheet.Range[$"E{j}"].Value),
                                Ingreso_Copiado_MXN = Convert.ToString(currentWorlsheet.Range[$"F{j}"].Value),
                                Imagen_PDF = Convert.ToString(currentWorlsheet.Range[$"G{j}"].Value),
                                Resultado = Convert.ToString(currentWorlsheet.Range[$"H{j}"].Value),
                                Imagen_Fnal = Convert.ToString(currentWorlsheet.Range[$"I{j}"].Value),
                                status = Convert.ToString(currentWorlsheet.Range[$"J{j}"].Value),
                            };

                            invoices.Add(invoice_);

                        }
                    }

                    j++;
                }


                invoices = invoice.ProceessInvoices(invoices);


                foreach (var item in invoices)
                {
                    currentWorlsheet.Range[$"J{item.rowIndexSheet}"].Value = item.status;
                }

            }
            catch (Exception)
            {

                throw;
            }
        }




    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Ribbon;
using UnifyInvoice.BO;
using UnifyInvoice.Properties;

namespace UnifyInvoice
{
    public partial class UnifyInvoicePDF
    {
        private void UnifyInvoicePDF_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void gpUnifyPDF_DialogLauncherClick(object sender, RibbonControlEventArgs e)
        {


        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {

            //Worksheet currentWorlsheet = Globals.ThisAddIn.Application.ActiveSheet;

            //currentWorlsheet.Range[$"A1"].Value = "Factura";

            //for (int i = 2; i <= 21; i++)
            //{
            //    currentWorlsheet.Range[$"A{i}"].Value = "FDFDF012" + (i % 2 == 0 ? 1.ToString() : 2.ToString());
            //}
            //currentWorlsheet.Columns.AutoFit();

        }




        private void btnUnifyInvoice_Click(object sender, RibbonControlEventArgs e)
        {

            ProcessInvoice();

        }

        private static void SetCursorToWaiting()
        {
            Microsoft.Office.Interop.Excel.Application application = Globals.ThisAddIn.Application;
            application.Cursor = XlMousePointer.xlWait;
        }
        private static void SetCursorToDefault()
        {
            Microsoft.Office.Interop.Excel.Application application = Globals.ThisAddIn.Application;
            application.Cursor = XlMousePointer.xlDefault;
        }
        public static void ProcessInvoice()
        {

            Worksheet currentWorlsheet = Globals.ThisAddIn.Application.ActiveSheet;

            try
            {

                SetCursorToWaiting();

                new Invoice().UnifyInvoice(ref currentWorlsheet);

                //Settings.Default.Save();

                MessageBox.Show("Proceso finalizado");

            }
            catch (Exception ex)
            {

                Worksheet nextWokShhet = Globals.ThisAddIn.Application.ActiveSheet;
                if (nextWokShhet == null)
                {


                    nextWokShhet.Range[$"X1"].Value = "Error";
                    nextWokShhet.Range[$"X2"].Value = ex.ToString();
                }

                MessageBox.Show("se producjo un Error");

            }
            finally
            {

                SetCursorToDefault();
            }
        }

        private void button1_Click_1(object sender, RibbonControlEventArgs e)
        {
            frmParamUnifyInvoice frmParamUnifyInvoice = new frmParamUnifyInvoice();

            if (frmParamUnifyInvoice.ShowDialog() == DialogResult.Yes) { 
            
            
            }

        }
    }

}

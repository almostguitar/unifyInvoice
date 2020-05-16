namespace UnifyInvoice
{
    partial class UnifyInvoicePDF : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public UnifyInvoicePDF()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Component Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.rhtab = this.Factory.CreateRibbonTab();
            this.gpUnifyPDF = this.Factory.CreateRibbonGroup();
            this.btnUnifyInvoice = this.Factory.CreateRibbonButton();
            this.btnUnifyForRange = this.Factory.CreateRibbonButton();
            this.rhtab.SuspendLayout();
            this.gpUnifyPDF.SuspendLayout();
            this.SuspendLayout();
            // 
            // rhtab
            // 
            this.rhtab.Groups.Add(this.gpUnifyPDF);
            this.rhtab.Label = "Ricoh";
            this.rhtab.Name = "rhtab";
            // 
            // gpUnifyPDF
            // 
            this.gpUnifyPDF.Items.Add(this.btnUnifyInvoice);
            this.gpUnifyPDF.Items.Add(this.btnUnifyForRange);
            this.gpUnifyPDF.Label = "Facturas";
            this.gpUnifyPDF.Name = "gpUnifyPDF";
            // 
            // btnUnifyInvoice
            // 
            this.btnUnifyInvoice.Label = "Unificar Facturas";
            this.btnUnifyInvoice.Name = "btnUnifyInvoice";
            this.btnUnifyInvoice.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnUnifyInvoice_Click);
            // 
            // btnUnifyForRange
            // 
            this.btnUnifyForRange.Label = "Unificar por Rangos";
            this.btnUnifyForRange.Name = "btnUnifyForRange";
            this.btnUnifyForRange.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button1_Click_1);
            // 
            // UnifyInvoicePDF
            // 
            this.Name = "UnifyInvoicePDF";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.rhtab);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.UnifyInvoicePDF_Load);
            this.rhtab.ResumeLayout(false);
            this.rhtab.PerformLayout();
            this.gpUnifyPDF.ResumeLayout(false);
            this.gpUnifyPDF.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab rhtab;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup gpUnifyPDF;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnUnifyInvoice;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnUnifyForRange;
    }

    partial class ThisRibbonCollection
    {
        internal UnifyInvoicePDF UnifyInvoicePDF
        {
            get { return this.GetRibbon<UnifyInvoicePDF>(); }
        }
    }
}

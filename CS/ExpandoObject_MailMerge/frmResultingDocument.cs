using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using DevExpress.XtraRichEdit.API.Native;


namespace ExpandoObject_MailMerge
{
    public partial class frmResultingDocument : Form
    {
        public frmResultingDocument()
        {
            InitializeComponent();
            //richEditControl1.Views.PrintLayoutView.ZoomFactor = 0.5f;
        }

        public Document Document { get { return richEditControl1.Document; } } 
    }
}

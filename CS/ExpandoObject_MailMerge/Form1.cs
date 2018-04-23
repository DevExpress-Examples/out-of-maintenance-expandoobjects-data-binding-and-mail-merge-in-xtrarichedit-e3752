using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml.Linq;
using System.Dynamic;
using DevExpress.XtraRichEdit.API.Native;

namespace ExpandoObject_MailMerge
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();

            ribbonControl1.SelectedPage = mailingsRibbonPage1;

            dynamic weathers = GetExpandoFromXml("weather.xml");
            richEditControl1.Options.MailMerge.DataSource = weathers;


            richEditControl1.LoadDocumentTemplate("weather_report.rtf");
            richEditControl1.Options.MailMerge.ViewMergedData = true;
        }


        public static IList<dynamic> GetExpandoFromXml(String file)
        {           
            var weathers = new List<dynamic>();

            var doc = XDocument.Load(file);
            var nodes = from node in doc.Root.Descendants("weather")
                        select node;
            foreach (var n in nodes) {
                dynamic MyData = new ExpandoObject();
                MyData.LastUpdateTime = String.Format("{0:o}",DateTime.Now);
                MyData.Weather = new ExpandoObject();
                foreach (var child in n.Descendants()) {

                    var w = MyData.Weather as IDictionary<String, object>;
                    XAttribute atb = child.Attribute("data");
                    if (atb != null)
                        w[child.Name.LocalName] = atb.Value;

                }

                weathers.Add(MyData);

                }
            return weathers;
    }


        private void barButtonMergeToNewDocument_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            MailMergeOptions options = richEditControl1.Document.CreateMailMergeOptions();
            options.MergeMode = MergeMode.NewSection;
            frmResultingDocument form = new frmResultingDocument();
            richEditControl1.Document.MailMerge(options,form.Document);
            form.Show(this);
        }

    }
}

Imports Microsoft.VisualBasic
Imports System
Imports System.Collections.Generic
Imports System.ComponentModel
Imports System.Data
Imports System.Drawing
Imports System.Linq
Imports System.Text
Imports System.Windows.Forms
Imports System.Xml.Linq
Imports System.Dynamic
Imports DevExpress.XtraRichEdit.API.Native

Namespace ExpandoObject_MailMerge
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()

			ribbonControl1.SelectedPage = mailingsRibbonPage1

			Dim weathers As Object = GetExpandoFromXml("weather.xml")
			richEditControl1.Options.MailMerge.DataSource = weathers


			richEditControl1.LoadDocumentTemplate("weather_report.rtf")
			richEditControl1.Options.MailMerge.ViewMergedData = True
		End Sub


        Public Shared Function GetExpandoFromXml(ByVal file As String) As IList(Of Object)
            Dim weathers = New List(Of Object)()

            Dim doc = XDocument.Load(file)
            Dim nodes = _
             From node In doc.Root.Descendants("weather") _
             Select node
            For Each n In nodes
                Dim MyData As Object = New ExpandoObject()
                MyData.LastUpdateTime = String.Format("{0:o}", DateTime.Now)
                MyData.Weather = New ExpandoObject()
                For Each child In n.Descendants()

                    Dim w = TryCast(MyData.Weather, IDictionary(Of String, Object))
                    Dim atb As XAttribute = child.Attribute("data")
                    If atb IsNot Nothing Then
                        w(child.Name.LocalName) = atb.Value
                    End If

                Next child

                weathers.Add(MyData)

            Next n
            Return weathers
        End Function


		Private Sub barButtonMergeToNewDocument_ItemClick(ByVal sender As Object, ByVal e As DevExpress.XtraBars.ItemClickEventArgs) Handles barButtonMergeToNewDocument.ItemClick
			Dim options As MailMergeOptions = richEditControl1.Document.CreateMailMergeOptions()
			options.MergeMode = MergeMode.NewSection
			Dim form As New frmResultingDocument()
			richEditControl1.Document.MailMerge(options,form.Document)
			form.Show(Me)
		End Sub

	End Class
End Namespace

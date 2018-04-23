Imports Microsoft.VisualBasic
Imports System
Imports System.Collections.Generic
Imports System.ComponentModel
Imports System.Data
Imports System.Drawing
Imports System.Linq
Imports System.Text
Imports System.Windows.Forms
Imports DevExpress.XtraRichEdit.API.Native


Namespace ExpandoObject_MailMerge
	Partial Public Class frmResultingDocument
		Inherits Form
		Public Sub New()
			InitializeComponent()
			'richEditControl1.Views.PrintLayoutView.ZoomFactor = 0.5f;
		End Sub

		Public ReadOnly Property Document() As Document
			Get
				Return richEditControl1.Document
			End Get
		End Property
	End Class
End Namespace

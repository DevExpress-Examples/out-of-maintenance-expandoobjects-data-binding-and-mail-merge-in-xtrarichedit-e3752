<!-- default file list -->
*Files to look at*:

* [Form1.cs](./CS/ExpandoObject_MailMerge/Form1.cs) (VB: [Form1.vb](./VB/ExpandoObject_MailMerge/Form1.vb))
<!-- default file list end -->
# ExpandoObjects - data binding and mail merge in XtraRichEdit


<p>This example illustrates that .NET 4.0 dynamic types (ExpandoObject) can be bound to XtraRichEdit and used for mail merge. <br />
RichEditControl has this capability starting from the 11.2.8 version. <br />
The example shows how to load XML data form a file into a List<dynamic> containing ExpandoObject instances. Sample data are weather reports obtained via Google weather service. <br />
You can use hierarchically organized dynamic types for mail merge. This example creates data composed of the current time stamp and weather data loaded from the file. Nested weather data are specified as the MERGEFIELD field parameter using dot notation.</p>

<br/>



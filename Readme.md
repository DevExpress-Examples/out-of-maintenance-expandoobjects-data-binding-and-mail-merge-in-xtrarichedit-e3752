<!-- default badges list -->
![](https://img.shields.io/endpoint?url=https://codecentral.devexpress.com/api/v1/VersionRange/128609157/11.2.8%2B)
[![](https://img.shields.io/badge/Open_in_DevExpress_Support_Center-FF7200?style=flat-square&logo=DevExpress&logoColor=white)](https://supportcenter.devexpress.com/ticket/details/E3752)
[![](https://img.shields.io/badge/📖_How_to_use_DevExpress_Examples-e9f6fc?style=flat-square)](https://docs.devexpress.com/GeneralInformation/403183)
<!-- default badges end -->
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



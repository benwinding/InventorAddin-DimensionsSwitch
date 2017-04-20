#r "ReferenceDDLs\Autodesk.Inventor.Interop.dll"
using Inventor;
using System.Runtime.InteropServices;
using System.ComponentModel;

const string ProgId = "Inventor.Application";
Inventor.Application app = (Inventor.Application) Marshal.GetActiveObject(ProgId);

private void WriteProps(object obj)
{
    Console.WriteLine(" ----- Getting Properties ----- ");
	foreach(PropertyDescriptor descriptor in TypeDescriptor.GetProperties(obj))
	{
	    string name=descriptor.Name;
	    object value=descriptor.GetValue(obj);
	    Console.WriteLine("{0}={1}",name,value);
	}
}

var doc = app.ActiveDocument as DrawingDocument;
var selectSet = doc.SelectSet;
object first = selectSet[1];
WriteProps(first);
var ordinate = (first as OrdinateDimension);
WriteProps(ordinate.OrdinateDimensionSet);
var ordinateSet = (ordinate.OrdinateDimensionSet as OrdinateDimension);
WriteProps(ordinateSet.Members[1]);

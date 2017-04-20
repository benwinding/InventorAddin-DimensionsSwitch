#r "ReferenceDDLs\Autodesk.Inventor.Interop.dll"
using Inventor;
using System.Runtime.InteropServices;
const string ProgId = "Inventor.Application";
Inventor.Application app = (Inventor.Application) Marshal.GetActiveObject(ProgId);

var doc = app.ActiveDocument as DrawingDocument;
var selectSet = doc.SelectSet;
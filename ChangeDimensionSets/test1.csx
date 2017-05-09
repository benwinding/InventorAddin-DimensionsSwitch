#r "C:\Users\Ben\Documents\Visual Studio 2015\Projects\InventorAddInTests\InventorAddIniLogiciPartMultiEditor\ReferenceDDLs\Autodesk.Inventor.Interop.dll"
using Inventor;
using System.Runtime.InteropServices;
using System.ComponentModel;

const string ProgId = "Inventor.Application";
Inventor.Application app = (Inventor.Application)Marshal.GetActiveObject(ProgId);

private void WriteProps(object obj)
{
    Console.WriteLine(" ----- Getting Properties ----- ");
    foreach (PropertyDescriptor descriptor in TypeDescriptor.GetProperties(obj))
    {
        string name = descriptor.Name;
        object value = descriptor.GetValue(obj);
        Console.WriteLine("{0}={1}", name, value);
    }
}

var oDrawing = app.ActiveDocument as DrawingDocument;

var oAssm = app.ActiveDocument as AssemblyDocument;
var occs = oAssm.ComponentDefinition.Occurrences;
var pd = occs[1].Definition as Inventor.PartComponentDefinition;


private void WritePropsRecurs(object obj, int levelsDeep)
{
    Console.WriteLine(" ----- Getting Properties ----- ");
    CallRecursion("o", obj, levelsDeep);
}

private void CallRecursion(string level, object obj, int levelsDeep)
{
    foreach (PropertyDescriptor descriptor in TypeDescriptor.GetProperties(obj))
    {
        string name = descriptor.Name;
        object value = descriptor.GetValue(obj);
        Console.WriteLine("{0}{1}={2}", level, name, value);
        if (levelsDeep < 1)
            return;
        if (value.ToString() == "System.__ComObject")
        {
            CallRecursion(level + "-", value, levelsDeep - 1);
        }
    }
}

WritePropsRecurs(pd.Parameters, 5);
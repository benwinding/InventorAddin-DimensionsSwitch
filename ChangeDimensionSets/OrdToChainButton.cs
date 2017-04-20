using System;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Inventor;

namespace InvAddIn
{
    internal class OrdToChainButton : Button
    {
        #region "Methods"

        public OrdToChainButton(string displayName, string internalName, CommandTypesEnum commandType, string clientId, string description, string tooltip, Icon standardIcon, Icon largeIcon, ButtonDisplayEnum buttonDisplayType)
            : base(displayName, internalName, commandType, clientId, description, tooltip, standardIcon, largeIcon, buttonDisplayType)
        {

        }

        protected override void ButtonDefinition_OnExecute(NameValueMap context)
        {
            //check to make sure a sketch is active
            var drawing = InventorApplication.ActiveDocument as DrawingDocument;
            if (drawing == null)
                return;
            if (drawing.SelectSet.Count > 0)
                return;
            if (drawing.SelectSet.OfType<ChainDimensionSet>() > 0)
                return;
            //if same session, combobox definitions will already exist
            ComboBoxDefinition slotWidthComboBoxDefinition;
            slotWidthComboBoxDefinition = (Inventor.ComboBoxDefinition)InventorApplication.CommandManager.ControlDefinitions["Autodesk:SimpleAddIn:SlotWidthCboBox"];

            ComboBoxDefinition slotHeightComboBoxDefinition;
            slotHeightComboBoxDefinition = (Inventor.ComboBoxDefinition)InventorApplication.CommandManager.ControlDefinitions["Autodesk:SimpleAddIn:SlotHeightCboBox"];

            //get the selected width from combo box
            double slotWidth;
            slotWidth = slotWidthComboBoxDefinition.ListIndex;

            //get the selected height from combo box
            double slotHeight;
            slotHeight = slotHeightComboBoxDefinition.ListIndex;

            if (slotWidth > 0 && slotHeight > 0)
            {
                //draw the sketch for the slot
                PlanarSketch planarSketch;
                planarSketch = (PlanarSketch)InventorApplication.ActiveEditObject;

                SketchLine[] lines = new SketchLine[2];
                SketchArc[] arcs = new SketchArc[2];

                TransientGeometry transientGeometry;
                transientGeometry = InventorApplication.TransientGeometry;

                //start a transaction so the slot will be within a single undo step
                Transaction createSlotTransaction;
                createSlotTransaction = InventorApplication.TransactionManager.StartTransaction(InventorApplication.ActiveDocument, "Create Slot");

                //draw the lines and arcs that make up the shape of the slot
                lines[0] = planarSketch.SketchLines.AddByTwoPoints(transientGeometry.CreatePoint2d(0, 0), transientGeometry.CreatePoint2d(slotWidth, 0));
                arcs[0] = planarSketch.SketchArcs.AddByCenterStartEndPoint(transientGeometry.CreatePoint2d(slotWidth, slotHeight / 2.0), lines[0].EndSketchPoint, transientGeometry.CreatePoint2d(slotWidth, slotHeight), true);

                lines[1] = planarSketch.SketchLines.AddByTwoPoints(arcs[0].EndSketchPoint, transientGeometry.CreatePoint2d(0, slotHeight));
                arcs[1] = planarSketch.SketchArcs.AddByCenterStartEndPoint(transientGeometry.CreatePoint2d(0, slotHeight / 2.0), lines[1].EndSketchPoint, lines[0].StartSketchPoint, true);

                //create the tangent constraints between the lines and arcs
                planarSketch.GeometricConstraints.AddTangent((SketchEntity)lines[0], (SketchEntity)arcs[0], null);
                planarSketch.GeometricConstraints.AddTangent((SketchEntity)lines[1], (SketchEntity)arcs[0], null);
                planarSketch.GeometricConstraints.AddTangent((SketchEntity)lines[1], (SketchEntity)arcs[1], null);
                planarSketch.GeometricConstraints.AddTangent((SketchEntity)lines[0], (SketchEntity)arcs[1], null);

                //create a parallel constraint between the two lines
                planarSketch.GeometricConstraints.AddParallel((SketchEntity)lines[0], (SketchEntity)lines[1], true, true);

                //end the transaction
                createSlotTransaction.End();
            }

            #endregion
        }
    }
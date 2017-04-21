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

        public OrdToChainButton(string displayName, string internalName, CommandTypesEnum commandType, string clientId,
            string description, string tooltip, Icon standardIcon, Icon largeIcon, ButtonDisplayEnum buttonDisplayType)
            : base(
                displayName, internalName, commandType, clientId, description, tooltip, standardIcon, largeIcon,
                buttonDisplayType)
        {

        }

        protected override void ButtonDefinition_OnExecute(NameValueMap context)
        {
            ConvertOrdinateSetToChainSet();
        }

        private void ConvertOrdinateSetToChainSet()
        {
            var drawing = InventorApplication.ActiveDocument as DrawingDocument;
            if (drawing == null)
                return;
            if (drawing.SelectSet.Count == 0)
                return;
            var selectedLinearInChainSet = drawing.SelectSet
                .OfType<OrdinateDimension>()
                .Where(d => d.IsOrdinateSetMember)
                .ToList();
            if (!selectedLinearInChainSet.Any())
                return;

            var ordinateDimensionSets = selectedLinearInChainSet
                .Select(x => x.OrdinateDimensionSet)
                .Distinct().ToList();
            foreach (var ordinateDimensionSet in ordinateDimensionSets)
            {
                ProccessOrdinateSet(drawing.ActiveSheet, ordinateDimensionSet);
                ordinateDimensionSet.Delete();
            }
        }

        private void ProccessOrdinateSet(Sheet drawingActiveSheet, OrdinateDimensionSet ordinateSet)
        {
            var intents = GetIntentsFromOrdinatesSet(ordinateSet);
            var placementPoint = GetPlacementPoint(ordinateSet);
            var dimensionType = ordinateSet.DimensionType;

            drawingActiveSheet.DrawingDimensions.ChainDimensionSets.Add(intents, placementPoint, dimensionType);
        }

        private Point2d GetPlacementPoint(OrdinateDimensionSet ordinateSet)
        {
            var origin = ordinateSet.Members[1];
            return origin.Text.Origin;
        }

        private ObjectCollection GetIntentsFromOrdinatesSet(OrdinateDimensionSet chainSet)
        {
            var memberCount = chainSet.Members.Count;
            var collection = InventorApplication.TransientObjects.CreateObjectCollection();
            for (int i = 1; i <= memberCount; i++)
            {
                var intent = (chainSet.Members[i] as OrdinateDimension).Intent;
                collection.Add(intent);
            }
            return collection;
        }

        #endregion
    }
}

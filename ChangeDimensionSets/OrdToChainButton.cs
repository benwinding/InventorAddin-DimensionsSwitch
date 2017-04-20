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
                .Distinct()
                .ToList();
            var ordinateCount = ordinateDimensionSets.Count;
            for (var i=0; i<ordinateCount; i++)
            {
                var ordinateDimensionSet = ordinateDimensionSets[i];
                ProccessOrdinateSet(drawing.ActiveSheet, ordinateDimensionSet);
                ordinateDimensionSet.Delete();
            }
        }

        private void ProccessOrdinateSet(Sheet drawingActiveSheet, OrdinateDimensionSet ordinateSet)
        {
            var intents = GetIntentsFromOrdinatesSet(ordinateSet);
            var placementPoint = (ordinateSet.Members[1] as OrdinateDimension).Text.Origin;
            var dimensionType = ordinateSet.DimensionType;

            drawingActiveSheet.DrawingDimensions.ChainDimensionSets.Add(intents, placementPoint, dimensionType);
        }

        private void ConvertChainSetToOrdinateSet()
        {
            var drawing = InventorApplication.ActiveDocument as DrawingDocument;
            if (drawing == null)
                return;
            if (drawing.SelectSet.Count == 0)
                return;
            var selectedLinearInChainSet = drawing.SelectSet
                .OfType<LinearGeneralDimension>()
                .Where(d => d.IsChainSetMember)
                .ToList();
            if (!selectedLinearInChainSet.Any())
                return;

            var chainSets = selectedLinearInChainSet
                .Select(x => x.ChainDimensionSet)
                .Distinct();

            foreach (var chainSet in chainSets)
            {
                ProccessChainSet(drawing.ActiveSheet, chainSet);
            }
        }

        private void ProccessChainSet(Sheet drawingActiveSheet, ChainDimensionSet chainSet)
        {
            var intents = GetIntentsFromChainSet(chainSet);
            var placementPoint = (chainSet.Members[1] as LinearGeneralDimension).Text.Origin;
            var dimensionType = chainSet.DimensionType;

            drawingActiveSheet.DrawingDimensions.OrdinateDimensionSets.Add(intents, placementPoint, dimensionType);
        }

        private ObjectCollection GetIntentsFromChainSet(ChainDimensionSet chainSet)
        {
            var memberCount = chainSet.Members.Count;
            var collection = InventorApplication.TransientObjects.CreateObjectCollection();
            for (int i = 1; i <= memberCount; i++)
            {
                var intent = (chainSet.Members[i] as LinearGeneralDimension).IntentOne;
                collection.Add(intent);
            }
            var intentLast = (chainSet.Members[memberCount] as LinearGeneralDimension).IntentTwo;
            collection.Add(intentLast);
            return collection;
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
//            var intentLast = (chainSet.Members[memberCount] as OrdinateDimension).Intent;
  //          collection.Add(intentLast);
            return collection;
        }

        #endregion
    }
}

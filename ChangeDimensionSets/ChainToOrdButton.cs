using System.Drawing;
using System.Linq;
using Inventor;

namespace InvAddIn
{
    internal class ChainToOrdButton : Button
    {
        #region "Methods"

        public ChainToOrdButton(string displayName, string internalName, CommandTypesEnum commandType, string clientId,
            string description, string tooltip, Icon standardIcon, Icon largeIcon, ButtonDisplayEnum buttonDisplayType)
            : base(
                displayName, internalName, commandType, clientId, description, tooltip, standardIcon, largeIcon,
                buttonDisplayType)
        {

        }

        protected override void ButtonDefinition_OnExecute(NameValueMap context)
        {
            ConvertChainSetToOrdinateSet();
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
                .Distinct().ToList();

            foreach (var chainSet in chainSets)
            {
                ProccessChainSet(drawing.ActiveSheet, chainSet);
                chainSet.Delete();
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

        #endregion
    }
}
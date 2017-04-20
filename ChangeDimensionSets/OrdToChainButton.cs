using System;
using System.Drawing;
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

        override protected void ButtonDefinition_OnExecute(NameValueMap context)
        {
            try
            {
                //if same session, combobox definitions will already exist
                ComboBoxDefinition slotWidthComboBoxDefinition;
                slotWidthComboBoxDefinition = (Inventor.ComboBoxDefinition)InventorApplication.CommandManager.ControlDefinitions["Autodesk:SimpleAddIn:SlotWidthCboBox"];

                ComboBoxDefinition slotHeightComboBoxDefinition;
                slotHeightComboBoxDefinition = (Inventor.ComboBoxDefinition)InventorApplication.CommandManager.ControlDefinitions["Autodesk:SimpleAddIn:SlotHeightCboBox"];
            }
            catch (Exception e)
            {
                MessageBox.Show(e.ToString());
            }
        }

        #endregion
    }
}
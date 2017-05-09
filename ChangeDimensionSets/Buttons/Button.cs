using System;
using System.Drawing;
using System.Windows.Forms;
using Inventor;
using Microsoft.VisualBasic.Compatibility.VB6;
using IPictureDisp = stdole.IPictureDisp;

namespace InvAddIn.Buttons
{
    internal abstract class Button
    {
        #region Data Members

        //private data members:
        private static Inventor.Application m_inventorApplication;

        private ButtonDefinition m_buttonDefinition;
        private ButtonDefinitionSink_OnExecuteEventHandler ButtonDefinition_OnExecuteEventDelegate;

        #endregion

        #region "Properties"

        public static Inventor.Application InventorApplication
        {
            set
            {
                m_inventorApplication = value;
            }
            get
            {
                return m_inventorApplication;
            }
        }

        public Inventor.ButtonDefinition ButtonDefinition
        {
            get
            {
                return m_buttonDefinition;
            }
        }

        #endregion

        #region "Methods"

        public Button(string displayName, string internalName, CommandTypesEnum commandType, string clientId, string description, string tooltip, Icon standardIcon, Icon largeIcon, ButtonDisplayEnum buttonDisplayType)
        {
            try
            {
                //get IPictureDisp for icons
                var standardIconIPictureDisp = (IPictureDisp) Support.IconToIPicture(standardIcon);
                var largeIconIPictureDisp = (IPictureDisp) Support.IconToIPicture(largeIcon);

                //create button definition
                m_buttonDefinition = m_inventorApplication.CommandManager.ControlDefinitions.AddButtonDefinition(displayName, internalName, commandType, clientId, description, tooltip, standardIconIPictureDisp, largeIconIPictureDisp, buttonDisplayType);

                //enable the button
                m_buttonDefinition.Enabled = true;

                //connect the button event sink
                ButtonDefinition_OnExecuteEventDelegate = ButtonDefinition_OnExecute;
                m_buttonDefinition.OnExecute += ButtonDefinition_OnExecuteEventDelegate;
            }
            catch (Exception e)
            {
                MessageBox.Show(e.ToString());
            }
        }

        public Button(string displayName, string internalName, CommandTypesEnum commandType, string clientId, string description, string tooltip, ButtonDisplayEnum buttonDisplayType)
        {
            try
            {
                //create button definition
                m_buttonDefinition = m_inventorApplication.CommandManager.ControlDefinitions.AddButtonDefinition(displayName, internalName, commandType, clientId, description, tooltip, Type.Missing, Type.Missing, buttonDisplayType);

                //enable the button
                m_buttonDefinition.Enabled = true;

                //connect the button event sink
                ButtonDefinition_OnExecuteEventDelegate = ButtonDefinition_OnExecute;
                m_buttonDefinition.OnExecute += ButtonDefinition_OnExecuteEventDelegate;
            }
            catch (Exception e)
            {
                MessageBox.Show(e.ToString());
            }
        }

        abstract protected void ButtonDefinition_OnExecute(NameValueMap context);

        #endregion
    }
}
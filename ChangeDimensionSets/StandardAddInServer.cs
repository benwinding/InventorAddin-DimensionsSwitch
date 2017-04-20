using System;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Inventor;

namespace InvAddIn
{
    /// <summary>
    /// This is the primary AddIn Server class that implements the ApplicationAddInServer interface
    /// that all Inventor AddIns are required to implement. The communication between Inventor and
    /// the AddIn is via the methods on this interface.
    /// </summary>
    [Guid("5698ae52-ba87-4b3f-b9f7-5871430abda3")]
    public class StandardAddInServer : Inventor.ApplicationAddInServer
    {

        // Inventor application object.
        private Inventor.Application m_inventorApplication;

        // Dimension switch - ribbon panel
        private RibbonPanel m_DimensionSwitch_ribbonPanel;

        // buttons
        private OrdToChainButton m_convertOrdinateToChain;

        public StandardAddInServer()
        {
        }

        #region ApplicationAddInServer Members

        public void Activate(Inventor.ApplicationAddInSite addInSiteObject, bool firstTime)
        {
            // This method is called by Inventor when it loads the addin.
            // The AddInSiteObject provides access to the Inventor Application object.
            // The FirstTime flag indicates if the addin is loaded for the first time.

            try
            {
                // Initialize AddIn members.
                m_inventorApplication = addInSiteObject.Application;
                Button.InventorApplication = m_inventorApplication;

                //retrieve the GUID for this class
                var addInClsidString = GetGUIDForThisClass();

                //create icons from files
                var ordToChainIcon = new Icon(GetType(), "ord-to-chain.ico");

                //create buttons
                var identifierPrefix = "Autodesk:ChangeDimensionSets:";
                m_convertOrdinateToChain = new OrdToChainButton(
                    "Ordinate To Chain",
                    identifierPrefix + "OrdToChainButton",
                    CommandTypesEnum.kShapeEditCmdType,
                    addInClsidString,
                    "Ordinate To Chain",
                    "Convert Ordinate To Chain Dimension",
                    ordToChainIcon,
                    ordToChainIcon,
                    ButtonDisplayEnum.kDisplayTextInLearningMode);

                //create command category
                var catName = "Dimensions";
                var commandCategory = m_inventorApplication.CommandManager.CommandCategories.Add(catName, 
                    identifierPrefix + "ChangeDimensionSetsCategory", addInClsidString);
                commandCategory.Add(m_convertOrdinateToChain.ButtonDefinition);

                if (firstTime == true)
                {
                    //access user interface manager
                    var userInterfaceManager = m_inventorApplication.UserInterfaceManager;
                    var interfaceStyle = userInterfaceManager.InterfaceStyle;

                    //create the UI for classic interface
                    if (interfaceStyle == InterfaceStyleEnum.kClassicInterface)
                    {
                        // Create command bar
                        var commandBar = userInterfaceManager.CommandBars.Add(catName,
                            identifierPrefix + "ChangeDimensionSetsToolbar", CommandBarTypeEnum.kRegularCommandBar,
                            addInClsidString);

                        //add buttons to toolbar
                        commandBar.Controls.AddButton(m_convertOrdinateToChain.ButtonDefinition);

                        //Get the 2d sketch environment base object
                        Inventor.Environment partSketchEnvironment =
                            userInterfaceManager.Environments["DLxDrawingEnvironment"];

                        //make this command bar accessible in the panel menu for the 2d sketch environment.
                        partSketchEnvironment.PanelBar.CommandBarList.Add(commandBar);
                    }
                    else
                    {
                        //get the ribbon associated with part document
                        var ribbons = userInterfaceManager.Ribbons;
                        var partRibbon = ribbons["Drawing"];

                        //get the tabls associated with part ribbon
                        var ribbonTabs = partRibbon.RibbonTabs;
                        var partSketchRibbonTab = ribbonTabs["id_TabPlaceViews"];

                        //create a new panel with the tab
                        var ribbonPanels = partSketchRibbonTab.RibbonPanels;
                        m_DimensionSwitch_ribbonPanel = ribbonPanels.Add(catName, 
                            identifierPrefix + "SlotRibbonPanel", addInClsidString);

                        //add controls to the slot panel
                        var partSketchSlotRibbonPanelCtrls = m_DimensionSwitch_ribbonPanel.CommandControls;

                        //add the buttons to the ribbon panel
                        partSketchSlotRibbonPanelCtrls.AddButton(m_convertOrdinateToChain.ButtonDefinition);
                    }
                }
            }
            catch(Exception e)
            {
                MessageBox.Show(e.Message.ToString());
                MessageBox.Show(e.StackTrace.ToString());
            }

        }

        public void Deactivate()
        {
            // This method is called by Inventor when the AddIn is unloaded.
            // The AddIn will be unloaded either manually by the user or
            // when the Inventor session is terminated
            
            // Release objects.
            m_inventorApplication = null;

            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        public void ExecuteCommand(int commandID)
        {
            // Note:this method is now obsolete, you should use the 
            // ControlDefinition functionality for implementing commands.
        }

        public object Automation
        {
            // This property is provided to allow the AddIn to expose an API 
            // of its own to other programs. Typically, this  would be done by
            // implementing the AddIn's API interface in a class and returning 
            // that class object through this property.

            get
            {
                // TODO: Add ApplicationAddInServer.Automation getter implementation
                return null;
            }
        }

        #endregion

        #region Helpers

        public string GetGUIDForThisClass()
        {
            //retrieve the GUID for this class
            GuidAttribute addInCLSID;
            addInCLSID = (GuidAttribute)GuidAttribute.GetCustomAttribute(typeof(StandardAddInServer), typeof(GuidAttribute));
            return "{" + addInCLSID.Value + "}";
        }

        #endregion

    }
}

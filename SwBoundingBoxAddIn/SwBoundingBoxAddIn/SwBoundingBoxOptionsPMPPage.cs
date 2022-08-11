using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using SolidWorks.Interop.swpublished;
using System;

namespace SwBoundingBoxAddIn
{
    class SwBoundingBoxOptionsPMPPage
    {
        //Local Objects
        IPropertyManagerPage2 swPropertyPage = null;
        SwBoundingBoxPMPEvents pmpEventHandler = null;


        ISldWorks iSwApp = null;
        SwBoundingBoxAddIn boundingBoxAddin = null;

        #region Property Manager Page Controls
        //Groups
        IPropertyManagerPageGroup pmp_BoundingBoxOpsGroup;       

               
        //checkBox controls
        public IPropertyManagerPageCheckbox cb_TighestFit;
        public IPropertyManagerPageCheckbox cb_CaluclateOnRebuild;
        public IPropertyManagerPageCheckbox cb_RecalculateOnClose;
        public IPropertyManagerPageCheckbox cb_UseDocumentLength; 

    

        //Control IDs
        public const int group_BBoxOpsID = 0;    

        
        public const int checkbox_TightFitID = 1;
        public const int checkbox_CalculateOnRebuildID = 2;
        public const int checkbox_RecalculateBeforeCloseID = 3;
        public const int checkbox_UseDocLengthID = 4;


        #endregion

        public SwBoundingBoxOptionsPMPPage(SwBoundingBoxAddIn addin)
        {
            boundingBoxAddin = addin;
            if (boundingBoxAddin != null)
            {
                iSwApp = (ISldWorks)boundingBoxAddin.SolidworksApp;
                CreatePropertyManagerPage();
            }
            else
            {
                System.Windows.Forms.MessageBox.Show("SwBoundingBox Addin is not Ready");
                
            }
        }


        protected void CreatePropertyManagerPage()
        {
            int errors = -1;
            int options = (int)swPropertyManagerPageOptions_e.swPropertyManagerOptions_OkayButton |
                (int)swPropertyManagerPageOptions_e.swPropertyManagerOptions_CancelButton;

            pmpEventHandler = new SwBoundingBoxPMPEvents(this);


            if (swPropertyPage != null && errors == (int)swPropertyManagerPageStatus_e.swPropertyManagerPage_Okay)
            {
                try
                {
                    AddControls();
                }
                catch (Exception e)
                {
                    iSwApp.SendMsgToUser2(e.Message, 0, 0);
                }
            }
        }


       /// <summary>
       /// Adds controls to the Property manager Page 
       /// </summary>
        protected void AddControls()
        {
            short controlType = -1;
            short align = -1;
            int options = -1;


            //Add the groups
            options = (int)swAddGroupBoxOptions_e.swGroupBoxOptions_Expanded |
                      (int)swAddGroupBoxOptions_e.swGroupBoxOptions_Visible;

            pmp_BoundingBoxOpsGroup = (IPropertyManagerPageGroup)swPropertyPage.AddGroupBox(group_BBoxOpsID, "BoundingBoxApp Options", options);

            options = (int)swAddGroupBoxOptions_e.swGroupBoxOptions_Checkbox |
                      (int)swAddGroupBoxOptions_e.swGroupBoxOptions_Visible;

            #region Adding Check box controls to the group 
            //checkbox1
            controlType = (int)swPropertyManagerPageControlType_e.swControlType_Checkbox;
            align = (int)swPropertyManagerPageControlLeftAlign_e.swControlAlign_LeftEdge;
            options = (int)swAddControlOptions_e.swControlOptions_Enabled |
                      (int)swAddControlOptions_e.swControlOptions_Visible;

            cb_TighestFit = (IPropertyManagerPageCheckbox)pmp_BoundingBoxOpsGroup.AddControl(checkbox_TightFitID, controlType, "Tighest fit", align, options, "Calculating using Body Extreme Points");
            cb_TighestFit.Checked = true;//Making this option selected by default

            //checkbox2
            controlType = (int)swPropertyManagerPageControlType_e.swControlType_Checkbox;
            align = (int)swPropertyManagerPageControlLeftAlign_e.swControlAlign_LeftEdge;
            options = (int)swAddControlOptions_e.swControlOptions_Enabled |
                      (int)swAddControlOptions_e.swControlOptions_Visible;

            cb_CaluclateOnRebuild = (IPropertyManagerPageCheckbox)pmp_BoundingBoxOpsGroup.AddControl(checkbox_CalculateOnRebuildID, controlType, "Recalculate upon rebuild", align, options, "Calculates on Rebuild");

            //checkbox3
            controlType = (int)swPropertyManagerPageControlType_e.swControlType_Checkbox;
            align = (int)swPropertyManagerPageControlLeftAlign_e.swControlAlign_LeftEdge;
            options = (int)swAddControlOptions_e.swControlOptions_Enabled |
                      (int)swAddControlOptions_e.swControlOptions_Visible;

            cb_RecalculateOnClose = (IPropertyManagerPageCheckbox)pmp_BoundingBoxOpsGroup.AddControl(checkbox_RecalculateBeforeCloseID, controlType, "Recalculate before closing (destroying)", align, options, "Calculates Before Closing");

            //checkbox4
            controlType = (int)swPropertyManagerPageControlType_e.swControlType_Checkbox;
            align = (int)swPropertyManagerPageControlLeftAlign_e.swControlAlign_LeftEdge;
            options = (int)swAddControlOptions_e.swControlOptions_Enabled |
                      (int)swAddControlOptions_e.swControlOptions_Visible;

            cb_UseDocumentLength = (IPropertyManagerPageCheckbox)pmp_BoundingBoxOpsGroup.AddControl(checkbox_UseDocLengthID, controlType, "Use document unit of length", align, options, "Boundng Box in Document Units");
            #endregion


        }

        public void Show()
        {
            if (swPropertyPage != null)
            {
                swPropertyPage.Show();
            }
        }
    }


    class SwBoundingBoxPMPEvents : IPropertyManagerPage2Handler8
    {
        private SwBoundingBoxOptionsPMPPage swBoundingBoxOptionsPMPPage;

        public SwBoundingBoxPMPEvents(SwBoundingBoxOptionsPMPPage swBoundingBoxOptionsPMPPage)
        {
            this.swBoundingBoxOptionsPMPPage = swBoundingBoxOptionsPMPPage;
        }

        public void AfterActivation()
        {
            
        }

        public void OnClose(int Reason)
        {
            if (Reason == (int)swPropertyManagerPageCloseReasons_e.swPropertyManagerPageClose_Okay)
            {
                if (swBoundingBoxOptionsPMPPage == null)
                    return;

                
            }
        }

        public void AfterClose()
        {
          
        }

        public bool OnHelp()
        {
            return true;
        }

        public bool OnPreviousPage()
        {
            return true;
        }

        public bool OnNextPage()
        {
            return true;
        }

        public bool OnPreview()
        {
            return true;
        }

        public void OnWhatsNew()
        {
          
        }

        public void OnUndo()
        {
            
        }

        public void OnRedo()
        {
           
        }

        public bool OnTabClicked(int Id)
        {
            return true;
        }

        public void OnGroupExpand(int Id, bool Expanded)
        {
           
        }

        public void OnGroupCheck(int Id, bool Checked)
        {
           
        }

        public void OnCheckboxCheck(int Id, bool Checked)
        {
           
        }

        public void OnOptionCheck(int Id)
        {
            
        }

        public void OnButtonPress(int Id)
        {
            
        }

        public void OnTextboxChanged(int Id, string Text)
        {
            
        }

        public void OnNumberboxChanged(int Id, double Value)
        {
            
        }

        public void OnComboboxEditChanged(int Id, string Text)
        {
            
        }

        public void OnComboboxSelectionChanged(int Id, int Item)
        {
           
        }

        public void OnListboxSelectionChanged(int Id, int Item)
        {
            
        }

        public void OnSelectionboxFocusChanged(int Id)
        {
            
        }

        public void OnSelectionboxListChanged(int Id, int Count)
        {
           
        }

        public void OnSelectionboxCalloutCreated(int Id)
        {
            
        }

        public void OnSelectionboxCalloutDestroyed(int Id)
        {
            
        }

        public bool OnSubmitSelection(int Id, object Selection, int SelType, ref string ItemText)
        {
            return true;
        }

        public int OnActiveXControlCreated(int Id, bool Status)
        {
            return 0;
        }

        public void OnSliderPositionChanged(int Id, double Value)
        {
           
        }

        public void OnSliderTrackingCompleted(int Id, double Value)
        {
            
        }

        public bool OnKeystroke(int Wparam, int Message, int Lparam, int Id)
        {
            return true;
        }

        public void OnPopupMenuItem(int Id)
        {
           
        }

        public void OnPopupMenuItemUpdate(int Id, ref int retval)
        {
            
        }

        public void OnGainedFocus(int Id)
        {
            
        }

        public void OnLostFocus(int Id)
        {
            
        }

        public int OnWindowFromHandleControlCreated(int Id, bool Status)
        {
            return 0;
        }

        public void OnListboxRMBUp(int Id, int PosX, int PosY)
        {
            
        }
    }
}

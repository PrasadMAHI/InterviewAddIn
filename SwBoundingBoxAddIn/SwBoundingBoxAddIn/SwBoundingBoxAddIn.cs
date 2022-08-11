using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Runtime.InteropServices;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

using SolidWorksTools.File;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swcommands;
using SolidWorks.Interop.swconst;
using SolidWorks.Interop.swpublished;

namespace SwBoundingBoxAddIn
{
    [ComVisible(true)]
    [Guid("7527A62E-21A5-4F35-B9A6-3043DBE1E07C")]
    [DisplayName("SwBoundingBoxApp")]
    [Description("Computing Part/Assembly bounding box")]
    public class SwBoundingBoxAddIn:ISwAddin
    {

        private SldWorks mSwApp;
        private ICommandManager mCmdMgr = null;
        private BitmapHandler mBmp;

        #region  declaration of public properties
        public SldWorks SolidworksApp
        {
            get { return mSwApp; }
        }
        #endregion



        /// <summary>
        /// This fucntion invokes, while connection to Solidworks, need to add Ui controls and events
        /// </summary>
        /// <param name="ThisSW">Solidworks Application</param>
        /// <param name="Cookie"></param>
        /// <returns></returns>
        public bool ConnectToSW(object thisSW, int cookie)
        {
            mSwApp = (SldWorks)thisSW;

            //Setup callbacks
            mSwApp.SetAddinCallbackInfo(0, this, cookie);

            #region Setup the Command Manager
            mCmdMgr = mSwApp.GetCommandManager(cookie);
            AddCommandMgr();
            #endregion

            return true;
        }

        private void AddCommandMgr()
        {
            mBmp = new BitmapHandler();
            string title = "Bounding Box", toolTip = "Calculates Bounding Box";


            int[] docTypes = new int[]{(int)swDocumentTypes_e.swDocASSEMBLY,
                                       (int)swDocumentTypes_e.swDocDRAWING,
                                       (int)swDocumentTypes_e.swDocPART};

            Assembly thisAssembly = System.Reflection.Assembly.GetAssembly(this.GetType());

            int refErrors = 0;
            ICommandGroup cmdGroup = mCmdMgr.CreateCommandGroup2(10, title, toolTip, "", -1, true, ref refErrors);
            cmdGroup.LargeIconList = mBmp.CreateFileFromResourceBitmap("SwBoundingBoxAddIn.ToolbarLarge.bmp", thisAssembly);
            cmdGroup.SmallIconList = mBmp.CreateFileFromResourceBitmap("SwBoundingBoxAddIn.ToolbarSmall.bmp", thisAssembly);
            cmdGroup.LargeMainIcon = mBmp.CreateFileFromResourceBitmap("SwBoundingBoxAddIn.MainIconLarge.bmp", thisAssembly);
            cmdGroup.SmallMainIcon = mBmp.CreateFileFromResourceBitmap("SwBoundingBoxAddIn.MainIconSmall.bmp", thisAssembly);

            int menuToolbarOption = (int)(swCommandItemType_e.swMenuItem | swCommandItemType_e.swToolbarItem);
            cmdGroup.AddCommandItem2(title, -1, "", toolTip, 0, "CalculateBoundingBox", "", 1, menuToolbarOption);

            cmdGroup.HasToolbar = true;
            cmdGroup.HasMenu = true;
            cmdGroup.Activate();

        }
        /// <summary>
        /// Invokes , while closing/disconnects from SW, Need to do some cleaning related to addin
        /// </summary>
        /// <returns></returns>
        public bool DisconnectFromSW()
        {
            mCmdMgr.RemoveCommandGroup2(10, false);
            return true;
        }
      

        #region SWAddinRegistration

        private const string ADDIN_KEY_TEMPLATE = @"SOFTWARE\SolidWorks\Addins\{{{0}}}";
        private const string ADDIN_STARTUP_KEY_TEMPLATE = @"Software\SolidWorks\AddInsStartup\{{{0}}}";
        private const string ADD_IN_TITLE_REG_KEY_NAME = "Title";
        private const string ADD_IN_DESCRIPTION_REG_KEY_NAME = "Description";

        /// <summary>
        /// Register the addin using regasm to run ROT (\codebase)
        /// </summary>
        /// <param name="addinType"></param>
        [ComRegisterFunction]
        public static void RegisterSwAddinFunction(Type addinType)
        {
            try
            {
                var addInTitle = "";
                var loadAtStartup = true;
                var addInDesc = "";

                var dispNameAtt = addinType.GetCustomAttributes(false).OfType<DisplayNameAttribute>().FirstOrDefault();

                if (dispNameAtt != null)
                {
                    addInTitle = dispNameAtt.DisplayName;
                }
                else
                {
                    addInTitle = addinType.ToString();
                }

                var descAtt = addinType.GetCustomAttributes(false).OfType<DescriptionAttribute>().FirstOrDefault();

                if (descAtt != null)
                {
                    addInDesc = descAtt.Description;
                }
                else
                {
                    addInDesc = addinType.ToString();
                }

                var addInkey = Microsoft.Win32.Registry.LocalMachine.CreateSubKey(
                    string.Format(ADDIN_KEY_TEMPLATE, addinType.GUID));

                addInkey.SetValue(null, 0);

                addInkey.SetValue(ADD_IN_TITLE_REG_KEY_NAME, addInTitle);
                addInkey.SetValue(ADD_IN_DESCRIPTION_REG_KEY_NAME, addInDesc);

                var addInStartupkey = Microsoft.Win32.Registry.CurrentUser.CreateSubKey(
                    string.Format(ADDIN_STARTUP_KEY_TEMPLATE, addinType.GUID));

                addInStartupkey.SetValue(null, Convert.ToInt32(loadAtStartup), Microsoft.Win32.RegistryValueKind.DWord);
            }
            catch (Exception ex)
            {

                Console.WriteLine("Addin Registraton Failed due to below error: \n " + ex.Message);
            }
        }

        /// <summary>
        /// Unregister the addin using regasm  (/u)
        /// </summary>
        /// <param name="addinType"></param>
        [ComUnregisterFunction]
        public static void UnregisterSwAddinFunction(Type addinType)
        {
            try
            {
                Microsoft.Win32.Registry.LocalMachine.DeleteSubKey(
                    string.Format(ADDIN_KEY_TEMPLATE, addinType.GUID));

                Microsoft.Win32.Registry.CurrentUser.DeleteSubKey(
                    string.Format(ADDIN_STARTUP_KEY_TEMPLATE, addinType.GUID));
            }
            catch (Exception e)
            {
                Console.WriteLine("Addin unregistration failed due to below error: \n" + e.Message);
            }
        }

        #endregion

        public void CalculateBoundingBox()
        {
            bool temp = false;
        }
    }
}

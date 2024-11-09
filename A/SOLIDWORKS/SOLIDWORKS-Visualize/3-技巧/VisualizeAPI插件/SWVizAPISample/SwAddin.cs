using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using SolidWorks.Interop.swpublished;
using SolidWorks.Visualize.Interfaces;
using SolidWorksTools;
using SolidWorksTools.File;

namespace SWVizAPISample
{
    /// <summary>
    /// Summary description for SWVizAPISample.
    /// </summary>
    [Guid("AFE93347-560E-470E-882C-DC705EB52DAC"), ComVisible(true)]
    [SwAddin(
        Description = "SOLIDWORKS Visualize API Sample Addin",
        Title = "SOLIDWORKS Visualize API Sample Addin",
        LoadAtStartup = true
        )]
    public class SwAddin : ISwAddin
    {
        #region Local Variables
        ISldWorks iSwApp = null;
        ICommandManager iCmdMgr = null;
        int addinID = 0;
        BitmapHandler iBmp;
        //int registerID;

        public const int mainCmdGroupID = 850629;
        public const int mainItemID1 = 750625;
        public const int mainItemID2 = 750626;
        public const int mainItemID3 = -1; // separator
        public const int mainItemID4 = 750627;
        
        public const string VisualizeAddinProgID = "SolidWorks.Visualize.Implementation.VisualizeAddin";

        string[] mainIcons = new string[6];
        string[] icons = new string[6];

        #region Event Handler Variables
        Hashtable openDocs = new Hashtable();
        SolidWorks.Interop.sldworks.SldWorks SwEventPtr = null;
        #endregion

        // Public Properties
        public ISldWorks SwApp
        {
            get { return iSwApp; }
        }
        public ICommandManager CmdMgr
        {
            get { return iCmdMgr; }
        }

        public Hashtable OpenDocs
        {
            get { return openDocs; }
        }

        #endregion

        #region SolidWorks Registration
        [ComRegisterFunctionAttribute]
        public static void RegisterFunction(Type t)
        {
            #region Get Custom Attribute: SwAddinAttribute
            SwAddinAttribute SWattr = null;
            Type type = typeof(SwAddin);

            foreach (System.Attribute attr in type.GetCustomAttributes(false).Cast<System.Attribute>())
            {
                if (attr is SwAddinAttribute)
                {
                    SWattr = attr as SwAddinAttribute;
                    break;
                }
            }

            #endregion

            try
            {
                Microsoft.Win32.RegistryKey hklm = Microsoft.Win32.Registry.LocalMachine;
                Microsoft.Win32.RegistryKey hkcu = Microsoft.Win32.Registry.CurrentUser;

                string keyname = "SOFTWARE\\SolidWorks\\SOLIDWORKS 2025\\Addins\\{" + t.GUID.ToString() + "}";
                Microsoft.Win32.RegistryKey addinkey = hklm.CreateSubKey(keyname);
                addinkey.SetValue(null, 0);

                addinkey.SetValue("Description", SWattr.Description);
                addinkey.SetValue("Title", SWattr.Title);

                keyname = "Software\\SolidWorks\\AddInsStartup\\{" + t.GUID.ToString() + "}";
                addinkey = hkcu.CreateSubKey(keyname);
                addinkey.SetValue(null, Convert.ToInt32(SWattr.LoadAtStartup), Microsoft.Win32.RegistryValueKind.DWord);
            }
            catch (System.NullReferenceException nl)
            {
                Console.WriteLine("There was a problem registering this dll: SWattr is null. \n\"" + nl.Message + "\"");
                System.Windows.Forms.MessageBox.Show("There was a problem registering this dll: SWattr is null.\n\"" + nl.Message + "\"");
            }

            catch (System.Exception e)
            {
                Console.WriteLine(e.Message);

                System.Windows.Forms.MessageBox.Show("There was a problem registering the function: \n\"" + e.Message + "\"");
            }
        }

        [ComUnregisterFunctionAttribute]
        public static void UnregisterFunction(Type t)
        {
            try
            {
                Microsoft.Win32.RegistryKey hklm = Microsoft.Win32.Registry.LocalMachine;
                Microsoft.Win32.RegistryKey hkcu = Microsoft.Win32.Registry.CurrentUser;

                string keyname = "SOFTWARE\\SolidWorks\\SOLIDWORKS 2025\\Addins\\{" + t.GUID.ToString() + "}";
                hklm.DeleteSubKey(keyname);

                keyname = "Software\\SolidWorks\\AddInsStartup\\{" + t.GUID.ToString() + "}";
                hkcu.DeleteSubKey(keyname);
            }
            catch (System.NullReferenceException nl)
            {
                Console.WriteLine("There was a problem unregistering this dll: " + nl.Message);
                System.Windows.Forms.MessageBox.Show("There was a problem unregistering this dll: \n\"" + nl.Message + "\"");
            }
            catch (System.Exception e)
            {
                Console.WriteLine("There was a problem unregistering this dll: " + e.Message);
                System.Windows.Forms.MessageBox.Show("There was a problem unregistering this dll: \n\"" + e.Message + "\"");
            }
        }

        #endregion

        #region ISwAddin Implementation
        public SwAddin()
        {
        }

        public bool ConnectToSW(object ThisSW, int cookie)
        {
            iSwApp = (ISldWorks)ThisSW;
            addinID = cookie;

            //Setup callbacks
            iSwApp.SetAddinCallbackInfo(0, this, addinID);

            #region Setup the Command Manager
            iCmdMgr = iSwApp.GetCommandManager(cookie);
            AddCommandMgr();
            #endregion

            #region Setup the Event Handlers
            SwEventPtr = (SolidWorks.Interop.sldworks.SldWorks)iSwApp;
            openDocs = new Hashtable();
            AttachEventHandlers();
            #endregion

            return true;
        }

        public bool DisconnectFromSW()
        {
            RemoveCommandMgr();
            DetachEventHandlers();

            Marshal.ReleaseComObject(iCmdMgr);
            iCmdMgr = null;
            Marshal.ReleaseComObject(iSwApp);
            iSwApp = null;
            //The addin _must_ call GC.Collect() here in order to retrieve all managed code pointers 
            GC.Collect();
            GC.WaitForPendingFinalizers();

            GC.Collect();
            GC.WaitForPendingFinalizers();

            return true;
        }
        #endregion

        #region UI Methods
        public void AddCommandMgr()
        {
            ICommandGroup cmdGroup;
            if (iBmp == null)
                iBmp = new BitmapHandler();
            Assembly thisAssembly;
            int cmdIndex0; // Simple Render
            int cmdIndex1; // Advanced Render
            int cmdIndex2; // Spacer
            int cmdIndex3; // Save
            string Title = "SOLIDWORKS Visualize API", ToolTip = "SOLIDWORKS Visualize API Sample Addin";


            int[] docTypes = new int[]{(int)swDocumentTypes_e.swDocASSEMBLY,
                                       (int)swDocumentTypes_e.swDocPART};

            thisAssembly = System.Reflection.Assembly.GetAssembly(this.GetType());


            int cmdGroupErr = 0;
            bool ignorePrevious = false;

            object registryIDs;
            //get the ID information stored in the registry
            bool getDataResult = iCmdMgr.GetGroupDataFromRegistry(mainCmdGroupID, out registryIDs);

            int[] knownIDs = { mainItemID1, mainItemID2, mainItemID3, mainItemID4 };

            if (getDataResult)
            {
                if (!CompareIDs((int[])registryIDs, knownIDs)) //if the IDs don't match, reset the commandGroup
                {
                    ignorePrevious = true;
                }
            }

            cmdGroup = iCmdMgr.CreateCommandGroup2(mainCmdGroupID, Title, ToolTip, "", -1, ignorePrevious, ref cmdGroupErr);

            // Add bitmaps to your project and set them as embedded resources or provide a direct path to the bitmaps.
            icons[0] = iBmp.CreateFileFromResourceBitmap("SWVizAPISample.Icons.VisualizeToolbarIcon_20.png", thisAssembly);
            icons[1] = iBmp.CreateFileFromResourceBitmap("SWVizAPISample.Icons.VisualizeToolbarIcon_32.png", thisAssembly);
            icons[2] = iBmp.CreateFileFromResourceBitmap("SWVizAPISample.Icons.VisualizeToolbarIcon_40.png", thisAssembly);
            icons[3] = iBmp.CreateFileFromResourceBitmap("SWVizAPISample.Icons.VisualizeToolbarIcon_64.png", thisAssembly);
            icons[4] = iBmp.CreateFileFromResourceBitmap("SWVizAPISample.Icons.VisualizeToolbarIcon_96.png", thisAssembly);
            icons[5] = iBmp.CreateFileFromResourceBitmap("SWVizAPISample.Icons.VisualizeToolbarIcon_128.png", thisAssembly);

            mainIcons[0] = iBmp.CreateFileFromResourceBitmap("SWVizAPISample.Icons.VisualizeMainIcon_20.png", thisAssembly);
            mainIcons[1] = iBmp.CreateFileFromResourceBitmap("SWVizAPISample.Icons.VisualizeMainIcon_32.png", thisAssembly);
            mainIcons[2] = iBmp.CreateFileFromResourceBitmap("SWVizAPISample.Icons.VisualizeMainIcon_40.png", thisAssembly);
            mainIcons[3] = iBmp.CreateFileFromResourceBitmap("SWVizAPISample.Icons.VisualizeMainIcon_64.png", thisAssembly);
            mainIcons[4] = iBmp.CreateFileFromResourceBitmap("SWVizAPISample.Icons.VisualizeMainIcon_96.png", thisAssembly);
            mainIcons[5] = iBmp.CreateFileFromResourceBitmap("SWVizAPISample.Icons.VisualizeMainIcon_128.png", thisAssembly);

            cmdGroup.MainIconList = mainIcons;
            cmdGroup.IconList = icons;

            int menuToolbarOption = (int)(swCommandItemType_e.swMenuItem | swCommandItemType_e.swToolbarItem);
            cmdIndex0 = cmdGroup.AddCommandItem2("简单渲染", 0, "用Visualize直接渲染保存成图片", "简单渲染", 0, "DoSimpleRender", "EnableForPartsAndAssemsInDesktop", mainItemID1, menuToolbarOption);
            cmdIndex1 = cmdGroup.AddCommandItem2("高级渲染", 1, "用Visualize设置输出图片的属性，在渲染导出", "高级渲染", 1, "DoAdvancedRender", "EnableForPartsAndAssemsInDesktop", mainItemID2, menuToolbarOption);
            cmdIndex2 = cmdGroup.AddSpacer2(2, menuToolbarOption); // mainItemID3
            cmdIndex3 = cmdGroup.AddCommandItem2("保存bc", 3, "保存成Visualize项目文件", "保存", 2, "DoSave", "EnableForPartsAndAssemsInDesktop", mainItemID4, menuToolbarOption);

            cmdGroup.HasToolbar = true;
            cmdGroup.HasMenu = true;
            cmdGroup.Activate();
            
            bool bResult;

            foreach (int type in docTypes)
            {
                CommandTab cmdTab;

                cmdTab = iCmdMgr.GetCommandTab(type, Title);

                if (cmdTab != null && (!getDataResult || ignorePrevious))//if tab exists, but we have ignored the registry info (or changed command group ID), re-create the tab.  Otherwise the ids won't matchup and the tab will be blank
                {
                    bool res = iCmdMgr.RemoveCommandTab(cmdTab);
                    cmdTab = null;
                }

                //if cmdTab is null, must be first load (possibly after reset), add the commands to the tabs
                if (cmdTab == null)
                {
                    cmdTab = iCmdMgr.AddCommandTab(type, Title);

                    CommandTabBox cmdBox = cmdTab.AddCommandTabBox();

                    int[] cmdIDs = new int[3];
                    int[] TextType = new int[3];

                    cmdIDs[0] = cmdGroup.get_CommandID(cmdIndex0);
                    TextType[0] = (int)swCommandTabButtonTextDisplay_e.swCommandTabButton_TextBelow;

                    cmdIDs[1] = cmdGroup.get_CommandID(cmdIndex1);
                    TextType[1] = (int)swCommandTabButtonTextDisplay_e.swCommandTabButton_TextBelow;

                    cmdIDs[2] = cmdGroup.get_CommandID(cmdIndex3);
                    TextType[2] = (int)swCommandTabButtonTextDisplay_e.swCommandTabButton_TextBelow;
                    
                    bResult = cmdBox.AddCommands(cmdIDs, TextType);

                    cmdTab.AddSeparator(cmdBox, cmdIDs[2]);
                }
            }

            thisAssembly = null;

        }

        public void RemoveCommandMgr()
        {
            iBmp.Dispose();

            iCmdMgr.RemoveCommandGroup(mainCmdGroupID);
            //iCmdMgr.RemoveFlyoutGroup(flyoutGroupID);
        }

        public bool CompareIDs(int[] storedIDs, int[] addinIDs)
        {
            List<int> storedList = new List<int>(storedIDs);
            List<int> addinList = new List<int>(addinIDs);

            addinList.Sort();
            storedList.Sort();

            if (addinList.Count != storedList.Count)
            {
                return false;
            }
            else
            {

                for (int i = 0; i < addinList.Count; i++)
                {
                    if (addinList[i] != storedList[i])
                    {
                        return false;
                    }
                }
            }
            return true;
        }

        #endregion

        #region UI Callbacks

        public void DoSimpleRender()
        {
            // Create a FolderBrowserDialog
            var folderDialog = new FolderBrowserDialog();

            // Show the dialog and get the result
            var result = folderDialog.ShowDialog();

            // If user selects a folder, update the OutputFolder property with the selected folder path
            if (result != DialogResult.OK)
            {
                return;
            }

            IVisualizeAddin vizAddin = SwApp.GetAddInObject(VisualizeAddinProgID);
            IVisualizeAddinManager vizAddingMgr = vizAddin?.GetAddinManager();
            if (vizAddingMgr == null)
            {
                return;
            }

            vizAddingMgr.RenderOptions.OutputFolder = folderDialog.SelectedPath;
            vizAddingMgr.Render();
        }

        public void DoAdvancedRender()
        {
            RenderDialog renderDialog = new RenderDialog("RenderDialog");
            renderDialog.ShowDialog();
            if (renderDialog.ViewModel.DialogCancel)
            {
                return;
            }

            IVisualizeAddin vizAddin = SwApp.GetAddInObject(VisualizeAddinProgID);
            if (vizAddin != null)
            {
                IVisualizeAddinManager vizAddingMgr = vizAddin.GetAddinManager();
                RenderHandler renderHandler = new RenderHandler();

                if (vizAddingMgr != null)
                {
                    IRenderOptions vizRenderOptions = vizAddingMgr.RenderOptions;
                    if (vizRenderOptions != null)
                    {
                        vizRenderOptions.ImageFormat = renderDialog.ViewModel.SelectedImageFormat;
                        vizRenderOptions.Width = (int)renderDialog.ViewModel.Width;
                        vizRenderOptions.Height = (int)renderDialog.ViewModel.Height;
                        vizRenderOptions.JobName = renderDialog.ViewModel.JobName;
                        vizRenderOptions.OutputFolder = renderDialog.ViewModel.OutputFolder;
                        vizRenderOptions.DenoiserEnabled = renderDialog.ViewModel.EnableDenoiser;
                        vizRenderOptions.IncludeAlpha = renderDialog.ViewModel.IncludeAlpha;
                        vizRenderOptions.TerminationMode = renderDialog.ViewModel.QualityMode
                            ? TerminationMode_e.QualityLevel
                            : TerminationMode_e.TimeLimit;
                        if (vizRenderOptions.TerminationMode == TerminationMode_e.QualityLevel)
                        {
                            vizRenderOptions.FrameCount = renderDialog.ViewModel.RenderPasses;
                            vizRenderOptions.Milliseconds = 0;
                            renderHandler.RenderProgressDialog.ViewModel.TotalFrames = renderDialog.ViewModel.RenderPasses;
                        }
                        else
                        {
                            vizRenderOptions.FrameCount = 0;
                            vizRenderOptions.Milliseconds = (long)renderDialog.ViewModel.TimeLimit.TotalMilliseconds;
                            renderHandler.RenderProgressDialog.ViewModel.TotalFrames = -1;
                        }

                        renderHandler.RenderProgressDialog.ViewModel.AspectRatio =
                            renderDialog.ViewModel.Width / renderDialog.ViewModel.Height;
                    }

                    vizAddingMgr.StartRender(renderHandler);

                    renderHandler.RenderProgressDialog.ShowDialog();
                }
            }
        }

        public void DoSave()
        {
            // Create an instance of SaveFileDialog
            SaveFileDialog saveFileDialog = new SaveFileDialog();

            // Set initial directory and file name filter if needed
            saveFileDialog.InitialDirectory = System.Environment.GetFolderPath(System.Environment.SpecialFolder.MyDocuments);
            saveFileDialog.FileName = GetModelName() + ".svpj";
            saveFileDialog.Filter = "SVPJ Files (*.svpj)|*.svpj|All files (*.*)|*.*";

            // Show the dialog and check if the user clicked OK
            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                string filePath = saveFileDialog.FileName; // Get the selected file path

                if (SwApp.GetAddInObject(VisualizeAddinProgID) is IVisualizeAddin vizAddin)
                {
                    IVisualizeAddinManager vizAddingMgr = vizAddin.GetAddinManager();
                    if (vizAddingMgr != null)
                    {
                        vizAddingMgr.Save(filePath);
                    }
                }
            }
        }

        public int PopupEnable()
        {
            if (iSwApp.ActiveDoc == null)
                return 0;
            else
                return 1;
        }

        public void TestCallback()
        {
            Debug.Print("Test Callback, CSharp");
        }

        public int EnableTest()
        {
            if (iSwApp.ActiveDoc == null)
                return 0;
            else
                return 1;
        }

        #endregion

        #region Event Methods
        public bool AttachEventHandlers()
        {
            AttachSwEvents();
            //Listen for events on all currently open docs
            AttachEventsToAllDocuments();
            return true;
        }

        private bool AttachSwEvents()
        {
            try
            {
                SwEventPtr.ActiveDocChangeNotify += new DSldWorksEvents_ActiveDocChangeNotifyEventHandler(OnDocChange);
                SwEventPtr.DocumentLoadNotify2 += new DSldWorksEvents_DocumentLoadNotify2EventHandler(OnDocLoad);
                SwEventPtr.FileNewNotify2 += new DSldWorksEvents_FileNewNotify2EventHandler(OnFileNew);
                SwEventPtr.ActiveModelDocChangeNotify += new DSldWorksEvents_ActiveModelDocChangeNotifyEventHandler(OnModelChange);
                SwEventPtr.FileOpenPostNotify += new DSldWorksEvents_FileOpenPostNotifyEventHandler(FileOpenPostNotify);
                return true;
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                return false;
            }
        }



        private bool DetachSwEvents()
        {
            try
            {
                SwEventPtr.ActiveDocChangeNotify -= new DSldWorksEvents_ActiveDocChangeNotifyEventHandler(OnDocChange);
                SwEventPtr.DocumentLoadNotify2 -= new DSldWorksEvents_DocumentLoadNotify2EventHandler(OnDocLoad);
                SwEventPtr.FileNewNotify2 -= new DSldWorksEvents_FileNewNotify2EventHandler(OnFileNew);
                SwEventPtr.ActiveModelDocChangeNotify -= new DSldWorksEvents_ActiveModelDocChangeNotifyEventHandler(OnModelChange);
                SwEventPtr.FileOpenPostNotify -= new DSldWorksEvents_FileOpenPostNotifyEventHandler(FileOpenPostNotify);
                return true;
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                return false;
            }

        }

        public void AttachEventsToAllDocuments()
        {
            ModelDoc2 modDoc = (ModelDoc2)iSwApp.GetFirstDocument();
            while (modDoc != null)
            {
                if (!openDocs.Contains(modDoc))
                {
                    AttachModelDocEventHandler(modDoc);
                }
                else if (openDocs.Contains(modDoc))
                {
                    bool connected = false;
                    DocumentEventHandler docHandler = (DocumentEventHandler)openDocs[modDoc];
                    if (docHandler != null)
                    {
                        connected = docHandler.ConnectModelViews();
                    }
                }

                modDoc = (ModelDoc2)modDoc.GetNext();
            }
        }

        public bool AttachModelDocEventHandler(ModelDoc2 modDoc)
        {
            if (modDoc == null)
                return false;

            DocumentEventHandler docHandler = null;

            if (!openDocs.Contains(modDoc))
            {
                switch (modDoc.GetType())
                {
                    case (int)swDocumentTypes_e.swDocPART:
                        {
                            docHandler = new PartEventHandler(modDoc, this);
                            break;
                        }
                    case (int)swDocumentTypes_e.swDocASSEMBLY:
                        {
                            docHandler = new AssemblyEventHandler(modDoc, this);
                            break;
                        }
                    case (int)swDocumentTypes_e.swDocDRAWING:
                        {
                            docHandler = new DrawingEventHandler(modDoc, this);
                            break;
                        }
                    default:
                        {
                            return false; //Unsupported document type
                        }
                }
                docHandler.AttachEventHandlers();
                openDocs.Add(modDoc, docHandler);
            }
            return true;
        }

        public bool DetachModelEventHandler(ModelDoc2 modDoc)
        {
            DocumentEventHandler docHandler;
            docHandler = (DocumentEventHandler)openDocs[modDoc];
            openDocs.Remove(modDoc);
            modDoc = null;
            docHandler = null;
            return true;
        }

        public bool DetachEventHandlers()
        {
            DetachSwEvents();

            //Close events on all currently open docs
            DocumentEventHandler docHandler;
            int numKeys = openDocs.Count;
            object[] keys = new Object[numKeys];

            //Remove all document event handlers
            openDocs.Keys.CopyTo(keys, 0);
            foreach (ModelDoc2 key in keys)
            {
                docHandler = (DocumentEventHandler)openDocs[key];
                docHandler.DetachEventHandlers(); //This also removes the pair from the hash
                docHandler = null;
            }
            return true;
        }
        #endregion

        #region Event Handlers
        //Events
        public int OnDocChange()
        {
            return 0;
        }

        public int OnDocLoad(string docTitle, string docPath)
        {
            return 0;
        }

        int FileOpenPostNotify(string FileName)
        {
            AttachEventsToAllDocuments();
            return 0;
        }

        public int OnFileNew(object newDoc, int docType, string templateName)
        {
            AttachEventsToAllDocuments();
            return 0;
        }

        public int OnModelChange()
        {
            return 0;
        }

        #endregion

        #region Helpers

        public string GetModelName() => Path.GetFileNameWithoutExtension((iSwApp.ActiveDoc as ModelDoc2).GetTitle());

        #endregion
    }
}

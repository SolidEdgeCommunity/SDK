using DemoAddIn.Properties;
using SolidEdgeFramework;
using SolidEdgeSDK.AddIn;
using SolidEdgeSDK.Extensions;
using SolidEdgeSDK.InteropServices;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;

// Please refer to Instructions.txt.
namespace DemoAddIn
{
    [Guid(MyConstants.AddInGuid), ProgId(MyConstants.AddInProgId), ComVisible(true)]
    [AddInImplementedCategory(CATID.SolidEdgeAddIn)]
    [AddInEnvironmentCategory(CATID.SEApplication)]
    [AddInEnvironmentCategory(CATID.SEAllDocumentEnvrionments)]
    [AddInCulture("en-US")]
    public class MyAddIn : SolidEdgeAddIn,
        SolidEdgeFramework.ISEAddInEventsEx2,
        SolidEdgeFramework.ISEAddInEdgeBarEventsEx,
        SolidEdgeFramework.ISEApplicationEvents,
        SolidEdgeFramework.ISEApplicationEventsEx,
        SolidEdgeFramework.ISEApplicationEventsEx2
    {
        #region SolidEdgeAddIn implementation

        public override void OnConnection(SolidEdgeFramework.SeConnectMode ConnectMode)
        {
            ComEventsManager = new ComEventsManager(this);

            // Prepend '\n' to the description to allow the addin to have its own Ribbon Tab.
            // Otherwise, addin commands will appear in the Add-Ins tab.
            AddInInstance.Description = $"\n{AddInInstance.Description}";

            // If you makes changes to your ribbon, be sure to increment the GuiVersion. That makes bFirstTime = true
            // next time Solid Edge is started. bFirstTime is used to setup the ribbon so if you make a change but don't
            // see the changes, that could be why.
            AddInInstance.GuiVersion = 2;

            // Example of how to hide addin from the Add-In Manager GUI.
            //AddInInstance.Visible = false;

            // Connect to select COM events.
            // 1) Modify class to implement desired event interface(s).
            // 2) Attach to each event set via ComEventsManager.Attach().
            ComEventsManager.Attach<SolidEdgeFramework.ISEAddInEdgeBarEventsEx>(AddInInstance);
            ComEventsManager.Attach<SolidEdgeFramework.ISEApplicationEvents>(Application);

            // Solid Edge 2020 or higher.
            if (SolidEdgeVersion.Major >= 220)
            {
                ComEventsManager.Attach<SolidEdgeFramework.ISEAddInEventsEx2>(AddInInstance);
                ComEventsManager.Attach<SolidEdgeFramework.ISEApplicationEventsEx>(Application);
                ComEventsManager.Attach<SolidEdgeFramework.ISEApplicationEventsEx2>(Application);
            }

            MyViewOverlay3D = new MyViewOverlay3D(this);
        }

        public override void OnConnectToEnvironment(Guid EnvCatID, SolidEdgeFramework.Environment environment, bool firstTime)
        {
            // If you are not seeing events for a certain environment, check your AddInEnvironmentCategory attribute(s).
            var myCommands = new MyCommand[] { };

            if (EnvCatID == new Guid(CATID.SEApplication))
            {
                myCommands = new MyCommand[]
                {
                    new MyCommand(MyCommandIds.About, Resources.BackstageCategory, null,  Resources.AboutCommand, NativeResources.PNG.Help_32x32, SeButtonStyle.seNoExitBackstageButton)
                };
            }
            else if (
                (EnvCatID == new Guid(CATID.SEAssembly)) ||
                (EnvCatID == new Guid(CATID.SEPart)) ||
                (EnvCatID == new Guid(CATID.SEDMPart)) ||
                (EnvCatID == new Guid(CATID.SESheetMetal)) ||
                (EnvCatID == new Guid(CATID.SEDMSheetMetal))
                )
            {
                myCommands = new MyCommand[]
                {
                    new MyCommand(MyCommandIds.About, Resources.BackstageCategory, null,  Resources.AboutCommand, NativeResources.PNG.Help_32x32, SeButtonStyle.seNoExitBackstageButton),
                    new MyCommand(MyCommandIds.Save, Resources.RibbonCaption, Resources.RibbonGroup1,  Resources.SaveCommand, NativeResources.PNG.Save_16x16, SeButtonStyle.seButtonAutomatic),
                    new MyCommand(MyCommandIds.Folder, Resources.RibbonCaption, Resources.RibbonGroup1,  Resources.FolderCommand, NativeResources.PNG.Folder_16x16, SeButtonStyle.seButtonAutomatic),
                    new MyCommand(MyCommandIds.Monitor, Resources.RibbonCaption, Resources.RibbonGroup1,  Resources.MonitorCommand, NativeResources.PNG.Monitor_16x16, SeButtonStyle.seButtonAutomatic),
                    new MyCommand(MyCommandIds.Box, Resources.RibbonCaption, Resources.RibbonGroup1,  Resources.BoxCommand, NativeResources.PNG.Box_32x32, SeButtonStyle.seButtonIconAndCaptionBelow),
                    new MyCommand(MyCommandIds.Camera, Resources.RibbonCaption, Resources.RibbonGroup2,  Resources.CameraCommand, NativeResources.PNG.Camera_32x32, SeButtonStyle.seButtonIconAndCaptionBelow),
                    new MyCommand(MyCommandIds.Photograph, Resources.RibbonCaption, Resources.RibbonGroup2,  Resources.PhotographCommand, NativeResources.PNG.Photograph_32x32, SeButtonStyle.seButtonIconAndCaptionBelow),
                    new MyCommand(MyCommandIds.Favorites, Resources.RibbonCaption, Resources.RibbonGroup2,  Resources.FavoritesCommand, NativeResources.PNG.Favorites_32x32, SeButtonStyle.seButtonIconAndCaptionBelow),
                    new MyCommand(MyCommandIds.Printer, Resources.RibbonCaption, Resources.RibbonGroup2,  Resources.PrinterCommand, NativeResources.PNG.Printer_32x32, SeButtonStyle.seButtonIconAndCaptionBelow),
                    new MyCommand(MyCommandIds.Tools, Resources.RibbonCaption, Resources.RibbonGroup3,  Resources.ToolsCommand, NativeResources.PNG.Tools_32x32, SeButtonStyle.seButtonIconAndCaptionBelow),
                    new MyCommand(MyCommandIds.CommandPrompt, Resources.RibbonCaption, Resources.RibbonGroup3,  Resources.CommandPromptCommand, NativeResources.PNG.CommandPrompt_32x32, SeButtonStyle.seButtonIconAndCaptionBelow),
                    new MyCommand(MyCommandIds.Notepad, Resources.RibbonCaption, Resources.RibbonGroup3,  Resources.NotepadCommand, NativeResources.PNG.Notepad_32x32, SeButtonStyle.seButtonIconAndCaptionBelow),
                    new MyCommand(MyCommandIds.Help, Resources.RibbonCaption, Resources.RibbonGroup4,  Resources.HelpCommand, NativeResources.PNG.Help_32x32, SeButtonStyle.seButtonIconAndCaptionBelow),
                    new MyCommand(MyCommandIds.Search, Resources.RibbonCaption, Resources.RibbonGroup4,  Resources.SearchCommand, NativeResources.PNG.Search_32x32, SeButtonStyle.seButtonIconAndCaptionBelow),
                    new MyCommand(MyCommandIds.Question, Resources.RibbonCaption, Resources.RibbonGroup5,  Resources.QuestionCommand, NativeResources.PNG.Question_32x32, SeButtonStyle.seButtonIconAndCaptionBelow),
                    new MyCommand(MyCommandIds.CheckBox1, Resources.RibbonCaption, Resources.RibbonGroup6,  Resources.Checkbox1Command, 0, SeButtonStyle.seCheckButton),
                    new MyCommand(MyCommandIds.CheckBox2, Resources.RibbonCaption, Resources.RibbonGroup6,  Resources.Checkbox2Command, 0, SeButtonStyle.seCheckButton),
                    new MyCommand(MyCommandIds.CheckBox3, Resources.RibbonCaption, Resources.RibbonGroup6,  Resources.Checkbox3Command, 0, SeButtonStyle.seCheckButton),
                    new MyCommand(MyCommandIds.RadioButton1, Resources.RibbonCaption, Resources.RibbonGroup6,  Resources.Radiobutton1Command, 0, SeButtonStyle.seRadioButton),
                    new MyCommand(MyCommandIds.RadioButton2, Resources.RibbonCaption, Resources.RibbonGroup6,  Resources.Radiobutton2Command, 0, SeButtonStyle.seRadioButton),
                    new MyCommand(MyCommandIds.RadioButton3, Resources.RibbonCaption, Resources.RibbonGroup6,  Resources.Radiobutton3Command, 0, SeButtonStyle.seRadioButton),
                    new MyCommand(MyCommandIds.BoundingBox, Resources.RibbonCaption, Resources.RibbonGroupOverlays,  Resources.BoundingBoxCommand, NativeResources.PNG.BoundingBox_32, SeButtonStyle.seButtonIconAndCaptionBelow),
                    new MyCommand(MyCommandIds.OpenGlBoxes, Resources.RibbonCaption, Resources.RibbonGroupOverlays,  Resources.OpenGlBoxesCommand, NativeResources.PNG.Boxes_32, SeButtonStyle.seButtonIconAndCaptionBelow),
                    new MyCommand(MyCommandIds.GdiPlus, Resources.RibbonCaption, Resources.RibbonGroupOverlays,  Resources.GdiPlusCommand, NativeResources.PNG.GdiPlus_32, SeButtonStyle.seButtonIconAndCaptionBelow)
                };
            }
            else if (EnvCatID == new Guid(CATID.SEDraft))
            {
                myCommands = new MyCommand[]
                {
                    new MyCommand(MyCommandIds.About, Resources.BackstageCategory, null,  Resources.AboutCommand, NativeResources.PNG.Help_32x32, SeButtonStyle.seNoExitBackstageButton),
                    new MyCommand(MyCommandIds.Save, Resources.RibbonCaption, Resources.RibbonGroup1,  Resources.SaveCommand, NativeResources.PNG.Save_16x16, SeButtonStyle.seButtonAutomatic),
                    new MyCommand(MyCommandIds.Folder, Resources.RibbonCaption, Resources.RibbonGroup1,  Resources.FolderCommand, NativeResources.PNG.Folder_16x16, SeButtonStyle.seButtonAutomatic),
                    new MyCommand(MyCommandIds.Monitor, Resources.RibbonCaption, Resources.RibbonGroup1,  Resources.MonitorCommand, NativeResources.PNG.Monitor_16x16, SeButtonStyle.seButtonAutomatic),
                    new MyCommand(MyCommandIds.Box, Resources.RibbonCaption, Resources.RibbonGroup1,  Resources.BoxCommand, NativeResources.PNG.Box_32x32, SeButtonStyle.seButtonIconAndCaptionBelow),
                    new MyCommand(MyCommandIds.Camera, Resources.RibbonCaption, Resources.RibbonGroup2,  Resources.CameraCommand, NativeResources.PNG.Camera_32x32, SeButtonStyle.seButtonIconAndCaptionBelow),
                    new MyCommand(MyCommandIds.Photograph, Resources.RibbonCaption, Resources.RibbonGroup2,  Resources.PhotographCommand, NativeResources.PNG.Photograph_32x32, SeButtonStyle.seButtonIconAndCaptionBelow),
                    new MyCommand(MyCommandIds.Favorites, Resources.RibbonCaption, Resources.RibbonGroup2,  Resources.FavoritesCommand, NativeResources.PNG.Favorites_32x32, SeButtonStyle.seButtonIconAndCaptionBelow),
                    new MyCommand(MyCommandIds.Printer, Resources.RibbonCaption, Resources.RibbonGroup2,  Resources.PrinterCommand, NativeResources.PNG.Printer_32x32, SeButtonStyle.seButtonIconAndCaptionBelow),
                    new MyCommand(MyCommandIds.Tools, Resources.RibbonCaption, Resources.RibbonGroup3,  Resources.ToolsCommand, NativeResources.PNG.Tools_32x32, SeButtonStyle.seButtonIconAndCaptionBelow),
                    new MyCommand(MyCommandIds.CommandPrompt, Resources.RibbonCaption, Resources.RibbonGroup3,  Resources.CommandPromptCommand, NativeResources.PNG.CommandPrompt_32x32, SeButtonStyle.seButtonIconAndCaptionBelow),
                    new MyCommand(MyCommandIds.Notepad, Resources.RibbonCaption, Resources.RibbonGroup3,  Resources.NotepadCommand, NativeResources.PNG.Notepad_32x32, SeButtonStyle.seButtonIconAndCaptionBelow),
                    new MyCommand(MyCommandIds.Help, Resources.RibbonCaption, Resources.RibbonGroup4,  Resources.HelpCommand, NativeResources.PNG.Help_32x32, SeButtonStyle.seButtonIconAndCaptionBelow),
                    new MyCommand(MyCommandIds.Search, Resources.RibbonCaption, Resources.RibbonGroup4,  Resources.SearchCommand, NativeResources.PNG.Search_32x32, SeButtonStyle.seButtonIconAndCaptionBelow),
                    new MyCommand(MyCommandIds.Question, Resources.RibbonCaption, Resources.RibbonGroup5,  Resources.QuestionCommand, NativeResources.PNG.Question_32x32, SeButtonStyle.seButtonIconAndCaptionBelow),
                    new MyCommand(MyCommandIds.CheckBox1, Resources.RibbonCaption, Resources.RibbonGroup6,  Resources.Checkbox1Command, 0, SeButtonStyle.seCheckButton),
                    new MyCommand(MyCommandIds.CheckBox2, Resources.RibbonCaption, Resources.RibbonGroup6,  Resources.Checkbox2Command, 0, SeButtonStyle.seCheckButton),
                    new MyCommand(MyCommandIds.CheckBox3, Resources.RibbonCaption, Resources.RibbonGroup6,  Resources.Checkbox3Command, 0, SeButtonStyle.seCheckButton),
                    new MyCommand(MyCommandIds.RadioButton1, Resources.RibbonCaption, Resources.RibbonGroup6,  Resources.Radiobutton1Command, 0, SeButtonStyle.seRadioButton),
                    new MyCommand(MyCommandIds.RadioButton2, Resources.RibbonCaption, Resources.RibbonGroup6,  Resources.Radiobutton2Command, 0, SeButtonStyle.seRadioButton),
                    new MyCommand(MyCommandIds.RadioButton3, Resources.RibbonCaption, Resources.RibbonGroup6,  Resources.Radiobutton3Command, 0, SeButtonStyle.seRadioButton)
                };
            }

            foreach (var myCommand in myCommands)
            {
                AddCommand(myCommand, EnvCatID, firstTime);
            }

            EnvironmentCommands.Add(EnvCatID, myCommands.ToDictionary(x => x.CommandId));
        }

        public override void OnDisconnection(SolidEdgeFramework.SeDisconnectMode DisconnectMode)
        {
            MyViewOverlay3D?.Dispose();
            MyViewOverlay3D = null;

            // Disconnect from all COM events.
            ComEventsManager.DetachAll();
            ComEventsManager = null;
        }

        #endregion        

        #region SolidEdgeFramework.ISEAddInEventsEx2

        void SolidEdgeFramework.ISEAddInEventsEx2.OnCommand(int CommandID, ShortCutMenuContextConstants Context, DocumentTypeConstants ActiveDocumentType, object pActiveDocument, object pActiveWindow, object pActiveSelectSet)
        {
            var EnvCatId = new Guid(this.Application.GetActiveEnvironment().CATID);
            var commandId = (MyCommandIds)CommandID;
            var myCommand = EnvironmentCommands[EnvCatId][commandId];

            switch (commandId)
            {
                case MyCommandIds.Save:
                    ShowSaveFileDialogDemo();
                    break;
                case MyCommandIds.Folder:
                    ShowFolderBrowserDialogDemo();
                    break;
                case MyCommandIds.Tools:
                    this.Application.StartCommand(SolidEdgeConstants.PartCommandConstants.PartToolsOptions);
                    break;
                case MyCommandIds.Monitor:
                case MyCommandIds.Box:
                case MyCommandIds.Camera:
                    ShowCustomDialogDemo();
                    break;
                case MyCommandIds.Printer:
                    this.Application.StartCommand(SolidEdgeConstants.PartCommandConstants.PartFilePrint);
                    break;
                case MyCommandIds.Help:
                    this.Application.StartCommand(SolidEdgeConstants.PartCommandConstants.PartHelpSolidEdgeontheWeb);
                    break;
                case MyCommandIds.CommandPrompt:
                    Process.Start("cmd.exe");
                    break;
                case MyCommandIds.Notepad:
                    Process.Start("notepad.exe");
                    break;
                case MyCommandIds.BoundingBox:
                    this.MyViewOverlay3D.ShowBoundingBox = !this.MyViewOverlay3D.ShowBoundingBox;
                    myCommand.Checked = this.MyViewOverlay3D.ShowBoundingBox;
                    break;
                case MyCommandIds.GdiPlus:
                    this.MyViewOverlay3D.ShowGDIPlusDemo = !this.MyViewOverlay3D.ShowGDIPlusDemo;
                    myCommand.Checked = this.MyViewOverlay3D.ShowGDIPlusDemo;
                    break;
                case MyCommandIds.OpenGlBoxes:
                    this.MyViewOverlay3D.ShowOpenGlBoxesDemo = !this.MyViewOverlay3D.ShowOpenGlBoxesDemo;
                    myCommand.Checked = this.MyViewOverlay3D.ShowOpenGlBoxesDemo;
                    break;
                case MyCommandIds.About:
                    ShowAboutDialog();
                    break;
                case MyCommandIds.CheckBox1:
                case MyCommandIds.CheckBox2:
                case MyCommandIds.CheckBox3:
                case MyCommandIds.RadioButton1:
                case MyCommandIds.RadioButton2:
                case MyCommandIds.RadioButton3:
                    // Demo toggle check state.
                    myCommand.Checked = !myCommand.Checked;
                    break;
                default:
                    //ShowCustomDialogDemo();
                    var owner = this.Application.GetNativeWindow();
                    MessageBox.Show(owner, myCommand.RuntimeCommandName);
                    break;
            }
        }

        void SolidEdgeFramework.ISEAddInEventsEx2.OnCommandHelp(int hFrameWnd, int HelpCommandID, int CommandID)
        {
        }

        void SolidEdgeFramework.ISEAddInEventsEx2.OnCommandOnLineHelp(int HelpCommandID, int CommandID, out string HelpURL)
        {
            HelpURL = null;
        }

        void SolidEdgeFramework.ISEAddInEventsEx2.OnCommandUpdateUI(int CommandID, ShortCutMenuContextConstants Context, DocumentTypeConstants ActiveDocumentType, object pActiveDocument, object pActiveWindow, object pActiveSelectSet, ref int CommandFlags, out string MenuItemText, ref int BitmapID)
        {
            MenuItemText = null;

            var EnvCatId = new Guid(this.Application.GetActiveEnvironment().CATID);
            var commandId = (MyCommandIds)CommandID;
            var myCommand = EnvironmentCommands[EnvCatId][commandId];
            var flags = default(SolidEdgeConstants.SECommandActivation);

            if (myCommand.Enabled)
            {
                flags |= SolidEdgeConstants.SECommandActivation.seCmdActive_Enabled;
            }

            if (myCommand.Checked)
            {
                flags |= SolidEdgeConstants.SECommandActivation.seCmdActive_Checked;
            }

            CommandFlags = (int)flags;
        }

        #endregion

        #region SolidEdgeFramework.ISEAddInEdgeBarEventsEx

        void SolidEdgeFramework.ISEAddInEdgeBarEventsEx.AddPage(object theDocument)
        {
            if (theDocument is SolidEdgeFramework.SolidEdgeDocument document)
            {
                // .NET WinForm UserControl Example
                {
                    var edgeBarPage1 = AddWinFormEdgeBarPage<DemoAddIn.MyEdgeBarPage, DemoAddIn.MyDocumentEdgeBarControl>(
                        new EdgeBarPageConfiguration
                        {
                            Caption = Resources.MyDocumentEdgeBarCaption1,
                            Index = 1,
                            NativeImageId = NativeResources.PNG.EdgeBar_20x20,
                            NativeResourcesDllPath = this.NativeResourcesDllPath,
                            Title = Resources.MyDocumentEdgeBarCaption1,
                            Tootip = Resources.MyDocumentEdgeBarCaption1
                        },
                        document: document);

                    // edgeBarPage1.ChildObject is an instance of DemoAddIn.MyDocumentEdgeBarControl.
                    var myDocumentEdgeBarControl = (DemoAddIn.MyDocumentEdgeBarControl)edgeBarPage1.ChildObject;
                    myDocumentEdgeBarControl.Application = Application;
                    myDocumentEdgeBarControl.Document = document;
                    myDocumentEdgeBarControl.EdgeBarPage = edgeBarPage1;
                }

                {
                    var edgeBarPage2 = AddWinFormEdgeBarPage<DemoAddIn.MyEdgeBarPage, DemoAddIn.NativeMessageEdgeBarControl>(
                        new EdgeBarPageConfiguration
                        {
                            Caption = "Native Messages",
                            Index = 2,
                            NativeImageId = NativeResources.PNG.EdgeBar_20x20,
                            NativeResourcesDllPath = this.NativeResourcesDllPath,
                            Title = "Native Messages",
                            Tootip = "Native Messages"
                        },
                        document: document);

                    // edgeBarPage1.ChildObject is an instance of DemoAddIn.MyDocumentEdgeBarControl.
                    var myDocumentEdgeBarControl = (DemoAddIn.NativeMessageEdgeBarControl)edgeBarPage2.ChildObject;
                    myDocumentEdgeBarControl.EdgeBarPage = edgeBarPage2;
                }

                // .NET WPF Page Example
                //{
                //    var edgeBarPage2 = AddWpfEdgeBarPage<DemoAddIn.WPF.MyWpfEdgeBarPage>(new EdgeBarPageConfiguration
                //    {
                //        Caption = Resources.MyDocumentEdgeBarCaption2,
                //        Index = 2,
                //        NativeImageId = NativeResources.PNG.EdgeBar_20x20,
                //        NativeResourcesDllPath = this.NativeResourcesDllPath,
                //        Title = Resources.MyDocumentEdgeBarCaption2,
                //        Tootip = Resources.MyDocumentEdgeBarCaption2
                //    },
                //    document: document);

                //    // edgeBarPage2.ChildObject is an instance of System.Windows.Interop.HwndSource.
                //    // edgeBarPage2.ChildObject.RootVisual is an instance of DemoAddIn.WPF.MyWpfEdgeBarPage.
                //    var hwndSource = (System.Windows.Interop.HwndSource)edgeBarPage2.ChildObject;
                //    var myWpfEdgeBarPage = (DemoAddIn.WPF.MyWpfEdgeBarPage)hwndSource.RootVisual;
                //}
            }
            else
            {
                //// .NET WinForm UserControl Example
                //{
                //    // Caution: ISEAddInEdgeBarEventsEx.AddPage(null) will be called multiple times.
                //    // We only want to add a global page once.
                //    if (this.GlobalEdgeBarPages.Any() == false)
                //    {
                //        var edgeBarPage1 = AddWinFormEdgeBarPage<DemoAddIn.MyGlobalEdgeBarControl>(new EdgeBarPageConfiguration
                //        {
                //            Caption = Resources.MyGlobalEdgeBarCaption,
                //            Index = 1,
                //            NativeImageId = NativeResources.PNG.EdgeBar_20x20,
                //            NativeResourcesDllPath = this.NativeResourcesDllPath,
                //            Title = Resources.MyGlobalEdgeBarCaption,
                //            Tootip = Resources.MyGlobalEdgeBarCaption
                //        });

                //        // edgeBarPage1.ChildObject is an instance of DemoAddIn.MyGlobalEdgeBarControl.
                //        var myGlobalEdgeBarControl = (DemoAddIn.MyGlobalEdgeBarControl)edgeBarPage1.ChildObject;
                //        myGlobalEdgeBarControl.Application = Application;
                //    }
                //}
            }
        }

        void SolidEdgeFramework.ISEAddInEdgeBarEventsEx.RemovePage(object theDocument)
        {
            var document = theDocument as SolidEdgeFramework.SolidEdgeDocument;
            RemoveEdgeBarPages(document);
        }

        void SolidEdgeFramework.ISEAddInEdgeBarEventsEx.IsPageDisplayable(object theDocument, string EnvironmentCatID, out bool vbIsPageDisplayable)
        {
            // Default to true.
            vbIsPageDisplayable = true;
        }

        #endregion

        #region SolidEdgeFramework.ISEApplicationEvents

        void SolidEdgeFramework.ISEApplicationEvents.AfterActiveDocumentChange(object theDocument)
        {
        }

        void SolidEdgeFramework.ISEApplicationEvents.AfterCommandRun(int theCommandID)
        {
        }

        void SolidEdgeFramework.ISEApplicationEvents.AfterDocumentOpen(object theDocument)
        {
        }

        void SolidEdgeFramework.ISEApplicationEvents.AfterDocumentPrint(object theDocument, int hDC, ref double ModelToDC, ref int Rect)
        {
        }

        void SolidEdgeFramework.ISEApplicationEvents.AfterDocumentSave(object theDocument)
        {
        }

        void SolidEdgeFramework.ISEApplicationEvents.AfterEnvironmentActivate(object theEnvironment)
        {
        }

        void SolidEdgeFramework.ISEApplicationEvents.AfterNewDocumentOpen(object theDocument)
        {
        }

        void SolidEdgeFramework.ISEApplicationEvents.AfterNewWindow(object theWindow)
        {
        }

        void SolidEdgeFramework.ISEApplicationEvents.AfterWindowActivate(object theWindow)
        {
            if (theWindow is SolidEdgeFramework.Window window)
            {
                MyViewOverlay3D.Window = window; ;
            }
        }

        void SolidEdgeFramework.ISEApplicationEvents.BeforeCommandRun(int theCommandID)
        {
        }

        void SolidEdgeFramework.ISEApplicationEvents.BeforeDocumentClose(object theDocument)
        {
        }

        void SolidEdgeFramework.ISEApplicationEvents.BeforeDocumentPrint(object theDocument, int hDC, ref double ModelToDC, ref int Rect)
        {
        }

        void SolidEdgeFramework.ISEApplicationEvents.BeforeEnvironmentDeactivate(object theEnvironment)
        {
        }

        void SolidEdgeFramework.ISEApplicationEvents.BeforeWindowDeactivate(object theWindow)
        {
        }

        void SolidEdgeFramework.ISEApplicationEvents.BeforeQuit()
        {
        }

        void SolidEdgeFramework.ISEApplicationEvents.BeforeDocumentSave(object theDocument)
        {
        }

        #endregion

        #region SolidEdgeFramework.ISEApplicationEventsEx

        void SolidEdgeFramework.ISEApplicationEventsEx.OnCommandUpdateUI(int CommandID, ref int CommandFlags, out string MenuItemTextD)
        {
            MenuItemTextD = null;

            // Example of how to remove select Solid Edge backstage commands.
            switch ((SolidEdgeConstants.ApplicationCommandConstants)CommandID)
            {
                case SolidEdgeConstants.ApplicationCommandConstants.ApplicationCadenasComponents:
                case SolidEdgeConstants.ApplicationCommandConstants.ApplicationHelpAbout:
                case SolidEdgeConstants.ApplicationCommandConstants.ApplicationSoldEdgeCommunityBlog:
                case SolidEdgeConstants.ApplicationCommandConstants.ApplicationSoldEdgeFacebook:
                //case SolidEdgeConstants.ApplicationCommandConstants.ApplicationSoldEdgeHomePage:
                case SolidEdgeConstants.ApplicationCommandConstants.ApplicationSoldEdgePortal:
                //case SolidEdgeConstants.ApplicationCommandConstants.ApplicationSolidEdgeUserCommunity:
                case SolidEdgeConstants.ApplicationCommandConstants.ApplicationTeamcenterHomePage:
                case SolidEdgeConstants.ApplicationCommandConstants.ApplicationTeamcenterUserCommunity:
                case SolidEdgeConstants.ApplicationCommandConstants.ApplicationTracePartsComponents:
                    CommandFlags |= (int)SolidEdgeConstants.SECommandActivation.seCmdActive_Remove;
                    break;
            }

            // Not all commands will be defined in ApplicationCommandConstants.
            // These we have to determine the command ID and filter accordingly.
            // Put a break in raw_BeforeCommandRun and run the command to get the ID.
            switch (CommandID)
            {
                case 11822: // Tools -> Components by 3Dfind.it
                case 11842: // Tools -> Search 3Dfind.it
                    CommandFlags |= (int)SolidEdgeConstants.SECommandActivation.seCmdActive_Remove;
                    break;
            }
        }

        #endregion

        #region SolidEdgeFramework.ISEApplicationEventsEx2

        void SolidEdgeFramework.ISEApplicationEventsEx2.OnBeforeDocumentOpen(ApplicationBeforeDocumentOpenEvent Context, string Filename, out bool CancelOpen)
        {
            CancelOpen = false;
        }

        #endregion

        #region Registration (regasm.exe) functions

        [ComRegisterFunction]
        static void OnComRegister(Type t)
        {
            if (Guid.Equals(t.GUID, typeof(MyAddIn).GUID))
            {
                var implementedCategories = System.Attribute
                    .GetCustomAttributes(t, typeof(AddInImplementedCategoryAttribute))
                    .Cast<AddInImplementedCategoryAttribute>()
                    .Select(a => a.Value)
                    .ToArray();

                var environmentCategories = System.Attribute
                    .GetCustomAttributes(t, typeof(AddInEnvironmentCategoryAttribute))
                    .Cast<AddInEnvironmentCategoryAttribute>()
                    .Select(a => a.Value)
                    .ToArray();

                var cultures = System.Attribute
                    .GetCustomAttributes(t, typeof(AddInCultureAttribute))
                    .Cast<AddInCultureAttribute>()
                    .Select(a => a.Value);

                var descriptors = cultures
                    .Select(culture => new AddInDescriptor(culture, typeof(Properties.Resources)))
                    .ToArray();

                var settings = new ComRegistrationSettings(t)
                {
                    Enabled = true,
                    ImplementedCategories = implementedCategories,
                    EnvironmentCategories = environmentCategories,
                    Descriptors = descriptors
                };

                ComRegisterSolidEdgeAddIn(settings);
            }
        }

        [ComUnregisterFunction]
        static void OnComUnregister(Type t)
        {
            if (Guid.Equals(t.GUID, typeof(MyAddIn).GUID))
            {
                ComUnregisterSolidEdgeAddIn(t);
            }
        }

        #endregion

        private void AddCommand(MyCommand command, Guid environmentCategory, bool firstTime)
        {
            if (String.IsNullOrWhiteSpace(command.Category))
            {
                throw new System.Exception($"'{nameof(command.Category)}' is empty.");
            }

            char commandSeparator = '\n';

            var commandPrefix = this.Guid.ToString();
            var commandId = (int)command.CommandId;
            var commandParts = command.CommandString
                .Split(new string[] { System.Environment.NewLine }, StringSplitOptions.None)
                .ToList();

            commandParts.Insert(0, $"{commandPrefix}_{commandId}");

            command.RuntimeCategoryName = String.Join(commandSeparator.ToString(), command.Category, command.Group).TrimEnd(commandSeparator);
            command.RuntimeCommandName = String.Join(commandSeparator.ToString(), commandParts.ToArray());

            var addInEx2 = (SolidEdgeFramework.ISEAddInEx2)AddInInstance;
            var EnvironmentCatID = environmentCategory.ToString("B");

            // Allocate command arrays. Please see the addin.doc in the SDK folder for details.
            Array commandNames = new string[] { command.RuntimeCommandName };
            Array commandIDs = new int[] { commandId };
            Array commandButtonStyles = new SolidEdgeFramework.SeButtonStyle[] { command.ButtonStyle };

            addInEx2.SetAddInInfoEx2(
                NativeResourcesDllPath,
                EnvironmentCatID,
                command.RuntimeCategoryName,
                command.ImageId,
                -1,
                -1,
                -1,
                commandNames.Length,
                ref commandNames,
                ref commandIDs,
                ref commandButtonStyles);

            if (firstTime)
            {
                // Shouldn't have to do this anymore but seems necessary for certain button styles.
                switch (command.ButtonStyle)
                {
                    case SeButtonStyle.seCheckButton:
                    case SeButtonStyle.seCheckButtonAndIcon:
                    case SeButtonStyle.seRadioButton:
                        // Add the command bar button.
                        var commandBarName = command.RuntimeCategoryName;
                        var commandBarButton = addInEx2.AddCommandBarButton(EnvironmentCatID, commandBarName, commandId);
                        commandBarButton.Style = command.ButtonStyle;
                        break;
                }
            }

            command.SolidEdgeCommandId = (int)commandIDs.GetValue(0);
        }

        private void ShowAboutDialog()
        {
            using (var dialog = new AboutDialog())
            {
                this.Application.ShowDialog(dialog);
            }
        }

        private void ShowSaveFileDialogDemo()
        {
            using (var dialog = new SaveFileDialog())
            {
                if (this.Application.ShowDialog(dialog) == DialogResult.OK)
                {
                }
            }
        }

        private void ShowFolderBrowserDialogDemo()
        {
            using (var dialog = new FolderBrowserDialog())
            {
                if (this.Application.ShowDialog(dialog) == DialogResult.OK)
                {
                }
            }
        }

        private void ShowCustomDialogDemo()
        {
            using (var dialog = new MyCustomDialog())
            {
                if (this.Application.ShowDialog(dialog) == DialogResult.OK)
                {
                }
            }
        }

        public ComEventsManager ComEventsManager { get; private set; }
        public Dictionary<Guid, Dictionary<MyCommandIds, MyCommand>> EnvironmentCommands { get; private set; } = new Dictionary<Guid, Dictionary<MyCommandIds, MyCommand>>();
        public MyViewOverlay3D MyViewOverlay3D { get; private set; }
    }
}
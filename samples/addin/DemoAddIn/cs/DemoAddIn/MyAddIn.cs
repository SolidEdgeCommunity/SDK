using DemoAddIn.Properties;
using SolidEdgeFramework;
using SolidEdgeSDK.AddIn;
using SolidEdgeSDK.InteropServices;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;

// Please refer to Instructions.txt.
namespace DemoAddIn
{
    [Guid(MyConstants.AddInGuid), ProgId(MyConstants.AddInProgId), ComVisible(true)]
    [AddInImplementedCategory(CATID.SolidEdgeAddIn)]
    [AddInEnvironmentCategory(CATID.SEApplication)]
    [AddInEnvironmentCategory(CATID.SEAllDocumentEnvrionments)]
    [AddInCulture("en-US")]
    [AddInCulture("zh-CN")]
    public class MyAddIn : SolidEdgeAddIn,
        SolidEdgeFramework.ISEAddInEventsEx,
        SolidEdgeFramework.ISEAddInEventsEx2,
        SolidEdgeFramework.ISEAddInEdgeBarEventsEx,
        SolidEdgeFramework.ISEApplicationEvents,
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
            ComEventsManager.Attach<SolidEdgeFramework.ISEAddInEventsEx>(AddInInstance);
            ComEventsManager.Attach<SolidEdgeFramework.ISEAddInEdgeBarEventsEx>(AddInInstance);
            ComEventsManager.Attach<SolidEdgeFramework.ISEApplicationEvents>(Application);

            // Solid Edge 2020 or higher.
            if (SolidEdgeVersion.Major >= 220)
            {
                ComEventsManager.Attach<SolidEdgeFramework.ISEAddInEventsEx2>(AddInInstance);
                ComEventsManager.Attach<SolidEdgeFramework.ISEApplicationEventsEx2>(Application);
            }

            My3dViewOverlay = new My3dViewOverlay(this);
        }

        public override void OnConnectToEnvironment(Guid EnvCatID, SolidEdgeFramework.Environment environment, bool firstTime)
        {
            // If you are not seeing events for a certain environment, check your AddInEnvironmentCategory attribute(s).

            var ribbon3dEnvironments = new Guid[]
            {
                new Guid(CATID.SEAssembly),
                new Guid(CATID.SEPart),
                new Guid(CATID.SEDMPart),
                new Guid(CATID.SESheetMetal),
                new Guid(CATID.SEDMSheetMetal)
            };

            var ribbon2dEnvironments = new Guid[]
            {
                new Guid(CATID.SEDraft)
            };

            if (ribbon3dEnvironments.Contains(EnvCatID))
            {
                AddRibbon<My3dRibbon>(EnvCatID, firstTime);
            }

            if (ribbon2dEnvironments.Contains(EnvCatID))
            {
                AddRibbon<My2dRibbon>(EnvCatID, firstTime);
            }
        }

        public override void OnDisconnection(SolidEdgeFramework.SeDisconnectMode DisconnectMode)
        {
            My3dViewOverlay?.Dispose();
            My3dViewOverlay = null;

            // Disconnect from all COM events.
            ComEventsManager.DetachAll();
            ComEventsManager = null;
        }

        #endregion

        #region SolidEdgeFramework.ISEAddInEvents

        void SolidEdgeFramework.ISEAddInEvents.OnCommand(int CommandID)
        {
            // Forward the call to the active ribbon if it implements this specific event.
            ((SolidEdgeFramework.ISEAddInEvents)ActiveRibbon)?.OnCommand(CommandID);
        }

        void SolidEdgeFramework.ISEAddInEvents.OnCommandHelp(int hFrameWnd, int HelpCommandID, int CommandID)
        {
            // Forward the call to the active ribbon if it implements this specific event.
            ((SolidEdgeFramework.ISEAddInEvents)ActiveRibbon)?.OnCommandHelp(hFrameWnd, HelpCommandID, CommandID);
        }

        void SolidEdgeFramework.ISEAddInEvents.OnCommandUpdateUI(int CommandID, ref int CommandFlags, out string MenuItemText, ref int BitmapID)
        {
            MenuItemText = null;

            // Forward the call to the active ribbon if it implements this specific event.
            ((SolidEdgeFramework.ISEAddInEvents)ActiveRibbon)?.OnCommandUpdateUI(CommandID, ref CommandFlags, out MenuItemText, ref BitmapID);
        }

        #endregion

        #region SolidEdgeFramework.ISEAddInEventsEx

        void SolidEdgeFramework.ISEAddInEventsEx.OnCommand(int CommandID)
        {
            // Forward the call to the active ribbon if it implements this specific event.
            ((SolidEdgeFramework.ISEAddInEventsEx)ActiveRibbon)?.OnCommand(CommandID);
        }

        void SolidEdgeFramework.ISEAddInEventsEx.OnCommandHelp(int hFrameWnd, int HelpCommandID, int CommandID)
        {
            // Forward the call to the active ribbon if it implements this specific event.
            ((SolidEdgeFramework.ISEAddInEventsEx)ActiveRibbon)?.OnCommandHelp(hFrameWnd, HelpCommandID, CommandID);
        }

        void SolidEdgeFramework.ISEAddInEventsEx.OnCommandUpdateUI(int CommandID, ref int CommandFlags, out string MenuItemText, ref int BitmapID)
        {
            MenuItemText = null;

            // Forward the call to the active ribbon if it implements this specific event.
            ((SolidEdgeFramework.ISEAddInEventsEx)ActiveRibbon)?.OnCommandUpdateUI(CommandID, ref CommandFlags, out MenuItemText, ref BitmapID);
        }

        void SolidEdgeFramework.ISEAddInEventsEx.OnCommandOnLineHelp(int HelpCommandID, int CommandID, out string HelpURL)
        {
            HelpURL = null;

            // Forward the call to the active ribbon if it implements this specific event.
            ((SolidEdgeFramework.ISEAddInEventsEx)ActiveRibbon)?.OnCommandOnLineHelp(HelpCommandID, CommandID, out HelpURL);
        }

        #endregion

        #region SolidEdgeFramework.ISEAddInEventsEx2

        void SolidEdgeFramework.ISEAddInEventsEx2.OnCommand(int CommandID, ShortCutMenuContextConstants Context, DocumentTypeConstants ActiveDocumentType, object pActiveDocument, object pActiveWindow, object pActiveSelectSet)
        {
            // Forward the call to the active ribbon if it implements this specific event.
            ((SolidEdgeFramework.ISEAddInEventsEx2)ActiveRibbon)?.OnCommand(CommandID, Context, ActiveDocumentType, pActiveDocument, pActiveWindow, pActiveSelectSet);
        }

        void SolidEdgeFramework.ISEAddInEventsEx2.OnCommandHelp(int hFrameWnd, int HelpCommandID, int CommandID)
        {
            // Forward the call to the active ribbon if it implements this specific event.
            ((SolidEdgeFramework.ISEAddInEventsEx2)ActiveRibbon)?.OnCommandHelp(hFrameWnd, HelpCommandID, CommandID);
        }

        void SolidEdgeFramework.ISEAddInEventsEx2.OnCommandUpdateUI(int CommandID, ShortCutMenuContextConstants Context, DocumentTypeConstants ActiveDocumentType, object pActiveDocument, object pActiveWindow, object pActiveSelectSet, ref int CommandFlags, out string MenuItemText, ref int BitmapID)
        {
            MenuItemText = null;

            // Forward the call to the active ribbon if it implements this specific event.
            ((SolidEdgeFramework.ISEAddInEventsEx2)ActiveRibbon)?.OnCommandUpdateUI(CommandID, Context, ActiveDocumentType, pActiveDocument, pActiveWindow, pActiveSelectSet, ref CommandFlags, out MenuItemText, ref BitmapID);
        }

        void SolidEdgeFramework.ISEAddInEventsEx2.OnCommandOnLineHelp(int HelpCommandID, int CommandID, out string HelpURL)
        {
            HelpURL = null;

            // Forward the call to the active ribbon if it implements this specific event.
            ((SolidEdgeFramework.ISEAddInEventsEx2)ActiveRibbon)?.OnCommandOnLineHelp(HelpCommandID, CommandID, out HelpURL);
        }

        #endregion

        #region SolidEdgeFramework.ISEAddInEdgeBarEventsEx

        void SolidEdgeFramework.ISEAddInEdgeBarEventsEx.AddPage(object theDocument)
        {
            if (theDocument is SolidEdgeFramework.SolidEdgeDocument document)
            {
                // .NET WinForm UserControl Example
                {
                    var edgeBarPage1 = AddWinFormEdgeBarPage<DemoAddIn.MyDocumentEdgeBarControl>(
                        new EdgeBarPageConfiguration
                        {
                            Caption = Resources.MyDocumentEdgeBarCaption1,
                            Index = 1,
                            NativeImageId = NativeResources.PNG.EdgeBar_20,
                            NativeResourcesDllPath = this.NativeResourcesDllPath,
                            Title = Resources.MyDocumentEdgeBarCaption1,
                            Tootip = Resources.MyDocumentEdgeBarCaption1
                        },
                        document: document);

                    // edgeBarPage1.ChildObject is an instance of DemoAddIn.MyDocumentEdgeBarControl.
                    var myDocumentEdgeBarControl = (DemoAddIn.MyDocumentEdgeBarControl)edgeBarPage1.ChildObject;
                    myDocumentEdgeBarControl.Application = Application;
                    myDocumentEdgeBarControl.Document = document;
                }

                // .NET WPF Page Example
                {
                    var edgeBarPage2 = AddWpfEdgeBarPage<DemoAddIn.WPF.MyWpfEdgeBarPage>(new EdgeBarPageConfiguration
                    {
                        Caption = Resources.MyDocumentEdgeBarCaption2,
                        Index = 2,
                        NativeImageId = NativeResources.PNG.EdgeBar_20,
                        NativeResourcesDllPath = this.NativeResourcesDllPath,
                        Title = Resources.MyDocumentEdgeBarCaption2,
                        Tootip = Resources.MyDocumentEdgeBarCaption2
                    },
                    document: document);

                    // edgeBarPage2.ChildObject is an instance of System.Windows.Interop.HwndSource.
                    // edgeBarPage2.ChildObject.RootVisual is an instance of DemoAddIn.WPF.MyWpfEdgeBarPage.
                    var hwndSource = (System.Windows.Interop.HwndSource)edgeBarPage2.ChildObject;
                    var myWpfEdgeBarPage = (DemoAddIn.WPF.MyWpfEdgeBarPage)hwndSource.RootVisual;
                }
            }
            else
            {
                // .NET WinForm UserControl Example
                {
                    // Caution: ISEAddInEdgeBarEventsEx.AddPage(null) will be called multiple times.
                    // We only want to add a global page once.
                    if (this.GlobalEdgeBarPages.Any() == false)
                    {
                        var edgeBarPage1 = AddWinFormEdgeBarPage<DemoAddIn.MyGlobalEdgeBarControl>(new EdgeBarPageConfiguration
                        {
                            Caption = Resources.MyGlobalEdgeBarCaption,
                            Index = 1,
                            NativeImageId = NativeResources.PNG.EdgeBar_20,
                            NativeResourcesDllPath = this.NativeResourcesDllPath,
                            Title = Resources.MyGlobalEdgeBarCaption,
                            Tootip = Resources.MyGlobalEdgeBarCaption
                        });

                        // edgeBarPage1.ChildObject is an instance of DemoAddIn.MyGlobalEdgeBarControl.
                        var myGlobalEdgeBarControl = (DemoAddIn.MyGlobalEdgeBarControl)edgeBarPage1.ChildObject;
                        myGlobalEdgeBarControl.Application = Application;
                    }
                }
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
                My3dViewOverlay.Window = window; ;
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

        public ComEventsManager ComEventsManager { get; private set; }
        public My3dViewOverlay My3dViewOverlay { get; private set; }
    }
}
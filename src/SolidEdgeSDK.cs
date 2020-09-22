//
// Licensed under the MIT license.
// See https://github.com/SolidEdgeCommunity/SDK/blob/master/LICENSE for full license information.
//

// This file is maintained at https://github.com/SolidEdgeCommunity/SDK.

// Requires additional references to PresentationCore.dll, PresentationFramework.dll & WindowsBase.dll.
//#define SE_SDK_WPF_SUPPORT

using Microsoft.Win32;
using SolidEdgeAssembly;
using SolidEdgeFramework;
using SolidEdgeGeometry;
using SolidEdgeSDK.Extensions;
using SolidEdgeSDK.InteropServices;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.Resources;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Text;
using System.Threading;

namespace SolidEdgeSDK
{
    /// <summary>
    /// Solid Edge category IDs imported from \sdk\include\secatids.h.
    /// </summary>
    public static class CATID
    {
        /// <summary>
        /// CATID_SolidEdgeAddIn
        /// </summary>
        public const string SolidEdgeAddIn = "26B1D2D1-2B03-11d2-B589-080036E8B802";

        /// <summary>
        /// CATID_SolidEdgeOEMAddIn
        /// </summary>
        public const string SolidEdgeOEMAddIn = "E71958D9-E29B-45F1-A502-A52231F24239";

        /// <summary>
        /// CATID_SEApplication
        /// </summary>
        public const string SEApplication = "26618394-09D6-11d1-BA07-080036230602";

        /// <summary>
        /// CATID_SEAssembly
        /// </summary>
        public const string SEAssembly = "26618395-09D6-11d1-BA07-080036230602";

        /// <summary>
        /// CATID_SEMotion
        /// </summary>
        public const string SEMotion = "67ED3F40-A351-11d3-A40B-0004AC969602";

        /// <summary>
        /// CATID_SEPart
        /// </summary>
        public const string SEPart = "26618396-09D6-11d1-BA07-080036230602";

        /// <summary>
        /// CATID_SEProfile
        /// </summary>
        public const string SEProfile = "26618397-09D6-11d1-BA07-080036230602";

        /// <summary>
        /// CATID_SEFeatureRecognition
        /// </summary>
        public const string SEFeatureRecognition = "E6F9C8DC-B256-11d3-A41E-0004AC969602";

        /// <summary>
        /// CATID_SESheetMetal
        /// </summary>
        public const string SESheetMetal = "26618398-09D6-11D1-BA07-080036230602";

        /// <summary>
        /// CATID_SEDraft
        /// </summary>
        public const string SEDraft = "08244193-B78D-11D2-9216-00C04F79BE98";

        /// <summary>
        /// CATID_SEWeldment
        /// </summary>
        public const string SEWeldment = "7313526A-276F-11D4-B64E-00C04F79B2BF";

        /// <summary>
        /// CATID_SEXpresRoute
        /// </summary>
        public const string SEXpresRoute = "1661432A-489C-4714-B1B2-61E85CFD0B71";

        /// <summary>
        /// CATID_SEExplode
        /// </summary>
        public const string SEExplode = "23BE4212-5810-478b-94FF-B4D682C1B538";

        /// <summary>
        /// CATID_SESimplify
        /// </summary>
        public const string SESimplify = "CE3DCEBF-E36E-4851-930A-ED892FE0772A";

        /// <summary>
        /// CATID_SEStudio
        /// </summary>
        public const string SEStudio = "D35550BF-0810-4f67-97D5-789EDBC23F4D";

        /// <summary>
        /// CATID_SELayout
        /// </summary>
        public const string SELayout = "27B34941-FFCD-4768-9102-0B6698656CEA";

        /// <summary>
        /// CATID_SESketch
        /// </summary>
        public const string SESketch = "0DDABC90-125E-4cfe-9CB7-DC97FB74CCF4";

        /// <summary>
        /// CATID_SEProfileHole
        /// </summary>
        public const string SEProfileHole = "0D5CC5F7-5BA3-4d2f-B6A9-31D9B401FE30";

        /// <summary>
        /// CATID_SEProfilePattern
        /// </summary>
        public const string SEProfilePattern = "7BD57D4B-BA47-4a79-A4E2-DFFD43B97ADF";

        /// <summary>
        /// CATID_SEProfileRevolved
        /// </summary>
        public const string SEProfileRevolved = "FB73C683-1536-4073-B792-E28B8D31146E";

        /// <summary>
        /// CATID_SEDrawingViewEdit
        /// </summary>
        public const string SEDrawingViewEdit = "8DBC3B5F-02D6-4241-BE96-B12EAF83FAE6";

        /// <summary>
        /// CATID_SERefAxis
        /// </summary>
        public const string SERefAxis = "B21CCFF8-1FDD-4f44-9417-F1EAE06888FA";

        /// <summary>
        /// CATID_SECuttingPlaneLine
        /// </summary>
        public const string SECuttingPlaneLine = "7C6F65F1-A02D-4c3c-8063-8F54B59B34E3";

        /// <summary>
        /// CATID_SEBrokenOutSectionProfile
        /// </summary>
        public const string SEBrokenOutSectionProfile = "534CAB66-8089-4e18-8FC4-6FA5A957E445";

        /// <summary>
        /// CATID_SEFrame
        /// </summary>
        public const string SEFrame = "D84119E8-F844-4823-B3A0-D4F31793028A";

        /// <summary>
        /// CATID_2dModel
        /// </summary>
        public const string SE2dModel = "F6031120-7D99-48a7-95FC-EEE8038D7996";

        /// <summary>
        /// CATID_EditBlockView
        /// </summary>
        public const string SEEditBlockView = "892A1CDA-12AE-4619-BB69-C5156C929832";

        /// <summary>
        /// CATID_EditBlockInPlace
        /// </summary>
        public const string EditBlockInPlace = "308A1927-CDCE-4b92-B654-241362608CDE";

        /// <summary>
        /// CATID_SEComponentSketchInPart
        /// </summary>
        public const string SEComponentSketchInPart = "FAB8DC23-00F4-4872-8662-18DD013F2095";

        /// <summary>
        /// CATID_SEComponentSketchInAsm
        /// </summary>
        public const string SEComponentSketchInAsm = "86D925FB-66ED-40d2-AA3D-D04E74838141";

        /// <summary>
        /// CATID_SEHarness
        /// </summary>
        public const string SEHarness = "5337A0AB-23ED-4261-A238-00E2070406FC";

        /// <summary>
        /// CATID_SEAll
        /// </summary>
        public const string SEAll = "C484ED57-DBB6-4a83-BEDB-C08600AF07BF";

        /// <summary>
        /// CATID_SEAllDocumentEnvrionments
        /// </summary>
        public const string SEAllDocumentEnvrionments = "BAD41B8D-18FF-42c9-9611-8A00E6921AE8";

        /// <summary>
        /// CATID_SEEditMV
        /// </summary>
        public const string SEEditMV = "C1D8CCB8-54D3-4fce-92AB-0668147FC7C3";

        /// <summary>
        /// CATID_SEEditMVPart
        /// </summary>
        public const string SEEditMVPart = "054BDB42-6C1E-41a4-9014-3D51BEE911EF";

        /// <summary>
        /// CATID_SEDMPart
        /// </summary>
        public const string SEDMPart = "D9B0BB85-3A6C-4086-A0BB-88A1AAD57A58";

        /// <summary>
        /// CATID_SEDMSheetMetal
        /// </summary>
        public const string SEDMSheetMetal = "9CBF2809-FF80-4dbc-98F2-B82DABF3530F";

        /// <summary>
        /// CATID_SEDMAssembly
        /// </summary>
        public const string SEDMAssembly = "2C3C2A72-3A4A-471d-98B5-E3A8CFA4A2BF";

        /// <summary>
        /// CATID_FEAResultsPart
        /// </summary>
        public const string FEAResultsPart = "B5965D1C-8819-4902-8252-64841537A16C";

        /// <summary>
        /// CATID_FEAResultsAsm
        /// </summary>
        public const string FEAResultsAsm = "986B2512-3AE9-4a57-8513-1D2A1E3520DD";

        /// <summary>
        /// CATID_SESimplifiedAssemblyPart
        /// </summary>
        public const string SESimplifiedAssemblyPart = "E7350DC3-6E7A-4D53-A53F-5B1C7A0709B3";

        /// <summary>
        /// CATID_Sketch3d
        /// </summary>
        public const string Sketch3d = "07F05BA4-18CD-4B87-8E2F-49864E71B41F";

        /// <summary>
        /// CATID_SEAssemblyViewer
        /// </summary>
        public const string SEAssemblyViewer = "F2483121-58BC-44AF-8B8F-D7B74DC8408B";
    }
}

namespace SolidEdgeSDK.AddIn
{
    public class AddInDescriptor
    {
        public AddInDescriptor()
        {
        }

        public AddInDescriptor(CultureInfo culture, Type resourceType)
        {
            Culture = culture;

            var resourceManager = new ResourceManager(resourceType.FullName, resourceType.Assembly);
            Description = resourceManager.GetString("AddInDescription", Culture);
            Summary = resourceManager.GetString("AddInSummary", Culture);
        }

        public CultureInfo Culture { get; set; }
        public string Description { get; set; }
        public string Summary { get; set; }
    }

    [AttributeUsage(AttributeTargets.Class, AllowMultiple = true, Inherited = false)]
    public class AddInEnvironmentCategoryAttribute : System.Attribute
    {
        public AddInEnvironmentCategoryAttribute(string guid)
            : this(new Guid(guid))
        {
        }

        public AddInEnvironmentCategoryAttribute(Guid guid)
        {
            Value = guid;
        }

        public Guid Value { get; private set; }
    }

    [AttributeUsage(AttributeTargets.Class, AllowMultiple = true, Inherited = false)]
    public class AddInImplementedCategoryAttribute : System.Attribute
    {
        public AddInImplementedCategoryAttribute(string guid)
            : this(new Guid(guid))
        {
        }

        public AddInImplementedCategoryAttribute(Guid guid)
        {
            Value = guid;
        }

        public Guid Value { get; private set; }
    }

    [AttributeUsage(AttributeTargets.Class, AllowMultiple = true, Inherited = false)]
    public sealed class AddInCultureAttribute : System.Attribute
    {
        public AddInCultureAttribute(string culture)
        {
            Value = new CultureInfo(culture);
        }

        public CultureInfo Value { get; private set; }
    }

    public sealed class ComRegistrationSettings
    {
        public ComRegistrationSettings(Type type)
        {
            this.Type = type;
        }

        public Type Type { get; private set; }
        public bool Enabled { get; set; }

        public Guid[] ImplementedCategories { get; set; } = new Guid[] { };
        public Guid[] EnvironmentCategories { get; set; } = new Guid[] { };
        public AddInDescriptor[] Descriptors { get; set; } = new AddInDescriptor[] { };
    }

    public class EdgeBarController : IDisposable
    {
        private bool _disposed = false;

        internal EdgeBarController(SolidEdgeAddIn addIn)
        {
            SolidEdgeAddIn = addIn ?? throw new ArgumentNullException(nameof(addIn));
        }

        /// <summary>
        /// Destructor
        /// </summary>
        ~EdgeBarController()
        {
            Dispose(false);
        }

        public EdgeBarPage AddWinFormPage<TControl>(EdgeBarPageConfiguration config) where TControl : System.Windows.Forms.Control, new()
        {
            return AddWinFormPage<TControl>(
                config: config,
                document: null);
        }

        public EdgeBarPage AddWinFormPage<TControl>(EdgeBarPageConfiguration config, SolidEdgeFramework.SolidEdgeDocument document) where TControl : System.Windows.Forms.Control, new()
        {
            TControl control = Activator.CreateInstance<TControl>();

            var edgeBarPage = AddPage(
                config: config,
                controlHandle: control.Handle,
                document: document);

            edgeBarPage.ChildObject = control;

            return edgeBarPage;
        }

#if SE_SDK_WPF_SUPPORT

        public EdgeBarPage AddWpfPage<TControl>(EdgeBarPageConfiguration config) where TControl : System.Windows.Controls.Page, new()
        {
            return AddWpfPage<TControl>(
                config: config,
                document: null);
        }

        public EdgeBarPage AddWpfPage<TControl>(EdgeBarPageConfiguration config, SolidEdgeFramework.SolidEdgeDocument document) where TControl : System.Windows.Controls.Page, new()
        {
            uint WS_VISIBLE = 0x10000000;
            uint WS_CHILD = 0x40000000;
            uint WS_MAXIMIZE = 0x01000000;

            TControl control = Activator.CreateInstance<TControl>();

            var edgeBarPage = AddPage(
                config: config,
                controlHandle: IntPtr.Zero,
                document: document);

            var hwndSource = new System.Windows.Interop.HwndSource(new System.Windows.Interop.HwndSourceParameters
            {
                PositionX = 0,
                PositionY = 0,
                Height = 0,
                Width = 0,
                ParentWindow = edgeBarPage.Handle,
                WindowStyle = (int)(WS_VISIBLE | WS_CHILD | WS_MAXIMIZE)
            })
            {
                RootVisual = control
            };

            edgeBarPage.ChildObject = hwndSource;

            return edgeBarPage;
        }

#endif

        public EdgeBarPage AddPage(EdgeBarPageConfiguration config, IntPtr controlHandle, SolidEdgeFramework.SolidEdgeDocument document)
        {
            if (config == null) throw new ArgumentNullException(nameof(config));

            var options = (int)config.GetOptions(document);
            var direction = (int)config.Direction;

            var hWndEdgeBarPage = SolidEdgeAddIn.EdgeBarEx2.AddPageEx2(
                theDocument: document,
                ResourceFilename: config.NativeResourcesDllPath,
                Index: config.Index,
                nBitmapID: config.NativeImageId,
                Tooltip: config.Tootip,
                Title: config.Title,
                Caption: config.Caption,
                nOptions: options,
                Direction: direction);

            var edgeBarPage = new EdgeBarPage(new IntPtr(hWndEdgeBarPage), config.Index)
            {
                ChildWindowHandle = controlHandle,
                Document = document
            };

            EdgeBarPages.Add(edgeBarPage);

            return edgeBarPage;
        }

        public void RemoveAllPages()
        {
            var edgeBarPages = EdgeBarPages.ToArray();

            foreach (var edgeBarPage in edgeBarPages)
            {
                RemovePage(edgeBarPage);
            }
        }

        public void RemoveDocumentPages(SolidEdgeFramework.SolidEdgeDocument document)
        {
            if (document == null) throw new ArgumentNullException(nameof(document));

            var edgeBarPages = DocumentEdgeBarPages.Where(x => x.Document == document);

            foreach (var edgeBarPage in edgeBarPages)
            {
                RemovePage(edgeBarPage);
            }
        }

        public void RemoveGlobalPages()
        {
            var edgeBarPages = GlobalEdgeBarPages.ToArray();

            foreach (var edgeBarPage in edgeBarPages)
            {
                RemovePage(edgeBarPage);
            }
        }

        private void RemovePage(EdgeBarPage edgeBarPage)
        {
            int hWnd = edgeBarPage.Handle.ToInt32();

            try
            {
                SolidEdgeAddIn.EdgeBarEx2.RemovePage(edgeBarPage.Document, hWnd, 0);
            }
            catch
            {
            }

            try
            {
                edgeBarPage.Dispose();
            }
            catch
            {
            }

            EdgeBarPages.Remove(edgeBarPage);
        }

        #region IDisposable implementation

        /// <summary>
        /// Disposes resources.
        /// </summary>
        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        void Dispose(bool disposing)
        {
            if (_disposed) return;

            if (disposing)
            {
                // Free managed objects here.
                EdgeBarPages.Clear();
            }

            // Free unmanaged objects here.
            _disposed = true;
        }

        #endregion

        public SolidEdgeAddIn SolidEdgeAddIn { get; private set; }
        private List<EdgeBarPage> EdgeBarPages { get; set; } = new List<EdgeBarPage>();

        public EdgeBarPage[] DocumentEdgeBarPages
        {
            get
            {
                return EdgeBarPages.Where(x => x.Document != null).ToArray();
            }
        }

        public EdgeBarPage[] GlobalEdgeBarPages
        {
            get
            {
                return EdgeBarPages.Where(x => x.Document == null).ToArray();
            }
        }
    }

    public class EdgeBarPage : System.Windows.Forms.NativeWindow, IDisposable
    {
        [DllImport("user32.dll", SetLastError = true)]
        static extern IntPtr SetParent(IntPtr hWndChild, IntPtr hWndNewParent);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        private bool _disposed = false;
        private IntPtr _hWndChildWindow = IntPtr.Zero;

        public EdgeBarPage(IntPtr hWnd, int index = 0)
        {
            if (hWnd == IntPtr.Zero) throw new System.ArgumentException($"{nameof(hWnd)} cannot be IntPtr.Zero.");

            Index = index;

            // Take ownership of HWND returned from ISolidEdgeBar::AddPage.
            AssignHandle(hWnd);
        }

        #region Properties

        public bool IsDisposed { get { return _disposed; } }
        public object ChildObject { get; set; }
        
        public IntPtr ChildWindowHandle
        {
            get { return _hWndChildWindow; }
            set
            {
                _hWndChildWindow = value;

                if (_hWndChildWindow != IntPtr.Zero)
                {
                    // Reparent child control to this hWnd.
                    SetParent(_hWndChildWindow, this.Handle);

                    // Show the child control and maximize it to fill the entire EdgeBarPage region.
                    ShowWindow(_hWndChildWindow, 3 /* SHOWMAXIMIZED */);
                }
            }
        }

        public int Index { get; private set; }
        public SolidEdgeFramework.SolidEdgeDocument Document { get; set; }
        public virtual bool Visible { get; set; } = true;

        #endregion

        #region IDisposable implementation

        ~EdgeBarPage()
        {
            Dispose(false);
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        void Dispose(bool disposing)
        {
            if (!_disposed)
            {
                if (disposing)
                {
                    try
                    {
                        ((IDisposable)this.ChildObject)?.Dispose();
                    }
                    catch
                    {
                    }
                }

                try
                {
                    ReleaseHandle();
                }
                catch
                {
                }

                ChildObject = null;
                _disposed = true;
            }
        }

        #endregion
    }

    public sealed class EdgeBarPageConfiguration
    {
        public int Index { get; set; }
        public string NativeResourcesDllPath { get; set; }
        public int NativeImageId { get; set; }
        public string Tootip { get; set; }
        public string Title { get; set; }
        public string Caption { get; set; }

        /// <summary>
        /// Hint as to where the docking pane should dock.
        /// </summary>
        public EdgeBarPageDock Direction { get; set; } = EdgeBarPageDock.Left;

        /// <summary>
        /// EdgeBarConstant.DONOT_MAKE_ACTIVE
        /// </summary>
        public bool MakeActive { get; set; } = false;

        /// <summary>
        /// EdgeBarConstant.NO_RESIZE_CHILD
        /// </summary>
        public bool ResizeChild { get; set; } = true;

        /// <summary>
        /// EdgeBarConstant.UPDATE_ON_PANE_SLIDING
        /// </summary>
        public bool UpdateOnPaneSliding { get; set; } = false;

        /// <summary>
        /// EdgeBarConstant.WANTS_ACCELERATORS
        /// </summary>
        public bool WantsAccelerators { get; set; } = false;

        public SolidEdgeConstants.EdgeBarConstant GetOptions(SolidEdgeFramework.SolidEdgeDocument document = null)
        {
            var options = default(SolidEdgeConstants.EdgeBarConstant);

            if (document == null)
            {
                options |= SolidEdgeConstants.EdgeBarConstant.TRACK_CLOSE_GLOBALLY;
            }
            else
            {
                options |= SolidEdgeConstants.EdgeBarConstant.TRACK_CLOSE_BYDOCUMENT;
            }

            if (MakeActive == false)
            {
                options |= SolidEdgeConstants.EdgeBarConstant.DONOT_MAKE_ACTIVE;
            }

            if (ResizeChild == false)
            {
                options |= SolidEdgeConstants.EdgeBarConstant.NO_RESIZE_CHILD;
            }

            if (UpdateOnPaneSliding == true)
            {
                options |= SolidEdgeConstants.EdgeBarConstant.UPDATE_ON_PANE_SLIDING;
            }

            if (WantsAccelerators == true)
            {
                options |= SolidEdgeConstants.EdgeBarConstant.WANTS_ACCELERATORS;
            }

            return options;
        }
    }

    public enum EdgeBarPageDock
    {
        Left = 0,
        Rigth = 1,
        Top = 2,
        Button = 3,
        None = 4
    }

    /// <summary>
    /// Base class for Solid Edge AddIns.
    /// </summary>
    /// <remarks>
    /// For the most part this class contains only plumbing code that generally should not need much modification.
    /// </remarks>
    public abstract class SolidEdgeAddIn : MarshalByRefObject, SolidEdgeFramework.ISolidEdgeAddIn
    {
        private static SolidEdgeAddIn _instance;
        private AppDomain _isolatedDomain;
        private SolidEdgeFramework.ISolidEdgeAddIn _isolatedAddIn;

        public SolidEdgeAddIn()
        {
            // Removing this call will break the addin.
            AppDomain.CurrentDomain.AssemblyResolve += (sender, args)
                => AppDomain.CurrentDomain.GetAssemblies().FirstOrDefault(x => x.FullName == args.Name);
        }

        void SolidEdgeFramework.ISolidEdgeAddIn.OnConnection(object Application, SolidEdgeFramework.SeConnectMode ConnectMode, SolidEdgeFramework.AddIn AddInInstance)
        {
            var type = this.GetType();

            if (AppDomain.CurrentDomain.IsDefaultAppDomain())
            {
                // We are in the "DefaultDomain".

                // Create a new (isolated) AppDomain and name it this.GetType().GUID (AddIn GUID).

                var appDomainSetup = AppDomain.CurrentDomain.SetupInformation;
                var evidence = AppDomain.CurrentDomain.Evidence;
                appDomainSetup.ApplicationBase = System.IO.Path.GetDirectoryName(type.Assembly.Location);

                var domainName = AddInInstance.GUID;
                _isolatedDomain = AppDomain.CreateDomain(domainName, evidence, appDomainSetup);
                var isolatedAddIn = _isolatedDomain.CreateInstanceAndUnwrap(type.Assembly.FullName, type.FullName);
                _isolatedAddIn = (SolidEdgeFramework.ISolidEdgeAddIn)isolatedAddIn;

                if (_isolatedAddIn != null)
                {
                    var application = (SolidEdgeFramework.Application)Application;

                    // Forward call to isolated AppDomain.
                    _isolatedAddIn.OnConnection(application, ConnectMode, AddInInstance);
                }
            }
            else
            {
                // We are in the isolated addin AppDomain.
                _instance = this;

                this.Application = UnwrapTransparentProxy<SolidEdgeFramework.Application>(Application);
                this.AddInInstance = UnwrapTransparentProxy<SolidEdgeFramework.AddIn>(AddInInstance);

                OnConnection(ConnectMode);
            }
        }

        void SolidEdgeFramework.ISolidEdgeAddIn.OnConnectToEnvironment(string EnvCatID, object pEnvironmentDispatch, bool bFirstTime)
        {
            if ((AppDomain.CurrentDomain.IsDefaultAppDomain()) && (_isolatedAddIn != null))
            {
                // We are in the "DefaultDomain".

                // Forward call to isolated AppDomain.
                _isolatedAddIn.OnConnectToEnvironment(EnvCatID, pEnvironmentDispatch, bFirstTime);
            }
            else
            {
                // We are in the isolated addin AppDomain.

                var environment = UnwrapTransparentProxy<SolidEdgeFramework.Environment>(pEnvironmentDispatch);
                var environmentCategory = new Guid(EnvCatID);

                OnConnectToEnvironment(environmentCategory, environment, bFirstTime);
            }
        }

        void SolidEdgeFramework.ISolidEdgeAddIn.OnDisconnection(SolidEdgeFramework.SeDisconnectMode DisconnectMode)
        {
            if (AppDomain.CurrentDomain.IsDefaultAppDomain())
            {
                // We are in the "DefaultDomain".

                if (_isolatedAddIn != null)
                {
                    // Forward call to isolated AppDomain.
                    _isolatedAddIn.OnDisconnection(DisconnectMode);

                    // Unload isolated domain.
                    if (_isolatedDomain != null)
                    {
                        AppDomain.Unload(_isolatedDomain);
                    }

                    _isolatedDomain = null;
                    _isolatedAddIn = null;
                }
            }
            else
            {
                // We are in the isolated addin AppDomain.

                OnDisconnection(DisconnectMode);
            }
        }

        public abstract void OnConnection(SolidEdgeFramework.SeConnectMode ConnectMode);
        public abstract void OnConnectToEnvironment(Guid EnvCatID, SolidEdgeFramework.Environment environment, bool firstTime);
        public abstract void OnDisconnection(SolidEdgeFramework.SeDisconnectMode DisconnectMode);

        public override object InitializeLifetimeService()
        {
            // Removing this call will break the addin.
            return null;
        }

        private TInterface UnwrapTransparentProxy<TInterface>(object rcw) where TInterface : class
        {
            if (System.Runtime.Remoting.RemotingServices.IsTransparentProxy(rcw))
            {
                IntPtr punk = Marshal.GetIUnknownForObject(rcw);

                try
                {
                    return (TInterface)Marshal.GetObjectForIUnknown(punk);
                }
                finally
                {
                    Marshal.Release(punk);
                }
            }

            return rcw as TInterface;
        }

        /// <summary>
        /// Writes HKEY_CLASSES_ROOT\CLSID\{ADDIN_GUID}.
        /// </summary>
        /// <param name="settings"></param>
        public static void ComRegisterSolidEdgeAddIn(ComRegistrationSettings settings)
        {
            var type = settings.Type;

            using (var baseKey = CreateBaseKey(type.GUID))
            {
                baseKey.SetValue("AutoConnect", settings.Enabled ? 1 : 0);

                foreach (var guid in settings.ImplementedCategories)
                {
                    var subkey = String.Join(@"\", "Implemented Categories", guid.ToString("B"));

                    baseKey.CreateSubKey(subkey);
                }

                foreach (var guid in settings.EnvironmentCategories)
                {
                    var subkey = String.Join(@"\", "Environment Categories", guid.ToString("B"));

                    baseKey.CreateSubKey(subkey);
                }

                foreach (var descriptor in settings.Descriptors)
                {
                    // Example Local ID (LCID)
                    // Description: English - United States
                    // int: 1033
                    // hex: 0x0409

                    // Convert LCID to hexadecimal.
                    var lcid = descriptor.Culture.LCID.ToString("X4");

                    // Remove leading zeros.
                    lcid = lcid.TrimStart('0');

                    // Write the title value.
                    // HKEY_CLASSES_ROOT\CLSID\{ADDIN_GUID}\409
                    baseKey.SetValue(lcid, descriptor.Description);

                    // Write the summary key.
                    // HKEY_CLASSES_ROOT\CLSID\{ADDIN_GUID}\Summary\409
                    using (var summaryKey = baseKey.CreateSubKey("Summary"))
                    {
                        summaryKey.SetValue(lcid, descriptor.Summary);
                    }
                }
            }
        }

        /// <summary>
        /// Deletes HKEY_CLASSES_ROOT\CLSID\{ADDIN_GUID}.
        /// </summary>
        /// <param name="t"></param>
        public static void ComUnregisterSolidEdgeAddIn(Type t)
        {
            string subkey = String.Join(@"\", "CLSID", t.GUID.ToString("B"));
            Registry.ClassesRoot.DeleteSubKeyTree(subkey, false);
        }

        static RegistryKey CreateBaseKey(Guid guid)
        {
            string subkey = String.Join(@"\", "CLSID", guid.ToString("B"));
            return Registry.ClassesRoot.CreateSubKey(subkey);
        }

        /// <summary>
        /// Returns a global static instance of the addin.
        /// </summary>
        public static SolidEdgeAddIn Instance { get { return _instance; } }

        public SolidEdgeFramework.Application Application { get; private set; }
        public SolidEdgeFramework.AddIn AddInInstance { get; private set; }

        public virtual System.IO.FileInfo FileInfo
        {
            get { return new System.IO.FileInfo(this.GetType().Assembly.Location); }
        }

        /// <summary>
        /// Returns the GUID of the AddIn.
        /// </summary>
        public Guid Guid { get { return new Guid(AddInInstance.GUID); } }

        /// <summary>
        /// Returns the version of the GUI for the AddIn.
        /// </summary>
        public int GuiVersion { get { return AddInInstance.GuiVersion; } }

        /// <summary>
        /// Returns the path of the .dll\.exe containing native Win32 resources. Typically the current assembly location.
        /// </summary>
        /// <remarks>
        /// It is only necessary to override if you have native resources in a separate .dll.
        /// </remarks>
        public virtual string NativeResourcesDllPath
        {
            get { return this.GetType().Assembly.Location; }
        }

        /// <summary>
        /// Returns an instance of SolidEdgeFramework.ISolidEdgeBarEx2.
        /// </summary>
        /// <remarks>Available in Solid Edge ST6 or greater.</remarks>
        public SolidEdgeFramework.ISolidEdgeBarEx2 EdgeBarEx2 { get { return AddInInstance as SolidEdgeFramework.ISolidEdgeBarEx2; } }

        public Version SolidEdgeVersion { get { return new Version(Application.Version); ; } }
    }

    public class RibbonTab
    {
        internal RibbonTab(Ribbon ribbon, string name)
        {
            Ribbon = ribbon;
            Name = name;
        }

        public Ribbon Ribbon { get; private set; }
        public string Name { get; private set; }
        public RibbonGroup[] Groups { get; set; } = new RibbonGroup[] { };

        public IEnumerable<RibbonControl> Controls
        {
            get
            {
                foreach (var group in this.Groups)
                {
                    foreach (var control in group.Controls)
                    {
                        yield return control;
                    }
                }
            }
        }

        public IEnumerable<RibbonButton> Buttons { get { return this.Controls.OfType<RibbonButton>(); } }
        public IEnumerable<RibbonCheckBox> CheckBoxes { get { return this.Controls.OfType<RibbonCheckBox>(); } }
        public IEnumerable<RibbonRadioButton> RadioButtons { get { return this.Controls.OfType<RibbonRadioButton>(); } }

        public override string ToString() { return Name; }
    }

    public class RibbonGroup
    {
        private List<RibbonControl> _controls = new List<RibbonControl>();

        public RibbonGroup(string name)
        {
            Name = name;
        }

        public string Name { get; private set; }

        public RibbonControl[] Controls { get; set; } = new RibbonControl[] { };

        public IEnumerable<RibbonButton> Buttons { get { return this.Controls.OfType<RibbonButton>(); } }
        public IEnumerable<RibbonCheckBox> CheckBoxes { get { return this.Controls.OfType<RibbonCheckBox>(); } }
        public IEnumerable<RibbonRadioButton> RadioButtons { get { return this.Controls.OfType<RibbonRadioButton>(); } }

        public override string ToString() { return Name; }
    }

    /// <summary>
    /// Controller class for working with ribbons.
    /// </summary>
    public sealed class RibbonController : IDisposable,
        SolidEdgeFramework.ISEAddInEvents,
        SolidEdgeFramework.ISEAddInEventsEx
    {
        private List<Ribbon> _ribbons = new List<Ribbon>();
        private Dictionary<IConnectionPoint, int> _connectionPointDictionary = new Dictionary<IConnectionPoint, int>();
        private bool _disposed = false;

        internal RibbonController(SolidEdgeAddIn addIn)
        {
            SolidEdgeAddIn = addIn ?? throw new ArgumentNullException(nameof(addIn));
            ComEventsManager = new ComEventsManager(this);

            if (SolidEdgeAddIn.SolidEdgeVersion.Major <= 105)
            {
                // Solid Edge ST5 or lower.
                ComEventsManager.Attach<SolidEdgeFramework.ISEAddInEvents>(SolidEdgeAddIn.AddInInstance);
            }
            else
            {
                // Solid Edge ST6 or higher.
                ComEventsManager.Attach<SolidEdgeFramework.ISEAddInEventsEx>(SolidEdgeAddIn.AddInInstance);
            }
        }

        /// <summary>
        /// Destructor
        /// </summary>
        ~RibbonController()
        {
            Dispose(false);
        }

        #region SolidEdgeFramework.ISEAddInEvents implentation

        void SolidEdgeFramework.ISEAddInEvents.OnCommand(int CommandID)
        {
            // Forward call to ISEAddInEventsEx implementation.
            ((SolidEdgeFramework.ISEAddInEventsEx)this).OnCommand(CommandID);
        }

        void SolidEdgeFramework.ISEAddInEvents.OnCommandHelp(int hFrameWnd, int HelpCommandID, int CommandID)
        {
            // Forward call to ISEAddInEventsEx implementation.
            ((SolidEdgeFramework.ISEAddInEventsEx)this).OnCommandHelp(hFrameWnd, HelpCommandID, CommandID);
        }

        void SolidEdgeFramework.ISEAddInEvents.OnCommandUpdateUI(int CommandID, ref int CommandFlags, out string MenuItemText, ref int BitmapID)
        {
            // Forward call to ISEAddInEventsEx implementation.
            MenuItemText = null;
            ((SolidEdgeFramework.ISEAddInEventsEx)this).OnCommandUpdateUI(CommandID, ref CommandFlags, out MenuItemText, ref BitmapID);
        }

        #endregion

        #region SolidEdgeFramework.ISEAddInEventsEx implementation

        void SolidEdgeFramework.ISEAddInEventsEx.OnCommand(int CommandID)
        {
            var ribbon = ActiveRibbon;

            if (ribbon != null)
            {
                var control = ribbon.Controls.Where(x => x.CommandId == CommandID).FirstOrDefault();

                if (control != null)
                {
                    control.DoClick();
                    ribbon.OnControlClick(control);

                    if (control is RibbonButton button)
                    {
                        ribbon.OnButtonClick(button);
                    }
                    else if (control is RibbonCheckBox checkBox)
                    {
                        ribbon.OnCheckBoxClick(checkBox);
                    }
                    else if (control is RibbonRadioButton radioButton)
                    {
                        ribbon.OnRadioButtonClick(radioButton);
                    }
                }
            }
        }

        void SolidEdgeFramework.ISEAddInEventsEx.OnCommandHelp(int hFrameWnd, int HelpCommandID, int CommandID)
        {
            var ribbon = ActiveRibbon;

            if (ribbon != null)
            {
                var control = ribbon.Controls.Where(x => x.CommandId == CommandID).FirstOrDefault();

                if (control != null)
                {
                    control.DoHelp(new IntPtr(hFrameWnd), HelpCommandID);
                }
            }
        }

        void SolidEdgeFramework.ISEAddInEventsEx.OnCommandOnLineHelp(int HelpCommandID, int CommandID, out string HelpURL)
        {
            var ribbon = ActiveRibbon;

            if (ribbon != null)
            {
                var control = ribbon.Controls.FirstOrDefault(x => x.CommandId == CommandID);

                if (control != null)
                {
                    HelpURL = control.WebHelpURL;
                    return;
                }
            }
            HelpURL = null;
        }

        void SolidEdgeFramework.ISEAddInEventsEx.OnCommandUpdateUI(int CommandID, ref int CommandFlags, out string MenuItemText, ref int BitmapID)
        {
            MenuItemText = null;
            var ribbon = ActiveRibbon;
            var flags = default(SolidEdgeConstants.SECommandActivation); ;

            if (ribbon != null)
            {
                var control = ribbon.Controls.Where(x => x.CommandId == CommandID).FirstOrDefault();

                if (control != null)
                {
                    if (control.Enabled)
                    {
                        flags |= SolidEdgeConstants.SECommandActivation.seCmdActive_Enabled;
                    }

                    if (control.Checked)
                    {
                        flags |= SolidEdgeConstants.SECommandActivation.seCmdActive_Checked;
                    }

                    if (control.UseDotMark)
                    {
                        flags |= SolidEdgeConstants.SECommandActivation.seCmdActive_UseDotMark;
                    }

                    flags |= SolidEdgeConstants.SECommandActivation.seCmdActive_ChangeText;
                    MenuItemText = control.Label;

                    CommandFlags = (int)flags;
                }
            }
        }

        #endregion

        #region Methods

        /// <summary>
        /// Adds a ribbon to the specified environment.
        /// </summary>
        /// <typeparam name="TRibbon"></typeparam>
        /// <param name="environmentCategory"></param>
        /// <param name="firstTime"></param>
        public TRibbon Add<TRibbon>(Guid environmentCategory, bool firstTime) where TRibbon : Ribbon
        {
            TRibbon ribbon = Activator.CreateInstance<TRibbon>();

            ribbon.EnvironmentCategory = environmentCategory;
            ribbon.SolidEdgeAddIn = this.SolidEdgeAddIn;
            ribbon.Initialize();

            Add(ribbon, environmentCategory, firstTime);

            return ribbon;
        }

        /// <summary>
        /// Adds a ribbon to the specified environment.
        /// </summary>
        /// <param name="ribbon"></param>
        /// <param name="environmentCategory"></param>
        /// <param name="firstTime"></param>
        public void Add(Ribbon ribbon, Guid environmentCategory, bool firstTime)
        {
            if (ribbon == null) throw new ArgumentNullException("ribbon");

            // Solid Edge ST or higher.
            var addInEx = (SolidEdgeFramework.ISEAddInEx)SolidEdgeAddIn.AddInInstance;

            // Solid Edge ST7 or higher.
            var addInEx2 = (SolidEdgeFramework.ISEAddInEx2)null;
            //var addInEx2 = _addIn.AddInEx2;

            var EnvironmentCatID = environmentCategory.ToString("B");

            if (_ribbons.Exists(x => x.EnvironmentCategory.Equals(ribbon.EnvironmentCategory)))
            {
                throw new System.Exception(String.Format("A ribbon has already been added for environment category {0}.", ribbon.EnvironmentCategory));
            }

            if (ribbon.EnvironmentCategory.Equals(Guid.Empty))
            {
                throw new System.Exception(String.Format("{0} is not a valid environment category.", ribbon.EnvironmentCategory));
            }

            foreach (var tab in ribbon.Tabs)
            {
                foreach (var group in tab.Groups)
                {
                    foreach (var control in group.Controls)
                    {
                        // Properly format the command bar name string.
                        var categoryName = tab.Name;

                        // Properly format the command name string.
                        StringBuilder commandName = new StringBuilder();

                        // Note: The command will not be added if it the name is not unique!
                        commandName.AppendFormat("{0}_{1}", SolidEdgeAddIn.Guid.ToString(), control.CommandId);

                        if (control is RibbonButton)
                        {
                            var ribbonButton = control as RibbonButton;
                            if (String.IsNullOrEmpty(ribbonButton.DropDownGroup) == false)
                            {
                                // Now append the description, tooltip, etc separated by \n.
                                commandName.AppendFormat("\n{0}\\\\{1}\n{2}\n{3}", ribbonButton.DropDownGroup, control.Label, control.SuperTip, control.ScreenTip);
                            }
                            else
                            {
                                // Now append the description, tooltip, etc separated by \n.
                                commandName.AppendFormat("\n{0}\n{1}\n{2}", control.Label, control.SuperTip, control.ScreenTip);
                            }
                        }
                        else
                        {
                            // Now append the description, tooltip, etc separated by \n.
                            commandName.AppendFormat("\n{0}\n{1}\n{2}", control.Label, control.SuperTip, control.ScreenTip);
                        }

                        // Append macro info if provided.
                        if (!String.IsNullOrEmpty(control.Macro))
                        {
                            commandName.AppendFormat("\n{0}", control.Macro);

                            if (!String.IsNullOrEmpty(control.MacroParameters))
                            {
                                commandName.AppendFormat("\n{0}", control.MacroParameters);
                            }
                        }

                        // Assign the control's CommandName property. Mostly just for reference.
                        control.CommandName = commandName.ToString();

                        // Allocate command arrays. Please see the addin.doc in the SDK folder for details.
                        Array commandNames = new string[] { control.CommandName };
                        Array commandIDs = new int[] { control.CommandId };

                        if (addInEx2 != null)
                        {
                            // Currently having an issue with SetAddInInfoEx2() in that the commandButtonStyles don't seem to apply.
                            // Need to investigate further. For now, addInEx2 is set to null.

                            categoryName = String.Format("{0}\n{1}", tab.Name, group.Name);
                            Array commandButtonStyles = new SolidEdgeFramework.SeButtonStyle[] { control.Style };

                            // ST7 or higher.
                            addInEx2.SetAddInInfoEx2(
                                SolidEdgeAddIn.NativeResourcesDllPath,
                                EnvironmentCatID,
                                categoryName,
                                control.ImageId,
                                -1,
                                -1,
                                -1,
                                commandNames.Length,
                                ref commandNames,
                                ref commandIDs,
                                ref commandButtonStyles);
                        }
                        else if (addInEx != null)
                        {
                            // ST or higher
                            addInEx.SetAddInInfoEx(
                                SolidEdgeAddIn.NativeResourcesDllPath,
                                EnvironmentCatID,
                                categoryName,
                                control.ImageId,
                                -1,
                                -1,
                                -1,
                                commandNames.Length,
                                ref commandNames,
                                ref commandIDs);

                            if (firstTime)
                            {
                                var commandBarName = String.Format("{0}\n{1}", tab.Name, group.Name);

                                // Add the command bar button.
                                SolidEdgeFramework.CommandBarButton pButton = addInEx.AddCommandBarButton(EnvironmentCatID, commandBarName, control.CommandId);

                                // Set the button style.
                                if (pButton != null)
                                {
                                    pButton.Style = control.Style;
                                }
                            }
                        }

                        control.SolidEdgeCommandId = (int)commandIDs.GetValue(0);
                    }
                }
            }

            _ribbons.Add(ribbon);
        }

        #endregion

        #region Properties

        /// <summary>
        /// Returns the ribbon for the current environment.
        /// </summary>
        public Ribbon ActiveRibbon
        {
            get
            {
                var application = SolidEdgeAddIn.Application;
                var environment = application.GetActiveEnvironment();
                var envCatId = Guid.Parse(environment.CATID);
                return _ribbons.Where(x => x.EnvironmentCategory.Equals(envCatId)).FirstOrDefault();
            }
        }

        /// <summary>
        /// Returns an enumerable collection of ribbons.
        /// </summary>
        public IEnumerable<Ribbon> Ribbons { get { return _ribbons.AsEnumerable(); } }

        public SolidEdgeAddIn SolidEdgeAddIn { get; private set; }
        public ComEventsManager ComEventsManager { get; set; }

        #endregion

        #region IDisposable implementation

        /// <summary>
        /// Disposes resources.
        /// </summary>
        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        void Dispose(bool disposing)
        {
            if (_disposed) return;

            if (disposing)
            {
                // Free managed objects here.

                _ribbons.Clear();
            }

            // Free unmanaged objects here. 
            ComEventsManager.DetachAll();

            _disposed = true;
        }

        #endregion
    }

    [Serializable]
    public delegate void RibbonControlClickEventHandler(RibbonControl control);

    [Serializable]
    public delegate void RibbonControlHelpEventHandler(RibbonControl control, IntPtr hWndFrame, int helpCommandID);

    /// <summary>
    /// Abstract base class for all ribbon controls.
    /// </summary>
    public abstract class RibbonControl
    {
        /// <summary>
        /// Raised when a user clicks the control.
        /// </summary>
        public event RibbonControlClickEventHandler Click;

        /// <summary>
        /// Raised when user requests help for the control.
        /// </summary>
        public event RibbonControlHelpEventHandler Help;

        internal RibbonControl(int commandId)
        {
            CommandId = commandId;
        }

        public bool UseDotMark { get; set; }

        /// <summary>
        /// Gets or set a value indicating whether the control is in the checked state.
        /// </summary>
        public bool Checked { get; set; }

        /// <summary>
        /// Gets the command id of the control.
        /// </summary>
        /// <remarks>
        /// This is the command id used by OnCommand SolidEdgeFramework.AddInEvents.
        /// </remarks>
        public int CommandId { get; set; }

        /// <summary>
        /// Returns the generated command name used when calling SetAddInInfo().
        /// </summary>
        public string CommandName { get; set; }

        /// <summary>
        /// Gets or set a value indicating whether the control is enabled.
        /// </summary>
        public bool Enabled { get; set; } = true;

        /// <summary>
        /// Gets ribbon group that the control is assigned to.
        /// </summary>
        public RibbonGroup Group { get; set; }

        /// <summary>
        /// Gets or set a value referencing an image embedded into the assembly as a native resource using the NativeResourceAttribute.
        /// </summary>
        /// <remarks>
        /// Changing this value after the ribbon has been initialized has no impact.
        /// </remarks>
        public int ImageId { get; set; }

        /// <summary>
        /// Gets or set a value indicating the label (caption) of the control.
        /// </summary>
        public string Label { get; set; }

        /// <summary>
        /// Gets or set a value indicating a macro (.exe) assigned to the control.
        /// </summary>
        public string Macro { get; set; }

        /// <summary>
        /// Gets or set a value indicating the macro (.exe) parameters assigned to the control.
        /// </summary>
        public string MacroParameters { get; set; }

        /// <summary>
        /// Gets or set a value indicating screentip of the control.
        /// </summary>
        /// <remarks>
        /// Changing this value after the ribbon has been initialized has no impact.
        /// </remarks>
        public string ScreenTip { get; set; }

        /// <summary>
        /// Gets or set a value indicating whether to show the image assigned to the control.
        /// </summary>
        public bool ShowImage { get; set; }

        /// <summary>
        /// Gets or set a value indicating whether to show the label (caption) assigned to the control.
        /// </summary>
        public bool ShowLabel { get; set; }

        /// <summary>
        /// Gets the Solid Edge assigned runtime command id of the control.
        /// </summary>
        /// <remarks>
        /// This is the command id used by the BeforeCommand and AfterCommand SolidEdgeFramework.ApplicationEvents. 
        /// You also use this command id when calling SolidEdgeFramework.Application.StartCommand().
        /// </remarks>
        public int SolidEdgeCommandId { get; set; }

        internal abstract SolidEdgeFramework.SeButtonStyle Style { get; }

        /// <summary>
        /// Gets or set a value indicating the supertip of the control.
        /// </summary>
        /// <remarks>
        /// Changing this value after the ribbon has been initialized has no impact.
        /// </remarks>
        public string SuperTip { get; set; }

        /// <summary>
        /// Gets or set the telp URL that is shown in the browser if the user asks for help by using the F1 key.
        /// </summary>
        public string WebHelpURL { get; set; }

        internal virtual void DoClick()
        {
            Click?.Invoke(this);
        }

        internal void DoHelp(IntPtr hWndFrame, int helpCommandID)
        {
            Help?.Invoke(this, hWndFrame, helpCommandID);
        }
    }

    public class RibbonButton : RibbonControl
    {
        public RibbonButton(int id)
            : base(id)
        {
        }

        public RibbonButton(int id, string label, string screentip, string supertip, int imageId, RibbonButtonSize size = RibbonButtonSize.Normal)
            : base(id)
        {
            Label = label;
            ScreenTip = screentip;
            SuperTip = supertip;
            ImageId = imageId;
            Size = size;
        }

        public string DropDownGroup { get; set; }
        public RibbonButtonSize Size { get; set; } = RibbonButtonSize.Normal;

        internal override SolidEdgeFramework.SeButtonStyle Style
        {
            get
            {
                var style = SolidEdgeFramework.SeButtonStyle.seButtonAutomatic;

                if (Size == RibbonButtonSize.Large)
                {
                    style = SolidEdgeFramework.SeButtonStyle.seButtonIconAndCaptionBelow;
                }
                else
                {
                    if (ShowImage && ShowLabel)
                    {
                        style = SolidEdgeFramework.SeButtonStyle.seButtonIconAndCaption;
                    }
                    else if (ShowImage)
                    {
                        style = SolidEdgeFramework.SeButtonStyle.seButtonIcon;
                    }
                    else if (ShowLabel)
                    {
                        style = SolidEdgeFramework.SeButtonStyle.seButtonCaption;
                    }
                }

                return style;
            }

        }
    }

    public class RibbonCheckBox : RibbonControl
    {
        public RibbonCheckBox(int id)
            : base(id)
        {
        }

        public RibbonCheckBox(int id, string label, string screentip, string supertip)
            : base(id)
        {
            Label = label;
            ScreenTip = screentip;
            SuperTip = supertip;
        }

        internal override SolidEdgeFramework.SeButtonStyle Style
        {
            get
            {
                var style = SolidEdgeFramework.SeButtonStyle.seCheckButton;

                if (this.ImageId >= 0)
                {
                    style = ShowImage == true ? SolidEdgeFramework.SeButtonStyle.seCheckButtonAndIcon : SolidEdgeFramework.SeButtonStyle.seCheckButton;
                }

                return style;
            }
        }
    }

    public class RibbonRadioButton : RibbonControl
    {
        public RibbonRadioButton(int id)
            : base(id)
        {
        }

        public RibbonRadioButton(int id, string label, string screentip, string supertip)
            : base(id)
        {
            Label = label;
            ScreenTip = screentip;
            SuperTip = supertip;
        }

        internal override SolidEdgeFramework.SeButtonStyle Style { get { return SolidEdgeFramework.SeButtonStyle.seRadioButton; } }
    }

    public enum RibbonButtonSize
    {
        Normal,
        Large
    }

    public abstract class Ribbon
    {
        public abstract void Initialize();

        protected void Initialize(RibbonTab[] tabs)
        {
            if (Tabs.Any())
            {
                throw new System.Exception($"{this.GetType().FullName} has already been initialized.");
            }

            Tabs = tabs;
        }

        public virtual void OnControlClick(RibbonControl control)
        {
        }

        public virtual void OnButtonClick(RibbonButton button)
        {
        }

        public virtual void OnCheckBoxClick(RibbonCheckBox checkBox)
        {
        }

        public virtual void OnRadioButtonClick(RibbonRadioButton radiobutton)
        {
        }

        public RibbonButton GetButton(int commandId)
        {
            return Buttons.FirstOrDefault(x => x.CommandId == commandId);
        }

        public RibbonCheckBox GetCheckBox(int commandId)
        {
            return CheckBoxes.FirstOrDefault(x => x.CommandId == commandId);
        }

        public TRibbonControl GetControl<TRibbonControl>(int commandId) where TRibbonControl : RibbonControl
        {
            return Controls.OfType<TRibbonControl>().FirstOrDefault(x => x.CommandId == commandId);
        }

        public RibbonRadioButton GetRadioButton(int commandId)
        {
            return RadioButtons.FirstOrDefault(x => x.CommandId == commandId);
        }

        public IEnumerable<RibbonControl> Controls
        {
            get
            {
                foreach (var tab in Tabs)
                {
                    foreach (var control in tab.Controls)
                    {
                        yield return control;
                    }
                }
            }
        }

        public SolidEdgeAddIn SolidEdgeAddIn { get; internal set; }
        public Guid EnvironmentCategory { get; set; }

        public IEnumerable<RibbonButton> Buttons { get { return this.Controls.OfType<RibbonButton>(); } }
        public IEnumerable<RibbonCheckBox> CheckBoxes { get { return this.Controls.OfType<RibbonCheckBox>(); } }
        public IEnumerable<RibbonRadioButton> RadioButtons { get { return this.Controls.OfType<RibbonRadioButton>(); } }

        public RibbonControl this[int commandId] { get { return this.Controls.Where(x => x.CommandId == commandId).FirstOrDefault(); } }
        public RibbonTab[] Tabs { get; private set; } = new RibbonTab[] { };
    }
}

namespace SolidEdgeSDK.Extensions
{
    /// <summary>
    /// SolidEdgeFramework.Application extension methods.
    /// </summary>
    public static class ApplicationExtensions
    {
        /// <summary>
        /// Returns the current active document.
        /// </summary>
        public static SolidEdgeFramework.SolidEdgeDocument GetActiveDocument(this SolidEdgeFramework.Application application, bool throwOnError = true)
        {
            try
            {
                // ActiveDocument will throw an exception if no document is open.
                return (SolidEdgeFramework.SolidEdgeDocument)application.ActiveDocument;
            }
            catch
            {
                if (throwOnError) throw;
            }

            return null;
        }

        /// <summary>
        /// Returns the currently active document.
        /// </summary>
        /// <typeparam name="T">The type to return.</typeparam>
        public static T GetActiveDocument<T>(this SolidEdgeFramework.Application application, bool throwOnError = true) where T : class
        {
            try
            {
                // ActiveDocument will throw an exception if no document is open.
                return (T)application.ActiveDocument;
            }
            catch
            {
                if (throwOnError) throw;
            }

            return null;
        }

        /// <summary>
        /// Returns the environment that belongs to the current active window of the document.
        /// </summary>
        public static SolidEdgeFramework.Environment GetActiveEnvironment(this SolidEdgeFramework.Application application)
        {
            SolidEdgeFramework.Environments environments = application.Environments;
            return environments.Item(application.ActiveEnvironment);
        }

        /// <summary>
        /// Returns an environment specified by CATID.
        /// </summary>
        public static SolidEdgeFramework.Environment GetEnvironment(this SolidEdgeFramework.Application application, string CATID)
        {
            var guid1 = new Guid(CATID);

            foreach (var environment in application.Environments.OfType<SolidEdgeFramework.Environment>())
            {
                var guid2 = new Guid(environment.CATID);
                if (guid1.Equals(guid2))
                {
                    return environment;
                }
            }

            return null;
        }

        /// <summary>
        /// Returns the value of a specified global constant.
        /// </summary>
        public static object GetGlobalParameter(this SolidEdgeFramework.Application application, SolidEdgeFramework.ApplicationGlobalConstants globalConstant)
        {
            object value = null;
            application.GetGlobalParameter(globalConstant, ref value);
            return value;
        }

        /// <summary>
        /// Returns the value of a specified global constant.
        /// </summary>
        /// <typeparam name="T">The type to return.</typeparam>
        public static T GetGlobalParameter<T>(this SolidEdgeFramework.Application application, SolidEdgeFramework.ApplicationGlobalConstants globalConstant)
        {
            object value = null;
            application.GetGlobalParameter(globalConstant, ref value);
            return (T)value;
        }

        /// <summary>
        /// Returns a NativeWindow object that represents the main application window.
        /// </summary>
        public static System.Windows.Forms.NativeWindow GetNativeWindow(this SolidEdgeFramework.Application application)
        {
            return System.Windows.Forms.NativeWindow.FromHandle(new IntPtr(application.hWnd));
        }

        /// <summary>
        /// Returns a Process object that represents the application prcoess.
        /// </summary>
        public static Process GetProcess(this SolidEdgeFramework.Application application)
        {
            return Process.GetProcessById(application.ProcessID);
        }

        /// <summary>
        /// Returns a Version object that represents the application version.
        /// </summary>
        public static Version GetVersion(this SolidEdgeFramework.Application application)
        {
            return new Version(application.Version);
        }

        /// <summary>
        /// Returns a point object to the application main window.
        /// </summary>
        public static IntPtr GetWindowHandle(this SolidEdgeFramework.Application application)
        {
            return new IntPtr(application.hWnd);
        }

        /// <summary>
        /// Activates a specified Solid Edge command.
        /// </summary>
        public static void StartCommand(this SolidEdgeFramework.Application application, SolidEdgeFramework.SolidEdgeCommandConstants CommandID)
        {
            application.StartCommand((SolidEdgeFramework.SolidEdgeCommandConstants)CommandID);
        }

        /// <summary>
        /// Shows the form with the application main window as the owner.
        /// </summary>
        public static void Show(this SolidEdgeFramework.Application application, System.Windows.Forms.Form form)
        {
            if (form == null) throw new ArgumentNullException("form");

            form.Show(application.GetNativeWindow());
        }

        /// <summary>
        /// Shows the form as a modal dialog box with the application main window as the owner.
        /// </summary>
        public static System.Windows.Forms.DialogResult ShowDialog(this SolidEdgeFramework.Application application, System.Windows.Forms.Form form)
        {
            if (form == null) throw new ArgumentNullException("form");

            return form.ShowDialog(application.GetNativeWindow());
        }

        /// <summary>
        /// Shows the form as a modal dialog box with the application main window as the owner.
        /// </summary>
        public static System.Windows.Forms.DialogResult ShowDialog(this SolidEdgeFramework.Application application, System.Windows.Forms.CommonDialog dialog)
        {
            if (dialog == null) throw new ArgumentNullException("dialog");
            return dialog.ShowDialog(application.GetNativeWindow());
        }

        public static void UpdateActiveWindow(this SolidEdgeFramework.Application application)
        {
            if (application.ActiveWindow is SolidEdgeFramework.Window window)
            {
                // 3D Window. Update view Window.View.
                window.View?.Update();
            }
            else if (application.ActiveWindow is SolidEdgeDraft.SheetWindow sheetWindow)
            {
                // 2D Window
                sheetWindow.Update();
            }
        }
    }

    public static class SheetExtensions
    {
        [DllImport("user32.dll")]
        static extern bool CloseClipboard();

        [DllImport("user32.dll", CharSet = CharSet.Unicode, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
        static extern IntPtr GetClipboardData(uint format);

        [DllImport("user32.dll")]
        static extern IntPtr GetClipboardOwner();

        [DllImport("user32.dll")]
        static extern bool IsClipboardFormatAvailable(uint format);

        [DllImport("user32.dll")]
        static extern bool OpenClipboard(IntPtr hWndNewOwner);

        [DllImport("gdi32.dll")]
        static extern bool DeleteEnhMetaFile(IntPtr hemf);

        [DllImport("gdi32.dll")]
        static extern uint GetEnhMetaFileBits(IntPtr hemf, uint cbBuffer, [Out] byte[] lpbBuffer);

        const uint CF_ENHMETAFILE = 14;

        public static System.Drawing.Imaging.Metafile GetEnhancedMetafile(this SolidEdgeDraft.Sheet sheet)
        {
            try
            {
                // Copy the sheet as an EMF to the windows clipboard.
                sheet.CopyEMFToClipboard();

                if (OpenClipboard(IntPtr.Zero))
                {
                    if (IsClipboardFormatAvailable(CF_ENHMETAFILE))
                    {
                        // Get the handle to the EMF.
                        IntPtr henhmetafile = GetClipboardData(CF_ENHMETAFILE);

                        return new System.Drawing.Imaging.Metafile(henhmetafile, true);
                    }
                    else
                    {
                        throw new System.Exception("CF_ENHMETAFILE is not available in clipboard.");
                    }
                }
                else
                {
                    throw new System.Exception("Error opening clipboard.");
                }
            }
            finally
            {
                CloseClipboard();
            }
        }

        public static void SaveAsEnhancedMetafile(this SolidEdgeDraft.Sheet sheet, string filename)
        {
            try
            {
                // Copy the sheet as an EMF to the windows clipboard.
                sheet.CopyEMFToClipboard();

                if (OpenClipboard(IntPtr.Zero))
                {
                    if (IsClipboardFormatAvailable(CF_ENHMETAFILE))
                    {
                        // Get the handle to the EMF.
                        IntPtr hEMF = GetClipboardData(CF_ENHMETAFILE);

                        // Query the size of the EMF.
                        uint len = GetEnhMetaFileBits(hEMF, 0, null);
                        byte[] rawBytes = new byte[len];

                        // Get all of the bytes of the EMF.
                        GetEnhMetaFileBits(hEMF, len, rawBytes);

                        // Write all of the bytes to a file.
                        System.IO.File.WriteAllBytes(filename, rawBytes);

                        // Delete the EMF handle.
                        DeleteEnhMetaFile(hEMF);
                    }
                    else
                    {
                        throw new System.Exception("CF_ENHMETAFILE is not available in clipboard.");
                    }
                }
                else
                {
                    throw new System.Exception("Error opening clipboard.");
                }
            }
            finally
            {
                CloseClipboard();
            }
        }
    }

    public static class UnitsOfMeasureExtensions
    {
        public static T FormatUnit<T>(this SolidEdgeFramework.UnitsOfMeasure unitsOfMeasure, SolidEdgeFramework.UnitTypeConstants unitType, double value) where T : class
        {
            return (T)unitsOfMeasure.FormatUnit((int)unitType, value, null);
        }

        public static T FormatUnit<T>(this SolidEdgeFramework.UnitsOfMeasure unitsOfMeasure, SolidEdgeFramework.UnitTypeConstants unitType, double value, SolidEdgeConstants.PrecisionConstants precision) where T : class
        {
            return (T)unitsOfMeasure.FormatUnit((int)unitType, value, precision);
        }
    }

    /// <summary>
    /// SolidEdgeFramework.Window extension methods.
    /// </summary>
    public static class WindowExtensions
    {
        /// <summary>
        /// Returns an IntPtr representing the window handle.
        /// </summary>
        public static IntPtr GetDrawHandle(this SolidEdgeFramework.Window window)
        {
            return new IntPtr(window.DrawHwnd);
        }

        /// <summary>
        /// Returns an IntPtr representing the window handle.
        /// </summary>
        public static IntPtr GetHandle(this SolidEdgeFramework.Window window)
        {
            return new IntPtr(window.hWnd);
        }
    }
}

namespace SolidEdgeSDK.InteropServices
{
    /// <summary>
    /// Default controller that handles connecting\disconnecting to COM events via IConnectionPointContainer and IConnectionPoint interfaces.
    /// </summary>
    public class ComEventsManager
    {
        public ComEventsManager(object sink)
        {
            if (sink == null) throw new ArgumentNullException(nameof(sink));
            Sink = sink;
        }

        /// <summary>
        /// Establishes a connection between a connection point object and the client's sink.
        /// </summary>
        /// <typeparam name="TInterface">Interface type of the outgoing interface whose connection point object is being requested.</typeparam>
        /// <param name="container">An object that implements the IConnectionPointContainer inferface.</param>
        public void Attach<TInterface>(object container) where TInterface : class
        {
            bool lockTaken = false;

            try
            {
                Monitor.Enter(this, ref lockTaken);

                // Prevent multiple event Advise() calls on same sink.
                if (IsAttached<TInterface>(container))
                {
                    return;
                }

                IConnectionPointContainer cpc = null;
                IConnectionPoint cp = null;
                int cookie = 0;

                cpc = (IConnectionPointContainer)container;
                cpc.FindConnectionPoint(typeof(TInterface).GUID, out cp);

                if (cp != null)
                {
                    cp.Advise(Sink, out cookie);
                    ConnectionPoints.Add(cp, cookie);
                }
            }
            finally
            {
                if (lockTaken)
                {
                    Monitor.Exit(this);
                }
            }
        }

        /// <summary>
        /// Determines if a connection between a connection point object and the client's sink is established.
        /// </summary>
        /// <param name="container">An object that implements the IConnectionPointContainer inferface.</param>
        public bool IsAttached<TInterface>(object container) where TInterface : class
        {
            bool lockTaken = false;

            try
            {
                Monitor.Enter(this, ref lockTaken);

                IConnectionPointContainer cpc = null;
                IConnectionPoint cp = null;
                int cookie = 0;

                cpc = (IConnectionPointContainer)container;
                cpc.FindConnectionPoint(typeof(TInterface).GUID, out cp);

                if (cp != null)
                {
                    if (ConnectionPoints.ContainsKey(cp))
                    {
                        cookie = ConnectionPoints[cp];
                        return true;
                    }
                }
            }
            finally
            {
                if (lockTaken)
                {
                    Monitor.Exit(this);
                }
            }

            return false;
        }

        /// <summary>
        /// Terminates an advisory connection previously established between a connection point object and a client's sink.
        /// </summary>
        /// <typeparam name="TInterface">Interface type of the interface whose connection point object is being requested to be removed.</typeparam>
        /// <param name="container">An object that implements the IConnectionPointContainer inferface.</param>
        public void Detach<TInterface>(object container) where TInterface : class
        {
            bool lockTaken = false;

            try
            {
                Monitor.Enter(this, ref lockTaken);

                IConnectionPointContainer cpc = null;
                IConnectionPoint cp = null;
                int cookie = 0;

                cpc = (IConnectionPointContainer)container;
                cpc.FindConnectionPoint(typeof(TInterface).GUID, out cp);

                if (cp != null)
                {
                    if (ConnectionPoints.ContainsKey(cp))
                    {
                        cookie = ConnectionPoints[cp];
                        cp.Unadvise(cookie);
                        ConnectionPoints.Remove(cp);
                    }
                }
            }
            finally
            {
                if (lockTaken)
                {
                    Monitor.Exit(this);
                }
            }
        }

        /// <summary>
        /// Terminates all advisory connections previously established.
        /// </summary>
        public void DetachAll()
        {
            bool lockTaken = false;

            try
            {
                Monitor.Enter(this, ref lockTaken);
                Dictionary<IConnectionPoint, int>.Enumerator enumerator = ConnectionPoints.GetEnumerator();
                while (enumerator.MoveNext())
                {
                    enumerator.Current.Key.Unadvise(enumerator.Current.Value);
                }
            }
            finally
            {
                ConnectionPoints.Clear();

                if (lockTaken)
                {
                    Monitor.Exit(this);
                }
            }
        }

        public object Sink { get; private set; }
        public Dictionary<IConnectionPoint, int> ConnectionPoints { get; private set; } = new Dictionary<IConnectionPoint, int>();
    }

    /// <summary>
    /// COM object wrapper class.
    /// </summary>
    public static class ComObject
    {
        const int LOCALE_SYSTEM_DEFAULT = 2048;

        /// <summary>
        /// Using IDispatch, returns the ITypeInfo of the specified object.
        /// </summary>
        /// <param name="comObject"></param>
        /// <returns></returns>
        public static ITypeInfo GetITypeInfo(object comObject)
        {
            if (Marshal.IsComObject(comObject) == false) throw new InvalidComObjectException();

            var dispatch = comObject as IDispatch;

            if (dispatch != null)
            {
                return dispatch.GetTypeInfo(0, LOCALE_SYSTEM_DEFAULT);
            }

            return null;
        }

        /// <summary>
        /// Returns a strongly typed property by name using the specified COM object.
        /// </summary>
        /// <typeparam name="T">The type of the property to return.</typeparam>
        /// <param name="comObject"></param>
        /// <param name="name">The name of the property to retrieve.</param>
        /// <returns></returns>
        public static T GetPropertyValue<T>(object comObject, string name)
        {
            if (Marshal.IsComObject(comObject) == false) throw new InvalidComObjectException();

            var type = comObject.GetType();
            var value = type.InvokeMember(name, System.Reflection.BindingFlags.GetProperty, null, comObject, null);

            return (T)value;
        }

        /// <summary>
        /// Returns a strongly typed property by name using the specified COM object.
        /// </summary>
        /// <typeparam name="T">The type of the property to return.</typeparam>
        /// <param name="comObject"></param>
        /// <param name="name">The name of the property to retrieve.</param>
        /// <param name="defaultValue">The value to return if the property does not exist.</param>
        /// <returns></returns>
        public static T GetPropertyValue<T>(object comObject, string name, T defaultValue)
        {
            if (Marshal.IsComObject(comObject) == false) throw new InvalidComObjectException();

            var type = comObject.GetType();

            try
            {
                var value = type.InvokeMember(name, System.Reflection.BindingFlags.GetProperty, null, comObject, null);
                return (T)value;
            }
            catch
            {
                return defaultValue;
            }
        }

        /// <summary>
        /// Using IDispatch, determine the managed type of the specified object.
        /// </summary>
        /// <param name="comObject"></param>
        /// <returns></returns>
        public static Type GetType(object comObject)
        {
            if (Marshal.IsComObject(comObject) == false) throw new InvalidComObjectException();

            Type type = null;
            var dispatch = comObject as IDispatch;
            ITypeInfo typeInfo = null;
            var pTypeAttr = IntPtr.Zero;
            var typeAttr = default(System.Runtime.InteropServices.ComTypes.TYPEATTR);

            try
            {
                if (dispatch != null)
                {
                    typeInfo = dispatch.GetTypeInfo(0, LOCALE_SYSTEM_DEFAULT);
                    typeInfo.GetTypeAttr(out pTypeAttr);
                    typeAttr = (System.Runtime.InteropServices.ComTypes.TYPEATTR)Marshal.PtrToStructure(pTypeAttr, typeof(System.Runtime.InteropServices.ComTypes.TYPEATTR));

                    // Type can technically be defined in any loaded assembly.
                    var assemblies = AppDomain.CurrentDomain.GetAssemblies();

                    // Scan each assembly for a type with a matching GUID.
                    foreach (var assembly in assemblies)
                    {
                        type = assembly.GetTypes()
                            .Where(x => x.IsInterface)
                            .Where(x => x.GUID.Equals(typeAttr.guid))
                            .FirstOrDefault();

                        if (type != null)
                        {
                            // Found what we're looking for so break out of the loop.
                            break;
                        }
                    }
                }
            }
            finally
            {
                if (typeInfo != null)
                {
                    typeInfo.ReleaseTypeAttr(pTypeAttr);
                    Marshal.ReleaseComObject(typeInfo);
                }
            }

            return type;
        }
    }

    [ComImport]
    [Guid("00020400-0000-0000-C000-000000000046")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    interface IDispatch
    {
        int GetTypeInfoCount();
        ITypeInfo GetTypeInfo(int iTInfo, int lcid);
    }
}

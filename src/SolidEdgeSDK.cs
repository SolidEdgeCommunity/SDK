//
// This file is maintained at https://github.com/SolidEdgeCommunity/SDK.
//
// Licensed under the MIT license.
// See https://github.com/SolidEdgeCommunity/SDK/blob/master/LICENSE for full
// license information.
//
// Required references:
//  - System.Core.dll
//  - System.Drawing.dll
//  - System.Windows.Forms.dll
//
// Required Solid Edge types from COM references:
//      - assembly.tlb
//      - constant.tlb
//      - draft.tlb
//      - framewrk.tlb
//      - fwksupp.tlb
//      - geometry.tlb
//      - Part.tlb
// It does not matter how you import these types to .NET. You can reference the
// type libraries directly from your project or pre generate the Interop
// Assemblies.

#region WPF Support

// Required references:
//  - PresentationCore.dll
//  - PresentationFramework.dll
//  - WindowsBase.dll

//#define ENABLE_WPF_SUPPORT

#endregion

using SolidEdgeSDK.Extensions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;

namespace SolidEdgeSDK.AddIn
{
    public class AddInDescriptor
    {
        public AddInDescriptor()
        {
        }

        public AddInDescriptor(System.Globalization.CultureInfo culture, Type resourceType)
        {
            Culture = culture;

            var resourceManager = new System.Resources.ResourceManager(resourceType.FullName, resourceType.Assembly);
            Description = resourceManager.GetString("AddInDescription", Culture);
            Summary = resourceManager.GetString("AddInSummary", Culture);
        }

        public System.Globalization.CultureInfo Culture { get; set; }
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
            Value = new System.Globalization.CultureInfo(culture);
        }

        public System.Globalization.CultureInfo Value { get; private set; }
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

        #region SolidEdgeFramework.ISolidEdgeAddIn

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

        #endregion

        #region EdgeBar

        public EdgeBarPage AddWinFormEdgeBarPage<TControl>(EdgeBarPageConfiguration config) where TControl : System.Windows.Forms.Control, new()
        {
            return AddWinFormEdgeBarPage<TControl>(
                config: config,
                document: null);
        }

        public EdgeBarPage AddWinFormEdgeBarPage<TControl>(EdgeBarPageConfiguration config, SolidEdgeFramework.SolidEdgeDocument document) where TControl : System.Windows.Forms.Control, new()
        {
            TControl control = Activator.CreateInstance<TControl>();

            var edgeBarPage = AddEdgeBarPage(
                config: config,
                controlHandle: control.Handle,
                document: document);

            edgeBarPage.ChildObject = control;

            return edgeBarPage;
        }

        public EdgeBarPage AddEdgeBarPage(EdgeBarPageConfiguration config, IntPtr controlHandle, SolidEdgeFramework.SolidEdgeDocument document)
        {
            if (config == null) throw new ArgumentNullException(nameof(config));

            var options = (int)config.GetOptions(document);
            var direction = (int)config.Direction;

            var hWndEdgeBarPage = ((SolidEdgeFramework.ISolidEdgeBarEx2)AddInInstance).AddPageEx2(
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

        public void RemoveAllEdgeBarPages()
        {
            var edgeBarPages = EdgeBarPages.ToArray();

            foreach (var edgeBarPage in edgeBarPages)
            {
                RemoveEdgeBarPage(edgeBarPage);
            }
        }

        public void RemoveEdgeBarPages(SolidEdgeFramework.SolidEdgeDocument document)
        {
            var edgeBarPages = EdgeBarPages
                .Where(x => x.Document == document)
                .ToArray();

            foreach (var edgeBarPage in edgeBarPages)
            {
                RemoveEdgeBarPage(edgeBarPage);
            }
        }

        private void RemoveEdgeBarPage(EdgeBarPage edgeBarPage)
        {
            int hWnd = edgeBarPage.Handle.ToInt32();

            try
            {
                ((SolidEdgeFramework.ISolidEdgeBarEx2)AddInInstance).RemovePage(edgeBarPage.Document, hWnd, 0);
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

        internal List<EdgeBarPage> EdgeBarPages { get; set; } = new List<EdgeBarPage>();

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

        #endregion

#if ENABLE_WPF_SUPPORT
        public EdgeBarPage AddWpfEdgeBarPage<TControl>(EdgeBarPageConfiguration config, SolidEdgeFramework.SolidEdgeDocument document = null) where TControl : System.Windows.Controls.Page, new()
        {
            uint WS_VISIBLE = 0x10000000;
            uint WS_CHILD = 0x40000000;
            uint WS_MAXIMIZE = 0x01000000;

            TControl control = Activator.CreateInstance<TControl>();

            var edgeBarPage = AddEdgeBarPage(
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
            Microsoft.Win32.Registry.ClassesRoot.DeleteSubKeyTree(subkey, false);
        }

        static Microsoft.Win32.RegistryKey CreateBaseKey(Guid guid)
        {
            string subkey = String.Join(@"\", "CLSID", guid.ToString("B"));
            return Microsoft.Win32.Registry.ClassesRoot.CreateSubKey(subkey);
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

        public Version SolidEdgeVersion { get { return new Version(Application.Version); ; } }
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
        public static System.Diagnostics.Process GetProcess(this SolidEdgeFramework.Application application)
        {
            return System.Diagnostics.Process.GetProcessById(application.ProcessID);
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
        public static void StartCommand(this SolidEdgeFramework.Application application, SolidEdgeConstants.AssemblyCommandConstants CommandID)
        {
            application.StartCommand((SolidEdgeFramework.SolidEdgeCommandConstants)CommandID);
        }

        /// <summary>
        /// Activates a specified Solid Edge command.
        /// </summary>
        public static void StartCommand(this SolidEdgeFramework.Application application, SolidEdgeConstants.CuttingPlaneLineCommandConstants CommandID)
        {
            application.StartCommand((SolidEdgeFramework.SolidEdgeCommandConstants)CommandID);
        }

        /// <summary>
        /// Activates a specified Solid Edge command.
        /// </summary>
        public static void StartCommand(this SolidEdgeFramework.Application application, SolidEdgeConstants.DetailCommandConstants CommandID)
        {
            application.StartCommand((SolidEdgeFramework.SolidEdgeCommandConstants)CommandID);
        }

        /// <summary>
        /// Activates a specified Solid Edge command.
        /// </summary>
        public static void StartCommand(this SolidEdgeFramework.Application application, SolidEdgeConstants.DrawingViewEditCommandConstants CommandID)
        {
            application.StartCommand((SolidEdgeFramework.SolidEdgeCommandConstants)CommandID);
        }

        /// <summary>
        /// Activates a specified Solid Edge command.
        /// </summary>
        public static void StartCommand(this SolidEdgeFramework.Application application, SolidEdgeConstants.ExplodeCommandConstants CommandID)
        {
            application.StartCommand((SolidEdgeFramework.SolidEdgeCommandConstants)CommandID);
        }

        /// <summary>
        /// Activates a specified Solid Edge command.
        /// </summary>
        public static void StartCommand(this SolidEdgeFramework.Application application, SolidEdgeConstants.LayoutCommandConstants CommandID)
        {
            application.StartCommand((SolidEdgeFramework.SolidEdgeCommandConstants)CommandID);
        }

        /// <summary>
        /// Activates a specified Solid Edge command.
        /// </summary>
        public static void StartCommand(this SolidEdgeFramework.Application application, SolidEdgeConstants.LayoutInPartCommandConstants CommandID)
        {
            application.StartCommand((SolidEdgeFramework.SolidEdgeCommandConstants)CommandID);
        }

        /// <summary>
        /// Activates a specified Solid Edge command.
        /// </summary>
        public static void StartCommand(this SolidEdgeFramework.Application application, SolidEdgeConstants.MotionCommandConstants CommandID)
        {
            application.StartCommand((SolidEdgeFramework.SolidEdgeCommandConstants)CommandID);
        }

        /// <summary>
        /// Activates a specified Solid Edge command.
        /// </summary>
        public static void StartCommand(this SolidEdgeFramework.Application application, SolidEdgeConstants.PartCommandConstants CommandID)
        {
            application.StartCommand((SolidEdgeFramework.SolidEdgeCommandConstants)CommandID);
        }

        /// <summary>
        /// Activates a specified Solid Edge command.
        /// </summary>
        public static void StartCommand(this SolidEdgeFramework.Application application, SolidEdgeConstants.ProfileCommandConstants CommandID)
        {
            application.StartCommand((SolidEdgeFramework.SolidEdgeCommandConstants)CommandID);
        }

        /// <summary>
        /// Activates a specified Solid Edge command.
        /// </summary>
        public static void StartCommand(this SolidEdgeFramework.Application application, SolidEdgeConstants.ProfileHoleCommandConstants CommandID)
        {
            application.StartCommand((SolidEdgeFramework.SolidEdgeCommandConstants)CommandID);
        }

        /// <summary>
        /// Activates a specified Solid Edge command.
        /// </summary>
        public static void StartCommand(this SolidEdgeFramework.Application application, SolidEdgeConstants.ProfilePatternCommandConstants CommandID)
        {
            application.StartCommand((SolidEdgeFramework.SolidEdgeCommandConstants)CommandID);
        }

        /// <summary>
        /// Activates a specified Solid Edge command.
        /// </summary>
        public static void StartCommand(this SolidEdgeFramework.Application application, SolidEdgeConstants.ProfileRevolvedCommandConstants CommandID)
        {
            application.StartCommand((SolidEdgeFramework.SolidEdgeCommandConstants)CommandID);
        }

        /// <summary>
        /// Activates a specified Solid Edge command.
        /// </summary>
        public static void StartCommand(this SolidEdgeFramework.Application application, SolidEdgeConstants.SheetMetalCommandConstants CommandID)
        {
            application.StartCommand((SolidEdgeFramework.SolidEdgeCommandConstants)CommandID);
        }

        /// <summary>
        /// Activates a specified Solid Edge command.
        /// </summary>
        public static void StartCommand(this SolidEdgeFramework.Application application, SolidEdgeConstants.SimplifyCommandConstants CommandID)
        {
            application.StartCommand((SolidEdgeFramework.SolidEdgeCommandConstants)CommandID);
        }

        /// <summary>
        /// Activates a specified Solid Edge command.
        /// </summary>
        public static void StartCommand(this SolidEdgeFramework.Application application, SolidEdgeConstants.StudioCommandConstants CommandID)
        {
            application.StartCommand((SolidEdgeFramework.SolidEdgeCommandConstants)CommandID);
        }

        /// <summary>
        /// Activates a specified Solid Edge command.
        /// </summary>
        public static void StartCommand(this SolidEdgeFramework.Application application, SolidEdgeConstants.TubingCommandConstants CommandID)
        {
            application.StartCommand((SolidEdgeFramework.SolidEdgeCommandConstants)CommandID);
        }

        /// <summary>
        /// Activates a specified Solid Edge command.
        /// </summary>
        public static void StartCommand(this SolidEdgeFramework.Application application, SolidEdgeConstants.WeldmentCommandConstants CommandID)
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
                System.Threading.Monitor.Enter(this, ref lockTaken);

                // Prevent multiple event Advise() calls on same sink.
                if (IsAttached<TInterface>(container))
                {
                    return;
                }

                System.Runtime.InteropServices.ComTypes.IConnectionPointContainer cpc = null;
                System.Runtime.InteropServices.ComTypes.IConnectionPoint cp = null;
                int cookie = 0;

                cpc = (System.Runtime.InteropServices.ComTypes.IConnectionPointContainer)container;
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
                    System.Threading.Monitor.Exit(this);
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
                System.Threading.Monitor.Enter(this, ref lockTaken);

                System.Runtime.InteropServices.ComTypes.IConnectionPointContainer cpc = null;
                System.Runtime.InteropServices.ComTypes.IConnectionPoint cp = null;
                int cookie = 0;

                cpc = (System.Runtime.InteropServices.ComTypes.IConnectionPointContainer)container;
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
                    System.Threading.Monitor.Exit(this);
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
                System.Threading.Monitor.Enter(this, ref lockTaken);

                System.Runtime.InteropServices.ComTypes.IConnectionPointContainer cpc = null;
                System.Runtime.InteropServices.ComTypes.IConnectionPoint cp = null;
                int cookie = 0;

                cpc = (System.Runtime.InteropServices.ComTypes.IConnectionPointContainer)container;
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
                    System.Threading.Monitor.Exit(this);
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
                System.Threading.Monitor.Enter(this, ref lockTaken);
                Dictionary<System.Runtime.InteropServices.ComTypes.IConnectionPoint, int>.Enumerator enumerator = ConnectionPoints.GetEnumerator();
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
                    System.Threading.Monitor.Exit(this);
                }
            }
        }

        public object Sink { get; private set; }
        public Dictionary<System.Runtime.InteropServices.ComTypes.IConnectionPoint, int> ConnectionPoints { get; private set; } = new Dictionary<System.Runtime.InteropServices.ComTypes.IConnectionPoint, int>();
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
        public static System.Runtime.InteropServices.ComTypes.ITypeInfo GetITypeInfo(object comObject)
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

        public static string GetComTypeFullName(object comObject)
        {
            if (Marshal.IsComObject(comObject) == false) throw new InvalidComObjectException();

            if (comObject is IDispatch dispatch)
            {
                var typeInfo = dispatch.GetTypeInfo(0, LOCALE_SYSTEM_DEFAULT);
                typeInfo.GetContainingTypeLib(out ITypeLib typeLib, out int pIndex);
                typeInfo.GetDocumentation(-1, out string typeName, out string typeDescription, out int typeHelpContext, out string typeHelpFile);
                typeLib.GetDocumentation(-1, out string typeLibName, out string typeLibDescription, out int typeLibHelpContext, out string typeLibHelpFile);

                return String.Join(".", typeLibName, typeName);
            }

            return "IUnknown";
        }
    }

    [ComImport]
    [Guid("00020400-0000-0000-C000-000000000046")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    interface IDispatch
    {
        int GetTypeInfoCount();
        System.Runtime.InteropServices.ComTypes.ITypeInfo GetTypeInfo(int iTInfo, int lcid);
    }

    /// <summary>
    /// IGL Graphics Library Interface. Ported from \sdk\include\igl.h.
    /// </summary>
    [ComImport]
    [Guid("0002D280-0000-0000-C000-000000000046")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IGL
    {
        void glAccum(uint op, float value);
        void glAlphaFunc(uint func, float aRef);
        void glBegin(uint mode);
        void glBitmap(int width, int height, float xorig, float yorig, float xmove, float ymove, [MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.U1)] byte[] bitmap);
        void glBlendFunc(uint sfactor, uint dfactor);
        void glCallList(uint list);
        void glCallLists(int n, uint type, [MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.U1)] byte[] lists);
        void glClear(uint mask);
        void glClearAccum(float red, float green, float blue, float alpha);
        void glClearColor(float red, float green, float blue, float alpha);
        void glClearDepth(double depth);
        void glClearIndex(float c);
        void glClearStencil(int s);
        void glClipPlane(uint plane, [MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.R8)] double[] equation);
        void glColor3b(sbyte red, sbyte green, sbyte blue);
        void glColor3bv([MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.I1)] sbyte[] v);
        void glColor3d(double red, double green, double blue);
        void glColor3dv([MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.R8)] double[] v);
        void glColor3f(float red, float green, float blue);
        void glColor3fv([MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.R4)] float[] v);
        void glColor3i(int red, int green, int blue);
        void glColor3iv([MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.I4)] int[] v);
        void glColor3s(short red, short green, short blue);
        void glColor3sv([MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.I2)] short[] v);
        void glColor3ub(byte red, byte green, byte blue);
        void glColor3ubv([MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.U1)] byte[] v);
        void glColor3ui(uint red, uint green, uint blue);
        void glColor3uiv([MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.U4)] uint[] v);
        void glColor3us(ushort red, ushort green, ushort blue);
        void glColor3usv([MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.U2)] ushort[] v);
        void glColor4b(sbyte red, sbyte green, sbyte blue, sbyte alpha);
        void glColor4bv([MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.I1)] sbyte[] v);
        void glColor4d(double red, double green, double blue, double alpha);
        void glColor4dv([MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.R8)] double[] v);
        void glColor4f(float red, float green, float blue, float alpha);
        void glColor4fv([MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.R4)] float[] v);
        void glColor4i(int red, int green, int blue, int alpha);
        void glColor4iv([MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.I4)] int[] v);
        void glColor4s(short red, short green, short blue, short alpha);
        void glColor4sv([MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.I2)] short[] v);
        void glColor4ub(byte red, byte green, byte blue, byte alpha);
        void glColor4ubv([MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.U1)] byte[] v);
        void glColor4ui(uint red, uint green, uint blue, uint alpha);
        void glColor4uiv([MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.U4)] uint[] v);
        void glColor4us(ushort red, ushort green, ushort blue, ushort alpha);
        void glColor4usv([MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.U2)] ushort[] v);
        void glColorMask(byte red, byte green, byte blue, byte alpha);
        void glColorMaterial(uint face, uint mode);
        void glCopyPixels(int x, int y, int width, int height, uint type);
        void glCullFace(uint mode);
        void glDeleteLists(uint list, int range);
        void glDepthFunc(uint func);
        void glDepthMask(byte flag);
        void glDepthRange(double zNear, double zFar);
        void glDisable(uint cap);
        void glDrawBuffer(uint mode);
        void glDrawPixels(int width, int height, uint format, uint type, [MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.U1)] byte[] pixels);
        void glEdgeFlag(byte flag);
        void glEdgeFlagv([MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.U1)] byte[] flag);
        void glEnable(uint cap);
        void glEnd();
        void glEndList();
        void glEvalCoord1d(double u);
        void glEvalCoord1dv([MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.R8)] double[] u);
        void glEvalCoord1f(float u);
        void glEvalCoord1fv([MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.R4)] float[] u);
        void glEvalCoord2d(double u, double v);
        void glEvalCoord2dv([MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.R8)] double[] u);
        void glEvalCoord2f(float u, float v);
        void glEvalCoord2fv([MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.R4)] float[] u);
        void glEvalMesh1(uint mode, int i1, int i2);
        void glEvalMesh2(uint mode, int i1, int i2, int j1, int j2);
        void glEvalPoint1(int i);
        void glEvalPoint2(int i, int j);
        void glFeedbackBuffer(int size, uint type, [MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.R4)] float[] buffer);
        void glFinish();
        void glFlush();
        void glFogf(uint pname, float param);
        void glFogfv(uint pname, [MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.R4)] float[] someParams);
        void glFogi(uint pname, int param);
        void glFogiv(uint pname, [MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.I4)] int[] someParams);
        void glFrontFace(uint mode);
        void glFrustum(double left, double right, double bottom, double top, double zNear, double zFar);
        uint glGenLists(int range);
        void glGetBooleanv(uint pname, [MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.U1)] byte[] someParams);
        void glGetClipPlane(uint plane, [MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.R8)] double[] equation);
        void glGetDoublev(uint pname, [MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.R8)] double[] someParams);
        uint glGetError();
        void glGetFloatv(uint pname, [MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.R4)] float[] someParams);
        void glGetIntegerv(uint pname, ref int someParams);
        void glGetLightfv(uint light, uint pname, [MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.R4)] float[] someParams);
        void glGetLightiv(uint light, uint pname, [MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.I4)] int[] someParams);
        void glGetMapdv(uint target, uint query, [MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.R8)] double[] v);
        void glGetMapfv(uint target, uint query, [MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.R4)] float[] v);
        void glGetMapiv(uint target, uint query, [MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.I4)] int[] v);
        void glGetMaterialfv(uint face, uint pname, [MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.R4)] float[] someParams);
        void glGetMaterialiv(uint face, uint pname, [MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.I4)] int[] someParams);
        void glGetPixelMapfv(uint map, [MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.R4)] float[] values);
        void glGetPixelMapuiv(uint map, [MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.U4)] uint[] values);
        void glGetPixelMapusv(uint map, [MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.U2)] ushort[] values);
        void glGetPolygonStipple([MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.U1)] byte[] mask);
        sbyte[] glGetString(uint name);
        void glGetTexEnvfv(uint target, uint pname, [MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.R4)] float[] someParams);
        void glGetTexEnviv(uint target, uint pname, [MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.I4)] int[] someParams);
        void glGetTexGendv(uint coord, uint pname, [MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.R8)] double[] someParams);
        void glGetTexGenfv(uint coord, uint pname, [MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.R4)] float[] someParams);
        void glGetTexGeniv(uint coord, uint pname, [MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.I4)] int[] someParams);
        void glGetTexImage(uint target, int level, uint format, uint type, [MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.U1)] byte[] pixels);
        void glGetTexLevelParameterfv(uint target, int level, uint pname, [MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.R4)] float[] someParams);
        void glGetTexLevelParameteriv(uint target, int level, uint pname, [MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.I4)] int[] someParams);
        void glGetTexParameterfv(uint target, uint pname, [MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.R4)] float[] someParams);
        void glGetTexParameteriv(uint target, uint pname, [MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.I4)] int[] someParams);
        void glHint(uint target, uint mode);
        void glIndexMask(uint mask);
        void glIndexd(double c);
        void glIndexdv([MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.R8)] double[] c);
        void glIndexf(float c);
        void glIndexfv([MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.R4)] float[] c);
        void glIndexi(int c);
        void glIndexiv([MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.I4)] int[] c);
        void glIndexs(short c);
        void glIndexsv([MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.I2)] short[] c);
        void glInitNames();
        byte glIsEnabled(uint cap);
        byte glIsList(uint list);
        void glLightModelf(uint pname, float param);
        void glLightModelfv(uint pname, [MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.R4)] float[] someParams);
        void glLightModeli(uint pname, int param);
        void glLightModeliv(uint pname, [MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.I4)] int[] someParams);
        void glLightf(uint light, uint pname, float param);
        void glLightfv(uint light, uint pname, [MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.R4)] float[] someParams);
        void glLighti(uint light, uint pname, int param);
        void glLightiv(uint light, uint pname, [MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.I4)] int[] someParams);
        void glLineStipple(int factor, ushort pattern);
        void glLineWidth(float width);
        void glListBase(uint aBase);
        void glLoadIdentity();
        void glLoadMatrixd([MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.R8)] double[] m);
        void glLoadMatrixf([MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.R4)] float[] m);
        void glLoadName(uint name);
        void glLogicOp(uint opcode);
        void glMap1d(uint target, double u1, double u2, int stride, int order, [MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.R8)] double[] points);
        void glMap1f(uint target, float u1, float u2, int stride, int order, [MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.R4)] float[] points);
        void glMap2d(uint target, double u1, double u2, int ustride, int uorder, double v1, double v2, int vstride, int vorder, [MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.R8)] double[] points);
        void glMap2f(uint target, float u1, float u2, int ustride, int uorder, float v1, float v2, int vstride, int vorder, [MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.R4)] float[] points);
        void glMapGrid1d(int un, double u1, double u2);
        void glMapGrid1f(int un, float u1, float u2);
        void glMapGrid2d(int un, double u1, double u2, int vn, double v1, double v2);
        void glMapGrid2f(int un, float u1, float u2, int vn, float v1, float v2);
        void glMaterialf(uint face, uint pname, float param);
        void glMaterialfv(uint face, uint pname, [MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.R4)] float[] someParams);
        void glMateriali(uint face, uint pname, int param);
        void glMaterialiv(uint face, uint pname, [MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.I4)] int[] someParams);
        void glMatrixMode(uint mode);
        void glMultMatrixd([MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.R8)] double[] m);
        void glMultMatrixf([MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.R4)] float[] m);
        void glNewList(uint list, uint mode);
        void glNormal3b(sbyte nx, sbyte ny, sbyte nz);
        void glNormal3bv([MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.I1)] sbyte[] v);
        void glNormal3d(double nx, double ny, double nz);
        void glNormal3dv([MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.R8)] double[] v);
        void glNormal3f(float nx, float ny, float nz);
        void glNormal3fv([MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.R4)] float[] v);
        void glNormal3i(int nx, int ny, int nz);
        void glNormal3iv([MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.I4)] int[] v);
        void glNormal3s(short nx, short ny, short nz);
        void glNormal3sv([MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.I2)] short[] v);
        void glOrtho(double left, double right, double bottom, double top, double zNear, double zFar);
        void glPassThrough(float token);
        void glPixelMapfv(uint map, int mapsize, [MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.R4)] float[] values);
        void glPixelMapuiv(uint map, int mapsize, [MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.U4)] uint[] values);
        void glPixelMapusv(uint map, int mapsize, [MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.U2)] ushort[] values);
        void glPixelStoref(uint pname, float param);
        void glPixelStorei(uint pname, int param);
        void glPixelTransferf(uint pname, float param);
        void glPixelTransferi(uint pname, int param);
        void glPixelZoom(float xfactor, float yfactor);
        void glPointSize(float size);
        void glPolygonMode(uint face, uint mode);
        void glPolygonStipple([MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.U1)] byte[] mask);
        void glPopAttrib();
        void glPopMatrix();
        void glPopName();
        void glPushAttrib(uint mask);
        void glPushMatrix();
        void glPushName(uint name);
        void glRasterPos2d(double x, double y);
        void glRasterPos2dv([MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.R8)] double[] v);
        void glRasterPos2f(float x, float y);
        void glRasterPos2fv([MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.R4)] float[] v);
        void glRasterPos2i(int x, int y);
        void glRasterPos2iv([MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.I4)] int[] v);
        void glRasterPos2s(short x, short y);
        void glRasterPos2sv([MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.I2)] short[] v);
        void glRasterPos3d(double x, double y, double z);
        void glRasterPos3dv([MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.R8)] double[] v);
        void glRasterPos3f(float x, float y, float z);
        void glRasterPos3fv([MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.R4)] float[] v);
        void glRasterPos3i(int x, int y, int z);
        void glRasterPos3iv([MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.I4)] int[] v);
        void glRasterPos3s(short x, short y, short z);
        void glRasterPos3sv([MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.I2)] short[] v);
        void glRasterPos4d(double x, double y, double z, double w);
        void glRasterPos4dv([MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.R8)] double[] v);
        void glRasterPos4f(float x, float y, float z, float w);
        void glRasterPos4fv([MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.R4)] float[] v);
        void glRasterPos4i(int x, int y, int z, int w);
        void glRasterPos4iv([MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.I4)] int[] v);
        void glRasterPos4s(short x, short y, short z, short w);
        void glRasterPos4sv([MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.I2)] short[] v);
        void glReadBuffer(uint mode);
        void glReadPixels(int x, int y, int width, int height, uint format, uint type, [MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.U1)] byte[] pixels);
        void glRectd(double x1, double y1, double x2, double y2);
        void glRectdv([MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.R8)] double[] v1, [MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.R8)] double[] v2);
        void glRectf(float x1, float y1, float x2, float y2);
        void glRectfv([MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.R4)] float[] v1, [MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.R4)] float[] v2);
        void glRecti(int x1, int y1, int x2, int y2);
        void glRectiv([MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.I4)] int[] v1, [MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.I4)] int[] v2);
        void glRects(short x1, short y1, short x2, short y2);
        void glRectsv([MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.I2)] short[] v1, [MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.I2)] short[] v2);
        int glRenderMode(uint mode);
        void glRotated(double angle, double x, double y, double z);
        void glRotatef(float angle, float x, float y, float z);
        void glScaled(double x, double y, double z);
        void glScalef(float x, float y, float z);
        void glScissor(int x, int y, int width, int height);
        void glSelectBuffer(int size, [MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.U4)] uint[] buffer);
        void glShadeModel(uint mode);
        void glStencilFunc(uint func, int aRef, uint mask);
        void glStencilMask(uint mask);
        void glStencilOp(uint fail, uint zfail, uint zpass);
        void glTexCoord1d(double s);
        void glTexCoord1dv([MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.R8)] double[] v);
        void glTexCoord1f(float s);
        void glTexCoord1fv([MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.R4)] float[] v);
        void glTexCoord1i(int s);
        void glTexCoord1iv([MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.I4)] int[] v);
        void glTexCoord1s(short s);
        void glTexCoord1sv([MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.I2)] short[] v);
        void glTexCoord2d(double s, double t);
        void glTexCoord2dv([MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.R8)] double[] v);
        void glTexCoord2f(float s, float t);
        void glTexCoord2fv([MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.R4)] float[] v);
        void glTexCoord2i(int s, int t);
        void glTexCoord2iv([MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.I4)] int[] v);
        void glTexCoord2s(short s, short t);
        void glTexCoord2sv([MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.I2)] short[] v);
        void glTexCoord3d(double s, double t, double r);
        void glTexCoord3dv([MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.R8)] double[] v);
        void glTexCoord3f(float s, float t, float r);
        void glTexCoord3fv([MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.R4)] float[] v);
        void glTexCoord3i(int s, int t, int r);
        void glTexCoord3iv([MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.I4)] int[] v);
        void glTexCoord3s(short s, short t, short r);
        void glTexCoord3sv([MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.I2)] short[] v);
        void glTexCoord4d(double s, double t, double r, double q);
        void glTexCoord4dv([MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.R8)] double[] v);
        void glTexCoord4f(float s, float t, float r, float q);
        void glTexCoord4fv([MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.R4)] float[] v);
        void glTexCoord4i(int s, int t, int r, int q);
        void glTexCoord4iv([MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.I4)] int[] v);
        void glTexCoord4s(short s, short t, short r, short q);
        void glTexCoord4sv([MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.I2)] short[] v);
        void glTexEnvf(uint target, uint pname, float param);
        void glTexEnvfv(uint target, uint pname, [MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.R4)] float[] someParams);
        void glTexEnvi(uint target, uint pname, int param);
        void glTexEnviv(uint target, uint pname, [MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.I4)] int[] someParams);
        void glTexGend(uint coord, uint pname, double param);
        void glTexGendv(uint coord, uint pname, [MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.R8)] double[] someParams);
        void glTexGenf(uint coord, uint pname, float param);
        void glTexGenfv(uint coord, uint pname, [MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.R4)] float[] someParams);
        void glTexGeni(uint coord, uint pname, int param);
        void glTexGeniv(uint coord, uint pname, [MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.I4)] int[] someParams);
        void glTexImage1D(uint target, int level, int components, int width, int border, uint format, uint type, [MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.U1)] byte[] pixels);
        void glTexImage2D(uint target, int level, int components, int width, int height, int border, uint format, uint type, [MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.U1)] byte[] pixels);
        void glTexParameterf(uint target, uint pname, float param);
        void glTexParameterfv(uint target, uint pname, [MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.R4)] float[] someParams);
        void glTexParameteri(uint target, uint pname, int param);
        void glTexParameteriv(uint target, uint pname, [MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.I4)] int[] someParams);
        void glTranslated(double x, double y, double z);
        void glTranslatef(float x, float y, float z);
        void glVertex2d(double x, double y);
        void glVertex2dv([MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.R8)] double[] v);
        void glVertex2f(float x, float y);
        void glVertex2fv([MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.R4)] float[] v);
        void glVertex2i(int x, int y);
        void glVertex2iv([MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.I4)] int[] v);
        void glVertex2s(short x, short y);
        void glVertex2sv([MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.I2)] short[] v);
        void glVertex3d(double x, double y, double z);
        void glVertex3dv([MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.R8)] double[] v);
        void glVertex3f(float x, float y, float z);
        void glVertex3fv([MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.R4)] float[] v);
        void glVertex3i(int x, int y, int z);
        void glVertex3iv([MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.I4)] int[] v);
        void glVertex3s(short x, short y, short z);
        void glVertex3sv([MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.I2)] short[] v);
        void glVertex4d(double x, double y, double z, double w);
        void glVertex4dv([MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.R8)] double[] v);
        void glVertex4f(float x, float y, float z, float w);
        void glVertex4fv([MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.R4)] float[] v);
        void glVertex4i(int x, int y, int z, int w);
        void glVertex4iv([MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.I4)] int[] v);
        void glVertex4s(short x, short y, short z, short w);
        void glVertex4sv([MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.I2)] short[] v);
        void glViewport(int x, int y, int width, int height);
        void glPolygonOffset(float factor, float units);
        void glBindTexture(uint target, uint texture);
    }

    /// <summary>
    /// IWGL Windows Graphics Layer Interface. Ported from \sdk\include\igl.h.
    /// </summary>
    [ComImport]
    [Guid("0002D282-0000-0000-C000-000000000046")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IWGL
    {
        IntPtr wglCreateContext(IntPtr hdc);
        int wglDeleteContext(IntPtr hglrc);
        IntPtr wglGetCurrentContext();
        IntPtr wglGetCurrentDC();
        int wglMakeCurrent(IntPtr hdc, IntPtr hglrc);
        int wglUseFontBitmapsA(IntPtr hDC, uint first, uint count, uint listbase);
        int wglUseFontBitmapsW(IntPtr hDC, uint first, uint count, uint listbase);
        int SwapBuffers(IntPtr hdc);
    }
}
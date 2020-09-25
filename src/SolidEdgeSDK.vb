'
' This file is maintained at https://github.com/SolidEdgeCommunity/SDK.
'
' Licensed under the MIT license.
' See https://github.com/SolidEdgeCommunity/SDK/blob/master/LICENSE for full
' license information.
'
' Required references:
'  - System.Core.dll
'  - System.Drawing.dll
'  - System.Windows.Forms.dll
'
' Required Solid Edge types from COM references:
'      - assembly.tlb
'      - constant.tlb
'      - draft.tlb
'      - framewrk.tlb
'      - fwksupp.tlb
'      - geometry.tlb
'      - Part.tlb
' It does not matter how you import these types to .NET. You can reference the
' type libraries directly from your project or pre generate the Interop
' Assemblies.

Imports SolidEdgeSDK.Extensions
Imports System
Imports System.Collections.Generic
Imports System.Linq
Imports System.Runtime.InteropServices
Imports System.Runtime.InteropServices.ComTypes
Imports System.Runtime.CompilerServices

Namespace SolidEdgeSDK.AddIn
    Public Class AddInDescriptor
        Public Sub New()
        End Sub

        Public Sub New(ByVal culture As System.Globalization.CultureInfo, ByVal resourceType As Type)
            Culture = culture
            Dim resourceManager = New System.Resources.ResourceManager(resourceType.FullName, resourceType.Assembly)
            Description = resourceManager.GetString("AddInDescription", Culture)
            Summary = resourceManager.GetString("AddInSummary", Culture)
        End Sub

        Public Property Culture As System.Globalization.CultureInfo
        Public Property Description As String
        Public Property Summary As String
    End Class

    <AttributeUsage(AttributeTargets.[Class], AllowMultiple:=True, Inherited:=False)>
    Public Class AddInEnvironmentCategoryAttribute
        Inherits System.Attribute

        Public Sub New(ByVal guid As String)
            Me.New(New Guid(guid))
        End Sub

        Public Sub New(ByVal guid As Guid)
            Value = guid
        End Sub

        Public Property Value As Guid
    End Class

    <AttributeUsage(AttributeTargets.[Class], AllowMultiple:=True, Inherited:=False)>
    Public Class AddInImplementedCategoryAttribute
        Inherits System.Attribute

        Public Sub New(ByVal guid As String)
            Me.New(New Guid(guid))
        End Sub

        Public Sub New(ByVal guid As Guid)
            Value = guid
        End Sub

        Public Property Value As Guid
    End Class

    <AttributeUsage(AttributeTargets.[Class], AllowMultiple:=True, Inherited:=False)>
    Public NotInheritable Class AddInCultureAttribute
        Inherits System.Attribute

        Public Sub New(ByVal culture As String)
            Value = New System.Globalization.CultureInfo(culture)
        End Sub

        Public Property Value As System.Globalization.CultureInfo
    End Class

    Public NotInheritable Class ComRegistrationSettings
        Public Sub New(ByVal type As Type)
            Me.Type = type
        End Sub

        Public Property Type As Type
        Public Property Enabled As Boolean
        Public Property ImplementedCategories As Guid() = New Guid() {}
        Public Property EnvironmentCategories As Guid() = New Guid() {}
        Public Property Descriptors As AddInDescriptor() = New AddInDescriptor() {}
    End Class

    Public Class EdgeBarPage
        Inherits System.Windows.Forms.NativeWindow
        Implements IDisposable

        <DllImport("user32.dll", SetLastError:=True)>
        Private Shared Function SetParent(ByVal hWndChild As IntPtr, ByVal hWndNewParent As IntPtr) As IntPtr
        <DllImport("user32.dll")>
        <MarshalAs(UnmanagedType.Bool)>
        Private Shared Function ShowWindow(ByVal hWnd As IntPtr, ByVal nCmdShow As Integer) As Boolean
        Private _disposed As Boolean = False
        Private _hWndChildWindow As IntPtr = IntPtr.Zero

        Public Sub New(ByVal hWnd As IntPtr, ByVal Optional index As Integer = 0)
            If hWnd = IntPtr.Zero Then Throw New System.ArgumentException($"{NameOf(hWnd)} cannot be IntPtr.Zero.")
            Index = index
            AssignHandle(hWnd)
        End Sub

        Public ReadOnly Property IsDisposed As Boolean
            Get
                Return _disposed
            End Get
        End Property

        Public Property ChildObject As Object

        Public Property ChildWindowHandle As IntPtr
            Get
                Return _hWndChildWindow
            End Get
            Set(ByVal value As IntPtr)
                _hWndChildWindow = value

                If _hWndChildWindow <> IntPtr.Zero Then
                    SetParent(_hWndChildWindow, Me.Handle)
                    ShowWindow(_hWndChildWindow, 3)
                End If
            End Set
        End Property

        Public Property Index As Integer
        Public Property Document As SolidEdgeFramework.SolidEdgeDocument
        Public Overridable Property Visible As Boolean = True

        Protected Overrides Sub Finalize()
            Dispose(False)
        End Sub

        Public Sub Dispose()
            Dispose(True)
            GC.SuppressFinalize(Me)
        End Sub

        Private Sub Dispose(ByVal disposing As Boolean)
            If Not _disposed Then

                If disposing Then

                    Try
                        (CType(Me.ChildObject, IDisposable))?.Dispose()
                    Catch
                    End Try
                End If

                Try
                    ReleaseHandle()
                Catch
                End Try

                ChildObject = Nothing
                _disposed = True
            End If
        End Sub
    End Class

    Public NotInheritable Class EdgeBarPageConfiguration
        Public Property Index As Integer
        Public Property NativeResourcesDllPath As String
        Public Property NativeImageId As Integer
        Public Property Tootip As String
        Public Property Title As String
        Public Property Caption As String
        Public Property Direction As EdgeBarPageDock = EdgeBarPageDock.Left
        Public Property MakeActive As Boolean = False
        Public Property ResizeChild As Boolean = True
        Public Property UpdateOnPaneSliding As Boolean = False
        Public Property WantsAccelerators As Boolean = False

        Public Function GetOptions(ByVal Optional document As SolidEdgeFramework.SolidEdgeDocument = Nothing) As SolidEdgeConstants.EdgeBarConstant
            Dim options = Nothing

            If document Is Nothing Then
                options = options Or SolidEdgeConstants.EdgeBarConstant.TRACK_CLOSE_GLOBALLY
            Else
                options = options Or SolidEdgeConstants.EdgeBarConstant.TRACK_CLOSE_BYDOCUMENT
            End If

            If MakeActive = False Then
                options = options Or SolidEdgeConstants.EdgeBarConstant.DONOT_MAKE_ACTIVE
            End If

            If ResizeChild = False Then
                options = options Or SolidEdgeConstants.EdgeBarConstant.NO_RESIZE_CHILD
            End If

            If UpdateOnPaneSliding = True Then
                options = options Or SolidEdgeConstants.EdgeBarConstant.UPDATE_ON_PANE_SLIDING
            End If

            If WantsAccelerators = True Then
                options = options Or SolidEdgeConstants.EdgeBarConstant.WANTS_ACCELERATORS
            End If

            Return options
        End Function
    End Class

    Public Enum EdgeBarPageDock
        Left = 0
        Rigth = 1
        Top = 2
        Button = 3
        None = 4
    End Enum

    Public MustInherit Partial Class SolidEdgeAddIn
        Inherits MarshalByRefObject
        Implements SolidEdgeFramework.ISolidEdgeAddIn

        Private Shared _instance As SolidEdgeAddIn
        Private _isolatedDomain As AppDomain
        Private _isolatedAddIn As SolidEdgeFramework.ISolidEdgeAddIn
        Private _ribbons As Dictionary(Of Guid, Ribbon) = New Dictionary(Of Guid, Ribbon)()

        Public Sub New()
            AddHandler AppDomain.CurrentDomain.AssemblyResolve, Function(sender, args) AppDomain.CurrentDomain.GetAssemblies().FirstOrDefault(Function(x) x.FullName = args.Name)
        End Sub

        Private Sub OnConnection(ByVal Application As Object, ByVal ConnectMode As SolidEdgeFramework.SeConnectMode, ByVal AddInInstance As SolidEdgeFramework.AddIn)
            Dim type = Me.[GetType]()

            If AppDomain.CurrentDomain.IsDefaultAppDomain() Then
                Dim appDomainSetup = AppDomain.CurrentDomain.SetupInformation
                Dim evidence = AppDomain.CurrentDomain.Evidence
                appDomainSetup.ApplicationBase = System.IO.Path.GetDirectoryName(type.Assembly.Location)
                Dim domainName = AddInInstance.GUID
                _isolatedDomain = AppDomain.CreateDomain(domainName, evidence, appDomainSetup)
                Dim isolatedAddIn = _isolatedDomain.CreateInstanceAndUnwrap(type.Assembly.FullName, type.FullName)
                _isolatedAddIn = CType(isolatedAddIn, SolidEdgeFramework.ISolidEdgeAddIn)

                If _isolatedAddIn IsNot Nothing Then
                    Dim application = CType(Application, SolidEdgeFramework.Application)
                    _isolatedAddIn.OnConnection(application, ConnectMode, AddInInstance)
                End If
            Else
                _instance = Me
                Me.Application = UnwrapTransparentProxy(Of SolidEdgeFramework.Application)(Application)
                Me.AddInInstance = UnwrapTransparentProxy(Of SolidEdgeFramework.AddIn)(AddInInstance)
                OnConnection(ConnectMode)
            End If
        End Sub

        Private Sub OnConnectToEnvironment(ByVal EnvCatID As String, ByVal pEnvironmentDispatch As Object, ByVal bFirstTime As Boolean)
            If (AppDomain.CurrentDomain.IsDefaultAppDomain()) AndAlso (_isolatedAddIn IsNot Nothing) Then
                _isolatedAddIn.OnConnectToEnvironment(EnvCatID, pEnvironmentDispatch, bFirstTime)
            Else
                Dim environment = UnwrapTransparentProxy(Of SolidEdgeFramework.Environment)(pEnvironmentDispatch)
                Dim environmentCategory = New Guid(EnvCatID)
                OnConnectToEnvironment(environmentCategory, environment, bFirstTime)
            End If
        End Sub

        Private Sub OnDisconnection(ByVal DisconnectMode As SolidEdgeFramework.SeDisconnectMode)
            If AppDomain.CurrentDomain.IsDefaultAppDomain() Then

                If _isolatedAddIn IsNot Nothing Then
                    _isolatedAddIn.OnDisconnection(DisconnectMode)

                    If _isolatedDomain IsNot Nothing Then
                        AppDomain.Unload(_isolatedDomain)
                    End If

                    _isolatedDomain = Nothing
                    _isolatedAddIn = Nothing
                End If
            Else
                OnDisconnection(DisconnectMode)
            End If
        End Sub

        Public Function AddWinFormEdgeBarPage(Of TControl As {System.Windows.Forms.Control, New})(ByVal config As EdgeBarPageConfiguration) As EdgeBarPage
            Return AddWinFormEdgeBarPage(Of TControl)(config:=config, document:=Nothing)
        End Function

        Public Function AddWinFormEdgeBarPage(Of TControl As {System.Windows.Forms.Control, New})(ByVal config As EdgeBarPageConfiguration, ByVal document As SolidEdgeFramework.SolidEdgeDocument) As EdgeBarPage
            Dim control As TControl = Activator.CreateInstance(Of TControl)()
            Dim edgeBarPage = AddEdgeBarPage(config:=config, controlHandle:=control.Handle, document:=document)
            edgeBarPage.ChildObject = control
            Return edgeBarPage
        End Function

        Public Function AddEdgeBarPage(ByVal config As EdgeBarPageConfiguration, ByVal controlHandle As IntPtr, ByVal document As SolidEdgeFramework.SolidEdgeDocument) As EdgeBarPage
            If config Is Nothing Then Throw New ArgumentNullException(NameOf(config))
            Dim options = CInt(config.GetOptions(document))
            Dim direction = CInt(config.Direction)
            Dim hWndEdgeBarPage = (CType(AddInInstance, SolidEdgeFramework.ISolidEdgeBarEx2)).AddPageEx2(theDocument:=document, ResourceFilename:=config.NativeResourcesDllPath, Index:=config.Index, nBitmapID:=config.NativeImageId, Tooltip:=config.Tootip, Title:=config.Title, Caption:=config.Caption, nOptions:=options, Direction:=direction)
            Dim edgeBarPage = New EdgeBarPage(New IntPtr(hWndEdgeBarPage), config.Index) With {
                .ChildWindowHandle = controlHandle,
                .Document = document
            }
            EdgeBarPages.Add(edgeBarPage)
            Return edgeBarPage
        End Function

        Public Sub RemoveAllEdgeBarPages()
            Dim edgeBarPages = EdgeBarPages.ToArray()

            For Each edgeBarPage In edgeBarPages
                RemoveEdgeBarPage(edgeBarPage)
            Next
        End Sub

        Public Sub RemoveEdgeBarPages(ByVal document As SolidEdgeFramework.SolidEdgeDocument)
            Dim edgeBarPages = EdgeBarPages.Where(Function(x) x.Document = document).ToArray()

            For Each edgeBarPage In edgeBarPages
                RemoveEdgeBarPage(edgeBarPage)
            Next
        End Sub

        Private Sub RemoveEdgeBarPage(ByVal edgeBarPage As EdgeBarPage)
            Dim hWnd As Integer = edgeBarPage.Handle.ToInt32()

            Try
                (CType(AddInInstance, SolidEdgeFramework.ISolidEdgeBarEx2)).RemovePage(edgeBarPage.Document, hWnd, 0)
            Catch
            End Try

            Try
                edgeBarPage.Dispose()
            Catch
            End Try

            EdgeBarPages.Remove(edgeBarPage)
        End Sub

        Friend Property EdgeBarPages As List(Of EdgeBarPage) = New List(Of EdgeBarPage)()

        Public ReadOnly Property DocumentEdgeBarPages As EdgeBarPage()
            Get
                Return EdgeBarPages.Where(Function(x) x.Document IsNot Nothing).ToArray()
            End Get
        End Property

        Public ReadOnly Property GlobalEdgeBarPages As EdgeBarPage()
            Get
                Return EdgeBarPages.Where(Function(x) x.Document Is Nothing).ToArray()
            End Get
        End Property

        Public Function AddRibbon(Of TRibbon As Ribbon)(ByVal environmentCategory As Guid, ByVal firstTime As Boolean) As TRibbon
            Dim ribbon As TRibbon = Activator.CreateInstance(Of TRibbon)()
            ribbon.EnvironmentCategory = environmentCategory
            ribbon.SolidEdgeAddIn = Me
            ribbon.Initialize()
            Add(ribbon, environmentCategory, firstTime)
            Return ribbon
        End Function

        Public Sub Add(ByVal ribbon As Ribbon, ByVal environmentCategory As Guid, ByVal firstTime As Boolean)
            If ribbon Is Nothing Then Throw New ArgumentNullException("ribbon")
            Dim addInEx = CType(AddInInstance, SolidEdgeFramework.ISEAddInEx)
            Dim addInEx2 = CType(Nothing, SolidEdgeFramework.ISEAddInEx2)
            Dim EnvironmentCatID = environmentCategory.ToString("B")

            If _ribbons.ContainsKey(ribbon.EnvironmentCategory) Then
                Throw New System.Exception(String.Format("A ribbon has already been added for environment category {0}.", ribbon.EnvironmentCategory))
            End If

            If ribbon.EnvironmentCategory.Equals(Guid.Empty) Then
                Throw New System.Exception(String.Format("{0} is not a valid environment category.", ribbon.EnvironmentCategory))
            End If

            For Each tab In ribbon.Tabs

                For Each group In tab.Groups

                    For Each control In group.Controls
                        Dim categoryName = tab.Name
                        Dim commandName = New System.Text.StringBuilder()
                        commandName.AppendFormat("{0}_{1}", Me.Guid.ToString(), control.CommandId)

                        If TypeOf control Is RibbonButton Then
                            Dim ribbonButton = TryCast(control, RibbonButton)

                            If String.IsNullOrEmpty(ribbonButton.DropDownGroup) = False Then
                                commandName.AppendFormat(vbLf & "{0}\\{1}" & vbLf & "{2}" & vbLf & "{3}", ribbonButton.DropDownGroup, control.Label, control.SuperTip, control.ScreenTip)
                            Else
                                commandName.AppendFormat(vbLf & "{0}" & vbLf & "{1}" & vbLf & "{2}", control.Label, control.SuperTip, control.ScreenTip)
                            End If
                        Else
                            commandName.AppendFormat(vbLf & "{0}" & vbLf & "{1}" & vbLf & "{2}", control.Label, control.SuperTip, control.ScreenTip)
                        End If

                        If Not String.IsNullOrEmpty(control.Macro) Then
                            commandName.AppendFormat(vbLf & "{0}", control.Macro)

                            If Not String.IsNullOrEmpty(control.MacroParameters) Then
                                commandName.AppendFormat(vbLf & "{0}", control.MacroParameters)
                            End If
                        End If

                        control.CommandName = commandName.ToString()
                        Dim commandNames As Array = New String() {control.CommandName}
                        Dim commandIDs As Array = New Integer() {control.CommandId}

                        If addInEx2 IsNot Nothing Then
                            categoryName = String.Format("{0}" & vbLf & "{1}", tab.Name, group.Name)
                            Dim commandButtonStyles As Array = New SolidEdgeFramework.SeButtonStyle() {control.Style}
                            addInEx2.SetAddInInfoEx2(NativeResourcesDllPath, EnvironmentCatID, categoryName, control.ImageId, -1, -1, -1, commandNames.Length, commandNames, commandIDs, commandButtonStyles)
                        ElseIf addInEx IsNot Nothing Then
                            addInEx.SetAddInInfoEx(NativeResourcesDllPath, EnvironmentCatID, categoryName, control.ImageId, -1, -1, -1, commandNames.Length, commandNames, commandIDs)

                            If firstTime Then
                                Dim commandBarName = String.Format("{0}" & vbLf & "{1}", tab.Name, group.Name)
                                Dim pButton As SolidEdgeFramework.CommandBarButton = addInEx.AddCommandBarButton(EnvironmentCatID, commandBarName, control.CommandId)

                                If pButton IsNot Nothing Then
                                    pButton.Style = control.Style
                                End If
                            End If
                        End If

                        control.SolidEdgeCommandId = CInt(commandIDs.GetValue(0))
                    Next
                Next
            Next

            _ribbons.Add(ribbon.EnvironmentCategory, ribbon)
        End Sub

        Public ReadOnly Property ActiveRibbon As Ribbon
            Get
                Dim environment = Me.Application.GetActiveEnvironment()
                Dim envCatId = Guid.Parse(environment.CATID)
                Dim value As Ribbon = Nothing

                If _ribbons.TryGetValue(envCatId, value) Then
                    Return value
                End If

                Return Nothing
            End Get
        End Property

        Public ReadOnly Property Ribbons As IEnumerable(Of Ribbon)
            Get
                Return _ribbons.Values
            End Get
        End Property

        Public MustOverride Sub OnConnection(ByVal ConnectMode As SolidEdgeFramework.SeConnectMode)
        Public MustOverride Sub OnConnectToEnvironment(ByVal EnvCatID As Guid, ByVal environment As SolidEdgeFramework.Environment, ByVal firstTime As Boolean)
        Public MustOverride Sub OnDisconnection(ByVal DisconnectMode As SolidEdgeFramework.SeDisconnectMode)

        Public Overrides Function InitializeLifetimeService() As Object
            Return Nothing
        End Function

        Private Function UnwrapTransparentProxy(Of TInterface As Class)(ByVal rcw As Object) As TInterface
            If System.Runtime.Remoting.RemotingServices.IsTransparentProxy(rcw) Then
                Dim punk As IntPtr = Marshal.GetIUnknownForObject(rcw)

                Try
                    Return CType(Marshal.GetObjectForIUnknown(punk), TInterface)
                Finally
                    Marshal.Release(punk)
                End Try
            End If

            Return TryCast(rcw, TInterface)
        End Function

        Public Shared Sub ComRegisterSolidEdgeAddIn(ByVal settings As ComRegistrationSettings)
            Dim type = settings.Type

            Using baseKey = CreateBaseKey(type.GUID)
                baseKey.SetValue("AutoConnect", If(settings.Enabled, 1, 0))

                For Each guid In settings.ImplementedCategories
                    Dim subkey = String.Join("\", "Implemented Categories", guid.ToString("B"))
                    baseKey.CreateSubKey(subkey)
                Next

                For Each guid In settings.EnvironmentCategories
                    Dim subkey = String.Join("\", "Environment Categories", guid.ToString("B"))
                    baseKey.CreateSubKey(subkey)
                Next

                For Each descriptor In settings.Descriptors
                    Dim lcid = descriptor.Culture.LCID.ToString("X4")
                    lcid = lcid.TrimStart("0"c)
                    baseKey.SetValue(lcid, descriptor.Description)

                    Using summaryKey = baseKey.CreateSubKey("Summary")
                        summaryKey.SetValue(lcid, descriptor.Summary)
                    End Using
                Next
            End Using
        End Sub

        Public Shared Sub ComUnregisterSolidEdgeAddIn(ByVal t As Type)
            Dim subkey As String = String.Join("\", "CLSID", t.GUID.ToString("B"))
            Microsoft.Win32.Registry.ClassesRoot.DeleteSubKeyTree(subkey, False)
        End Sub

        Private Shared Function CreateBaseKey(ByVal guid As Guid) As Microsoft.Win32.RegistryKey
            Dim subkey As String = String.Join("\", "CLSID", guid.ToString("B"))
            Return Microsoft.Win32.Registry.ClassesRoot.CreateSubKey(subkey)
        End Function

        Public Shared ReadOnly Property Instance As SolidEdgeAddIn
            Get
                Return _instance
            End Get
        End Property

        Public Property Application As SolidEdgeFramework.Application
        Public Property AddInInstance As SolidEdgeFramework.AddIn

        Public Overridable ReadOnly Property FileInfo As System.IO.FileInfo
            Get
                Return New System.IO.FileInfo(Me.[GetType]().Assembly.Location)
            End Get
        End Property

        Public ReadOnly Property Guid As Guid
            Get
                Return New Guid(AddInInstance.GUID)
            End Get
        End Property

        Public ReadOnly Property GuiVersion As Integer
            Get
                Return AddInInstance.GuiVersion
            End Get
        End Property

        Public Overridable ReadOnly Property NativeResourcesDllPath As String
            Get
                Return Me.[GetType]().Assembly.Location
            End Get
        End Property

        Public ReadOnly Property SolidEdgeVersion As Version
            Get
                Return New Version(Application.Version)
            End Get
        End Property
    End Class

    Public Class RibbonTab
        Friend Sub New(ByVal ribbon As Ribbon, ByVal name As String)
            Ribbon = ribbon
            Name = name
        End Sub

        Public Property Ribbon As Ribbon
        Public Property Name As String
        Public Property Groups As RibbonGroup() = New RibbonGroup() {}

        Public ReadOnly Iterator Property Controls As IEnumerable(Of RibbonControl)
            Get

                For Each group In Me.Groups

                    For Each control In group.Controls
                        Yield control
                    Next
                Next
            End Get
        End Property

        Public ReadOnly Property Buttons As IEnumerable(Of RibbonButton)
            Get
                Return Me.Controls.OfType(Of RibbonButton)()
            End Get
        End Property

        Public ReadOnly Property CheckBoxes As IEnumerable(Of RibbonCheckBox)
            Get
                Return Me.Controls.OfType(Of RibbonCheckBox)()
            End Get
        End Property

        Public ReadOnly Property RadioButtons As IEnumerable(Of RibbonRadioButton)
            Get
                Return Me.Controls.OfType(Of RibbonRadioButton)()
            End Get
        End Property

        Public Overrides Function ToString() As String
            Return Name
        End Function
    End Class

    Public Class RibbonGroup
        Private _controls As List(Of RibbonControl) = New List(Of RibbonControl)()

        Public Sub New(ByVal name As String)
            Name = name
        End Sub

        Public Property Name As String
        Public Property Controls As RibbonControl() = New RibbonControl() {}

        Public ReadOnly Property Buttons As IEnumerable(Of RibbonButton)
            Get
                Return Me.Controls.OfType(Of RibbonButton)()
            End Get
        End Property

        Public ReadOnly Property CheckBoxes As IEnumerable(Of RibbonCheckBox)
            Get
                Return Me.Controls.OfType(Of RibbonCheckBox)()
            End Get
        End Property

        Public ReadOnly Property RadioButtons As IEnumerable(Of RibbonRadioButton)
            Get
                Return Me.Controls.OfType(Of RibbonRadioButton)()
            End Get
        End Property

        Public Overrides Function ToString() As String
            Return Name
        End Function
    End Class

    <Serializable>
    Public Delegate Sub RibbonControlClickEventHandler(ByVal control As RibbonControl)
    <Serializable>
    Public Delegate Sub RibbonControlHelpEventHandler(ByVal control As RibbonControl, ByVal hWndFrame As IntPtr, ByVal helpCommandID As Integer)

    Public MustInherit Class RibbonControl
        Friend Sub New(ByVal commandId As Integer)
            CommandId = commandId
        End Sub

        Public Property UseDotMark As Boolean
        Public Property Checked As Boolean
        Public Property CommandId As Integer
        Public Property CommandName As String
        Public Property Enabled As Boolean = True
        Public Property Group As RibbonGroup
        Public Property ImageId As Integer
        Public Property Label As String
        Public Property Macro As String
        Public Property MacroParameters As String
        Public Property ScreenTip As String
        Public Property ShowImage As Boolean
        Public Property ShowLabel As Boolean
        Public Property SolidEdgeCommandId As Integer
        Friend MustOverride ReadOnly Property Style As SolidEdgeFramework.SeButtonStyle
        Public Property SuperTip As String
        Public Property WebHelpURL As String
    End Class

    Public Class RibbonButton
        Inherits RibbonControl

        Public Sub New(ByVal id As Integer)
            MyBase.New(id)
        End Sub

        Public Sub New(ByVal id As Integer, ByVal label As String, ByVal screentip As String, ByVal supertip As String, ByVal imageId As Integer, ByVal Optional size As RibbonButtonSize = RibbonButtonSize.Normal)
            MyBase.New(id)
            Label = label
            ScreenTip = screentip
            SuperTip = supertip
            ImageId = imageId
            Size = size
        End Sub

        Public Property DropDownGroup As String
        Public Property Size As RibbonButtonSize = RibbonButtonSize.Normal

        Friend Overrides ReadOnly Property Style As SolidEdgeFramework.SeButtonStyle
            Get
                Dim style = SolidEdgeFramework.SeButtonStyle.seButtonAutomatic

                If Size = RibbonButtonSize.Large Then
                    style = SolidEdgeFramework.SeButtonStyle.seButtonIconAndCaptionBelow
                Else

                    If ShowImage AndAlso ShowLabel Then
                        style = SolidEdgeFramework.SeButtonStyle.seButtonIconAndCaption
                    ElseIf ShowImage Then
                        style = SolidEdgeFramework.SeButtonStyle.seButtonIcon
                    ElseIf ShowLabel Then
                        style = SolidEdgeFramework.SeButtonStyle.seButtonCaption
                    End If
                End If

                Return style
            End Get
        End Property
    End Class

    Public Class RibbonCheckBox
        Inherits RibbonControl

        Public Sub New(ByVal id As Integer)
            MyBase.New(id)
        End Sub

        Public Sub New(ByVal id As Integer, ByVal label As String, ByVal screentip As String, ByVal supertip As String)
            MyBase.New(id)
            Label = label
            ScreenTip = screentip
            SuperTip = supertip
        End Sub

        Friend Overrides ReadOnly Property Style As SolidEdgeFramework.SeButtonStyle
            Get
                Dim style = SolidEdgeFramework.SeButtonStyle.seCheckButton

                If Me.ImageId >= 0 Then
                    style = If(ShowImage = True, SolidEdgeFramework.SeButtonStyle.seCheckButtonAndIcon, SolidEdgeFramework.SeButtonStyle.seCheckButton)
                End If

                Return style
            End Get
        End Property
    End Class

    Public Class RibbonRadioButton
        Inherits RibbonControl

        Public Sub New(ByVal id As Integer)
            MyBase.New(id)
        End Sub

        Public Sub New(ByVal id As Integer, ByVal label As String, ByVal screentip As String, ByVal supertip As String)
            MyBase.New(id)
            Label = label
            ScreenTip = screentip
            SuperTip = supertip
        End Sub

        Friend Overrides ReadOnly Property Style As SolidEdgeFramework.SeButtonStyle
            Get
                Return SolidEdgeFramework.SeButtonStyle.seRadioButton
            End Get
        End Property
    End Class

    Public Enum RibbonButtonSize
        Normal
        Large
    End Enum

    Public MustInherit Class Ribbon
        Public MustOverride Sub Initialize()

        Protected Sub Initialize(ByVal tabs As RibbonTab())
            If Tabs.Any() Then
                Throw New System.Exception($"{Me.[GetType]().FullName} has already been initialized.")
            End If

            Dim duplicateCommandIds = tabs.SelectMany(Function(x) x.Controls).[Select](Function(x) x.CommandId).GroupBy(Function(x) x).Where(Function(x) x.Count() > 1)

            If duplicateCommandIds.Any() Then
                Throw New System.Exception($"Duplicate command IDs detected. The duplicated command IDs are: '{String.Join(",", duplicateCommandIds.[Select](Function(x) x.Key))}'.")
            End If

            Tabs = tabs
        End Sub

        Public Function GetButton(ByVal commandId As Integer) As RibbonButton
            Return Buttons.FirstOrDefault(Function(x) x.CommandId = commandId)
        End Function

        Public Function GetCheckBox(ByVal commandId As Integer) As RibbonCheckBox
            Return CheckBoxes.FirstOrDefault(Function(x) x.CommandId = commandId)
        End Function

        Public Function GetControl(ByVal commandId As Integer) As RibbonControl
            Return Controls.FirstOrDefault(Function(x) x.CommandId = commandId)
        End Function

        Public Function GetControl(Of TRibbonControl As RibbonControl)(ByVal commandId As Integer) As TRibbonControl
            Return Controls.OfType(Of TRibbonControl)().FirstOrDefault(Function(x) x.CommandId = commandId)
        End Function

        Public Function GetRadioButton(ByVal commandId As Integer) As RibbonRadioButton
            Return RadioButtons.FirstOrDefault(Function(x) x.CommandId = commandId)
        End Function

        Public ReadOnly Iterator Property Controls As IEnumerable(Of RibbonControl)
            Get

                For Each tab In Tabs

                    For Each control In tab.Controls
                        Yield control
                    Next
                Next
            End Get
        End Property

        Public Property SolidEdgeAddIn As SolidEdgeAddIn
        Public Property EnvironmentCategory As Guid

        Public ReadOnly Property Buttons As IEnumerable(Of RibbonButton)
            Get
                Return Me.Controls.OfType(Of RibbonButton)()
            End Get
        End Property

        Public ReadOnly Property CheckBoxes As IEnumerable(Of RibbonCheckBox)
            Get
                Return Me.Controls.OfType(Of RibbonCheckBox)()
            End Get
        End Property

        Public ReadOnly Property RadioButtons As IEnumerable(Of RibbonRadioButton)
            Get
                Return Me.Controls.OfType(Of RibbonRadioButton)()
            End Get
        End Property

        Default Public ReadOnly Property Item(ByVal commandId As Integer) As RibbonControl
            Get
                Return Me.Controls.Where(Function(x) x.CommandId = commandId).FirstOrDefault()
            End Get
        End Property

        Public Property Tabs As RibbonTab() = New RibbonTab() {}
    End Class
End Namespace

Namespace SolidEdgeSDK.Extensions
    Module ApplicationExtensions
        <Extension()>
        Function GetActiveDocument(ByVal application As SolidEdgeFramework.Application, ByVal Optional throwOnError As Boolean = True) As SolidEdgeFramework.SolidEdgeDocument
            Try
                Return CType(application.ActiveDocument, SolidEdgeFramework.SolidEdgeDocument)
            Catch
                If throwOnError Then Throw
            End Try

            Return Nothing
        End Function

        <Extension()>
        Function GetActiveDocument(Of T As Class)(ByVal application As SolidEdgeFramework.Application, ByVal Optional throwOnError As Boolean = True) As T
            Try
                Return CType(application.ActiveDocument, T)
            Catch
                If throwOnError Then Throw
            End Try

            Return Nothing
        End Function

        <Extension()>
        Function GetActiveEnvironment(ByVal application As SolidEdgeFramework.Application) As SolidEdgeFramework.Environment
            Dim environments As SolidEdgeFramework.Environments = application.Environments
            Return environments.Item(application.ActiveEnvironment)
        End Function

        <Extension()>
        Function GetEnvironment(ByVal application As SolidEdgeFramework.Application, ByVal CATID As String) As SolidEdgeFramework.Environment
            Dim guid1 = New Guid(CATID)

            For Each environment In application.Environments.OfType(Of SolidEdgeFramework.Environment)()
                Dim guid2 = New Guid(environment.CATID)

                If guid1.Equals(guid2) Then
                    Return environment
                End If
            Next

            Return Nothing
        End Function

        <Extension()>
        Function GetGlobalParameter(ByVal application As SolidEdgeFramework.Application, ByVal globalConstant As SolidEdgeFramework.ApplicationGlobalConstants) As Object
            Dim value As Object = Nothing
            application.GetGlobalParameter(globalConstant, value)
            Return value
        End Function

        <Extension()>
        Function GetGlobalParameter(Of T)(ByVal application As SolidEdgeFramework.Application, ByVal globalConstant As SolidEdgeFramework.ApplicationGlobalConstants) As T
            Dim value As Object = Nothing
            application.GetGlobalParameter(globalConstant, value)
            Return CType(value, T)
        End Function

        <Extension()>
        Function GetNativeWindow(ByVal application As SolidEdgeFramework.Application) As System.Windows.Forms.NativeWindow
            Return System.Windows.Forms.NativeWindow.FromHandle(New IntPtr(application.hWnd))
        End Function

        <Extension()>
        Function GetProcess(ByVal application As SolidEdgeFramework.Application) As System.Diagnostics.Process
            Return System.Diagnostics.Process.GetProcessById(application.ProcessID)
        End Function

        <Extension()>
        Function GetVersion(ByVal application As SolidEdgeFramework.Application) As Version
            Return New Version(application.Version)
        End Function

        <Extension()>
        Function GetWindowHandle(ByVal application As SolidEdgeFramework.Application) As IntPtr
            Return New IntPtr(application.hWnd)
        End Function

        <Extension()>
        Sub StartCommand(ByVal application As SolidEdgeFramework.Application, ByVal CommandID As SolidEdgeConstants.AssemblyCommandConstants)
            application.StartCommand(CType(CommandID, SolidEdgeFramework.SolidEdgeCommandConstants))
        End Sub

        <Extension()>
        Sub StartCommand(ByVal application As SolidEdgeFramework.Application, ByVal CommandID As SolidEdgeConstants.CuttingPlaneLineCommandConstants)
            application.StartCommand(CType(CommandID, SolidEdgeFramework.SolidEdgeCommandConstants))
        End Sub

        <Extension()>
        Sub StartCommand(ByVal application As SolidEdgeFramework.Application, ByVal CommandID As SolidEdgeConstants.DetailCommandConstants)
            application.StartCommand(CType(CommandID, SolidEdgeFramework.SolidEdgeCommandConstants))
        End Sub

        <Extension()>
        Sub StartCommand(ByVal application As SolidEdgeFramework.Application, ByVal CommandID As SolidEdgeConstants.DrawingViewEditCommandConstants)
            application.StartCommand(CType(CommandID, SolidEdgeFramework.SolidEdgeCommandConstants))
        End Sub

        <Extension()>
        Sub StartCommand(ByVal application As SolidEdgeFramework.Application, ByVal CommandID As SolidEdgeConstants.ExplodeCommandConstants)
            application.StartCommand(CType(CommandID, SolidEdgeFramework.SolidEdgeCommandConstants))
        End Sub

        <Extension()>
        Sub StartCommand(ByVal application As SolidEdgeFramework.Application, ByVal CommandID As SolidEdgeConstants.LayoutCommandConstants)
            application.StartCommand(CType(CommandID, SolidEdgeFramework.SolidEdgeCommandConstants))
        End Sub

        <Extension()>
        Sub StartCommand(ByVal application As SolidEdgeFramework.Application, ByVal CommandID As SolidEdgeConstants.LayoutInPartCommandConstants)
            application.StartCommand(CType(CommandID, SolidEdgeFramework.SolidEdgeCommandConstants))
        End Sub

        <Extension()>
        Sub StartCommand(ByVal application As SolidEdgeFramework.Application, ByVal CommandID As SolidEdgeConstants.MotionCommandConstants)
            application.StartCommand(CType(CommandID, SolidEdgeFramework.SolidEdgeCommandConstants))
        End Sub

        <Extension()>
        Sub StartCommand(ByVal application As SolidEdgeFramework.Application, ByVal CommandID As SolidEdgeConstants.PartCommandConstants)
            application.StartCommand(CType(CommandID, SolidEdgeFramework.SolidEdgeCommandConstants))
        End Sub

        <Extension()>
        Sub StartCommand(ByVal application As SolidEdgeFramework.Application, ByVal CommandID As SolidEdgeConstants.ProfileCommandConstants)
            application.StartCommand(CType(CommandID, SolidEdgeFramework.SolidEdgeCommandConstants))
        End Sub

        <Extension()>
        Sub StartCommand(ByVal application As SolidEdgeFramework.Application, ByVal CommandID As SolidEdgeConstants.ProfileHoleCommandConstants)
            application.StartCommand(CType(CommandID, SolidEdgeFramework.SolidEdgeCommandConstants))
        End Sub

        <Extension()>
        Sub StartCommand(ByVal application As SolidEdgeFramework.Application, ByVal CommandID As SolidEdgeConstants.ProfilePatternCommandConstants)
            application.StartCommand(CType(CommandID, SolidEdgeFramework.SolidEdgeCommandConstants))
        End Sub

        <Extension()>
        Sub StartCommand(ByVal application As SolidEdgeFramework.Application, ByVal CommandID As SolidEdgeConstants.ProfileRevolvedCommandConstants)
            application.StartCommand(CType(CommandID, SolidEdgeFramework.SolidEdgeCommandConstants))
        End Sub

        <Extension()>
        Sub StartCommand(ByVal application As SolidEdgeFramework.Application, ByVal CommandID As SolidEdgeConstants.SheetMetalCommandConstants)
            application.StartCommand(CType(CommandID, SolidEdgeFramework.SolidEdgeCommandConstants))
        End Sub

        <Extension()>
        Sub StartCommand(ByVal application As SolidEdgeFramework.Application, ByVal CommandID As SolidEdgeConstants.SimplifyCommandConstants)
            application.StartCommand(CType(CommandID, SolidEdgeFramework.SolidEdgeCommandConstants))
        End Sub

        <Extension()>
        Sub StartCommand(ByVal application As SolidEdgeFramework.Application, ByVal CommandID As SolidEdgeConstants.StudioCommandConstants)
            application.StartCommand(CType(CommandID, SolidEdgeFramework.SolidEdgeCommandConstants))
        End Sub

        <Extension()>
        Sub StartCommand(ByVal application As SolidEdgeFramework.Application, ByVal CommandID As SolidEdgeConstants.TubingCommandConstants)
            application.StartCommand(CType(CommandID, SolidEdgeFramework.SolidEdgeCommandConstants))
        End Sub

        <Extension()>
        Sub StartCommand(ByVal application As SolidEdgeFramework.Application, ByVal CommandID As SolidEdgeConstants.WeldmentCommandConstants)
            application.StartCommand(CType(CommandID, SolidEdgeFramework.SolidEdgeCommandConstants))
        End Sub

        <Extension()>
        Sub Show(ByVal application As SolidEdgeFramework.Application, ByVal form As System.Windows.Forms.Form)
            If form Is Nothing Then Throw New ArgumentNullException("form")
            form.Show(application.GetNativeWindow())
        End Sub

        <Extension()>
        Function ShowDialog(ByVal application As SolidEdgeFramework.Application, ByVal form As System.Windows.Forms.Form) As System.Windows.Forms.DialogResult
            If form Is Nothing Then Throw New ArgumentNullException("form")
            Return form.ShowDialog(application.GetNativeWindow())
        End Function

        <Extension()>
        Function ShowDialog(ByVal application As SolidEdgeFramework.Application, ByVal dialog As System.Windows.Forms.CommonDialog) As System.Windows.Forms.DialogResult
            If dialog Is Nothing Then Throw New ArgumentNullException("dialog")
            Return dialog.ShowDialog(application.GetNativeWindow())
        End Function

        <Extension()>
        Sub UpdateActiveWindow(ByVal application As SolidEdgeFramework.Application)
            Dim window As SolidEdgeFramework.Window = Nothing, sheetWindow As SolidEdgeDraft.SheetWindow = Nothing

            If CSharpImpl.__Assign(window, TryCast(application.ActiveWindow, SolidEdgeFramework.Window)) IsNot Nothing Then
                window.View?.Update()
            ElseIf CSharpImpl.__Assign(sheetWindow, TryCast(application.ActiveWindow, SolidEdgeDraft.SheetWindow)) IsNot Nothing Then
                sheetWindow.Update()
            End If
        End Sub

        Private Class CSharpImpl
            <Obsolete("Please refactor calling code to use normal Visual Basic assignment")>
            Shared Function __Assign(Of T)(ByRef target As T, value As T) As T
                target = value
                Return value
            End Function
        End Class
    End Module

    Module SheetExtensions
        <DllImport("user32.dll")>
        Private Function CloseClipboard() As Boolean
        <DllImport("user32.dll", CharSet:=CharSet.Unicode, ExactSpelling:=True, CallingConvention:=CallingConvention.StdCall)>
        Private Function GetClipboardData(ByVal format As UInteger) As IntPtr
        <DllImport("user32.dll")>
        Private Function GetClipboardOwner() As IntPtr
        <DllImport("user32.dll")>
        Private Function IsClipboardFormatAvailable(ByVal format As UInteger) As Boolean
        <DllImport("user32.dll")>
        Private Function OpenClipboard(ByVal hWndNewOwner As IntPtr) As Boolean
        <DllImport("gdi32.dll")>
        Private Function DeleteEnhMetaFile(ByVal hemf As IntPtr) As Boolean
        <DllImport("gdi32.dll")>
        Private Function GetEnhMetaFileBits(ByVal hemf As IntPtr, ByVal cbBuffer As UInteger,
        <Out> ByVal lpbBuffer As Byte()) As UInteger
        Const CF_ENHMETAFILE As UInteger = 14

        <Extension()>
        Function GetEnhancedMetafile(ByVal sheet As SolidEdgeDraft.Sheet) As System.Drawing.Imaging.Metafile
            Try
                sheet.CopyEMFToClipboard()

                If OpenClipboard(IntPtr.Zero) Then

                    If IsClipboardFormatAvailable(CF_ENHMETAFILE) Then
                        Dim henhmetafile As IntPtr = GetClipboardData(CF_ENHMETAFILE)
                        Return New System.Drawing.Imaging.Metafile(henhmetafile, True)
                    Else
                        Throw New System.Exception("CF_ENHMETAFILE is not available in clipboard.")
                    End If
                Else
                    Throw New System.Exception("Error opening clipboard.")
                End If

            Finally
                CloseClipboard()
            End Try
        End Function

        <Extension()>
        Sub SaveAsEnhancedMetafile(ByVal sheet As SolidEdgeDraft.Sheet, ByVal filename As String)
            Try
                sheet.CopyEMFToClipboard()

                If OpenClipboard(IntPtr.Zero) Then

                    If IsClipboardFormatAvailable(CF_ENHMETAFILE) Then
                        Dim hEMF As IntPtr = GetClipboardData(CF_ENHMETAFILE)
                        Dim len As UInteger = GetEnhMetaFileBits(hEMF, 0, Nothing)
                        Dim rawBytes As Byte() = New Byte(len - 1) {}
                        GetEnhMetaFileBits(hEMF, len, rawBytes)
                        System.IO.File.WriteAllBytes(filename, rawBytes)
                        DeleteEnhMetaFile(hEMF)
                    Else
                        Throw New System.Exception("CF_ENHMETAFILE is not available in clipboard.")
                    End If
                Else
                    Throw New System.Exception("Error opening clipboard.")
                End If

            Finally
                CloseClipboard()
            End Try
        End Sub

        Private Class CSharpImpl
            <Obsolete("Please refactor calling code to use normal Visual Basic assignment")>
            Shared Function __Assign(Of T)(ByRef target As T, value As T) As T
                target = value
                Return value
            End Function
        End Class
    End Module

    Module UnitsOfMeasureExtensions
        <Extension()>
        Function FormatUnit(Of T As Class)(ByVal unitsOfMeasure As SolidEdgeFramework.UnitsOfMeasure, ByVal unitType As SolidEdgeFramework.UnitTypeConstants, ByVal value As Double) As T
            Return CType(unitsOfMeasure.FormatUnit(CInt(unitType), value, Nothing), T)
        End Function

        <Extension()>
        Function FormatUnit(Of T As Class)(ByVal unitsOfMeasure As SolidEdgeFramework.UnitsOfMeasure, ByVal unitType As SolidEdgeFramework.UnitTypeConstants, ByVal value As Double, ByVal precision As SolidEdgeConstants.PrecisionConstants) As T
            Return CType(unitsOfMeasure.FormatUnit(CInt(unitType), value, precision), T)
        End Function

        Private Class CSharpImpl
            <Obsolete("Please refactor calling code to use normal Visual Basic assignment")>
            Shared Function __Assign(Of T)(ByRef target As T, value As T) As T
                target = value
                Return value
            End Function
        End Class
    End Module

    Module WindowExtensions
        <Extension()>
        Function GetDrawHandle(ByVal window As SolidEdgeFramework.Window) As IntPtr
            Return New IntPtr(window.DrawHwnd)
        End Function

        <Extension()>
        Function GetHandle(ByVal window As SolidEdgeFramework.Window) As IntPtr
            Return New IntPtr(window.hWnd)
        End Function

        Private Class CSharpImpl
            <Obsolete("Please refactor calling code to use normal Visual Basic assignment")>
            Shared Function __Assign(Of T)(ByRef target As T, value As T) As T
                target = value
                Return value
            End Function
        End Class
    End Module
End Namespace

Namespace SolidEdgeSDK.InteropServices
    Module CATID
        Public Const SolidEdgeAddIn As String = "26B1D2D1-2B03-11d2-B589-080036E8B802"
        Public Const SolidEdgeOEMAddIn As String = "E71958D9-E29B-45F1-A502-A52231F24239"
        Public Const SEApplication As String = "26618394-09D6-11d1-BA07-080036230602"
        Public Const SEAssembly As String = "26618395-09D6-11d1-BA07-080036230602"
        Public Const SEMotion As String = "67ED3F40-A351-11d3-A40B-0004AC969602"
        Public Const SEPart As String = "26618396-09D6-11d1-BA07-080036230602"
        Public Const SEProfile As String = "26618397-09D6-11d1-BA07-080036230602"
        Public Const SEFeatureRecognition As String = "E6F9C8DC-B256-11d3-A41E-0004AC969602"
        Public Const SESheetMetal As String = "26618398-09D6-11D1-BA07-080036230602"
        Public Const SEDraft As String = "08244193-B78D-11D2-9216-00C04F79BE98"
        Public Const SEWeldment As String = "7313526A-276F-11D4-B64E-00C04F79B2BF"
        Public Const SEXpresRoute As String = "1661432A-489C-4714-B1B2-61E85CFD0B71"
        Public Const SEExplode As String = "23BE4212-5810-478b-94FF-B4D682C1B538"
        Public Const SESimplify As String = "CE3DCEBF-E36E-4851-930A-ED892FE0772A"
        Public Const SEStudio As String = "D35550BF-0810-4f67-97D5-789EDBC23F4D"
        Public Const SELayout As String = "27B34941-FFCD-4768-9102-0B6698656CEA"
        Public Const SESketch As String = "0DDABC90-125E-4cfe-9CB7-DC97FB74CCF4"
        Public Const SEProfileHole As String = "0D5CC5F7-5BA3-4d2f-B6A9-31D9B401FE30"
        Public Const SEProfilePattern As String = "7BD57D4B-BA47-4a79-A4E2-DFFD43B97ADF"
        Public Const SEProfileRevolved As String = "FB73C683-1536-4073-B792-E28B8D31146E"
        Public Const SEDrawingViewEdit As String = "8DBC3B5F-02D6-4241-BE96-B12EAF83FAE6"
        Public Const SERefAxis As String = "B21CCFF8-1FDD-4f44-9417-F1EAE06888FA"
        Public Const SECuttingPlaneLine As String = "7C6F65F1-A02D-4c3c-8063-8F54B59B34E3"
        Public Const SEBrokenOutSectionProfile As String = "534CAB66-8089-4e18-8FC4-6FA5A957E445"
        Public Const SEFrame As String = "D84119E8-F844-4823-B3A0-D4F31793028A"
        Public Const SE2dModel As String = "F6031120-7D99-48a7-95FC-EEE8038D7996"
        Public Const SEEditBlockView As String = "892A1CDA-12AE-4619-BB69-C5156C929832"
        Public Const EditBlockInPlace As String = "308A1927-CDCE-4b92-B654-241362608CDE"
        Public Const SEComponentSketchInPart As String = "FAB8DC23-00F4-4872-8662-18DD013F2095"
        Public Const SEComponentSketchInAsm As String = "86D925FB-66ED-40d2-AA3D-D04E74838141"
        Public Const SEHarness As String = "5337A0AB-23ED-4261-A238-00E2070406FC"
        Public Const SEAll As String = "C484ED57-DBB6-4a83-BEDB-C08600AF07BF"
        Public Const SEAllDocumentEnvrionments As String = "BAD41B8D-18FF-42c9-9611-8A00E6921AE8"
        Public Const SEEditMV As String = "C1D8CCB8-54D3-4fce-92AB-0668147FC7C3"
        Public Const SEEditMVPart As String = "054BDB42-6C1E-41a4-9014-3D51BEE911EF"
        Public Const SEDMPart As String = "D9B0BB85-3A6C-4086-A0BB-88A1AAD57A58"
        Public Const SEDMSheetMetal As String = "9CBF2809-FF80-4dbc-98F2-B82DABF3530F"
        Public Const SEDMAssembly As String = "2C3C2A72-3A4A-471d-98B5-E3A8CFA4A2BF"
        Public Const FEAResultsPart As String = "B5965D1C-8819-4902-8252-64841537A16C"
        Public Const FEAResultsAsm As String = "986B2512-3AE9-4a57-8513-1D2A1E3520DD"
        Public Const SESimplifiedAssemblyPart As String = "E7350DC3-6E7A-4D53-A53F-5B1C7A0709B3"
        Public Const Sketch3d As String = "07F05BA4-18CD-4B87-8E2F-49864E71B41F"
        Public Const SEAssemblyViewer As String = "F2483121-58BC-44AF-8B8F-D7B74DC8408B"

        Private Class CSharpImpl
            <Obsolete("Please refactor calling code to use normal Visual Basic assignment")>
            Shared Function __Assign(Of T)(ByRef target As T, value As T) As T
                target = value
                Return value
            End Function
        End Class
    End Module

    Public Class ComEventsManager
        Public Sub New(ByVal sink As Object)
            If sink Is Nothing Then Throw New ArgumentNullException(NameOf(sink))
            Sink = sink
        End Sub

        Public Sub Attach(Of TInterface As Class)(ByVal container As Object)
            Dim lockTaken As Boolean = False

            Try
                System.Threading.Monitor.Enter(Me, lockTaken)

                If IsAttached(Of TInterface)(container) Then
                    Return
                End If

                Dim cpc As System.Runtime.InteropServices.ComTypes.IConnectionPointContainer = Nothing
                Dim cp As System.Runtime.InteropServices.ComTypes.IConnectionPoint = Nothing
                Dim cookie As Integer = 0
                cpc = CType(container, System.Runtime.InteropServices.ComTypes.IConnectionPointContainer)
                cpc.FindConnectionPoint(GetType(TInterface).GUID, cp)

                If cp IsNot Nothing Then
                    cp.Advise(Sink, cookie)
                    ConnectionPoints.Add(cp, cookie)
                End If

            Finally

                If lockTaken Then
                    System.Threading.Monitor.[Exit](Me)
                End If
            End Try
        End Sub

        Public Function IsAttached(Of TInterface As Class)(ByVal container As Object) As Boolean
            Dim lockTaken As Boolean = False

            Try
                System.Threading.Monitor.Enter(Me, lockTaken)
                Dim cpc As System.Runtime.InteropServices.ComTypes.IConnectionPointContainer = Nothing
                Dim cp As System.Runtime.InteropServices.ComTypes.IConnectionPoint = Nothing
                Dim cookie As Integer = 0
                cpc = CType(container, System.Runtime.InteropServices.ComTypes.IConnectionPointContainer)
                cpc.FindConnectionPoint(GetType(TInterface).GUID, cp)

                If cp IsNot Nothing Then

                    If ConnectionPoints.ContainsKey(cp) Then
                        cookie = ConnectionPoints(cp)
                        Return True
                    End If
                End If

            Finally

                If lockTaken Then
                    System.Threading.Monitor.[Exit](Me)
                End If
            End Try

            Return False
        End Function

        Public Sub Detach(Of TInterface As Class)(ByVal container As Object)
            Dim lockTaken As Boolean = False

            Try
                System.Threading.Monitor.Enter(Me, lockTaken)
                Dim cpc As System.Runtime.InteropServices.ComTypes.IConnectionPointContainer = Nothing
                Dim cp As System.Runtime.InteropServices.ComTypes.IConnectionPoint = Nothing
                Dim cookie As Integer = 0
                cpc = CType(container, System.Runtime.InteropServices.ComTypes.IConnectionPointContainer)
                cpc.FindConnectionPoint(GetType(TInterface).GUID, cp)

                If cp IsNot Nothing Then

                    If ConnectionPoints.ContainsKey(cp) Then
                        cookie = ConnectionPoints(cp)
                        cp.Unadvise(cookie)
                        ConnectionPoints.Remove(cp)
                    End If
                End If

            Finally

                If lockTaken Then
                    System.Threading.Monitor.[Exit](Me)
                End If
            End Try
        End Sub

        Public Sub DetachAll()
            Dim lockTaken As Boolean = False

            Try
                System.Threading.Monitor.Enter(Me, lockTaken)
                Dim enumerator As Dictionary(Of System.Runtime.InteropServices.ComTypes.IConnectionPoint, Integer).Enumerator = ConnectionPoints.GetEnumerator()

                While enumerator.MoveNext()
                    enumerator.Current.Key.Unadvise(enumerator.Current.Value)
                End While

            Finally
                ConnectionPoints.Clear()

                If lockTaken Then
                    System.Threading.Monitor.[Exit](Me)
                End If
            End Try
        End Sub

        Public Property Sink As Object
        Public Property ConnectionPoints As Dictionary(Of System.Runtime.InteropServices.ComTypes.IConnectionPoint, Integer) = New Dictionary(Of System.Runtime.InteropServices.ComTypes.IConnectionPoint, Integer)()

        Private Class CSharpImpl
            <Obsolete("Please refactor calling code to use normal Visual Basic assignment")>
            Shared Function __Assign(Of T)(ByRef target As T, value As T) As T
                target = value
                Return value
            End Function
        End Class
    End Class

    Module ComObject
        Const LOCALE_SYSTEM_DEFAULT As Integer = 2048

        Function GetITypeInfo(ByVal comObject As Object) As System.Runtime.InteropServices.ComTypes.ITypeInfo
            If Marshal.IsComObject(comObject) = False Then Throw New InvalidComObjectException()
            Dim dispatch = TryCast(comObject, IDispatch)

            If dispatch IsNot Nothing Then
                Return dispatch.GetTypeInfo(0, LOCALE_SYSTEM_DEFAULT)
            End If

            Return Nothing
        End Function

        Function GetPropertyValue(Of T)(ByVal comObject As Object, ByVal name As String) As T
            Dim type = comObject.[GetType]()
            Dim value = type.InvokeMember(name, System.Reflection.BindingFlags.GetProperty, Nothing, comObject, Nothing)
            Return CType(value, T)
        End Function

        Function GetPropertyValue(Of T)(ByVal comObject As Object, ByVal name As String, ByVal defaultValue As T) As T
            Dim type = comObject.[GetType]()

            Try
                Dim value = type.InvokeMember(name, System.Reflection.BindingFlags.GetProperty, Nothing, comObject, Nothing)
                Return CType(value, T)
            Catch
                Return defaultValue
            End Try
        End Function

        Function GetComTypeFullName(ByVal comObject As Object) As String
            If Marshal.IsComObject(comObject) = False Then Throw New InvalidComObjectException()
            Dim typeLib As ITypeLib = Nothing, pIndex As Integer = Nothing, typeName As String = Nothing, typeDescription As String = Nothing, typeHelpContext As Integer = Nothing, typeHelpFile As String = Nothing, typeLibName As String = Nothing, typeLibDescription As String = Nothing, typeLibHelpContext As Integer = Nothing, typeLibHelpFile As String = Nothing, dispatch As IDispatch = Nothing

            If CSharpImpl.__Assign(dispatch, TryCast(comObject, IDispatch)) IsNot Nothing Then
                Dim typeInfo = dispatch.GetTypeInfo(0, LOCALE_SYSTEM_DEFAULT)
                typeInfo.GetContainingTypeLib(typeLib, pIndex)
                typeInfo.GetDocumentation(-1, typeName, typeDescription, typeHelpContext, typeHelpFile)
                typeLib.GetDocumentation(-1, typeLibName, typeLibDescription, typeLibHelpContext, typeLibHelpFile)
                Return String.Join(".", typeLibName, typeName)
            End If

            Return "IUnknown"
        End Function

        Private Class CSharpImpl
            <Obsolete("Please refactor calling code to use normal Visual Basic assignment")>
            Shared Function __Assign(Of T)(ByRef target As T, value As T) As T
                target = value
                Return value
            End Function
        End Class
    End Module

    <ComImport>
    <Guid("00020400-0000-0000-C000-000000000046")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Interface IDispatch
        Function GetTypeInfoCount() As Integer
        Function GetTypeInfo(ByVal iTInfo As Integer, ByVal lcid As Integer) As System.Runtime.InteropServices.ComTypes.ITypeInfo
    End Interface

    <ComImport>
    <Guid("0002D280-0000-0000-C000-000000000046")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Interface IGL
        Sub glAccum(ByVal op As UInteger, ByVal value As Single)
        Sub glAlphaFunc(ByVal func As UInteger, ByVal aRef As Single)
        Sub glBegin(ByVal mode As UInteger)
        Sub glBitmap(ByVal width As Integer, ByVal height As Integer, ByVal xorig As Single, ByVal yorig As Single, ByVal xmove As Single, ByVal ymove As Single,
<MarshalAs(UnmanagedType.LPArray, ArraySubType:=UnmanagedType.U1)> ByVal bitmap As Byte())
        Sub glBlendFunc(ByVal sfactor As UInteger, ByVal dfactor As UInteger)
        Sub glCallList(ByVal list As UInteger)
        Sub glCallLists(ByVal n As Integer, ByVal type As UInteger,
<MarshalAs(UnmanagedType.LPArray, ArraySubType:=UnmanagedType.U1)> ByVal lists As Byte())
        Sub glClear(ByVal mask As UInteger)
        Sub glClearAccum(ByVal red As Single, ByVal green As Single, ByVal blue As Single, ByVal alpha As Single)
        Sub glClearColor(ByVal red As Single, ByVal green As Single, ByVal blue As Single, ByVal alpha As Single)
        Sub glClearDepth(ByVal depth As Double)
        Sub glClearIndex(ByVal c As Single)
        Sub glClearStencil(ByVal s As Integer)
        Sub glClipPlane(ByVal plane As UInteger,
<MarshalAs(UnmanagedType.LPArray, ArraySubType:=UnmanagedType.R8)> ByVal equation As Double())
        Sub glColor3b(ByVal red As SByte, ByVal green As SByte, ByVal blue As SByte)
        Sub glColor3bv(
<MarshalAs(UnmanagedType.LPArray, ArraySubType:=UnmanagedType.I1)> ByVal v As SByte())
        Sub glColor3d(ByVal red As Double, ByVal green As Double, ByVal blue As Double)
        Sub glColor3dv(
<MarshalAs(UnmanagedType.LPArray, ArraySubType:=UnmanagedType.R8)> ByVal v As Double())
        Sub glColor3f(ByVal red As Single, ByVal green As Single, ByVal blue As Single)
        Sub glColor3fv(
<MarshalAs(UnmanagedType.LPArray, ArraySubType:=UnmanagedType.R4)> ByVal v As Single())
        Sub glColor3i(ByVal red As Integer, ByVal green As Integer, ByVal blue As Integer)
        Sub glColor3iv(
<MarshalAs(UnmanagedType.LPArray, ArraySubType:=UnmanagedType.I4)> ByVal v As Integer())
        Sub glColor3s(ByVal red As Short, ByVal green As Short, ByVal blue As Short)
        Sub glColor3sv(
<MarshalAs(UnmanagedType.LPArray, ArraySubType:=UnmanagedType.I2)> ByVal v As Short())
        Sub glColor3ub(ByVal red As Byte, ByVal green As Byte, ByVal blue As Byte)
        Sub glColor3ubv(
<MarshalAs(UnmanagedType.LPArray, ArraySubType:=UnmanagedType.U1)> ByVal v As Byte())
        Sub glColor3ui(ByVal red As UInteger, ByVal green As UInteger, ByVal blue As UInteger)
        Sub glColor3uiv(
<MarshalAs(UnmanagedType.LPArray, ArraySubType:=UnmanagedType.U4)> ByVal v As UInteger())
        Sub glColor3us(ByVal red As UShort, ByVal green As UShort, ByVal blue As UShort)
        Sub glColor3usv(
<MarshalAs(UnmanagedType.LPArray, ArraySubType:=UnmanagedType.U2)> ByVal v As UShort())
        Sub glColor4b(ByVal red As SByte, ByVal green As SByte, ByVal blue As SByte, ByVal alpha As SByte)
        Sub glColor4bv(
<MarshalAs(UnmanagedType.LPArray, ArraySubType:=UnmanagedType.I1)> ByVal v As SByte())
        Sub glColor4d(ByVal red As Double, ByVal green As Double, ByVal blue As Double, ByVal alpha As Double)
        Sub glColor4dv(
<MarshalAs(UnmanagedType.LPArray, ArraySubType:=UnmanagedType.R8)> ByVal v As Double())
        Sub glColor4f(ByVal red As Single, ByVal green As Single, ByVal blue As Single, ByVal alpha As Single)
        Sub glColor4fv(
<MarshalAs(UnmanagedType.LPArray, ArraySubType:=UnmanagedType.R4)> ByVal v As Single())
        Sub glColor4i(ByVal red As Integer, ByVal green As Integer, ByVal blue As Integer, ByVal alpha As Integer)
        Sub glColor4iv(
<MarshalAs(UnmanagedType.LPArray, ArraySubType:=UnmanagedType.I4)> ByVal v As Integer())
        Sub glColor4s(ByVal red As Short, ByVal green As Short, ByVal blue As Short, ByVal alpha As Short)
        Sub glColor4sv(
<MarshalAs(UnmanagedType.LPArray, ArraySubType:=UnmanagedType.I2)> ByVal v As Short())
        Sub glColor4ub(ByVal red As Byte, ByVal green As Byte, ByVal blue As Byte, ByVal alpha As Byte)
        Sub glColor4ubv(
<MarshalAs(UnmanagedType.LPArray, ArraySubType:=UnmanagedType.U1)> ByVal v As Byte())
        Sub glColor4ui(ByVal red As UInteger, ByVal green As UInteger, ByVal blue As UInteger, ByVal alpha As UInteger)
        Sub glColor4uiv(
<MarshalAs(UnmanagedType.LPArray, ArraySubType:=UnmanagedType.U4)> ByVal v As UInteger())
        Sub glColor4us(ByVal red As UShort, ByVal green As UShort, ByVal blue As UShort, ByVal alpha As UShort)
        Sub glColor4usv(
<MarshalAs(UnmanagedType.LPArray, ArraySubType:=UnmanagedType.U2)> ByVal v As UShort())
        Sub glColorMask(ByVal red As Byte, ByVal green As Byte, ByVal blue As Byte, ByVal alpha As Byte)
        Sub glColorMaterial(ByVal face As UInteger, ByVal mode As UInteger)
        Sub glCopyPixels(ByVal x As Integer, ByVal y As Integer, ByVal width As Integer, ByVal height As Integer, ByVal type As UInteger)
        Sub glCullFace(ByVal mode As UInteger)
        Sub glDeleteLists(ByVal list As UInteger, ByVal range As Integer)
        Sub glDepthFunc(ByVal func As UInteger)
        Sub glDepthMask(ByVal flag As Byte)
        Sub glDepthRange(ByVal zNear As Double, ByVal zFar As Double)
        Sub glDisable(ByVal cap As UInteger)
        Sub glDrawBuffer(ByVal mode As UInteger)
        Sub glDrawPixels(ByVal width As Integer, ByVal height As Integer, ByVal format As UInteger, ByVal type As UInteger,
<MarshalAs(UnmanagedType.LPArray, ArraySubType:=UnmanagedType.U1)> ByVal pixels As Byte())
        Sub glEdgeFlag(ByVal flag As Byte)
        Sub glEdgeFlagv(
<MarshalAs(UnmanagedType.LPArray, ArraySubType:=UnmanagedType.U1)> ByVal flag As Byte())
        Sub glEnable(ByVal cap As UInteger)
        Sub glEnd()
        Sub glEndList()
        Sub glEvalCoord1d(ByVal u As Double)
        Sub glEvalCoord1dv(
<MarshalAs(UnmanagedType.LPArray, ArraySubType:=UnmanagedType.R8)> ByVal u As Double())
        Sub glEvalCoord1f(ByVal u As Single)
        Sub glEvalCoord1fv(
<MarshalAs(UnmanagedType.LPArray, ArraySubType:=UnmanagedType.R4)> ByVal u As Single())
        Sub glEvalCoord2d(ByVal u As Double, ByVal v As Double)
        Sub glEvalCoord2dv(
<MarshalAs(UnmanagedType.LPArray, ArraySubType:=UnmanagedType.R8)> ByVal u As Double())
        Sub glEvalCoord2f(ByVal u As Single, ByVal v As Single)
        Sub glEvalCoord2fv(
<MarshalAs(UnmanagedType.LPArray, ArraySubType:=UnmanagedType.R4)> ByVal u As Single())
        Sub glEvalMesh1(ByVal mode As UInteger, ByVal i1 As Integer, ByVal i2 As Integer)
        Sub glEvalMesh2(ByVal mode As UInteger, ByVal i1 As Integer, ByVal i2 As Integer, ByVal j1 As Integer, ByVal j2 As Integer)
        Sub glEvalPoint1(ByVal i As Integer)
        Sub glEvalPoint2(ByVal i As Integer, ByVal j As Integer)
        Sub glFeedbackBuffer(ByVal size As Integer, ByVal type As UInteger,
<MarshalAs(UnmanagedType.LPArray, ArraySubType:=UnmanagedType.R4)> ByVal buffer As Single())
        Sub glFinish()
        Sub glFlush()
        Sub glFogf(ByVal pname As UInteger, ByVal param As Single)
        Sub glFogfv(ByVal pname As UInteger,
<MarshalAs(UnmanagedType.LPArray, ArraySubType:=UnmanagedType.R4)> ByVal someParams As Single())
        Sub glFogi(ByVal pname As UInteger, ByVal param As Integer)
        Sub glFogiv(ByVal pname As UInteger,
<MarshalAs(UnmanagedType.LPArray, ArraySubType:=UnmanagedType.I4)> ByVal someParams As Integer())
        Sub glFrontFace(ByVal mode As UInteger)
        Sub glFrustum(ByVal left As Double, ByVal right As Double, ByVal bottom As Double, ByVal top As Double, ByVal zNear As Double, ByVal zFar As Double)
        Function glGenLists(ByVal range As Integer) As UInteger
        Sub glGetBooleanv(ByVal pname As UInteger,
<MarshalAs(UnmanagedType.LPArray, ArraySubType:=UnmanagedType.U1)> ByVal someParams As Byte())
        Sub glGetClipPlane(ByVal plane As UInteger,
<MarshalAs(UnmanagedType.LPArray, ArraySubType:=UnmanagedType.R8)> ByVal equation As Double())
        Sub glGetDoublev(ByVal pname As UInteger,
<MarshalAs(UnmanagedType.LPArray, ArraySubType:=UnmanagedType.R8)> ByVal someParams As Double())
        Function glGetError() As UInteger
        Sub glGetFloatv(ByVal pname As UInteger,
<MarshalAs(UnmanagedType.LPArray, ArraySubType:=UnmanagedType.R4)> ByVal someParams As Single())
        Sub glGetIntegerv(ByVal pname As UInteger, ByRef someParams As Integer)
        Sub glGetLightfv(ByVal light As UInteger, ByVal pname As UInteger,
<MarshalAs(UnmanagedType.LPArray, ArraySubType:=UnmanagedType.R4)> ByVal someParams As Single())
        Sub glGetLightiv(ByVal light As UInteger, ByVal pname As UInteger,
<MarshalAs(UnmanagedType.LPArray, ArraySubType:=UnmanagedType.I4)> ByVal someParams As Integer())
        Sub glGetMapdv(ByVal target As UInteger, ByVal query As UInteger,
<MarshalAs(UnmanagedType.LPArray, ArraySubType:=UnmanagedType.R8)> ByVal v As Double())
        Sub glGetMapfv(ByVal target As UInteger, ByVal query As UInteger,
<MarshalAs(UnmanagedType.LPArray, ArraySubType:=UnmanagedType.R4)> ByVal v As Single())
        Sub glGetMapiv(ByVal target As UInteger, ByVal query As UInteger,
<MarshalAs(UnmanagedType.LPArray, ArraySubType:=UnmanagedType.I4)> ByVal v As Integer())
        Sub glGetMaterialfv(ByVal face As UInteger, ByVal pname As UInteger,
<MarshalAs(UnmanagedType.LPArray, ArraySubType:=UnmanagedType.R4)> ByVal someParams As Single())
        Sub glGetMaterialiv(ByVal face As UInteger, ByVal pname As UInteger,
<MarshalAs(UnmanagedType.LPArray, ArraySubType:=UnmanagedType.I4)> ByVal someParams As Integer())
        Sub glGetPixelMapfv(ByVal map As UInteger,
<MarshalAs(UnmanagedType.LPArray, ArraySubType:=UnmanagedType.R4)> ByVal values As Single())
        Sub glGetPixelMapuiv(ByVal map As UInteger,
<MarshalAs(UnmanagedType.LPArray, ArraySubType:=UnmanagedType.U4)> ByVal values As UInteger())
        Sub glGetPixelMapusv(ByVal map As UInteger,
<MarshalAs(UnmanagedType.LPArray, ArraySubType:=UnmanagedType.U2)> ByVal values As UShort())
        Sub glGetPolygonStipple(
<MarshalAs(UnmanagedType.LPArray, ArraySubType:=UnmanagedType.U1)> ByVal mask As Byte())
        Function glGetString(ByVal name As UInteger) As SByte()
        Sub glGetTexEnvfv(ByVal target As UInteger, ByVal pname As UInteger,
<MarshalAs(UnmanagedType.LPArray, ArraySubType:=UnmanagedType.R4)> ByVal someParams As Single())
        Sub glGetTexEnviv(ByVal target As UInteger, ByVal pname As UInteger,
<MarshalAs(UnmanagedType.LPArray, ArraySubType:=UnmanagedType.I4)> ByVal someParams As Integer())
        Sub glGetTexGendv(ByVal coord As UInteger, ByVal pname As UInteger,
<MarshalAs(UnmanagedType.LPArray, ArraySubType:=UnmanagedType.R8)> ByVal someParams As Double())
        Sub glGetTexGenfv(ByVal coord As UInteger, ByVal pname As UInteger,
<MarshalAs(UnmanagedType.LPArray, ArraySubType:=UnmanagedType.R4)> ByVal someParams As Single())
        Sub glGetTexGeniv(ByVal coord As UInteger, ByVal pname As UInteger,
<MarshalAs(UnmanagedType.LPArray, ArraySubType:=UnmanagedType.I4)> ByVal someParams As Integer())
        Sub glGetTexImage(ByVal target As UInteger, ByVal level As Integer, ByVal format As UInteger, ByVal type As UInteger,
<MarshalAs(UnmanagedType.LPArray, ArraySubType:=UnmanagedType.U1)> ByVal pixels As Byte())
        Sub glGetTexLevelParameterfv(ByVal target As UInteger, ByVal level As Integer, ByVal pname As UInteger,
<MarshalAs(UnmanagedType.LPArray, ArraySubType:=UnmanagedType.R4)> ByVal someParams As Single())
        Sub glGetTexLevelParameteriv(ByVal target As UInteger, ByVal level As Integer, ByVal pname As UInteger,
<MarshalAs(UnmanagedType.LPArray, ArraySubType:=UnmanagedType.I4)> ByVal someParams As Integer())
        Sub glGetTexParameterfv(ByVal target As UInteger, ByVal pname As UInteger,
<MarshalAs(UnmanagedType.LPArray, ArraySubType:=UnmanagedType.R4)> ByVal someParams As Single())
        Sub glGetTexParameteriv(ByVal target As UInteger, ByVal pname As UInteger,
<MarshalAs(UnmanagedType.LPArray, ArraySubType:=UnmanagedType.I4)> ByVal someParams As Integer())
        Sub glHint(ByVal target As UInteger, ByVal mode As UInteger)
        Sub glIndexMask(ByVal mask As UInteger)
        Sub glIndexd(ByVal c As Double)
        Sub glIndexdv(
<MarshalAs(UnmanagedType.LPArray, ArraySubType:=UnmanagedType.R8)> ByVal c As Double())
        Sub glIndexf(ByVal c As Single)
        Sub glIndexfv(
<MarshalAs(UnmanagedType.LPArray, ArraySubType:=UnmanagedType.R4)> ByVal c As Single())
        Sub glIndexi(ByVal c As Integer)
        Sub glIndexiv(
<MarshalAs(UnmanagedType.LPArray, ArraySubType:=UnmanagedType.I4)> ByVal c As Integer())
        Sub glIndexs(ByVal c As Short)
        Sub glIndexsv(
<MarshalAs(UnmanagedType.LPArray, ArraySubType:=UnmanagedType.I2)> ByVal c As Short())
        Sub glInitNames()
        Function glIsEnabled(ByVal cap As UInteger) As Byte
        Function glIsList(ByVal list As UInteger) As Byte
        Sub glLightModelf(ByVal pname As UInteger, ByVal param As Single)
        Sub glLightModelfv(ByVal pname As UInteger,
<MarshalAs(UnmanagedType.LPArray, ArraySubType:=UnmanagedType.R4)> ByVal someParams As Single())
        Sub glLightModeli(ByVal pname As UInteger, ByVal param As Integer)
        Sub glLightModeliv(ByVal pname As UInteger,
<MarshalAs(UnmanagedType.LPArray, ArraySubType:=UnmanagedType.I4)> ByVal someParams As Integer())
        Sub glLightf(ByVal light As UInteger, ByVal pname As UInteger, ByVal param As Single)
        Sub glLightfv(ByVal light As UInteger, ByVal pname As UInteger,
<MarshalAs(UnmanagedType.LPArray, ArraySubType:=UnmanagedType.R4)> ByVal someParams As Single())
        Sub glLighti(ByVal light As UInteger, ByVal pname As UInteger, ByVal param As Integer)
        Sub glLightiv(ByVal light As UInteger, ByVal pname As UInteger,
<MarshalAs(UnmanagedType.LPArray, ArraySubType:=UnmanagedType.I4)> ByVal someParams As Integer())
        Sub glLineStipple(ByVal factor As Integer, ByVal pattern As UShort)
        Sub glLineWidth(ByVal width As Single)
        Sub glListBase(ByVal aBase As UInteger)
        Sub glLoadIdentity()
        Sub glLoadMatrixd(
<MarshalAs(UnmanagedType.LPArray, ArraySubType:=UnmanagedType.R8)> ByVal m As Double())
        Sub glLoadMatrixf(
<MarshalAs(UnmanagedType.LPArray, ArraySubType:=UnmanagedType.R4)> ByVal m As Single())
        Sub glLoadName(ByVal name As UInteger)
        Sub glLogicOp(ByVal opcode As UInteger)
        Sub glMap1d(ByVal target As UInteger, ByVal u1 As Double, ByVal u2 As Double, ByVal stride As Integer, ByVal order As Integer,
<MarshalAs(UnmanagedType.LPArray, ArraySubType:=UnmanagedType.R8)> ByVal points As Double())
        Sub glMap1f(ByVal target As UInteger, ByVal u1 As Single, ByVal u2 As Single, ByVal stride As Integer, ByVal order As Integer,
<MarshalAs(UnmanagedType.LPArray, ArraySubType:=UnmanagedType.R4)> ByVal points As Single())
        Sub glMap2d(ByVal target As UInteger, ByVal u1 As Double, ByVal u2 As Double, ByVal ustride As Integer, ByVal uorder As Integer, ByVal v1 As Double, ByVal v2 As Double, ByVal vstride As Integer, ByVal vorder As Integer,
<MarshalAs(UnmanagedType.LPArray, ArraySubType:=UnmanagedType.R8)> ByVal points As Double())
        Sub glMap2f(ByVal target As UInteger, ByVal u1 As Single, ByVal u2 As Single, ByVal ustride As Integer, ByVal uorder As Integer, ByVal v1 As Single, ByVal v2 As Single, ByVal vstride As Integer, ByVal vorder As Integer,
<MarshalAs(UnmanagedType.LPArray, ArraySubType:=UnmanagedType.R4)> ByVal points As Single())
        Sub glMapGrid1d(ByVal un As Integer, ByVal u1 As Double, ByVal u2 As Double)
        Sub glMapGrid1f(ByVal un As Integer, ByVal u1 As Single, ByVal u2 As Single)
        Sub glMapGrid2d(ByVal un As Integer, ByVal u1 As Double, ByVal u2 As Double, ByVal vn As Integer, ByVal v1 As Double, ByVal v2 As Double)
        Sub glMapGrid2f(ByVal un As Integer, ByVal u1 As Single, ByVal u2 As Single, ByVal vn As Integer, ByVal v1 As Single, ByVal v2 As Single)
        Sub glMaterialf(ByVal face As UInteger, ByVal pname As UInteger, ByVal param As Single)
        Sub glMaterialfv(ByVal face As UInteger, ByVal pname As UInteger,
<MarshalAs(UnmanagedType.LPArray, ArraySubType:=UnmanagedType.R4)> ByVal someParams As Single())
        Sub glMateriali(ByVal face As UInteger, ByVal pname As UInteger, ByVal param As Integer)
        Sub glMaterialiv(ByVal face As UInteger, ByVal pname As UInteger,
<MarshalAs(UnmanagedType.LPArray, ArraySubType:=UnmanagedType.I4)> ByVal someParams As Integer())
        Sub glMatrixMode(ByVal mode As UInteger)
        Sub glMultMatrixd(
<MarshalAs(UnmanagedType.LPArray, ArraySubType:=UnmanagedType.R8)> ByVal m As Double())
        Sub glMultMatrixf(
<MarshalAs(UnmanagedType.LPArray, ArraySubType:=UnmanagedType.R4)> ByVal m As Single())
        Sub glNewList(ByVal list As UInteger, ByVal mode As UInteger)
        Sub glNormal3b(ByVal nx As SByte, ByVal ny As SByte, ByVal nz As SByte)
        Sub glNormal3bv(
<MarshalAs(UnmanagedType.LPArray, ArraySubType:=UnmanagedType.I1)> ByVal v As SByte())
        Sub glNormal3d(ByVal nx As Double, ByVal ny As Double, ByVal nz As Double)
        Sub glNormal3dv(
<MarshalAs(UnmanagedType.LPArray, ArraySubType:=UnmanagedType.R8)> ByVal v As Double())
        Sub glNormal3f(ByVal nx As Single, ByVal ny As Single, ByVal nz As Single)
        Sub glNormal3fv(
<MarshalAs(UnmanagedType.LPArray, ArraySubType:=UnmanagedType.R4)> ByVal v As Single())
        Sub glNormal3i(ByVal nx As Integer, ByVal ny As Integer, ByVal nz As Integer)
        Sub glNormal3iv(
<MarshalAs(UnmanagedType.LPArray, ArraySubType:=UnmanagedType.I4)> ByVal v As Integer())
        Sub glNormal3s(ByVal nx As Short, ByVal ny As Short, ByVal nz As Short)
        Sub glNormal3sv(
<MarshalAs(UnmanagedType.LPArray, ArraySubType:=UnmanagedType.I2)> ByVal v As Short())
        Sub glOrtho(ByVal left As Double, ByVal right As Double, ByVal bottom As Double, ByVal top As Double, ByVal zNear As Double, ByVal zFar As Double)
        Sub glPassThrough(ByVal token As Single)
        Sub glPixelMapfv(ByVal map As UInteger, ByVal mapsize As Integer,
<MarshalAs(UnmanagedType.LPArray, ArraySubType:=UnmanagedType.R4)> ByVal values As Single())
        Sub glPixelMapuiv(ByVal map As UInteger, ByVal mapsize As Integer,
<MarshalAs(UnmanagedType.LPArray, ArraySubType:=UnmanagedType.U4)> ByVal values As UInteger())
        Sub glPixelMapusv(ByVal map As UInteger, ByVal mapsize As Integer,
<MarshalAs(UnmanagedType.LPArray, ArraySubType:=UnmanagedType.U2)> ByVal values As UShort())
        Sub glPixelStoref(ByVal pname As UInteger, ByVal param As Single)
        Sub glPixelStorei(ByVal pname As UInteger, ByVal param As Integer)
        Sub glPixelTransferf(ByVal pname As UInteger, ByVal param As Single)
        Sub glPixelTransferi(ByVal pname As UInteger, ByVal param As Integer)
        Sub glPixelZoom(ByVal xfactor As Single, ByVal yfactor As Single)
        Sub glPointSize(ByVal size As Single)
        Sub glPolygonMode(ByVal face As UInteger, ByVal mode As UInteger)
        Sub glPolygonStipple(
<MarshalAs(UnmanagedType.LPArray, ArraySubType:=UnmanagedType.U1)> ByVal mask As Byte())
        Sub glPopAttrib()
        Sub glPopMatrix()
        Sub glPopName()
        Sub glPushAttrib(ByVal mask As UInteger)
        Sub glPushMatrix()
        Sub glPushName(ByVal name As UInteger)
        Sub glRasterPos2d(ByVal x As Double, ByVal y As Double)
        Sub glRasterPos2dv(
<MarshalAs(UnmanagedType.LPArray, ArraySubType:=UnmanagedType.R8)> ByVal v As Double())
        Sub glRasterPos2f(ByVal x As Single, ByVal y As Single)
        Sub glRasterPos2fv(
<MarshalAs(UnmanagedType.LPArray, ArraySubType:=UnmanagedType.R4)> ByVal v As Single())
        Sub glRasterPos2i(ByVal x As Integer, ByVal y As Integer)
        Sub glRasterPos2iv(
<MarshalAs(UnmanagedType.LPArray, ArraySubType:=UnmanagedType.I4)> ByVal v As Integer())
        Sub glRasterPos2s(ByVal x As Short, ByVal y As Short)
        Sub glRasterPos2sv(
<MarshalAs(UnmanagedType.LPArray, ArraySubType:=UnmanagedType.I2)> ByVal v As Short())
        Sub glRasterPos3d(ByVal x As Double, ByVal y As Double, ByVal z As Double)
        Sub glRasterPos3dv(
<MarshalAs(UnmanagedType.LPArray, ArraySubType:=UnmanagedType.R8)> ByVal v As Double())
        Sub glRasterPos3f(ByVal x As Single, ByVal y As Single, ByVal z As Single)
        Sub glRasterPos3fv(
<MarshalAs(UnmanagedType.LPArray, ArraySubType:=UnmanagedType.R4)> ByVal v As Single())
        Sub glRasterPos3i(ByVal x As Integer, ByVal y As Integer, ByVal z As Integer)
        Sub glRasterPos3iv(
<MarshalAs(UnmanagedType.LPArray, ArraySubType:=UnmanagedType.I4)> ByVal v As Integer())
        Sub glRasterPos3s(ByVal x As Short, ByVal y As Short, ByVal z As Short)
        Sub glRasterPos3sv(
<MarshalAs(UnmanagedType.LPArray, ArraySubType:=UnmanagedType.I2)> ByVal v As Short())
        Sub glRasterPos4d(ByVal x As Double, ByVal y As Double, ByVal z As Double, ByVal w As Double)
        Sub glRasterPos4dv(
<MarshalAs(UnmanagedType.LPArray, ArraySubType:=UnmanagedType.R8)> ByVal v As Double())
        Sub glRasterPos4f(ByVal x As Single, ByVal y As Single, ByVal z As Single, ByVal w As Single)
        Sub glRasterPos4fv(
<MarshalAs(UnmanagedType.LPArray, ArraySubType:=UnmanagedType.R4)> ByVal v As Single())
        Sub glRasterPos4i(ByVal x As Integer, ByVal y As Integer, ByVal z As Integer, ByVal w As Integer)
        Sub glRasterPos4iv(
<MarshalAs(UnmanagedType.LPArray, ArraySubType:=UnmanagedType.I4)> ByVal v As Integer())
        Sub glRasterPos4s(ByVal x As Short, ByVal y As Short, ByVal z As Short, ByVal w As Short)
        Sub glRasterPos4sv(
<MarshalAs(UnmanagedType.LPArray, ArraySubType:=UnmanagedType.I2)> ByVal v As Short())
        Sub glReadBuffer(ByVal mode As UInteger)
        Sub glReadPixels(ByVal x As Integer, ByVal y As Integer, ByVal width As Integer, ByVal height As Integer, ByVal format As UInteger, ByVal type As UInteger,
<MarshalAs(UnmanagedType.LPArray, ArraySubType:=UnmanagedType.U1)> ByVal pixels As Byte())
        Sub glRectd(ByVal x1 As Double, ByVal y1 As Double, ByVal x2 As Double, ByVal y2 As Double)
        Sub glRectdv(
<MarshalAs(UnmanagedType.LPArray, ArraySubType:=UnmanagedType.R8)> ByVal v1 As Double(),
<MarshalAs(UnmanagedType.LPArray, ArraySubType:=UnmanagedType.R8)> ByVal v2 As Double())
        Sub glRectf(ByVal x1 As Single, ByVal y1 As Single, ByVal x2 As Single, ByVal y2 As Single)
        Sub glRectfv(
<MarshalAs(UnmanagedType.LPArray, ArraySubType:=UnmanagedType.R4)> ByVal v1 As Single(),
<MarshalAs(UnmanagedType.LPArray, ArraySubType:=UnmanagedType.R4)> ByVal v2 As Single())
        Sub glRecti(ByVal x1 As Integer, ByVal y1 As Integer, ByVal x2 As Integer, ByVal y2 As Integer)
        Sub glRectiv(
<MarshalAs(UnmanagedType.LPArray, ArraySubType:=UnmanagedType.I4)> ByVal v1 As Integer(),
<MarshalAs(UnmanagedType.LPArray, ArraySubType:=UnmanagedType.I4)> ByVal v2 As Integer())
        Sub glRects(ByVal x1 As Short, ByVal y1 As Short, ByVal x2 As Short, ByVal y2 As Short)
        Sub glRectsv(
<MarshalAs(UnmanagedType.LPArray, ArraySubType:=UnmanagedType.I2)> ByVal v1 As Short(),
<MarshalAs(UnmanagedType.LPArray, ArraySubType:=UnmanagedType.I2)> ByVal v2 As Short())
        Function glRenderMode(ByVal mode As UInteger) As Integer
        Sub glRotated(ByVal angle As Double, ByVal x As Double, ByVal y As Double, ByVal z As Double)
        Sub glRotatef(ByVal angle As Single, ByVal x As Single, ByVal y As Single, ByVal z As Single)
        Sub glScaled(ByVal x As Double, ByVal y As Double, ByVal z As Double)
        Sub glScalef(ByVal x As Single, ByVal y As Single, ByVal z As Single)
        Sub glScissor(ByVal x As Integer, ByVal y As Integer, ByVal width As Integer, ByVal height As Integer)
        Sub glSelectBuffer(ByVal size As Integer,
<MarshalAs(UnmanagedType.LPArray, ArraySubType:=UnmanagedType.U4)> ByVal buffer As UInteger())
        Sub glShadeModel(ByVal mode As UInteger)
        Sub glStencilFunc(ByVal func As UInteger, ByVal aRef As Integer, ByVal mask As UInteger)
        Sub glStencilMask(ByVal mask As UInteger)
        Sub glStencilOp(ByVal fail As UInteger, ByVal zfail As UInteger, ByVal zpass As UInteger)
        Sub glTexCoord1d(ByVal s As Double)
        Sub glTexCoord1dv(
<MarshalAs(UnmanagedType.LPArray, ArraySubType:=UnmanagedType.R8)> ByVal v As Double())
        Sub glTexCoord1f(ByVal s As Single)
        Sub glTexCoord1fv(
<MarshalAs(UnmanagedType.LPArray, ArraySubType:=UnmanagedType.R4)> ByVal v As Single())
        Sub glTexCoord1i(ByVal s As Integer)
        Sub glTexCoord1iv(
<MarshalAs(UnmanagedType.LPArray, ArraySubType:=UnmanagedType.I4)> ByVal v As Integer())
        Sub glTexCoord1s(ByVal s As Short)
        Sub glTexCoord1sv(
<MarshalAs(UnmanagedType.LPArray, ArraySubType:=UnmanagedType.I2)> ByVal v As Short())
        Sub glTexCoord2d(ByVal s As Double, ByVal t As Double)
        Sub glTexCoord2dv(
<MarshalAs(UnmanagedType.LPArray, ArraySubType:=UnmanagedType.R8)> ByVal v As Double())
        Sub glTexCoord2f(ByVal s As Single, ByVal t As Single)
        Sub glTexCoord2fv(
<MarshalAs(UnmanagedType.LPArray, ArraySubType:=UnmanagedType.R4)> ByVal v As Single())
        Sub glTexCoord2i(ByVal s As Integer, ByVal t As Integer)
        Sub glTexCoord2iv(
<MarshalAs(UnmanagedType.LPArray, ArraySubType:=UnmanagedType.I4)> ByVal v As Integer())
        Sub glTexCoord2s(ByVal s As Short, ByVal t As Short)
        Sub glTexCoord2sv(
<MarshalAs(UnmanagedType.LPArray, ArraySubType:=UnmanagedType.I2)> ByVal v As Short())
        Sub glTexCoord3d(ByVal s As Double, ByVal t As Double, ByVal r As Double)
        Sub glTexCoord3dv(
<MarshalAs(UnmanagedType.LPArray, ArraySubType:=UnmanagedType.R8)> ByVal v As Double())
        Sub glTexCoord3f(ByVal s As Single, ByVal t As Single, ByVal r As Single)
        Sub glTexCoord3fv(
<MarshalAs(UnmanagedType.LPArray, ArraySubType:=UnmanagedType.R4)> ByVal v As Single())
        Sub glTexCoord3i(ByVal s As Integer, ByVal t As Integer, ByVal r As Integer)
        Sub glTexCoord3iv(
<MarshalAs(UnmanagedType.LPArray, ArraySubType:=UnmanagedType.I4)> ByVal v As Integer())
        Sub glTexCoord3s(ByVal s As Short, ByVal t As Short, ByVal r As Short)
        Sub glTexCoord3sv(
<MarshalAs(UnmanagedType.LPArray, ArraySubType:=UnmanagedType.I2)> ByVal v As Short())
        Sub glTexCoord4d(ByVal s As Double, ByVal t As Double, ByVal r As Double, ByVal q As Double)
        Sub glTexCoord4dv(
<MarshalAs(UnmanagedType.LPArray, ArraySubType:=UnmanagedType.R8)> ByVal v As Double())
        Sub glTexCoord4f(ByVal s As Single, ByVal t As Single, ByVal r As Single, ByVal q As Single)
        Sub glTexCoord4fv(
<MarshalAs(UnmanagedType.LPArray, ArraySubType:=UnmanagedType.R4)> ByVal v As Single())
        Sub glTexCoord4i(ByVal s As Integer, ByVal t As Integer, ByVal r As Integer, ByVal q As Integer)
        Sub glTexCoord4iv(
<MarshalAs(UnmanagedType.LPArray, ArraySubType:=UnmanagedType.I4)> ByVal v As Integer())
        Sub glTexCoord4s(ByVal s As Short, ByVal t As Short, ByVal r As Short, ByVal q As Short)
        Sub glTexCoord4sv(
<MarshalAs(UnmanagedType.LPArray, ArraySubType:=UnmanagedType.I2)> ByVal v As Short())
        Sub glTexEnvf(ByVal target As UInteger, ByVal pname As UInteger, ByVal param As Single)
        Sub glTexEnvfv(ByVal target As UInteger, ByVal pname As UInteger,
<MarshalAs(UnmanagedType.LPArray, ArraySubType:=UnmanagedType.R4)> ByVal someParams As Single())
        Sub glTexEnvi(ByVal target As UInteger, ByVal pname As UInteger, ByVal param As Integer)
        Sub glTexEnviv(ByVal target As UInteger, ByVal pname As UInteger,
<MarshalAs(UnmanagedType.LPArray, ArraySubType:=UnmanagedType.I4)> ByVal someParams As Integer())
        Sub glTexGend(ByVal coord As UInteger, ByVal pname As UInteger, ByVal param As Double)
        Sub glTexGendv(ByVal coord As UInteger, ByVal pname As UInteger,
<MarshalAs(UnmanagedType.LPArray, ArraySubType:=UnmanagedType.R8)> ByVal someParams As Double())
        Sub glTexGenf(ByVal coord As UInteger, ByVal pname As UInteger, ByVal param As Single)
        Sub glTexGenfv(ByVal coord As UInteger, ByVal pname As UInteger,
<MarshalAs(UnmanagedType.LPArray, ArraySubType:=UnmanagedType.R4)> ByVal someParams As Single())
        Sub glTexGeni(ByVal coord As UInteger, ByVal pname As UInteger, ByVal param As Integer)
        Sub glTexGeniv(ByVal coord As UInteger, ByVal pname As UInteger,
<MarshalAs(UnmanagedType.LPArray, ArraySubType:=UnmanagedType.I4)> ByVal someParams As Integer())
        Sub glTexImage1D(ByVal target As UInteger, ByVal level As Integer, ByVal components As Integer, ByVal width As Integer, ByVal border As Integer, ByVal format As UInteger, ByVal type As UInteger,
<MarshalAs(UnmanagedType.LPArray, ArraySubType:=UnmanagedType.U1)> ByVal pixels As Byte())
        Sub glTexImage2D(ByVal target As UInteger, ByVal level As Integer, ByVal components As Integer, ByVal width As Integer, ByVal height As Integer, ByVal border As Integer, ByVal format As UInteger, ByVal type As UInteger,
<MarshalAs(UnmanagedType.LPArray, ArraySubType:=UnmanagedType.U1)> ByVal pixels As Byte())
        Sub glTexParameterf(ByVal target As UInteger, ByVal pname As UInteger, ByVal param As Single)
        Sub glTexParameterfv(ByVal target As UInteger, ByVal pname As UInteger,
<MarshalAs(UnmanagedType.LPArray, ArraySubType:=UnmanagedType.R4)> ByVal someParams As Single())
        Sub glTexParameteri(ByVal target As UInteger, ByVal pname As UInteger, ByVal param As Integer)
        Sub glTexParameteriv(ByVal target As UInteger, ByVal pname As UInteger,
<MarshalAs(UnmanagedType.LPArray, ArraySubType:=UnmanagedType.I4)> ByVal someParams As Integer())
        Sub glTranslated(ByVal x As Double, ByVal y As Double, ByVal z As Double)
        Sub glTranslatef(ByVal x As Single, ByVal y As Single, ByVal z As Single)
        Sub glVertex2d(ByVal x As Double, ByVal y As Double)
        Sub glVertex2dv(
<MarshalAs(UnmanagedType.LPArray, ArraySubType:=UnmanagedType.R8)> ByVal v As Double())
        Sub glVertex2f(ByVal x As Single, ByVal y As Single)
        Sub glVertex2fv(
<MarshalAs(UnmanagedType.LPArray, ArraySubType:=UnmanagedType.R4)> ByVal v As Single())
        Sub glVertex2i(ByVal x As Integer, ByVal y As Integer)
        Sub glVertex2iv(
<MarshalAs(UnmanagedType.LPArray, ArraySubType:=UnmanagedType.I4)> ByVal v As Integer())
        Sub glVertex2s(ByVal x As Short, ByVal y As Short)
        Sub glVertex2sv(
<MarshalAs(UnmanagedType.LPArray, ArraySubType:=UnmanagedType.I2)> ByVal v As Short())
        Sub glVertex3d(ByVal x As Double, ByVal y As Double, ByVal z As Double)
        Sub glVertex3dv(
<MarshalAs(UnmanagedType.LPArray, ArraySubType:=UnmanagedType.R8)> ByVal v As Double())
        Sub glVertex3f(ByVal x As Single, ByVal y As Single, ByVal z As Single)
        Sub glVertex3fv(
<MarshalAs(UnmanagedType.LPArray, ArraySubType:=UnmanagedType.R4)> ByVal v As Single())
        Sub glVertex3i(ByVal x As Integer, ByVal y As Integer, ByVal z As Integer)
        Sub glVertex3iv(
<MarshalAs(UnmanagedType.LPArray, ArraySubType:=UnmanagedType.I4)> ByVal v As Integer())
        Sub glVertex3s(ByVal x As Short, ByVal y As Short, ByVal z As Short)
        Sub glVertex3sv(
<MarshalAs(UnmanagedType.LPArray, ArraySubType:=UnmanagedType.I2)> ByVal v As Short())
        Sub glVertex4d(ByVal x As Double, ByVal y As Double, ByVal z As Double, ByVal w As Double)
        Sub glVertex4dv(
<MarshalAs(UnmanagedType.LPArray, ArraySubType:=UnmanagedType.R8)> ByVal v As Double())
        Sub glVertex4f(ByVal x As Single, ByVal y As Single, ByVal z As Single, ByVal w As Single)
        Sub glVertex4fv(
<MarshalAs(UnmanagedType.LPArray, ArraySubType:=UnmanagedType.R4)> ByVal v As Single())
        Sub glVertex4i(ByVal x As Integer, ByVal y As Integer, ByVal z As Integer, ByVal w As Integer)
        Sub glVertex4iv(
<MarshalAs(UnmanagedType.LPArray, ArraySubType:=UnmanagedType.I4)> ByVal v As Integer())
        Sub glVertex4s(ByVal x As Short, ByVal y As Short, ByVal z As Short, ByVal w As Short)
        Sub glVertex4sv(
<MarshalAs(UnmanagedType.LPArray, ArraySubType:=UnmanagedType.I2)> ByVal v As Short())
        Sub glViewport(ByVal x As Integer, ByVal y As Integer, ByVal width As Integer, ByVal height As Integer)
        Sub glPolygonOffset(ByVal factor As Single, ByVal units As Single)
        Sub glBindTexture(ByVal target As UInteger, ByVal texture As UInteger)
    End Interface

    <ComImport>
    <Guid("0002D282-0000-0000-C000-000000000046")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Interface IWGL
        Function wglCreateContext(ByVal hdc As IntPtr) As IntPtr
        Function wglDeleteContext(ByVal hglrc As IntPtr) As Integer
        Function wglGetCurrentContext() As IntPtr
        Function wglGetCurrentDC() As IntPtr
        Function wglMakeCurrent(ByVal hdc As IntPtr, ByVal hglrc As IntPtr) As Integer
        Function wglUseFontBitmapsA(ByVal hDC As IntPtr, ByVal first As UInteger, ByVal count As UInteger, ByVal listbase As UInteger) As Integer
        Function wglUseFontBitmapsW(ByVal hDC As IntPtr, ByVal first As UInteger, ByVal count As UInteger, ByVal listbase As UInteger) As Integer
        Function SwapBuffers(ByVal hdc As IntPtr) As Integer
    End Interface
End Namespace

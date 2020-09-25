'
' This file is maintained at https://github.com/SolidEdgeCommunity/SDK.
'
' Licensed under the MIT license.
' See https://github.com/SolidEdgeCommunity/SDK/blob/master/LICENSE for full
' license information.
'
' Required references:
'  - PresentationCore.dll
'  - PresentationFramework.dll
'  - WindowsBase.dll

Imports System

Namespace SolidEdgeSDK.AddIn
    Public Partial Class SolidEdgeAddIn
        Public Function AddWpfEdgeBarPage(Of TControl As {System.Windows.Controls.Page, New})(ByVal config As EdgeBarPageConfiguration) As EdgeBarPage
            Return AddWpfEdgeBarPage(Of TControl)(config:=config, document:=Nothing)
        End Function

        Public Function AddWpfEdgeBarPage(Of TControl As {System.Windows.Controls.Page, New})(ByVal config As EdgeBarPageConfiguration, ByVal document As SolidEdgeFramework.SolidEdgeDocument) As EdgeBarPage
            Dim WS_VISIBLE As UInteger = &H10000000
            Dim WS_CHILD As UInteger = &H40000000
            Dim WS_MAXIMIZE As UInteger = &H01000000
            Dim control As TControl = Activator.CreateInstance(Of TControl)()
            Dim edgeBarPage = AddEdgeBarPage(config:=config, controlHandle:=IntPtr.Zero, document:=document)
            Dim hwndSource = New System.Windows.Interop.HwndSource(New System.Windows.Interop.HwndSourceParameters With {
                .PositionX = 0,
                .PositionY = 0,
                .Height = 0,
                .Width = 0,
                .ParentWindow = edgeBarPage.Handle,
                .WindowStyle = CInt((WS_VISIBLE Or WS_CHILD Or WS_MAXIMIZE))
            }) With {
                .RootVisual = control
            }
            edgeBarPage.ChildObject = hwndSource
            Return edgeBarPage
        End Function
    End Class
End Namespace

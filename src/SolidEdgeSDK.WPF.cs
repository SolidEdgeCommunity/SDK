//
// This file is maintained at https://github.com/SolidEdgeCommunity/SDK.
//
// Licensed under the MIT license.
// See https://github.com/SolidEdgeCommunity/SDK/blob/master/LICENSE for full
// license information.
//
// Required references:
//  - PresentationCore.dll
//  - PresentationFramework.dll
//  - WindowsBase.dll

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

namespace SolidEdgeSDK.AddIn
{
    public partial class EdgeBarController : IDisposable
    {
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
    }
}

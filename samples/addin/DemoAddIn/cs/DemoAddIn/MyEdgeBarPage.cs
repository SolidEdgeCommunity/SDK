using SolidEdgeSDK.AddIn;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace DemoAddIn
{
    public class MyEdgeBarPage : EdgeBarPage
    {
        public MyEdgeBarPage()
        {
        }

        protected override void WndProc(ref Message m)
        {
            const int WM_SETFOCUS = 0x0007;
            const int WM_KILLFOCUS = 0x0008;
            const int WM_ENABLE = 0x000A;

            switch (m.Msg)
            {
                case WM_SETFOCUS:
                    break;
                case WM_KILLFOCUS:
                    break;
                case WM_ENABLE:
                    break;
            }

            if (this.ChildObject is NativeMessageEdgeBarControl control)
            {
                control.LogMessage(m);
            }

            base.WndProc(ref m);
        }
    }
}

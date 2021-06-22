using SolidEdgeSDK.AddIn;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace DemoAddIn
{
    public partial class NativeMessageEdgeBarControl : UserControl
    {
        public NativeMessageEdgeBarControl()
        {
            InitializeComponent();
        }

        public void LogMessage(Message m)
        {
            tbMessages.AppendText($"{m.ToString()}{Environment.NewLine}");
        }

        public EdgeBarPage EdgeBarPage { get; set; }
    }
}

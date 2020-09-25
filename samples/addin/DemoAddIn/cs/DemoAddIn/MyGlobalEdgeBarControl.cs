using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace DemoAddIn
{
    public partial class MyGlobalEdgeBarControl : UserControl
    {
        public MyGlobalEdgeBarControl()
        {
            InitializeComponent();
        }

        private void MyGlobalEdgeBarControl_Load(object sender, EventArgs e)
        {
        }

        public SolidEdgeFramework.Application Application { get; set; }
    }
}
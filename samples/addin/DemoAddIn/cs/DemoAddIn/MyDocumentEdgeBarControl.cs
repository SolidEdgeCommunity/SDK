using SolidEdgeFramework;
using SolidEdgeSDK;
using SolidEdgeSDK.Extensions;
using SolidEdgeSDK.InteropServices;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Reflection;


namespace DemoAddIn
{
    public partial class MyDocumentEdgeBarControl : UserControl,
        SolidEdgeFramework.ISEDocumentEvents
    {
        private SolidEdgeFramework.SolidEdgeDocument _document;

        public MyDocumentEdgeBarControl()
        {
            InitializeComponent();
        }

        private void MyEdgeBarControl_Load(object sender, EventArgs e)
        {
            // Trick to use the default system font.
            Font = SystemFonts.MessageBoxFont;

            ComEventsManager = new ComEventsManager(this);
        }

        private void listView_SelectedIndexChanged(object sender, EventArgs e)
        {
            propertyGrid.SelectedObject = null;

            if (listView.SelectedItems.Count > 0)
            {
                var item = listView.SelectedItems[0];
                propertyGrid.SelectedObject = item.Tag;
            }
        }

        private void UpdateListView(object[] items)
        {
            listView.Items.Clear();
            propertyGrid.SelectedObject = null;

            var listViewItems = new List<ListViewItem>();

            for (int i = 0; i < items.Length; i++)
            {
                listViewItems.Add(new ListViewItem(new string[] { $"{i + 1}", ComObject.GetComTypeFullName(items[i]) })
                {
                    Tag = items[i]
                });
            }

            listView.Items.AddRange(listViewItems.ToArray());

            foreach (ColumnHeader header in listView.Columns)
            {
                header.Width = -2;
            }

            if (listView.Items.Count > 0)
            {
                listView.Items[0].Selected = true;
            }
        }

        #region ISEDocumentEvents

        void ISEDocumentEvents.BeforeClose()
        {
            ComEventsManager.DetachAll();
        }

        void ISEDocumentEvents.BeforeSave()
        {
        }

        void ISEDocumentEvents.AfterSave()
        {
        }

        void ISEDocumentEvents.SelectSetChanged(object SelectSet)
        {
            var items = new object[] { };

            if (SelectSet is SolidEdgeFramework.SelectSet selectSet)
            {
                items = selectSet.OfType<object>().ToArray();
            }

            if (items.Length == 0)
            {
                items = new object[] { Document };
            }

            UpdateListView(items);
        }

        #endregion

        public SolidEdgeFramework.Application Application { get; set; }

        public SolidEdgeFramework.SolidEdgeDocument Document
        {
            get { return _document; }
            set
            {
                _document = value;

                if (_document == null)
                {
                    propertyGrid.SelectedObject = null;
                    ComEventsManager.DetachAll();
                }
                else
                {
                    propertyGrid.SelectedObject = _document;
                    ComEventsManager.Attach<ISEDocumentEvents>(_document.DocumentEvents);

                    UpdateListView(new object[] { _document });
                }
            }
        }

        private ComEventsManager ComEventsManager { get; set; }
    }
}

using DemoAddIn.Properties;
using SolidEdgeFramework;
using SolidEdgeSDK.AddIn;
using SolidEdgeSDK.Extensions;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace DemoAddIn
{
    class My3dRibbon : Ribbon,
        SolidEdgeFramework.ISEAddInEventsEx,
        SolidEdgeFramework.ISEAddInEventsEx2
    {
        enum CommandIds : int
        {
            Save,
            Folder,
            Monitor,
            Box,
            Camera,
            Photograph,
            Favorites,
            Printer,
            Tools,
            CommandPrompt,
            Notepad,
            Help,
            Search,
            Question,
            CheckBox1,
            CheckBox2,
            CheckBox3,
            RadioButton1,
            RadioButton2,
            RadioButton3,
            BoundingBox,
            OpenGlBoxes,
            GdiPlus
        }

        public My3dRibbon()
            : base()
        {
        }

        public override void Initialize()
        {
            // Build your Ribbon here.
            // IMPORTANT: If and when you make changes to your ribbon, you must increment AddInInstance.GuiVersion in MyAddIn.cs.
            // If you fail to do this, your ribbon will likely no display correctly.

            var tabs = new RibbonTab[]
            {
                new RibbonTab(this, Resources.MyRibbon1)
                {
                    Groups = new RibbonGroup[]
                    {
                        new RibbonGroup("Group 1")
                        {
                            Controls = new RibbonControl[]
                            {
                                new RibbonButton((int)CommandIds.Save, "Save", "Save Screentip", "Save Supertip", NativeResources.PNG.Save_16, RibbonButtonSize.Normal),
                                new RibbonButton((int)CommandIds.Folder, "Folder", "Folder Screentip", "Folder Supertip", NativeResources.PNG.Folder_16, RibbonButtonSize.Normal),
                                new RibbonButton((int)CommandIds.Monitor, "Monitor", "Monitor Screentip", "Monitor Supertip", NativeResources.PNG.Monitor_16, RibbonButtonSize.Normal),
                                new RibbonButton((int)CommandIds.Box, "Box", "Box Screentip", "Box Supertip", NativeResources.PNG.Box_32, RibbonButtonSize.Large)
                            }
                        },
                        new RibbonGroup("Group 2")
                        {
                            Controls = new RibbonControl[]
                            {
                                new RibbonButton((int)CommandIds.Camera, "Camera", "Camera Screentip", "Camera Supertip", NativeResources.PNG.Camera_32, RibbonButtonSize.Large),
                                new RibbonButton((int)CommandIds.Photograph, "Photograph", "Photograph Screentip", "Photograph Supertip", NativeResources.PNG.Photograph_32, RibbonButtonSize.Large),
                                new RibbonButton((int)CommandIds.Favorites, "Favorites", "Favorites Screentip", "Favorites Supertip", NativeResources.PNG.Favorites_32, RibbonButtonSize.Large),
                                new RibbonButton((int)CommandIds.Printer, "Printer", "Printer Screentip", "Printer Supertip", NativeResources.PNG.Printer_32, RibbonButtonSize.Large)
                            }
                        },
                        new RibbonGroup("Group 3")
                        {
                            Controls = new RibbonControl[]
                            {
                                new RibbonButton((int)CommandIds.Tools, "Tools", "Tools Screentip", "Tools Supertip", NativeResources.PNG.Tools_32, RibbonButtonSize.Large),
                                new RibbonButton((int)CommandIds.CommandPrompt, "Command Prompt", "Command Prompt Screentip", "Command Prompt Supertip", NativeResources.PNG.CommandPrompt_32, RibbonButtonSize.Large),
                                new RibbonButton((int)CommandIds.Notepad, "Notepad", "Notepad Screentip", "Notepad Supertip", NativeResources.PNG.Notepad_32, RibbonButtonSize.Large)
                            }
                        },
                        new RibbonGroup("Group 4")
                        {
                            Controls = new RibbonControl[]
                            {
                                new RibbonButton((int)CommandIds.Help, "Help", "Help Screentip", "Help Supertip", NativeResources.PNG.Help_32, RibbonButtonSize.Large),
                                new RibbonButton((int)CommandIds.Search, "Search", "Search Screentip", "Search Supertip", NativeResources.PNG.Search_32, RibbonButtonSize.Large)
                            }
                        },
                        new RibbonGroup("Group 5")
                        {
                            Controls = new RibbonControl[]
                            {
                                new RibbonButton((int)CommandIds.Question, "Question", "Question Screentip", "Question Supertip", NativeResources.PNG.Question_32, RibbonButtonSize.Large)
                            }
                        },
                        new RibbonGroup("Group 6")
                        {
                            Controls = new RibbonControl[]
                            {
                                new RibbonCheckBox((int)CommandIds.CheckBox1, "Checkbox1", "Checkbox1 Screentip", "Checkbox1 Supertip"),
                                new RibbonCheckBox((int)CommandIds.CheckBox2, "Checkbox2", "Checkbox2 Screentip", "Checkbox2 Supertip"),
                                new RibbonCheckBox((int)CommandIds.CheckBox3, "Checkbox3", "Checkbox3 Screentip", "Checkbox3 Supertip"),
                                new RibbonRadioButton((int)CommandIds.RadioButton1, "Radiobutton1", "Radiobutton1 Screentip", "Radiobutton1 Supertip"),
                                new RibbonRadioButton((int)CommandIds.RadioButton2, "Radiobutton2", "Radiobutton2 Screentip", "Radiobutton2 Supertip"),
                                new RibbonRadioButton((int)CommandIds.RadioButton3, "Radiobutton3", "Radiobutton3 Screentip", "Radiobutton3 Supertip")
                            }
                        },
                        new RibbonGroup("Overlays")
                        {
                            Controls = new RibbonControl[]
                            {
                                new RibbonButton((int)CommandIds.BoundingBox, "Bounding Box", "Bounding Box Screentip", "Bounding Box Supertip", NativeResources.PNG.BoundingBox_32, RibbonButtonSize.Large),
                                new RibbonButton((int)CommandIds.OpenGlBoxes, "OpenGL Boxes", "OpenGL Boxes Screentip", "OpenGL Boxes Supertip", NativeResources.PNG.Boxes_32, RibbonButtonSize.Large),
                                new RibbonButton((int)CommandIds.GdiPlus, "GDI+", "GDI+ Screentip", "GDI+ Supertip", NativeResources.PNG.GdiPlus_32, RibbonButtonSize.Large)
                            }
                        }
                    }
                }
            };

            base.Initialize(tabs);
        }

        #region ISEAddInEvents, ISEAddInEventsEx & ISEAddInEventsEx2

        public void OnCommand(int CommandID)
        {
            var Context = default(ShortCutMenuContextConstants);
            var ActiveDocumentType = default(DocumentTypeConstants);
            object pActiveDocument = null;
            object pActiveWindow = null;
            object pActiveSelectSet = null;

            // Forward call to newer implementation.
            OnCommand(CommandID, Context, ActiveDocumentType, pActiveDocument, pActiveWindow, pActiveSelectSet);
        }

        public void OnCommand(int CommandID, ShortCutMenuContextConstants Context, DocumentTypeConstants ActiveDocumentType, object pActiveDocument, object pActiveWindow, object pActiveSelectSet)
        {
            var myAddIn = (MyAddIn)this.SolidEdgeAddIn;
            var application = myAddIn.Application;
            var control = GetControl(CommandID);
            var commandId = (CommandIds)control.CommandId;

            switch (commandId)
            {
                case CommandIds.Save:
                    ShowSaveFileDialogDemo();
                    break;
                case CommandIds.Folder:
                    ShowFolderBrowserDialogDemo();
                    break;
                case CommandIds.Tools:
                    application.StartCommand(SolidEdgeConstants.PartCommandConstants.PartToolsOptions);
                    break;
                case CommandIds.Printer:
                    application.StartCommand(SolidEdgeConstants.PartCommandConstants.PartFilePrint);
                    break;
                case CommandIds.Help:
                    application.StartCommand(SolidEdgeConstants.PartCommandConstants.PartHelpSolidEdgeontheWeb);
                    break;
                case CommandIds.CommandPrompt:
                    Process.Start("cmd.exe");
                    break;
                case CommandIds.Notepad:
                    Process.Start("notepad.exe");
                    break;
                case CommandIds.BoundingBox:
                    control.Checked = !control.Checked;
                    myAddIn.My3dViewOverlay.ShowBoundingBox = control.Checked;
                    break;
                case CommandIds.GdiPlus:
                    control.Checked = !control.Checked;
                    myAddIn.My3dViewOverlay.ShowGDIPlusDemo = control.Checked;
                    break;
                case CommandIds.OpenGlBoxes:
                    control.Checked = !control.Checked;
                    myAddIn.My3dViewOverlay.ShowOpenGlBoxesDemo = control.Checked;
                    break;
                default:
                    if (control is SolidEdgeSDK.AddIn.RibbonButton)
                    {
                        ShowCustomDialogDemo();
                    }
                    break;
            }

            // Example of how to determine what type of control was clicked.
            if (control is RibbonButton ribbonButton)
            {
            }
            else if (control is RibbonCheckBox checkBox)
            {
                // Demo toggling state.
                checkBox.Checked = !checkBox.Checked;
            }
            else if (control is RibbonRadioButton radioButton)
            {
                // Demo toggling state.
                radioButton.Checked = !radioButton.Checked;
            }
        }

        public void OnCommandHelp(int hFrameWnd, int HelpCommandID, int CommandID)
        {
        }

        public void OnCommandOnLineHelp(int HelpCommandID, int CommandID, out string HelpURL)
        {
            HelpURL = null;

            var myAddIn = (MyAddIn)this.SolidEdgeAddIn;
            var control = this.GetControl(CommandID);

            HelpURL = control?.WebHelpURL;
        }

        public void OnCommandUpdateUI(int CommandID, ref int CommandFlags, out string MenuItemText, ref int BitmapID)
        {
            var Context = default(ShortCutMenuContextConstants);
            var ActiveDocumentType = default(DocumentTypeConstants);
            object pActiveDocument = null;
            object pActiveWindow = null;
            object pActiveSelectSet = null;

            // Forward call to newer implementation.
            OnCommandUpdateUI(CommandID, Context, ActiveDocumentType, pActiveDocument, pActiveWindow, pActiveSelectSet, ref CommandFlags, out MenuItemText, ref BitmapID);
        }

        public void OnCommandUpdateUI(int CommandID, ShortCutMenuContextConstants Context, DocumentTypeConstants ActiveDocumentType, object pActiveDocument, object pActiveWindow, object pActiveSelectSet, ref int CommandFlags, out string MenuItemText, ref int BitmapID)
        {
            MenuItemText = null;
            var myAddIn = (MyAddIn)this.SolidEdgeAddIn;
            var flags = default(SolidEdgeConstants.SECommandActivation);
            var control = Controls.First(x => x.CommandId == CommandID);
            var commandId = (CommandIds)CommandID;

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
                CommandFlags = (int)flags;
                MenuItemText = control.Label;
            }
        }

        #endregion

        private void ShowSaveFileDialogDemo()
        {
            using (var dialog = new SaveFileDialog())
            {
                // The ShowDialog() extension method is exposed by:
                // using SolidEdgeFramework.Extensions (SolidEdge.Community.dll)
                if (SolidEdgeAddIn.Application.ShowDialog(dialog) == DialogResult.OK)
                {
                }
            }
        }

        private void ShowFolderBrowserDialogDemo()
        {
            using (var dialog = new FolderBrowserDialog())
            {
                // The ShowDialog() extension method is exposed by:
                // using SolidEdgeFramework.Extensions (SolidEdge.Community.dll)
                if (SolidEdgeAddIn.Application.ShowDialog(dialog) == DialogResult.OK)
                {
                }
            }
        }

        private void ShowCustomDialogDemo()
        {
            using (var dialog = new MyCustomDialog())
            {
                // The ShowDialog() extension method is exposed by:
                // using SolidEdgeFramework.Extensions (SolidEdge.Community.dll)
                if (SolidEdgeAddIn.Application.ShowDialog(dialog) == DialogResult.OK)
                {
                }
            }
        }
    }
}

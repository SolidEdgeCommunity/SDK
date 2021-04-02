using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DemoAddIn
{
    /// <summary>
    /// Command Ids
    /// </summary>
    /// <remarks>
    /// When you make changes to your commands, you must increment AddInInstance.GuiVersion in MyAddIn.cs.
    /// </remarks>
    public enum MyCommandIds : int
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
        GdiPlus,
        About
    }

    public class MyCommand
    {
        public MyCommand()
        {
        }

        public MyCommand(MyCommandIds commandId, string category, string group, string command, int imageId, SolidEdgeFramework.SeButtonStyle buttonStyle)
        {
            CommandId = commandId;
            Category = category;
            Group = group;
            CommandString = command;
            ImageId = imageId;
            ButtonStyle = buttonStyle;
        }

        public MyCommandIds CommandId { get; set; }
        public int SolidEdgeCommandId { get; set; }
        public string Category { get; set; }
        public string Group { get; set; }
        public string CommandString { get; set; }
        public int ImageId { get; set; }
        public SolidEdgeFramework.SeButtonStyle ButtonStyle { get; set; } = SolidEdgeFramework.SeButtonStyle.seButtonAutomatic;
        public string RuntimeCategoryName { get; set; }
        public string RuntimeCommandName { get; set; }
        public bool Enabled { get; set; } = true;
        public bool Checked { get; set; }

        public override string ToString()
        {
            if (String.IsNullOrWhiteSpace(RuntimeCommandName) == false)
            {
                return RuntimeCommandName;
            }
            else if (String.IsNullOrWhiteSpace(CommandString) == false)
            {
                return CommandString;
            }

            return base.ToString();
        }
    }
}

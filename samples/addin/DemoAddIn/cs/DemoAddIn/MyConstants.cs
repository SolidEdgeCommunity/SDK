using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace DemoAddIn
{
    static class MyConstants
    {
        /// <summary>
        /// Unique GUID for the addin.
        /// </summary>
        /// <remarks>
        /// You must generate a new unique GUID for your addin.
        /// </remarks>
        /// <example>
        /// A6D088BE-0640-480B-93D2-FC9A7F49ADF5
        /// </example>
        public const string AddInGuid = "03504538-C2A2-44BC-994C-42876F57363F";

        /// <summary>
        /// Unique PROGID for the addin.
        /// </summary>
        /// <remarks>
        /// You must specify a unique PROGID for your addin.
        /// </remarks>
        public const string AddInProgId = "SolidEdgeCommunity.DemoAddIn";
    }

    /// <summary>
    /// Native.rc constants.
    /// </summary>
    /// <remarks>
    /// This class and Native.rc must be manually maintained.
    /// </remarks>
    static class NativeResources
    {
        /// <summary>
        /// Native BMP resources.
        /// </summary>
        public static class BMP
        {
            //public const int Example = 5000;
        }

        /// <summary>
        /// Native PNG resources.
        /// </summary>        
        public static class PNG
        {
            public const int BoundingBox_32 = 6000;
            public const int Boxes_32 = 6001;
            public const int Box_32 = 6002;
            public const int Camera_32 = 6003;
            public const int CommandPrompt_32 = 6004;
            public const int EdgeBar_20 = 6005;
            public const int Favorites_32 = 6006;
            public const int Folder_16 = 6007;
            public const int GdiPlus_32 = 6008;
            public const int Help_32 = 6009;
            public const int Monitor_16 = 6010;
            public const int Notepad_32 = 6011;
            public const int Photograph_32 = 6012;
            public const int Printer_32 = 6013;
            public const int Question_32 = 6014;
            public const int Save_16 = 6015;
            public const int Search_32 = 6016;
            public const int Tools_32 = 6017;
        }
    }
}

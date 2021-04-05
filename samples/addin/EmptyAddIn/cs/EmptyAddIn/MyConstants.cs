using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EmptyAddIn
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
        /// 7EC6031B-F90B-4AE2-8999-40EFED6664BB
        /// </example>
        public const string AddInGuid = "7EC6031B-F90B-4AE2-8999-40EFED6664BB";

        /// <summary>
        /// Unique PROGID for the addin.
        /// </summary>
        /// <remarks>
        /// You must specify a unique PROGID for your addin.
        /// </remarks>
        public const string AddInProgId = "SolidEdgeCommunity.EmptyAddIn";
    }
}

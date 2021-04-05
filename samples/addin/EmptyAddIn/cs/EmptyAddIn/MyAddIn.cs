using SolidEdgeSDK.AddIn;
using SolidEdgeSDK.InteropServices;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;

// Please refer to Instructions.txt.
namespace EmptyAddIn
{
    // MUST UPDATE MyConstants.cs!
    [Guid(MyConstants.AddInGuid), ProgId(MyConstants.AddInProgId), ComVisible(true)]
    [AddInImplementedCategory(CATID.SolidEdgeAddIn)]
    [AddInEnvironmentCategory(CATID.SEApplication)]
    [AddInEnvironmentCategory(CATID.SEAllDocumentEnvrionments)]
    [AddInCulture("en-US")]
    public class MyAddIn : SolidEdgeAddIn
    {
        #region SolidEdgeAddIn implementation

        public override void OnConnection(SolidEdgeFramework.SeConnectMode ConnectMode)
        {
        }

        public override void OnConnectToEnvironment(Guid EnvCatID, SolidEdgeFramework.Environment environment, bool firstTime)
        {
        }

        public override void OnDisconnection(SolidEdgeFramework.SeDisconnectMode DisconnectMode)
        {
        }

        #endregion

        #region Registration (regasm.exe) functions

        [ComRegisterFunction]
        static void OnComRegister(Type t)
        {
            if (Guid.Equals(t.GUID, typeof(MyAddIn).GUID))
            {
                var implementedCategories = System.Attribute
                    .GetCustomAttributes(t, typeof(AddInImplementedCategoryAttribute))
                    .Cast<AddInImplementedCategoryAttribute>()
                    .Select(a => a.Value)
                    .ToArray();

                var environmentCategories = System.Attribute
                    .GetCustomAttributes(t, typeof(AddInEnvironmentCategoryAttribute))
                    .Cast<AddInEnvironmentCategoryAttribute>()
                    .Select(a => a.Value)
                    .ToArray();

                var cultures = System.Attribute
                    .GetCustomAttributes(t, typeof(AddInCultureAttribute))
                    .Cast<AddInCultureAttribute>()
                    .Select(a => a.Value);

                var descriptors = cultures
                    .Select(culture => new AddInDescriptor(culture, typeof(Properties.Resources)))
                    .ToArray();

                var settings = new ComRegistrationSettings(t)
                {
                    Enabled = true,
                    ImplementedCategories = implementedCategories,
                    EnvironmentCategories = environmentCategories,
                    Descriptors = descriptors
                };

                ComRegisterSolidEdgeAddIn(settings);
            }
        }

        [ComUnregisterFunction]
        static void OnComUnregister(Type t)
        {
            if (Guid.Equals(t.GUID, typeof(MyAddIn).GUID))
            {
                ComUnregisterSolidEdgeAddIn(t);
            }
        }

        #endregion
    }
}
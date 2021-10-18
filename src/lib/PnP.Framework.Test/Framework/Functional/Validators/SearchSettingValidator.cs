using PnP.Framework.Provisioning.Providers.Xml;
using System;

namespace PnP.Framework.Test.Framework.Functional.Validators
{

    public class SerializedSearchSetting
    {
        public string SchemaXml { get; set; }
    }

    public class SearchSettingValidator : ValidatorBase
    {
        #region construction        
        public SearchSettingValidator() : base()
        {
            // optionally override schema version
            SchemaVersion = XMLConstants.PROVISIONING_SCHEMA_NAMESPACE_2021_03;
        }
        #endregion

        #region Validation logic
        public bool Validate(String sourceSearchSetting, string targetSearchSetting)
        {
            if (!String.IsNullOrEmpty(sourceSearchSetting))
            {
                if (String.IsNullOrEmpty(targetSearchSetting))
                {
                    return false;
                }
            }

            return true;
        }
        #endregion
    }
}

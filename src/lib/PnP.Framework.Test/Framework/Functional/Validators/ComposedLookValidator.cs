using Microsoft.VisualStudio.TestTools.UnitTesting;
using PnP.Framework.Provisioning.Model;

namespace PnP.Framework.Tests.Framework.Functional.Validators
{
    [TestClass]
    public class ComposedLookValidator : ValidatorBase
    {
        public bool Validate(ComposedLook source, ComposedLook target)
        {
            if (!source.BackgroundFile.Trim().ToLower().Equals(target.BackgroundFile.Trim().ToLower())) { return false; }

            if (!source.ColorFile.Trim().ToLower().Equals(target.ColorFile.Trim().ToLower())) { return false; }

            if (!source.FontFile.Trim().ToLower().Equals(target.FontFile.Trim().ToLower())) { return false; }

            if (!source.Name.Trim().ToLower().Equals(target.Name.Trim().ToLower())) { return false; }

            if (!source.Version.Equals(target.Version)) { return false; }

            return true;
        }
    }
}

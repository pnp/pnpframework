using Microsoft.SharePoint.Client;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using PnP.Framework.Modernization.Publishing;
using System;
using System.Collections.Generic;
using System.Linq;

namespace PnP.Framework.Modernization.Tests.Transform.Publishing
{
    [TestClass]
    public class EnterpriseWikiMetadataTests
    {
        [TestMethod]
        public void OptedInEnterpriseWikiMappingCopiesAllFiveSameNameFields()
        {
            var model = new PageLayoutManager().LoadDefaultPageLayoutMappingFile(true);
            var layout = model.PageLayouts.Single(l => l.Name == "EnterpriseWiki");
            foreach (var name in new[] { "Wiki_x0020_Page_x0020_Categories", "RatedBy", "Ratings", "RatingCount", "AverageRating" })
            {
                Assert.AreEqual(name, layout.MetaData.Field.Single(f => f.Name == name).TargetFieldName);
            }
            Assert.IsTrue(EnterpriseWikiRatings.UsesSameNameMapping(layout));
            Assert.IsFalse(EnterpriseWikiRatings.UsesSameNameMapping(model.PageLayouts.First(l => l.Name != "EnterpriseWiki")));
        }

        [TestMethod]
        public void CustomRatingFunctionsAreNotOverridden()
        {
            var layout = new PageLayoutManager().LoadDefaultPageLayoutMappingFile(true).PageLayouts.Single(l => l.Name == "EnterpriseWiki");
            layout.MetaData.Field.Single(f => f.Name == "Ratings").Functions = "CustomRatings()";
            Assert.IsFalse(EnterpriseWikiRatings.UsesSameNameMapping(layout));
        }

        [TestMethod]
        public void SwitchingFlagRestoresLegacyMappingAndSidebarBehavior()
        {
            using (var context = new ClientContext("https://contoso.sharepoint.com/sites/wiki"))
            {
                var transformator = new PublishingPageTransformator(context, context, new PageTransformation(), null);
                var information = new PublishingPageTransformationInformation(null) { IsEnterpriseWikiPage = true };
                var legacy = transformator.GetPublishingModel(information);
                var legacyLayout = legacy.PageLayouts.Single(l => l.Name == "EnterpriseWiki");
                Assert.IsFalse(EnterpriseWikiRatings.UsesSameNameMapping(legacyLayout));
                Assert.IsTrue(legacyLayout.MetaData.Field.All(f => string.IsNullOrEmpty(f.TargetFieldName)));
                Assert.IsFalse(information.ShouldIncludeEnterpriseWikiSidebar);

                information.EnableEnterpriseWikiCapabilities = true;
                var enabled = transformator.GetPublishingModel(information);
                Assert.IsTrue(EnterpriseWikiRatings.UsesSameNameMapping(enabled.PageLayouts.Single(l => l.Name == "EnterpriseWiki")));
                Assert.IsTrue(information.ShouldIncludeEnterpriseWikiSidebar);

                information.EnableEnterpriseWikiCapabilities = false;
                Assert.AreSame(legacy, transformator.GetPublishingModel(information));
                Assert.IsTrue(legacyLayout.MetaData.Field.All(f => string.IsNullOrEmpty(f.TargetFieldName)));
                Assert.IsFalse(information.ShouldIncludeEnterpriseWikiSidebar);
            }
        }

        [TestMethod]
        public void CustomMetadataMappingIsHonoredInBothModes()
        {
            var custom = new PageLayoutManager().LoadDefaultPageLayoutMappingFile();
            var layout = custom.PageLayouts.Single(l => l.Name == "EnterpriseWiki");
            layout.MetaData.Field.Single(f => f.Name == EnterpriseWikiPagePreparation.CategoryFieldName).TargetFieldName = "CustomCategories";
            custom.PageLayouts = new[] { layout };
            using (var context = new ClientContext("https://contoso.sharepoint.com/sites/wiki"))
            {
                var transformator = new PublishingPageTransformator(context, context, new PageTransformation(), custom);
                foreach (var enabled in new[] { false, true })
                {
                    var information = new PublishingPageTransformationInformation(null) { EnableEnterpriseWikiCapabilities = enabled };
                    var selected = transformator.GetPublishingModel(information).PageLayouts.Single(l => l.Name == "EnterpriseWiki");
                    Assert.AreEqual("CustomCategories", selected.MetaData.Field.Single(f => f.Name == EnterpriseWikiPagePreparation.CategoryFieldName).TargetFieldName);
                    Assert.IsFalse(EnterpriseWikiRatings.UsesSameNameMapping(selected));
                }
            }
        }

        [TestMethod]
        public void EmptyHistoryIsValid()
        {
            var snapshot = EnterpriseWikiRatings.Read(new Dictionary<string, object>());
            Assert.AreEqual(0, snapshot.Voters.Length);
            Assert.AreEqual(0, snapshot.Scores.Length);
            Assert.AreEqual(0, snapshot.Count);
            Assert.AreEqual(0, snapshot.Average);
        }

        [DataTestMethod]
        [DataRow("6,")]
        [DataRow("3,,")]
        public void RejectsInvalidScoreSerialization(string ratings)
        {
            Assert.ThrowsException<InvalidOperationException>(() => EnterpriseWikiRatings.Read(History(new[] { 1 }, ratings, 1, 3)));
        }

        [TestMethod]
        public void RejectsIncompleteOrInconsistentHistory()
        {
            Assert.ThrowsException<InvalidOperationException>(() => EnterpriseWikiRatings.Read(History(Array.Empty<int>(), "5,5,", 2, 5)));
            Assert.ThrowsException<InvalidOperationException>(() => EnterpriseWikiRatings.Read(History(new[] { 1 }, "3,", 2, 3)));
            Assert.ThrowsException<InvalidOperationException>(() => EnterpriseWikiRatings.Read(History(new[] { 1 }, "3,", 1, 4)));
            Assert.ThrowsException<InvalidOperationException>(() => EnterpriseWikiRatings.Read(History(new[] { 1, 1 }, "3,3,", 2, 3)));
        }

        [TestMethod]
        public void RemapsUsersWithoutReorderingOrDroppingScores()
        {
            var source = EnterpriseWikiRatings.Read(History(new[] { 41, 12 }, "2,5,", 2, 3.5));
            var target = source.MapVoters(v => new FieldUserValue { LookupId = v.LookupId == 41 ? 8 : 99 });
            CollectionAssert.AreEqual(new[] { 8, 99 }, target.Voters.Select(v => v.LookupId).ToArray());
            CollectionAssert.AreEqual(new[] { 2, 5 }, target.Scores);
            Assert.AreEqual("2,5,", target.Ratings);
            Assert.AreEqual(3.5, target.Average);
            CollectionAssert.AreEqual(new[] { 41, 12 }, source.Voters.Select(v => v.LookupId).ToArray());
            Assert.ThrowsException<InvalidOperationException>(() => source.MapVoters(_ => null));
            Assert.ThrowsException<InvalidOperationException>(() => source.MapVoters(_ => new FieldUserValue { LookupId = 8 }));
        }

        private static Dictionary<string, object> History(int[] voters, string scores, int count, double average)
        {
            return new Dictionary<string, object>
            {
                ["RatedBy"] = voters.Select(id => new FieldUserValue { LookupId = id }).ToArray(),
                ["Ratings"] = scores,
                ["RatingCount"] = count,
                ["AverageRating"] = average
            };
        }
    }
}

using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Taxonomy;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;

namespace PnP.Framework.Modernization.Tests.Extensions
{
    [TestClass]
    public class ListItemExtensionsTests
    {
        [TestMethod]
        public void GetDifferences_WithIdenticalValues_ReturnsEmptyList()
        {
            // Arrange
            var listItem = CreateMockListItem(new Dictionary<string, object>
            {
                ["Title"] = "Test Title",
                ["Description"] = "Test Description"
            });

            var newValues = new Dictionary<string, object>
            {
                ["Title"] = "Test Title",
                ["Description"] = "Test Description"
            };

            // Act
            var differences = listItem.GetDifferences(newValues, false);

            // Assert
            Assert.AreEqual(0, differences.Count);
        }

        [TestMethod]
        public void GetDifferences_WithDifferentValues_ReturnsExpectedChanges()
        {
            // Arrange
            var listItem = CreateMockListItem(new Dictionary<string, object>
            {
                ["Title"] = "Old Title",
                ["Description"] = "Old Description"
            });

            var newValues = new Dictionary<string, object>
            {
                ["Title"] = "New Title",
                ["Description"] = "Old Description"
            };

            // Act
            var differences = listItem.GetDifferences(newValues, false);

            // Assert
            Assert.AreEqual(1, differences.Count);
            Assert.AreEqual("Title", differences[0].FieldInternalName);
            Assert.AreEqual("New Title", differences[0].NewValue);
            Assert.AreEqual("Old Title", differences[0].CurrentValue);
        }

        [TestMethod]
        public void GetDifferences_WithEmptyStringTreatAsNull_True_HandlesEmptyStringsAsNull()
        {
            // Arrange
            var listItem = CreateMockListItem(new Dictionary<string, object>
            {
                ["Title"] = "",
                ["Description"] = null
            });

            var newValues = new Dictionary<string, object>
            {
                ["Title"] = null,
                ["Description"] = ""
            };

            // Act
            var differences = listItem.GetDifferences(newValues, true);

            // Assert
            Assert.AreEqual(0, differences.Count, "Empty strings and nulls should be treated as equal when treatEmptyStringAsNull is true");
        }

        [TestMethod]
        public void GetDifferences_WithEmptyStringTreatAsNull_False_HandlesEmptyStringsAsDifferent()
        {
            // Arrange
            var listItem = CreateMockListItem(new Dictionary<string, object>
            {
                ["Title"] = "",
                ["Description"] = null
            });

            var newValues = new Dictionary<string, object>
            {
                ["Title"] = null,
                ["Description"] = ""
            };

            // Act
            var differences = listItem.GetDifferences(newValues, false);

            // Assert
            Assert.AreEqual(2, differences.Count, "Empty strings and nulls should be treated as different when treatEmptyStringAsNull is false");
        }

        [TestMethod]
        public void GetDifferences_WithNullCurrentValue_ReturnsCorrectChange()
        {
            // Arrange
            var listItem = CreateMockListItem(new Dictionary<string, object>
            {
                ["Title"] = null
            });

            var newValues = new Dictionary<string, object>
            {
                ["Title"] = "New Title"
            };

            // Act
            var differences = listItem.GetDifferences(newValues, false);

            // Assert
            Assert.AreEqual(1, differences.Count);
            Assert.AreEqual("Title", differences[0].FieldInternalName);
            Assert.AreEqual("New Title", differences[0].NewValue);
            Assert.IsNull(differences[0].CurrentValue);
        }

        [TestMethod]
        public void GetDifferences_WithMissingField_TreatsAsNull()
        {
            // Arrange
            var listItem = CreateMockListItem(new Dictionary<string, object>());

            var newValues = new Dictionary<string, object>
            {
                ["Title"] = "New Title"
            };

            // Act
            var differences = listItem.GetDifferences(newValues, false);

            // Assert
            Assert.AreEqual(1, differences.Count);
            Assert.AreEqual("Title", differences[0].FieldInternalName);
            Assert.AreEqual("New Title", differences[0].NewValue);
            Assert.IsNull(differences[0].CurrentValue);
        }

        [TestMethod]
        public void GetDifferences_WithVariousDataTypes_HandlesCorrectly()
        {
            // Arrange
            var currentDate = DateTime.UtcNow.AddDays(-1);
            var newDate = DateTime.UtcNow;
            
            var listItem = CreateMockListItem(new Dictionary<string, object>
            {
                ["Title"] = "Old Title",
                ["NumberField"] = 100,
                ["DateField"] = currentDate,
                ["UserField"] = new FieldUserValue { LookupId = 1},
                ["MultiLookupField"] = new FieldLookupValue[] 
                { 
                    new FieldLookupValue { LookupId = 1} 
                },
                ["TaxonomyField"] = new TaxonomyFieldValue { TermGuid = "old-guid", Label = "OldTerm" },
                ["GeoLocationField"] = new FieldGeolocationValue { Latitude = 40.0, Longitude = -70.0, Altitude = 100.0, Measure = 0.0 }
            });

            var newValues = new Dictionary<string, object>
            {
                ["Title"] = "New Title",
                ["NumberField"] = 123,
                ["DateField"] = newDate,
                ["UserField"] = new FieldUserValue {LookupId = 5 },
                ["MultiLookupField"] = new FieldLookupValue[] 
                { 
                    new FieldLookupValue { LookupId = 2 }, 
                    new FieldLookupValue { LookupId = 3 } 
                },
                ["TaxonomyField"] = new TaxonomyFieldValue { TermGuid = "abcd-efgh-ijkl-mnop", Label = "Term1" },
                ["GeoLocationField"] = new FieldGeolocationValue { Latitude = 47.6097, Longitude = -122.3331, Altitude = 0, Measure = 0 }
            };

            // Act
            var differences = listItem.GetDifferences(newValues, false);

            // Assert
            Assert.AreEqual(7, differences.Count, "All fields should show differences");
            
            // Verify each field change
            var titleChange = differences.FirstOrDefault(d => d.FieldInternalName == "Title");
            Assert.IsNotNull(titleChange);
            Assert.AreEqual("New Title", titleChange.NewValue);
            Assert.AreEqual("Old Title", titleChange.CurrentValue);

            var numberChange = differences.FirstOrDefault(d => d.FieldInternalName == "NumberField");
            Assert.IsNotNull(numberChange);
            Assert.AreEqual(123, numberChange.NewValue);
            Assert.AreEqual(100, numberChange.CurrentValue);

            var dateChange = differences.FirstOrDefault(d => d.FieldInternalName == "DateField");
            Assert.IsNotNull(dateChange);
            Assert.AreEqual(newDate, dateChange.NewValue);
            Assert.AreEqual(currentDate, dateChange.CurrentValue);
        }

        [TestMethod]
        public void GetDifferences_WithIdenticalComplexTypes_ReturnsEmptyList()
        {
            // Arrange
            var userValue = new FieldUserValue {LookupId = 5 };
            var lookupValues = new FieldLookupValue[] 
            { 
                new FieldLookupValue { LookupId = 2 }, 
                new FieldLookupValue { LookupId = 3 } 
            };
            var taxonomyValue = new TaxonomyFieldValue { TermGuid = "abcd-efgh-ijkl-mnop", Label = "Term1" };
            var geoValue = new FieldGeolocationValue { Latitude = 47.6097, Longitude = -122.3331, Altitude = 0, Measure = 0 };

            var listItem = CreateMockListItem(new Dictionary<string, object>
            {
                ["UserField"] = userValue,
                ["MultiLookupField"] = lookupValues,
                ["TaxonomyField"] = taxonomyValue,
                ["GeoLocationField"] = geoValue
            });

            var newValues = new Dictionary<string, object>
            {
                ["UserField"] = new FieldUserValue {LookupId = 5 },
                ["MultiLookupField"] = new FieldLookupValue[] 
                { 
                    new FieldLookupValue { LookupId = 2 }, 
                    new FieldLookupValue { LookupId = 3 } 
                },
                ["TaxonomyField"] = new TaxonomyFieldValue { TermGuid = "abcd-efgh-ijkl-mnop", Label = "Term1" },
                ["GeoLocationField"] = new FieldGeolocationValue { Latitude = 47.6097, Longitude = -122.3331, Altitude = 0, Measure = 0 }
            };

            // Act
            var differences = listItem.GetDifferences(newValues, false);

            // Assert
            Assert.AreEqual(0, differences.Count, "Identical complex types should not show differences");
        }

        [TestMethod]
        public void GetDifferences_WithDifferentUserFields_ReturnsCorrectChanges()
        {
            // Arrange
            var listItem = CreateMockListItem(new Dictionary<string, object>
            {
                ["UserField"] = new FieldUserValue {LookupId = 1 }
            });

            var newValues = new Dictionary<string, object>
            {
                ["UserField"] = new FieldUserValue {LookupId = 5 }
            };

            // Act
            var differences = listItem.GetDifferences(newValues, false);

            // Assert
            Assert.AreEqual(1, differences.Count);
            Assert.AreEqual("UserField", differences[0].FieldInternalName);
            
            var newUserValue = differences[0].NewValue as FieldUserValue;
            var currentUserValue = differences[0].CurrentValue as FieldUserValue;
            
            Assert.IsNotNull(newUserValue);
            Assert.IsNotNull(currentUserValue);
            Assert.AreEqual(5, newUserValue.LookupId);
            Assert.AreEqual(1, currentUserValue.LookupId);
        }

        [TestMethod]
        public void GetDifferences_WithDifferentMultiLookupFields_ReturnsCorrectChanges()
        {
            // Arrange
            var listItem = CreateMockListItem(new Dictionary<string, object>
            {
                ["MultiLookupField"] = new FieldLookupValue[] 
                { 
                    new FieldLookupValue { LookupId = 1 } 
                }
            });

            var newValues = new Dictionary<string, object>
            {
                ["MultiLookupField"] = new FieldLookupValue[] 
                { 
                    new FieldLookupValue { LookupId = 2 }, 
                    new FieldLookupValue { LookupId = 3 } 
                }
            };

            // Act
            var differences = listItem.GetDifferences(newValues, false);

            // Assert
            Assert.AreEqual(1, differences.Count);
            Assert.AreEqual("MultiLookupField", differences[0].FieldInternalName);
            
            var newLookupValues = differences[0].NewValue as FieldLookupValue[];
            var currentLookupValues = differences[0].CurrentValue as FieldLookupValue[];
            
            Assert.IsNotNull(newLookupValues);
            Assert.IsNotNull(currentLookupValues);
            Assert.AreEqual(2, newLookupValues.Length);
            Assert.AreEqual(1, currentLookupValues.Length);
        }

        [TestMethod]
        public void GetDifferences_WithDifferentTaxonomyFields_ReturnsCorrectChanges()
        {
            // Arrange
            var listItem = CreateMockListItem(new Dictionary<string, object>
            {
                ["TaxonomyField"] = new TaxonomyFieldValue { TermGuid = "old-guid", Label = "OldTerm" }
            });

            var newValues = new Dictionary<string, object>
            {
                ["TaxonomyField"] = new TaxonomyFieldValue { TermGuid = "new-guid", Label = "NewTerm" }
            };

            // Act
            var differences = listItem.GetDifferences(newValues, false);

            // Assert
            Assert.AreEqual(1, differences.Count);
            Assert.AreEqual("TaxonomyField", differences[0].FieldInternalName);
            
            var newTaxValue = differences[0].NewValue as TaxonomyFieldValue;
            var currentTaxValue = differences[0].CurrentValue as TaxonomyFieldValue;
            
            Assert.IsNotNull(newTaxValue);
            Assert.IsNotNull(currentTaxValue);
            Assert.AreEqual("new-guid", newTaxValue.TermGuid);
            Assert.AreEqual("old-guid", currentTaxValue.TermGuid);
        }

        [TestMethod]
        public void GetDifferences_WithDifferentGeolocationFields_ReturnsCorrectChanges()
        {
            // Arrange
            var listItem = CreateMockListItem(new Dictionary<string, object>
            {
                ["GeoLocationField"] = new FieldGeolocationValue { Latitude = 40.0, Longitude = -70.0, Altitude = 100.0, Measure = 0.0 }
            });

            var newValues = new Dictionary<string, object>
            {
                ["GeoLocationField"] = new FieldGeolocationValue { Latitude = 47.6097, Longitude = -122.3331, Altitude = 0, Measure = 0 }
            };

            // Act
            var differences = listItem.GetDifferences(newValues, false);

            // Assert
            Assert.AreEqual(1, differences.Count);
            Assert.AreEqual("GeoLocationField", differences[0].FieldInternalName);
            
            var newGeoValue = differences[0].NewValue as FieldGeolocationValue;
            var currentGeoValue = differences[0].CurrentValue as FieldGeolocationValue;
            
            Assert.IsNotNull(newGeoValue);
            Assert.IsNotNull(currentGeoValue);
            Assert.AreEqual(47.6097, newGeoValue.Latitude, 0.0001);
            Assert.AreEqual(40.0, currentGeoValue.Latitude, 0.0001);
        }

        [TestMethod]
        public void GetDifferences_WithFieldUrlValue_ReturnsCorrectChanges()
        {
            // Arrange
            var listItem = CreateMockListItem(new Dictionary<string, object>
            {
                ["UrlField"] = new FieldUrlValue { Url = "https://old.com", Description = "Old Site" }
            });

            var newValues = new Dictionary<string, object>
            {
                ["UrlField"] = new FieldUrlValue { Url = "https://new.com", Description = "New Site" }
            };

            // Act
            var differences = listItem.GetDifferences(newValues, false);

            // Assert
            Assert.AreEqual(1, differences.Count);
            Assert.AreEqual("UrlField", differences[0].FieldInternalName);

            var newUrlValue = differences[0].NewValue as FieldUrlValue;
            var currentUrlValue = differences[0].CurrentValue as FieldUrlValue;

            Assert.IsNotNull(newUrlValue);
            Assert.IsNotNull(currentUrlValue);
            Assert.AreEqual("https://new.com", newUrlValue.Url);
            Assert.AreEqual("https://old.com", currentUrlValue.Url);
        }

        [TestMethod]
        public void GetDifferences_WithStringArrayFields_ReturnsCorrectChanges()
        {
            // Arrange
            var listItem = CreateMockListItem(new Dictionary<string, object>
            {
                ["StringArrayField"] = new string[] { "Option1", "Option2" }
            });

            var newValues = new Dictionary<string, object>
            {
                ["StringArrayField"] = new string[] { "Option3", "Option4", "Option5" }
            };

            // Act
            var differences = listItem.GetDifferences(newValues, false);

            // Assert
            Assert.AreEqual(1, differences.Count);
            Assert.AreEqual("StringArrayField", differences[0].FieldInternalName);

            var newStringArray = differences[0].NewValue as string[];
            var currentStringArray = differences[0].CurrentValue as string[];

            Assert.IsNotNull(newStringArray);
            Assert.IsNotNull(currentStringArray);
            Assert.AreEqual(3, newStringArray.Length);
            Assert.AreEqual(2, currentStringArray.Length);
        }

        private ListItem CreateMockListItem(Dictionary<string, object> fieldValues)
        {
            return new SimpleTestableListItem(fieldValues);
        }
    }

    #region ListItem Test Object

    public class SimpleTestableListItem : ListItem
    {
        public SimpleTestableListItem(Dictionary<string, object> fieldValues)
            : base(CreateMockContext(), null)
        {
            var actualFieldValues = base.FieldValues;

            foreach (var kvp in fieldValues)
            {
                actualFieldValues[kvp.Key] = kvp.Value;
            }
        }

        private static ClientRuntimeContext CreateMockContext()
        {
            return (ClientRuntimeContext)FormatterServices.GetUninitializedObject(typeof(ClientRuntimeContext));
        }
    }

    #endregion
}
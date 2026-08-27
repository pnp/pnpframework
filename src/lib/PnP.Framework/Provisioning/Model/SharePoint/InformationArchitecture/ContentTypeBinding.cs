using System;
using System.Linq;

namespace PnP.Framework.Provisioning.Model
{
    /// <summary>
    /// Domain Object for Content Type Binding in the Provisioning Template 
    /// </summary>
    public partial class ContentTypeBinding : BaseModel, IEquatable<ContentTypeBinding>
    {
        #region Private Members

        private string _contentTypeId;
        private FieldRefCollection _fieldRefs;

        #endregion

        #region Constructors

        /// <summary>
        /// Constructor for ContentTypeBinding
        /// </summary>
        public ContentTypeBinding()
        {
            this._fieldRefs = new FieldRefCollection(this.ParentTemplate);
        }

        #endregion

        #region Properties
        /// <summary>
        /// Gets or Sets the Content Type ID 
        /// </summary>
        public string ContentTypeId { get { return _contentTypeId; } set { _contentTypeId = value; } }

        /// <summary>
        /// Gets or Sets if the Content Type should be the default Content Type in the library
        /// </summary>
        public bool Default { get; set; }

        /// <summary>
        /// Declares if the Content Type should be Removed from the list or library
        /// </summary>
        public bool Remove { get; set; } = false;

        /// <summary>
        /// Declares if the Content Type should be Hidden from New button of the list or library, optional attribute.
        /// </summary>
        public bool Hidden { get; set; }

        /// <summary>
        /// The FieldRefs entries of the List Instance
        /// </summary>
        public FieldRefCollection FieldRefs
        {
            get { return this._fieldRefs; }
            private set { this._fieldRefs = value; }
        }

        /// <summary>
        /// Declares the ID of the SPFx Client Side Component to customize the Display Form of the Content Type.
        /// </summary>
        public string DisplayFormClientSideComponentId { get; set; }

        /// <summary>
        /// Defines the properties of the SPFx Client Side Component to customize the Display Form of the Content Type.
        /// </summary>
        public string DisplayFormClientSideComponentProperties { get; set; }

        /// <summary>
        /// Declares the ID of the SPFx Client Side Component to customize the New Form of the Content Type.
        /// </summary>
        public string NewFormClientSideComponentId { get; set; }

        /// <summary>
        /// Defines the properties of the SPFx Client Side Component to customize the New Form of the Content Type.
        /// </summary>
        public string NewFormClientSideComponentProperties { get; set; }

        /// <summary>
        /// Declares the ID of the SPFx Client Side Component to customize the Edit Form of the Content Type.
        /// </summary>
        public string EditFormClientSideComponentId { get; set; }

        /// <summary>
        /// Defines the properties of the SPFx Client Side Component to customize the Edit Form of the Content Type.
        /// </summary>
        public string EditFormClientSideComponentProperties { get; set; }

        #endregion

        #region Comparison code

        /// <summary>
        /// Gets the hash code
        /// </summary>
        /// <returns>Returns HashCode</returns>
        public override int GetHashCode()
        {
            return (String.Format("{0}|{1}|{2}|{3}|{4}|{5}|{6}|{7}|{8}|{9}|{10}",
                (this.ContentTypeId != null ? this.ContentTypeId.GetHashCode() : 0),
                this.Default.GetHashCode(),
                this.Remove.GetHashCode(),
                this.Hidden.GetHashCode(),
                (this.FieldRefs != null ? this.FieldRefs.Aggregate(0, (acc, next) => acc + (next != null ? next.GetHashCode() : 0)) : 0),
                (this.DisplayFormClientSideComponentId != null ? this.DisplayFormClientSideComponentId.GetHashCode() : 0),
                (this.DisplayFormClientSideComponentProperties != null ? this.DisplayFormClientSideComponentProperties.GetHashCode() : 0),
                (this.NewFormClientSideComponentId != null ? this.NewFormClientSideComponentId.GetHashCode() : 0),
                (this.NewFormClientSideComponentProperties != null ? this.NewFormClientSideComponentProperties.GetHashCode() : 0),
                (this.EditFormClientSideComponentId != null ? this.EditFormClientSideComponentId.GetHashCode() : 0),
                (this.EditFormClientSideComponentProperties != null ? this.EditFormClientSideComponentProperties.GetHashCode() : 0)
            ).GetHashCode());
        }

        /// <summary>
        /// Compares object with ContentTypeBinding
        /// </summary>
        /// <param name="obj">Object that represents ContentTypeBinding</param>
        /// <returns>true if the current object is equal to the ContentTypeBinding</returns>
        public override bool Equals(object obj)
        {
            if (!(obj is ContentTypeBinding))
            {
                return (false);
            }
            return (Equals((ContentTypeBinding)obj));
        }

        /// <summary>
        /// Compares ContentTypeBinding object based on ContentTypeId, Default and Remove properties.
        /// </summary>
        /// <param name="other">ContentTypeBinding object</param>
        /// <returns>true if the ContentTypeBinding object is equal to the current object; otherwise, false.</returns>
        public bool Equals(ContentTypeBinding other)
        {
            if (other == null)
            {
                return (false);
            }

            return (this.ContentTypeId == other.ContentTypeId &&
                this.Default == other.Default &&
                this.Remove == other.Remove &&
                this.Hidden == other.Hidden
                );
        }

        #endregion
    }
}

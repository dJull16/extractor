namespace Extractor.Core.Models
{
    /// <summary>
    /// Represents a collection of metadata grouped together
    /// </summary>
    public class GroupCollection
    {
        /// <summary>
        /// Unique identifier for the group
        /// </summary>
        public int Id { get; set; }

        /// <summary>
        /// Display name of the group
        /// </summary>
        public string Name { get; set; } = string.Empty;

        /// <summary>
        /// List of attribute metadata in this group
        /// </summary>
        public List<AttributeMetadata> AttributeMetadata { get; set; } = new();

        /// <summary>
        /// List of taxonomy metadata in this group
        /// </summary>
        public List<TaxonomyMetadata> TaxonomyMetadata { get; set; } = new();
    }
}
namespace Extractor.Core.Models
{
    /// <summary>
    /// Represents taxonomy metadata with hierarchical relationships
    /// </summary>
    public class TaxonomyMetadata
    {
        /// <summary>
        /// Unique identifier for the taxonomy
        /// </summary>
        public int Id { get; set; }

        /// <summary>
        /// Display name of the taxonomy
        /// </summary>
        public string Name { get; set; } = string.Empty;

        /// <summary>
        /// List of taxonomy node values
        /// </summary>
        public List<TaxonomyNodeValue> Values { get; set; } = new();

        /// <summary>
        /// List of relationships between taxonomy node values
        /// </summary>
        public List<TaxonomyNodeValueRelationship> Relationships { get; set; } = new();
    }
}
namespace Extractor.Core.Models
{
    /// <summary>
    /// Represents a value in a taxonomy node
    /// </summary>
    public class TaxonomyNodeValue
    {
        /// <summary>
        /// Unique identifier for the taxonomy node value
        /// </summary>
        public int Id { get; set; }

        /// <summary>
        /// Display value
        /// </summary>
        public string Value { get; set; } = string.Empty;
    }
}
namespace Extractor.Core.Models
{
    /// <summary>
    /// Represents metadata for an attribute with associated values
    /// </summary>
    public class AttributeMetadata
    {
        /// <summary>
        /// Unique identifier for the attribute
        /// </summary>
        public int Id { get; set; }

        /// <summary>
        /// Display name of the attribute
        /// </summary>
        public string Name { get; set; } = string.Empty;

        /// <summary>
        /// List of possible values for this attribute
        /// </summary>
        public List<string> Values { get; set; } = new();
    }
}
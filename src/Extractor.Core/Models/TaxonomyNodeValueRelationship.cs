namespace Extractor.Core.Models
{
    /// <summary>
    /// Represents a relationship between taxonomy node values for dependent dropdowns
    /// </summary>
    public class TaxonomyNodeValueRelationship
    {
        /// <summary>
        /// Child taxonomy node value ID (depends on parent)
        /// </summary>
        public int FirstTaxonomyNodeValueId { get; set; }

        /// <summary>
        /// Parent taxonomy node value ID (independent)
        /// </summary>
        public int SecondTaxonomyNodeValueId { get; set; }
    }
}
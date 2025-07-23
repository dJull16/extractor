namespace Extractor.Core.Models
{
    /// <summary>
    /// Represents a collection of user permissions with associated metadata
    /// </summary>
    public class UserPermissionCollection
    {
        /// <summary>
        /// Unique identifier for the user permission collection
        /// </summary>
        public int Id { get; set; }

        /// <summary>
        /// Display name or description of the permission collection
        /// </summary>
        public string Name { get; set; } = string.Empty;

        /// <summary>
        /// List of group collections containing metadata permissions
        /// </summary>
        public List<GroupCollection> GroupCollections { get; set; } = new();
    }
}
using Extractor.Core.Models;

namespace Extractor.Core.Interfaces
{
    /// <summary>
    /// Interface for Excel Template Generator Service
    /// </summary>
    public interface IExcelTemplateGeneratorService
    {
        /// <summary>
        /// Generates an Excel template based on user permissions
        /// </summary>
        /// <param name="userPermissions">User permission collection containing metadata</param>
        /// <returns>Excel file as byte array</returns>
        /// <exception cref="ArgumentNullException">Thrown when userPermissions is null</exception>
        /// <exception cref="InvalidOperationException">Thrown when there's an error generating the template</exception>
        Task<byte[]> GenerateTemplateAsync(UserPermissionCollection userPermissions);

        /// <summary>
        /// Generates an Excel template and saves it to the specified file path
        /// </summary>
        /// <param name="userPermissions">User permission collection containing metadata</param>
        /// <param name="filePath">Path where the Excel file should be saved</param>
        /// <exception cref="ArgumentNullException">Thrown when userPermissions or filePath is null</exception>
        /// <exception cref="InvalidOperationException">Thrown when there's an error generating or saving the template</exception>
        Task GenerateTemplateToFileAsync(UserPermissionCollection userPermissions, string filePath);
    }
}
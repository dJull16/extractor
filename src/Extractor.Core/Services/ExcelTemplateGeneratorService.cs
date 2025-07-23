using Extractor.Core.Interfaces;
using Extractor.Core.Models;
using OfficeOpenXml;
using OfficeOpenXml.DataValidation;

namespace Extractor.Core.Services
{
    /// <summary>
    /// Service for generating Excel templates based on user permissions
    /// </summary>
    public class ExcelTemplateGeneratorService : IExcelTemplateGeneratorService
    {
        private const string DataSheetName = "Data";
        private const string LookupSheetName = "Lookups";
        private const string FileNameHeader = "File Name";
        private const string TitleHeader = "Title";

        // Static constructor removed to avoid issues with license setting

        /// <inheritdoc />
        public async Task<byte[]> GenerateTemplateAsync(UserPermissionCollection userPermissions)
        {
            if (userPermissions == null)
                throw new ArgumentNullException(nameof(userPermissions));

            try
            {
                using var package = new ExcelPackage();
                
                // Set license context for EPPlus 7.x
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                
                // Create main data worksheet
                var dataWorksheet = package.Workbook.Worksheets.Add(DataSheetName);
                
                // Create hidden lookup worksheet for dropdown data
                var lookupWorksheet = package.Workbook.Worksheets.Add(LookupSheetName);
                lookupWorksheet.Hidden = eWorkSheetHidden.Hidden;

                // Generate headers and extract metadata
                var headers = GenerateHeaders(userPermissions);
                var attributeMetadata = ExtractAttributeMetadata(userPermissions);
                var taxonomyMetadata = ExtractTaxonomyMetadata(userPermissions);

                // Setup main worksheet
                SetupDataWorksheet(dataWorksheet, headers);
                
                // Setup lookup data
                SetupLookupWorksheet(lookupWorksheet, attributeMetadata, taxonomyMetadata);
                
                // Apply data validation (dropdowns)
                ApplyDataValidation(dataWorksheet, lookupWorksheet, headers, attributeMetadata, taxonomyMetadata);

                return await Task.FromResult(package.GetAsByteArray());
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Error generating Excel template: {ex.Message}", ex);
            }
        }

        /// <inheritdoc />
        public async Task GenerateTemplateToFileAsync(UserPermissionCollection userPermissions, string filePath)
        {
            if (userPermissions == null)
                throw new ArgumentNullException(nameof(userPermissions));
            
            if (string.IsNullOrWhiteSpace(filePath))
                throw new ArgumentNullException(nameof(filePath));

            try
            {
                var templateBytes = await GenerateTemplateAsync(userPermissions);
                await File.WriteAllBytesAsync(filePath, templateBytes);
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Error saving Excel template to file: {ex.Message}", ex);
            }
        }

        /// <summary>
        /// Generates headers for the Excel template
        /// </summary>
        private List<string> GenerateHeaders(UserPermissionCollection userPermissions)
        {
            var headers = new List<string> { FileNameHeader, TitleHeader };

            foreach (var group in userPermissions.GroupCollections)
            {
                headers.AddRange(group.AttributeMetadata.Select(a => a.Name));
                headers.AddRange(group.TaxonomyMetadata.Select(t => t.Name));
            }

            return headers;
        }

        /// <summary>
        /// Extracts all attribute metadata from user permissions
        /// </summary>
        private List<AttributeMetadata> ExtractAttributeMetadata(UserPermissionCollection userPermissions)
        {
            var attributes = new List<AttributeMetadata>();
            
            foreach (var group in userPermissions.GroupCollections)
            {
                attributes.AddRange(group.AttributeMetadata);
            }

            return attributes;
        }

        /// <summary>
        /// Extracts all taxonomy metadata from user permissions
        /// </summary>
        private List<TaxonomyMetadata> ExtractTaxonomyMetadata(UserPermissionCollection userPermissions)
        {
            var taxonomies = new List<TaxonomyMetadata>();
            
            foreach (var group in userPermissions.GroupCollections)
            {
                taxonomies.AddRange(group.TaxonomyMetadata);
            }

            return taxonomies;
        }

        /// <summary>
        /// Sets up the main data worksheet with headers
        /// </summary>
        private void SetupDataWorksheet(ExcelWorksheet worksheet, List<string> headers)
        {
            // Add headers in row 1
            for (int i = 0; i < headers.Count; i++)
            {
                worksheet.Cells[1, i + 1].Value = headers[i];
                worksheet.Cells[1, i + 1].Style.Font.Bold = true;
            }

            // Auto-fit columns
            worksheet.Cells[1, 1, 1, headers.Count].AutoFitColumns();
        }

        /// <summary>
        /// Sets up the lookup worksheet with named ranges for dropdown data
        /// </summary>
        private void SetupLookupWorksheet(ExcelWorksheet worksheet, List<AttributeMetadata> attributes, List<TaxonomyMetadata> taxonomies)
        {
            int currentRow = 1;

            // Create named ranges for attribute values
            foreach (var attribute in attributes)
            {
                if (attribute.Values.Any())
                {
                    CreateNamedRange(worksheet, attribute.Name, attribute.Values, ref currentRow);
                }
            }

            // Create named ranges for taxonomy values
            foreach (var taxonomy in taxonomies)
            {
                if (taxonomy.Values.Any())
                {
                    var values = taxonomy.Values.Select(v => v.Value).ToList();
                    CreateNamedRange(worksheet, taxonomy.Name, values, ref currentRow);

                    // Create parent-specific named ranges for dependent dropdowns
                    CreateDependentDropdownRanges(worksheet, taxonomy, ref currentRow);
                }
            }
        }

        /// <summary>
        /// Creates a named range for dropdown values
        /// </summary>
        private void CreateNamedRange(ExcelWorksheet worksheet, string name, List<string> values, ref int currentRow)
        {
            // Clean name for Excel named range (remove spaces and special characters)
            var cleanName = CleanNameForExcel(name);
            
            for (int i = 0; i < values.Count; i++)
            {
                worksheet.Cells[currentRow + i, 1].Value = values[i];
            }

            // Create named range
            var range = worksheet.Cells[currentRow, 1, currentRow + values.Count - 1, 1];
            worksheet.Workbook.Names.Add(cleanName, range);

            currentRow += values.Count + 1; // Add gap between ranges
        }

        /// <summary>
        /// Creates named ranges for dependent dropdown relationships
        /// </summary>
        private void CreateDependentDropdownRanges(ExcelWorksheet worksheet, TaxonomyMetadata taxonomy, ref int currentRow)
        {
            // Group relationships by parent (SecondTaxonomyNodeValueId)
            var parentGroups = taxonomy.Relationships.GroupBy(r => r.SecondTaxonomyNodeValueId);

            foreach (var parentGroup in parentGroups)
            {
                var parentValue = taxonomy.Values.FirstOrDefault(v => v.Id == parentGroup.Key);
                if (parentValue == null) continue;

                var childValues = new List<string>();
                foreach (var relationship in parentGroup)
                {
                    var childValue = taxonomy.Values.FirstOrDefault(v => v.Id == relationship.FirstTaxonomyNodeValueId);
                    if (childValue != null)
                    {
                        childValues.Add(childValue.Value);
                    }
                }

                if (childValues.Any())
                {
                    var parentName = CleanNameForExcel($"{taxonomy.Name}_{parentValue.Value}");
                    CreateNamedRange(worksheet, parentName, childValues, ref currentRow);
                }
            }
        }

        /// <summary>
        /// Applies data validation (dropdowns) to the data worksheet
        /// </summary>
        private void ApplyDataValidation(ExcelWorksheet dataWorksheet, ExcelWorksheet lookupWorksheet, 
            List<string> headers, List<AttributeMetadata> attributes, List<TaxonomyMetadata> taxonomies)
        {
            for (int col = 1; col <= headers.Count; col++)
            {
                var header = headers[col - 1];
                
                // Skip default headers
                if (header == FileNameHeader || header == TitleHeader)
                    continue;

                // Apply validation for attributes (simple dropdowns)
                var attribute = attributes.FirstOrDefault(a => a.Name == header);
                if (attribute != null && attribute.Values.Any())
                {
                    ApplySimpleDropdown(dataWorksheet, col, CleanNameForExcel(attribute.Name));
                    continue;
                }

                // Apply validation for taxonomies (dependent dropdowns)
                var taxonomy = taxonomies.FirstOrDefault(t => t.Name == header);
                if (taxonomy != null && taxonomy.Values.Any())
                {
                    if (taxonomy.Relationships.Any())
                    {
                        ApplyDependentDropdown(dataWorksheet, col, taxonomy, headers);
                    }
                    else
                    {
                        ApplySimpleDropdown(dataWorksheet, col, CleanNameForExcel(taxonomy.Name));
                    }
                }
            }
        }

        /// <summary>
        /// Applies simple dropdown validation to a column
        /// </summary>
        private void ApplySimpleDropdown(ExcelWorksheet worksheet, int column, string namedRange)
        {
            var validation = worksheet.DataValidations.AddListValidation($"{GetColumnLetter(column)}2:{GetColumnLetter(column)}1048576");
            validation.Formula.ExcelFormula = namedRange;
            validation.ShowErrorMessage = true;
            validation.ErrorTitle = "Invalid Selection";
            validation.Error = "Please select a value from the dropdown list.";
        }

        /// <summary>
        /// Applies dependent dropdown validation to a column
        /// </summary>
        private void ApplyDependentDropdown(ExcelWorksheet worksheet, int column, TaxonomyMetadata taxonomy, List<string> headers)
        {
            // Find parent column (this is a simplified approach - in real scenarios, you'd need more complex logic)
            // For now, we'll assume the parent is the previous taxonomy column
            var parentColumns = new List<int>();
            for (int i = 0; i < headers.Count; i++)
            {
                if (i < column - 1)
                {
                    var otherTaxonomy = ExtractTaxonomyMetadata(new UserPermissionCollection())
                        .FirstOrDefault(t => t.Name == headers[i]);
                    if (otherTaxonomy != null)
                    {
                        parentColumns.Add(i + 1);
                    }
                }
            }

            // Apply INDIRECT formula for dependent dropdown
            var validation = worksheet.DataValidations.AddListValidation($"{GetColumnLetter(column)}2:{GetColumnLetter(column)}1048576");
            
            if (parentColumns.Any())
            {
                var parentColumn = parentColumns.Last(); // Use the last parent column
                var cleanTaxonomyName = CleanNameForExcel(taxonomy.Name);
                validation.Formula.ExcelFormula = $"INDIRECT({GetColumnLetter(parentColumn)}2&\"_{cleanTaxonomyName}\")";
            }
            else
            {
                // Fallback to simple dropdown if no parent found
                validation.Formula.ExcelFormula = CleanNameForExcel(taxonomy.Name);
            }

            validation.ShowErrorMessage = true;
            validation.ErrorTitle = "Invalid Selection";
            validation.Error = "Please select a value from the dropdown list.";
        }

        /// <summary>
        /// Cleans a name to be valid for Excel named ranges
        /// </summary>
        private string CleanNameForExcel(string name)
        {
            return name.Replace(" ", "_")
                      .Replace("-", "_")
                      .Replace("(", "")
                      .Replace(")", "")
                      .Replace(".", "_");
        }

        /// <summary>
        /// Gets the Excel column letter for a given column number
        /// </summary>
        private string GetColumnLetter(int column)
        {
            string result = "";
            while (column > 0)
            {
                column--;
                result = (char)('A' + column % 26) + result;
                column /= 26;
            }
            return result;
        }
    }
}
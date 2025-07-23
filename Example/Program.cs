using Extractor.Core.Models;
using Extractor.Core.Services;
using OfficeOpenXml;

namespace Extractor.Example
{
    /// <summary>
    /// Example usage of the Excel Template Generator Service
    /// </summary>
    class Program
    {
        static async Task Main(string[] args)
        {
            // Set EPPlus license for the example
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            // Create the service
            var service = new ExcelTemplateGeneratorService();

            // Create sample user permissions with complex metadata
            var userPermissions = CreateSampleUserPermissions();

            try
            {
                Console.WriteLine("Generating Excel template...");
                
                // Generate Excel template
                var excelBytes = await service.GenerateTemplateAsync(userPermissions);
                
                // Save to file
                var outputPath = Path.Combine(Environment.CurrentDirectory, "SampleTemplate.xlsx");
                await File.WriteAllBytesAsync(outputPath, excelBytes);
                
                Console.WriteLine($"Excel template generated successfully!");
                Console.WriteLine($"File saved to: {outputPath}");
                Console.WriteLine($"File size: {excelBytes.Length} bytes");
                
                // Verify the Excel file by reading it back
                using var package = new ExcelPackage(new FileInfo(outputPath));
                var dataSheet = package.Workbook.Worksheets["Data"];
                var lookupSheet = package.Workbook.Worksheets["Lookups"];
                
                Console.WriteLine("\nWorksheet analysis:");
                Console.WriteLine($"- Data sheet: {dataSheet.Dimension?.Address ?? "Empty"}");
                Console.WriteLine($"- Lookup sheet: {(lookupSheet?.Hidden == OfficeOpenXml.eWorkSheetHidden.Hidden ? "Hidden" : "Visible")}");
                Console.WriteLine($"- Named ranges: {package.Workbook.Names.Count}");
                
                Console.WriteLine("\nHeaders in the template:");
                for (int col = 1; col <= 10; col++)
                {
                    var headerValue = dataSheet.Cells[1, col].Value?.ToString();
                    if (!string.IsNullOrEmpty(headerValue))
                    {
                        Console.WriteLine($"  Column {col}: {headerValue}");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        private static UserPermissionCollection CreateSampleUserPermissions()
        {
            return new UserPermissionCollection
            {
                Id = 1,
                Name = "Sample User Permissions",
                GroupCollections = new List<GroupCollection>
                {
                    new GroupCollection
                    {
                        Id = 1,
                        Name = "Document Properties",
                        AttributeMetadata = new List<AttributeMetadata>
                        {
                            new AttributeMetadata
                            {
                                Id = 1,
                                Name = "Department",
                                Values = new List<string> { "Human Resources", "Information Technology", "Finance", "Marketing", "Operations" }
                            },
                            new AttributeMetadata
                            {
                                Id = 2,
                                Name = "Priority",
                                Values = new List<string> { "Critical", "High", "Medium", "Low" }
                            },
                            new AttributeMetadata
                            {
                                Id = 3,
                                Name = "Document Type",
                                Values = new List<string> { "Report", "Presentation", "Spreadsheet", "Document", "Image" }
                            }
                        },
                        TaxonomyMetadata = new List<TaxonomyMetadata>
                        {
                            new TaxonomyMetadata
                            {
                                Id = 1,
                                Name = "Location",
                                Values = new List<TaxonomyNodeValue>
                                {
                                    // Parent locations
                                    new TaxonomyNodeValue { Id = 1, Value = "North America" },
                                    new TaxonomyNodeValue { Id = 2, Value = "Europe" },
                                    new TaxonomyNodeValue { Id = 3, Value = "Asia Pacific" },
                                    
                                    // Child locations - North America
                                    new TaxonomyNodeValue { Id = 4, Value = "United States" },
                                    new TaxonomyNodeValue { Id = 5, Value = "Canada" },
                                    new TaxonomyNodeValue { Id = 6, Value = "Mexico" },
                                    
                                    // Child locations - Europe
                                    new TaxonomyNodeValue { Id = 7, Value = "United Kingdom" },
                                    new TaxonomyNodeValue { Id = 8, Value = "Germany" },
                                    new TaxonomyNodeValue { Id = 9, Value = "France" },
                                    
                                    // Child locations - Asia Pacific
                                    new TaxonomyNodeValue { Id = 10, Value = "Japan" },
                                    new TaxonomyNodeValue { Id = 11, Value = "Australia" },
                                    new TaxonomyNodeValue { Id = 12, Value = "Singapore" }
                                },
                                Relationships = new List<TaxonomyNodeValueRelationship>
                                {
                                    // North America children
                                    new TaxonomyNodeValueRelationship { FirstTaxonomyNodeValueId = 4, SecondTaxonomyNodeValueId = 1 }, // US -> North America
                                    new TaxonomyNodeValueRelationship { FirstTaxonomyNodeValueId = 5, SecondTaxonomyNodeValueId = 1 }, // Canada -> North America
                                    new TaxonomyNodeValueRelationship { FirstTaxonomyNodeValueId = 6, SecondTaxonomyNodeValueId = 1 }, // Mexico -> North America
                                    
                                    // Europe children
                                    new TaxonomyNodeValueRelationship { FirstTaxonomyNodeValueId = 7, SecondTaxonomyNodeValueId = 2 }, // UK -> Europe
                                    new TaxonomyNodeValueRelationship { FirstTaxonomyNodeValueId = 8, SecondTaxonomyNodeValueId = 2 }, // Germany -> Europe
                                    new TaxonomyNodeValueRelationship { FirstTaxonomyNodeValueId = 9, SecondTaxonomyNodeValueId = 2 }, // France -> Europe
                                    
                                    // Asia Pacific children
                                    new TaxonomyNodeValueRelationship { FirstTaxonomyNodeValueId = 10, SecondTaxonomyNodeValueId = 3 }, // Japan -> Asia Pacific
                                    new TaxonomyNodeValueRelationship { FirstTaxonomyNodeValueId = 11, SecondTaxonomyNodeValueId = 3 }, // Australia -> Asia Pacific
                                    new TaxonomyNodeValueRelationship { FirstTaxonomyNodeValueId = 12, SecondTaxonomyNodeValueId = 3 }  // Singapore -> Asia Pacific
                                }
                            }
                        }
                    },
                    new GroupCollection
                    {
                        Id = 2,
                        Name = "Classification",
                        AttributeMetadata = new List<AttributeMetadata>
                        {
                            new AttributeMetadata
                            {
                                Id = 4,
                                Name = "Security Level",
                                Values = new List<string> { "Public", "Internal", "Confidential", "Restricted" }
                            }
                        },
                        TaxonomyMetadata = new List<TaxonomyMetadata>
                        {
                            new TaxonomyMetadata
                            {
                                Id = 2,
                                Name = "Project Category",
                                Values = new List<TaxonomyNodeValue>
                                {
                                    // Parent categories
                                    new TaxonomyNodeValue { Id = 13, Value = "Technology" },
                                    new TaxonomyNodeValue { Id = 14, Value = "Business" },
                                    
                                    // Child categories - Technology
                                    new TaxonomyNodeValue { Id = 15, Value = "Software Development" },
                                    new TaxonomyNodeValue { Id = 16, Value = "Infrastructure" },
                                    new TaxonomyNodeValue { Id = 17, Value = "Data Analytics" },
                                    
                                    // Child categories - Business
                                    new TaxonomyNodeValue { Id = 18, Value = "Process Improvement" },
                                    new TaxonomyNodeValue { Id = 19, Value = "Strategic Planning" },
                                    new TaxonomyNodeValue { Id = 20, Value = "Market Research" }
                                },
                                Relationships = new List<TaxonomyNodeValueRelationship>
                                {
                                    // Technology children
                                    new TaxonomyNodeValueRelationship { FirstTaxonomyNodeValueId = 15, SecondTaxonomyNodeValueId = 13 }, // Software Dev -> Technology
                                    new TaxonomyNodeValueRelationship { FirstTaxonomyNodeValueId = 16, SecondTaxonomyNodeValueId = 13 }, // Infrastructure -> Technology
                                    new TaxonomyNodeValueRelationship { FirstTaxonomyNodeValueId = 17, SecondTaxonomyNodeValueId = 13 }, // Data Analytics -> Technology
                                    
                                    // Business children
                                    new TaxonomyNodeValueRelationship { FirstTaxonomyNodeValueId = 18, SecondTaxonomyNodeValueId = 14 }, // Process Improvement -> Business
                                    new TaxonomyNodeValueRelationship { FirstTaxonomyNodeValueId = 19, SecondTaxonomyNodeValueId = 14 }, // Strategic Planning -> Business
                                    new TaxonomyNodeValueRelationship { FirstTaxonomyNodeValueId = 20, SecondTaxonomyNodeValueId = 14 }  // Market Research -> Business
                                }
                            }
                        }
                    }
                }
            };
        }
    }
}
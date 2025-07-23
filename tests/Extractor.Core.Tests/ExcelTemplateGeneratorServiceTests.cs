using Extractor.Core.Models;
using Extractor.Core.Services;
using OfficeOpenXml;
using Xunit;

namespace Extractor.Core.Tests
{
    public class ExcelTemplateGeneratorServiceTests : IDisposable
    {
        private readonly ExcelTemplateGeneratorService _service;

        public ExcelTemplateGeneratorServiceTests()
        {
            // Initialize EPPlus license for tests (EPPlus 7.x)
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            
            _service = new ExcelTemplateGeneratorService();
        }

        [Fact]
        public async Task GenerateTemplateAsync_WithNullUserPermissions_ThrowsArgumentNullException()
        {
            // Act & Assert
            await Assert.ThrowsAsync<ArgumentNullException>(() => _service.GenerateTemplateAsync(null));
        }

        [Fact]
        public async Task GenerateTemplateAsync_WithEmptyUserPermissions_ReturnsValidExcelFile()
        {
            // Arrange
            var userPermissions = new UserPermissionCollection
            {
                Id = 1,
                Name = "Test Permissions",
                GroupCollections = new List<GroupCollection>()
            };

            // Act
            var result = await _service.GenerateTemplateAsync(userPermissions);

            // Assert
            Assert.NotNull(result);
            Assert.NotEmpty(result);

            // Verify it's a valid Excel file
            using var stream = new MemoryStream(result);
            using var package = new ExcelPackage(stream);
            
            Assert.NotNull(package.Workbook);
            Assert.True(package.Workbook.Worksheets.Count >= 1);
            
            var dataSheet = package.Workbook.Worksheets["Data"];
            Assert.NotNull(dataSheet);
            
            // Should have default headers
            Assert.Equal("File Name", dataSheet.Cells[1, 1].Value?.ToString());
            Assert.Equal("Title", dataSheet.Cells[1, 2].Value?.ToString());
        }

        [Fact]
        public async Task GenerateTemplateAsync_WithAttributeMetadata_CreatesCorrectHeaders()
        {
            // Arrange
            var userPermissions = new UserPermissionCollection
            {
                Id = 1,
                Name = "Test Permissions",
                GroupCollections = new List<GroupCollection>
                {
                    new GroupCollection
                    {
                        Id = 1,
                        Name = "Group 1",
                        AttributeMetadata = new List<AttributeMetadata>
                        {
                            new AttributeMetadata
                            {
                                Id = 1,
                                Name = "Department",
                                Values = new List<string> { "HR", "IT", "Finance" }
                            },
                            new AttributeMetadata
                            {
                                Id = 2,
                                Name = "Priority",
                                Values = new List<string> { "High", "Medium", "Low" }
                            }
                        }
                    }
                }
            };

            // Act
            var result = await _service.GenerateTemplateAsync(userPermissions);

            // Assert
            using var stream = new MemoryStream(result);
            using var package = new ExcelPackage(stream);
            
            var dataSheet = package.Workbook.Worksheets["Data"];
            
            // Verify headers
            Assert.Equal("File Name", dataSheet.Cells[1, 1].Value?.ToString());
            Assert.Equal("Title", dataSheet.Cells[1, 2].Value?.ToString());
            Assert.Equal("Department", dataSheet.Cells[1, 3].Value?.ToString());
            Assert.Equal("Priority", dataSheet.Cells[1, 4].Value?.ToString());
        }

        [Fact]
        public async Task GenerateTemplateAsync_WithTaxonomyMetadata_CreatesCorrectHeaders()
        {
            // Arrange
            var userPermissions = new UserPermissionCollection
            {
                Id = 1,
                Name = "Test Permissions",
                GroupCollections = new List<GroupCollection>
                {
                    new GroupCollection
                    {
                        Id = 1,
                        Name = "Group 1",
                        TaxonomyMetadata = new List<TaxonomyMetadata>
                        {
                            new TaxonomyMetadata
                            {
                                Id = 1,
                                Name = "Location",
                                Values = new List<TaxonomyNodeValue>
                                {
                                    new TaxonomyNodeValue { Id = 1, Value = "North America" },
                                    new TaxonomyNodeValue { Id = 2, Value = "Europe" },
                                    new TaxonomyNodeValue { Id = 3, Value = "USA" },
                                    new TaxonomyNodeValue { Id = 4, Value = "Canada" },
                                    new TaxonomyNodeValue { Id = 5, Value = "UK" },
                                    new TaxonomyNodeValue { Id = 6, Value = "Germany" }
                                },
                                Relationships = new List<TaxonomyNodeValueRelationship>
                                {
                                    new TaxonomyNodeValueRelationship { FirstTaxonomyNodeValueId = 3, SecondTaxonomyNodeValueId = 1 }, // USA -> North America
                                    new TaxonomyNodeValueRelationship { FirstTaxonomyNodeValueId = 4, SecondTaxonomyNodeValueId = 1 }, // Canada -> North America
                                    new TaxonomyNodeValueRelationship { FirstTaxonomyNodeValueId = 5, SecondTaxonomyNodeValueId = 2 }, // UK -> Europe
                                    new TaxonomyNodeValueRelationship { FirstTaxonomyNodeValueId = 6, SecondTaxonomyNodeValueId = 2 }  // Germany -> Europe
                                }
                            }
                        }
                    }
                }
            };

            // Act
            var result = await _service.GenerateTemplateAsync(userPermissions);

            // Assert
            using var stream = new MemoryStream(result);
            using var package = new ExcelPackage(stream);
            
            var dataSheet = package.Workbook.Worksheets["Data"];
            
            // Verify headers
            Assert.Equal("File Name", dataSheet.Cells[1, 1].Value?.ToString());
            Assert.Equal("Title", dataSheet.Cells[1, 2].Value?.ToString());
            Assert.Equal("Location", dataSheet.Cells[1, 3].Value?.ToString());
        }

        [Fact]
        public async Task GenerateTemplateAsync_WithComplexData_CreatesLookupSheet()
        {
            // Arrange
            var userPermissions = new UserPermissionCollection
            {
                Id = 1,
                Name = "Test Permissions",
                GroupCollections = new List<GroupCollection>
                {
                    new GroupCollection
                    {
                        Id = 1,
                        Name = "Group 1",
                        AttributeMetadata = new List<AttributeMetadata>
                        {
                            new AttributeMetadata
                            {
                                Id = 1,
                                Name = "Department",
                                Values = new List<string> { "HR", "IT", "Finance" }
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
                                    new TaxonomyNodeValue { Id = 1, Value = "North America" },
                                    new TaxonomyNodeValue { Id = 2, Value = "USA" }
                                },
                                Relationships = new List<TaxonomyNodeValueRelationship>
                                {
                                    new TaxonomyNodeValueRelationship { FirstTaxonomyNodeValueId = 2, SecondTaxonomyNodeValueId = 1 }
                                }
                            }
                        }
                    }
                }
            };

            // Act
            var result = await _service.GenerateTemplateAsync(userPermissions);

            // Assert
            using var stream = new MemoryStream(result);
            using var package = new ExcelPackage(stream);
            
            // Verify lookup sheet exists and is hidden
            var lookupSheet = package.Workbook.Worksheets["Lookups"];
            Assert.NotNull(lookupSheet);
            Assert.Equal(OfficeOpenXml.eWorkSheetHidden.Hidden, lookupSheet.Hidden);
            
            // Verify named ranges exist
            Assert.Contains(package.Workbook.Names, n => n.Name == "Department");
            Assert.Contains(package.Workbook.Names, n => n.Name == "Location");
        }

        [Fact]
        public async Task GenerateTemplateToFileAsync_WithValidData_CreatesFile()
        {
            // Arrange
            var userPermissions = new UserPermissionCollection
            {
                Id = 1,
                Name = "Test Permissions",
                GroupCollections = new List<GroupCollection>()
            };
            
            var tempFile = Path.GetTempFileName();
            var excelFile = Path.ChangeExtension(tempFile, ".xlsx");

            try
            {
                // Act
                await _service.GenerateTemplateToFileAsync(userPermissions, excelFile);

                // Assert
                Assert.True(File.Exists(excelFile));
                
                var fileInfo = new FileInfo(excelFile);
                Assert.True(fileInfo.Length > 0);
                
                // Verify it's a valid Excel file
                using var package = new ExcelPackage(new FileInfo(excelFile));
                Assert.NotNull(package.Workbook);
                Assert.True(package.Workbook.Worksheets.Count >= 1);
            }
            finally
            {
                // Cleanup
                if (File.Exists(tempFile)) File.Delete(tempFile);
                if (File.Exists(excelFile)) File.Delete(excelFile);
            }
        }

        [Fact]
        public async Task GenerateTemplateToFileAsync_WithNullUserPermissions_ThrowsArgumentNullException()
        {
            // Arrange
            var tempFile = Path.GetTempFileName();

            try
            {
                // Act & Assert
                await Assert.ThrowsAsync<ArgumentNullException>(() => 
                    _service.GenerateTemplateToFileAsync(null, tempFile));
            }
            finally
            {
                // Cleanup
                if (File.Exists(tempFile)) File.Delete(tempFile);
            }
        }

        [Fact]
        public async Task GenerateTemplateToFileAsync_WithNullFilePath_ThrowsArgumentNullException()
        {
            // Arrange
            var userPermissions = new UserPermissionCollection
            {
                Id = 1,
                Name = "Test Permissions",
                GroupCollections = new List<GroupCollection>()
            };

            // Act & Assert
            await Assert.ThrowsAsync<ArgumentNullException>(() => 
                _service.GenerateTemplateToFileAsync(userPermissions, null));
        }

        [Fact]
        public async Task GenerateTemplateAsync_WithMultipleGroups_CombinesAllMetadata()
        {
            // Arrange
            var userPermissions = new UserPermissionCollection
            {
                Id = 1,
                Name = "Test Permissions",
                GroupCollections = new List<GroupCollection>
                {
                    new GroupCollection
                    {
                        Id = 1,
                        Name = "Group 1",
                        AttributeMetadata = new List<AttributeMetadata>
                        {
                            new AttributeMetadata { Id = 1, Name = "Department", Values = new List<string> { "HR", "IT" } }
                        }
                    },
                    new GroupCollection
                    {
                        Id = 2,
                        Name = "Group 2",
                        AttributeMetadata = new List<AttributeMetadata>
                        {
                            new AttributeMetadata { Id = 2, Name = "Priority", Values = new List<string> { "High", "Low" } }
                        },
                        TaxonomyMetadata = new List<TaxonomyMetadata>
                        {
                            new TaxonomyMetadata
                            {
                                Id = 1,
                                Name = "Category",
                                Values = new List<TaxonomyNodeValue>
                                {
                                    new TaxonomyNodeValue { Id = 1, Value = "Category A" },
                                    new TaxonomyNodeValue { Id = 2, Value = "Category B" }
                                }
                            }
                        }
                    }
                }
            };

            // Act
            var result = await _service.GenerateTemplateAsync(userPermissions);

            // Assert
            using var stream = new MemoryStream(result);
            using var package = new ExcelPackage(stream);
            
            var dataSheet = package.Workbook.Worksheets["Data"];
            
            // Verify all headers are present
            Assert.Equal("File Name", dataSheet.Cells[1, 1].Value?.ToString());
            Assert.Equal("Title", dataSheet.Cells[1, 2].Value?.ToString());
            Assert.Equal("Department", dataSheet.Cells[1, 3].Value?.ToString());
            Assert.Equal("Priority", dataSheet.Cells[1, 4].Value?.ToString());
            Assert.Equal("Category", dataSheet.Cells[1, 5].Value?.ToString());
        }

        public void Dispose()
        {
            // Cleanup if needed
        }
    }
}
# Excel Template Generator Service

A complete .NET implementation of an Excel Template Generator Service that creates dynamic Excel templates based on user permissions and metadata structures.

## Features

- **Dynamic Headers**: Automatically generates headers from user permissions (File Name, Title + metadata)
- **Simple Dropdowns**: Data validation for attribute metadata with predefined values
- **Dependent Dropdowns**: Hierarchical taxonomy relationships using Excel INDIRECT formulas
- **Named Ranges**: Proper Excel named ranges for dropdown data management
- **Hidden Lookup Sheet**: Contains all dropdown data and relationship mappings
- **Comprehensive Error Handling**: Robust exception handling with detailed error messages
- **Async/Await Support**: Modern async patterns for performance
- **Production Ready**: Full unit test coverage and proper logging support

## Architecture

### Core Models

- **`UserPermissionCollection`**: Root container for user metadata permissions
- **`GroupCollection`**: Groups related metadata together
- **`AttributeMetadata`**: Simple attributes with predefined values
- **`TaxonomyMetadata`**: Hierarchical metadata with parent-child relationships
- **`TaxonomyNodeValue`**: Individual values in a taxonomy
- **`TaxonomyNodeValueRelationship`**: Defines parent-child relationships

### Service Interface

```csharp
public interface IExcelTemplateGeneratorService
{
    Task<byte[]> GenerateTemplateAsync(UserPermissionCollection userPermissions);
    Task GenerateTemplateToFileAsync(UserPermissionCollection userPermissions, string filePath);
}
```

## Usage

### Basic Usage

```csharp
using Extractor.Core.Models;
using Extractor.Core.Services;
using OfficeOpenXml;

// Set EPPlus license
ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

// Create service
var service = new ExcelTemplateGeneratorService();

// Create user permissions with metadata
var userPermissions = new UserPermissionCollection
{
    Id = 1,
    Name = "User Permissions",
    GroupCollections = new List<GroupCollection>
    {
        new GroupCollection
        {
            AttributeMetadata = new List<AttributeMetadata>
            {
                new AttributeMetadata
                {
                    Name = "Department",
                    Values = new List<string> { "HR", "IT", "Finance" }
                }
            }
        }
    }
};

// Generate Excel template
var excelBytes = await service.GenerateTemplateAsync(userPermissions);
await File.WriteAllBytesAsync("template.xlsx", excelBytes);
```

### Advanced Usage with Dependent Dropdowns

```csharp
var taxonomyMetadata = new TaxonomyMetadata
{
    Name = "Location",
    Values = new List<TaxonomyNodeValue>
    {
        new TaxonomyNodeValue { Id = 1, Value = "North America" },
        new TaxonomyNodeValue { Id = 2, Value = "United States" },
        new TaxonomyNodeValue { Id = 3, Value = "Canada" }
    },
    Relationships = new List<TaxonomyNodeValueRelationship>
    {
        // US and Canada are children of North America
        new TaxonomyNodeValueRelationship { FirstTaxonomyNodeValueId = 2, SecondTaxonomyNodeValueId = 1 },
        new TaxonomyNodeValueRelationship { FirstTaxonomyNodeValueId = 3, SecondTaxonomyNodeValueId = 1 }
    }
};
```

## Excel Output Structure

### Data Sheet
- **Row 1**: Headers (File Name, Title, + metadata names)
- **Row 2+**: Data entry rows with dropdown validation
- **Column Width**: Auto-fitted for optimal display

### Lookups Sheet (Hidden)
- Contains all dropdown data in named ranges
- Parent-child relationship ranges for dependent dropdowns
- Named according to Excel naming conventions (spaces replaced with underscores)

### Data Validation
- **Simple Dropdowns**: Direct reference to named ranges
- **Dependent Dropdowns**: INDIRECT formulas referencing parent selections
- **Error Messages**: User-friendly validation error messages

## Project Structure

```
src/
├── Extractor.Core/
│   ├── Models/
│   │   ├── AttributeMetadata.cs
│   │   ├── GroupCollection.cs
│   │   ├── TaxonomyMetadata.cs
│   │   ├── TaxonomyNodeValue.cs
│   │   ├── TaxonomyNodeValueRelationship.cs
│   │   └── UserPermissionCollection.cs
│   ├── Interfaces/
│   │   └── IExcelTemplateGeneratorService.cs
│   └── Services/
│       └── ExcelTemplateGeneratorService.cs
tests/
└── Extractor.Core.Tests/
    └── ExcelTemplateGeneratorServiceTests.cs
Example/
└── Program.cs (Demo application)
```

## Dependencies

- **EPPlus 7.4.2**: Excel file generation and manipulation
- **.NET 8.0**: Target framework
- **xUnit**: Unit testing framework

## Testing

The project includes comprehensive unit tests covering:

- Null parameter validation
- Empty user permissions handling
- Simple attribute metadata processing
- Complex taxonomy metadata with relationships
- Multiple group collections
- File generation and validation
- Excel structure verification

Run tests with:

```bash
dotnet test
```

## Example Output

The service generates Excel files with:

- **8 Headers**: File Name, Title, Department, Priority, Document Type, Location, Security Level, Project Category
- **11 Named Ranges**: For all dropdown data and relationships
- **Hidden Lookup Sheet**: Contains all reference data
- **Data Validation**: Applied to all metadata columns

## Error Handling

The service provides comprehensive error handling:

- **ArgumentNullException**: For null parameters
- **InvalidOperationException**: For Excel generation errors
- **Detailed Error Messages**: With inner exception details
- **Graceful Degradation**: Continues processing when possible

## License

Uses EPPlus with NonCommercial license context for demonstration purposes.

## Running the Example

To see the service in action:

```bash
cd Example
dotnet run
```

This will generate a `SampleTemplate.xlsx` file demonstrating all features including dependent dropdowns.
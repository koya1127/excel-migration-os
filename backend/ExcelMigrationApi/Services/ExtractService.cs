using System.IO.Compression;
using System.Xml.Linq;
using ExcelMigrationApi.Models;

namespace ExcelMigrationApi.Services;

public class ExtractService
{
    public ExtractReport Extract(List<string> filePaths)
    {
        var report = new ExtractReport
        {
            GeneratedUtc = DateTime.UtcNow.ToString("o"),
            FileCount = filePaths.Count
        };

        foreach (var filePath in filePaths)
        {
            var ext = Path.GetExtension(filePath).ToLowerInvariant();
            if (ext != ".xlsm")
            {
                continue;
            }

            var fileName = Path.GetFileName(filePath);

            // Extract VBA modules via EPPlus
            try
            {
                ExtractVbaModules(filePath, fileName, report);
            }
            catch (Exception ex)
            {
                report.Modules.Add(new VbaModule
                {
                    SourceFile = fileName,
                    ModuleName = "(error)",
                    ModuleType = "Error",
                    Code = $"VBA extraction failed: {ex.Message}"
                });
            }

            // Extract form controls from VML drawings inside the ZIP
            try
            {
                ExtractFormControls(filePath, fileName, report);
            }
            catch
            {
                // Form control extraction is best-effort
            }
        }

        report.ModuleCount = report.Modules.Count;
        return report;
    }

    private void ExtractVbaModules(string filePath, string fileName, ExtractReport report)
    {
        var modules = VbaExtractor.ExtractModules(filePath);
        if (modules.Count == 0) return;

        foreach (var module in modules)
        {
            var code = module.Code ?? string.Empty;
            var codeLines = string.IsNullOrWhiteSpace(code) ? 0 : code.Split('\n').Length;

            report.Modules.Add(new VbaModule
            {
                SourceFile = fileName,
                ModuleName = module.Name,
                ModuleType = module.Type,
                CodeLines = codeLines,
                Code = code
            });
        }
    }

    private void ExtractFormControls(string filePath, string fileName, ExtractReport report)
    {
        using var zip = ZipFile.OpenRead(filePath);

        foreach (var entry in zip.Entries)
        {
            if (!entry.FullName.StartsWith("xl/drawings/vmlDrawing", StringComparison.OrdinalIgnoreCase) ||
                !entry.FullName.EndsWith(".vml", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            // Try UTF-8 first, then cp932 for Japanese files
            string xmlContent;
            try
            {
                using var stream = entry.Open();
                using var reader = new StreamReader(stream, System.Text.Encoding.UTF8);
                xmlContent = reader.ReadToEnd();
            }
            catch
            {
                try
                {
                    using var stream = entry.Open();
                    System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
                    using var reader = new StreamReader(stream, System.Text.Encoding.GetEncoding(932));
                    xmlContent = reader.ReadToEnd();
                }
                catch
                {
                    continue;
                }
            }

            ParseVmlForControls(xmlContent, fileName, report);
        }
    }

    private void ParseVmlForControls(string xmlContent, string fileName, ExtractReport report)
    {
        // VML uses mixed namespaces; parse with XDocument
        try
        {
            // Wrap in root if needed to handle namespace declarations
            var doc = XDocument.Parse(xmlContent);

            XNamespace v = "urn:schemas-microsoft-com:vml";
            XNamespace x = "urn:schemas-microsoft-com:office:excel";
            XNamespace o = "urn:schemas-microsoft-com:office:office";

            var shapes = doc.Descendants(v + "shape")
                .Concat(doc.Descendants("shape"))
                .ToList();

            foreach (var shape in shapes)
            {
                // Look for ClientData with ObjectType="Button"
                var clientData = shape.Descendants(x + "ClientData")
                    .Concat(shape.Descendants("ClientData"))
                    .FirstOrDefault();

                if (clientData == null) continue;

                var objectType = clientData.Attribute("ObjectType")?.Value ?? string.Empty;
                if (!objectType.Equals("Button", StringComparison.OrdinalIgnoreCase) &&
                    !objectType.Equals("Drop", StringComparison.OrdinalIgnoreCase) &&
                    !objectType.Equals("Checkbox", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                // Extract macro name
                var fmlaLink = clientData.Elements(x + "FmlaMacro")
                    .Concat(clientData.Elements("FmlaMacro"))
                    .FirstOrDefault()?.Value ?? string.Empty;

                // Strip [0]! prefix if present
                if (fmlaLink.Contains("!"))
                {
                    fmlaLink = fmlaLink.Substring(fmlaLink.IndexOf('!') + 1);
                }

                // Extract label from textbox
                var label = string.Empty;
                var textbox = shape.Descendants(v + "textbox")
                    .Concat(shape.Descendants("textbox"))
                    .FirstOrDefault();
                if (textbox != null)
                {
                    // Get inner text content
                    label = string.Join("", textbox.Descendants()
                        .Where(e => !e.HasElements)
                        .Select(e => e.Value.Trim())
                        .Where(t => !string.IsNullOrEmpty(t)));
                    if (string.IsNullOrEmpty(label))
                    {
                        label = textbox.Value.Trim();
                    }
                }

                report.Controls.Add(new FormControl
                {
                    SourceFile = fileName,
                    SheetName = string.Empty, // VML doesn't directly indicate sheet name
                    ControlName = shape.Attribute("id")?.Value ?? string.Empty,
                    ControlType = objectType,
                    Label = label,
                    Macro = fmlaLink
                });
            }
        }
        catch
        {
            // VML parsing can be fragile; skip on failure
        }
    }
}

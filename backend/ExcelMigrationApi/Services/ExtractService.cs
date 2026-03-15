using System.IO.Compression;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using ExcelMigrationApi.Models;

namespace ExcelMigrationApi.Services;

public class ExtractService
{
    private readonly ILogger<ExtractService> _logger;

    // VBA event → GAS trigger mapping
    private static readonly Dictionary<string, (string GasTriggerType, string GasNotes)> EventMapping = new(StringComparer.OrdinalIgnoreCase)
    {
        ["Workbook_Open"]           = ("onOpen",        "Map to function onOpen(e) — simple trigger, runs automatically"),
        ["Auto_Open"]               = ("onOpen",        "Map to function onOpen(e) — simple trigger, runs automatically"),
        ["Workbook_BeforeClose"]    = ("installable",   "Requires installable trigger: ScriptApp.newTrigger('fn').forSpreadsheet(ss).onOpen().create() — no direct beforeClose equivalent"),
        ["Workbook_BeforeSave"]     = ("installable",   "No direct GAS equivalent. Use onChange installable trigger as approximation"),
        ["Workbook_SheetChange"]    = ("onEdit",        "Map to function onEdit(e) — simple trigger. Use e.range, e.value, e.source"),
        ["Worksheet_Change"]        = ("onEdit",        "Map to function onEdit(e) — simple trigger. Use e.range, e.value, e.source. Filter by sheet name if needed"),
        ["Worksheet_SelectionChange"] = ("installable", "Requires installable onSelectionChange trigger: ScriptApp.newTrigger('fn').forSpreadsheet(ss).onSelectionChange().create()"),
        ["Worksheet_Activate"]      = ("unsupported",   "No GAS equivalent. Add TODO comment suggesting onOpen or custom menu alternative"),
        ["Worksheet_Deactivate"]    = ("unsupported",   "No GAS equivalent. Add TODO comment"),
        ["Workbook_NewSheet"]       = ("installable",   "Use onChange installable trigger with e.changeType == 'INSERT_SHEET'"),
        ["Worksheet_BeforeDoubleClick"] = ("installable", "No direct equivalent. Consider onEdit trigger or custom menu"),
        ["Worksheet_BeforeRightClick"]  = ("unsupported", "No GAS equivalent. Add TODO comment suggesting custom menu"),
        ["Worksheet_Calculate"]     = ("installable",   "Use onChange installable trigger as approximation"),
    };

    // Regex to detect VBA event handlers: Private/Public Sub EventName(...)
    private static readonly Regex EventRegex = new(
        @"(?:Private|Public)?\s*Sub\s+((?:Workbook|Worksheet|Auto)_\w+)\s*\(",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);

    public ExtractService(ILogger<ExtractService> logger)
    {
        _logger = logger;
    }

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
            var fileName = Path.GetFileName(filePath);

            if (ext == ".xls")
            {
                report.Modules.Add(new VbaModule
                {
                    SourceFile = fileName,
                    ModuleName = "(unsupported)",
                    ModuleType = "Warning",
                    Code = ".xls 形式のVBA抽出は現在サポートされていません。.xlsm に変換してから再度お試しください。",
                    CodeLines = 0
                });
                continue;
            }

            if (ext != ".xlsm")
            {
                continue;
            }

            // Build sheet metadata from ZIP (codeName→sheetName mapping, VML→sheet mapping)
            var sheetMeta = new SheetMetadata();
            try
            {
                sheetMeta = BuildSheetMetadata(filePath);
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, "Sheet metadata extraction failed for {FileName}", fileName);
            }

            // Extract VBA modules
            try
            {
                ExtractVbaModules(filePath, fileName, report, sheetMeta);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "VBA extraction failed for {FileName}", fileName);
                report.Modules.Add(new VbaModule
                {
                    SourceFile = fileName,
                    ModuleName = "(error)",
                    ModuleType = "Error",
                    Code = "VBA抽出中にエラーが発生しました"
                });
            }

            // Extract form controls from VML drawings inside the ZIP
            try
            {
                ExtractFormControls(filePath, fileName, report, sheetMeta);
            }
            catch
            {
                // Form control extraction is best-effort
            }
        }

        report.ModuleCount = report.Modules.Count;
        return report;
    }

    /// <summary>
    /// Build metadata mapping from ZIP internals:
    /// - codeName → sheet name (from workbook.xml)
    /// - VML file path → sheet name (from rels files)
    /// </summary>
    private SheetMetadata BuildSheetMetadata(string filePath)
    {
        var meta = new SheetMetadata();

        using var zip = ZipFile.OpenRead(filePath);

        // Step 1: Parse xl/workbook.xml to get rId → sheet name mapping
        // and codeName → sheet name mapping
        var workbookEntry = zip.GetEntry("xl/workbook.xml");
        if (workbookEntry == null) return meta;

        XDocument workbookDoc;
        using (var stream = workbookEntry.Open())
        {
            workbookDoc = XDocument.Load(stream);
        }

        XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        XNamespace r = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

        var sheets = workbookDoc.Descendants(ns + "sheet").ToList();
        var rIdToSheetName = new Dictionary<string, string>();

        foreach (var sheet in sheets)
        {
            var sheetName = sheet.Attribute("name")?.Value ?? string.Empty;
            var rId = sheet.Attribute(r + "id")?.Value ?? string.Empty;
            var codeName = sheet.Attribute("codeName")?.Value ?? string.Empty;

            if (!string.IsNullOrEmpty(rId))
                rIdToSheetName[rId] = sheetName;

            // codeName is the VBA module name for Document modules (e.g., "Sheet1" code name → "売上データ" sheet name)
            if (!string.IsNullOrEmpty(codeName))
                meta.CodeNameToSheetName[codeName] = sheetName;

            // Also map the sheet name itself (in case codeName equals the sheet name)
            if (!string.IsNullOrEmpty(sheetName))
                meta.CodeNameToSheetName[sheetName] = sheetName;
        }

        // Step 2: Parse xl/_rels/workbook.xml.rels to get rId → worksheet path
        var workbookRelsEntry = zip.GetEntry("xl/_rels/workbook.xml.rels");
        if (workbookRelsEntry == null) return meta;

        XDocument relsDoc;
        using (var stream = workbookRelsEntry.Open())
        {
            relsDoc = XDocument.Load(stream);
        }

        XNamespace relsNs = "http://schemas.openxmlformats.org/package/2006/relationships";
        var rIdToWorksheetPath = new Dictionary<string, string>();

        foreach (var rel in relsDoc.Descendants(relsNs + "Relationship"))
        {
            var id = rel.Attribute("Id")?.Value ?? string.Empty;
            var target = rel.Attribute("Target")?.Value ?? string.Empty;
            if (target.StartsWith("worksheets/", StringComparison.OrdinalIgnoreCase))
            {
                rIdToWorksheetPath[id] = target; // e.g. "worksheets/sheet1.xml"
            }
        }

        // Build worksheetPath → sheetName
        var worksheetPathToSheetName = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var (rId, sheetName) in rIdToSheetName)
        {
            if (rIdToWorksheetPath.TryGetValue(rId, out var wsPath))
            {
                worksheetPathToSheetName[wsPath] = sheetName;
            }
        }

        // Step 3: Parse each worksheet's rels to find VML drawing → sheet name
        foreach (var entry in zip.Entries)
        {
            // Look for xl/worksheets/_rels/sheetN.xml.rels
            if (!entry.FullName.StartsWith("xl/worksheets/_rels/", StringComparison.OrdinalIgnoreCase) ||
                !entry.FullName.EndsWith(".rels", StringComparison.OrdinalIgnoreCase))
                continue;

            // Derive the worksheet path from the rels path
            // e.g. "xl/worksheets/_rels/sheet1.xml.rels" → "worksheets/sheet1.xml"
            var relsFileName = Path.GetFileName(entry.FullName); // "sheet1.xml.rels"
            var worksheetFileName = relsFileName.Replace(".rels", ""); // "sheet1.xml"
            var worksheetPath = "worksheets/" + worksheetFileName;

            if (!worksheetPathToSheetName.TryGetValue(worksheetPath, out var sheetName))
                continue;

            XDocument sheetRelsDoc;
            using (var stream = entry.Open())
            {
                sheetRelsDoc = XDocument.Load(stream);
            }

            foreach (var rel in sheetRelsDoc.Descendants(relsNs + "Relationship"))
            {
                var target = rel.Attribute("Target")?.Value ?? string.Empty;
                // VML drawings are referenced as "../drawings/vmlDrawingN.vml"
                if (target.Contains("vmlDrawing", StringComparison.OrdinalIgnoreCase))
                {
                    // Normalize to "xl/drawings/vmlDrawingN.vml"
                    var vmlPath = target.Replace("../", "xl/");
                    meta.VmlPathToSheetName[vmlPath] = sheetName;
                }
            }
        }

        return meta;
    }

    private void ExtractVbaModules(string filePath, string fileName, ExtractReport report, SheetMetadata sheetMeta)
    {
        var modules = VbaExtractor.ExtractModules(filePath);
        if (modules.Count == 0) return;

        foreach (var module in modules)
        {
            var code = module.Code ?? string.Empty;
            var codeLines = string.IsNullOrWhiteSpace(code) ? 0 : code.Split('\n').Length;

            // Resolve sheet name for Document modules
            var sheetName = string.Empty;
            if (module.Type == "Document" && !string.IsNullOrEmpty(module.Name))
            {
                sheetMeta.CodeNameToSheetName.TryGetValue(module.Name, out sheetName);
                sheetName ??= string.Empty;
            }

            // Detect VBA events in code
            var detectedEvents = DetectEvents(code, module.Name, sheetName);

            report.Modules.Add(new VbaModule
            {
                SourceFile = fileName,
                ModuleName = module.Name,
                ModuleType = module.Type,
                CodeLines = codeLines,
                Code = code,
                SheetName = sheetName,
                DetectedEvents = detectedEvents
            });
        }
    }

    /// <summary>
    /// Scan VBA code for known event handler patterns and map them to GAS equivalents.
    /// </summary>
    private static List<VbaEvent> DetectEvents(string code, string moduleName, string sheetName)
    {
        var events = new List<VbaEvent>();
        if (string.IsNullOrWhiteSpace(code)) return events;

        var matches = EventRegex.Matches(code);
        foreach (Match match in matches)
        {
            var eventName = match.Groups[1].Value; // e.g. "Workbook_Open", "Worksheet_Change"

            // Normalize: find the matching key in our mapping
            var mappingKey = EventMapping.Keys.FirstOrDefault(k =>
                k.Equals(eventName, StringComparison.OrdinalIgnoreCase));

            string gasTriggerType;
            string gasNotes;

            if (mappingKey != null)
            {
                (gasTriggerType, gasNotes) = EventMapping[mappingKey];
            }
            else
            {
                gasTriggerType = "unsupported";
                gasNotes = $"Unknown VBA event '{eventName}'. Add TODO comment in GAS.";
            }

            events.Add(new VbaEvent
            {
                VbaEventName = eventName,
                SheetName = sheetName,
                GasTriggerType = gasTriggerType,
                GasNotes = gasNotes
            });
        }

        return events;
    }

    private void ExtractFormControls(string filePath, string fileName, ExtractReport report, SheetMetadata sheetMeta)
    {
        using var zip = ZipFile.OpenRead(filePath);

        foreach (var entry in zip.Entries)
        {
            if (!entry.FullName.StartsWith("xl/drawings/vmlDrawing", StringComparison.OrdinalIgnoreCase) ||
                !entry.FullName.EndsWith(".vml", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            // Resolve sheet name from VML path
            var vmlSheetName = string.Empty;
            sheetMeta.VmlPathToSheetName.TryGetValue(entry.FullName, out vmlSheetName);
            vmlSheetName ??= string.Empty;

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

            ParseVmlForControls(xmlContent, fileName, vmlSheetName, report);
        }
    }

    private void ParseVmlForControls(string xmlContent, string fileName, string sheetName, ExtractReport report)
    {
        try
        {
            var doc = XDocument.Parse(xmlContent);

            XNamespace v = "urn:schemas-microsoft-com:vml";
            XNamespace x = "urn:schemas-microsoft-com:office:excel";

            var shapes = doc.Descendants(v + "shape")
                .Concat(doc.Descendants("shape"))
                .ToList();

            foreach (var shape in shapes)
            {
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

                var fmlaLink = clientData.Elements(x + "FmlaMacro")
                    .Concat(clientData.Elements("FmlaMacro"))
                    .FirstOrDefault()?.Value ?? string.Empty;

                if (fmlaLink.Contains("!"))
                {
                    fmlaLink = fmlaLink.Substring(fmlaLink.IndexOf('!') + 1);
                }

                var label = string.Empty;
                var textbox = shape.Descendants(v + "textbox")
                    .Concat(shape.Descendants("textbox"))
                    .FirstOrDefault();
                if (textbox != null)
                {
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
                    SheetName = sheetName,
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

    /// <summary>
    /// Internal metadata resolved from ZIP structure.
    /// </summary>
    private class SheetMetadata
    {
        /// <summary>VBA codeName (e.g. "Sheet1") → worksheet display name (e.g. "売上データ")</summary>
        public Dictionary<string, string> CodeNameToSheetName { get; } = new(StringComparer.OrdinalIgnoreCase);

        /// <summary>VML file path (e.g. "xl/drawings/vmlDrawing1.vml") → sheet name</summary>
        public Dictionary<string, string> VmlPathToSheetName { get; } = new(StringComparer.OrdinalIgnoreCase);
    }
}

using System.IO.Compression;
using System.Text;
using OpenMcdf;

namespace ExcelMigrationApi.Services;

/// <summary>
/// Extracts VBA module source code from .xlsm files without EPPlus.
/// Uses ZIP to access vbaProject.bin, OpenMcdf for OLE2 parsing,
/// dir stream for MODULEOFFSET/CODEPAGE, and MS-OVBA decompression.
/// </summary>
public static class VbaExtractor
{
    public class VbaModuleInfo
    {
        public string Name { get; set; } = string.Empty;
        public string Code { get; set; } = string.Empty;
        public string Type { get; set; } = "Standard";
    }

    /// <summary>
    /// Extract all VBA modules from an .xlsm file.
    /// </summary>
    public static List<VbaModuleInfo> ExtractModules(string xlsmPath)
    {
        var modules = new List<VbaModuleInfo>();

        byte[]? vbaProjectBytes = ExtractVbaProjectBin(xlsmPath);
        if (vbaProjectBytes == null) return modules;

        using var cf = new CompoundFile(new MemoryStream(vbaProjectBytes));

        // Get module types from PROJECT text stream
        if (!cf.RootStorage.TryGetStream("PROJECT", out var projectStream) || projectStream == null)
            return modules;

        var projectText = Encoding.UTF8.GetString(projectStream.GetData());
        var typeMap = ParseProjectStreamTypes(projectText);

        // Get VBA storage
        if (!cf.RootStorage.TryGetStorage("VBA", out var vbaStorage) || vbaStorage == null)
            return modules;

        // Parse dir stream for module stream names, offsets, and codepage
        if (!vbaStorage.TryGetStream("dir", out var dirStream) || dirStream == null)
            return modules;

        var dirData = Decompress(dirStream.GetData());
        var dirInfo = ParseDirStream(dirData);

        // Read each module's source
        foreach (var entry in dirInfo.Modules)
        {
            try
            {
                if (!vbaStorage.TryGetStream(entry.StreamName, out var moduleStream) || moduleStream == null)
                    continue;

                var rawData = moduleStream.GetData();
                if (entry.TextOffset >= rawData.Length) continue;

                // Decompress source from TextOffset
                var compressedLen = rawData.Length - entry.TextOffset;
                var compressed = new byte[compressedLen];
                Array.Copy(rawData, entry.TextOffset, compressed, 0, compressedLen);
                var decompressedBytes = Decompress(compressed);

                // Use codepage from dir stream for correct Japanese decoding
                var encoding = GetEncodingFromCodepage(dirInfo.CodePage);
                var sourceCode = encoding.GetString(decompressedBytes);

                // Determine type from PROJECT stream, fallback to dir stream
                var type = typeMap.TryGetValue(entry.Name, out var t) ? t : entry.Type;

                modules.Add(new VbaModuleInfo
                {
                    Name = entry.Name,
                    Code = sourceCode,
                    Type = type
                });
            }
            catch
            {
                // Skip modules that fail to decompress
            }
        }

        return modules;
    }

    /// <summary>
    /// Count VBA modules without fully extracting source code.
    /// </summary>
    public static int CountModules(string xlsmPath)
    {
        try
        {
            byte[]? vbaProjectBytes = ExtractVbaProjectBin(xlsmPath);
            if (vbaProjectBytes == null) return 0;

            using var cf = new CompoundFile(new MemoryStream(vbaProjectBytes));
            if (!cf.RootStorage.TryGetStorage("VBA", out var vbaStorage) || vbaStorage == null)
                return 0;
            if (!vbaStorage.TryGetStream("dir", out var dirStream) || dirStream == null)
                return 0;

            var dirData = Decompress(dirStream.GetData());
            return ParseDirStream(dirData).Modules.Count;
        }
        catch
        {
            return 0;
        }
    }

    private static byte[]? ExtractVbaProjectBin(string xlsmPath)
    {
        using var zip = ZipFile.OpenRead(xlsmPath);
        var entry = zip.GetEntry("xl/vbaProject.bin");
        if (entry == null) return null;

        using var stream = entry.Open();
        using var ms = new MemoryStream();
        stream.CopyTo(ms);
        return ms.ToArray();
    }

    private static Encoding GetEncodingFromCodepage(int codepage)
    {
        if (codepage <= 0) codepage = 932; // Default to Shift_JIS for Japanese
        try
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            return Encoding.GetEncoding(codepage);
        }
        catch
        {
            return Encoding.UTF8;
        }
    }

    #region PROJECT stream parsing (types only)

    /// <summary>
    /// Parse the PROJECT text stream to build a module name -> type map.
    /// </summary>
    private static Dictionary<string, string> ParseProjectStreamTypes(string projectText)
    {
        var types = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        foreach (var line in projectText.Split('\n'))
        {
            var trimmed = line.Trim('\r', ' ');
            if (string.IsNullOrEmpty(trimmed)) continue;

            var eqIndex = trimmed.IndexOf('=');
            if (eqIndex <= 0) continue;

            var key = trimmed[..eqIndex];
            var value = trimmed[(eqIndex + 1)..];

            switch (key)
            {
                case "Module":
                    types[value] = "Standard";
                    break;
                case "Class":
                    types[value] = "Class";
                    break;
                case "BaseClass":
                    types[value] = "Form";
                    break;
                case "Document":
                    var slashIndex = value.IndexOf('/');
                    var name = slashIndex > 0 ? value[..slashIndex] : value;
                    types[name] = "Document";
                    break;
            }
        }

        return types;
    }

    #endregion

    #region dir stream parsing

    private class DirInfo
    {
        public int CodePage { get; set; }
        public List<ModuleEntry> Modules { get; set; } = new();
    }

    private class ModuleEntry
    {
        public string Name { get; set; } = string.Empty;
        public string StreamName { get; set; } = string.Empty;
        public string Type { get; set; } = "Standard";
        public int TextOffset { get; set; }
    }

    /// <summary>
    /// Parse the decompressed dir stream according to [MS-OVBA] 2.3.4.2.
    /// Extracts PROJECTCODEPAGE, and for each module: name, stream name, type, and text offset.
    /// </summary>
    private static DirInfo ParseDirStream(byte[] data)
    {
        var info = new DirInfo();
        int pos = 0;
        ModuleEntry? current = null;

        while (pos + 6 <= data.Length)
        {
            var recordId = BitConverter.ToUInt16(data, pos);
            var recordSize = BitConverter.ToInt32(data, pos + 2);
            pos += 6;

            if (recordSize < 0 || pos + recordSize > data.Length)
                break;

            switch (recordId)
            {
                case 0x0003: // PROJECTCODEPAGE
                    if (recordSize >= 2)
                        info.CodePage = BitConverter.ToUInt16(data, pos);
                    break;

                case 0x0009: // PROJECTVERSION — 4 bytes MajorVersion, then 2 extra bytes MinorVersion
                    pos += recordSize;
                    // MinorVersion: 2 bytes not counted in recordSize
                    if (pos + 2 <= data.Length) pos += 2;
                    continue; // skip normal pos += recordSize below

                case 0x0019: // MODULENAME
                    current = new ModuleEntry();
                    if (recordSize > 0)
                    {
                        var nameEncoding = GetEncodingFromCodepage(info.CodePage);
                        current.Name = nameEncoding.GetString(data, pos, recordSize).TrimEnd('\0');
                        current.StreamName = current.Name;
                    }
                    break;

                case 0x0047: // MODULENAMEUNICODE
                    if (current != null && recordSize >= 2)
                    {
                        var unicodeName = Encoding.Unicode.GetString(data, pos, recordSize).TrimEnd('\0');
                        if (!string.IsNullOrEmpty(unicodeName))
                            current.Name = unicodeName;
                        // StreamName stays as the MBCS name from 0x0019 (used for OLE2 stream lookup)
                    }
                    break;

                case 0x001A: // MODULESTREAMNAME (MBCS)
                    if (current != null && recordSize > 0)
                    {
                        var nameEncoding = GetEncodingFromCodepage(info.CodePage);
                        current.StreamName = nameEncoding.GetString(data, pos, recordSize).TrimEnd('\0');
                    }
                    break;

                case 0x0032: // MODULESTREAMNAMEUNICODE — skip, keep MBCS stream name for OLE2
                    break;

                case 0x0031: // MODULEOFFSET
                    if (current != null && recordSize >= 4)
                        current.TextOffset = BitConverter.ToInt32(data, pos);
                    break;

                case 0x0021: // MODULETYPE procedural
                    if (current != null)
                        current.Type = "Standard";
                    break;

                case 0x0022: // MODULETYPE class/document
                    if (current != null)
                        current.Type = "Class";
                    break;

                case 0x002B: // MODULEEND
                    if (current != null)
                    {
                        info.Modules.Add(current);
                        current = null;
                    }
                    break;
            }

            pos += recordSize;
        }

        return info;
    }

    #endregion

    #region MS-OVBA Decompression (MS-OVBA 2.4.1)

    private static byte[] Decompress(byte[] data)
    {
        if (data.Length == 0) return data;

        int pos = 0;
        if (data[pos++] != 0x01) return Array.Empty<byte>();

        var output = new List<byte>(data.Length * 2);

        while (pos < data.Length)
        {
            if (pos + 2 > data.Length) break;

            var header = (ushort)(data[pos] | (data[pos + 1] << 8));
            pos += 2;

            var chunkSize = (header & 0x0FFF) + 3;
            var isCompressed = (header & 0x8000) != 0;

            var chunkEnd = pos + chunkSize - 2;
            if (chunkEnd > data.Length) chunkEnd = data.Length;

            if (!isCompressed)
            {
                while (pos < chunkEnd)
                    output.Add(data[pos++]);
                continue;
            }

            var chunkStart = output.Count;

            while (pos < chunkEnd)
            {
                if (pos >= data.Length) break;
                var flagByte = data[pos++];

                for (int i = 0; i < 8 && pos < chunkEnd; i++)
                {
                    if ((flagByte & (1 << i)) == 0)
                    {
                        if (pos >= data.Length) return output.ToArray();
                        output.Add(data[pos++]);
                    }
                    else
                    {
                        if (pos + 2 > data.Length) return output.ToArray();
                        var token = (ushort)(data[pos] | (data[pos + 1] << 8));
                        pos += 2;

                        var decompressedCurrent = output.Count - chunkStart;
                        int bitCount = GetBitCount(decompressedCurrent);
                        int lengthMask = (ushort)(0xFFFF >> bitCount);

                        var length = (token & lengthMask) + 3;
                        var offset = (token >> (16 - bitCount)) + 1;

                        var copyFrom = output.Count - offset;
                        if (copyFrom < 0) copyFrom = 0;

                        for (int j = 0; j < length; j++)
                        {
                            var idx = copyFrom + j;
                            output.Add(idx < output.Count ? output[idx] : (byte)0);
                        }
                    }
                }
            }
        }

        return output.ToArray();
    }

    private static int GetBitCount(int decompressedCurrent)
    {
        if (decompressedCurrent <= 1) return 4;

        int bits = 4;
        int threshold = 1 << 4;
        while (threshold < decompressedCurrent && bits < 12)
        {
            threshold <<= 1;
            bits++;
        }
        return bits;
    }

    #endregion
}

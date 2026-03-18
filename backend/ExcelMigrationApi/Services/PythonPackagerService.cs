using System.IO.Compression;
using System.Text;
using ExcelMigrationApi.Models;

namespace ExcelMigrationApi.Services;

public class PythonPackagerService
{
    private const string PythonEmbedPath = "/app/python-embed";

    public PythonPackage Package(string originalFileName, List<PythonConvertResult> results, string? spreadsheetId)
    {
        var baseName = Path.GetFileNameWithoutExtension(originalFileName);
        var safeName = MakeSafeName(baseName);

        using var memoryStream = new MemoryStream();
        using (var archive = new ZipArchive(memoryStream, ZipArchiveMode.Create, true))
        {
            var successResults = results.Where(r => r.Status == "success").ToList();

            // modules/*.py
            foreach (var result in successResults)
            {
                var moduleName = MakeSafeName(result.ModuleName).ToLowerInvariant();
                var entry = archive.CreateEntry($"{safeName}_local/modules/{moduleName}.py");
                using var writer = new StreamWriter(entry.Open(), Encoding.UTF8);
                writer.Write(result.PythonCode);
            }

            // modules/__init__.py
            var initEntry = archive.CreateEntry($"{safeName}_local/modules/__init__.py");
            using (var writer = new StreamWriter(initEntry.Open(), Encoding.UTF8))
            {
                writer.WriteLine("# Auto-generated module init");
            }

            // main.py
            var mainContent = BuildMainPy(baseName, successResults);
            var mainEntry = archive.CreateEntry($"{safeName}_local/main.py");
            using (var writer = new StreamWriter(mainEntry.Open(), Encoding.UTF8))
            {
                writer.Write(mainContent);
            }

            // config.json
            var configContent = BuildConfig(spreadsheetId);
            var configEntry = archive.CreateEntry($"{safeName}_local/config.json");
            using (var writer = new StreamWriter(configEntry.Open(), Encoding.UTF8))
            {
                writer.Write(configContent);
            }

            // python-embed/ (bundled Python runtime for Windows)
            if (Directory.Exists(PythonEmbedPath))
            {
                AddDirectoryToZip(archive, PythonEmbedPath, $"{safeName}_local/python-embed");
            }

            // run.bat (uses embedded Python — no install required)
            var runEntry = archive.CreateEntry($"{safeName}_local/run.bat");
            using (var writer = new StreamWriter(runEntry.Open(), Encoding.UTF8))
            {
                writer.WriteLine("@echo off");
                writer.WriteLine("chcp 65001 >nul 2>&1");
                writer.WriteLine("\"%~dp0python-embed\\python.exe\" \"%~dp0main.py\"");
                writer.WriteLine("pause");
            }

            // README.txt
            var readmeContent = BuildReadme(baseName, successResults, spreadsheetId);
            var readmeEntry = archive.CreateEntry($"{safeName}_local/README.txt");
            using (var writer = new StreamWriter(readmeEntry.Open(), Encoding.UTF8))
            {
                writer.Write(readmeContent);
            }
        }

        return new PythonPackage
        {
            FileName = $"{safeName}_local.zip",
            ZipData = memoryStream.ToArray(),
            ModuleNames = results.Where(r => r.Status == "success").Select(r => r.ModuleName).ToList(),
            ReadmeContent = BuildReadme(baseName, results.Where(r => r.Status == "success").ToList(), spreadsheetId)
        };
    }

    private static void AddDirectoryToZip(ZipArchive archive, string sourceDir, string zipPrefix)
    {
        foreach (var filePath in Directory.GetFiles(sourceDir, "*", SearchOption.AllDirectories))
        {
            var relativePath = Path.GetRelativePath(sourceDir, filePath).Replace('\\', '/');
            var entry = archive.CreateEntry($"{zipPrefix}/{relativePath}", CompressionLevel.Optimal);
            using var entryStream = entry.Open();
            using var fileStream = File.OpenRead(filePath);
            fileStream.CopyTo(entryStream);
        }
    }

    private static string BuildMainPy(string baseName, List<PythonConvertResult> results)
    {
        var sb = new StringBuilder();
        sb.AppendLine("#!/usr/bin/env python3");
        sb.AppendLine("# -*- coding: utf-8 -*-");
        sb.AppendLine($"\"\"\"");
        sb.AppendLine($"{baseName} ローカル版");
        sb.AppendLine($"Excel Migration OS により自動生成");
        sb.AppendLine($"\"\"\"");
        sb.AppendLine();
        sb.AppendLine("import sys");
        sb.AppendLine("import os");
        sb.AppendLine("import json");
        sb.AppendLine();

        // Import modules
        foreach (var result in results)
        {
            var moduleName = MakeSafeName(result.ModuleName).ToLowerInvariant();
            sb.AppendLine($"from modules import {moduleName}");
        }

        sb.AppendLine();
        sb.AppendLine();

        // Extract callable functions with full signatures and docstrings
        var menuItems = new List<(string ModuleName, string FuncName, string DisplayName, List<FuncParam> Params)>();
        var funcRegex = new System.Text.RegularExpressions.Regex(
            @"^def\s+(\w+)\s*\(([^)]*)\).*?:\s*\n(\s*""""""(.*?)"""""")?",
            System.Text.RegularExpressions.RegexOptions.Multiline | System.Text.RegularExpressions.RegexOptions.Singleline);

        var skipFuncs = new HashSet<string> { "get_spreadsheet", "get_path", "main" };

        foreach (var result in results)
        {
            if (string.IsNullOrEmpty(result.PythonCode)) continue;
            var moduleSafe = MakeSafeName(result.ModuleName).ToLowerInvariant();

            foreach (System.Text.RegularExpressions.Match m in funcRegex.Matches(result.PythonCode))
            {
                var funcName = m.Groups[1].Value;
                if (funcName.StartsWith("_") || skipFuncs.Contains(funcName))
                    continue;

                // Parse parameters
                var paramStr = m.Groups[2].Value.Trim();
                var funcParams = ParseFuncParams(paramStr);

                // Extract docstring for display name
                var docstring = m.Groups[4].Success ? m.Groups[4].Value.Trim().Split('\n')[0].Trim() : "";
                var display = !string.IsNullOrEmpty(docstring) ? docstring : funcName.Replace("_", " ");

                menuItems.Add((moduleSafe, funcName, display, funcParams));
            }
        }

        // Note: gspread/workbook helpers removed — Python local version runs 100% offline

        sb.AppendLine("def main():");
        sb.AppendLine("    print(\"=\" * 50)");
        sb.AppendLine($"    print(\"  {baseName} ローカル版\")");
        sb.AppendLine("    print(\"=\" * 50)");
        sb.AppendLine("    print()");
        sb.AppendLine("    print(\"実行したい機能の番号を入力してください:\")");
        sb.AppendLine("    print()");

        if (menuItems.Count > 0)
        {
            for (var i = 0; i < menuItems.Count; i++)
            {
                sb.AppendLine($"    print(\"  {i + 1}. {menuItems[i].DisplayName}\")");
            }
        }
        else
        {
            for (var i = 0; i < results.Count; i++)
            {
                sb.AppendLine($"    print(\"  {i + 1}. {results[i].ModuleName}\")");
            }
        }
        sb.AppendLine("    print(\"  0. 終了\")");
        sb.AppendLine("    print()");
        sb.AppendLine();
        sb.AppendLine("    while True:");
        sb.AppendLine("        choice = input(\"番号を入力: \").strip()");
        sb.AppendLine();
        sb.AppendLine("        if choice == \"0\":");
        sb.AppendLine("            print(\"終了します。\")");
        sb.AppendLine("            break");

        if (menuItems.Count > 0)
        {
            for (var i = 0; i < menuItems.Count; i++)
            {
                var (modName, funcName, displayName, funcParams) = menuItems[i];
                sb.AppendLine($"        elif choice == \"{i + 1}\":");
                sb.AppendLine($"            try:");
                sb.AppendLine($"                print(\"実行中: {displayName}...\")");

                // Build argument list
                var args = new List<string>();
                foreach (var p in funcParams)
                {
                    if (IsWorkbookParam(p))
                    {
                        args.Add("get_workbook()");
                    }
                    else if (p.TypeHint.Contains("Worksheet"))
                    {
                        // Worksheet params: prompt for sheet name and get from workbook
                        sb.AppendLine($"                _sheet_name = input(\"{p.Name}のシート名を入力: \").strip()");
                        args.Add("get_workbook().worksheet(_sheet_name)");
                    }
                    else if (p.TypeHint.Contains("int") || p.Name.Contains("row") || p.Name.Contains("col") || p.Name.Contains("pos"))
                    {
                        sb.AppendLine($"                _{p.Name} = int(input(\"{ParamToJapanese(p.Name)}を入力: \").strip())");
                        args.Add($"_{p.Name}");
                    }
                    else
                    {
                        // Default: string input
                        sb.AppendLine($"                _{p.Name} = input(\"{ParamToJapanese(p.Name)}を入力: \").strip()");
                        args.Add($"_{p.Name}");
                    }
                }

                sb.AppendLine($"                {modName}.{funcName}({string.Join(", ", args)})");
                sb.AppendLine($"            except Exception as e:");
                sb.AppendLine($"                print(f\"エラーが発生しました: {{e}}\")");
            }
        }
        else
        {
            for (var i = 0; i < results.Count; i++)
            {
                var moduleName = MakeSafeName(results[i].ModuleName).ToLowerInvariant();
                sb.AppendLine($"        elif choice == \"{i + 1}\":");
                sb.AppendLine($"            try:");
                sb.AppendLine($"                print(\"実行中: {results[i].ModuleName}...\")");
                sb.AppendLine($"                print(\"※ modules/{moduleName}.py の関数を確認してください\")");
                sb.AppendLine($"            except Exception as e:");
                sb.AppendLine($"                print(f\"エラーが発生しました: {{e}}\")");
            }
        }

        sb.AppendLine("        else:");
        sb.AppendLine("            print(\"無効な番号です。もう一度入力してください。\")");
        sb.AppendLine("        print()");
        sb.AppendLine();
        sb.AppendLine();
        sb.AppendLine("if __name__ == \"__main__\":");
        sb.AppendLine("    main()");

        return sb.ToString();
    }

    private record FuncParam(string Name, string TypeHint);

    private static List<FuncParam> ParseFuncParams(string paramStr)
    {
        var result = new List<FuncParam>();
        if (string.IsNullOrWhiteSpace(paramStr)) return result;

        foreach (var raw in paramStr.Split(','))
        {
            var param = raw.Trim();
            if (string.IsNullOrEmpty(param) || param == "self") continue;

            // Handle default values: skip params with defaults (they're optional)
            if (param.Contains('=')) continue;

            // Parse "name: type" or just "name"
            var parts = param.Split(':');
            var name = parts[0].Trim();
            var typeHint = parts.Length > 1 ? parts[1].Trim() : "";
            result.Add(new FuncParam(name, typeHint));
        }
        return result;
    }

    private static bool IsWorkbookParam(FuncParam p)
    {
        return p.Name == "workbook" ||
               p.TypeHint.Contains("Spreadsheet") ||
               (p.Name == "wb" && p.TypeHint.Contains("gspread"));
    }

    private static string ParamToJapanese(string paramName)
    {
        return paramName switch
        {
            "caller" => "呼び出し元 (例: search, buzaibuhin, hinmei)",
            "code" => "検索コード",
            "sheet_name" => "シート名",
            "csv_file" => "CSVファイル名",
            "key_type" => "種別 (buzai/buhin)",
            "search_code" => "検索コード",
            "str_key" => "検索キー",
            "str_ext" => "拡張子",
            "index_path" => "インデックスパス",
            "index_sheet_name" => "インデックスシート名",
            "sub_addr" => "サブアドレス",
            "infile" => "ファイル名",
            "conv_type" => "変換タイプ (0: エスケープ, 1: 復元)",
            "stat" => "ステータス",
            "save_dest" => "保存先パス",
            _ => paramName
        };
    }

    private static string BuildConfig(string? spreadsheetId)
    {
        var sb = new StringBuilder();
        sb.AppendLine("{");
        sb.AppendLine("  \"data_dir\": \"data\"");
        sb.AppendLine("}");
        return sb.ToString();
    }

    private static string BuildReadme(string baseName, List<PythonConvertResult> results, string? spreadsheetId)
    {
        var sb = new StringBuilder();
        sb.AppendLine($"========================================");
        sb.AppendLine($"  {baseName} ローカル版");
        sb.AppendLine($"  使い方ガイド");
        sb.AppendLine($"========================================");
        sb.AppendLine();
        sb.AppendLine("【はじめに】");
        sb.AppendLine("このプログラムは、Excelマクロの機能をパソコン上で動かすためのものです。");
        sb.AppendLine("スプレッドシート（Google Sheets）では動かせなかった機能がここに含まれています。");
        sb.AppendLine();
        sb.AppendLine("【使い方】");
        sb.AppendLine("1. このフォルダを好きな場所に置いてください");
        sb.AppendLine("2. 「run.bat」をダブルクリックしてください");
        sb.AppendLine("3. メニューが表示されるので、番号を入力して機能を選んでください");
        sb.AppendLine();
        sb.AppendLine("※ Pythonのインストールは不要です（同梱済み）");
        sb.AppendLine();

        // No Google Sheets connection — this runs 100% locally

        sb.AppendLine("【含まれている機能】");
        foreach (var result in results)
        {
            sb.AppendLine($"  - {result.ModuleName}");
        }
        sb.AppendLine();
        sb.AppendLine("【困ったときは】");
        sb.AppendLine("・「run.bat」でエラーが出る → フォルダ内の「python-embed」フォルダが存在するか確認してください");
        sb.AppendLine("・機能が正しく動かない → 元のExcelファイルで同じ操作を確認してみてください");
        sb.AppendLine();
        sb.AppendLine("========================================");
        sb.AppendLine("  Excel Migration OS により自動生成");
        sb.AppendLine($"  生成日: {DateTime.Now:yyyy/MM/dd}");
        sb.AppendLine("========================================");

        return sb.ToString();
    }

    private static string MakeSafeName(string name)
    {
        // Replace characters that are unsafe for filenames/Python identifiers
        var safe = name.Replace(" ", "_").Replace("-", "_").Replace(".", "_");
        // Remove any remaining non-alphanumeric/underscore characters
        safe = System.Text.RegularExpressions.Regex.Replace(safe, @"[^\w]", "_");
        // Ensure doesn't start with digit
        if (safe.Length > 0 && char.IsDigit(safe[0]))
            safe = "_" + safe;
        return safe;
    }
}

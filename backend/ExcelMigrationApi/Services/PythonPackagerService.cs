using System.IO.Compression;
using System.Text;
using ExcelMigrationApi.Models;

namespace ExcelMigrationApi.Services;

public class PythonPackagerService
{
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

            // requirements.txt
            var reqEntry = archive.CreateEntry($"{safeName}_local/requirements.txt");
            using (var writer = new StreamWriter(reqEntry.Open(), Encoding.UTF8))
            {
                writer.WriteLine("gspread>=5.0.0");
                writer.WriteLine("google-auth>=2.0.0");
                writer.WriteLine("openpyxl>=3.0.0");
            }

            // setup.bat
            var setupEntry = archive.CreateEntry($"{safeName}_local/setup.bat");
            using (var writer = new StreamWriter(setupEntry.Open(), Encoding.UTF8))
            {
                writer.WriteLine("@echo off");
                writer.WriteLine("echo セットアップを開始します...");
                writer.WriteLine("pip install -r requirements.txt");
                writer.WriteLine("echo.");
                writer.WriteLine("echo セットアップが完了しました。");
                writer.WriteLine("echo 「run.bat」をダブルクリックして実行してください。");
                writer.WriteLine("pause");
            }

            // run.bat
            var runEntry = archive.CreateEntry($"{safeName}_local/run.bat");
            using (var writer = new StreamWriter(runEntry.Open(), Encoding.UTF8))
            {
                writer.WriteLine("@echo off");
                writer.WriteLine("python main.py");
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

        // Check if any function needs workbook → add helper
        var needsWorkbook = menuItems.Any(mi => mi.Params.Any(p => IsWorkbookParam(p)));
        if (needsWorkbook)
        {
            sb.AppendLine("_workbook_cache = None");
            sb.AppendLine();
            sb.AppendLine("def get_workbook():");
            sb.AppendLine("    global _workbook_cache");
            sb.AppendLine("    if _workbook_cache is not None:");
            sb.AppendLine("        return _workbook_cache");
            sb.AppendLine("    # 各モジュールのget_spreadsheet()を探す");
            sb.AppendLine("    for mod_name in dir():");
            sb.AppendLine("        pass");
            // Directly call the first module's get_spreadsheet
            var firstModWithSpreadsheet = results.FirstOrDefault(r =>
                !string.IsNullOrEmpty(r.PythonCode) && r.PythonCode.Contains("def get_spreadsheet("));
            if (firstModWithSpreadsheet != null)
            {
                var modName = MakeSafeName(firstModWithSpreadsheet.ModuleName).ToLowerInvariant();
                sb.AppendLine($"    _workbook_cache = {modName}.get_spreadsheet()");
            }
            else
            {
                sb.AppendLine("    import gspread");
                sb.AppendLine("    from google.oauth2.service_account import Credentials");
                sb.AppendLine("    with open('config.json', 'r') as f:");
                sb.AppendLine("        config = json.load(f)");
                sb.AppendLine("    scopes = ['https://www.googleapis.com/auth/spreadsheets']");
                sb.AppendLine("    creds = Credentials.from_service_account_file(config['credentials_file'], scopes=scopes)");
                sb.AppendLine("    gc = gspread.authorize(creds)");
                sb.AppendLine("    _workbook_cache = gc.open_by_key(config['spreadsheet_id'])");
            }
            sb.AppendLine("    return _workbook_cache");
            sb.AppendLine();
            sb.AppendLine();
        }

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
        sb.AppendLine($"  \"spreadsheet_id\": \"{spreadsheetId ?? "ここにGoogle SpreadsheetのIDを入力"}\",");
        sb.AppendLine("  \"credentials_file\": \"credentials.json\"");
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
        sb.AppendLine("【初回セットアップ】");
        sb.AppendLine("1. Pythonがインストールされていることを確認してください");
        sb.AppendLine("   （https://www.python.org からダウンロード可能）");
        sb.AppendLine("2. 「setup.bat」をダブルクリックしてください");
        sb.AppendLine("   必要なライブラリが自動でインストールされます");
        sb.AppendLine();
        sb.AppendLine("【実行方法】");
        sb.AppendLine("「run.bat」をダブルクリックしてください。");
        sb.AppendLine("メニューが表示されるので、番号を入力して機能を選んでください。");
        sb.AppendLine();

        if (!string.IsNullOrEmpty(spreadsheetId))
        {
            sb.AppendLine("【Google Sheetsとの連携】");
            sb.AppendLine("このプログラムはGoogle Sheetsのデータを読み書きできます。");
            sb.AppendLine($"連携先: https://docs.google.com/spreadsheets/d/{spreadsheetId}");
            sb.AppendLine("※ 初回はGoogleのサービスアカウント認証ファイル（credentials.json）が必要です。");
            sb.AppendLine();
        }

        sb.AppendLine("【含まれている機能】");
        foreach (var result in results)
        {
            sb.AppendLine($"  - {result.ModuleName}");
        }
        sb.AppendLine();
        sb.AppendLine("【困ったときは】");
        sb.AppendLine("・「setup.bat」でエラーが出る → Pythonがインストールされていない可能性があります");
        sb.AppendLine("・「run.bat」でエラーが出る → setup.batを先に実行してください");
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

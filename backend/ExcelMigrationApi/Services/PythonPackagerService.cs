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
                writer.WriteLine("pywin32>=300");
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

        // Extract callable functions from each module's Python code
        var menuItems = new List<(string ModuleName, string FuncName, string DisplayName)>();
        var funcRegex = new System.Text.RegularExpressions.Regex(
            @"^def\s+(\w+)\s*\(", System.Text.RegularExpressions.RegexOptions.Multiline);

        foreach (var result in results)
        {
            if (string.IsNullOrEmpty(result.PythonCode)) continue;
            var moduleSafe = MakeSafeName(result.ModuleName).ToLowerInvariant();

            foreach (System.Text.RegularExpressions.Match m in funcRegex.Matches(result.PythonCode))
            {
                var funcName = m.Groups[1].Value;
                // Skip private/internal functions and get_spreadsheet helper
                if (funcName.StartsWith("_") || funcName == "get_spreadsheet" || funcName == "get_path")
                    continue;
                // Use function name as display, replacing underscores with spaces
                var display = funcName.Replace("_", " ");
                menuItems.Add((moduleSafe, funcName, display));
            }
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
            // Fallback: list modules
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
                var (modName, funcName, displayName) = menuItems[i];
                // Check if function takes no args or has defaults
                sb.AppendLine($"        elif choice == \"{i + 1}\":");
                sb.AppendLine($"            try:");
                sb.AppendLine($"                print(\"実行中: {displayName}...\")");
                sb.AppendLine($"                {modName}.{funcName}()");
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

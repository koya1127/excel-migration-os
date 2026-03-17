using System.Text.RegularExpressions;
using ExcelMigrationApi.Models;

namespace ExcelMigrationApi.Services;

public class TrackRouterService
{
    private static readonly (Regex Pattern, string Category, string Description)[] Track2Patterns =
    {
        // File I/O
        (new Regex(@"Open\s+.+\s+For\s+(Input|Output|Append|Binary)", RegexOptions.IgnoreCase | RegexOptions.Compiled),
            "FileIO", "ローカルファイル読み書き"),
        (new Regex(@"\bFileCopy\b", RegexOptions.IgnoreCase | RegexOptions.Compiled),
            "FileIO", "ファイルコピー"),
        (new Regex(@"\bKill\b\s+", RegexOptions.IgnoreCase | RegexOptions.Compiled),
            "FileIO", "ファイル削除"),
        (new Regex(@"\bDir\s*\(", RegexOptions.IgnoreCase | RegexOptions.Compiled),
            "FileIO", "フォルダ内ファイル一覧"),
        (new Regex(@"\bMkDir\b", RegexOptions.IgnoreCase | RegexOptions.Compiled),
            "FileIO", "フォルダ作成"),
        (new Regex(@"\bRmDir\b", RegexOptions.IgnoreCase | RegexOptions.Compiled),
            "FileIO", "フォルダ削除"),
        (new Regex(@"FileSystemObject", RegexOptions.IgnoreCase | RegexOptions.Compiled),
            "FileIO", "FileSystemObject"),
        (new Regex(@"\bFreeFile\b", RegexOptions.IgnoreCase | RegexOptions.Compiled),
            "FileIO", "ファイルハンドル操作"),
        (new Regex(@"Print\s+#\d+", RegexOptions.IgnoreCase | RegexOptions.Compiled),
            "FileIO", "ファイル書き込み"),
        (new Regex(@"Line\s+Input\s+#", RegexOptions.IgnoreCase | RegexOptions.Compiled),
            "FileIO", "ファイル行読み取り"),
        (new Regex(@"Input\s+#\d+", RegexOptions.IgnoreCase | RegexOptions.Compiled),
            "FileIO", "ファイル読み取り"),

        // COM / Shell
        (new Regex(@"CreateObject\s*\(", RegexOptions.IgnoreCase | RegexOptions.Compiled),
            "COM", "COM オブジェクト生成"),
        (new Regex(@"WScript\.", RegexOptions.IgnoreCase | RegexOptions.Compiled),
            "COM", "Windows Script Host"),
        (new Regex(@"\bShell\b\s*[\(\""]", RegexOptions.IgnoreCase | RegexOptions.Compiled),
            "Shell", "外部プログラム実行"),
        (new Regex(@"\bSendKeys\b", RegexOptions.IgnoreCase | RegexOptions.Compiled),
            "Shell", "キー入力シミュレーション"),

        // ActiveX / Form state
        (new Regex(@"ControlFormat\.Value", RegexOptions.IgnoreCase | RegexOptions.Compiled),
            "FormControl", "フォーム部品の状態取得"),
        (new Regex(@"OLEObjects", RegexOptions.IgnoreCase | RegexOptions.Compiled),
            "FormControl", "ActiveXコントロール"),
        (new Regex(@"\.OLEType", RegexOptions.IgnoreCase | RegexOptions.Compiled),
            "FormControl", "OLEオブジェクト"),

        // Clipboard
        (new Regex(@"DataObject", RegexOptions.IgnoreCase | RegexOptions.Compiled),
            "Clipboard", "クリップボード操作"),
        (new Regex(@"MSForms\.DataObject", RegexOptions.IgnoreCase | RegexOptions.Compiled),
            "Clipboard", "クリップボード操作"),

        // Win API
        (new Regex(@"Declare\s+(PtrSafe\s+)?Function", RegexOptions.IgnoreCase | RegexOptions.Compiled),
            "WinAPI", "Windows API呼び出し"),
        (new Regex(@"Declare\s+(PtrSafe\s+)?Sub", RegexOptions.IgnoreCase | RegexOptions.Compiled),
            "WinAPI", "Windows API呼び出し"),

        // Registry
        (new Regex(@"SaveSetting\b|GetSetting\b|DeleteSetting\b", RegexOptions.IgnoreCase | RegexOptions.Compiled),
            "Registry", "レジストリ操作"),

        // External macro call
        (new Regex(@"Application\.Run\s+""[^""]*!",  RegexOptions.IgnoreCase | RegexOptions.Compiled),
            "ExternalCall", "外部マクロ呼び出し"),

        // StrConv (partial — GAS can't do natively)
        (new Regex(@"\bStrConv\b", RegexOptions.IgnoreCase | RegexOptions.Compiled),
            "StringConv", "全角半角変換"),
        (new Regex(@"\bASC\s*\(", RegexOptions.IgnoreCase | RegexOptions.Compiled),
            "StringConv", "全角→半角変換"),
    };

    public TrackResult Route(List<VbaModule> modules)
    {
        var result = new TrackResult();

        foreach (var module in modules)
        {
            if (string.IsNullOrEmpty(module.Code) || module.CodeLines <= 1)
            {
                // Skip empty modules
                continue;
            }

            var detectedPatterns = new List<string>();
            var reasons = new List<string>();

            foreach (var (pattern, category, description) in Track2Patterns)
            {
                if (pattern.IsMatch(module.Code))
                {
                    detectedPatterns.Add($"{category}: {description}");
                    if (!reasons.Contains(description))
                        reasons.Add(description);
                }
            }

            var decision = new TrackDecision
            {
                ModuleName = module.ModuleName,
                SourceFile = module.SourceFile,
                DetectedPatterns = detectedPatterns
            };

            if (detectedPatterns.Count > 0)
            {
                decision.Track = 2;
                decision.Reason = string.Join("、", reasons);
                result.Track2Modules.Add(module);
            }
            else
            {
                decision.Track = 1;
                decision.Reason = "ローカル専用パターンなし（GAS変換可能）";
                result.Track1Modules.Add(module);
            }

            result.Decisions.Add(decision);
        }

        return result;
    }
}

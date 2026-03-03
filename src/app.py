"""Streamlit Web UI for Excel Migration OS."""

from __future__ import annotations

import contextlib
import io
import json
import os
import sys
from dataclasses import asdict
from pathlib import Path

import streamlit as st
from dotenv import load_dotenv

# Ensure src/ is importable
sys.path.insert(0, str(Path(__file__).parent))

from main import (
    ConvertReport,
    DeployReport,
    ExtractReport,
    GroupSummary,
    ScanReport,
    UploadReport,
    build_group_summaries,
    run_convert,
    run_deploy,
    run_extract,
    run_scan,
    run_upload,
)

load_dotenv()

# ---------------------------------------------------------------------------
# Page config
# ---------------------------------------------------------------------------
st.set_page_config(page_title="Excel Migration OS", page_icon="📊", layout="wide")

# ---------------------------------------------------------------------------
# Sidebar - command selection & common settings
# ---------------------------------------------------------------------------
st.sidebar.title("Excel Migration OS")
page = st.sidebar.radio(
    "コマンド選択",
    ["スキャン", "抽出", "変換", "アップロード", "デプロイ", "マイグレーション"],
)

st.sidebar.markdown("---")
st.sidebar.subheader("共通設定")
credentials_path = st.sidebar.text_input(
    "OAuth認証ファイル", value="credentials/client_secret.json"
)
token_path = st.sidebar.text_input("トークンファイル", value="credentials/token.json")


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------
def _capture(func, *args, **kwargs):
    """Run *func* while capturing stdout. Returns (result, captured_text).

    Also catches SystemExit raised by run_convert / run_deploy on error.
    """
    buf = io.StringIO()
    result = None
    error = None
    with contextlib.redirect_stdout(buf):
        try:
            result = func(*args, **kwargs)
        except SystemExit as exc:
            error = str(exc)
        except Exception as exc:
            error = str(exc)
    return result, buf.getvalue(), error


def _show_log(log_text: str):
    if log_text.strip():
        with st.expander("実行ログ", expanded=False):
            st.code(log_text, language="text")


def _resolve(path_str: str) -> Path:
    return Path(path_str).expanduser().resolve()


def _show_workbook_context(output_path_str: str):
    """Display workbook context from extracted workbook_context.json files."""
    output_dir = _resolve(output_path_str)
    ctx_files = sorted(output_dir.rglob("workbook_context.json"))
    if not ctx_files:
        return

    st.subheader("ワークブックコンテキスト")
    for ctx_file in ctx_files:
        try:
            ctx = json.loads(ctx_file.read_text(encoding="utf-8"))
        except Exception:
            continue

        wb_name = ctx_file.parent.name
        summary = ctx.get("summary", {})

        with st.expander(f"{wb_name} - コンテキスト情報", expanded=False):
            # Summary metrics
            mc1, mc2, mc3, mc4 = st.columns(4)
            mc1.metric("シート数", summary.get("sheet_count", 0))
            mc2.metric("非空セル", summary.get("total_non_empty_cells", 0))
            mc3.metric("数式", summary.get("total_formulas", 0))
            mc4.metric("入力規則", summary.get("total_data_validations", 0))

            s_incompat = summary.get("total_incompatible_functions", 0)
            s_udf = summary.get("total_udf_calls", 0)
            if s_incompat or s_udf:
                mc5, mc6 = st.columns(2)
                mc5.metric("非互換関数", s_incompat)
                mc6.metric("UDF呼出し", s_udf)

            # Sheet tabs
            sheets = ctx.get("sheets", [])
            if sheets:
                sheet_names = [s["name"] for s in sheets]
                selected_sheet = st.selectbox(
                    "シート選択", sheet_names, key=f"ctx_sheet_{wb_name}"
                )
                for sd in sheets:
                    if sd["name"] != selected_sheet:
                        continue

                    st.caption(
                        f"{sd.get('row_count', 0)} rows x {sd.get('col_count', 0)} cols"
                    )

                    # Cell data
                    cells = sd.get("cells", {})
                    if cells:
                        cell_rows = [
                            {"セル": coord, "値": str(info["value"]), "型": info["type"]}
                            for coord, info in sorted(cells.items())
                        ]
                        st.write(f"セルデータ ({len(cell_rows)} cells)")
                        st.dataframe(cell_rows, use_container_width=True, height=200)

                    # Data validations
                    dvs = sd.get("data_validations", [])
                    if dvs:
                        st.write("入力規則")
                        dv_rows = [
                            {
                                "種別": dv.get("type", ""),
                                "範囲": dv.get("sqref", ""),
                                "式1": dv.get("formula1", ""),
                                "式2": dv.get("formula2", ""),
                            }
                            for dv in dvs
                        ]
                        st.dataframe(dv_rows, use_container_width=True)

                    # Formulas
                    formulas = sd.get("formulas", [])
                    if formulas:
                        st.write("数式一覧")
                        fm_rows = [
                            {"セル": f["cell"], "数式": f["formula"]}
                            for f in formulas
                        ]
                        st.dataframe(fm_rows, use_container_width=True)
                    break

            # Named ranges
            named_ranges = ctx.get("named_ranges", [])
            if named_ranges:
                st.write("名前付き範囲")
                nr_rows = [
                    {"名前": nr["name"], "値": nr["value"], "スコープ": nr["scope"]}
                    for nr in named_ranges
                ]
                st.dataframe(nr_rows, use_container_width=True)

            # Inter-sheet references
            inter_refs = ctx.get("inter_sheet_references", [])
            if inter_refs:
                st.write("シート間参照")
                ref_rows = [
                    {
                        "参照元シート": ref["source_sheet"],
                        "セル": ref["source_cell"],
                        "数式": ref["formula"],
                        "参照先": ", ".join(ref["target_sheets"]),
                    }
                    for ref in inter_refs
                ]
                st.dataframe(ref_rows, use_container_width=True)

            # Function compatibility
            func_compat = ctx.get("function_compatibility", {})
            fc_summary = func_compat.get("summary", {})
            n_incompat = fc_summary.get("total_incompatible", 0)
            n_udf = fc_summary.get("total_udf_calls", 0)
            if n_incompat or n_udf:
                st.write("関数互換性")
                fc1, fc2, fc3, fc4 = st.columns(4)
                fc1.metric("非互換関数", n_incompat)
                fc2.metric("Missing", fc_summary.get("missing_count", 0))
                fc3.metric("Partial", fc_summary.get("partial_count", 0))
                fc4.metric("UDF呼出し", n_udf)

                if n_incompat:
                    st.warning(
                        f"{n_incompat} 件の非互換Excel関数が検出されました。"
                        "Google Sheetsでは動作しない、または挙動が異なる可能性があります。"
                    )
                    incompat_rows = [
                        {
                            "シート": item["sheet"],
                            "セル": item["cell"],
                            "関数": item["function"],
                            "カテゴリ": item["category"],
                            "数式": item["formula"],
                        }
                        for item in func_compat.get("incompatible_formulas", [])
                    ]
                    if incompat_rows:
                        st.dataframe(incompat_rows, use_container_width=True)

                if n_udf:
                    st.warning(
                        f"{n_udf} 件のVBA UDF（ユーザー定義関数）がセル数式で使用されています。"
                        "Google Sheetsアップロード時にこれらの数式は消失します。"
                    )
                    udf_rows = [
                        {
                            "シート": item["sheet"],
                            "セル": item["cell"],
                            "UDF名": item["udf_name"],
                            "モジュール": item["module"],
                            "数式": item["formula"],
                        }
                        for item in func_compat.get("udf_calls", [])
                    ]
                    if udf_rows:
                        st.dataframe(udf_rows, use_container_width=True)


# ---------------------------------------------------------------------------
# Page: スキャン
# ---------------------------------------------------------------------------
def _difficulty_icon(difficulty: str) -> str:
    if difficulty == "Easy":
        return "\U0001f7e2"
    if difficulty == "Medium":
        return "\U0001f7e1"
    return "\U0001f534"


def page_scan():
    st.header("スキャン - Excelファイル分析")
    st.caption("フォルダ内のExcelファイルを分析し、リスクスコア付きレポートを生成します。")

    col1, col2 = st.columns(2)
    with col1:
        input_path = st.text_input("入力パス（ファイルまたはフォルダ）", key="scan_input")
    with col2:
        output_path = st.text_input("出力先フォルダ", value="output/scan", key="scan_output")

    group_mode = st.radio(
        "グルーピングモード",
        ["なし", "プレフィックス（ファイル名_前）", "サブフォルダ"],
        horizontal=True,
        key="scan_group_mode",
    )
    group_by_map = {
        "なし": "none",
        "プレフィックス（ファイル名_前）": "prefix",
        "サブフォルダ": "subfolder",
    }
    group_by = group_by_map[group_mode]

    if st.button("スキャン実行", type="primary", key="scan_run"):
        if not input_path:
            st.error("入力パスを指定してください。")
            return

        with st.spinner("スキャン中..."):
            report, log, err = _capture(
                run_scan, _resolve(input_path), _resolve(output_path), group_by
            )

        _show_log(log)

        if err:
            st.error(f"エラー: {err}")
            return

        st.session_state["scan_report"] = report
        st.session_state["scan_group_by"] = group_by

    report: ScanReport | None = st.session_state.get("scan_report")
    if report is None:
        return

    # Re-group if user changed mode after scan
    current_group_by = st.session_state.get("scan_group_by", "none")
    if group_by != current_group_by:
        report.groups = build_group_summaries(
            report.files, group_by, report.input_root
        )
        report.group_by = group_by
        st.session_state["scan_group_by"] = group_by

    # Overall metrics
    st.subheader("サマリー")
    c1, c2, c3 = st.columns(3)
    c1.metric("ファイル数", report.file_count)
    macro_count = sum(1 for f in report.files if f.has_macro)
    c2.metric("マクロ有り", macro_count)
    avg_risk = (
        round(sum(f.risk_score for f in report.files) / len(report.files), 1)
        if report.files
        else 0
    )
    c3.metric("平均リスクスコア", avg_risk)

    # Group-level display
    if report.groups:
        easy = sum(1 for g in report.groups if g.migration_difficulty == "Easy")
        medium = sum(1 for g in report.groups if g.migration_difficulty == "Medium")
        hard = sum(1 for g in report.groups if g.migration_difficulty == "Hard")

        st.subheader("グループ別サマリー")
        gc1, gc2, gc3, gc4 = st.columns(4)
        gc1.metric("グループ数", len(report.groups))
        gc2.metric("\U0001f7e2 Easy", easy)
        gc3.metric("\U0001f7e1 Medium", medium)
        gc4.metric("\U0001f534 Hard", hard)

        # Group summary table
        group_rows = []
        for g in report.groups:
            group_rows.append(
                {
                    "グループ名": g.group_name,
                    "ファイル数": g.file_count,
                    "マクロ有り": g.macro_file_count,
                    "平均リスク": g.avg_risk_score,
                    "最大リスク": g.max_risk_score,
                    "非互換関数": g.total_incompatible_functions,
                    "難易度": f"{_difficulty_icon(g.migration_difficulty)} {g.migration_difficulty}",
                }
            )
        st.dataframe(group_rows, use_container_width=True)

        # Drill-down per group
        st.subheader("グループ詳細")
        for g in report.groups:
            icon = _difficulty_icon(g.migration_difficulty)
            label = f"{icon} {g.group_name} ({g.file_count}ファイル, {g.migration_difficulty})"
            with st.expander(label):
                file_rows = []
                for idx in g.file_indices:
                    f = report.files[idx]
                    file_rows.append(
                        {
                            "ファイル": Path(f.path).name,
                            "拡張子": f.extension,
                            "サイズ(KB)": round(f.size_bytes / 1024, 1),
                            "マクロ": "有" if f.has_macro else "無",
                            "数式数": f.formula_count,
                            "非互換関数": f.incompatible_function_count,
                            "リスク": f.risk_score,
                            "備考": f.notes,
                        }
                    )
                st.dataframe(file_rows, use_container_width=True)
    else:
        # Flat file list (group_by=none)
        st.subheader("ファイル一覧")
        rows = []
        for f in report.files:
            rows.append(
                {
                    "ファイル": Path(f.path).name,
                    "拡張子": f.extension,
                    "サイズ(KB)": round(f.size_bytes / 1024, 1),
                    "マクロ": "有" if f.has_macro else "無",
                    "数式数": f.formula_count,
                    "リスク": f.risk_score,
                    "備考": f.notes,
                }
            )
        st.dataframe(rows, use_container_width=True)


# ---------------------------------------------------------------------------
# Page: 抽出
# ---------------------------------------------------------------------------
def page_extract():
    st.header("抽出 - VBAモジュール抽出")
    st.caption(".xlsmファイルからVBAコードとフォームコントロールを抽出します。")

    col1, col2 = st.columns(2)
    with col1:
        input_path = st.text_input("入力パス（.xlsmファイルまたはフォルダ）", key="ext_input")
    with col2:
        output_path = st.text_input("出力先フォルダ", value="output/extract", key="ext_output")

    if st.button("抽出実行", type="primary", key="ext_run"):
        if not input_path:
            st.error("入力パスを指定してください。")
            return

        with st.spinner("抽出中..."):
            report, log, err = _capture(
                run_extract, _resolve(input_path), _resolve(output_path)
            )

        _show_log(log)

        if err:
            st.error(f"エラー: {err}")
            return

        st.session_state["extract_report"] = report

    report: ExtractReport | None = st.session_state.get("extract_report")
    if report is None:
        return

    # Metrics
    st.subheader("サマリー")
    c1, c2, c3 = st.columns(3)
    c1.metric("処理ファイル数", report.file_count)
    c2.metric("VBAモジュール数", report.module_count)
    c3.metric("フォームコントロール数", len(report.controls))

    # VBA modules table
    if report.modules:
        st.subheader("VBAモジュール一覧")
        mod_rows = []
        for m in report.modules:
            mod_rows.append(
                {
                    "モジュール名": m.module_name,
                    "種別": m.module_type,
                    "行数": m.code_lines,
                    "出力先": m.output_path,
                }
            )
        st.dataframe(mod_rows, use_container_width=True)

    # Form controls table
    if report.controls:
        st.subheader("フォームコントロール一覧")
        ctrl_rows = []
        for c in report.controls:
            ctrl_rows.append(
                {
                    "シート": c.sheet_name,
                    "コントロール名": c.control_name,
                    "種別": c.control_type,
                    "ラベル": c.label,
                    "マクロ": c.macro,
                    "位置": c.anchor,
                }
            )
        st.dataframe(ctrl_rows, use_container_width=True)

    # Workbook context display
    _show_workbook_context(output_path)


# ---------------------------------------------------------------------------
# Page: 変換
# ---------------------------------------------------------------------------
def page_convert():
    st.header("変換 - VBA → Google Apps Script")
    st.caption("Claude APIを使用してVBAコードをApps Scriptに変換します。")

    col1, col2 = st.columns(2)
    with col1:
        input_path = st.text_input("入力フォルダ（.basファイル）", key="conv_input")
    with col2:
        output_path = st.text_input("出力先フォルダ", value="output/convert", key="conv_output")

    col3, col4 = st.columns(2)
    with col3:
        api_key = st.text_input(
            "Anthropic API Key",
            type="password",
            value=os.environ.get("ANTHROPIC_API_KEY", ""),
            key="conv_api_key",
            help=".envファイルから自動読込。手動入力で上書き可能。",
        )
    with col4:
        model = st.selectbox(
            "モデル",
            ["claude-sonnet-4-6", "claude-haiku-4-5-20251001", "claude-opus-4-6"],
            key="conv_model",
        )

    if st.button("変換実行", type="primary", key="conv_run"):
        if not input_path:
            st.error("入力フォルダを指定してください。")
            return
        if not api_key:
            st.error("API Keyを指定してください。")
            return

        with st.spinner("変換中（Claude API呼び出し）..."):
            report, log, err = _capture(
                run_convert,
                _resolve(input_path),
                _resolve(output_path),
                api_key or None,
                model,
            )

        _show_log(log)

        if err:
            st.error(f"エラー: {err}")
            return

        st.session_state["convert_report"] = report
        st.session_state["_conv_log"] = log

    report: ConvertReport | None = st.session_state.get("convert_report")
    if report is None:
        return

    # Metrics
    st.subheader("サマリー")
    c1, c2, c3 = st.columns(3)
    c1.metric("合計", report.total)
    c2.metric("成功", report.success)
    c3.metric("失敗", report.failed)

    # Results table
    st.subheader("変換結果")
    result_rows = []
    for r in report.results:
        result_rows.append(
            {
                "モジュール名": r.module_name,
                "種別": r.module_type,
                "ステータス": r.status,
                "エラー": r.error or "-",
            }
        )
    st.dataframe(result_rows, use_container_width=True)

    # Button migration note
    log = st.session_state.get("_conv_log", "")
    if "form control buttons" in log:
        st.info(
            "Excelのフォームコントロールボタンは onOpen() カスタムメニューに変換されました。\n\n"
            "Google Sheets ではプログラムによるボタン作成ができないため、"
            "シート上にボタンが必要な場合は手動で追加してください:\n\n"
            "**挿入 > 図形描画** でボタンを作成 → 右クリック → **スクリプトを割り当て**"
        )

    # Incompatible function / UDF warnings from convert log
    if "incompatible Excel function" in log:
        st.warning(
            "非互換Excel関数が検出されました。変換コードに setupFormulas() または"
            "コメント付き代替処理が含まれています。出力コードを確認してください。"
        )
    if "VBA UDF call" in log:
        st.warning(
            "VBA UDF（ユーザー定義関数）がセル数式で使用されていました。"
            "@customfunction 付きのApps Scriptカスタム関数として変換されています。"
        )

    # Code preview
    success_results = [r for r in report.results if r.status == "success"]
    if success_results:
        st.subheader("コードプレビュー")
        selected = st.selectbox(
            "ファイルを選択",
            [r.module_name for r in success_results],
            key="conv_preview_select",
        )
        for r in success_results:
            if r.module_name == selected:
                try:
                    code = Path(r.output_gs).read_text(encoding="utf-8")
                    st.code(code, language="javascript")
                except Exception as exc:
                    st.warning(f"ファイル読み込みエラー: {exc}")
                break


# ---------------------------------------------------------------------------
# Page: アップロード
# ---------------------------------------------------------------------------
def page_upload():
    st.header("アップロード - Google Drive")
    st.caption("ExcelファイルをGoogle Driveにアップロードします。")

    col1, col2 = st.columns(2)
    with col1:
        input_path = st.text_input("入力パス（ファイルまたはフォルダ）", key="up_input")
    with col2:
        output_path = st.text_input("出力先フォルダ", value="output/upload", key="up_output")

    col3, col4 = st.columns(2)
    with col3:
        drive_folder_id = st.text_input(
            "DriveフォルダID（任意）", key="up_folder_id"
        )
    with col4:
        convert_to_sheets = st.checkbox(
            "Google Sheets形式に変換", value=True, key="up_convert"
        )

    if st.button("アップロード実行", type="primary", key="up_run"):
        if not input_path:
            st.error("入力パスを指定してください。")
            return

        cred = _resolve(credentials_path)
        tok = _resolve(token_path)

        if not cred.exists():
            st.error(f"認証ファイルが見つかりません: {cred}")
            return

        st.info("初回実行時はブラウザでGoogle OAuth認証が必要です。")

        with st.spinner("アップロード中..."):
            report, log, err = _capture(
                run_upload,
                _resolve(input_path),
                _resolve(output_path),
                cred,
                tok,
                drive_folder_id or None,
                convert_to_sheets,
            )

        _show_log(log)

        if err:
            st.error(f"エラー: {err}")
            return

        st.session_state["upload_report"] = report

    report: UploadReport | None = st.session_state.get("upload_report")
    if report is None:
        return

    # Metrics
    st.subheader("サマリー")
    c1, c2, c3 = st.columns(3)
    c1.metric("合計", report.file_count)
    c2.metric("成功", report.success_count)
    c3.metric("失敗", report.failure_count)

    # Results table with links
    st.subheader("アップロード結果")
    for f in report.files:
        col_a, col_b, col_c = st.columns([3, 1, 2])
        col_a.write(f.file_name)
        col_b.write(f.status)
        if f.drive_web_view_link:
            col_c.markdown(f"[Driveで開く]({f.drive_web_view_link})")
        elif f.error:
            col_c.write(f.error)


# ---------------------------------------------------------------------------
# Page: デプロイ
# ---------------------------------------------------------------------------
def page_deploy():
    st.header("デプロイ - Apps Script")
    st.caption(".gsファイルをGoogle SheetsのApps Scriptプロジェクトにデプロイします。")

    col1, col2 = st.columns(2)
    with col1:
        input_path = st.text_input("入力フォルダ（.gsファイル）", key="dep_input")
    with col2:
        spreadsheet_id = st.text_input("スプレッドシートID", key="dep_ss_id")

    script_id = st.text_input(
        "既存スクリプトID（任意 - 空欄で新規作成）", key="dep_script_id"
    )

    if st.button("デプロイ実行", type="primary", key="dep_run"):
        if not input_path:
            st.error("入力フォルダを指定してください。")
            return
        if not spreadsheet_id:
            st.error("スプレッドシートIDを指定してください。")
            return

        cred = _resolve(credentials_path)
        tok = _resolve(token_path)

        if not cred.exists():
            st.error(f"認証ファイルが見つかりません: {cred}")
            return

        st.info("初回実行時はブラウザでGoogle OAuth認証が必要です。")

        with st.spinner("デプロイ中..."):
            report, log, err = _capture(
                run_deploy,
                _resolve(input_path),
                spreadsheet_id,
                cred,
                tok,
                script_id or None,
            )

        _show_log(log)

        if err:
            st.error(f"エラー: {err}")
            return

        st.session_state["deploy_report"] = report

    report: DeployReport | None = st.session_state.get("deploy_report")
    if report is None:
        return

    # Results
    st.subheader("デプロイ結果")
    c1, c2 = st.columns(2)
    c1.metric("デプロイファイル数", report.file_count)
    c2.metric("スクリプトID", report.script_id)

    st.subheader("デプロイ済みファイル")
    for name in report.files_deployed:
        st.write(f"- {name}")

    editor_url = f"https://script.google.com/d/{report.script_id}/edit"
    st.markdown(f"[スクリプトエディタを開く]({editor_url})")


# ---------------------------------------------------------------------------
# Page: マイグレーション
# ---------------------------------------------------------------------------
def page_migrate():
    st.header("マイグレーション - End-to-End")
    st.caption("xlsmをGoogle Sheetsにアップロードし、.gsファイルをデプロイします。")

    col1, col2 = st.columns(2)
    with col1:
        xlsm_path = st.text_input("xlsmファイルパス", key="mig_xlsm")
    with col2:
        gs_dir = st.text_input(".gsファイルフォルダ", key="mig_gs_dir")

    output_path = st.text_input("出力先フォルダ", value="output/migrate", key="mig_output")

    if st.button("マイグレーション実行", type="primary", key="mig_run"):
        if not xlsm_path:
            st.error("xlsmファイルを指定してください。")
            return
        if not gs_dir:
            st.error(".gsフォルダを指定してください。")
            return

        xlsm = _resolve(xlsm_path)
        gs = _resolve(gs_dir)
        out = _resolve(output_path)
        cred = _resolve(credentials_path)
        tok = _resolve(token_path)

        if not xlsm.exists():
            st.error(f"xlsmファイルが見つかりません: {xlsm}")
            return
        if not cred.exists():
            st.error(f"認証ファイルが見つかりません: {cred}")
            return

        st.info("初回実行時はブラウザでGoogle OAuth認証が必要です。")

        # Step 1: Upload
        progress = st.progress(0, text="Step 1/2: アップロード中...")
        upload_report, up_log, up_err = _capture(
            run_upload, xlsm, out, cred, tok, None, True
        )

        _show_log(up_log)

        if up_err:
            st.error(f"アップロードエラー: {up_err}")
            return

        if upload_report.failure_count > 0 or not upload_report.files:
            st.error("アップロードに失敗しました。")
            return

        uploaded = upload_report.files[0]
        spreadsheet_id = uploaded.drive_file_id
        st.success(f"アップロード完了: {uploaded.file_name}")

        # Step 2: Deploy
        progress.progress(50, text="Step 2/2: デプロイ中...")
        deploy_report, dep_log, dep_err = _capture(
            run_deploy, gs, spreadsheet_id, cred, tok, None
        )

        _show_log(dep_log)

        if dep_err:
            st.error(f"デプロイエラー: {dep_err}")
            return

        progress.progress(100, text="完了!")

        st.session_state["migrate_result"] = {
            "upload": upload_report,
            "deploy": deploy_report,
            "spreadsheet_link": uploaded.drive_web_view_link,
        }

    result = st.session_state.get("migrate_result")
    if result is None:
        return

    st.subheader("マイグレーション結果")

    deploy_report: DeployReport = result["deploy"]

    st.markdown(f"**スプレッドシート**: [{result['spreadsheet_link']}]({result['spreadsheet_link']})")

    editor_url = f"https://script.google.com/d/{deploy_report.script_id}/edit"
    st.markdown(f"**スクリプトエディタ**: [{editor_url}]({editor_url})")

    st.write(f"デプロイファイル数: {deploy_report.file_count}")
    for name in deploy_report.files_deployed:
        st.write(f"- {name}")


# ---------------------------------------------------------------------------
# Router
# ---------------------------------------------------------------------------
PAGES = {
    "スキャン": page_scan,
    "抽出": page_extract,
    "変換": page_convert,
    "アップロード": page_upload,
    "デプロイ": page_deploy,
    "マイグレーション": page_migrate,
}

PAGES[page]()

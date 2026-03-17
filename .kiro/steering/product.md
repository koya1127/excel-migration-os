# Product Overview

Excel/VBAマクロを含むファイルをGoogle Sheets + Apps Scriptに自動移行するSaaSツール。
ターゲットは社内にExcelマクロが大量にある企業。エンドユーザーはIT素人の事務担当者。

## Core Capabilities

1. **Scan** — Excelファイルの移行リスク・互換性を自動分析
2. **Convert** — VBA→GAS / VBA→Python の二系統自動変換（Claude API）
3. **Deploy** — 変換コードをGoogle Sheetsに自動デプロイ + 使い方シート生成
4. **Migrate** — Upload→Extract→Track振り分け→Convert→Deploy を一括実行
5. **Track分岐** — GASで不可能な機能（ファイルI/O、COM等）をPythonローカル版で救済

## Target Use Cases

- Excel→Google Workspace移行プロジェクトでマクロの自動変換が必要な企業
- IT部門がマクロ移行の工数を削減したいケース
- 移行後のスプレッドシートを非IT部門がそのまま使えることが必須

## Value Proposition

- VBA→GASの変換可能な部分は自動変換、不可能な部分はPythonで補完する二系統アーキテクチャ
- 変換後スプレッドシートに「使い方」シートを自動生成（IT素人でも初見で使える）
- 従量課金（¥3/1Kトークン）で小規模からスタート可能

---
_Focus on patterns and purpose, not exhaustive feature lists_

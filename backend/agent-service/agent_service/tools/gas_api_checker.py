"""Check that GAS API calls in converted code are valid."""
import re


# Known valid GAS classes and their common methods
VALID_GAS_APIS: dict[str, set[str]] = {
    "SpreadsheetApp": {
        "getActiveSpreadsheet", "getActive", "getUi", "openById", "openByUrl",
        "create", "flush", "enableAllDataSourcesExecution",
    },
    "Spreadsheet": {
        "getSheetByName", "getSheets", "getActiveSheet", "setActiveSheet",
        "insertSheet", "deleteSheet", "getName", "getUrl", "getId",
        "getRange", "getRangeByName", "getNamedRanges", "toast",
        "addMenu", "getEditors", "addEditor", "removeEditor",
    },
    "Sheet": {
        "getRange", "getDataRange", "getLastRow", "getLastColumn",
        "getName", "setName", "getMaxRows", "getMaxColumns",
        "insertRows", "insertColumns", "deleteRows", "deleteColumns",
        "insertRowAfter", "insertRowBefore", "insertColumnAfter", "insertColumnBefore",
        "hideRow", "hideColumn", "showRow", "showColumn",
        "getFormUrl", "getCharts", "newChart",
        "activate", "clear", "clearContents", "clearFormats",
        "copyTo", "getFilter", "sort", "autoResizeColumn",
        "setColumnWidth", "setRowHeight", "getColumnWidth", "getRowHeight",
        "protect", "getProtections", "isSheetHidden", "showSheet", "hideSheet",
        "appendRow", "moveRows", "moveColumns",
    },
    "Range": {
        "getValue", "getValues", "setValue", "setValues",
        "getDisplayValue", "getDisplayValues",
        "getFormula", "getFormulas", "setFormula", "setFormulas",
        "getRow", "getColumn", "getNumRows", "getNumColumns",
        "getA1Notation", "getSheet", "offset",
        "setBackground", "setBackgrounds", "getBackground", "getBackgrounds",
        "setFontColor", "setFontColors", "setFontSize", "setFontWeight",
        "setBorder", "setNumberFormat", "setNumberFormats",
        "setHorizontalAlignment", "setVerticalAlignment",
        "setWrap", "setWrapStrategy",
        "merge", "breakApart", "isPartOfMerge", "getMergedRanges",
        "activate", "clear", "clearContent", "clearFormat",
        "copyTo", "moveTo", "getNote", "setNote", "clearNote",
        "protect", "isBlank", "setDataValidation", "getDataValidation",
        "createTextFinder", "getFilter", "sort",
        "getGridId", "getRichTextValue", "setRichTextValue",
        "insertCheckboxes", "removeCheckboxes", "isChecked",
    },
    "Ui": {
        "alert", "prompt", "createMenu", "createAddonMenu",
        "showSidebar", "showDialog", "showModalDialog", "showModelessDialog",
        "Button", "ButtonSet",
    },
    "Menu": {
        "addItem", "addSeparator", "addSubMenu", "addToUi",
    },
    "ScriptApp": {
        "newTrigger", "getProjectTriggers", "deleteTrigger",
        "getOAuthToken", "getScriptId", "getService",
    },
    "Logger": {
        "log", "getLog", "clear",
    },
    "Utilities": {
        "formatDate", "formatString", "sleep", "jsonParse", "jsonStringify",
        "base64Encode", "base64Decode", "newBlob", "zip", "unzip",
        "computeDigest", "computeHmacSignature",
    },
    "DriveApp": {
        "getFileById", "getFolderById", "createFile", "createFolder",
        "searchFiles", "searchFolders", "getRootFolder",
    },
    "UrlFetchApp": {
        "fetch", "fetchAll", "getRequest",
    },
    "PropertiesService": {
        "getScriptProperties", "getUserProperties", "getDocumentProperties",
    },
    "CacheService": {
        "getScriptCache", "getUserCache", "getDocumentCache",
    },
    "HtmlService": {
        "createHtmlOutput", "createHtmlOutputFromFile",
        "createTemplate", "createTemplateFromFile",
    },
    "ContentService": {
        "createTextOutput",
    },
    "MailApp": {
        "sendEmail", "getRemainingDailyQuota",
    },
    "GmailApp": {
        "sendEmail", "search", "getInboxThreads",
    },
    "Session": {
        "getActiveUser", "getEffectiveUser", "getScriptTimeZone",
        "getTemporaryActiveUserKey",
    },
    "Browser": {
        "msgBox", "inputBox",
    },
    "LockService": {
        "getScriptLock", "getUserLock", "getDocumentLock",
    },
}

# Methods that don't exist in GAS (common VBA conversion errors)
INVALID_PATTERNS = [
    (r"\.getDataTable\(\)", "getDataTable() does not exist in GAS. Use getDataRange() or getRange()"),
    (r"\.PrintOut\b", "PrintOut does not exist in GAS. Consider creating a PDF with getAs('application/pdf')"),
    (r"\.Copy\b(?!\()", "Sheet.Copy does not exist as-is in GAS. Use copyTo()"),
    (r"\.Paste\b", "Paste does not exist in GAS. Use Range.copyTo() with copyFormatToRange"),
    (r"\.Select\b", "Select does not exist in GAS. Use activate() or setActiveRange()"),
    (r"Application\.", "Application object does not exist in GAS. Use SpreadsheetApp"),
    (r"ActiveCell", "ActiveCell does not exist in GAS. Use getActiveRange()"),
    (r"\.Visible\s*=", "Sheet.Visible property does not exist in GAS. Use hideSheet()/showSheet()"),
    (r"\.QueryTable\b", "QueryTable does not exist in GAS. Use UrlFetchApp for external data"),
    (r"\.ListObjects?\b", "ListObjects does not exist in GAS. Use getDataRange() or getRange()"),
]


def check_gas_apis(code: str) -> dict:
    """Check GAS code for invalid API calls.

    Returns:
        {"valid": True} or {"valid": False, "issues": [...]}
    """
    issues = []

    # Check for known invalid patterns
    for pattern, message in INVALID_PATTERNS:
        matches = list(re.finditer(pattern, code))
        if matches:
            for m in matches:
                line_num = code[:m.start()].count("\n") + 1
                issues.append({
                    "line": line_num,
                    "pattern": m.group(),
                    "message": message,
                })

    # Check for method calls on known GAS objects that don't exist
    # Pattern: SomeGasClass.methodName() or someVar.unknownMethod()
    for cls, methods in VALID_GAS_APIS.items():
        # Find direct class usage like SpreadsheetApp.unknownMethod()
        pattern = rf"{cls}\.(\w+)\("
        for m in re.finditer(pattern, code):
            method = m.group(1)
            if method not in methods:
                line_num = code[:m.start()].count("\n") + 1
                issues.append({
                    "line": line_num,
                    "pattern": f"{cls}.{method}()",
                    "message": f"{method}() is not a known method of {cls}",
                })

    if issues:
        return {"valid": False, "issues": issues}
    return {"valid": True}

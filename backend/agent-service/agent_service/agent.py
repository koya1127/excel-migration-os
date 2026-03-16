"""Claude Agent SDK based VBA→GAS conversion agent."""
import asyncio
import json
import logging
import os

from claude_agent_sdk import (
    tool,
    create_sdk_mcp_server,
    ClaudeSDKClient,
    ClaudeAgentOptions,
    AssistantMessage,
    ResultMessage,
    TextBlock,
)

from .config import MODEL_LARGE, MODEL_SMALL, SMALL_MODULE_THRESHOLD, MAX_AGENT_TURNS
from .models import ConvertRequest, ConvertResult
from .tools.syntax_validator import validate_gas_syntax
from .tools.gas_api_checker import check_gas_apis
from .tools.duplicate_checker import check_duplicates

logger = logging.getLogger(__name__)

SYSTEM_PROMPT = """You are a specialist in converting Excel VBA macros to Google Apps Script.
Your job is to convert VBA code to valid, working GAS code, then validate it.

## Workflow
1. Analyze the VBA code and convert it to Google Apps Script
2. Call validate_syntax to check for JavaScript syntax errors
3. Call check_gas_api to verify all GAS API calls are valid
4. If ANY errors are found, fix the code and re-validate
5. Once the code passes both checks, call submit_result with the final code

## VBA → GAS Event Handler Mapping
| VBA Event | GAS Equivalent | Notes |
|---|---|---|
| Workbook_Open / Auto_Open | function onOpen(e) | Simple trigger |
| Worksheet_Change | function onEdit(e) | Simple trigger. Filter by e.range.getSheet().getName() |
| Worksheet_SelectionChange | installable trigger | Requires setupTriggers() |
| Workbook_BeforeClose | — | No GAS equivalent. Add TODO comment |
| Workbook_BeforeSave | installable onChange | Approximate with onChange trigger |
| Worksheet_Activate/Deactivate | — | No equivalent. Add TODO comment |

## VBA Patterns to Remove or Replace
- Application.ScreenUpdating → DELETE
- Application.EnableEvents → DELETE
- Application.DisplayAlerts → DELETE
- Application.Calculation = xlManual/xlAutomatic → DELETE
- On Error Resume Next → try/catch
- On Error GoTo → try/catch
- DoEvents → DELETE
- ActiveWorkbook → SpreadsheetApp.getActiveSpreadsheet()
- ActiveSheet → SpreadsheetApp.getActiveSheet()
- ThisWorkbook → SpreadsheetApp.getActiveSpreadsheet()
- Cells(row, col) → sheet.getRange(row, col)
- Range("A1") → sheet.getRange("A1")
- Sheets("name") → ss.getSheetByName("name")
- .Value / .Value2 → .getValue() / .getValues()
- MsgBox → SpreadsheetApp.getUi().alert()
- InputBox → SpreadsheetApp.getUi().prompt()
- PrintOut → Create PDF via getAs('application/pdf') and show download link, or show alert to print manually

## CRITICAL: GAS APIs that do NOT exist (common conversion errors)
- getDataTable() → Use getDataRange() or named range
- QueryTable.Refresh → Use UrlFetchApp.fetch() for external data
- ListObjects → Use getDataRange() or getRange()
- PrintOut → No direct print. Use getAs('application/pdf') + DriveApp or alert
- Select → Use activate() for sheets, setActiveRange() for ranges
- Application.* → Does not exist. Use SpreadsheetApp.*

## Installable Triggers
When needed, generate a setupTriggers() function and add a comment at top:
// Run setupTriggers() once to install event triggers

## Button/Menu Handling
- If button context is provided, create a custom menu in onOpen()
- Group menu items by sheet when buttons come from multiple sheets
- Merge menu creation into existing onOpen() if it already exists

## Output Rules
- Use V8 runtime syntax (let, const, arrow functions OK)
- Output ONLY the .gs code — no explanations, no markdown fences
- CRITICAL: Ensure all braces, parentheses, and brackets are properly closed
- CRITICAL: The output must be syntactically valid JavaScript
- Never leave code outside of a function or string literal
- If VBA is too complex to fully convert, output a working stub with TODO comments
- For empty modules (only Attribute lines, no actual code), return an empty string"""


# --- Custom MCP Tools ---

@tool("validate_syntax", "Validate GAS code for JavaScript syntax errors", {
    "code": str,
})
async def validate_syntax_tool(args: dict) -> dict:
    """Run JS syntax validation on GAS code."""
    code = args["code"]
    result = validate_gas_syntax(code)
    if result.get("valid"):
        return {"content": [{"type": "text", "text": "Syntax is valid. No errors found."}]}
    errors = result.get("errors", [])
    error_text = "Syntax errors found:\n"
    for e in errors:
        line = e.get("line", "?")
        msg = e.get("message", "unknown error")
        error_text += f"  Line {line}: {msg}\n"
    return {"content": [{"type": "text", "text": error_text}]}


@tool("check_gas_api", "Check if GAS API calls in the code are valid", {
    "code": str,
})
async def check_gas_api_tool(args: dict) -> dict:
    """Check for invalid GAS API usage."""
    code = args["code"]
    result = check_gas_apis(code)
    if result.get("valid"):
        return {"content": [{"type": "text", "text": "All GAS API calls are valid."}]}
    issues = result.get("issues", [])
    issue_text = "Invalid GAS API calls found:\n"
    for i in issues:
        line = i.get("line", "?")
        pattern = i.get("pattern", "")
        msg = i.get("message", "")
        issue_text += f"  Line {line}: {pattern} — {msg}\n"
    return {"content": [{"type": "text", "text": issue_text}]}


@tool("submit_result", "Submit the final validated GAS code", {
    "gas_code": str,
    "notes": str,
})
async def submit_result_tool(args: dict) -> dict:
    """Terminal tool: submit the conversion result."""
    return {"content": [{"type": "text", "text": "Result submitted."}]}


# Create MCP server with custom tools
mcp_server = create_sdk_mcp_server(
    "convert-tools",
    tools=[validate_syntax_tool, check_gas_api_tool, submit_result_tool],
)


def _build_user_message(request: ConvertRequest) -> str:
    """Build the user message from a ConvertRequest, same logic as C# ConvertService."""
    parts = []
    parts.append(f"Module Name: {request.module_name}")
    parts.append(f"Module Type: {request.module_type}")
    if request.sheet_name:
        parts.append(f"Sheet Name: {request.sheet_name}")
    parts.append("")

    if request.detected_events:
        parts.append("Detected VBA Events in this module:")
        for evt in request.detected_events:
            parts.append(f"- {evt.vba_event_name} → GAS: {evt.gas_trigger_type} | {evt.gas_notes}")
        parts.append("")

    parts.append("VBA Code:")
    parts.append("```vba")
    parts.append(request.vba_code)
    parts.append("```")

    if request.button_context:
        parts.append("")
        parts.append("Button Context (form controls that trigger macros):")
        # Group by sheet
        by_sheet: dict[str, list] = {}
        for btn in request.button_context:
            sheet = btn.sheet_name or "(unknown sheet)"
            by_sheet.setdefault(sheet, []).append(btn)
        for sheet, btns in by_sheet.items():
            parts.append(f"  Sheet: {sheet}")
            for btn in btns:
                parts.append(f'    - {btn.control_type} "{btn.label}" → calls "{btn.macro}"')
        parts.append("")
        parts.append("Generate an onOpen() function that creates a custom menu with these buttons.")

        has_open = any(
            e.vba_event_name.lower() in ("workbook_open", "auto_open")
            for e in (request.detected_events or [])
        )
        if has_open:
            parts.append("IMPORTANT: This module already contains Workbook_Open/Auto_Open. Merge menu creation INTO the converted onOpen().")

    return "\n".join(parts)


async def convert_single_module(request: ConvertRequest) -> ConvertResult:
    """Convert a single VBA module using Claude Agent SDK."""
    # Skip empty modules
    code_lines = [
        line for line in request.vba_code.split("\n")
        if not line.strip().startswith("Attribute ") and line.strip()
    ]
    if len(code_lines) <= 1:
        return ConvertResult(
            module_name=request.module_name,
            source_file=request.source_file,
            gas_code="",
            status="success",
        )

    user_message = _build_user_message(request)

    # Choose model based on module size
    total_lines = request.vba_code.count("\n") + 1
    model = MODEL_SMALL if total_lines <= SMALL_MODULE_THRESHOLD else MODEL_LARGE

    try:
        # Explicitly pass ANTHROPIC_API_KEY so Claude CLI uses it for auth
        env = {}
        api_key = os.environ.get("ANTHROPIC_API_KEY", "")
        if api_key:
            env["ANTHROPIC_API_KEY"] = api_key

        options = ClaudeAgentOptions(
            model=model,
            system_prompt=SYSTEM_PROMPT,
            max_turns=MAX_AGENT_TURNS,
            mcp_servers={"convert-tools": mcp_server},
            permission_mode="bypassPermissions",
            env=env,
        )

        gas_code = ""
        total_input = 0
        total_output = 0

        async with ClaudeSDKClient(options=options) as client:
            await client.query(user_message)
            async for message in client.receive_response():
                if isinstance(message, AssistantMessage):
                    for block in message.content:
                        # Capture submit_result tool call input
                        if hasattr(block, "name") and "submit_result" in (block.name or ""):
                            input_data = block.input if hasattr(block, "input") else {}
                            if isinstance(input_data, dict):
                                gas_code = input_data.get("gas_code", gas_code)

                # Capture final text as fallback (if agent doesn't use submit_result)
                if isinstance(message, ResultMessage):
                    if message.result and not gas_code:
                        gas_code = message.result
                    # Get token usage from ResultMessage
                    if hasattr(message, "usage") and message.usage:
                        usage = message.usage
                        total_input = usage.get("input_tokens", 0)
                        total_input += usage.get("cache_creation_input_tokens", 0)
                        total_input += usage.get("cache_read_input_tokens", 0)
                        total_output = usage.get("output_tokens", 0)

        # Strip markdown fences if present
        gas_code = _strip_code_fences(gas_code)

        return ConvertResult(
            module_name=request.module_name,
            source_file=request.source_file,
            gas_code=gas_code,
            status="success",
            input_tokens=total_input,
            output_tokens=total_output,
        )

    except Exception as e:
        logger.error("Agent conversion failed for %s: %s", request.module_name, str(e))
        return ConvertResult(
            module_name=request.module_name,
            source_file=request.source_file,
            status="error",
            error=f"エージェント変換エラー: {str(e)[:200]}",
        )


async def convert_batch(requests: list[ConvertRequest]) -> list[ConvertResult]:
    """Convert multiple modules with cross-module coordination."""
    # Phase 1: Convert each module individually (concurrent)
    sem = asyncio.Semaphore(5)

    async def convert_with_limit(req: ConvertRequest) -> ConvertResult:
        async with sem:
            return await convert_single_module(req)

    results = await asyncio.gather(*[convert_with_limit(r) for r in requests])

    # Phase 2: Check for duplicate triggers across modules
    successful = [
        {"name": r.module_name, "code": r.gas_code}
        for r in results
        if r.status == "success" and r.gas_code
    ]

    if len(successful) > 1:
        dup_check = check_duplicates(successful)
        if dup_check["has_duplicates"]:
            results = await _merge_duplicates(list(results), dup_check["duplicates"])

    return list(results)


async def _merge_duplicates(
    results: list[ConvertResult],
    duplicates: list[dict],
) -> list[ConvertResult]:
    """Use Claude to merge duplicate trigger functions across modules."""
    # Build merge prompt
    parts = ["The following modules have duplicate trigger functions that need to be merged:\n"]

    for dup in duplicates:
        parts.append(f"Duplicate function: {dup['function']}()")
        parts.append(f"Found in modules: {', '.join(dup['modules'])}\n")

    parts.append("\nCurrent module codes:\n")
    for r in results:
        if r.gas_code:
            parts.append(f"--- {r.module_name} ---")
            parts.append(r.gas_code)
            parts.append("")

    parts.append("""
Rules for merging:
1. Keep ONE onOpen() in the primary module (ThisWorkbook or first Document module)
2. Merge all menu.addItem() calls into that single onOpen()
3. Remove onOpen() from other modules
4. For duplicate onEdit(), merge into a single function that dispatches by sheet name
5. Output each module's final code using submit_merged_results tool
6. Validate syntax of each merged module before submitting
""")

    merge_prompt = "\n".join(parts)

    @tool("submit_merged_results", "Submit merged module codes", {
        "modules": str,  # JSON string: [{"name": str, "code": str}]
    })
    async def submit_merged_tool(args: dict) -> dict:
        return {"content": [{"type": "text", "text": "Merged results submitted."}]}

    merge_server = create_sdk_mcp_server(
        "merge-tools",
        tools=[validate_syntax_tool, check_gas_api_tool, submit_merged_tool],
    )

    try:
        env = {}
        api_key = os.environ.get("ANTHROPIC_API_KEY", "")
        if api_key:
            env["ANTHROPIC_API_KEY"] = api_key

        options = ClaudeAgentOptions(
            model=MODEL_LARGE,
            system_prompt="You are an expert at merging Google Apps Script modules. Merge duplicate trigger functions across modules while preserving all functionality.",
            max_turns=MAX_AGENT_TURNS,
            mcp_servers={"merge-tools": merge_server},
            permission_mode="bypassPermissions",
            env=env,
        )

        merged_modules = {}

        async with ClaudeSDKClient(options=options) as client:
            await client.query(merge_prompt)
            async for message in client.receive_response():
                if isinstance(message, AssistantMessage):
                    for block in message.content:
                        if hasattr(block, "name") and "submit_merged_results" in (block.name or ""):
                            input_data = block.input if hasattr(block, "input") else {}
                            if isinstance(input_data, dict):
                                try:
                                    modules_json = json.loads(input_data.get("modules", "[]"))
                                    for mod in modules_json:
                                        merged_modules[mod["name"]] = mod["code"]
                                except (json.JSONDecodeError, KeyError, TypeError):
                                    pass

        # Apply merged code back to results
        for r in results:
            if r.module_name in merged_modules:
                r.gas_code = _strip_code_fences(merged_modules[r.module_name])

    except Exception as e:
        logger.warning("Merge failed, keeping individual results: %s", str(e))

    return results


def _strip_code_fences(code: str) -> str:
    """Remove markdown code fences if present."""
    if not code:
        return code
    lines = code.split("\n")
    if lines and lines[0].strip().startswith("```"):
        lines = lines[1:]
    if lines and lines[-1].strip() == "```":
        lines = lines[:-1]
    return "\n".join(lines).strip()

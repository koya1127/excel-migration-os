"""GAS/JavaScript syntax validation using Node.js acorn parser."""
import subprocess
import json
import tempfile
import os


def validate_gas_syntax(code: str) -> dict:
    """Validate GAS code syntax using acorn JS parser.

    Returns:
        {"valid": True} or {"valid": False, "errors": [...]}
    """
    # Write code to temp file
    with tempfile.NamedTemporaryFile(mode="w", suffix=".js", delete=False, encoding="utf-8") as f:
        f.write(code)
        tmp_path = f.name

    try:
        # Use acorn to parse (sourceType: module to allow top-level functions)
        check_script = f"""
const acorn = require('acorn');
const fs = require('fs');
const code = fs.readFileSync('{tmp_path.replace(os.sep, "/")}', 'utf-8');
try {{
    acorn.parse(code, {{
        ecmaVersion: 2020,
        sourceType: 'script',
        allowReturnOutsideFunction: true
    }});
    console.log(JSON.stringify({{valid: true}}));
}} catch (e) {{
    console.log(JSON.stringify({{
        valid: false,
        errors: [{{
            message: e.message,
            line: e.loc ? e.loc.line : null,
            column: e.loc ? e.loc.column : null
        }}]
    }}));
}}
"""
        result = subprocess.run(
            ["node", "-e", check_script],
            capture_output=True, text=True, timeout=10
        )

        if result.returncode == 0 and result.stdout.strip():
            return json.loads(result.stdout.strip())

        # If acorn not installed, fall back to node --check approach
        if "Cannot find module 'acorn'" in result.stderr:
            return _fallback_syntax_check(code, tmp_path)

        return {"valid": False, "errors": [{"message": f"Validator error: {result.stderr[:200]}"}]}

    finally:
        os.unlink(tmp_path)


def _fallback_syntax_check(code: str, tmp_path: str) -> dict:
    """Fallback: wrap code in a function and use node --check."""
    wrapped_path = tmp_path + ".check.js"
    try:
        # Wrap in function to allow function declarations at top level
        with open(wrapped_path, "w", encoding="utf-8") as f:
            f.write(f"(function() {{\n{code}\n}})();")

        result = subprocess.run(
            ["node", "--check", wrapped_path],
            capture_output=True, text=True, timeout=10
        )
        if result.returncode == 0:
            return {"valid": True}

        return {
            "valid": False,
            "errors": [{"message": result.stderr.strip()[:500]}]
        }
    finally:
        if os.path.exists(wrapped_path):
            os.unlink(wrapped_path)

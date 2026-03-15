"""Check for duplicate trigger functions across multiple modules."""
import re


TRIGGER_FUNCTIONS = [
    "onOpen",
    "onEdit",
    "onSelectionChange",
    "onChange",
    "onFormSubmit",
    "onInstall",
    "setupTriggers",
]


def check_duplicates(modules: list[dict]) -> dict:
    """Check for duplicate trigger/event functions across modules.

    Args:
        modules: [{"name": str, "code": str}, ...]

    Returns:
        {"has_duplicates": bool, "duplicates": [{"function": str, "modules": [str]}]}
    """
    # Map function name -> list of module names where it appears
    func_locations: dict[str, list[str]] = {}

    for mod in modules:
        code = mod.get("code", "")
        name = mod.get("name", "")

        for func_name in TRIGGER_FUNCTIONS:
            # Match function declaration: function onOpen(... or function onOpen (...
            pattern = rf"\bfunction\s+{func_name}\s*\("
            if re.search(pattern, code):
                func_locations.setdefault(func_name, []).append(name)

    duplicates = []
    for func_name, module_names in func_locations.items():
        if len(module_names) > 1:
            duplicates.append({
                "function": func_name,
                "modules": module_names,
            })

    return {
        "has_duplicates": len(duplicates) > 0,
        "duplicates": duplicates,
    }

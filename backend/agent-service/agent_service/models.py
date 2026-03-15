"""Pydantic models matching C# ConvertModels."""
from pydantic import BaseModel


class FormControl(BaseModel):
    control_type: str = ""
    label: str = ""
    macro: str = ""
    sheet_name: str = ""


class VbaEvent(BaseModel):
    vba_event_name: str = ""
    gas_trigger_type: str = ""
    gas_notes: str = ""


class ConvertRequest(BaseModel):
    vba_code: str = ""
    module_name: str = ""
    module_type: str = ""
    source_file: str = ""
    sheet_name: str = ""
    button_context: list[FormControl] | None = None
    detected_events: list[VbaEvent] | None = None


class ConvertResult(BaseModel):
    module_name: str = ""
    gas_code: str = ""
    status: str = ""  # "success" or "error"
    error: str = ""
    source_file: str = ""
    input_tokens: int = 0
    output_tokens: int = 0


class ConvertReport(BaseModel):
    generated_utc: str = ""
    total: int = 0
    success: int = 0
    failed: int = 0
    total_input_tokens: int = 0
    total_output_tokens: int = 0
    results: list[ConvertResult] = []


class BatchConvertRequest(BaseModel):
    modules: list[ConvertRequest]

"""FastAPI application for the VBA→GAS conversion agent service."""
import asyncio
import logging
from datetime import datetime, timezone

from fastapi import FastAPI, HTTPException
from pydantic import BaseModel

from .models import ConvertRequest, ConvertResult, ConvertReport, BatchConvertRequest
from .agent import convert_single_module, convert_batch
from .config import PORT

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

app = FastAPI(title="VBA→GAS Agent Service", version="1.0.0")


@app.get("/health")
async def health():
    return {"status": "ok"}


@app.post("/convert", response_model=ConvertResult)
async def convert_single(request: ConvertRequest):
    """Convert a single VBA module to GAS with agent validation loop."""
    logger.info("Converting module: %s", request.module_name)
    result = await convert_single_module(request)
    if result.status == "error":
        logger.error("Conversion failed for %s: %s", request.module_name, result.error)
    else:
        logger.info("Conversion succeeded for %s (tokens: %d/%d)",
                     request.module_name, result.input_tokens, result.output_tokens)
    return result


@app.post("/convert/batch", response_model=ConvertReport)
async def convert_batch_endpoint(request: BatchConvertRequest):
    """Convert multiple VBA modules with cross-module coordination."""
    if len(request.modules) > 50:
        raise HTTPException(status_code=400, detail="最大50モジュールまでです")

    logger.info("Batch converting %d modules", len(request.modules))

    results = await convert_batch(request.modules)

    report = ConvertReport(
        generated_utc=datetime.now(timezone.utc).isoformat(),
        total=len(results),
        success=sum(1 for r in results if r.status == "success"),
        failed=sum(1 for r in results if r.status == "error"),
        total_input_tokens=sum(r.input_tokens for r in results),
        total_output_tokens=sum(r.output_tokens for r in results),
        results=results,
    )

    logger.info("Batch complete: %d/%d succeeded, %d tokens",
                 report.success, report.total,
                 report.total_input_tokens + report.total_output_tokens)

    return report


if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=PORT)

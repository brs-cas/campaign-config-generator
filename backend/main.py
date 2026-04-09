from datetime import date, datetime
from typing import Optional

from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import Response
from pydantic import BaseModel, field_validator, model_validator

from generator import CampaignConfigGenerator

app = FastAPI(title="Campaign Config Generator")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

generator = CampaignConfigGenerator()


# --- Request / response models ---


class Phase(BaseModel):
    name: str
    start_date: str
    end_date: str

    @field_validator("end_date")
    @classmethod
    def end_after_start(cls, v, info):
        start = info.data.get("start_date")
        if start and v:
            try:
                if datetime.strptime(v, "%Y-%m-%d") <= datetime.strptime(start, "%Y-%m-%d"):
                    raise ValueError("end_date must be after start_date")
            except ValueError as e:
                if "end_date must be after start_date" in str(e):
                    raise
                raise ValueError("Dates must be in YYYY-MM-DD format")
        return v


class CampaignConfig(BaseModel):
    campaign_name: str
    campaign_type: str  # "bau" or "peak"
    markets: list[str]
    phases: list[Phase]
    global_promo_text: str = ""
    coupon_code: Optional[str] = None

    @field_validator("campaign_name")
    @classmethod
    def name_not_empty(cls, v):
        if not v.strip():
            raise ValueError("campaign_name must not be empty")
        return v.strip()

    @field_validator("campaign_type")
    @classmethod
    def valid_type(cls, v):
        if v.lower() not in ("bau", "peak"):
            raise ValueError("campaign_type must be 'bau' or 'peak'")
        return v.lower()

    @field_validator("markets")
    @classmethod
    def at_least_one_market(cls, v):
        if not v:
            raise ValueError("At least one market must be selected")
        valid = {"US", "AU", "SG", "UK", "CA"}
        for m in v:
            if m.upper() not in valid:
                raise ValueError(f"Invalid market: {m}. Must be one of {valid}")
        return [m.upper() for m in v]

    @model_validator(mode="after")
    def validate_phases(self):
        if self.campaign_type == "peak" and not self.phases:
            raise ValueError("Peak campaigns must have at least one phase with valid dates")
        return self


class TemplateSaveRequest(BaseModel):
    campaign_name: str
    config: dict


# --- Helper ---


def _check_phase_overlaps(phases: list[Phase]) -> list[str]:
    """Return warnings for overlapping phase dates (non-blocking)."""
    warnings = []
    parsed = []
    for p in phases:
        try:
            s = datetime.strptime(p.start_date, "%Y-%m-%d").date()
            e = datetime.strptime(p.end_date, "%Y-%m-%d").date()
            parsed.append((p.name, s, e))
        except ValueError:
            continue

    for i in range(len(parsed)):
        for j in range(i + 1, len(parsed)):
            n1, s1, e1 = parsed[i]
            n2, s2, e2 = parsed[j]
            if s1 <= e2 and s2 <= e1:
                warnings.append(f"Phases '{n1}' and '{n2}' have overlapping dates")
    return warnings


# --- Endpoints ---


@app.get("/health")
def health():
    return {"status": "ok"}


@app.post("/generate")
def generate_spreadsheet(config: CampaignConfig):
    warnings = _check_phase_overlaps(config.phases)

    phases = [p.model_dump() for p in config.phases]

    xlsx_bytes = generator.generate(
        campaign_name=config.campaign_name,
        campaign_type=config.campaign_type,
        markets=config.markets,
        phases=phases,
        global_promo_text=config.global_promo_text,
        coupon_code=config.coupon_code,
    )

    safe_name = config.campaign_name.replace(" ", "_")
    filename = f"{safe_name}_config.xlsx"

    headers = {"Content-Disposition": f'attachment; filename="{filename}"'}
    if warnings:
        headers["X-Warnings"] = "; ".join(warnings)

    return Response(
        content=xlsx_bytes,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers=headers,
    )


@app.get("/templates")
def list_templates():
    return {"templates": generator.list_templates()}


@app.get("/templates/{filename}")
def get_template(filename: str):
    try:
        return generator.load_template(filename)
    except FileNotFoundError:
        raise HTTPException(status_code=404, detail="Template not found")


@app.post("/templates")
def save_template(request: TemplateSaveRequest):
    filename = generator.save_template(request.campaign_name, request.config)
    return {"filename": filename}


@app.get("/config/touchpoints")
def get_touchpoints():
    return generator.touchpoints_config


@app.get("/config/markets")
def get_markets():
    return generator.market_defaults

"""M0 -- tenant configuration (the input side of SDD 5.1's Template Registry).

v1 originally had exactly one onboarded tenant (the reference/sample deck, for
dogfooding -- spike Doc S1). M7 (SDD v1.9) added a second, "meridian" -- a
different template, manifest, brand, shape-naming convention, and slide
ordering -- to prove the engine is genuinely tenant-agnostic, not implicitly
coupled to the reference tenant's specifics. This dict is still a hardcoded
stand-in for a real (DB-backed) tenant registry; the shape -- tenant_id ->
template + manifest paths -- is what a future real registry extends, not
replaces.
"""
from dataclasses import dataclass
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parent.parent


@dataclass(frozen=True)
class TenantConfig:
    tenant_id: str
    template_path: Path
    manifest_path: Path


_TENANTS = {
    "default": TenantConfig(
        tenant_id="default",
        template_path=REPO_ROOT / "Templates" / "MasterDeck.pptx",
        manifest_path=REPO_ROOT / "Templates" / "MasterDeck_manifest.json",
    ),
    "meridian": TenantConfig(
        tenant_id="meridian",
        template_path=REPO_ROOT / "Templates" / "Meridian.pptx",
        manifest_path=REPO_ROOT / "Templates" / "Meridian_manifest.json",
    ),
}

DEFAULT_TENANT_ID = "default"
# claude-sonnet-5 (Anthropic) is the SDD's reference model for the Content
# Generator (D6/SS8); ollama and gemini are supported swappable alternatives
# (see llm_providers.py) -- per-provider default model names live there.
DEFAULT_PROVIDER = "anthropic"


def get_tenant_config(tenant_id: str) -> TenantConfig:
    try:
        return _TENANTS[tenant_id]
    except KeyError:
        raise KeyError(f"unknown tenant_id '{tenant_id}'; registered tenants: {list(_TENANTS)}")

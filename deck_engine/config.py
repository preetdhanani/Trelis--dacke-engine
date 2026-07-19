"""M0 -- tenant configuration (the input side of SDD 5.1's Template Registry).

v1 has exactly one onboarded tenant (the reference/sample deck, for dogfooding -- spike Doc
S1). Multi-tenant registration (SDD M7) will replace this hardcoded dict with
a real lookup (DB-backed); the shape -- tenant_id -> template + manifest paths
-- is what M7 extends, not replaces.
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

import json
import tempfile
import unittest
from pathlib import Path

from deck_engine.config import get_tenant_config, TenantConfig
from deck_engine.template_registry import (
    ManifestValidationError,
    load_tenant_assets,
    load_tenant_assets_from_config,
)


class TestTemplateRegistry(unittest.TestCase):
    def test_real_default_manifest_loads_and_validates(self):
        assets = load_tenant_assets("default")
        self.assertEqual(assets.tenant_id, "default")
        self.assertEqual(len(assets.manifest.layouts), 8)
        layout_ids = {l.layout_id for l in assets.manifest.layouts}
        self.assertIn("text_bullets", layout_ids)
        self.assertIn("title", layout_ids)

    def test_open_template_returns_pristine_presentation_each_call(self):
        assets = load_tenant_assets("default")
        prs1 = assets.open_template()
        prs2 = assets.open_template()
        self.assertIsNot(prs1, prs2)
        self.assertEqual(len(prs1.slides), len(prs2.slides))

    def test_unknown_tenant_raises_keyerror(self):
        with self.assertRaises(KeyError):
            get_tenant_config("not_a_real_tenant")

    def test_bad_shape_name_in_manifest_is_fatal_at_load(self):
        real_config = get_tenant_config("default")
        manifest = json.loads(real_config.manifest_path.read_text(encoding="utf-8"))
        manifest["layouts"][0]["slots"][0]["shape_name"] = "IM_THIS_SHAPE_DOES_NOT_EXIST"

        with tempfile.TemporaryDirectory() as tmp:
            broken_manifest_path = Path(tmp) / "broken_manifest.json"
            broken_manifest_path.write_text(json.dumps(manifest), encoding="utf-8")
            broken_config = TenantConfig(
                tenant_id="default_broken",
                template_path=real_config.template_path,
                manifest_path=broken_manifest_path,
            )
            with self.assertRaises(ManifestValidationError) as ctx:
                load_tenant_assets_from_config(broken_config)
            self.assertIn("IM_THIS_SHAPE_DOES_NOT_EXIST", str(ctx.exception))

    def test_out_of_range_slide_index_is_fatal_at_load(self):
        real_config = get_tenant_config("default")
        manifest = json.loads(real_config.manifest_path.read_text(encoding="utf-8"))
        manifest["layouts"][0]["slide_index"] = 999

        with tempfile.TemporaryDirectory() as tmp:
            broken_manifest_path = Path(tmp) / "broken_manifest.json"
            broken_manifest_path.write_text(json.dumps(manifest), encoding="utf-8")
            broken_config = TenantConfig(
                tenant_id="default_broken",
                template_path=real_config.template_path,
                manifest_path=broken_manifest_path,
            )
            with self.assertRaises(ManifestValidationError) as ctx:
                load_tenant_assets_from_config(broken_config)
            self.assertIn("out of range", str(ctx.exception))

    def test_missing_template_file_is_fatal_at_load(self):
        real_config = get_tenant_config("default")
        missing_config = TenantConfig(
            tenant_id="default_missing",
            template_path=Path("does_not_exist.pptx"),
            manifest_path=real_config.manifest_path,
        )
        with self.assertRaises(ManifestValidationError):
            load_tenant_assets_from_config(missing_config)


if __name__ == "__main__":
    unittest.main()

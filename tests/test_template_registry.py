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


class TestSecondTenant(unittest.TestCase):
    """M7 (SDD v1.9): a second tenant ("meridian") with a different template,
    manifest, brand, shape-naming convention, and slide ordering, proving the
    engine is genuinely tenant-agnostic (PRD FR-11/FR-12) rather than
    implicitly coupled to the reference tenant's specifics."""

    def test_meridian_manifest_loads_and_validates(self):
        assets = load_tenant_assets("meridian")
        self.assertEqual(assets.tenant_id, "meridian")
        self.assertEqual(len(assets.manifest.layouts), 8)
        layout_ids = {l.layout_id for l in assets.manifest.layouts}
        # same layout_id contract as "default" (CONTENT_LAYOUT_IDS + bookends)...
        self.assertEqual(layout_ids, {
            "title", "closing_contact", "section_divider", "text_bullets",
            "image_only", "two_column", "exhibit_data", "quote_callout",
        })

    def test_meridian_uses_a_different_shape_naming_convention(self):
        assets = load_tenant_assets("meridian")
        title = assets.manifest.layout_by_id("title")
        shape_names = {s.shape_name for s in title.slots}
        self.assertTrue(all(name.startswith("MP_") for name in shape_names))
        self.assertFalse(any(name.startswith("IM_") for name in shape_names))

    def test_meridian_slide_index_ordering_differs_from_default(self):
        # deliberately shuffled at authoring time (tools/build_meridian_template.py)
        # -- proves layout_id -> slide_index is manifest-driven, not assumed fixed.
        default_title_idx = load_tenant_assets("default").manifest.layout_by_id("title").slide_index
        meridian_title_idx = load_tenant_assets("meridian").manifest.layout_by_id("title").slide_index
        self.assertNotEqual(default_title_idx, meridian_title_idx)

    def test_meridian_brand_colors_and_fonts_differ_from_default(self):
        default_brand = load_tenant_assets("default").manifest.brand
        meridian_brand = load_tenant_assets("meridian").manifest.brand
        self.assertNotEqual(set(default_brand["colors"]), set(meridian_brand["colors"]))
        self.assertNotEqual(default_brand["fonts"]["display"], meridian_brand["fonts"]["display"])

    def test_loading_one_tenant_does_not_leak_into_the_other(self):
        default_assets = load_tenant_assets("default")
        meridian_assets = load_tenant_assets("meridian")
        self.assertNotEqual(default_assets.manifest.template_name, meridian_assets.manifest.template_name)
        # opening one tenant's template must never return the other's slide count/shape names
        default_prs = default_assets.open_template()
        meridian_prs = meridian_assets.open_template()
        self.assertNotEqual(len(default_prs.slides), 0)
        default_shape_names = {s.name for slide in default_prs.slides for s in slide.shapes}
        meridian_shape_names = {s.name for slide in meridian_prs.slides for s in slide.shapes}
        self.assertTrue(default_shape_names.isdisjoint(meridian_shape_names))


if __name__ == "__main__":
    unittest.main()

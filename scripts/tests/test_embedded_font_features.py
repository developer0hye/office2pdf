"""Check resolved features, since merely skipping font loading still ships assets."""

import pathlib
import subprocess
import unittest


ROOT = pathlib.Path(__file__).resolve().parents[2]


class EmbeddedFontFeatures(unittest.TestCase):
    def assert_fonts(
        self, package: str, target: str, enabled: bool, *flags: str
    ) -> None:
        output = subprocess.check_output(
            [
                "cargo", "tree", "--locked", "-p", package,
                "--target", target, "-e", "normal", "--prefix", "none",
                "--format", "{p} features={f}", *flags,
            ],
            cwd=ROOT,
            text=True,
        )
        assets = [line for line in output.splitlines() if line.startswith("typst-assets ")]
        self.assertTrue(assets, "Typst still uses non-font assets")
        for line in assets:
            features = line.split("features=", 1)[1].split(" ", 1)[0].split(",")
            self.assertEqual("fonts" in features, enabled, line)

    def test_native_defaults_include_fonts(self) -> None:
        for package in ("office2pdf", "office2pdf-cli"):
            with self.subTest(package=package):
                self.assert_fonts(package, "aarch64-apple-darwin", True)

    def test_native_opt_out_excludes_fonts(self) -> None:
        for package in ("office2pdf", "office2pdf-cli"):
            for target in (
                "aarch64-apple-darwin",
                "x86_64-unknown-linux-gnu",
                "x86_64-pc-windows-msvc",
            ):
                with self.subTest(package=package, target=target):
                    self.assert_fonts(package, target, False, "--no-default-features")

    def test_explicit_opt_in_includes_fonts(self) -> None:
        for package in ("office2pdf", "office2pdf-cli"):
            with self.subTest(package=package):
                self.assert_fonts(package, "aarch64-apple-darwin", True,
                                  "--no-default-features", "--features", "embedded-fonts")

    def test_wasm_retains_fonts_without_default_features(self) -> None:
        self.assert_fonts("office2pdf", "wasm32-unknown-unknown", True, "--no-default-features")


if __name__ == "__main__":
    unittest.main()

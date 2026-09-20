"""Contract tests for reproducible workspace dependency resolution."""

from __future__ import annotations

import re
import subprocess
import tomllib
import unittest
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parents[2]
LOCKFILE = PROJECT_ROOT / "Cargo.lock"
RESOLVING_COMMAND = re.compile(
    r"\b(?:cargo (?:build|check|clippy|metadata|test)|wasm-pack (?:build|test))\b"
)


class DependencyLockContractTest(unittest.TestCase):
    def test_workspace_lockfile_is_tracked_and_not_ignored(self) -> None:
        self.assertTrue(LOCKFILE.is_file(), "the workspace Cargo.lock must exist")

        tracked = subprocess.run(
            ["git", "ls-files", "--error-unmatch", "Cargo.lock"],
            cwd=PROJECT_ROOT,
            capture_output=True,
            text=True,
            check=False,
        )
        self.assertEqual(tracked.returncode, 0, tracked.stderr)

        ignored = subprocess.run(
            ["git", "check-ignore", "--no-index", "--quiet", "Cargo.lock"],
            cwd=PROJECT_ROOT,
            check=False,
        )
        self.assertNotEqual(ignored.returncode, 0, "Cargo.lock must not be ignored")

    def test_branch_patches_have_reviewable_exact_revisions(self) -> None:
        manifest = tomllib.loads((PROJECT_ROOT / "Cargo.toml").read_text())
        lock = tomllib.loads(LOCKFILE.read_text())

        for name, patch in manifest["patch"]["crates-io"].items():
            if "branch" not in patch:
                continue

            sources = [
                package.get("source", "")
                for package in lock["package"]
                if package["name"] == name
                and package.get("source", "").startswith(f"git+{patch['git']}")
            ]
            self.assertEqual(len(sources), 1, f"expected one locked source for {name}")
            self.assertRegex(
                sources[0],
                r"#[0-9a-f]{40}$",
                f"{name} must expose its exact Git revision in Cargo.lock",
            )

    def test_workflow_dependency_commands_refuse_lockfile_updates(self) -> None:
        for relative_path in (
            ".github/workflows/ci.yml",
            ".github/workflows/release.yml",
        ):
            for line_number, line in enumerate(
                (PROJECT_ROOT / relative_path).read_text().splitlines(), start=1
            ):
                command = line.strip()
                if command.startswith(("#", "echo ")):
                    continue
                if RESOLVING_COMMAND.search(command):
                    self.assertIn(
                        "--locked",
                        command,
                        f"{relative_path}:{line_number} can update Cargo.lock: {command}",
                    )

    def test_publish_verification_keeps_checking_the_consumer_graph(self) -> None:
        workflow = (PROJECT_ROOT / ".github/workflows/ci.yml").read_text()
        without_lock = workflow.index("rm Cargo.lock")
        publish = workflow.index("cargo publish --dry-run -p office2pdf --allow-dirty")
        self.assertLess(without_lock, publish)
        self.assertNotIn("cargo publish --locked --dry-run", workflow)

    def test_documented_visual_builds_use_the_reviewed_resolution(self) -> None:
        agent_rules = (PROJECT_ROOT / "AGENTS.md").read_text()
        business = (PROJECT_ROOT / "tests/golden_mocks/business/README.md").read_text()
        readme = (PROJECT_ROOT / "README.md").read_text()

        self.assertIn(
            "cargo test --locked -p office2pdf --test artifact_generator",
            agent_rules,
        )
        self.assertIn("cargo build --locked -p office2pdf-cli", business)
        wasm_builds = [
            line
            for line in readme.splitlines()
            if line.startswith("wasm-pack build crates/office2pdf")
        ]
        self.assertEqual(len(wasm_builds), 2)
        self.assertTrue(all("--locked" in line for line in wasm_builds))


if __name__ == "__main__":
    unittest.main()

from __future__ import annotations

import json
import shutil
import unittest
import uuid
from pathlib import Path

from Scripts.queue_board import QueueStore
from Scripts.security_settings import (
    AppSettingsStore,
    normalize_allowed_domains,
    validate_link,
)


class QueueStoreTests(unittest.TestCase):
    def setUp(self):
        self.root = Path.cwd() / "tests" / "_runtime" / uuid.uuid4().hex
        self.root.mkdir(parents=True)
        self.data_file = self.root / "Data" / "queue_board.json"

    def tearDown(self):
        shutil.rmtree(self.root)

    def test_add_reorder_move_and_reload(self):
        store = QueueStore(self.data_file)
        first_id = store.add("general", "First", "#112233")
        linked_file = (self.root / "reference.pdf").resolve()
        second_id = store.add("general", "Second", "#ABCDEF", str(linked_file))
        self.assertTrue(store.reorder("general", second_id, 0))
        store.update(
            first_id,
            column="weekly",
            name="First weekly",
            color="#445566",
            link=str((self.root / "document.xlsx").resolve()),
        )
        store.save()

        loaded = QueueStore(self.data_file)
        loaded.load()
        self.assertEqual(
            [brick["name"] for brick in loaded.columns["general"]], ["Second"]
        )
        self.assertEqual(
            [brick["name"] for brick in loaded.columns["weekly"]], ["First weekly"]
        )
        self.assertEqual(loaded.columns["general"][0]["link"], str(linked_file))

    def test_moves_bricks_between_columns_at_requested_position(self):
        store = QueueStore(self.data_file)
        first = store.add("general", "First", "#112233")
        store.add("weekly", "Existing", "#445566")

        self.assertTrue(store.move(first, "weekly", 0))
        self.assertEqual(store.columns["general"], [])
        self.assertEqual(
            [brick["name"] for brick in store.columns["weekly"]],
            ["First", "Existing"],
        )
        self.assertFalse(store.move(first, "weekly", 0))
        self.assertEqual(
            [brick["name"] for brick in store.columns["weekly"]],
            ["First", "Existing"],
        )

    def test_save_uses_expected_json_shape(self):
        store = QueueStore(self.data_file)
        store.add("monthly", "Month end", "#F2C94C")
        store.save()

        payload = json.loads(self.data_file.read_text(encoding="utf-8"))
        self.assertEqual(payload["version"], 1)
        self.assertEqual(payload["columns"]["monthly"][0]["name"], "Month end")
        self.assertEqual(payload["columns"]["archive"], [])

    def test_archives_only_rejected_links_and_restores_after_edit(self):
        self.data_file.parent.mkdir(parents=True)
        payload = {
            "version": 1,
            "columns": {
                "general": [
                    {
                        "id": "good",
                        "name": "Allowed sheet",
                        "color": "#112233",
                        "link": "https://docs.google.com/spreadsheets/d/example/edit",
                    },
                    {
                        "id": "blocked",
                        "name": "Drive folder",
                        "color": "#445566",
                        "link": "https://drive.google.com/drive/folders/example",
                    },
                ],
                "weekly": [],
                "monthly": [],
            },
        }
        self.data_file.write_text(json.dumps(payload), encoding="utf-8")

        store = QueueStore(self.data_file, ["docs.google.com"])
        store.load()
        self.assertEqual([item["id"] for item in store.columns["general"]], ["good"])
        self.assertEqual(
            [item["id"] for item in store.columns["archive"]],
            ["blocked"],
        )
        archived = store.columns["archive"][0]
        self.assertEqual(archived["source_column"], "general")
        self.assertIn("not allowed", archived["link_error"])
        self.assertEqual(store.archived_during_load, 1)
        store.save()

        restored = QueueStore(
            self.data_file,
            ["docs.google.com", "drive.google.com"],
        )
        restored.load()
        restored.update(
            "blocked",
            column="weekly",
            name="Drive folder",
            color="#445566",
            link="https://drive.google.com/drive/folders/example",
        )
        self.assertEqual(restored.columns["archive"], [])
        self.assertEqual(restored.columns["weekly"][0]["id"], "blocked")
        self.assertNotIn("link_error", restored.columns["weekly"][0])

    def test_manual_archive_and_restore_preserve_original_stack(self):
        store = QueueStore(self.data_file)
        brick_id = store.add("monthly", "Month end", "#F2C94C")

        store.archive(brick_id)
        self.assertEqual(store.columns["monthly"], [])
        self.assertEqual(store.columns["archive"][0]["source_column"], "monthly")
        self.assertEqual(
            store.columns["archive"][0]["link_error"],
            "Archived manually.",
        )

        restored_column = store.unarchive(brick_id)
        self.assertEqual(restored_column, "monthly")
        self.assertEqual(store.columns["archive"], [])
        self.assertEqual(store.columns["monthly"][0]["id"], brick_id)

    def test_rejects_relative_file_link(self):
        with self.assertRaisesRegex(ValueError, "absolute"):
            validate_link("reports/monthly.xlsx", ["docs.google.com"])

    def test_allows_google_workspace_links_on_allowed_docs_domain(self):
        links = [
            "https://docs.google.com/spreadsheets/d/example/edit",
            "https://docs.google.com/document/d/example/edit",
            "https://docs.google.com/presentation/d/example/edit",
            "https://docs.google.com/forms/d/example/edit",
        ]
        for link in links:
            with self.subTest(link=link):
                self.assertEqual(validate_link(link, ["docs.google.com"]), link)

        with self.assertRaisesRegex(ValueError, "not allowed"):
            validate_link(
                "https://drive.google.com/drive/folders/example",
                ["docs.google.com"],
            )

    def test_rejects_unlisted_insecure_and_executable_links(self):
        with self.assertRaisesRegex(ValueError, "not allowed"):
            validate_link("https://example.com/report", ["docs.google.com"])
        with self.assertRaisesRegex(ValueError, "HTTPS"):
            validate_link("http://docs.google.com/spreadsheets/d/x", ["docs.google.com"])
        with self.assertRaisesRegex(ValueError, "blocked"):
            validate_link(
                str((self.root / "program.exe").resolve()),
                ["docs.google.com"],
            )
        with self.assertRaisesRegex(ValueError, "blocked"):
            validate_link(
                str((self.root / "legacy-macro-capable.xls").resolve()),
                ["docs.google.com"],
            )
        with self.assertRaisesRegex(ValueError, "Network-share"):
            validate_link(
                r"\\server\shared\report.xlsx",
                ["docs.google.com"],
            )

    def test_normalizes_domain_allowlist_and_settings(self):
        domains = normalize_allowed_domains(["Example.COM", "*.Example.org"])
        self.assertEqual(domains, ["example.com", "*.example.org"])
        settings = AppSettingsStore(self.root / "Data" / "settings.json")
        settings.allowed_domains = domains
        settings.save()

        loaded = AppSettingsStore(settings.settings_file)
        loaded.load()
        self.assertEqual(loaded.allowed_domains, domains)

    def test_corrupt_data_is_not_silently_replaced(self):
        self.data_file.parent.mkdir(parents=True)
        self.data_file.write_text("not-json", encoding="utf-8")
        store = QueueStore(self.data_file)

        with self.assertRaisesRegex(ValueError, "Could not read"):
            store.load()
        self.assertEqual(self.data_file.read_text(encoding="utf-8"), "not-json")


if __name__ == "__main__":
    unittest.main()

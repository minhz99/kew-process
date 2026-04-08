import os
import tempfile
import unittest
import zipfile

from PIL import Image

try:
    from generate_kew_synopex import build_synopex_report, etree
except Exception as exc:  # pragma: no cover - dependency/import guard
    build_synopex_report = None
    etree = None
    IMPORT_ERROR = exc
else:
    IMPORT_ERROR = None


@unittest.skipIf(build_synopex_report is None, f"Không import được generator: {IMPORT_ERROR}")
@unittest.skipIf(etree is None, "Thiếu lxml trong môi trường test")
class GenerateKewSynopexSmokeTest(unittest.TestCase):
    def test_build_report_from_minimal_image_set(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            machine_dir = os.path.join(temp_dir, "S0001 - Tu bom")
            os.makedirs(machine_dir, exist_ok=True)

            for filename in ["a1.png"] + [f"ps-sd{i}.png" for i in range(1, 7)]:
                image = Image.new("RGB", (240, 256), color=(255, 255, 255))
                image.save(os.path.join(machine_dir, filename))

            output_path = os.path.join(temp_dir, "report.docx")
            generated_path = build_synopex_report(base_dir=temp_dir, output_file=output_path)

            self.assertEqual(generated_path, output_path)
            self.assertTrue(os.path.exists(output_path))

            with zipfile.ZipFile(output_path, "r") as archive:
                names = set(archive.namelist())

            self.assertIn("word/document.xml", names)
            self.assertIn("word/styles.xml", names)
            self.assertIn("word/media/image1.png", names)


if __name__ == "__main__":
    unittest.main()

import io
import unittest
import zipfile
from pathlib import Path
from unittest import mock

from app import app


def make_zip_bytes(entries):
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as archive:
        for name, content in entries.items():
            archive.writestr(name, content)
    buffer.seek(0)
    return buffer


class SynopexApiTestCase(unittest.TestCase):
    def setUp(self):
        app.config["TESTING"] = True
        self.client = app.test_client()

    def test_generate_requires_zip(self):
        response = self.client.post("/api/synopex/generate", data={})

        self.assertEqual(response.status_code, 400)
        self.assertEqual(response.get_json()["error"], "Cần upload file ZIP dữ liệu.")

    def test_generate_rejects_invalid_zip(self):
        response = self.client.post(
            "/api/synopex/generate",
            data={
                "data_zip": (io.BytesIO(b"not-a-zip"), "input.zip"),
                "output_name": "report.docx",
            },
            content_type="multipart/form-data",
        )

        self.assertEqual(response.status_code, 400)
        self.assertEqual(response.get_json()["error"], "File ZIP dữ liệu không hợp lệ.")

    @mock.patch("generate_kew_synopex.build_synopex_report")
    def test_generate_returns_docx_file(self, build_report_mock):
        def fake_build_report(*, base_dir, output_file, template_file=None, tesseract_cmd=None):
            self.assertTrue(Path(base_dir).exists())
            Path(output_file).write_bytes(b"fake-docx")
            return output_file

        build_report_mock.side_effect = fake_build_report

        zip_buffer = make_zip_bytes(
            {
                "root/S0001 - May 1/a1.png": b"trend",
                "root/S0001 - May 1/ps-sd1.png": b"1",
                "root/S0001 - May 1/ps-sd2.png": b"2",
                "root/S0001 - May 1/ps-sd3.png": b"3",
                "root/S0001 - May 1/ps-sd4.png": b"4",
                "root/S0001 - May 1/ps-sd5.png": b"5",
                "root/S0001 - May 1/ps-sd6.png": b"6",
            }
        )

        response = self.client.post(
            "/api/synopex/generate",
            data={
                "data_zip": (zip_buffer, "input.zip"),
                "output_name": "Bao cao tu dong",
            },
            content_type="multipart/form-data",
        )

        self.assertEqual(response.status_code, 200)
        self.assertEqual(
            response.headers["Content-Type"],
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )
        self.assertIn("Bao cao tu dong.docx", response.headers["Content-Disposition"])
        self.assertEqual(response.data, b"fake-docx")


if __name__ == "__main__":
    unittest.main()

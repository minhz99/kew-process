import io
import unittest
import zipfile

from PIL import Image

from app import app

try:
    from generate_kew_synopex import etree
except Exception as exc:  # pragma: no cover - dependency/import guard
    etree = None
    IMPORT_ERROR = exc
else:
    IMPORT_ERROR = None


def make_image_zip_bytes():
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as archive:
        for filename in ["a1.png"] + [f"ps-sd{i}.png" for i in range(1, 7)]:
            image_buffer = io.BytesIO()
            image = Image.new("RGB", (240, 256), color=(255, 255, 255))
            image.save(image_buffer, format="PNG")
            archive.writestr(f"batch/S0001 - Tu bom/{filename}", image_buffer.getvalue())
    buffer.seek(0)
    return buffer


@unittest.skipIf(etree is None, f"Thiếu lxml hoặc import generator lỗi: {IMPORT_ERROR}")
class SynopexEndToEndTestCase(unittest.TestCase):
    def setUp(self):
        app.config["TESTING"] = True
        self.client = app.test_client()

    def test_generate_endpoint_builds_docx(self):
        response = self.client.post(
            "/api/synopex/generate",
            data={
                "data_zip": (make_image_zip_bytes(), "input.zip"),
                "output_name": "Ket qua synopex",
            },
            content_type="multipart/form-data",
        )

        self.assertEqual(response.status_code, 200)
        self.assertEqual(
            response.headers["Content-Type"],
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )
        self.assertIn("Ket qua synopex.docx", response.headers["Content-Disposition"])

        with zipfile.ZipFile(io.BytesIO(response.data), "r") as archive:
            names = set(archive.namelist())

        self.assertIn("word/document.xml", names)
        self.assertIn("word/styles.xml", names)
        self.assertIn("word/media/image1.png", names)


if __name__ == "__main__":
    unittest.main()

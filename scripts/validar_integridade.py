from __future__ import annotations

import json
import shutil
import sys
import tempfile
import zipfile
from datetime import date
from decimal import Decimal
from hashlib import sha256
from io import BytesIO
from pathlib import Path

from PIL import Image

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

import app  # noqa: E402


class FakePackage:
    def __init__(self, data: bytes) -> None:
        self.data = data

    def getbuffer(self) -> memoryview:
        return memoryview(self.data)


def image_bytes() -> bytes:
    buffer = BytesIO()
    Image.new("RGB", (640, 900), "white").save(buffer, "JPEG")
    return buffer.getvalue()


def package_bytes(*, corrupt_hash: bool = False, duplicate_photo: bool = False, invalid_image: bool = False) -> bytes:
    first = b"imagem-invalida" if invalid_image else image_bytes()
    second = image_bytes()
    first_name = "notas/001.jpg"
    second_name = first_name if duplicate_photo else "notas/002.jpg"
    expenses = [
        {
            "sequence": 1,
            "date": "2026-06-07",
            "category": "Refeições",
            "description": "Almoço",
            "amount": "25.00",
            "photo": first_name,
            "photo_size": len(first),
            "photo_sha256": "0" * 64 if corrupt_hash else sha256(first).hexdigest(),
        },
        {
            "sequence": 2,
            "date": "2026-06-07",
            "category": "Refeições",
            "description": "Jantar",
            "amount": "35.00",
            "photo": second_name,
            "photo_size": len(second),
            "photo_sha256": sha256(second).hexdigest(),
        },
    ]
    manifest = {
        "format": "relatorio-despesas-mobile",
        "version": 2,
        "expense_count": len(expenses),
        "expenses": expenses,
    }
    buffer = BytesIO()
    with zipfile.ZipFile(buffer, "w", compression=zipfile.ZIP_STORED) as archive:
        archive.writestr(app.MOBILE_PACKAGE_MANIFEST, json.dumps(manifest, ensure_ascii=False))
        archive.writestr(first_name, first)
        if second_name != first_name:
            archive.writestr(second_name, second)
    return buffer.getvalue()


def expect_rejected(data: bytes, expected_text: str) -> None:
    with tempfile.TemporaryDirectory(dir=app.UPLOAD_WORK_DIR) as temp_dir:
        try:
            app.import_mobile_package(FakePackage(data), Path(temp_dir))
        except ValueError as exc:
            if expected_text.lower() not in str(exc).lower():
                raise AssertionError(f"Erro inesperado: {exc}") from exc
        else:
            raise AssertionError(f"Pacote deveria ter sido rejeitado: {expected_text}")


def validate_package_integrity() -> None:
    app.UPLOAD_WORK_DIR.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory(dir=app.UPLOAD_WORK_DIR) as temp_dir:
        notes, detections = app.import_mobile_package(FakePackage(package_bytes()), Path(temp_dir))
        assert len(notes) == 2 and len(detections) == 2
    expect_rejected(package_bytes(corrupt_hash=True), "integridade")
    expect_rejected(package_bytes(duplicate_photo=True), "mais de uma despesa")
    expect_rejected(package_bytes(invalid_image=True), "corrompida")
    expect_rejected(b"nao-e-zip", "ZIP válido")


def validate_final_zip() -> None:
    with tempfile.TemporaryDirectory(dir=app.UPLOAD_WORK_DIR) as temp_dir:
        root = Path(temp_dir)
        first = root / "relatorio.xlsx"
        second = root / "notas.pdf"
        first.write_bytes(b"arquivo-excel")
        second.write_bytes(b"%PDF-1.4\narquivo-pdf")
        result = app.make_download_zip([first, second])
        with zipfile.ZipFile(BytesIO(result)) as archive:
            assert archive.testzip() is None
            assert sorted(archive.namelist()) == ["notas.pdf", "relatorio.xlsx"]


def validate_isolated_outputs() -> None:
    app.UPLOAD_WORK_DIR.mkdir(parents=True, exist_ok=True)
    output_dirs: list[Path] = []
    with tempfile.TemporaryDirectory(dir=app.UPLOAD_WORK_DIR) as temp_dir:
        notes, detections = app.import_mobile_package(FakePackage(package_bytes()), Path(temp_dir))
        expenses = app.detections_to_frame(detections, date(2026, 6, 7))
        for _ in range(2):
            result = app.generate_report(
                employee={"name": "", "cpf": "", "cost_center": "", "bank": "", "agency": "", "account": ""},
                trip={
                    "reason": "Teste de isolamento",
                    "start_date": date(2026, 6, 7),
                    "end_date": date(2026, 6, 7),
                    "report_date": date(2026, 6, 7),
                    "auto_dates": False,
                },
                base_name="MESMO_NOME",
                raw_expenses=expenses,
                raw_mileage=app.default_mileage(),
                km_rate=Decimal("0.80"),
                advance=Decimal("0.00"),
                note_map=notes,
            )
            output_dirs.append(Path(result["output_dir"]))
        assert output_dirs[0] != output_dirs[1]
        assert all((path / "MESMO_NOME.xlsx").exists() for path in output_dirs)
    for path in output_dirs:
        shutil.rmtree(path, ignore_errors=True)


def main() -> None:
    try:
        validate_package_integrity()
        print("OK - pacotes íntegros aceitos e pacotes corrompidos rejeitados")
        validate_final_zip()
        print("OK - ZIP final verificado após a geração")
        validate_isolated_outputs()
        print("OK - gerações simultâneas com o mesmo nome ficam isoladas")
    finally:
        shutil.rmtree(app.UPLOAD_WORK_DIR / "mobile_packages", ignore_errors=True)
    print("VALIDAÇÃO DE INTEGRIDADE CONCLUÍDA")


if __name__ == "__main__":
    main()

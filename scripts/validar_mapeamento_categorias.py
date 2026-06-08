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

from openpyxl import load_workbook
from PIL import Image, ImageDraw

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

import app  # noqa: E402


class FakePackage:
    def __init__(self, data: bytes) -> None:
        self._data = data

    def getbuffer(self) -> memoryview:
        return memoryview(self._data)


def note_bytes(label: str) -> bytes:
    image = Image.new("RGB", (900, 1200), "white")
    draw = ImageDraw.Draw(image)
    draw.text((60, 70), label, fill="black")
    buffer = BytesIO()
    image.save(buffer, "JPEG", quality=90)
    return buffer.getvalue()


def build_mobile_package() -> tuple[bytes, int]:
    expense_date = "2026-06-08"
    items = [
        ("alimentacao", "Almoco", "25.00"),
        (app.CATEGORY_ORDER[1], "Jantar", "35.00"),
        (app.CATEGORY_ORDER[0], "Hotel", "101.00"),
        (app.CATEGORY_ORDER[2], "Pedagio", "12.00"),
        (app.CATEGORY_ORDER[3], "Frete", "13.00"),
        (app.CATEGORY_ORDER[4], "Combustivel", "14.00"),
        (app.CATEGORY_ORDER[5], "Material", "15.00"),
        (app.CATEGORY_ORDER[6], "Locacao", "16.00"),
        (app.CATEGORY_ORDER[7], "Taxi", "17.00"),
        (app.CATEGORY_ORDER[8], "Lavacao de carro", "18.00"),
    ]

    images: dict[str, bytes] = {}
    expenses: list[dict[str, object]] = []
    for sequence, (category, description, amount) in enumerate(items, start=1):
        photo_name = f"notas/{sequence:03d}.jpg"
        image_data = note_bytes(f"{description} R$ {amount}")
        images[photo_name] = image_data
        expenses.append(
            {
                "sequence": sequence,
                "date": expense_date,
                "category": category,
                "description": description,
                "amount": amount,
                "photo": photo_name,
                "photo_size": len(image_data),
                "photo_sha256": sha256(image_data).hexdigest(),
            }
        )

    manifest = {
        "format": "relatorio-despesas-mobile",
        "version": 2,
        "expense_count": len(expenses),
        "expenses": expenses,
    }

    buffer = BytesIO()
    with zipfile.ZipFile(buffer, "w", compression=zipfile.ZIP_STORED) as archive:
        archive.writestr(app.MOBILE_PACKAGE_MANIFEST, json.dumps(manifest, ensure_ascii=False))
        for photo_name, image_data in images.items():
            archive.writestr(photo_name, image_data)
    return buffer.getvalue(), len(expenses)


def assert_number(value: object, expected: float, label: str) -> None:
    if value is None or round(float(value), 2) != expected:
        raise AssertionError(f"{label}: esperado {expected}, recebido {value!r}")


def assert_formula(value: object, expected: str, label: str) -> None:
    if value != expected:
        raise AssertionError(f"{label}: esperado {expected}, recebido {value!r}")


def assert_date(value: object, expected: date, label: str) -> None:
    loaded = value.date() if hasattr(value, "date") else value
    if loaded != expected:
        raise AssertionError(f"{label}: esperado {expected}, recebido {value!r}")


def main() -> None:
    package_data, expected_count = build_mobile_package()
    app.UPLOAD_WORK_DIR.mkdir(parents=True, exist_ok=True)
    output_dir: Path | None = None

    try:
        with tempfile.TemporaryDirectory(dir=app.UPLOAD_WORK_DIR) as temp_dir:
            note_map, detections = app.import_mobile_package(FakePackage(package_data), Path(temp_dir))
            if len(note_map) != expected_count or len(detections) != expected_count:
                raise AssertionError("Nem todas as notas foram importadas do pacote movel.")
            if any(not path.exists() or path.stat().st_size <= 0 for path in note_map.values()):
                raise AssertionError("Alguma foto importada esta ausente ou vazia.")

            expenses = app.detections_to_frame(detections, date(2026, 6, 8))
            normalized = app.normalize_expenses(expenses)
            if normalized.empty or len(normalized) != expected_count:
                raise AssertionError("A normalizacao perdeu despesas validas.")
            if normalized.loc[normalized["descricao"] == "Almoco", "categoria"].iloc[0] != app.CATEGORY_ORDER[1]:
                raise AssertionError("Alimentacao nao foi normalizada para Refeicoes.")
            if normalized.loc[normalized["descricao"] == "Hotel", "categoria"].iloc[0] != app.CATEGORY_ORDER[0]:
                raise AssertionError("Hospedagem nao permaneceu em Hospedagem.")

            result = app.generate_report(
                employee={"name": "", "cpf": "", "cost_center": "", "bank": "", "agency": "", "account": ""},
                trip={
                    "reason": "Validacao de mapeamento",
                    "start_date": date(2026, 6, 8),
                    "end_date": date(2026, 6, 8),
                    "report_date": date(2026, 6, 8),
                    "auto_dates": False,
                },
                base_name="VALIDACAO_MAPEAMENTO_CATEGORIAS",
                raw_expenses=expenses,
                raw_mileage=app.default_mileage(),
                km_rate=Decimal("0.80"),
                advance=Decimal("0.00"),
                note_map=note_map,
            )
            output_dir = Path(result["output_dir"])

            xlsx_path = Path(result["xlsx"])
            pdf_path = Path(result["notes_pdf"])
            if not zipfile.is_zipfile(xlsx_path):
                raise AssertionError("Excel gerado nao e um arquivo XLSX valido.")
            if not pdf_path.read_bytes().startswith(b"%PDF-"):
                raise AssertionError("PDF de notas gerado esta invalido.")

            workbook = load_workbook(xlsx_path, data_only=False)
            try:
                worksheet = workbook.worksheets[0]
                assert_date(worksheet["D19"].value, date(2026, 6, 8), "Data da coluna de despesas")
                assert_number(worksheet["D21"].value, 101.00, "Hospedagem deve preencher D21")
                assert_formula(worksheet["D24"].value, "=25.00+35.00", "Refeicoes devem preencher D24 com formula")
                assert_number(worksheet["D27"].value, 12.00, "Pedagio deve preencher D27")
                assert_number(worksheet["D30"].value, 13.00, "Frete deve preencher D30")
                assert_number(worksheet["D33"].value, 14.00, "Combustivel deve preencher D33")
                assert_number(worksheet["D36"].value, 15.00, "Material deve preencher D36")
                assert_number(worksheet["D39"].value, 16.00, "Locacoes deve preencher D39")
                assert_number(worksheet["D42"].value, 17.00, "Transporte deve preencher D42")
                assert_number(worksheet["D45"].value, 18.00, "Outras despesas deve preencher D45")
                assert_date(worksheet["B51"].value, date(2026, 6, 8), "Outras despesas deve detalhar data em B51")
                if worksheet["D51"].value != "Lavacao de carro":
                    raise AssertionError(f"Descricao de outras despesas inesperada: {worksheet['D51'].value!r}")
                assert_number(worksheet["X51"].value, 18.00, "Outras despesas deve detalhar valor em X51")
            finally:
                workbook.close()

            final_zip = app.make_download_zip([xlsx_path, pdf_path])
            with zipfile.ZipFile(BytesIO(final_zip)) as archive:
                if archive.testzip() is not None:
                    raise AssertionError("ZIP final esta corrompido.")
                expected_files = sorted([xlsx_path.name, pdf_path.name])
                if sorted(archive.namelist()) != expected_files:
                    raise AssertionError(f"ZIP final contem arquivos inesperados: {archive.namelist()}")

        print("OK - alimentacao/refeicoes preencheram a linha correta no Excel")
        print("OK - hospedagem, pedagio, frete, combustivel, material, locacoes, transporte e outras despesas conferidos")
        print("OK - todas as fotos do coletor foram importadas e incluidas no fluxo final")
        print("VALIDACAO DE MAPEAMENTO DE CATEGORIAS CONCLUIDA")
    finally:
        if output_dir is not None:
            shutil.rmtree(output_dir, ignore_errors=True)
        shutil.rmtree(app.UPLOAD_WORK_DIR / "mobile_packages", ignore_errors=True)


if __name__ == "__main__":
    main()

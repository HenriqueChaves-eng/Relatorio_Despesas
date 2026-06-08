from __future__ import annotations

import json
import shutil
import sys
import tempfile
import zipfile
from datetime import date, datetime
from decimal import Decimal
from io import BytesIO
from pathlib import Path

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
    image = Image.new("RGB", (1000, 1400), "white")
    draw = ImageDraw.Draw(image)
    draw.text((80, 100), label, fill="black")
    buffer = BytesIO()
    image.save(buffer, "JPEG")
    return buffer.getvalue()


def create_package() -> bytes:
    first_photo = "notas/001_2026-05-23_REFEICOES.jpg"
    second_photo = "notas/002_2026-05-23_REFEICOES.jpg"
    manifest = {
        "format": "relatorio-despesas-mobile",
        "version": 1,
        "expenses": [
            {
                "sequence": 1,
                "date": "2026-05-23",
                "category": app.CATEGORY_ORDER[1],
                "description": "Almoco",
                "amount": "50.00",
                "photo": first_photo,
            },
            {
                "sequence": 2,
                "date": "2026-05-23",
                "category": app.CATEGORY_ORDER[1],
                "description": "Jantar",
                "amount": "30.00",
                "photo": second_photo,
            },
        ],
    }
    buffer = BytesIO()
    with zipfile.ZipFile(buffer, "w", compression=zipfile.ZIP_STORED) as archive:
        archive.writestr(app.MOBILE_PACKAGE_MANIFEST, json.dumps(manifest, ensure_ascii=False))
        archive.writestr(first_photo, note_bytes("ALMOCO R$ 50,00"))
        archive.writestr(second_photo, note_bytes("JANTAR R$ 30,00"))
    return buffer.getvalue()


def main() -> None:
    base_name = f"VALIDACAO_FLUXO_MOBILE_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
    app.UPLOAD_WORK_DIR.mkdir(exist_ok=True)
    generated_dir: Path | None = None

    try:
        with tempfile.TemporaryDirectory(dir=app.UPLOAD_WORK_DIR) as temp_dir:
            note_map, detections = app.import_mobile_package(FakePackage(create_package()), Path(temp_dir))
            expenses = app.detections_to_frame(detections, date(2026, 5, 23))
            result = app.generate_report(
                employee={
                    "name": "Usuario Teste",
                    "cpf": "",
                    "cost_center": "1600",
                    "bank": "",
                    "agency": "",
                    "account": "",
                },
                trip={
                    "reason": "Validacao coletor mobile",
                    "start_date": date(2026, 5, 23),
                    "end_date": date(2026, 5, 23),
                    "report_date": date(2026, 6, 6),
                    "auto_dates": False,
                },
                base_name=base_name,
                raw_expenses=expenses,
                raw_mileage=app.default_mileage(),
                km_rate=Decimal("0.80"),
                advance=Decimal("0.00"),
                note_map=note_map,
            )

        generated_dir = Path(result["output_dir"])
        final_files = sorted(path.name for path in generated_dir.iterdir() if path.is_file())
        expected = [f"{base_name}.xlsx", f"{base_name}_NOTAS.pdf"]
        if final_files != expected:
            raise AssertionError(f"Arquivos finais inesperados: {final_files}")
        if result["missing_notes"]:
            raise AssertionError(f"Notas ausentes: {result['missing_notes']}")

        print("OK - pacote movel importado")
        print("OK - Excel preenchido gerado")
        print("OK - PDF das notas gerado em sequencia")
        print(f"OK - somente dois arquivos finais: {final_files}")
    finally:
        if generated_dir is not None:
            shutil.rmtree(generated_dir, ignore_errors=True)
        shutil.rmtree(app.UPLOAD_WORK_DIR / "mobile_packages", ignore_errors=True)


if __name__ == "__main__":
    main()

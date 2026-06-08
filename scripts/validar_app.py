from __future__ import annotations

import json
import sys
import tempfile
import time
import zipfile
from datetime import date
from decimal import Decimal
from io import BytesIO
from pathlib import Path

import pandas as pd
from PIL import Image

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

import app  # noqa: E402


def check(condition: bool, message: str) -> None:
    if not condition:
        raise AssertionError(message)


def validate_text_and_parsers() -> None:
    check(app.repair_text_encoding("RefeiÃ§Ãµes") == "Refeições", "reparo de acentuação falhou")
    check(app.slugify("Relatório São Domingos") == "RELATORIO_SAO_DOMINGOS", "slugify falhou")
    check(app.parse_money("R$ 1.234,56") == Decimal("1234.56"), "parse_money BR falhou")
    check(app.parse_money("89.9") == Decimal("89.90"), "parse_money decimal falhou")
    check(
        app.parse_detected_date("EMISSAO: 29/04/2026", date(2026, 4, 1), date(2026, 4, 30))
        == date(2026, 4, 29),
        "data numerica nao detectada",
    )
    check(
        app.parse_detected_amount("VALOR TOTAL R$ 89,90\nTROCO R$ 0,00") == Decimal("89.90"),
        "valor total nao detectado",
    )
    check(app.normalize_ai_category("alimentacao") == "Refeições", "categoria alimentacao falhou")
    check(app.normalize_ai_category("uber") == "Transporte/ taxi", "categoria transporte falhou")
    check(app.normalize_ai_category("lavação de carro") == "Outras Despesas", "categoria lavacao falhou")
    check(
        app.infer_category("LAVACAO DE CARRO VALOR TOTAL R$ 120,00", "nota_lavacao.jpg") == "Outras Despesas",
        "lavacao de carro foi confundida com locacao",
    )
    check(app.excel_sum_formula([Decimal("50.00"), Decimal("70.00")]) == "=50.00+70.00", "formula excel falhou")
    check(".heic" in app.IMAGE_EXTENSIONS and ".heif" in app.IMAGE_EXTENSIONS, "suporte iPad HEIC/HEIF ausente")


def validate_ai_schema_and_detection() -> None:
    schema = app.receipt_ai_json_schema()
    audit_schema = app.receipt_audit_json_schema()
    required = set(schema["required"])
    check({"expense_date", "amount", "category", "description", "confidence"}.issubset(required), "schema IA incompleto")
    check("approved" in audit_schema["properties"], "schema de auditoria sem approved")

    with tempfile.TemporaryDirectory(dir=ROOT / "saida") as temp_dir:
        image_path = Path(temp_dir) / "nota.jpg"
        Image.new("RGB", (1200, 1600), "white").save(image_path, "JPEG")
        detection = app.note_detection_from_ai(
            note_name="nota.jpg",
            note_path=image_path,
            start_date=date(2026, 4, 1),
            end_date=date(2026, 4, 30),
            payload={
                "expense_date": "2026-04-29",
                "amount": 42.75,
                "category": "refeicao",
                "description": "Restaurante teste",
                "confidence": "alta",
                "warnings": [],
                "evidence": "VALOR TOTAL R$ 42,75",
            },
            source="IA Google AI Studio",
        )

    check(detection.expense_date == date(2026, 4, 29), "deteccao IA perdeu a data")
    check(detection.value == Decimal("42.75"), "deteccao IA perdeu o valor")
    check(detection.category == "Refeições", "deteccao IA perdeu a categoria")
    check(detection.source == "IA Google AI Studio", "deteccao IA perdeu a origem")


def validate_upload_signature_stability() -> None:
    class FakeUpload:
        name = "nota_teste.jpg"

        def __init__(self, data: bytes) -> None:
            self._data = data

        def getbuffer(self) -> memoryview:
            return memoryview(self._data)

    image_bytes = b"fake-image-content-for-upload-stability"
    with tempfile.TemporaryDirectory(dir=ROOT / "saida") as temp_dir:
        work_dir = Path(temp_dir)
        note_map_1 = {note.name: note.path for note in app.save_uploaded_notes([FakeUpload(image_bytes)], work_dir)}
        signature_1 = app.note_signature(note_map_1)
        note_path = next(iter(note_map_1.values()))
        mtime_1 = note_path.stat().st_mtime_ns
        time.sleep(0.05)
        note_map_2 = {note.name: note.path for note in app.save_uploaded_notes([FakeUpload(image_bytes)], work_dir)}
        signature_2 = app.note_signature(note_map_2)
        mtime_2 = next(iter(note_map_2.values())).stat().st_mtime_ns

    check(signature_1 == signature_2, "assinatura de upload mudou em rerun sem alteracao")
    check(mtime_1 == mtime_2, "upload identico foi regravado e pode apagar a revisao")


def validate_mobile_package_import() -> None:
    image_buffer = BytesIO()
    Image.new("RGB", (800, 1200), "white").save(image_buffer, "JPEG")
    manifest = {
        "format": "relatorio-despesas-mobile",
        "version": 1,
        "expenses": [
            {
                "sequence": 1,
                "date": "2026-05-23",
                "category": "Refeições",
                "description": "Almoco",
                "amount": "50.00",
                "photo": "notas/001_2026-05-23_REFEICOES.jpg",
            },
            {
                "sequence": 2,
                "date": "2026-05-23",
                "category": "Refeições",
                "description": "Jantar",
                "amount": "30.00",
                "photo": "notas/002_2026-05-23_REFEICOES.jpg",
            },
        ],
    }
    package_buffer = BytesIO()
    with zipfile.ZipFile(package_buffer, "w", compression=zipfile.ZIP_STORED) as archive:
        archive.writestr(app.MOBILE_PACKAGE_MANIFEST, json.dumps(manifest, ensure_ascii=False))
        archive.writestr("notas/001_2026-05-23_REFEICOES.jpg", image_buffer.getvalue())
        archive.writestr("notas/002_2026-05-23_REFEICOES.jpg", image_buffer.getvalue())

    class FakePackage:
        def getbuffer(self) -> memoryview:
            return memoryview(package_buffer.getvalue())

    with tempfile.TemporaryDirectory(dir=ROOT / "saida") as temp_dir:
        note_map, detections = app.import_mobile_package(FakePackage(), Path(temp_dir))
        check(len(note_map) == 2, "pacote movel nao importou as fotos")
        check(len(detections) == 2, "pacote movel nao importou as despesas")
        check(detections[0].value == Decimal("50.00"), "pacote movel perdeu o primeiro valor")
        check(detections[1].value == Decimal("30.00"), "pacote movel perdeu o segundo valor")
        check(detections[0].source == "Coletor móvel", "pacote movel perdeu a origem manual")


def validate_review_rules() -> None:
    frame = pd.DataFrame(
        [
            {
                "usar": True,
                "data": date(2026, 4, 29),
                "categoria": "Refeições",
                "descricao": "Almoco",
                "valor": 50.0,
                "nota": "nota-1.jpg",
                "confianca": "Alta",
                "origem": "OCR local",
                "alertas": "",
            },
            {
                "usar": True,
                "data": date(2026, 4, 29),
                "categoria": "Refeições",
                "descricao": "Almoco duplicado",
                "valor": 50.0,
                "nota": "nota-1.jpg",
                "confianca": "Alta",
                "origem": "IA Google AI Studio",
                "alertas": "",
            },
            {
                "usar": True,
                "data": "data ruim",
                "categoria": "Categoria inexistente",
                "descricao": "Linha invalida",
                "valor": 0,
                "nota": "",
                "confianca": "Baixa",
                "origem": "IA indisponivel",
                "alertas": "",
            },
        ]
    )
    reviewed = app.add_review_alerts(
        frame,
        note_map={"nota-1.jpg": ROOT / "entrada_notas" / "nota-1.jpg"},
        start_date=date(2026, 4, 1),
        end_date=date(2026, 4, 30),
    )
    blockers, warnings = app.review_alerts(reviewed)
    blocker_text = "\n".join(blockers)
    warning_text = "\n".join(warnings)
    check("Nota usada em mais de uma linha" in blocker_text, "duplicidade de nota nao bloqueou")
    check("Valor zerado" in blocker_text, "valor zerado nao bloqueou")
    check("Data invalida" in blocker_text, "data invalida nao bloqueou")
    check("Categoria invalida" in blocker_text, "categoria invalida nao bloqueou")
    check("IA indisponivel" in blocker_text, "IA indisponivel nao bloqueou")
    check("Leitura por OCR local" in warning_text, "OCR local nao gerou aviso")


def validate_payload_and_limits() -> None:
    expenses = app.normalize_expenses(
        pd.DataFrame(
            [
                {
                    "data": date(2026, 4, 29),
                    "categoria": "Refeições",
                    "descricao": "Almoco",
                    "valor": "R$ 50,00",
                    "nota": "nota-1.jpg",
                },
                {
                    "data": date(2026, 4, 29),
                    "categoria": "Refeições",
                    "descricao": "Jantar",
                    "valor": "70,00",
                    "nota": "nota-3.jpg",
                },
                {
                    "data": date(2026, 4, 30),
                    "categoria": "Pedágio",
                    "descricao": "Pedagio",
                    "valor": "12,40",
                    "nota": "nota-2.jpg",
                },
            ]
        )
    )
    payload = app.build_payload(
        output_dir=ROOT / "saida" / "VALIDACAO_APP",
        base_name="VALIDACAO_APP",
        employee={"name": "Henrique", "department": "Teste"},
        trip={
            "start_date": date(2026, 4, 29),
            "end_date": date(2026, 4, 30),
            "report_date": date(2026, 5, 19),
            "reason": "Teste",
        },
        expenses=expenses,
        mileage=[],
        km_rate=Decimal("1.00"),
        advance=Decimal("0.00"),
    )
    check(len(payload["dates"]) == 2, "payload nao preservou as datas")
    check(len(payload["expense_cells"]) == 2, "payload nao criou as celulas de despesa")
    meal_cell = next(cell for cell in payload["expense_cells"] if cell["category"] == "Refeicoes")
    check(meal_cell["value"] == 120.0, "payload nao somou refeicoes do dia")
    check(meal_cell["formula"] == "=50.00+70.00", "payload nao gerou formula de composicao")
    check(payload["trip"]["start_date"] == "2026-04-29", "payload perdeu data inicial")

    too_many_dates = pd.DataFrame(
        [{"data": date(2026, 4, day), "categoria": "Refeições"} for day in range(1, app.MAX_DATES + 2)]
    )
    limit_alerts = app.report_limit_alerts(too_many_dates)
    check(any("datas diferentes" in alert for alert in limit_alerts), "limite de datas nao gerou alerta")


def main() -> None:
    validations = [
        validate_text_and_parsers,
        validate_ai_schema_and_detection,
        validate_upload_signature_stability,
        validate_mobile_package_import,
        validate_review_rules,
        validate_payload_and_limits,
    ]
    for validation in validations:
        validation()
        print(f"OK - {validation.__name__}")
    print("VALIDACAO CONCLUIDA: app.py passou nos testes locais.")


if __name__ == "__main__":
    main()

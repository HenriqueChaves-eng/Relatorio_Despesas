from __future__ import annotations

import base64
import hmac
import html
import json
import os
import re
import shutil
import subprocess
import tempfile
import time
import unicodedata
import urllib.error
import urllib.request
import zipfile
from collections import defaultdict
from dataclasses import dataclass
from datetime import date, datetime
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from hashlib import sha1, sha256
from io import BytesIO
from pathlib import Path
from typing import Any, Iterable
from uuid import uuid4

import pandas as pd
import streamlit as st
from fpdf import FPDF
from PIL import Image, ImageEnhance, ImageFilter, ImageOps, ImageStat
from excel_generator import generate_excel_report

try:
    from pillow_heif import register_heif_opener
except Exception:
    register_heif_opener = None

if register_heif_opener is not None:
    register_heif_opener()


ROOT = Path(__file__).resolve().parent
TEMPLATE_PATH = ROOT / "templates" / "modelo_relatorio_despesas.xlsx"
OCR_SCRIPT_PATH = ROOT / "scripts" / "ocr_windows.ps1"
OUTPUT_ROOT = ROOT / "saida"
NOTES_INPUT_DIR = ROOT / "entrada_notas"
UPLOAD_WORK_DIR = ROOT / ".streamlit_uploads"
AI_CACHE_DIR = UPLOAD_WORK_DIR / "ai_cache"
AI_MEMORY_PATH = UPLOAD_WORK_DIR / "memoria_notas.json"
MOBILE_PACKAGE_DIR = UPLOAD_WORK_DIR / "mobile_packages"
MOBILE_PACKAGE_MANIFEST = "dados_relatorio_mobile.json"
BRAND_LOGO_PATH = (
    ROOT / "agres-only-logo.svg"
    if (ROOT / "agres-only-logo.svg").exists()
    else ROOT / "mobile_app" / "agres-only-logo.svg"
)

POWERSHELL_EXE = Path(r"C:\WINDOWS\System32\WindowsPowerShell\v1.0\powershell.exe")
DEFAULT_AI_MODEL = "gpt-4.1-mini"
AI_FALLBACK_MODELS = ("gpt-4o-mini",)
DEFAULT_GEMINI_MODEL = "gemini-2.5-flash"
GEMINI_FALLBACK_MODELS = ("gemini-2.0-flash",)
AI_PROMPT_VERSION = "receipt-ai-v2"

CATEGORY_ORDER = [
    "Hospedagem",
    "Refeições",
    "Pedágio",
    "Frete",
    "Combustível",
    "Material de Uso e Consumo",
    "Locações de Veículos",
    "Transporte/ taxi",
    "Outras Despesas",
]

EXCEL_CATEGORY_KEYS = {
    "Hospedagem": "Hospedagem",
    "Refeições": "Refeicoes",
    "Pedágio": "Pedagio",
    "Frete": "Frete",
    "Combustível": "Combustivel",
    "Material de Uso e Consumo": "Material de Uso e Consumo",
    "Locações de Veículos": "Locacoes de Veiculos",
    "Transporte/ taxi": "Transporte/ taxi",
    "Outras Despesas": "Outras Despesas",
}

MAX_DATES = 8
MAX_OTHER_EXPENSES = 4
MAX_MILEAGE_ROWS = 4

DATE_COLUMNS = ["D", "F", "H", "J", "L", "N", "P", "R"]
IMAGE_EXTENSIONS = {".jpg", ".jpeg", ".png", ".webp", ".bmp", ".tif", ".tiff", ".heic", ".heif"}
MAX_MOBILE_PACKAGE_BYTES = 250 * 1024 * 1024
MAX_MOBILE_UNCOMPRESSED_BYTES = 300 * 1024 * 1024
MAX_MOBILE_IMAGE_BYTES = 25 * 1024 * 1024
MAX_MOBILE_ENTRIES = 200
MAX_ZIP_MEMBERS = 250
STALE_WORKSPACE_SECONDS = 24 * 60 * 60

TRIP_PRESETS = {
    "Personalizado": {"place": "", "motive": ""},
    "Herbicat - Bauru/SP": {"place": "Herbicat - Bauru/SP", "motive": "Suporte ISO31OFP"},
    "Usina São Domingos - Catanduva/SP": {
        "place": "Usina São Domingos - Catanduva/SP",
        "motive": "Suporte agroNave12+isoBox Sprayer - vinhaça",
    },
    "Agres - Curitiba/PR": {"place": "Agres - Curitiba/PR", "motive": "Suporte alinhamento engenharia"},
    "Rimaquinas - Goiânia/GO": {"place": "Rimaquinas - Goiânia/GO", "motive": "Instalação ISO31OFP"},
}

MANAGED_REVIEW_ALERTS = {
    "Nota usada em mais de uma linha",
    "Despesa possivelmente duplicada",
    "Valor zerado",
    "Valor alto",
    "Data invalida",
    "Categoria invalida",
    "Nota sem imagem vinculada",
    "Data fora do periodo",
    "Baixa confianca",
    "IA indisponivel",
    "Leitura por OCR local",
}

BLOCKING_REVIEW_ALERTS = {
    "Nota usada em mais de uma linha",
    "Valor zerado",
    "Data invalida",
    "Categoria invalida",
    "Nota sem imagem vinculada",
    "Data fora do periodo",
    "IA indisponivel",
}

DATE_PATTERN = re.compile(r"(?<!\d)([0-3]?\d)\s*[/.,'\-\s]\s*([01]?\d)\s*[/.,'\-\s]\s*(20\d{2}|\d{2})(?=\D|$)")
TEXT_DATE_PATTERN = re.compile(
    r"(?<!\d)([0-3]?\d)\s*(?:DE\s+)?"
    r"(JAN(?:EIRO)?|FEV(?:EREIRO)?|MAR(?:CO)?|ABR(?:IL)?|MAI(?:O)?|JUN(?:HO)?|JUL(?:HO)?|AGO(?:STO)?|SET(?:EMBRO)?|OUT(?:UBRO)?|NOV(?:EMBRO)?|DEZ(?:EMBRO)?)"
    r"\.?\s*(?:DE\s+)?(20\d{2}|\d{2})?",
    re.IGNORECASE,
)
MONEY_PATTERN = re.compile(r"(?<!\d)(?:R\s*[$S]?\s*)?(\d{1,3}(?:[.\s]\d{3})*[,\.]\d{2})(?=\D|$)")
MONEY_WITHOUT_SEPARATOR_PATTERN = re.compile(r"R\s*[$S]?\s*(\d{3,7})\b", re.IGNORECASE)
BROAD_MONEY_PATTERN = re.compile(r"(?:R\s*[$S]?\s*)?([0-9OQDCMBFILIS|!.,\s]{2,18})", re.IGNORECASE)
PT_MONTHS = {
    "JAN": 1,
    "JANEIRO": 1,
    "FEV": 2,
    "FEVEREIRO": 2,
    "MAR": 3,
    "MARCO": 3,
    "ABR": 4,
    "ABRIL": 4,
    "MAI": 5,
    "MAIO": 5,
    "JUN": 6,
    "JUNHO": 6,
    "JUL": 7,
    "JULHO": 7,
    "AGO": 8,
    "AGOSTO": 8,
    "SET": 9,
    "SETEMBRO": 9,
    "OUT": 10,
    "OUTUBRO": 10,
    "NOV": 11,
    "NOVEMBRO": 11,
    "DEZ": 12,
    "DEZEMBRO": 12,
}
TOTAL_KEYWORDS = (
    "TOTAL",
    "VALOR TOTAL",
    "VALOR PAGO",
    "VALOR DO SERVICO",
    "VALOR LIQUIDO",
    "VL TOTAL",
    "VLR TOTAL",
    "A PAGAR",
    "PAGAMENTO",
)
BAD_AMOUNT_CONTEXT = (
    "TRIBUTO",
    "IBPT",
    "LEI FEDERAL",
    "FEDERAL",
    "ESTADUAL",
    "MUNICIPAL",
    "APROXIMADO",
    "APROXINADO",
    "CNPJ",
    "CPF",
    "CHAVE",
    "ACESSO",
    "PROTOCOLO",
    "AUTORIZACAO",
    "SERIE",
)
PAYMENT_KEYWORDS = (
    "A PAGAR",
    "VALOR PAGO",
    "TOTAL PAGO",
)
TABLE_TOTAL_WORDS = (
    "QTD",
    "QTDE",
    "ITEM",
    "ITENS",
    "UNIT",
    "UN ",
)

CATEGORY_HINTS = {
    "Hospedagem": (
        "HOTEL",
        "POUSADA",
        "HOSPEDAGEM",
        "HOSPEDE",
        "DIARIA",
        "DIARIA",
        "BOOKING",
    ),
    "Refeições": (
        "RESTAURANTE",
        "LANCHONETE",
        "PADARIA",
        "PIZZARIA",
        "CHURRASCARIA",
        "CAFETERIA",
        "CAFE",
        "LANCHES",
        "BURGER",
        "CHECKOUT",
        "HAMBURGAO",
        "PAO DE QUEIJO",
        "MCDONALD",
        "IFOOD",
        "REFEICAO",
        "ALIMENTACAO",
        "MARMIT",
        "PIZZA",
        "SALGADO",
        "SALGADOS",
        "SUBWAY",
        "GASTRO",
        "TEMPER",
        "SUSHI",
        "HORTIFRUTI",
    ),
    "Pedágio": (
        "PEDAGIO",
        "PEDAG",
        "PRACA",
        "PRAÇA",
        "RODOVIA",
        "CONCESSIONARIA",
        "SEM PARAR",
        "TAG",
        "RECARGA AUTOMATICA",
        "TCF7D72",
        "ARAUCA",
        "WITMARSUM",
        "CONECTCAR",
        "ECOPISTAS",
        "ECOVIAS",
        "CCR",
        "ARTESP",
        "EIXO SP",
        "VIAPAULISTA",
    ),
    "Frete": (
        "FRETE",
        "TRANSPORTADORA",
        "REMESSA",
        "ENVIO",
    ),
    "Combustível": (
        "POSTO",
        "COMBUSTIVEL",
        "GASOLINA",
        "ETANOL",
        "DIESEL",
        "ABAST",
        "LITRO",
        "LITROS",
        "IPIRANGA",
        "SHELL",
        "PETROBRAS",
        "RAIZEN",
    ),
    "Material de Uso e Consumo": (
        "MATERIAL",
        "PAPELARIA",
        "FERRAGEM",
        "FERRAMENTA",
        "CASA CONSTRUCAO",
        "ELETRONICA",
        "MULTIMETRO",
        "RESISTENCIA",
        "SOLDA",
        "UTILIDADES",
        "SUPRIMENTOS",
    ),
    "Locações de Veículos": (
        "LOCADORA",
        "LOCALIZA",
        "MOVIDA",
        "UNIDAS",
        "ALUGUEL DE VEICULO",
        "LOCACAO",
        "RENT A CAR",
    ),
    "Transporte/ taxi": (
        "UBER",
        "OBRIGADO POR VIAJAR",
        "99",
        "TAXI",
        "CORRIDA",
        "TRANSPORTE",
        "ESTACIONAMENTO",
        "PARK",
    ),
    "Outras Despesas": (
        "LAVAGEM",
        "LAVACAO",
        "LAVAÇÃO",
        "LAVAGEM DE CARRO",
        "LAVACAO DE CARRO",
        "LAVA JATO",
        "LAVA-JATO",
        "LAVANDERIA",
        "LAVA RAPIDO",
        "LAVA-RAPIDO",
        "LAVA AUTO",
        "ESTETICA AUTOMOTIVA",
        "DIVERSOS",
    ),
}


@dataclass
class NoteFile:
    name: str
    path: Path


@dataclass
class NoteDetection:
    note_name: str
    note_path: Path
    expense_date: date | None
    category: str
    description: str
    value: Decimal
    confidence: str
    source: str
    ocr_text: str
    warnings: list[str]


def configure_page() -> None:
    st.set_page_config(page_title="Relatório de Despesas | Agres", layout="wide")
    st.markdown(
        """
        <style>
        :root {
            --agres-dark: #242528; --agres-muted: #66686c; --agres-line: #d9d9d7;
            --agres-bg: #f5f6f5; --agres-soft: #edf0ee; --agres-green: #35634f;
            --agres-green-dark: #284d3d; --agres-green-soft: #edf5f1; --agres-success: #2f604b;
        }
        .stApp { background: var(--agres-bg); }
        [data-testid="stHeader"], [data-testid="stToolbar"], [data-testid="stDecoration"], footer { display: none; }
        .block-container { padding-top: 1.5rem; padding-bottom: 3rem; max-width: 1140px; }
        h1, h2, h3, p { letter-spacing: 0; }
        h1 { font-size: 1.8rem; margin-bottom: .15rem; }
        h2, h3 { color: var(--agres-dark); }
        [data-testid="stMetric"] {
            min-height: 92px; padding: 14px 16px; background: #fff; border-color: var(--agres-line);
            border-radius: 8px;
        }
        [data-testid="stMetricLabel"] { color: var(--agres-muted); }
        [data-testid="stMetricValue"] { color: var(--agres-dark); font-size: 1.12rem; }
        [data-testid="stMetricValue"] > div {
            overflow: visible !important; text-overflow: clip !important; white-space: normal !important;
            line-height: 1.25;
        }
        div[data-testid="stForm"] { border: 0; padding: 0; }
        .stButton button, .stDownloadButton button {
            border-radius: 6px; min-height: 2.9rem; font-weight: 700;
        }
        .stButton button[kind="primary"], .stDownloadButton button[kind="primary"] {
            background: var(--agres-green); border-color: var(--agres-green);
        }
        [data-testid="stFileUploaderDropzone"] {
            min-height: 76px; padding: 11px 14px; background: var(--agres-green-soft);
            border: 1px dashed #8ca99a; border-radius: 8px;
        }
        [data-testid="stFileUploaderDropzone"] button {
            min-width: 0 !important; width: auto !important; min-height: 38px !important; padding: 7px 14px !important;
            color: white; background: var(--agres-green); border: 0; border-radius: 6px; font-size: .82rem;
            font-weight: 700;
        }
        [data-testid="stFileUploaderDropzone"] button:hover { background: var(--agres-green-dark); }
        [data-testid="stFileUploaderDropzoneInstructions"] { padding: 0; }
        [data-testid="stFileUploaderDropzoneInstructions"] span { color: var(--agres-dark); font-size: .82rem; }
        [data-testid="stFileUploaderDropzoneInstructions"] small { color: var(--agres-muted); font-size: .68rem; }
        [data-testid="stFileUploader"] label p { color: var(--agres-dark); font-size: .82rem; font-weight: 700; }
        [data-testid="stFileUploaderFile"] {
            min-height: 54px; margin-top: 8px; padding: 8px 10px; background: #fff;
            border: 1px solid var(--agres-line); border-radius: 7px;
        }
        [data-testid="stFileUploaderFile"] [data-testid="stFileUploaderFileName"] {
            color: var(--agres-dark); font-size: .82rem; font-weight: 650;
        }
        [data-testid="stFileUploaderFile"] small { color: var(--agres-muted); font-size: .68rem; }
        [data-testid="stFileUploaderFile"] button,
        [data-testid="stFileUploaderDeleteBtn"] button {
            width: 32px !important; min-width: 32px !important; min-height: 32px !important; padding: 4px !important;
            color: #6e4141; background: #f8eeee; border: 1px solid #ead3d3; border-radius: 50%;
        }
        [data-testid="stExpander"] {
            background: #fff; border-color: var(--agres-line); border-radius: 8px;
        }
        [data-testid="stDataFrame"] { border: 1px solid var(--agres-line); border-radius: 8px; overflow: hidden; }
        [data-testid="stAlert"] { border-radius: 6px; }
        .brand-header {
            display: grid; grid-template-columns: minmax(180px, 290px) 1fr auto; gap: 30px; align-items: center;
            padding: 18px 0 20px; margin-bottom: 4px; border-bottom: 1px solid var(--agres-line);
        }
        .brand-header img { display: block; width: 100%; max-height: 42px; object-fit: contain; object-position: left center; }
        .brand-header h1 { margin: 0; color: var(--agres-dark); font-size: 1.5rem; }
        .brand-header p { margin: 5px 0 0; color: var(--agres-muted); font-size: .88rem; }
        .product-badge {
            padding: 6px 10px; color: #4e5053; background: var(--agres-soft); border: 1px solid var(--agres-line);
            border-radius: 999px; font-size: .68rem; font-weight: 750; text-transform: uppercase; white-space: nowrap;
        }
        .workflow {
            display: grid; grid-template-columns: repeat(4, 1fr); gap: 0; margin: 20px 0 26px;
            background: #fff; border: 1px solid var(--agres-line); border-radius: 8px; overflow: hidden;
        }
        .workflow-item {
            display: flex; align-items: center; gap: 9px; padding: 11px 13px; color: var(--agres-muted);
            border-right: 1px solid var(--agres-line); font-size: .73rem; font-weight: 700;
        }
        .workflow-item:last-child { border-right: 0; }
        .workflow-item span {
            display: inline-grid; width: 23px; height: 23px; place-items: center; flex: 0 0 23px;
            color: #fff; background: var(--agres-green); border-radius: 50%; font-size: .68rem;
        }
        .flow-step {
            display: grid; grid-template-columns: 34px 1fr; gap: 11px; align-items: center;
            margin: 28px 0 12px; padding-top: 2px;
        }
        .flow-step span {
            display: inline-grid; width: 30px; height: 30px; place-items: center; color: #fff;
            background: var(--agres-green); border-radius: 50%; font-size: .74rem; font-weight: 750;
        }
        .flow-step strong { display: block; color: var(--agres-dark); font-size: 1.05rem; }
        .flow-step small { display: block; margin-top: 2px; color: var(--agres-muted); font-size: .78rem; font-weight: 500; }
        .import-status {
            display: flex; align-items: center; justify-content: space-between; gap: 18px; padding: 13px 15px;
            margin: 10px 0 14px; color: #284e3d; background: #edf4f0; border: 1px solid #cbded4; border-radius: 6px;
        }
        .import-status strong { font-size: .88rem; }
        .import-status span { color: #4f6d5e; font-size: .78rem; }
        .section-caption { margin: -4px 0 12px; color: var(--agres-muted); font-size: .78rem; }
        .final-panel {
            padding: 16px; margin: 8px 0 14px; color: #31443a; background: #edf4f0;
            border: 1px solid #cbded4; border-radius: 8px;
        }
        .final-panel strong { display: block; margin-bottom: 3px; color: #284e3d; }
        .final-panel span { font-size: .8rem; }
        hr { border-color: var(--agres-line); }
        @media (max-width: 760px) {
            .workflow { grid-template-columns: 1fr 1fr; }
            .workflow-item:nth-child(2) { border-right: 0; }
            .workflow-item:nth-child(-n+2) { border-bottom: 1px solid var(--agres-line); }
        }
        @media (max-width: 900px) {
            .block-container { padding-left: .85rem; padding-right: .85rem; padding-top: .8rem; }
            h1 { font-size: 1.45rem; }
            .stButton button, .stDownloadButton button { width: 100%; min-height: 3.1rem; font-size: 1rem; }
            [data-testid="stFileUploaderDropzone"] { min-height: 76px; }
            div[data-testid="column"] { width: 100% !important; flex: 1 1 100% !important; }
            .brand-header { grid-template-columns: 1fr; gap: 14px; padding: 16px; }
            .brand-header img { max-width: 260px; }
            .product-badge { width: fit-content; }
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


def brl(value: Decimal | float | int) -> str:
    amount = Decimal(str(value)).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    text = f"{amount:,.2f}"
    return f"R$ {text}".replace(",", "X").replace(".", ",").replace("X", ".")


def parse_money(value: object) -> Decimal:
    if value is None:
        return Decimal("0")
    if isinstance(value, Decimal):
        return value.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    if isinstance(value, (int, float)):
        return Decimal(str(value)).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)

    text = str(value).strip()
    if not text:
        return Decimal("0")

    text = text.replace("R$", "").replace(" ", "")
    if "," in text:
        text = text.replace(".", "").replace(",", ".")

    try:
        return Decimal(text).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    except InvalidOperation as exc:
        raise ValueError(f"Valor inválido: {value}") from exc


def to_iso(value: object) -> str:
    if isinstance(value, datetime):
        return value.date().isoformat()
    if isinstance(value, date):
        return value.isoformat()
    if pd.isna(value):
        return ""
    text = str(value).strip()
    if not text:
        return ""
    if re.fullmatch(r"\d{4}-\d{2}-\d{2}", text):
        return text
    parsed = pd.to_datetime(text, dayfirst=True, errors="raise")
    return parsed.date().isoformat()


def to_date(value: object) -> date | None:
    iso = to_iso(value)
    if not iso:
        return None
    return date.fromisoformat(iso)


def slugify(value: str, fallback: str = "RELATORIO") -> str:
    text = unicodedata.normalize("NFKD", repair_text_encoding(value))
    text = "".join(ch for ch in text if not unicodedata.combining(ch))
    text = text.upper()
    text = re.sub(r"[^A-Z0-9]+", "_", text)
    text = re.sub(r"_+", "_", text).strip("_")
    return text[:110] or fallback


def repair_text_encoding(value: object) -> str:
    text = "" if value is None else str(value)
    if not any(marker in text for marker in ("Ã", "Â", "â")):
        return text

    marker_count = sum(text.count(marker) for marker in ("Ã", "Â", "â"))
    try:
        repaired = text.encode("latin-1").decode("utf-8")
    except (UnicodeEncodeError, UnicodeDecodeError):
        return text

    repaired_marker_count = sum(repaired.count(marker) for marker in ("Ã", "Â", "â"))
    return repaired if repaired_marker_count < marker_count else text


def ascii_key(value: str) -> str:
    text = unicodedata.normalize("NFKD", repair_text_encoding(value))
    text = "".join(ch for ch in text if not unicodedata.combining(ch))
    text = re.sub(r"[^A-Za-z0-9]+", " ", text)
    return re.sub(r"\s+", " ", text).strip().upper()


def note_signature(note_map: dict[str, Path]) -> str:
    parts: list[str] = []
    for name, path in sorted(note_map.items()):
        try:
            stat = path.stat()
            parts.append(f"{name}:{stat.st_size}:{file_sha1(path)}")
        except FileNotFoundError:
            parts.append(f"{name}:missing")
    return sha1("|".join(parts).encode("utf-8")).hexdigest()


def memory_signature() -> str:
    try:
        stat = AI_MEMORY_PATH.stat()
    except FileNotFoundError:
        return "no-memory"
    return f"{stat.st_size}:{int(stat.st_mtime)}"


def file_sha1(path: Path) -> str:
    digest = sha1()
    with path.open("rb") as file:
        for chunk in iter(lambda: file.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def atomic_write_bytes(path: Path, data: bytes) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    temp_path = path.with_name(f".{path.name}.{uuid4().hex}.tmp")
    try:
        with temp_path.open("wb") as file:
            file.write(data)
            file.flush()
            os.fsync(file.fileno())
        temp_path.replace(path)
    finally:
        temp_path.unlink(missing_ok=True)


def validate_image_bytes(data: bytes, label: str) -> None:
    if not data:
        raise ValueError(f"Imagem vazia no pacote: {label}")
    if len(data) > MAX_MOBILE_IMAGE_BYTES:
        raise ValueError(f"Imagem excede {MAX_MOBILE_IMAGE_BYTES // 1024 // 1024} MB: {label}")
    try:
        with Image.open(BytesIO(data)) as image:
            image.verify()
    except Exception as exc:
        raise ValueError(f"Imagem inválida ou corrompida no pacote: {label}") from exc


def prepare_ocr_image(path: Path, mode: str = "document") -> Path:
    cache_dir = UPLOAD_WORK_DIR / "ocr"
    cache_dir.mkdir(parents=True, exist_ok=True)
    try:
        stat = path.stat()
        key = sha1(f"v3:{mode}:{path.resolve()}:{stat.st_size}:{stat.st_mtime}".encode("utf-8")).hexdigest()
    except FileNotFoundError:
        return path

    target = cache_dir / f"{key}.png"
    if target.exists():
        return target

    with Image.open(path) as image:
        image = ImageOps.exif_transpose(image)
        if mode == "original":
            image = image.convert("RGB")
        elif mode == "document":
            image = ImageOps.autocontrast(image.convert("L"))
            image = ImageEnhance.Contrast(image).enhance(1.4)
            image = image.filter(ImageFilter.SHARPEN)
        elif mode == "contrast":
            image = ImageOps.autocontrast(image.convert("L"))
            image = ImageEnhance.Contrast(image).enhance(2.0)
            image = image.filter(ImageFilter.SHARPEN)
        elif mode == "binary":
            image = ImageOps.autocontrast(image.convert("L"))
            image = ImageEnhance.Contrast(image).enhance(1.7)
            image = image.point(lambda px: 255 if px > 178 else 0)
        else:
            raise ValueError(f"Modo de OCR desconhecido: {mode}")

        width, height = image.size
        target_width = 1900 if height >= width else 2300
        if width < target_width:
            scale = target_width / width
            image = image.resize((int(width * scale), int(height * scale)), Image.Resampling.LANCZOS)
        image.save(target, "PNG")

    return target


@st.cache_data(show_spinner=False)
def run_windows_ocr_cached(path_text: str, size: int, mtime: float) -> str:
    _ = (size, mtime)
    source_path = Path(path_text)
    command = [
        str(POWERSHELL_EXE),
        "-NoProfile",
        "-ExecutionPolicy",
        "Bypass",
        "-File",
        str(OCR_SCRIPT_PATH),
        "-ImagePath",
        str(source_path),
    ]
    result = subprocess.run(command, capture_output=True, text=True, encoding="utf-8", errors="replace", cwd=ROOT)
    if result.returncode != 0:
        raise RuntimeError((result.stderr or result.stdout or "Falha ao ler imagem com OCR.").strip())
    payload = json.loads(result.stdout)
    return str(payload.get("text", ""))


def run_windows_ocr(path: Path) -> str:
    stat = path.stat()
    return run_windows_ocr_cached(str(path), stat.st_size, stat.st_mtime)


def ocr_quality(text: str) -> int:
    lines = [line for line in text.splitlines() if line.strip()]
    score = min(len(lines), 25)
    if date_candidates_from_text(text):
        score += 20
    if amount_candidates_from_text(text):
        score += 20
    normalized = ascii_key(text)
    if any(keyword in normalized for keyword in TOTAL_KEYWORDS):
        score += 8
    return score


def ocr_note(path: Path) -> str:
    texts: list[str] = []
    for mode in ("document", "original", "contrast", "binary"):
        prepared = prepare_ocr_image(path, mode)
        try:
            text = run_windows_ocr(prepared)
        except Exception:
            continue
        if text.strip():
            texts.append(text)
        if ocr_quality(text) >= 55:
            break

    seen: set[str] = set()
    merged_lines: list[str] = []
    for text in sorted(texts, key=ocr_quality, reverse=True):
        for line in text.splitlines():
            clean = line.strip()
            key = ascii_key(clean)
            if not clean or key in seen:
                continue
            seen.add(key)
            merged_lines.append(clean)
    return "\n".join(merged_lines)


def infer_year_for_text_date(month: int, start_date: date | None, end_date: date | None) -> int:
    if start_date and end_date:
        for candidate_year in sorted({start_date.year, end_date.year}):
            try:
                probe = date(candidate_year, month, 1)
            except ValueError:
                continue
            if start_date.replace(day=1) <= probe <= end_date.replace(day=28):
                return candidate_year
        return start_date.year
    if start_date:
        return start_date.year
    return date.today().year


def date_candidates_from_text(
    text: str,
    start_date: date | None = None,
    end_date: date | None = None,
) -> list[date]:
    candidates: list[date] = []
    for day_text, month_text, year_text in DATE_PATTERN.findall(text):
        day = int(day_text)
        month = int(month_text)
        year = int(year_text)
        if year < 100:
            year += 2000
        try:
            candidates.append(date(year, month, day))
        except ValueError:
            continue

    normalized = ascii_key(text)
    for day_text, month_text, year_text in TEXT_DATE_PATTERN.findall(normalized):
        day = int(day_text)
        month = PT_MONTHS.get(month_text.upper().rstrip("."))
        if not month:
            continue
        if year_text:
            year = int(year_text)
            if year < 100:
                year += 2000
        else:
            year = infer_year_for_text_date(month, start_date, end_date)
        try:
            candidates.append(date(year, month, day))
        except ValueError:
            continue

    unique: list[date] = []
    for candidate in candidates:
        if candidate not in unique:
            unique.append(candidate)
    return unique


def parse_detected_date(text: str, start_date: date | None, end_date: date | None) -> date | None:
    candidates = date_candidates_from_text(text, start_date, end_date)
    if not candidates:
        return None

    if start_date and end_date:
        in_range = [candidate for candidate in candidates if start_date <= candidate <= end_date]
        if in_range:
            return in_range[0]

    if start_date:
        return min(candidates, key=lambda candidate: abs((candidate - start_date).days))

    return candidates[0]


def parse_decimal_text(value: str, force_cents: bool = False) -> Decimal | None:
    text = value.strip().replace(" ", "")
    if not text:
        return None
    if force_cents and text.isdigit():
        return (Decimal(text) / Decimal("100")).quantize(Decimal("0.01"))
    if "," in text:
        text = text.replace(".", "").replace(",", ".")
    try:
        return Decimal(text).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    except InvalidOperation:
        return None


def clean_ocr_money_token(value: str) -> str:
    text = value.upper().strip()
    text = text.replace("R$", "").replace("RS", "")
    text = text.replace(" ", "")
    substitutions = str.maketrans(
        {
            "O": "0",
            "Q": "0",
            "D": "0",
            "C": "0",
            "B": "0",
            "F": "1",
            "I": "1",
            "L": "1",
            "|": "1",
            "!": "1",
            "S": "5",
        }
    )
    text = text.translate(substitutions)
    text = text.replace("M", "00")
    return re.sub(r"[^0-9,\.]", "", text)


def parse_ocr_money_token(value: str, force_cents: bool = False) -> Decimal | None:
    text = clean_ocr_money_token(value)
    if not text:
        return None

    digits = re.sub(r"\D", "", text)
    if len(digits) < 2:
        return None

    separators = [index for index, char in enumerate(text) if char in ",."]
    if separators:
        separator = separators[-1]
        cents = re.sub(r"\D", "", text[separator + 1 :])
        whole = re.sub(r"\D", "", text[:separator])
        if len(cents) >= 2:
            cents = cents[:2]
            whole = whole or "0"
            return Decimal(f"{int(whole)}.{cents}").quantize(Decimal("0.01"))
        if len(cents) == 1:
            whole = whole or "0"
            return Decimal(f"{int(whole)}.{cents}0").quantize(Decimal("0.01"))

    if force_cents and len(digits) >= 3:
        return (Decimal(digits) / Decimal("100")).quantize(Decimal("0.01"))
    if len(digits) >= 4 and any(char.isalpha() for char in value):
        return Decimal(digits).quantize(Decimal("0.01"))
    return None


def has_currency_marker(line: str) -> bool:
    return re.search(r"\bR\s*(?:[$S]\s*)\d|\bR\s+\d", line, re.IGNORECASE) is not None


def amount_candidates_from_line(line: str, allow_loose: bool = False) -> list[Decimal]:
    values: list[Decimal] = []
    for match in MONEY_PATTERN.findall(line):
        value = parse_decimal_text(match)
        if value is not None:
            values.append(value)

    for match in MONEY_WITHOUT_SEPARATOR_PATTERN.findall(line):
        value = parse_decimal_text(match, force_cents=True)
        if value is not None:
            values.append(value)

    normalized = ascii_key(line)
    has_currency_prefix = has_currency_marker(line)
    if has_currency_prefix or "VALOR" in normalized or "TOTAL" in normalized or "PAGAR" in normalized:
        for match in BROAD_MONEY_PATTERN.findall(line):
            has_original_digit = re.search(r"\d", match) is not None
            has_decimal_hint = "," in match or has_currency_prefix
            if not has_original_digit and not has_currency_prefix:
                continue
            if not has_decimal_hint:
                continue
            value = parse_ocr_money_token(match, force_cents=("R" in normalized or "RS" in normalized))
            if value is not None:
                values.append(value)

    context_allows_loose = allow_loose or any(token in normalized for token in ("VALOR", "TOTAL", "PAGAR", "PAGO"))
    if context_allows_loose:
        for whole, cents in re.findall(r"(?<!\d)(\d{1,4})[,.](\d{1,2})(?!\d)", line):
            cents = cents.ljust(2, "0")
            values.append(Decimal(f"{int(whole)}.{cents}").quantize(Decimal("0.01")))
        for whole, cents in re.findall(r"(?<!\d)(\d{1,4})\s+(\d{2})(?!\d)", line):
            values.append(Decimal(f"{int(whole)}.{cents}").quantize(Decimal("0.01")))

    if allow_loose:
        compact = line.strip()
        loose_match = re.fullmatch(r"(?:R\s*[$S]?\s*)?(\d{1,4})[,.](\d{1,2})", compact, flags=re.IGNORECASE)
        if loose_match:
            whole, cents = loose_match.groups()
            cents = cents.ljust(2, "0")
            values.append(Decimal(f"{int(whole)}.{cents}").quantize(Decimal("0.01")))
        split_match = re.fullmatch(r"(?:R\s*[$S]?\s*)?(\d{1,4})\s+(\d{2})", compact, flags=re.IGNORECASE)
        if split_match:
            whole, cents = split_match.groups()
            values.append(Decimal(f"{int(whole)}.{cents}").quantize(Decimal("0.01")))

    filtered = [value for value in values if Decimal("0.01") <= value <= Decimal("50000")]
    unique: list[Decimal] = []
    for value in filtered:
        if value not in unique:
            unique.append(value)
    return unique


def amount_candidates_from_text(text: str) -> list[Decimal]:
    values: list[Decimal] = []
    for line in text.splitlines():
        values.extend(amount_candidates_from_line(line))
    unique: list[Decimal] = []
    for value in values:
        if value not in unique:
            unique.append(value)
    return unique


def is_final_amount_context(context: str) -> bool:
    if any(keyword in context for keyword in PAYMENT_KEYWORDS):
        return True
    if any(word in context for word in TABLE_TOTAL_WORDS):
        return False
    return any(keyword in context for keyword in TOTAL_KEYWORDS)


def parse_ride_receipt_amount(text: str) -> Decimal | None:
    normalized = ascii_key(text)
    if "OBRIGADO POR VIAJAR" not in normalized and "UBER" not in normalized:
        return None

    seen_total = False
    for line in [line.strip() for line in text.splitlines() if line.strip()]:
        line_key = ascii_key(line)
        if "TOTAL" == line_key or line_key.startswith("TOTAL "):
            seen_total = True
            continue
        if not seen_total:
            continue
        values = amount_candidates_from_line(line)
        if values and has_currency_marker(line):
            return max(values).quantize(Decimal("0.01"))
    return None


def parse_tag_statement_amount(text: str) -> Decimal | None:
    normalized = ascii_key(text)
    if not any(token in normalized for token in ("TCF7D72", "RECARGA AUTOMATICA", "SALDO DA TAG", "SEM PARAR")):
        return None

    recharge_values: list[Decimal] = []
    debit_values: list[Decimal] = []
    for line in [line.strip() for line in text.splitlines() if line.strip()]:
        line_key = ascii_key(line)
        values = amount_candidates_from_line(line, allow_loose=True)
        if not values:
            continue
        if "RECARGA" in line_key or "ADICIONADO AO SALDO" in line_key:
            recharge_values.extend(values)
            continue
        if re.search(r"(^|\s)-\s*R\s*[$S]?", line, flags=re.IGNORECASE) or line.startswith("-"):
            debit_values.extend(values)

    if recharge_values:
        return max(recharge_values).quantize(Decimal("0.01"))
    if debit_values:
        return sum(debit_values, Decimal("0")).quantize(Decimal("0.01"))
    return None


def parse_detected_amount(text: str) -> Decimal:
    tag_amount = parse_tag_statement_amount(text)
    if tag_amount is not None:
        return tag_amount

    ride_amount = parse_ride_receipt_amount(text)
    if ride_amount is not None:
        return ride_amount

    lines = [line.strip() for line in text.splitlines() if line.strip()]
    scored: list[tuple[int, Decimal]] = []

    recent_context: list[str] = []
    line_keys = [ascii_key(line) for line in lines]
    for index, line in enumerate(lines):
        normalized = ascii_key(line)
        has_recent_total = any(is_final_amount_context(context) for context in recent_context[-5:])
        has_recent_payment = any(any(keyword in context for keyword in PAYMENT_KEYWORDS) for context in recent_context[-3:])
        has_next_payment = any(
            any(keyword in context for keyword in PAYMENT_KEYWORDS)
            for context in line_keys[index + 1 : index + 3]
        )
        line_values = amount_candidates_from_line(line, allow_loose=has_recent_total)
        if not line_values:
            recent_context.append(normalized)
            continue
        score = 2
        has_currency_prefix = has_currency_marker(line)
        if has_next_payment:
            score = 10
        elif any(keyword in normalized for keyword in PAYMENT_KEYWORDS):
            score = 10
        elif is_final_amount_context(normalized):
            score = 8
        elif has_recent_payment:
            score = 9
        elif has_recent_total:
            score = 7
        elif has_currency_prefix:
            score = 9
        if "TROCO" in normalized or "DESCONTO" in normalized:
            score -= 3
        if any(keyword in normalized for keyword in BAD_AMOUNT_CONTEXT):
            score -= 8
        for value in line_values:
            scored.append((score, value))
        recent_context.append(normalized)

    if not scored:
        for value in amount_candidates_from_text(text):
            scored.append((1, value))

    if not scored:
        return Decimal("0")

    return max(scored, key=lambda item: (item[0], item[1]))[1].quantize(Decimal("0.01"))


def infer_category(text: str, note_name: str) -> str:
    source = f"{text}\n{note_name}"
    normalized = ascii_key(source)
    if is_vehicle_wash_text(source):
        return "Outras Despesas"
    scores: dict[str, int] = {}
    for category, hints in CATEGORY_HINTS.items():
        score = 0
        for hint in hints:
            if ascii_key(hint) in normalized:
                score += 1
        if score:
            scores[category] = score

    if not scores:
        return "Outras Despesas"
    return max(scores.items(), key=lambda item: item[1])[0]


def is_vehicle_wash_text(*values: object) -> bool:
    source = " ".join(repair_text_encoding(value) for value in values if value is not None)
    normalized = ascii_key(source)
    wash_tokens = (
        "LAVAGEM",
        "LAVACAO",
        "LAVA RAPIDO",
        "LAVA JATO",
        "LAVA AUTO",
        "ESTETICA AUTOMOTIVA",
        "HIGIENIZACAO",
    )
    vehicle_tokens = ("CARRO", "VEICULO", "AUTO", "AUTOMOTIVA", "AUTOMOTIVO")
    if any(token in normalized for token in ("LAVA RAPIDO", "LAVA JATO", "LAVA AUTO")):
        return True
    return any(token in normalized for token in wash_tokens) and (
        any(token in normalized for token in vehicle_tokens) or "LAVANDERIA" not in normalized
    )


def compact_description(text: str, note_name: str, category: str) -> str:
    lines = [line.strip() for line in text.splitlines() if line.strip()]
    ignored = ("CNPJ", "CPF", "IE", "TOTAL", "VALOR", "CUPOM", "EXTRATO", "SAT", "NFC", "CHAVE")
    for line in lines[:8]:
        normalized = ascii_key(line)
        if len(normalized) < 4:
            continue
        if any(token in normalized for token in ignored):
            continue
        return line[:80]
    return Path(note_name).stem[:80] or category


def normalize_ai_category(value: object) -> str:
    key = ascii_key(str(value or ""))
    if not key:
        return "Outras Despesas"

    direct_aliases = {
        "REFEICAO": CATEGORY_ORDER[1],
        "REFEICOES": CATEGORY_ORDER[1],
        "ALIMENTACAO": CATEGORY_ORDER[1],
        "PEDAGIO": CATEGORY_ORDER[2],
        "COMBUSTIVEL": CATEGORY_ORDER[4],
        "LAVAGEM": CATEGORY_ORDER[8],
        "LAVACAO": CATEGORY_ORDER[8],
        "LAVAGEM DE CARRO": CATEGORY_ORDER[8],
        "LAVACAO DE CARRO": CATEGORY_ORDER[8],
        "LAVA RAPIDO": CATEGORY_ORDER[8],
        "LAVA JATO": CATEGORY_ORDER[8],
        "TAXI": CATEGORY_ORDER[7],
        "TRANSPORTE": CATEGORY_ORDER[7],
        "TRANSPORTE TAXI": CATEGORY_ORDER[7],
        "UBER": CATEGORY_ORDER[7],
        "HOTEL": CATEGORY_ORDER[0],
        "HOSPEDAGEM": CATEGORY_ORDER[0],
        "LOCACAO": CATEGORY_ORDER[6],
        "LOCACOES DE VEICULOS": CATEGORY_ORDER[6],
    }
    if key in direct_aliases:
        return direct_aliases[key]

    for category in CATEGORY_ORDER:
        category_key = ascii_key(category)
        if key == category_key or key in category_key or category_key in key:
            return category

    aliases = {
        "REFEICAO": "REFEICOES",
        "REFEICOES": "REFEICOES",
        "ALIMENTACAO": "REFEICOES",
        "PEDAGIO": "PEDAGIO",
        "COMBUSTIVEL": "COMBUSTIVEL",
        "LAVAGEM": "OUTRAS DESPESAS",
        "LAVACAO": "OUTRAS DESPESAS",
        "LAVAGEM DE CARRO": "OUTRAS DESPESAS",
        "LAVACAO DE CARRO": "OUTRAS DESPESAS",
        "LAVA RAPIDO": "OUTRAS DESPESAS",
        "LAVA JATO": "OUTRAS DESPESAS",
        "TAXI": "TRANSPORTE TAXI",
        "TRANSPORTE": "TRANSPORTE TAXI",
        "UBER": "TRANSPORTE TAXI",
        "HOTEL": "HOSPEDAGEM",
        "HOSPEDAGEM": "HOSPEDAGEM",
        "LOCACAO": "LOCACOES DE VEICULOS",
    }
    target_key = aliases.get(key)
    if target_key:
        for category in CATEGORY_ORDER:
            if ascii_key(category).replace("/", " ") == target_key or ascii_key(category) == target_key:
                return category
    return "Outras Despesas"


def normalize_ai_confidence(value: object, warnings: list[str], amount: Decimal, expense_date: date | None) -> str:
    key = ascii_key(str(value or ""))
    if key.startswith("ALT") and not warnings:
        return "Alta"
    if key.startswith("BAIX"):
        return "Baixa"
    if warnings:
        return "Media" if amount > 0 or expense_date else "Baixa"
    return "Media"


def coerce_ai_date(value: object, start_date: date | None, end_date: date | None) -> date | None:
    if value is None:
        return None
    text = str(value).strip()
    if not text or text.lower() in {"null", "none", "nao encontrado"}:
        return None
    try:
        return date.fromisoformat(text[:10])
    except ValueError:
        pass
    try:
        return to_date(text)
    except Exception:
        pass

    parsed = parse_detected_date(text, start_date, end_date)
    return parsed


def receipt_ai_json_schema() -> dict[str, Any]:
    return {
        "type": "object",
        "additionalProperties": False,
        "properties": {
            "expense_date": {
                "type": ["string", "null"],
                "description": "Data do gasto em YYYY-MM-DD, ou null se nao estiver legivel.",
            },
            "amount": {
                "type": ["number", "null"],
                "description": "Valor final pago/reembolsavel, com centavos.",
            },
            "category": {
                "type": "string",
                "enum": CATEGORY_ORDER,
            },
            "description": {
                "type": "string",
                "description": "Nome curto do estabelecimento, servico ou rota.",
            },
            "confidence": {
                "type": "string",
                "enum": ["Alta", "Media", "Baixa"],
            },
            "needs_review": {
                "type": "boolean",
            },
            "warnings": {
                "type": "array",
                "items": {"type": "string"},
            },
            "evidence": {
                "type": "string",
                "description": "Resumo curto da linha/parte da nota usada para decidir.",
            },
        },
        "required": [
            "expense_date",
            "amount",
            "category",
            "description",
            "confidence",
            "needs_review",
            "warnings",
            "evidence",
        ],
    }


def receipt_audit_json_schema() -> dict[str, Any]:
    return {
        "type": "object",
        "additionalProperties": False,
        "properties": {
            "approved": {
                "type": "boolean",
                "description": "true somente quando data, valor e categoria parecem corretos.",
            },
            "expense_date": {"type": ["string", "null"]},
            "amount": {"type": ["number", "null"]},
            "category": {"type": "string", "enum": CATEGORY_ORDER},
            "description": {"type": "string"},
            "confidence": {"type": "string", "enum": ["Alta", "Media", "Baixa"]},
            "warnings": {"type": "array", "items": {"type": "string"}},
            "evidence": {"type": "string"},
        },
        "required": [
            "approved",
            "expense_date",
            "amount",
            "category",
            "description",
            "confidence",
            "warnings",
            "evidence",
        ],
    }


def build_receipt_ai_prompt(note_name: str, start_date: date | None, end_date: date | None) -> str:
    categories = "\n".join(f"- {category}" for category in CATEGORY_ORDER)
    period = "nao informado"
    if start_date and end_date:
        period = f"{start_date.isoformat()} ate {end_date.isoformat()}"
    elif start_date:
        period = f"a partir de {start_date.isoformat()}"

    return f"""
Voce e um especialista em prestacao de contas no Brasil. Leia a imagem da nota e extraia os campos para um relatorio de despesas.

Arquivo: {note_name}
Periodo da viagem: {period}

Categorias validas:
{categories}

Regras:
- Use o valor final pago/reembolsavel, nao CNPJ, CPF, chave de acesso, tributos, troco, desconto, subtotal, preco unitario ou quantidade.
- Em NFC-e/cupom fiscal, prefira "Valor a pagar", "Valor pago", "Total", "Valor total" ou equivalente.
- Em NFS-e/hotel, use o valor total da nota/servico.
- Em combustivel, use o total pago, nao o preco por litro nem a quantidade de litros.
- Em Uber/taxi/app, use a linha "Total". Ignore taxa, preco intermediario, custo fixo ou repasses.
- Lavação, lavagem, lava rapido, lava-jato ou limpeza de carro/veiculo e "Outras Despesas"; nunca classifique como Locacoes de Veiculos.
- Em pedagio/tag/extrato, se nao houver total explicito, some apenas lancamentos de pedagio/debito visiveis; ignore recargas, saldo e creditos.
- Em recibo manuscrito, leia valor, data e descricao mesmo se o texto for parcial.
- Se a data aparecer sem ano (ex.: "18 de abril"), use o ano do periodo da viagem quando fizer sentido.
- Se houver mais de um valor possivel, escolha o que melhor representa o total pago e marque needs_review=true.
- Nao invente dados. Se um campo nao estiver legivel, use null/0 quando aplicavel e explique em warnings.
- Retorne somente o JSON definido pelo schema, sem texto extra.
""".strip()


def build_receipt_audit_prompt(
    note_name: str,
    detection: NoteDetection,
    start_date: date | None,
    end_date: date | None,
) -> str:
    extracted = {
        "expense_date": detection.expense_date.isoformat() if detection.expense_date else None,
        "amount": float(detection.value),
        "category": detection.category,
        "description": detection.description,
        "confidence": detection.confidence,
        "source": detection.source,
        "warnings": detection.warnings,
    }
    period = "nao informado"
    if start_date and end_date:
        period = f"{start_date.isoformat()} ate {end_date.isoformat()}"
    elif start_date:
        period = f"a partir de {start_date.isoformat()}"

    return f"""
Voce e um auditor de prestacao de contas. Revise a imagem da nota e a extracao feita por outra IA.

Arquivo: {note_name}
Periodo da viagem: {period}
Extracao inicial:
{json.dumps(extracted, ensure_ascii=False, indent=2)}

Tarefa:
- Confirme se data, valor final pago e categoria estao corretos.
- Se algum campo estiver errado, retorne o valor corrigido.
- Nao aprove se houver duvida relevante, valor ilegivel, data conflitante, duplicidade aparente ou categoria incerta.
- Para pedagio/tag, use somente gasto reembolsavel de pedagio/debito; nao use saldo.
- Para combustivel, use o total pago, nao litros nem preco unitario.
- Para Uber/taxi/app, use Total.
- Lavação, lavagem, lava rapido, lava-jato ou limpeza de carro/veiculo e "Outras Despesas"; nao e locacao de veiculo.
- Retorne somente o JSON definido pelo schema.
""".strip()


def receipt_image_jpeg_bytes(path: Path) -> bytes:
    with Image.open(path) as image:
        image = ImageOps.exif_transpose(image)
        image.thumbnail((2200, 2200), Image.Resampling.LANCZOS)
        if image.mode in {"RGBA", "LA"}:
            alpha = image.getchannel("A")
            background = Image.new("RGB", image.size, "white")
            background.paste(image.convert("RGB"), mask=alpha)
            image = background
        else:
            image = image.convert("RGB")

        buffer = BytesIO()
        image.save(buffer, format="JPEG", quality=90, optimize=True)
        return buffer.getvalue()


def receipt_image_data_url(path: Path) -> str:
    encoded = base64.b64encode(receipt_image_jpeg_bytes(path)).decode("ascii")
    return f"data:image/jpeg;base64,{encoded}"


def image_quality_warnings(path: Path) -> list[str]:
    try:
        with Image.open(path) as image:
            image = ImageOps.exif_transpose(image)
            width, height = image.size
            gray = image.convert("L")
            gray.thumbnail((320, 320), Image.Resampling.LANCZOS)
            stat = ImageStat.Stat(gray)
            brightness = stat.mean[0]
            contrast = stat.stddev[0]
    except Exception as exc:
        return [f"Nao foi possivel validar qualidade da imagem: {exc}"]

    warnings: list[str] = []
    if min(width, height) < 650:
        warnings.append("Imagem pequena; conferir leitura")
    if max(width, height) / max(min(width, height), 1) > 4:
        warnings.append("Imagem muito estreita/cortada; conferir")
    if brightness < 45:
        warnings.append("Imagem escura; conferir")
    elif brightness > 235:
        warnings.append("Imagem muito clara; conferir")
    if contrast < 18:
        warnings.append("Baixo contraste/possivel borrada; conferir")
    return warnings


def read_openai_output_text(payload: dict[str, Any]) -> str:
    direct = payload.get("output_text")
    if isinstance(direct, str) and direct.strip():
        return direct.strip()

    output = payload.get("output")
    if isinstance(output, list):
        chunks: list[str] = []
        for item in output:
            if not isinstance(item, dict):
                continue
            content = item.get("content")
            if not isinstance(content, list):
                continue
            for content_item in content:
                if not isinstance(content_item, dict):
                    continue
                text = content_item.get("text")
                if isinstance(text, str):
                    chunks.append(text)
        if chunks:
            return "\n".join(chunks).strip()

    choices = payload.get("choices")
    if isinstance(choices, list) and choices:
        message = choices[0].get("message", {}) if isinstance(choices[0], dict) else {}
        content = message.get("content") if isinstance(message, dict) else None
        if isinstance(content, str):
            return content.strip()
    return ""


def parse_json_object(text: str) -> dict[str, Any]:
    try:
        payload = json.loads(text)
    except json.JSONDecodeError:
        start = text.find("{")
        end = text.rfind("}")
        if start == -1 or end == -1 or end <= start:
            raise
        payload = json.loads(text[start : end + 1])
    if not isinstance(payload, dict):
        raise ValueError("A IA retornou um JSON que nao e um objeto.")
    return payload


def ai_provider_label(provider: str) -> str:
    return "Google AI Studio" if provider == "gemini" else "IA OpenAI"


def ai_model_candidates(provider: str, model: str) -> list[str]:
    candidates: list[str] = []
    fallbacks = GEMINI_FALLBACK_MODELS if provider == "gemini" else AI_FALLBACK_MODELS
    for candidate in (model, *fallbacks):
        clean = candidate.strip()
        if clean and clean not in candidates:
            candidates.append(clean)
    return candidates


def should_try_next_ai_model(exc: Exception) -> bool:
    message = str(exc)
    upper = message.upper()
    if "HTTP 401" in message or "HTTP 403" in message:
        return False
    if "INSUFFICIENT_QUOTA" in upper or "INVALID_API_KEY" in upper or "API_KEY_INVALID" in upper:
        return False
    if "API KEY NOT VALID" in upper:
        return False
    return True


def ai_cache_path(
    note_path: Path,
    provider: str,
    model: str,
    start_date: date | None,
    end_date: date | None,
    purpose: str = "extract",
    extra: str = "",
) -> Path:
    start = start_date.isoformat() if start_date else ""
    end = end_date.isoformat() if end_date else ""
    key = sha1(
        f"{AI_PROMPT_VERSION}|{purpose}|{provider}|{file_sha1(note_path)}|{model}|{start}|{end}|{extra}".encode(
            "utf-8"
        )
    ).hexdigest()
    return AI_CACHE_DIR / f"{key}.json"


def read_ai_cache(
    note_path: Path,
    provider: str,
    model: str,
    start_date: date | None,
    end_date: date | None,
    purpose: str = "extract",
    extra: str = "",
) -> dict[str, Any] | None:
    path = ai_cache_path(note_path, provider, model, start_date, end_date, purpose, extra)
    if not path.exists():
        return None
    try:
        with path.open("r", encoding="utf-8") as file:
            payload = json.load(file)
        return payload if isinstance(payload, dict) else None
    except Exception:
        return None


def write_ai_cache(
    note_path: Path,
    provider: str,
    model: str,
    start_date: date | None,
    end_date: date | None,
    payload: dict[str, Any],
    purpose: str = "extract",
    extra: str = "",
) -> None:
    AI_CACHE_DIR.mkdir(parents=True, exist_ok=True)
    path = ai_cache_path(note_path, provider, model, start_date, end_date, purpose, extra)
    with path.open("w", encoding="utf-8") as file:
        json.dump(payload, file, ensure_ascii=False, indent=2)


def call_openai_receipt_ai(
    note_name: str,
    note_path: Path,
    start_date: date | None,
    end_date: date | None,
    api_key: str,
    model: str,
) -> dict[str, Any]:
    body = {
        "model": model,
        "input": [
            {
                "role": "user",
                "content": [
                    {"type": "input_text", "text": build_receipt_ai_prompt(note_name, start_date, end_date)},
                    {"type": "input_image", "image_url": receipt_image_data_url(note_path), "detail": "high"},
                ],
            }
        ],
        "text": {
            "format": {
                "type": "json_schema",
                "name": "expense_receipt",
                "schema": receipt_ai_json_schema(),
                "strict": True,
            }
        },
        "max_output_tokens": 900,
    }
    request = urllib.request.Request(
        "https://api.openai.com/v1/responses",
        data=json.dumps(body).encode("utf-8"),
        headers={
            "Authorization": f"Bearer {api_key}",
            "Content-Type": "application/json",
        },
        method="POST",
    )
    try:
        with urllib.request.urlopen(request, timeout=120) as response:
            raw = response.read().decode("utf-8", errors="replace")
    except urllib.error.HTTPError as exc:
        details = exc.read().decode("utf-8", errors="replace")
        raise RuntimeError(f"OpenAI retornou HTTP {exc.code}: {details[:500]}") from exc
    except urllib.error.URLError as exc:
        raise RuntimeError(f"Nao foi possivel conectar na OpenAI: {exc.reason}") from exc

    response_payload = json.loads(raw)
    output_text = read_openai_output_text(response_payload)
    if not output_text:
        raise RuntimeError("A OpenAI nao retornou texto estruturado.")
    return parse_json_object(output_text)


def call_openai_receipt_audit(
    note_name: str,
    note_path: Path,
    detection: NoteDetection,
    start_date: date | None,
    end_date: date | None,
    api_key: str,
    model: str,
) -> dict[str, Any]:
    body = {
        "model": model,
        "input": [
            {
                "role": "user",
                "content": [
                    {"type": "input_text", "text": build_receipt_audit_prompt(note_name, detection, start_date, end_date)},
                    {"type": "input_image", "image_url": receipt_image_data_url(note_path), "detail": "high"},
                ],
            }
        ],
        "text": {
            "format": {
                "type": "json_schema",
                "name": "expense_receipt_audit",
                "schema": receipt_audit_json_schema(),
                "strict": True,
            }
        },
        "max_output_tokens": 900,
    }
    request = urllib.request.Request(
        "https://api.openai.com/v1/responses",
        data=json.dumps(body).encode("utf-8"),
        headers={
            "Authorization": f"Bearer {api_key}",
            "Content-Type": "application/json",
        },
        method="POST",
    )
    try:
        with urllib.request.urlopen(request, timeout=120) as response:
            raw = response.read().decode("utf-8", errors="replace")
    except urllib.error.HTTPError as exc:
        details = exc.read().decode("utf-8", errors="replace")
        raise RuntimeError(f"OpenAI retornou HTTP {exc.code}: {details[:500]}") from exc
    except urllib.error.URLError as exc:
        raise RuntimeError(f"Nao foi possivel conectar na OpenAI: {exc.reason}") from exc

    output_text = read_openai_output_text(json.loads(raw))
    if not output_text:
        raise RuntimeError("A OpenAI nao retornou auditoria estruturada.")
    return parse_json_object(output_text)


def test_openai_connection(api_key: str, model: str) -> str:
    body = {
        "model": model,
        "input": "Responda somente OK.",
        "max_output_tokens": 16,
    }
    request = urllib.request.Request(
        "https://api.openai.com/v1/responses",
        data=json.dumps(body).encode("utf-8"),
        headers={
            "Authorization": f"Bearer {api_key}",
            "Content-Type": "application/json",
        },
        method="POST",
    )
    try:
        with urllib.request.urlopen(request, timeout=30) as response:
            raw = response.read().decode("utf-8", errors="replace")
    except urllib.error.HTTPError as exc:
        details = exc.read().decode("utf-8", errors="replace")
        raise RuntimeError(f"OpenAI retornou HTTP {exc.code}: {details[:500]}") from exc
    except urllib.error.URLError as exc:
        raise RuntimeError(f"Nao foi possivel conectar na OpenAI: {exc.reason}") from exc
    return read_openai_output_text(json.loads(raw)) or "OK"


def read_gemini_output_text(payload: dict[str, Any]) -> str:
    candidates = payload.get("candidates")
    if not isinstance(candidates, list):
        return ""

    chunks: list[str] = []
    for candidate in candidates:
        if not isinstance(candidate, dict):
            continue
        content = candidate.get("content")
        if not isinstance(content, dict):
            continue
        parts = content.get("parts")
        if not isinstance(parts, list):
            continue
        for part in parts:
            if not isinstance(part, dict):
                continue
            text = part.get("text")
            if isinstance(text, str) and text.strip():
                chunks.append(text.strip())
    return "\n".join(chunks).strip()


def post_gemini_generate_content(model: str, api_key: str, body: dict[str, Any]) -> dict[str, Any]:
    model_path = model if model.startswith("models/") else f"models/{model}"
    request = urllib.request.Request(
        f"https://generativelanguage.googleapis.com/v1beta/{model_path}:generateContent",
        data=json.dumps(body).encode("utf-8"),
        headers={
            "x-goog-api-key": api_key,
            "Content-Type": "application/json",
        },
        method="POST",
    )
    try:
        with urllib.request.urlopen(request, timeout=120) as response:
            raw = response.read().decode("utf-8", errors="replace")
    except urllib.error.HTTPError as exc:
        details = exc.read().decode("utf-8", errors="replace")
        raise RuntimeError(f"Google AI Studio retornou HTTP {exc.code}: {details[:500]}") from exc
    except urllib.error.URLError as exc:
        raise RuntimeError(f"Nao foi possivel conectar no Google AI Studio: {exc.reason}") from exc
    return json.loads(raw)


def test_gemini_connection(api_key: str, model: str) -> str:
    body = {
        "contents": [{"role": "user", "parts": [{"text": "Responda somente OK."}]}],
        "generationConfig": {"maxOutputTokens": 16},
    }
    payload = post_gemini_generate_content(model, api_key, body)
    return read_gemini_output_text(payload) or "OK"


def test_ai_connection(provider: str, api_key: str, model: str) -> str:
    if provider == "gemini":
        return test_gemini_connection(api_key, model)
    return test_openai_connection(api_key, model)


def call_gemini_receipt_ai(
    note_name: str,
    note_path: Path,
    start_date: date | None,
    end_date: date | None,
    api_key: str,
    model: str,
) -> dict[str, Any]:
    image_data = base64.b64encode(receipt_image_jpeg_bytes(note_path)).decode("ascii")
    parts = [
        {"text": build_receipt_ai_prompt(note_name, start_date, end_date)},
        {"inline_data": {"mime_type": "image/jpeg", "data": image_data}},
    ]
    body = {
        "contents": [{"role": "user", "parts": parts}],
        "generationConfig": {
            "responseMimeType": "application/json",
            "responseJsonSchema": receipt_ai_json_schema(),
            "maxOutputTokens": 900,
        },
    }
    try:
        response_payload = post_gemini_generate_content(model, api_key, body)
    except RuntimeError as exc:
        # Older/stable API variants can reject responseJsonSchema. JSON mode still keeps parsing reliable.
        if "responseJsonSchema" not in str(exc) and "Unknown name" not in str(exc):
            raise
        body["generationConfig"] = {
            "responseMimeType": "application/json",
            "maxOutputTokens": 900,
        }
        response_payload = post_gemini_generate_content(model, api_key, body)

    output_text = read_gemini_output_text(response_payload)
    if not output_text:
        raise RuntimeError("O Google AI Studio nao retornou texto estruturado.")
    return parse_json_object(output_text)


def call_gemini_receipt_audit(
    note_name: str,
    note_path: Path,
    detection: NoteDetection,
    start_date: date | None,
    end_date: date | None,
    api_key: str,
    model: str,
) -> dict[str, Any]:
    image_data = base64.b64encode(receipt_image_jpeg_bytes(note_path)).decode("ascii")
    body = {
        "contents": [
            {
                "role": "user",
                "parts": [
                    {"text": build_receipt_audit_prompt(note_name, detection, start_date, end_date)},
                    {"inline_data": {"mime_type": "image/jpeg", "data": image_data}},
                ],
            }
        ],
        "generationConfig": {
            "responseMimeType": "application/json",
            "responseJsonSchema": receipt_audit_json_schema(),
            "maxOutputTokens": 900,
        },
    }
    try:
        response_payload = post_gemini_generate_content(model, api_key, body)
    except RuntimeError as exc:
        if "responseJsonSchema" not in str(exc) and "Unknown name" not in str(exc):
            raise
        body["generationConfig"] = {
            "responseMimeType": "application/json",
            "maxOutputTokens": 900,
        }
        response_payload = post_gemini_generate_content(model, api_key, body)

    output_text = read_gemini_output_text(response_payload)
    if not output_text:
        raise RuntimeError("O Google AI Studio nao retornou auditoria estruturada.")
    return parse_json_object(output_text)


def note_detection_from_ai(
    note_name: str,
    note_path: Path,
    start_date: date | None,
    end_date: date | None,
    payload: dict[str, Any],
    source: str = "IA OpenAI",
) -> NoteDetection:
    warnings = image_quality_warnings(note_path)
    for item in payload.get("warnings", []):
        warning = repair_text_encoding(item).strip()
        if warning and warning not in warnings:
            warnings.append(warning)
    expense_date = coerce_ai_date(payload.get("expense_date"), start_date, end_date)
    value = parse_money(payload.get("amount"))
    category = normalize_ai_category(payload.get("category"))
    description = repair_text_encoding(payload.get("description") or "").strip()[:80] or Path(note_name).stem[:80] or category
    evidence = repair_text_encoding(payload.get("evidence") or "").strip()
    if category == "Locações de Veículos" and is_vehicle_wash_text(description, evidence, note_name):
        category = "Outras Despesas"
        if "Lavação/lavagem de veículo classificada como Outras Despesas" not in warnings:
            warnings.append("Lavação/lavagem de veículo classificada como Outras Despesas")

    if expense_date is None:
        warnings.append("Data nao encontrada pela IA")
    elif start_date and end_date and not (start_date <= expense_date <= end_date):
        warnings.append("Data fora do periodo da viagem")
    if value <= 0:
        warnings.append("Valor nao encontrado pela IA")
    elif value > Decimal("3000.00"):
        warnings.append("Valor alto; conferir")
    detail_payload = {
        "data": expense_date.isoformat() if expense_date else None,
        "categoria": category,
        "valor": str(value),
        "descricao": description,
        "evidencia": evidence,
        "alertas": warnings,
    }
    return NoteDetection(
        note_name=note_name,
        note_path=note_path,
        expense_date=expense_date,
        category=category,
        description=description,
        value=value,
        confidence=normalize_ai_confidence(payload.get("confidence"), warnings, value, expense_date),
        source=source,
        ocr_text="Extracao por IA:\n" + json.dumps(detail_payload, ensure_ascii=False, indent=2),
        warnings=warnings,
    )


def audit_cache_extra(detection: NoteDetection) -> str:
    payload = {
        "date": detection.expense_date.isoformat() if detection.expense_date else None,
        "category": detection.category,
        "description": detection.description,
        "value": str(detection.value),
        "source": detection.source,
    }
    return sha1(json.dumps(payload, sort_keys=True, ensure_ascii=False).encode("utf-8")).hexdigest()


def apply_audit_payload(
    detection: NoteDetection,
    payload: dict[str, Any],
    start_date: date | None,
    end_date: date | None,
    source: str,
) -> NoteDetection:
    audit_warnings = [repair_text_encoding(item).strip() for item in payload.get("warnings", []) if str(item).strip()]
    expense_date = coerce_ai_date(payload.get("expense_date"), start_date, end_date) or detection.expense_date
    try:
        value = parse_money(payload.get("amount"))
    except Exception:
        value = detection.value
    if value <= 0:
        value = detection.value

    category = normalize_ai_category(payload.get("category") or detection.category)
    description = repair_text_encoding(payload.get("description") or detection.description).strip()[:80]
    approved = bool(payload.get("approved"))
    evidence = repair_text_encoding(payload.get("evidence") or "").strip()
    warnings = list(detection.warnings)
    if category == "Locações de Veículos" and is_vehicle_wash_text(description, evidence, detection.note_name):
        category = "Outras Despesas"
        if "Lavação/lavagem de veículo classificada como Outras Despesas" not in warnings:
            warnings.append("Lavação/lavagem de veículo classificada como Outras Despesas")

    for warning in audit_warnings:
        if warning not in warnings:
            warnings.append(warning)
    if not approved:
        warnings.append("Auditoria da IA pediu conferencia")

    detail = {
        "aprovado": approved,
        "data": expense_date.isoformat() if expense_date else None,
        "categoria": category,
        "valor": str(value),
        "descricao": description,
        "evidencia": evidence,
        "alertas": audit_warnings,
    }
    confidence = normalize_ai_confidence(payload.get("confidence"), warnings, value, expense_date)
    return NoteDetection(
        note_name=detection.note_name,
        note_path=detection.note_path,
        expense_date=expense_date,
        category=category,
        description=description or detection.description,
        value=value,
        confidence=confidence,
        source=f"{detection.source} + auditoria",
        ocr_text=detection.ocr_text
        + "\n\nAuditoria por IA:\n"
        + json.dumps(detail, ensure_ascii=False, indent=2),
        warnings=warnings,
    )


def detect_note_with_ai(
    note_name: str,
    note_path: Path,
    start_date: date | None,
    end_date: date | None,
    provider: str,
    api_key: str,
    model: str,
    fallback_local: bool,
    audit_ai: bool,
) -> NoteDetection:
    last_exc: Exception | None = None
    label = ai_provider_label(provider)
    for candidate_model in ai_model_candidates(provider, model):
        try:
            payload = read_ai_cache(note_path, provider, candidate_model, start_date, end_date)
            if payload is None:
                if provider == "gemini":
                    payload = call_gemini_receipt_ai(note_name, note_path, start_date, end_date, api_key, candidate_model)
                else:
                    payload = call_openai_receipt_ai(note_name, note_path, start_date, end_date, api_key, candidate_model)
                write_ai_cache(note_path, provider, candidate_model, start_date, end_date, payload)
            detection = note_detection_from_ai(
                note_name,
                note_path,
                start_date,
                end_date,
                payload,
                source=f"{label} ({candidate_model})",
            )
            if audit_ai:
                try:
                    extra = audit_cache_extra(detection)
                    audit_payload = read_ai_cache(
                        note_path, provider, candidate_model, start_date, end_date, purpose="audit", extra=extra
                    )
                    if audit_payload is None:
                        if provider == "gemini":
                            audit_payload = call_gemini_receipt_audit(
                                note_name, note_path, detection, start_date, end_date, api_key, candidate_model
                            )
                        else:
                            audit_payload = call_openai_receipt_audit(
                                note_name, note_path, detection, start_date, end_date, api_key, candidate_model
                            )
                        write_ai_cache(
                            note_path,
                            provider,
                            candidate_model,
                            start_date,
                            end_date,
                            audit_payload,
                            purpose="audit",
                            extra=extra,
                        )
                    detection = apply_audit_payload(detection, audit_payload, start_date, end_date, label)
                except Exception as exc:
                    detection.warnings.append(f"Auditoria da IA falhou: {exc}")
                    detection.confidence = "Media" if detection.confidence == "Alta" else detection.confidence
            return detection
        except Exception as exc:
            last_exc = exc
            if not should_try_next_ai_model(exc):
                break

    failure = last_exc or RuntimeError("IA falhou sem detalhes.")
    if fallback_local:
        detection = detect_note(note_name, note_path, start_date, end_date)
        detection.source = "OCR local (IA indisponivel)"
        detection.confidence = "Baixa" if detection.confidence == "Alta" else detection.confidence
        detection.ocr_text = f"IA indisponivel: {failure}\n\nOCR local:\n{detection.ocr_text}"
        return detection
    return NoteDetection(
        note_name=note_name,
        note_path=note_path,
        expense_date=None,
        category="Outras Despesas",
        description=Path(note_name).stem[:80],
        value=Decimal("0"),
        confidence="Baixa",
        source=label,
        ocr_text="",
        warnings=[f"IA falhou: {failure}"],
    )


def load_training_memory() -> dict[str, dict[str, Any]]:
    if not AI_MEMORY_PATH.exists():
        return {}
    try:
        with AI_MEMORY_PATH.open("r", encoding="utf-8") as file:
            payload = json.load(file)
        return payload if isinstance(payload, dict) else {}
    except Exception:
        return {}


def save_training_memory(memory: dict[str, dict[str, Any]]) -> None:
    AI_MEMORY_PATH.parent.mkdir(parents=True, exist_ok=True)
    with AI_MEMORY_PATH.open("w", encoding="utf-8") as file:
        json.dump(memory, file, ensure_ascii=False, indent=2)


def read_memory_detection(note_name: str, note_path: Path) -> NoteDetection | None:
    memory = load_training_memory()
    item = memory.get(file_sha1(note_path))
    if not isinstance(item, dict):
        return None
    try:
        expense_date = to_date(item.get("data"))
        value = parse_money(item.get("valor"))
        category = normalize_ai_category(item.get("categoria"))
    except Exception:
        return None
    if expense_date is None or value <= 0:
        return None

    description = repair_text_encoding(item.get("descricao") or Path(note_name).stem).strip()[:80]
    return NoteDetection(
        note_name=note_name,
        note_path=note_path,
        expense_date=expense_date,
        category=category,
        description=description,
        value=value,
        confidence="Memoria",
        source="Memoria local",
        ocr_text="Dados recuperados da memoria local de correcoes.",
        warnings=[],
    )


def clear_training_memory() -> None:
    if AI_MEMORY_PATH.exists():
        AI_MEMORY_PATH.unlink()


def clear_ai_cache() -> int:
    if not AI_CACHE_DIR.exists():
        return 0
    removed = 0
    for path in AI_CACHE_DIR.glob("*.json"):
        if path.is_file():
            path.unlink()
            removed += 1
    return removed


def confidence_label(expense_date: date | None, value: Decimal, text: str, warnings: list[str]) -> str:
    if any("Valor nao encontrado" in warning for warning in warnings):
        return "Baixa"
    if any("Data nao encontrada" in warning for warning in warnings):
        return "Média" if value > 0 and len(text.strip()) > 20 else "Baixa"
    if warnings:
        return "Média"

    score = 0
    if expense_date:
        score += 1
    if value > 0:
        score += 1
    if len(text.strip()) > 20:
        score += 1
    if score == 3:
        return "Alta"
    if score == 2:
        return "Média"
    return "Baixa"


def detect_note(note_name: str, note_path: Path, start_date: date | None, end_date: date | None) -> NoteDetection:
    warnings: list[str] = image_quality_warnings(note_path)
    try:
        text = ocr_note(note_path)
    except Exception as exc:
        text = ""
        warnings.append(f"OCR falhou: {exc}")

    expense_date = parse_detected_date(text, start_date, end_date)
    value = parse_detected_amount(text)
    category = infer_category(text, note_name)
    description = compact_description(text, note_name, category)

    if expense_date is None:
        warnings.append("Data nao encontrada")
    if value <= 0:
        warnings.append("Valor nao encontrado")
    elif value < Decimal("2.00"):
        warnings.append("Valor muito baixo; conferir")
    elif value > Decimal("3000.00"):
        warnings.append("Valor muito alto; conferir")
    if not text.strip():
        warnings.append("Texto nao encontrado")

    return NoteDetection(
        note_name=note_name,
        note_path=note_path,
        expense_date=expense_date,
        category=category,
        description=description,
        value=value,
        confidence=confidence_label(expense_date, value, text, warnings),
        source="OCR local",
        ocr_text=text,
        warnings=warnings,
    )


def ai_is_ready(options: dict[str, object]) -> bool:
    return bool(options.get("mode") == "ai" and options.get("api_key") and options.get("consent"))


def analyze_notes(
    note_map: dict[str, Path],
    start_date: date | None,
    end_date: date | None,
    options: dict[str, object] | None = None,
) -> list[NoteDetection]:
    options = options or {}
    use_ai = ai_is_ready(options)
    use_memory = bool(options.get("use_memory", True))
    fallback_local = bool(options.get("fallback_local", True))
    audit_ai = bool(options.get("audit_ai", True))
    provider = str(options.get("provider") or "openai")
    api_key = str(options.get("api_key") or "")
    model = str(options.get("model") or (DEFAULT_GEMINI_MODEL if provider == "gemini" else DEFAULT_AI_MODEL))

    detections: list[NoteDetection] = []
    for name, path in sorted(note_map.items()):
        memory_detection = read_memory_detection(name, path) if use_memory else None
        if memory_detection:
            detections.append(memory_detection)
        elif use_ai:
            detections.append(
                detect_note_with_ai(
                    name, path, start_date, end_date, provider, api_key, model, fallback_local, audit_ai
                )
            )
        else:
            detections.append(detect_note(name, path, start_date, end_date))
    return detections


def analyze_notes_cached(
    note_map: dict[str, Path],
    start_date: date | None,
    end_date: date | None,
    options: dict[str, object],
    signature: str,
) -> list[NoteDetection]:
    if st.session_state.get("detections_signature") == signature and "detections" in st.session_state:
        return st.session_state.detections
    detections = analyze_notes(note_map, start_date, end_date, options)
    st.session_state.detections_signature = signature
    st.session_state.detections = detections
    return detections


def detections_to_frame(detections: list[NoteDetection], fallback_date: date) -> pd.DataFrame:
    rows: list[dict[str, object]] = []
    for detection in detections:
        rows.append(
            {
                "usar": True,
                "data": detection.expense_date or fallback_date,
                "categoria": detection.category,
                "descricao": detection.description,
                "valor": float(detection.value),
                "nota": detection.note_name,
                "confianca": detection.confidence,
                "origem": detection.source,
                "alertas": " | ".join(detection.warnings),
            }
        )
    return pd.DataFrame(rows)


def append_alert(current: object, alert: str) -> str:
    alerts = [item.strip() for item in str(current or "").split("|") if item.strip()]
    if alert not in alerts:
        alerts.append(alert)
    return " | ".join(alerts)


def remove_managed_alerts(current: object) -> str:
    alerts = [item.strip() for item in str(current or "").split("|") if item.strip()]
    return " | ".join(alert for alert in alerts if alert not in MANAGED_REVIEW_ALERTS)


def add_review_alerts(
    frame: pd.DataFrame,
    note_map: dict[str, Path] | None = None,
    start_date: date | None = None,
    end_date: date | None = None,
) -> pd.DataFrame:
    if frame.empty:
        return frame

    reviewed = frame.copy()
    if "alertas" not in reviewed.columns:
        reviewed["alertas"] = ""
    reviewed["alertas"] = reviewed["alertas"].apply(remove_managed_alerts)
    notes = reviewed.get("nota", pd.Series(dtype=object)).fillna("").astype(str).str.strip()
    duplicated_notes = {note for note, count in notes.value_counts().items() if note and count > 1}
    duplicate_expenses: set[tuple[str, str, str]] = set()

    normalized_keys: list[tuple[str, str, str]] = []
    for _, row in reviewed.iterrows():
        try:
            item_date = to_iso(row.get("data"))
            value = str(parse_money(row.get("valor")))
        except Exception:
            item_date = ""
            value = ""
        category = repair_text_encoding(row.get("categoria", "")).strip()
        normalized_keys.append((item_date, category, value))

    counts: dict[tuple[str, str, str], int] = defaultdict(int)
    for key in normalized_keys:
        if all(key):
            counts[key] += 1
    duplicate_expenses = {key for key, count in counts.items() if count > 1}

    for position, (index, row) in enumerate(reviewed.iterrows()):
        note = str(row.get("nota", "") or "").strip()
        if note in duplicated_notes:
            reviewed.at[index, "alertas"] = append_alert(
                reviewed.at[index, "alertas"], "Nota usada em mais de uma linha"
            )
        if note_map is not None and (not note or note not in note_map):
            reviewed.at[index, "alertas"] = append_alert(reviewed.at[index, "alertas"], "Nota sem imagem vinculada")
        if normalized_keys[position] in duplicate_expenses:
            reviewed.at[index, "alertas"] = append_alert(
                reviewed.at[index, "alertas"], "Despesa possivelmente duplicada"
            )
        try:
            item_date = to_date(row.get("data"))
        except Exception:
            item_date = None
        if item_date is None:
            reviewed.at[index, "alertas"] = append_alert(reviewed.at[index, "alertas"], "Data invalida")
        elif start_date and end_date and not (start_date <= item_date <= end_date):
            reviewed.at[index, "alertas"] = append_alert(reviewed.at[index, "alertas"], "Data fora do periodo")
        if repair_text_encoding(row.get("categoria", "")).strip() not in CATEGORY_ORDER:
            reviewed.at[index, "alertas"] = append_alert(reviewed.at[index, "alertas"], "Categoria invalida")
        try:
            value = parse_money(row.get("valor"))
        except Exception:
            value = Decimal("0")
        if value <= 0:
            reviewed.at[index, "alertas"] = append_alert(reviewed.at[index, "alertas"], "Valor zerado")
        elif value > Decimal("3000.00"):
            reviewed.at[index, "alertas"] = append_alert(reviewed.at[index, "alertas"], "Valor alto")
        confidence = ascii_key(row.get("confianca", ""))
        source = ascii_key(row.get("origem", ""))
        if "BAIXA" in confidence:
            reviewed.at[index, "alertas"] = append_alert(reviewed.at[index, "alertas"], "Baixa confianca")
        if "IA INDISPONIVEL" in source:
            reviewed.at[index, "alertas"] = append_alert(reviewed.at[index, "alertas"], "IA indisponivel")
        elif source == "OCR LOCAL" or source.startswith("OCR LOCAL "):
            reviewed.at[index, "alertas"] = append_alert(reviewed.at[index, "alertas"], "Leitura por OCR local")
    return reviewed


def review_alerts(frame: pd.DataFrame) -> tuple[list[str], list[str]]:
    blockers: list[str] = []
    warnings: list[str] = []
    if frame.empty or "alertas" not in frame.columns:
        return blockers, warnings

    for _, row in frame.iterrows():
        if not bool(row.get("usar", True)):
            continue
        note = str(row.get("nota", "") or "").strip() or "linha sem nota"
        for alert in [item.strip() for item in str(row.get("alertas", "") or "").split("|") if item.strip()]:
            message = f"{note}: {alert}"
            if alert in BLOCKING_REVIEW_ALERTS:
                blockers.append(message)
            else:
                warnings.append(message)
    return blockers, warnings


def default_expenses() -> pd.DataFrame:
    return pd.DataFrame(
        [
            {
                "data": date.today(),
                "categoria": "Refeições",
                "descricao": "",
                "valor": 0.0,
                "nota": "",
            }
        ]
    )


def default_mileage() -> pd.DataFrame:
    return pd.DataFrame(
        columns=[
            "data",
            "distancia_km",
            "descricao",
        ]
    )


def list_notes_from_folder(folder: Path) -> list[NoteFile]:
    notes: list[NoteFile] = []
    for path in sorted(folder.iterdir()):
        if path.is_file() and path.suffix.lower() in IMAGE_EXTENSIONS:
            notes.append(NoteFile(name=path.name, path=path))
    return notes


def list_folder_notes(extra_folder: str = "") -> list[NoteFile]:
    NOTES_INPUT_DIR.mkdir(exist_ok=True)
    notes = list_notes_from_folder(NOTES_INPUT_DIR)
    extra_folder = repair_text_encoding(extra_folder).strip().strip('"')
    if not extra_folder:
        return notes

    extra_path = Path(extra_folder).expanduser()
    if not extra_path.exists() or not extra_path.is_dir():
        st.warning(f"Pasta adicional não encontrada: {extra_path}")
        return notes

    for note in list_notes_from_folder(extra_path):
        notes.append(NoteFile(name=f"{extra_path.name}_{note.name}", path=note.path))
    return notes


def save_uploaded_notes(uploaded_files: Iterable[object], work_dir: Path) -> list[NoteFile]:
    notes: list[NoteFile] = []
    upload_dir = work_dir / "uploads"
    upload_dir.mkdir(parents=True, exist_ok=True)

    for uploaded in uploaded_files:
        data = uploaded.getbuffer()
        original_name = re.sub(r'[<>:"/\\|?*]+', "_", uploaded.name).strip() or "nota.jpg"
        stem = Path(original_name).stem or "nota"
        suffix = Path(original_name).suffix.lower() or ".jpg"
        full_digest = sha1(data).hexdigest()
        digest = full_digest[:10]
        safe_name = f"{stem}_{digest}{suffix}"
        target = upload_dir / safe_name
        should_write = True
        if target.exists():
            try:
                should_write = file_sha1(target) != full_digest
            except OSError:
                should_write = True
        if should_write:
            atomic_write_bytes(target, bytes(data))
        notes.append(NoteFile(name=safe_name, path=target))

    return notes


def safe_mobile_member_name(value: object) -> str:
    text = repair_text_encoding(value).replace("\\", "/").strip()
    if not text or text.startswith("/") or re.match(r"^[A-Za-z]:", text):
        return ""
    raw_parts = text.split("/")
    if any(part == ".." for part in raw_parts):
        return ""
    parts = [part for part in raw_parts if part not in {"", "."}]
    return "/".join(parts)


def import_mobile_package(uploaded_file: object, work_dir: Path) -> tuple[dict[str, Path], list[NoteDetection]]:
    package_bytes = bytes(uploaded_file.getbuffer())
    if not package_bytes:
        raise ValueError("O arquivo ZIP está vazio.")
    if len(package_bytes) > MAX_MOBILE_PACKAGE_BYTES:
        raise ValueError(f"O pacote excede o limite de {MAX_MOBILE_PACKAGE_BYTES // 1024 // 1024} MB.")
    package_digest = sha1(package_bytes).hexdigest()
    package_dir = work_dir / "mobile_packages" / package_digest
    package_dir.mkdir(parents=True, exist_ok=True)

    try:
        archive_context = zipfile.ZipFile(BytesIO(package_bytes))
    except zipfile.BadZipFile as exc:
        raise ValueError("O arquivo enviado não é um ZIP válido ou está corrompido.") from exc

    with archive_context as archive:
        infos = archive.infolist()
        if len(infos) > MAX_ZIP_MEMBERS:
            raise ValueError(f"O pacote possui mais de {MAX_ZIP_MEMBERS} arquivos.")
        if sum(info.file_size for info in infos) > MAX_MOBILE_UNCOMPRESSED_BYTES:
            raise ValueError(f"O conteúdo descompactado excede {MAX_MOBILE_UNCOMPRESSED_BYTES // 1024 // 1024} MB.")
        if any(info.flag_bits & 0x1 for info in infos):
            raise ValueError("Pacotes ZIP protegidos por senha não são aceitos.")
        corrupted_member = archive.testzip()
        if corrupted_member:
            raise ValueError(f"O ZIP está corrompido no arquivo: {corrupted_member}")

        member_map: dict[str, zipfile.ZipInfo] = {}
        for info in infos:
            if info.is_dir():
                continue
            safe_name = safe_mobile_member_name(info.filename)
            if not safe_name:
                raise ValueError(f"Nome de arquivo inseguro ou inválido no ZIP: {info.filename}")
            if safe_name in member_map:
                raise ValueError(f"Arquivo duplicado no ZIP: {safe_name}")
            member_map[safe_name] = info

        manifest_names = [
            name for name in member_map if name.lower().endswith(MOBILE_PACKAGE_MANIFEST.lower())
        ]
        if not manifest_names:
            raise ValueError(f"Pacote móvel inválido: arquivo {MOBILE_PACKAGE_MANIFEST} não encontrado.")
        if len(manifest_names) > 1:
            raise ValueError(f"Pacote móvel inválido: existem vários arquivos {MOBILE_PACKAGE_MANIFEST}.")
        manifest_info = member_map[manifest_names[0]]
        if manifest_info.file_size > 2 * 1024 * 1024:
            raise ValueError("O manifesto do pacote é grande demais.")
        try:
            manifest = json.loads(archive.read(manifest_info).decode("utf-8-sig"))
        except (UnicodeDecodeError, json.JSONDecodeError) as exc:
            raise ValueError("O manifesto do pacote está inválido ou corrompido.") from exc
        if not isinstance(manifest, dict) or manifest.get("format") != "relatorio-despesas-mobile":
            raise ValueError("Pacote móvel inválido ou incompatível.")
        if manifest.get("version") not in (1, 2):
            raise ValueError("Versão do pacote móvel não suportada.")

        note_map: dict[str, Path] = {}
        detections: list[NoteDetection] = []
        entries = manifest.get("expenses", [])
        if not isinstance(entries, list):
            raise ValueError("Pacote móvel sem lista de despesas.")
        if not entries:
            raise ValueError("O pacote móvel não contém despesas.")
        if len(entries) > MAX_MOBILE_ENTRIES:
            raise ValueError(f"O pacote excede o limite de {MAX_MOBILE_ENTRIES} despesas.")
        if manifest.get("expense_count") not in (None, len(entries)):
            raise ValueError("A contagem de despesas do pacote não confere.")

        used_photos: set[str] = set()
        used_sequences: set[int] = set()
        for position, item in enumerate(entries, start=1):
            if not isinstance(item, dict):
                raise ValueError(f"Despesa {position} possui formato inválido.")
            try:
                sequence = int(item.get("sequence", position))
            except (TypeError, ValueError) as exc:
                raise ValueError(f"Despesa {position} possui sequência inválida.") from exc
            if sequence in used_sequences:
                raise ValueError(f"Sequência duplicada no pacote: {sequence}")
            used_sequences.add(sequence)
            member_name = safe_mobile_member_name(item.get("photo"))
            if not member_name:
                raise ValueError(f"Despesa {position} sem foto vinculada.")
            if member_name in used_photos:
                raise ValueError(f"Foto usada em mais de uma despesa: {member_name}")
            used_photos.add(member_name)
            member_info = member_map.get(member_name)
            if member_info is None:
                raise ValueError(f"Foto não encontrada no pacote: {member_name}")

            member_path = Path(member_name)
            suffix = member_path.suffix.lower()
            if suffix not in IMAGE_EXTENSIONS:
                raise ValueError(f"Formato de imagem não suportado: {member_name}")
            note_name = f"{position:03d}_{slugify(member_path.stem, 'NOTA')}{suffix}"
            target = package_dir / note_name
            image_bytes = archive.read(member_info)
            expected_size = item.get("photo_size")
            if expected_size not in (None, "") and int(expected_size) != len(image_bytes):
                raise ValueError(f"O tamanho da foto não confere: {member_name}")
            expected_sha256 = str(item.get("photo_sha256") or "").strip().lower()
            if expected_sha256 and sha256(image_bytes).hexdigest() != expected_sha256:
                raise ValueError(f"A verificação de integridade da foto falhou: {member_name}")
            validate_image_bytes(image_bytes, member_name)
            if not target.exists() or file_sha1(target) != sha1(image_bytes).hexdigest():
                atomic_write_bytes(target, image_bytes)

            try:
                expense_date = to_date(item.get("date"))
                amount = parse_money(item.get("amount"))
            except Exception as exc:
                raise ValueError(f"Despesa {position} possui data ou valor inválido.") from exc
            category = normalize_ai_category(item.get("category"))
            description = repair_text_encoding(item.get("description", "")).strip()[:80] or category
            if expense_date is None:
                raise ValueError(f"Despesa {position} não possui data.")
            if amount <= 0:
                raise ValueError(f"Despesa {position} não possui valor válido.")

            note_map[note_name] = target
            detections.append(
                NoteDetection(
                    note_name=note_name,
                    note_path=target,
                    expense_date=expense_date,
                    category=category,
                    description=description,
                    value=amount,
                    confidence="Manual",
                    source="Coletor móvel",
                    ocr_text="Lançamento informado manualmente no coletor móvel.",
                    warnings=[],
                )
            )

    return note_map, detections


def normalize_expenses(frame: pd.DataFrame) -> pd.DataFrame:
    if frame.empty:
        return frame

    normalized = frame.copy()
    normalized["data"] = normalized["data"].apply(to_date)
    normalized["categoria"] = normalized["categoria"].fillna("").apply(repair_text_encoding).astype(str).str.strip()
    normalized["descricao"] = normalized["descricao"].fillna("").apply(repair_text_encoding).astype(str).str.strip()
    normalized["nota"] = normalized["nota"].fillna("").apply(repair_text_encoding).astype(str).str.strip()
    normalized["valor_decimal"] = normalized["valor"].apply(parse_money)
    normalized = normalized[normalized["data"].notna()]
    normalized = normalized[normalized["categoria"].isin(CATEGORY_ORDER)]
    normalized = normalized[normalized["valor_decimal"] > 0]
    normalized["ordem"] = range(len(normalized))
    normalized["categoria_ordem"] = normalized["categoria"].map(CATEGORY_ORDER.index)
    return normalized.sort_values(["data", "categoria_ordem", "ordem"]).reset_index(drop=True)


def remember_review_corrections(frame: pd.DataFrame, note_map: dict[str, Path]) -> int:
    if frame.empty:
        return 0

    memory = load_training_memory()
    learned = 0
    for _, row in frame.iterrows():
        note_name = str(row.get("nota", "") or "").strip()
        if not note_name or note_name not in note_map:
            continue
        try:
            expense_date = to_date(row.get("data"))
            value = parse_money(row.get("valor"))
            category = normalize_ai_category(row.get("categoria"))
        except Exception:
            continue
        if expense_date is None or value <= 0:
            continue

        memory[file_sha1(note_map[note_name])] = {
            "nota": note_name,
            "data": expense_date.isoformat(),
            "categoria": category,
            "descricao": repair_text_encoding(row.get("descricao", "") or "").strip()[:80],
            "valor": str(value),
            "updated_at": datetime.now().isoformat(timespec="seconds"),
            "version": AI_PROMPT_VERSION,
        }
        learned += 1

    if learned:
        save_training_memory(memory)
    return learned


def normalize_mileage(frame: pd.DataFrame) -> list[dict[str, object]]:
    if frame.empty:
        return []

    items: list[dict[str, object]] = []
    for _, row in frame.iterrows():
        item_date = to_iso(row.get("data"))
        description = str(row.get("descricao", "") or "").strip()
        distance_raw = row.get("distancia_km")
        distance = parse_money(distance_raw)
        if not item_date and not description and distance == 0:
            continue
        if not item_date or distance <= 0:
            raise ValueError("Toda linha de quilometragem precisa ter data e distância maior que zero.")
        items.append(
            {
                "date": item_date,
                "distance": float(distance),
                "description": description,
            }
        )

    if len(items) > MAX_MILEAGE_ROWS:
        raise ValueError(f"O modelo aceita no máximo {MAX_MILEAGE_ROWS} linhas de quilometragem.")

    return items


def excel_sum_formula(amounts: list[Decimal]) -> str | None:
    clean_amounts = [amount.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP) for amount in amounts if amount > 0]
    if len(clean_amounts) <= 1:
        return None
    pieces = [format(amount, "f") for amount in clean_amounts]
    return "=" + "+".join(pieces)


def build_payload(
    output_dir: Path,
    base_name: str,
    employee: dict[str, str],
    trip: dict[str, object],
    expenses: pd.DataFrame,
    mileage: list[dict[str, object]],
    km_rate: Decimal,
    advance: Decimal,
) -> dict[str, object]:
    unique_dates = sorted(expenses["data"].unique().tolist())
    if len(unique_dates) > MAX_DATES:
        raise ValueError(f"O modelo aceita no máximo {MAX_DATES} datas diferentes de despesas.")

    date_to_column = {expense_date: DATE_COLUMNS[index] for index, expense_date in enumerate(unique_dates)}
    grouped: dict[tuple[str, str], list[Decimal]] = defaultdict(list)
    for _, row in expenses.iterrows():
        grouped[(EXCEL_CATEGORY_KEYS[row["categoria"]], row["data"].isoformat())].append(row["valor_decimal"])

    expense_cells = [
        {
            "category": category,
            "column": date_to_column[date.fromisoformat(date_text)],
            "value": float(sum(amounts, Decimal("0"))),
            "formula": excel_sum_formula(amounts),
            "components": [float(amount) for amount in amounts],
        }
        for (category, date_text), amounts in grouped.items()
        if sum(amounts, Decimal("0")) > 0
    ]

    other_expense_rows: list[dict[str, object]] = []
    other_expenses = expenses[expenses["categoria"] == "Outras Despesas"]
    if len(other_expenses) > MAX_OTHER_EXPENSES:
        raise ValueError(f"O modelo aceita no máximo {MAX_OTHER_EXPENSES} descrições em Outras Despesas.")

    for _, row in other_expenses.iterrows():
        other_expense_rows.append(
            {
                "date": row["data"].isoformat(),
                "description": row["descricao"],
                "value": float(row["valor_decimal"]),
            }
        )

    return {
        "template_path": str(TEMPLATE_PATH),
        "output_dir": str(output_dir),
        "xlsx_path": str(output_dir / f"{base_name}.xlsx"),
        "report_pdf_path": "",
        "employee": employee,
        "trip": {
            "start_date": to_iso(trip["start_date"]),
            "end_date": to_iso(trip["end_date"]),
            "report_date": to_iso(trip["report_date"]),
            "reason": str(trip["reason"]).strip(),
        },
        "dates": [expense_date.isoformat() for expense_date in unique_dates],
        "expense_cells": expense_cells,
        "other_expenses": other_expense_rows,
        "mileage": mileage,
        "km_rate": float(km_rate),
        "advance": float(advance),
    }


def run_excel_generator(payload: dict[str, object], output_dir: Path) -> None:
    _ = output_dir
    generate_excel_report(payload)


def clear_legacy_output_files(output_dir: Path, base_name: str) -> None:
    for path in (
        output_dir / f"{base_name}.pdf",
        output_dir / f"{base_name}.zip",
        output_dir / "dados_relatorio.json",
    ):
        path.unlink(missing_ok=True)


def fit_image_size(image_width: int, image_height: int, max_width: float, max_height: float) -> tuple[float, float]:
    ratio = min(max_width / image_width, max_height / image_height)
    return image_width * ratio, image_height * ratio


def display_note_image(path: Path, width: int = 340) -> None:
    try:
        with Image.open(path) as image:
            st.image(image.convert("RGB"), width=width)
    except Exception:
        st.image(str(path), width=width)


def generate_notes_pdf(expenses: pd.DataFrame, note_map: dict[str, Path], output_path: Path) -> list[str]:
    pdf = FPDF(unit="mm", format="A4")
    pdf.set_auto_page_break(auto=False)

    missing: list[str] = []
    with tempfile.TemporaryDirectory() as temp_dir:
        temp_path = Path(temp_dir)

        if expenses.empty:
            pdf.add_page()
            pdf.set_font("Arial", "B", 14)
            pdf.cell(0, 10, "Notas de despesas", ln=True)
            pdf.set_font("Arial", "", 11)
            pdf.multi_cell(0, 7, "Nenhuma despesa com nota foi informada.")
        else:
            for index, row in expenses.iterrows():
                note_name = str(row.get("nota", "") or "").strip()
                title = f"{row['data'].strftime('%d/%m/%Y')} - {row['categoria']} - {brl(row['valor_decimal'])}"
                description = str(row.get("descricao", "") or "").strip()

                pdf.add_page()
                pdf.set_font("Arial", "B", 12)
                pdf.cell(0, 7, title.encode("latin-1", "replace").decode("latin-1"), ln=True)
                if description:
                    pdf.set_font("Arial", "", 10)
                    pdf.multi_cell(0, 5, description.encode("latin-1", "replace").decode("latin-1"))
                pdf.ln(2)

                if not note_name or note_name not in note_map:
                    missing.append(title)
                    pdf.set_font("Arial", "", 11)
                    pdf.multi_cell(0, 7, "Nota nao anexada.")
                    continue

                try:
                    with Image.open(note_map[note_name]) as image:
                        image = ImageOps.exif_transpose(image)
                        image = image.convert("RGB")
                        image_path = temp_path / f"nota_{index}.jpg"
                        image.save(image_path, "JPEG", quality=92)
                        width, height = fit_image_size(image.width, image.height, 190, 254)
                        x = (210 - width) / 2
                        y = max(pdf.get_y(), 32)
                        pdf.image(str(image_path), x=x, y=y, w=width, h=height)
                except Exception as exc:
                    missing.append(f"{title} ({exc})")
                    pdf.set_font("Arial", "", 11)
                    pdf.multi_cell(0, 7, "Nao foi possivel inserir esta imagem no PDF.")

        temp_output = output_path.with_name(f".{output_path.name}.{uuid4().hex}.tmp.pdf")
        try:
            pdf.output(str(temp_output))
            pdf_bytes = temp_output.read_bytes()
            if len(pdf_bytes) < 100 or not pdf_bytes.startswith(b"%PDF-"):
                raise ValueError("O PDF gerado está vazio ou corrompido.")
            temp_output.replace(output_path)
        finally:
            temp_output.unlink(missing_ok=True)

    return missing


def make_zip(output_dir: Path, files: list[Path]) -> Path:
    zip_path = output_dir / f"{output_dir.name}.zip"
    with zipfile.ZipFile(zip_path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        for file in files:
            archive.write(file, arcname=file.name)
    return zip_path


def make_download_zip(files: list[Path]) -> bytes:
    for file in files:
        if not file.exists() or file.stat().st_size == 0:
            raise ValueError(f"Arquivo final ausente ou vazio: {file.name}")
    buffer = BytesIO()
    with zipfile.ZipFile(buffer, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        for file in files:
            archive.write(file, arcname=file.name)
    zip_bytes = buffer.getvalue()
    with zipfile.ZipFile(BytesIO(zip_bytes)) as archive:
        if archive.testzip():
            raise ValueError("O pacote final apresentou erro de integridade.")
        if sorted(archive.namelist()) != sorted(file.name for file in files):
            raise ValueError("O pacote final não contém todos os arquivos esperados.")
    return zip_bytes


def render_header() -> None:
    logo_data = ""
    if BRAND_LOGO_PATH.exists():
        logo_data = base64.b64encode(BRAND_LOGO_PATH.read_bytes()).decode("ascii")
    logo_html = f'<img src="data:image/svg+xml;base64,{logo_data}" alt="Agres">' if logo_data else ""
    st.markdown(
        f"""
        <div class="brand-header">
            {logo_html}
            <div>
                <h1>Relatório de Despesas</h1>
                <p>Conferência e geração dos documentos para prestação de contas.</p>
            </div>
            <span class="product-badge">Fluxo Offline</span>
        </div>
        <div class="workflow">
            <div class="workflow-item"><span>1</span>Importar Arquivo</div>
            <div class="workflow-item"><span>2</span>Dados da Viagem</div>
            <div class="workflow-item"><span>3</span>Conferir Despesas</div>
            <div class="workflow-item"><span>4</span>Gerar Documentos</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_flow_step(number: int, title: str, help_text: str = "") -> None:
    help_html = f"<small>{help_text}</small>" if help_text else ""
    st.markdown(
        f'<div class="flow-step"><span>{number}</span><div><strong>{title}</strong>{help_html}</div></div>',
        unsafe_allow_html=True,
    )


def configured_value(section: str, key: str, environment_name: str) -> str:
    try:
        section_values = st.secrets.get(section, {})
        if hasattr(section_values, "get") and section_values.get(key):
            return str(section_values[key]).strip()
    except Exception:
        pass
    return os.getenv(environment_name, "").strip()


def require_access() -> None:
    expected = configured_value("security", "app_password", "APP_PASSWORD")
    if not expected:
        return
    if st.session_state.get("access_granted"):
        return

    st.markdown("### Acesso ao Gerador")
    supplied = st.text_input("Senha de Acesso", type="password", key="access_password")
    if st.button("Entrar", type="primary", width="stretch"):
        if hmac.compare_digest(supplied, expected):
            st.session_state["access_granted"] = True
            st.rerun()
        st.error("Senha inválida.")
    st.stop()


def cleanup_stale_directories(root: Path, max_age_seconds: int = STALE_WORKSPACE_SECONDS) -> None:
    if not root.exists():
        return
    cutoff = time.time() - max_age_seconds
    for path in root.iterdir():
        try:
            if path.is_dir() and path.stat().st_mtime < cutoff:
                shutil.rmtree(path, ignore_errors=True)
        except OSError:
            continue


def render_employee_form(expense_dates: list[date]) -> tuple[dict[str, str], dict[str, object], str]:
    render_flow_step(2, "Dados da Viagem", "Informe o cliente, o local e o motivo. O período é preenchido pelos lançamentos importados.")

    col_a, col_b = st.columns([1.2, 1])
    with col_a:
        place = st.text_input(
            "Cliente / Local",
            placeholder="Ex.: Herbicat - Bauru/SP",
            key="report_place",
        )
    with col_b:
        motive = st.text_input(
            "Motivo",
            placeholder="Ex.: Suporte técnico",
            key="report_motive",
        )

    imported_start = min(expense_dates) if expense_dates else date.today()
    imported_end = max(expense_dates) if expense_dates else date.today()

    with st.expander("Configurações do Relatório", expanded=False, icon=":material/settings:"):
        report_date = st.date_input("Data de Emissão", value=imported_end)
        auto_dates = st.checkbox("Usar Período dos Lançamentos Importados", value=True)
        if auto_dates:
            start_date = imported_start
            end_date = imported_end
            st.caption(f"Período: {start_date.strftime('%d/%m/%Y')} a {end_date.strftime('%d/%m/%Y')}")
        else:
            col_c, col_d = st.columns([1, 1])
            with col_c:
                start_date = st.date_input("Data de Saída", value=imported_start)
            with col_d:
                end_date = st.date_input("Data de Chegada", value=imported_end)

    reason_parts = [part.strip() for part in (motive, place) if part and part.strip()]
    reason = " - ".join(reason_parts)

    with st.expander("Dados do Colaborador e Pagamento", expanded=False, icon=":material/account_balance:"):
        col_1, col_2, col_3 = st.columns([1.2, 1, .75])
        with col_1:
            name = st.text_input("Nome", value=configured_value("employee", "name", "EMPLOYEE_NAME"))
        with col_2:
            cpf = st.text_input("CPF", value=configured_value("employee", "cpf", "EMPLOYEE_CPF"))
        with col_3:
            cost_center = st.text_input(
                "Centro de Custo",
                value=configured_value("employee", "cost_center", "EMPLOYEE_COST_CENTER"),
            )

        col_4, col_5, col_6 = st.columns([1.2, .75, .9])
        with col_4:
            bank = st.text_input("Banco", value=configured_value("employee", "bank", "EMPLOYEE_BANK"))
        with col_5:
            agency = st.text_input("Agência", value=configured_value("employee", "agency", "EMPLOYEE_AGENCY"))
        with col_6:
            account = st.text_input(
                "Conta Corrente",
                value=configured_value("employee", "account", "EMPLOYEE_ACCOUNT"),
            )

    employee = {
        "name": name,
        "cpf": cpf,
        "cost_center": cost_center,
        "bank": bank,
        "agency": agency,
        "account": account,
    }
    trip = {
        "reason": reason,
        "start_date": start_date,
        "end_date": end_date,
        "report_date": report_date,
        "auto_dates": auto_dates,
    }

    generated_name = f"{report_date.strftime('%Y%m%d')}_DESPESAS_{slugify(place)}_{slugify(motive)}"
    base_name = generated_name
    return employee, trip, base_name


def render_notes_inputs(work_dir: Path) -> tuple[dict[str, Path], list[NoteDetection]]:
    render_flow_step(
        1,
        "Importar Arquivo do Coletor",
        "Selecione o ZIP gerado no iPhone ou iPad para carregar lançamentos e comprovantes.",
    )
    mobile_package = st.file_uploader(
        "Selecionar Arquivo ZIP",
        type=["zip"],
        accept_multiple_files=False,
        key="mobile_package_upload",
        help="Formato esperado: AAAAMMDD_RELATÓRIO_DESPESAS_HENRIQUE.zip",
    )
    mobile_notes: dict[str, Path] = {}
    mobile_detections: list[NoteDetection] = []
    if mobile_package is not None:
        try:
            mobile_notes, mobile_detections = import_mobile_package(mobile_package, work_dir)
            total = sum((detection.value for detection in mobile_detections), Decimal("0"))
            valid_dates = sorted(detection.expense_date for detection in mobile_detections if detection.expense_date)
            period = (
                f"{valid_dates[0].strftime('%d/%m/%Y')} a {valid_dates[-1].strftime('%d/%m/%Y')}"
                if valid_dates
                else "Não informado"
            )
            st.markdown(
                f"""
                <div class="import-status">
                    <strong>Arquivo Carregado com Sucesso</strong>
                    <span>{html.escape(repair_text_encoding(mobile_package.name))}</span>
                </div>
                """,
                unsafe_allow_html=True,
            )
            col_a, col_b, col_c, col_d = st.columns([1.1, .75, 1.65, .8])
            col_a.metric("Valor Total", brl(total), border=True)
            col_b.metric("Lançamentos", len(mobile_detections), border=True)
            col_c.metric("Período", period, border=True)
            col_d.metric("Comprovantes", len(mobile_notes), border=True)
        except Exception as exc:
            st.error(str(exc))

    note_map: dict[str, Path] = dict(mobile_notes)

    if note_map:
        st.caption("Os comprovantes serão organizados automaticamente na mesma sequência do relatório.")
    else:
        st.info("Selecione o arquivo ZIP para iniciar o fechamento.", icon=":material/upload_file:")

    return note_map, mobile_detections


def render_ai_settings() -> dict[str, object]:
    st.subheader("Leitura")
    mode_label = st.radio(
        "Motor de leitura",
        ["Google AI Studio", "OpenAI", "OCR local"],
        horizontal=True,
        index=0,
    )
    provider = "gemini" if mode_label == "Google AI Studio" else "openai"
    options: dict[str, object] = {
        "mode": "ocr" if mode_label == "OCR local" else "ai",
        "provider": provider,
        "api_key": "",
        "model": DEFAULT_GEMINI_MODEL if provider == "gemini" else DEFAULT_AI_MODEL,
        "consent": False,
        "fallback_local": True,
        "use_memory": True,
        "audit_ai": True,
    }

    if options["mode"] == "ocr":
        st.caption("OCR local selecionado. Nenhuma imagem sera enviada para fora deste computador.")
        return options

    if provider == "gemini":
        env_key = (os.getenv("GEMINI_API_KEY", "") or os.getenv("GOOGLE_API_KEY", "")).strip()
        env_model = os.getenv("GEMINI_RECEIPT_MODEL", DEFAULT_GEMINI_MODEL).strip() or DEFAULT_GEMINI_MODEL
        key_label = "Google AI Studio API key"
        key_placeholder = "Cole a chave aqui ou configure GEMINI_API_KEY"
        found_key_label = "GEMINI_API_KEY encontrada no ambiente."
    else:
        env_key = os.getenv("OPENAI_API_KEY", "").strip()
        env_model = os.getenv("OPENAI_RECEIPT_MODEL", DEFAULT_AI_MODEL).strip() or DEFAULT_AI_MODEL
        key_label = "OpenAI API key"
        key_placeholder = "Cole a chave aqui ou configure OPENAI_API_KEY"
        found_key_label = "OPENAI_API_KEY encontrada no ambiente."

    with st.expander("Configuracao da IA", expanded=True):
        if env_key:
            st.caption(found_key_label)
        api_key_input = st.text_input(
            key_label,
            type="password",
            placeholder=key_placeholder,
        )
        model = st.text_input("Modelo", value=env_model)
        consent = st.checkbox("Autorizo enviar as imagens das notas para leitura por IA.", value=False)
        audit_ai = st.checkbox("Auditar cada nota com segunda passagem da IA", value=True)
        fallback_local = st.checkbox("Se a IA falhar, tentar OCR local", value=True)
        use_memory = st.checkbox("Usar memoria local das minhas correcoes", value=True)
        test_key = st.button("Testar chave IA")

    api_key = (api_key_input or env_key).strip()
    default_model = DEFAULT_GEMINI_MODEL if provider == "gemini" else DEFAULT_AI_MODEL
    model = model.strip() or default_model
    if test_key:
        if not api_key:
            st.error(f"Informe a {key_label} antes de testar.")
        else:
            try:
                reply = test_ai_connection(provider, api_key, model)
                st.success(f"Conexao OK. Resposta: {reply[:80]}")
            except Exception as exc:
                st.error(str(exc))

    options.update(
        {
            "api_key": api_key,
            "model": model,
            "consent": consent,
            "fallback_local": fallback_local,
            "use_memory": use_memory,
            "audit_ai": audit_ai,
        }
    )

    if not api_key:
        st.warning(f"Informe a {key_label} para ativar a leitura por IA.")
    elif not consent:
        st.warning("Marque a autorizacao para enviar as imagens das notas para a IA.")
    return options


def render_memory_tools() -> None:
    with st.expander("Memória e reprocessamento", expanded=False):
        memory = load_training_memory()
        col_a, col_b, col_c = st.columns([1, 1, 1])
        col_a.metric("Correções salvas", len(memory))
        cache_count = len(list(AI_CACHE_DIR.glob("*.json"))) if AI_CACHE_DIR.exists() else 0
        col_b.metric("Leituras em cache", cache_count)
        with col_c:
            st.write("")
            st.write("")
            if st.button("Limpar cache de IA"):
                removed = clear_ai_cache()
                st.success(f"Cache limpo: {removed} leitura(s).")
                st.rerun()
        if memory:
            preview_rows = []
            for key, item in memory.items():
                if not isinstance(item, dict):
                    continue
                preview_rows.append(
                    {
                        "apagar": False,
                        "chave": key,
                        "nota": repair_text_encoding(item.get("nota", "")),
                        "data": item.get("data", ""),
                        "categoria": normalize_ai_category(item.get("categoria")),
                        "descricao": repair_text_encoding(item.get("descricao", "")),
                        "valor": item.get("valor", ""),
                    }
                )
            if preview_rows:
                edited_memory = st.data_editor(
                    pd.DataFrame(preview_rows),
                    hide_index=True,
                    width="stretch",
                    column_config={
                        "apagar": st.column_config.CheckboxColumn("Apagar"),
                        "chave": st.column_config.TextColumn("Chave"),
                        "nota": st.column_config.TextColumn("Nota"),
                        "data": st.column_config.TextColumn("Data"),
                        "categoria": st.column_config.SelectboxColumn("Categoria", options=CATEGORY_ORDER),
                        "descricao": st.column_config.TextColumn("Descrição"),
                        "valor": st.column_config.TextColumn("Valor"),
                    },
                    disabled=["chave", "nota"],
                    key="memory_editor",
                )
                if st.button("Salvar ajustes da memória"):
                    updated: dict[str, dict[str, Any]] = {}
                    skipped = 0
                    for _, row in edited_memory.iterrows():
                        key = str(row.get("chave", "") or "").strip()
                        if not key or bool(row.get("apagar", False)):
                            continue
                        try:
                            expense_date = to_date(row.get("data"))
                            value = parse_money(row.get("valor"))
                            category = normalize_ai_category(row.get("categoria"))
                        except Exception:
                            skipped += 1
                            continue
                        if expense_date is None or value <= 0:
                            skipped += 1
                            continue
                        original = memory.get(key, {}) if isinstance(memory.get(key), dict) else {}
                        updated[key] = {
                            **original,
                            "nota": repair_text_encoding(row.get("nota", "")),
                            "data": expense_date.isoformat(),
                            "categoria": category,
                            "descricao": repair_text_encoding(row.get("descricao", "")).strip()[:80],
                            "valor": str(value),
                            "updated_at": datetime.now().isoformat(timespec="seconds"),
                            "version": AI_PROMPT_VERSION,
                        }
                    save_training_memory(updated)
                    if skipped:
                        st.warning(f"Memória salva, mas {skipped} linha(s) inválida(s) foram ignoradas.")
                    else:
                        st.success("Memória salva.")
                    st.rerun()
            if st.button("Apagar memória de correções"):
                clear_training_memory()
                st.success("Memória apagada.")
                st.rerun()


def render_recognition_test(
    note_map: dict[str, Path],
    options: dict[str, object],
    start_date: date | None,
    end_date: date | None,
) -> None:
    with st.expander("Teste de reconhecimento", expanded=False):
        if not note_map:
            st.info("Envie fotos ou coloque imagens em entrada_notas para testar.")
            return
        if options.get("mode") == "ai" and not ai_is_ready(options):
            st.warning("Informe a chave, marque a autorizacao e depois rode o teste.")
            return

        col_a, col_b = st.columns([1, 1])
        max_notes = min(len(note_map), 5)
        with col_a:
            sample_size = int(
                st.number_input("Quantidade de notas no teste", min_value=1, max_value=max_notes, value=min(3, max_notes))
            )
        with col_b:
            force_ai = st.checkbox("Ignorar memoria neste teste", value=True)
        run_test = st.button("Rodar teste de reconhecimento")

        if not run_test:
            return

        sample = dict(list(sorted(note_map.items()))[:sample_size])
        test_options = dict(options)
        if force_ai:
            test_options["use_memory"] = False
        with st.spinner("Testando reconhecimento nas notas selecionadas..."):
            detections = analyze_notes(sample, start_date, end_date, test_options)

        test_frame = add_review_alerts(
            detections_to_frame(detections, start_date or date.today()),
            sample,
            start_date,
            end_date,
        )
        blockers, warnings = review_alerts(test_frame)
        if blockers:
            st.error(f"Teste concluido com {len(blockers)} bloqueio(s).")
        elif warnings:
            st.warning(f"Teste concluido com {len(warnings)} aviso(s).")
        else:
            st.success("Teste concluido sem alertas.")
        st.dataframe(test_frame, hide_index=True, width="stretch")
        render_visual_review(test_frame, detections, sample, start_date, end_date, "recognition_test")


def render_expenses_editor(note_names: list[str]) -> pd.DataFrame:
    st.subheader("Despesas")
    if "expenses_df" not in st.session_state:
        st.session_state.expenses_df = default_expenses()

    return st.data_editor(
        st.session_state.expenses_df,
        num_rows="dynamic",
        width="stretch",
        hide_index=True,
        column_config={
            "data": st.column_config.DateColumn("Data do gasto", format="DD/MM/YYYY", required=True),
            "categoria": st.column_config.SelectboxColumn("Categoria", options=CATEGORY_ORDER, required=True),
            "descricao": st.column_config.TextColumn("Descrição da nota"),
            "valor": st.column_config.NumberColumn("Valor", min_value=0.0, step=0.01, format="R$ %.2f", required=True),
            "nota": st.column_config.SelectboxColumn("Foto da nota", options=[""] + note_names),
        },
        key="expenses_editor",
    )


def render_visual_review(
    edited: pd.DataFrame,
    detections: list[NoteDetection],
    note_map: dict[str, Path] | None = None,
    start_date: date | None = None,
    end_date: date | None = None,
    key_prefix: str = "visual_review",
) -> None:
    if edited.empty:
        return

    detection_map = {detection.note_name: detection for detection in detections}
    note_map = note_map or {}
    with st.expander("Revisão visual nota por nota", expanded=True):
        for position, (index, row) in enumerate(edited.iterrows()):
            note_name = str(row.get("nota", "") or "").strip()
            detection = detection_map.get(note_name)
            note_path = detection.note_path if detection else note_map.get(note_name)
            title_date = row.get("data")
            try:
                title_date_text = to_date(title_date).strftime("%d/%m/%Y") if to_date(title_date) else "sem data"
            except Exception:
                title_date_text = "sem data"
            try:
                value_text = brl(parse_money(row.get("valor")))
            except Exception:
                value_text = "valor inválido"
            title = f"{title_date_text} - {row.get('categoria', '')} - {value_text}"
            if not bool(row.get("usar", True)):
                title = f"Ignorada - {title}"
            st.markdown(f"**{title}**")

            col_image, col_data = st.columns([0.9, 1.1])
            with col_image:
                if note_path and note_path.exists():
                    display_note_image(note_path, width=360)
                else:
                    st.info("Sem imagem vinculada.")
            with col_data:
                st.write(f"Nota: {note_name or 'sem nota'}")
                st.write(f"Descrição: {row.get('descricao', '')}")
                st.write(f"Leitura: {row.get('confianca', '')} - {row.get('origem', '')}")
                alerts = str(row.get("alertas", "") or "").strip()
                if alerts:
                    st.warning(alerts)
                else:
                    st.success("Sem alertas.")
                if detection:
                    with st.expander("Texto/JSON lido", expanded=False):
                        st.text(detection.ocr_text or "Sem texto lido.")
                form_hash = sha1(f"{position}:{note_name}:{key_prefix}".encode("utf-8")).hexdigest()[:10]
                with st.expander("Corrigir esta nota", expanded=False):
                    try:
                        current_date = to_date(row.get("data")) or date.today()
                    except Exception:
                        current_date = date.today()
                    current_category = normalize_ai_category(row.get("categoria"))
                    current_category_index = CATEGORY_ORDER.index(current_category)
                    try:
                        current_value = float(parse_money(row.get("valor")))
                    except Exception:
                        current_value = 0.0
                    current_description = repair_text_encoding(row.get("descricao", "")).strip()
                    current_use = bool(row.get("usar", True))

                    with st.form(f"quick_fix_form_{form_hash}"):
                        fix_use = st.checkbox("Usar no relatório", value=current_use, key=f"quick_fix_use_{form_hash}")
                        fix_date = st.date_input("Data", value=current_date, key=f"quick_fix_date_{form_hash}")
                        fix_category = st.selectbox(
                            "Categoria",
                            options=CATEGORY_ORDER,
                            index=current_category_index,
                            key=f"quick_fix_category_{form_hash}",
                        )
                        fix_description = st.text_input(
                            "Descrição",
                            value=current_description,
                            key=f"quick_fix_description_{form_hash}",
                        )
                        fix_value = st.number_input(
                            "Valor correto",
                            min_value=0.0,
                            value=current_value,
                            step=0.01,
                            format="%.2f",
                            key=f"quick_fix_value_{form_hash}",
                        )
                        submitted = st.form_submit_button("Aplicar correção nesta nota", width="stretch")

                    if submitted:
                        review_df = st.session_state.get("review_df")
                        if not isinstance(review_df, pd.DataFrame) or review_df.empty:
                            review_df = edited.copy()
                        else:
                            review_df = review_df.copy()
                        if index in review_df.index:
                            target_index = index
                        elif position < len(review_df.index):
                            target_index = review_df.index[position]
                        else:
                            target_index = index

                        for column in ["usar", "data", "categoria", "descricao", "valor", "nota", "confianca", "origem", "alertas"]:
                            if column not in review_df.columns:
                                review_df[column] = ""
                        review_df.at[target_index, "usar"] = fix_use
                        review_df.at[target_index, "data"] = fix_date
                        review_df.at[target_index, "categoria"] = fix_category
                        review_df.at[target_index, "descricao"] = repair_text_encoding(fix_description).strip()
                        review_df.at[target_index, "valor"] = float(parse_money(fix_value))
                        review_df.at[target_index, "nota"] = note_name
                        review_df.at[target_index, "confianca"] = "Manual"
                        review_df.at[target_index, "origem"] = "Revisão manual"
                        review_df.at[target_index, "alertas"] = ""
                        st.session_state.review_df = add_review_alerts(review_df, note_map, start_date, end_date)
                        for state_key in list(st.session_state.keys()):
                            if str(state_key).startswith("review_editor_"):
                                del st.session_state[state_key]
                        st.success("Correção aplicada.")
                        st.rerun()


def render_auto_review(
    detections: list[NoteDetection],
    note_names: list[str],
    note_map: dict[str, Path],
    fallback_date: date,
    signature: str,
    start_date: date | None,
    end_date: date | None,
) -> pd.DataFrame:
    st.subheader("Conferência automática")

    if not detections:
        saved_review = st.session_state.get("review_df")
        if (
            st.session_state.get("review_signature") == signature
            and isinstance(saved_review, pd.DataFrame)
            and not saved_review.empty
        ):
            st.info("Leitura anterior preservada. Confira os dados e gere os arquivos.")
            edited = st.data_editor(
                saved_review,
                num_rows="dynamic",
                width="stretch",
                hide_index=True,
                column_config={
                    "usar": st.column_config.CheckboxColumn("Usar", default=True),
                    "data": st.column_config.DateColumn("Data", format="DD/MM/YYYY", required=True),
                    "categoria": st.column_config.SelectboxColumn("Categoria", options=CATEGORY_ORDER, required=True),
                    "descricao": st.column_config.TextColumn("Descrição"),
                    "valor": st.column_config.NumberColumn("Valor", min_value=0.0, step=0.01, format="R$ %.2f", required=True),
                    "nota": st.column_config.SelectboxColumn("Nota", options=[""] + note_names),
                    "confianca": st.column_config.TextColumn("Leitura"),
                    "origem": st.column_config.TextColumn("Origem"),
                    "alertas": st.column_config.TextColumn("Alertas"),
                },
                disabled=["confianca", "origem", "alertas"],
                key=f"review_editor_saved_{signature}",
            )
            edited = add_review_alerts(edited, note_map, start_date, end_date)
            st.session_state.review_df = edited
            render_visual_review(edited, [], note_map, start_date, end_date, f"saved_{signature}")
            return edited
        st.info("Depois de enviar as fotos, a leitura aparece aqui.")
        return pd.DataFrame(columns=["usar", "data", "categoria", "descricao", "valor", "nota", "confianca", "origem"])

    if st.session_state.get("review_signature") != signature:
        st.session_state.review_signature = signature
        st.session_state.review_df = add_review_alerts(
            detections_to_frame(detections, fallback_date), note_map, start_date, end_date
        )

    warning_count = sum(1 for detection in detections if detection.warnings)
    if warning_count:
        st.warning(f"{warning_count} nota(s) precisam de uma olhada mais cuidadosa.")
    else:
        st.success("As notas foram lidas. Confira o resumo e gere os arquivos.")

    edited = st.data_editor(
        st.session_state.review_df,
        num_rows="dynamic",
        width="stretch",
        hide_index=True,
        column_config={
            "usar": st.column_config.CheckboxColumn("Usar", default=True),
            "data": st.column_config.DateColumn("Data", format="DD/MM/YYYY", required=True),
            "categoria": st.column_config.SelectboxColumn("Categoria", options=CATEGORY_ORDER, required=True),
            "descricao": st.column_config.TextColumn("Descrição"),
            "valor": st.column_config.NumberColumn("Valor", min_value=0.0, step=0.01, format="R$ %.2f", required=True),
            "nota": st.column_config.SelectboxColumn("Nota", options=[""] + note_names),
            "confianca": st.column_config.TextColumn("Leitura"),
            "origem": st.column_config.TextColumn("Origem"),
            "alertas": st.column_config.TextColumn("Alertas"),
        },
        disabled=["confianca", "origem", "alertas"],
        key=f"review_editor_{signature}",
    )
    edited = add_review_alerts(edited, note_map, start_date, end_date)
    st.session_state.review_df = edited

    st.caption("Ao gerar os arquivos, as correcoes feitas aqui entram na memoria local para as proximas leituras.")
    render_visual_review(edited, detections, note_map, start_date, end_date, f"review_{signature}")

    with st.expander("Ver detalhes da leitura", expanded=False):
        for detection in detections:
            status = " | ".join(detection.warnings) if detection.warnings else "ok"
            st.markdown(f"**{detection.note_name}** - {detection.source} - {detection.confidence} - {status}")
            st.text(detection.ocr_text or "Sem texto lido.")

    flagged = [detection for detection in detections if detection.warnings]
    if flagged:
        with st.expander("Ver notas que precisam de conferência", expanded=False):
            for detection in flagged:
                st.markdown(f"**{detection.note_name}**")
                st.caption(" | ".join(detection.warnings))
                display_note_image(detection.note_path, width=320)

    return edited


def render_offline_review(
    detections: list[NoteDetection],
    note_map: dict[str, Path],
    signature: str,
) -> pd.DataFrame:
    render_flow_step(
        3,
        "Conferir Despesas",
        "Revise os dados abaixo. Qualquer ajuste feito nesta tabela será aplicado aos documentos finais.",
    )
    st.markdown(
        '<p class="section-caption">Desmarque “Incluir” somente quando uma despesa não deve constar no relatório.</p>',
        unsafe_allow_html=True,
    )
    columns = ["usar", "data", "categoria", "descricao", "valor", "nota"]
    if st.session_state.get("offline_review_signature") != signature:
        st.session_state.offline_review_signature = signature
        initial = detections_to_frame(detections, date.today())
        st.session_state.offline_review_df = initial[columns].copy()

    saved = st.session_state.get("offline_review_df")
    if not isinstance(saved, pd.DataFrame):
        saved = pd.DataFrame(columns=columns)

    edited = st.data_editor(
        saved,
        width="stretch",
        hide_index=True,
        column_order=["usar", "data", "categoria", "descricao", "valor"],
        column_config={
            "usar": st.column_config.CheckboxColumn("Incluir", default=True),
            "data": st.column_config.DateColumn("Data", format="DD/MM/YYYY", required=True),
            "categoria": st.column_config.SelectboxColumn("Categoria", options=CATEGORY_ORDER, required=True),
            "descricao": st.column_config.TextColumn("Descrição"),
            "valor": st.column_config.NumberColumn("Valor", min_value=0.0, step=0.01, format="R$ %.2f", required=True),
            "nota": None,
        },
        key=f"offline_review_editor_{signature}",
    )
    st.session_state.offline_review_df = edited
    return edited


def render_mileage_editor() -> tuple[pd.DataFrame, Decimal, Decimal]:
    with st.expander("Informações Complementares", expanded=False, icon=":material/add_road:"):
        st.caption("Preencha somente quando houver quilometragem reembolsável ou adiantamento recebido.")
        col_rate, col_advance = st.columns([1, 1])
        with col_rate:
            km_rate = parse_money(st.number_input("Valor Pago por Km", min_value=0.0, value=0.8, step=0.01))
        with col_advance:
            advance = parse_money(st.number_input("Adiantamento Recebido", min_value=0.0, value=0.0, step=0.01))

        mileage_df = st.data_editor(
            default_mileage(),
            num_rows="dynamic",
            width="stretch",
            hide_index=True,
            column_config={
                "data": st.column_config.DateColumn("Data", format="DD/MM/YYYY"),
                "distancia_km": st.column_config.NumberColumn("Distância Percorrida", min_value=0.0, step=1.0),
                "descricao": st.column_config.TextColumn("Descrição"),
            },
            key="mileage_editor",
        )

    return mileage_df, km_rate, advance


def validate_inputs(trip: dict[str, object], expenses: pd.DataFrame) -> None:
    if not TEMPLATE_PATH.exists():
        raise FileNotFoundError(f"Modelo não encontrado: {TEMPLATE_PATH}")
    if to_date(trip["end_date"]) < to_date(trip["start_date"]):
        raise ValueError("A data de chegada não pode ser anterior à data de saída.")
    if expenses.empty:
        raise ValueError("Informe pelo menos uma despesa com data, categoria e valor.")


def render_summary(expenses: pd.DataFrame) -> None:
    total = sum(expenses["valor_decimal"], Decimal("0")) if not expenses.empty else Decimal("0")
    expense_count = len(expenses)
    note_count = expenses["nota"].replace("", pd.NA).dropna().nunique() if not expenses.empty else 0
    period = (
        f"{expenses['data'].min().strftime('%d/%m/%Y')} a {expenses['data'].max().strftime('%d/%m/%Y')}"
        if not expenses.empty
        else "Não informado"
    )

    col_a, col_b, col_c, col_d = st.columns([1.1, .75, 1.65, .8])
    col_a.metric("Valor a Reembolsar", brl(total), border=True)
    col_b.metric("Despesas Incluídas", expense_count, border=True)
    col_c.metric("Período", period, border=True)
    col_d.metric("Comprovantes", note_count, border=True)


def trip_with_inferred_dates(trip: dict[str, object], expenses: pd.DataFrame) -> dict[str, object]:
    if not bool(trip.get("auto_dates")) or expenses.empty:
        return trip

    normalized = normalize_expenses(expenses)
    if normalized.empty:
        return trip

    adjusted = dict(trip)
    adjusted["start_date"] = normalized["data"].min()
    adjusted["end_date"] = normalized["data"].max()
    return adjusted


def report_limit_alerts(expenses: pd.DataFrame) -> list[str]:
    alerts: list[str] = []
    if expenses.empty:
        alerts.append("Nenhuma despesa selecionada")
        return alerts
    unique_dates = expenses["data"].nunique()
    other_count = int((expenses["categoria"] == "Outras Despesas").sum())
    if unique_dates > MAX_DATES:
        alerts.append(f"Mais de {MAX_DATES} datas diferentes")
    if other_count > MAX_OTHER_EXPENSES:
        alerts.append(f"Mais de {MAX_OTHER_EXPENSES} linhas em Outras Despesas")
    return alerts


def render_readiness_panel(reviewed_expenses: pd.DataFrame, normalized_expenses: pd.DataFrame) -> tuple[bool, bool]:
    blockers, warnings = review_alerts(reviewed_expenses)
    limit_alerts = report_limit_alerts(normalized_expenses)
    blockers.extend(limit_alerts)

    if blockers:
        st.error(f"{len(blockers)} pendência(s) impedem a geração dos documentos.", icon=":material/error:")
        with st.expander("Ver Pendências", expanded=True, icon=":material/checklist:"):
            for item in blockers:
                st.write(f"- {item}")
    elif warnings:
        st.warning(f"{len(warnings)} ponto(s) de atenção. Revise antes de continuar.", icon=":material/warning:")
        with st.expander("Ver Pontos de Atenção", expanded=False, icon=":material/checklist:"):
            for item in warnings:
                st.write(f"- {item}")
    elif not normalized_expenses.empty:
        st.success("Conferência concluída. Os documentos estão prontos para geração.", icon=":material/check_circle:")

    allow_with_alerts = False
    if blockers or warnings:
        allow_with_alerts = st.checkbox("Autorizar Geração Mesmo com Pendências", value=False)
    can_generate = not normalized_expenses.empty and (not blockers or allow_with_alerts)
    return can_generate, False


def generate_report(
    employee: dict[str, str],
    trip: dict[str, object],
    base_name: str,
    raw_expenses: pd.DataFrame,
    raw_mileage: pd.DataFrame,
    km_rate: Decimal,
    advance: Decimal,
    note_map: dict[str, Path],
) -> dict[str, object]:
    expenses = normalize_expenses(raw_expenses)
    validate_inputs(trip, expenses)
    mileage = normalize_mileage(raw_mileage)

    output_dir = OUTPUT_ROOT / f"{slugify(base_name, 'RELATORIO')}_{uuid4().hex}"
    output_dir.mkdir(parents=True, exist_ok=True)

    payload = build_payload(output_dir, base_name, employee, trip, expenses, mileage, km_rate, advance)
    notes_pdf_path = output_dir / f"{base_name}_NOTAS.pdf"

    try:
        missing_notes = generate_notes_pdf(expenses, note_map, notes_pdf_path)
        if missing_notes:
            raise ValueError("Não foi possível validar todos os comprovantes. Corrija as imagens antes de gerar.")
        run_excel_generator(payload, output_dir)
        xlsx_path = Path(payload["xlsx_path"])
        if not zipfile.is_zipfile(xlsx_path):
            raise ValueError("O Excel gerado está inválido ou corrompido.")
        if not notes_pdf_path.read_bytes().startswith(b"%PDF-"):
            raise ValueError("O PDF dos comprovantes está inválido ou corrompido.")
    except Exception:
        shutil.rmtree(output_dir, ignore_errors=True)
        raise

    return {
        "output_dir": output_dir,
        "xlsx": xlsx_path,
        "notes_pdf": notes_pdf_path,
        "missing_notes": missing_notes,
        "expenses": expenses,
    }


def main() -> None:
    configure_page()
    require_access()
    cleanup_stale_directories(OUTPUT_ROOT)
    cleanup_stale_directories(UPLOAD_WORK_DIR)
    render_header()

    work_dir = UPLOAD_WORK_DIR
    work_dir.mkdir(exist_ok=True)

    note_map, mobile_detections = render_notes_inputs(work_dir)
    if not mobile_detections:
        return

    expense_dates = [detection.expense_date for detection in mobile_detections if detection.expense_date]
    employee, trip, base_name = render_employee_form(expense_dates)
    signature = sha1((note_signature(note_map) + "|offline-zip-v1").encode("utf-8")).hexdigest()
    raw_expenses = render_offline_review(mobile_detections, note_map, signature)
    raw_mileage, km_rate, advance = render_mileage_editor()

    selected_expenses = raw_expenses.copy()
    if not selected_expenses.empty and "usar" in selected_expenses.columns:
        selected_expenses = selected_expenses[selected_expenses["usar"] == True].copy()
    selected_expenses = add_review_alerts(
        selected_expenses,
        note_map,
        None,
        None,
    )

    preview_expenses = pd.DataFrame()
    can_generate = False
    render_flow_step(4, "Finalizar Relatório", "Confira o resumo e gere os dois documentos para envio ao financeiro.")
    try:
        preview_expenses = normalize_expenses(selected_expenses)
        render_summary(preview_expenses)
        can_generate, _ = render_readiness_panel(selected_expenses, preview_expenses)
    except Exception as exc:
        st.warning(str(exc))

    generate = st.button(
        "Gerar Documentos Finais",
        type="primary",
        icon=":material/task_alt:",
        width="stretch",
        disabled=not can_generate,
    )
    if not generate:
        return

    try:
        with st.spinner("Gerando Excel e PDF dos comprovantes..."):
            report_trip = trip_with_inferred_dates(trip, selected_expenses)
            result = generate_report(
                employee=employee,
                trip=report_trip,
                base_name=base_name,
                raw_expenses=selected_expenses,
                raw_mileage=raw_mileage,
                km_rate=km_rate,
                advance=advance,
                note_map=note_map,
            )
    except Exception as exc:
        st.error(str(exc))
        return

    st.markdown(
        """
        <div class="final-panel">
            <strong>Pacote Final Pronto</strong>
            <span>O ZIP contém o Excel preenchido e o PDF dos comprovantes, prontos para envio ao financeiro.</span>
        </div>
        """,
        unsafe_allow_html=True,
    )

    final_zip = make_download_zip([Path(result["xlsx"]), Path(result["notes_pdf"])])
    final_expenses = result["expenses"]
    final_date = final_expenses["data"].max() if not final_expenses.empty else date.today()
    final_zip_name = f"{final_date.strftime('%Y%m%d')}_RELATÓRIO_DESPESAS_HENRIQUE_FINAL.zip"
    st.download_button(
        "Baixar Pacote Final",
        final_zip,
        file_name=final_zip_name,
        mime="application/zip",
        type="primary",
        icon=":material/folder_zip:",
        on_click="ignore",
        width="stretch",
    )

    with st.expander("Ver Arquivos Incluídos", expanded=False, icon=":material/inventory_2:"):
        st.write(f"- {Path(result['xlsx']).name}")
        st.write(f"- {Path(result['notes_pdf']).name}")

    if result["missing_notes"]:
        st.warning("Algumas despesas foram geradas sem imagem de nota vinculada.")


if __name__ == "__main__":
    main()

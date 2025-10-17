"""Herramientas para convertir remitos en PDF a un Excel consolidado.

Este módulo utiliza pdfplumber para leer remitos con un formato tabular
tradicional y genera un archivo ``salida.xlsx`` con las columnas
``numero_remito``, ``codigo`` y ``cantidad``.
"""

from __future__ import annotations

import argparse
import logging
import re
from pathlib import Path
from typing import Dict, Iterable, List, Optional

import pandas as pd
import pdfplumber


LOGGER = logging.getLogger(__name__)

CODIGO_PATTERN = re.compile(r"^(?P<codigo>[A-Z0-9]{5,})\b")
REMITO_PATTERN = re.compile(r"REMITO\s+([0-9\-]+)")
UNIDAD_PATTERN = re.compile(r"(\d+)\s*(?=UNIDAD(?:ES)?\b)", re.IGNORECASE)

IGNORED_TOKENS = {
    "ORIGINAL",
    "DUPLICADO",
    "TRIPLICADO",
    "CUADRUPLICADO",
}


def extract_numero_remito(text: str) -> Optional[str]:
    """Obtiene el número de remito presente en ``text``.

    Parameters
    ----------
    text:
        Bloque de texto a inspeccionar.

    Returns
    -------
    Optional[str]
        El número de remito si se detecta, ``None`` en caso contrario.
    """

    match = REMITO_PATTERN.search(text)
    if match:
        return match.group(1)
    return None


def _normalize_line(line: str) -> str:
    """Normaliza una línea eliminando espacios extra y tabs."""

    return re.sub(r"\s+", " ", line.strip())


def _finalize_item(
    numero_remito: str,
    codigo: Optional[str],
    lines: Iterable[str],
) -> Optional[Dict[str, object]]:
    """Genera un registro de item a partir de ``lines``.

    Parameters
    ----------
    numero_remito:
        Identificador del remito al que pertenece el ítem.
    codigo:
        Código detectado para el ítem.
    lines:
        Secuencia de líneas que conforman el bloque del ítem.
    """

    if not codigo:
        return None

    block_text = " ".join(filter(None, lines))
    if not block_text:
        return None

    cantidad_matches = list(UNIDAD_PATTERN.finditer(block_text))
    if not cantidad_matches:
        LOGGER.debug("No se encontró cantidad para %s en bloque '%s'", codigo, block_text)
        return None

    cantidad = int(cantidad_matches[-1].group(1))

    return {
        "numero_remito": numero_remito,
        "codigo": codigo,
        "cantidad": cantidad,
    }


def parse_items_from_text(text: str, numero_remito: str) -> List[Dict[str, object]]:
    """Extrae los ítems de un bloque de texto perteneciente a un remito."""

    items: List[Dict[str, object]] = []
    current_lines: List[str] = []
    current_code: Optional[str] = None

    for raw_line in text.splitlines():
        line = _normalize_line(raw_line)
        if not line:
            continue

        upper_line = line.upper()
        if any(token in upper_line for token in IGNORED_TOKENS) or upper_line.startswith("P\u00c1GINA"):
            continue

        # Evita procesar encabezados de tabla.
        if upper_line.startswith("CODIGO") or upper_line.startswith("C\u00d3DIGO"):
            continue

        code_match = CODIGO_PATTERN.match(line)
        if code_match:
            # Finaliza el bloque anterior si corresponde.
            item = _finalize_item(numero_remito, current_code, current_lines)
            if item:
                items.append(item)

            current_code = code_match.group("codigo")
            current_lines = [line]
        else:
            if current_lines:
                current_lines.append(line)

    # Procesa el último bloque pendiente.
    item = _finalize_item(numero_remito, current_code, current_lines)
    if item:
        items.append(item)

    return items


def parse_pdf(path_pdf: Path) -> List[Dict[str, object]]:
    """Procesa un PDF y devuelve los registros de ítems detectados."""

    if not path_pdf.exists():
        raise FileNotFoundError(f"No se encuentra el archivo: {path_pdf}")

    records: List[Dict[str, object]] = []
    numero_remito: Optional[str] = None

    with pdfplumber.open(str(path_pdf)) as pdf:
        for page in pdf.pages:
            text = page.extract_text(x_tolerance=1, y_tolerance=1) or ""
            page_remito = extract_numero_remito(text)
            if page_remito:
                numero_remito = page_remito

            if not numero_remito:
                continue

            records.extend(parse_items_from_text(text, numero_remito))

    if not numero_remito:
        raise ValueError(f"No se encontró número de remito en {path_pdf}")

    if not records:
        raise ValueError(f"No se encontraron ítems en {path_pdf}")

    return records


def _collect_pdfs_from_directory(directory: Path) -> List[Path]:
    """Obtiene los PDF dentro de ``directory`` de manera ordenada."""

    pdf_files = sorted(
        [path for path in directory.iterdir() if path.is_file() and path.suffix.lower() == ".pdf"]
    )
    return pdf_files


def _aggregate_records(records: Iterable[Dict[str, object]]) -> pd.DataFrame:
    """Agrupa los registros por número de remito y código sumando cantidades."""

    if not records:
        raise ValueError("No se recibieron registros para agrupar")

    df = pd.DataFrame(records)
    grouped = (
        df.groupby(["numero_remito", "codigo"], as_index=False)["cantidad"].sum()
        .sort_values(["numero_remito", "codigo"])
        .reset_index(drop=True)
    )
    grouped["cantidad"] = grouped["cantidad"].astype(int)
    return grouped


def main() -> None:
    """Punto de entrada principal del script CLI."""

    parser = argparse.ArgumentParser(description="Convierte remitos PDF a un Excel consolidado")
    parser.add_argument("--pdf", type=str, help="Ruta al archivo PDF a procesar")
    parser.add_argument(
        "--dir", type=str, help="Ruta a la carpeta con archivos PDF para procesar"
    )
    parser.add_argument(
        "--log-level",
        default="WARNING",
        choices=["DEBUG", "INFO", "WARNING", "ERROR", "CRITICAL"],
        help="Nivel de log a utilizar",
    )

    args = parser.parse_args()

    logging.basicConfig(level=getattr(logging, args.log_level.upper(), logging.WARNING))

    pdf_paths: List[Path] = []

    if args.pdf:
        pdf_paths = [Path(args.pdf)]
    elif args.dir:
        directory = Path(args.dir)
        if not directory.exists() or not directory.is_dir():
            parser.error(f"La ruta proporcionada no es una carpeta válida: {directory}")
        pdf_paths = _collect_pdfs_from_directory(directory)
        if not pdf_paths:
            parser.error(f"No se encontraron archivos PDF en {directory}")
    else:
        parser.error("Debe proporcionar --pdf o --dir")

    all_records: List[Dict[str, object]] = []
    errores: Dict[Path, str] = {}

    for pdf_path in pdf_paths:
        try:
            parsed_records = parse_pdf(pdf_path)
            LOGGER.info("Procesados %d ítems desde %s", len(parsed_records), pdf_path)
            all_records.extend(parsed_records)
        except Exception as exc:  # pylint: disable=broad-except
            errores[pdf_path] = str(exc)
            LOGGER.error("Error al procesar %s: %s", pdf_path, exc)

    if errores and not all_records:
        parser.error(
            "No se pudo procesar ningún archivo. Errores: "
            + "; ".join(f"{path}: {msg}" for path, msg in errores.items())
        )

    if not all_records:
        parser.error("No se obtuvieron ítems para exportar")

    consolidated_df = _aggregate_records(all_records)

    print("Preview de los primeros registros detectados:")
    print(consolidated_df.head())

    output_path = Path("salida.xlsx")
    consolidated_df.to_excel(output_path, index=False)
    print(f"Archivo Excel generado en: {output_path.resolve()}")


if __name__ == "__main__":
    main()

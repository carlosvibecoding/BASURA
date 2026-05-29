"""
Lógica de análisis ΔΔCt para qPCR (PPIA y SYP como genes control).
Replica el flujo de PLACA 1 RGS10.
"""
from __future__ import annotations

import re
import statistics
from dataclasses import dataclass, field
from typing import Dict, List, Optional, Tuple

HOUSEKEEPING = frozenset({"PPIA", "SYP"})
SD_THRESHOLD = 0.3
SAMPLE_RE = re.compile(r"^[CS]\d+$", re.IGNORECASE)


@dataclass
class WellReading:
    ct_values: List[float] = field(default_factory=list)
    ct_mean: Optional[float] = None
    ct_sd: Optional[float] = None

    def add_row(self, ct: object, ct_mean: object, ct_sd: object) -> None:
        if ct != "" and ct is not None:
            try:
                self.ct_values.append(float(ct))
            except (TypeError, ValueError):
                pass
        if self.ct_mean is None and ct_mean != "" and ct_mean is not None:
            try:
                self.ct_mean = float(ct_mean)
            except (TypeError, ValueError):
                pass
        if self.ct_sd is None and ct_sd != "" and ct_sd is not None:
            try:
                self.ct_sd = float(ct_sd)
            except (TypeError, ValueError):
                pass


@dataclass
class SampleCalcs:
    sample: str
    goi: str
    ct1: Optional[float]
    ct2: Optional[float]
    goi_mean: float
    goi_sd: float
    flag_high_sd: bool
    ppi_dct: float
    ppi_dct_mean_c: float
    ppi_ddct: float
    ppi_fc: float
    syp_dct: float
    syp_dct_mean_c: float
    syp_ddct: float
    syp_fc: float


def normalize_sample(name: str) -> str:
    return str(name).strip().upper()


def normalize_target(name: str) -> str:
    return str(name).strip().upper()


def is_valid_sample(name: str) -> bool:
    return bool(SAMPLE_RE.match(normalize_sample(name)))


def _norm_header_label(cell: object) -> str:
    """Normaliza cabeceras (StepOne usa a veces 'Cт' con т cirílica)."""
    label = str(cell).strip().lower()
    return label.replace("\u0442", "t")  # т cirílica → t latina


def find_raw_header(rows: List[List[object]]) -> Tuple[int, Dict[str, int]]:
    """Devuelve índice de fila de cabecera y mapa nombre_columna -> índice."""
    for idx, row in enumerate(rows):
        headers = {}
        for col, cell in enumerate(row):
            label = _norm_header_label(cell)
            if not label:
                continue
            if "sample name" in label or label == "sample":
                headers["sample"] = col
            elif "target name" in label or label == "target":
                headers["target"] = col
            elif label in ("ct", "ct."):
                headers["ct"] = col
            elif "ct mean" in label:
                headers["ct_mean"] = col
            elif "ct sd" in label:
                headers["ct_sd"] = col
        if "sample" in headers and "target" in headers and "ct" in headers:
            return idx, headers
    raise ValueError('No se encontró la fila de cabecera con "Sample Name" y "Ct".')


def parse_raw_rows(rows: List[List[object]]) -> Tuple[str, Dict[str, Dict[str, WellReading]], List[str]]:
    """
    Parsea datos RAW del termociclador.
    Retorna: gen de interés, datos[sample][target], orden de muestras (primera aparición en GOI).
    """
    header_row, cols = find_raw_header(rows)
    data: Dict[str, Dict[str, WellReading]] = {}
    sample_order: List[str] = []
    targets_seen: List[str] = []

    for row in rows[header_row + 1 :]:
        if not row or max(len(str(c)) for c in row) == 0:
            continue
        sample_col = cols.get("sample", 1)
        target_col = cols.get("target", 2)
        if sample_col >= len(row) or target_col >= len(row):
            continue
        sample = normalize_sample(row[sample_col])
        target = normalize_target(row[target_col])
        if not is_valid_sample(sample) or not target:
            continue

        ct = row[cols["ct"]] if "ct" in cols and cols["ct"] < len(row) else ""
        ct_mean = row[cols["ct_mean"]] if "ct_mean" in cols and cols["ct_mean"] < len(row) else ""
        ct_sd = row[cols["ct_sd"]] if "ct_sd" in cols and cols["ct_sd"] < len(row) else ""

        data.setdefault(sample, {})
        data[sample].setdefault(target, WellReading())
        data[sample][target].add_row(ct, ct_mean, ct_sd)

        if target not in targets_seen:
            targets_seen.append(target)

    goi_candidates = [t for t in targets_seen if t not in HOUSEKEEPING]
    if len(goi_candidates) != 1:
        raise ValueError(
            f"Se esperaba un único gen de interés (distinto de PPIA/SYP). Encontrados: {goi_candidates}"
        )
    goi = goi_candidates[0]

    for sample in data:
        if goi in data[sample]:
            sample_order.append(sample)

    for ref in ("PPIA", "SYP"):
        if ref not in targets_seen:
            raise ValueError(f"Falta el gen control {ref} en los datos RAW.")

    return goi, data, sample_order


def _mean_or_none(reading: WellReading) -> Optional[float]:
    if reading.ct_mean is not None:
        return reading.ct_mean
    if reading.ct_values:
        return statistics.mean(reading.ct_values)
    return None


def calculate_all(
    goi: str, data: Dict[str, Dict[str, WellReading]], sample_order: List[str]
) -> List[SampleCalcs]:
    """Calcula ΔCt, media C, ΔΔCt y 2^(-ΔΔCt) para PPIA y SYP."""
    # ΔCt por muestra y referencia
    dct_ppi: Dict[str, float] = {}
    dct_syp: Dict[str, float] = {}
    goi_info: Dict[str, Tuple[float, float, List[float]]] = {}

    for sample in sample_order:
        if goi not in data[sample]:
            continue
        goi_r = data[sample][goi]
        goi_mean = _mean_or_none(goi_r)
        if goi_mean is None:
            continue
        goi_sd = goi_r.ct_sd if goi_r.ct_sd is not None else 0.0
        cts = goi_r.ct_values[:2]
        while len(cts) < 2:
            cts.append(None)
        ct1, ct2 = (cts + [None, None])[:2]

        ppi_mean = _mean_or_none(data[sample].get("PPIA", WellReading()))
        syp_mean = _mean_or_none(data[sample].get("SYP", WellReading()))
        if ppi_mean is None or syp_mean is None:
            continue

        dct_ppi[sample] = goi_mean - ppi_mean
        dct_syp[sample] = goi_mean - syp_mean
        goi_info[sample] = (goi_mean, goi_sd, [x for x in (ct1, ct2) if x is not None])

        if ct1 is None and goi_r.ct_values:
            ct1 = goi_r.ct_values[0]
        if ct2 is None and len(goi_r.ct_values) > 1:
            ct2 = goi_r.ct_values[1]

    c_samples = [s for s in sample_order if s.startswith("C") and s in dct_ppi]
    if not c_samples:
        raise ValueError("No hay muestras control (C…) para calcular el promedio del paso 2.")

    avg_ppi = statistics.mean(dct_ppi[s] for s in c_samples)
    avg_syp = statistics.mean(dct_syp[s] for s in c_samples)

    results: List[SampleCalcs] = []
    for sample in sample_order:
        if sample not in dct_ppi:
            continue
        goi_mean, goi_sd, cts = goi_info[sample]
        ct1 = cts[0] if len(cts) > 0 else None
        ct2 = cts[1] if len(cts) > 1 else None
        ppi_ddct = dct_ppi[sample] - avg_ppi
        syp_ddct = dct_syp[sample] - avg_syp
        results.append(
            SampleCalcs(
                sample=sample,
                goi=goi,
                ct1=ct1,
                ct2=ct2,
                goi_mean=goi_mean,
                goi_sd=goi_sd,
                flag_high_sd=goi_sd > SD_THRESHOLD,
                ppi_dct=dct_ppi[sample],
                ppi_dct_mean_c=avg_ppi,
                ppi_ddct=ppi_ddct,
                ppi_fc=2 ** (-ppi_ddct),
                syp_dct=dct_syp[sample],
                syp_dct_mean_c=avg_syp,
                syp_ddct=syp_ddct,
                syp_fc=2 ** (-syp_ddct),
            )
        )
    return results


def goi_display_name(goi: str) -> str:
    """Ej.: RGS10r -> RGS10 para títulos."""
    name = goi.rstrip("rR").strip()
    return name if name else goi

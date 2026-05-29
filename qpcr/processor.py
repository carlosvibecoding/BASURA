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
FC_EXTREME = 1000.0
SAMPLE_RE = re.compile(r"^[CS]\d+$", re.IGNORECASE)
INDET_RE = re.compile(r"undetermin|indetermin", re.I)


def is_ct_indeterminate(ct: object) -> bool:
    if ct is None or ct == "":
        return True
    if isinstance(ct, (int, float)):
        return False
    return bool(INDET_RE.search(str(ct)))


def parse_ct_value(ct: object) -> Optional[float]:
    """Parsea Ct (coma/punto, locale ES) a float en rango plausible."""
    if is_ct_indeterminate(ct):
        return None
    if isinstance(ct, (int, float)):
        v = float(ct)
        return v if 5 <= v <= 50 else None
    s = str(ct).strip().replace("\u00a0", "").replace(" ", "")
    if not s:
        return None
    n_dot = s.count(".")
    n_com = s.count(",")
    if n_com == 1 and n_dot == 0:
        s = s.replace(",", ".")
    elif n_com == 1 and n_dot >= 1:
        s = s.replace(".", "").replace(",", ".")
    elif n_dot == 1 and n_com == 0:
        left, right = s.split(".", 1)
        if len(left) <= 2:
            s = f"{left}.{right}"
        else:
            s = s.replace(".", "")
    elif n_dot > 1:
        s = s.replace(".", "")
    try:
        v = float(s)
    except ValueError:
        return None
    return v if 5 <= v <= 50 else None


def sample_sort_key(sample: str) -> tuple:
    s = sample.upper()
    prefix = 0 if s.startswith("C") else 1
    num = int(s[1:]) if s[1:].isdigit() else 0
    return (prefix, num)


@dataclass
class WellReading:
    ct_values: List[float] = field(default_factory=list)
    ct_raw: List[str] = field(default_factory=list)
    ct_mean: Optional[float] = None
    ct_sd: Optional[float] = None

    def add_row(self, ct: object, ct_mean: object, ct_sd: object) -> None:
        self.ct_raw.append("" if ct is None else str(ct).strip())
        parsed = parse_ct_value(ct)
        if parsed is not None:
            self.ct_values.append(parsed)
        if self.ct_mean is None:
            m = parse_ct_value(ct_mean)
            if m is not None:
                self.ct_mean = m
        if self.ct_sd is None and ct_sd != "" and ct_sd is not None:
            try:
                sd = float(str(ct_sd).replace(",", "."))
                if 0 <= sd <= 5:
                    self.ct_sd = sd
            except (TypeError, ValueError):
                pass

    @property
    def is_indeterminate(self) -> bool:
        return len(self.ct_values) == 0

    @property
    def single_replicate(self) -> bool:
        return len(self.ct_values) == 1


@dataclass
class SampleCalcs:
    sample: str
    goi: str
    ct1: Optional[float]
    ct2: Optional[float]
    goi_mean: float
    goi_sd: float
    flag_high_sd: bool
    flag_indeterminate: bool
    flag_single_rep: bool
    ct1_display: Optional[str]
    ct2_display: Optional[str]
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
    sample_order.sort(key=sample_sort_key)

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
        ppi_r = data[sample].get("PPIA", WellReading())
        syp_r = data[sample].get("SYP", WellReading())
        if goi_r.is_indeterminate or ppi_r.is_indeterminate or syp_r.is_indeterminate:
            continue
        goi_mean = _mean_or_none(goi_r)
        ppi_mean = _mean_or_none(ppi_r)
        syp_mean = _mean_or_none(syp_r)
        if goi_mean is None or ppi_mean is None or syp_mean is None:
            continue
        goi_sd = goi_r.ct_sd if goi_r.ct_sd is not None else 0.0
        cts = goi_r.ct_values[:2]
        ct1, ct2 = (cts + [None, None])[:2]
        raw = goi_r.ct_raw[:2]
        while len(raw) < 2:
            raw.append("")
        ct1_d, ct2_d = raw[0], raw[1] if len(raw) > 1 else ""

        dct_ppi[sample] = goi_mean - ppi_mean
        dct_syp[sample] = goi_mean - syp_mean
        goi_info[sample] = (
            goi_mean,
            goi_sd,
            ct1,
            ct2,
            ct1_d,
            ct2_d,
            goi_r.single_replicate,
            goi_r.is_indeterminate,
        )

    c_samples = [s for s in sample_order if s.startswith("C") and s in dct_ppi]
    if not c_samples:
        raise ValueError("No hay muestras control (C…) para calcular el promedio del paso 2.")

    avg_ppi = statistics.mean(dct_ppi[s] for s in c_samples)
    avg_syp = statistics.mean(dct_syp[s] for s in c_samples)

    results: List[SampleCalcs] = []
    for sample in sample_order:
        goi_r = data[sample].get(goi, WellReading())
        ppi_r = data[sample].get("PPIA", WellReading())
        syp_r = data[sample].get("SYP", WellReading())
        raw = goi_r.ct_raw[:2]
        while len(raw) < 2:
            raw.append("")
        indet = goi_r.is_indeterminate or ppi_r.is_indeterminate or syp_r.is_indeterminate
        if sample in dct_ppi:
            goi_mean, goi_sd, ct1, ct2, ct1_d, ct2_d, one_rep, _ = goi_info[sample]
            ppi_ddct = dct_ppi[sample] - avg_ppi
            syp_ddct = dct_syp[sample] - avg_syp
            ppi_fc = 2 ** (-ppi_ddct)
            syp_fc = 2 ** (-syp_ddct)
        else:
            goi_mean = _mean_or_none(goi_r) or 0.0
            goi_sd = goi_r.ct_sd or 0.0
            ct1 = goi_r.ct_values[0] if goi_r.ct_values else None
            ct2 = goi_r.ct_values[1] if len(goi_r.ct_values) > 1 else None
            ct1_d, ct2_d = raw[0], raw[1]
            one_rep = goi_r.single_replicate
            ppi_fc = syp_fc = 0.0
            ppi_ddct = syp_ddct = 0.0
        results.append(
            SampleCalcs(
                sample=sample,
                goi=goi,
                ct1=ct1 if not indet else None,
                ct2=ct2 if not indet else None,
                goi_mean=goi_mean,
                goi_sd=goi_sd,
                flag_high_sd=goi_sd > SD_THRESHOLD,
                flag_indeterminate=indet,
                flag_single_rep=one_rep if not indet else False,
                ct1_display=ct1_d or None,
                ct2_display=ct2_d or None,
                ppi_dct=dct_ppi.get(sample, 0.0),
                ppi_dct_mean_c=avg_ppi,
                ppi_ddct=ppi_ddct if sample in dct_ppi else 0.0,
                ppi_fc=ppi_fc,
                syp_dct=dct_syp.get(sample, 0.0),
                syp_dct_mean_c=avg_syp,
                syp_ddct=syp_ddct if sample in dct_ppi else 0.0,
                syp_fc=syp_fc,
            )
        )
    return results


def goi_display_name(goi: str) -> str:
    """Ej.: RGS10r -> RGS10 para títulos."""
    name = goi.rstrip("rR").strip()
    return name if name else goi

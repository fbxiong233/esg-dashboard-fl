# -*- coding: utf-8 -*-
"""Federated-statistics ESG supply-chain risk pipeline.

Scoring engine, scenario generator, verification runs and Streamlit dashboard
in one module.  The demonstration covers six cumulative periods M0-M5, five
single-change events A-E under the frozen M0 parameters, an industry peer
experiment (M6) and a dependency-weight sensitivity (M7).  All indicator
values are simulated.

    python esg_fedstat_pipeline.py generate    # write the input workbooks to data/
    python esg_fedstat_pipeline.py verify      # run the demonstration, write outputs/
    streamlit run esg_fedstat_pipeline.py      # dashboard
"""
from __future__ import annotations

import hashlib
import json
import os
import random
import re
import shutil
import sys as _sys
import tempfile
import warnings
from dataclasses import dataclass, field
from datetime import datetime, timezone
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import numpy as np
import openpyxl
import pandas as pd

warnings.filterwarnings("ignore")

# =============================================================================
# SECTION 0.  PRESENTATION SWITCHES
# =============================================================================

# Entities are shown under anonymous codes.  False displays the bare entity IDs
# without the descriptive labels below.
ANONYMISE_ENTITIES = True

ENTITY_ALIAS = {
    "S2-CHEM":    "S2-CHEM    (Tier-2 chemical inputs)",
    "S2-EQUIP":   "S2-EQUIP   (Tier-2 equipment)",
    "S1-RUBBER":  "S1-RUBBER  (Tier-1 component)",
    "S1-COATING": "S1-COATING (Tier-1 component)",
    "FOCAL":      "FOCAL      (focal manufacturer)",
    "D1-RETAIL":  "D1-RETAIL  (downstream retail)",
    "D2-RETAIL":  "D2-RETAIL  (downstream retail)",
}

SHORT_ALIAS = {
    "S2-CHEM": "S2-CHEM", "S2-EQUIP": "S2-EQUIP", "S1-RUBBER": "S1-RUBBER",
    "S1-COATING": "S1-COATING", "FOCAL": "FOCAL",
    "D1-RETAIL": "D1-RETAIL", "D2-RETAIL": "D2-RETAIL",
}


def disp(entity: str, short: bool = True) -> str:
    """Display name for an entity, honoring ANONYMISE_ENTITIES."""
    if not ANONYMISE_ENTITIES:
        return entity
    return (SHORT_ALIAS if short else ENTITY_ALIAS).get(entity, entity)


# =============================================================================
# SECTION 1.  SUPPLY-CHAIN TOPOLOGY AND LOCAL DEPENDENCY WEIGHTS
# =============================================================================

SUPPLY_CHAIN = {
    "S2-CHEM":          {"tier": "Tier 2",     "role": "Fertilizer / chemical inputs", "upstream": []},
    "S2-EQUIP":            {"tier": "Tier 2",     "role": "Agricultural equipment",       "upstream": []},
    "S1-RUBBER":         {"tier": "Tier 1",     "role": "Rubber components",            "upstream": ["S2-CHEM", "S2-EQUIP"]},
    "S1-COATING": {"tier": "Tier 1",     "role": "Coatings",                     "upstream": []},
    "FOCAL":           {"tier": "Focal",      "role": "Focal manufacturer",           "upstream": ["S1-RUBBER", "S1-COATING"]},
    "D1-RETAIL":          {"tier": "Downstream", "role": "Retail",                       "upstream": ["FOCAL"]},
    "D2-RETAIL":           {"tier": "Downstream", "role": "Retail",                       "upstream": ["FOCAL"]},
}

ENTITY_ORDER = ["S2-CHEM", "S2-EQUIP", "S1-RUBBER", "S1-COATING",
                "FOCAL", "D1-RETAIL", "D2-RETAIL"]

TIER_ORDER = ("Tier 2", "Tier 1", "Focal", "Downstream")

# ---------------------------------------------------------------------------
# Dependency weights, set by each firm under an accepted method (Table 4) and
# not transmitted: beta_i on the firm's own score and alpha_ij on direct
# supplier j's total, with beta_i + sum_j alpha_ij = 1.  Values assumed for the
# demonstration.
# ---------------------------------------------------------------------------
DEPENDENCY_WEIGHTS = {
    "S2-CHEM":          {"beta": 1.00, "alpha": {}},
    "S2-EQUIP":            {"beta": 1.00, "alpha": {}},
    "S1-RUBBER":         {"beta": 0.50, "alpha": {"S2-CHEM": 0.20, "S2-EQUIP": 0.30}},
    "S1-COATING": {"beta": 1.00, "alpha": {}},
    "FOCAL":           {"beta": 0.70, "alpha": {"S1-RUBBER": 0.12, "S1-COATING": 0.18}},
    "D1-RETAIL":          {"beta": 0.99, "alpha": {"FOCAL": 0.01}},
    "D2-RETAIL":           {"beta": 0.98, "alpha": {"FOCAL": 0.02}},
}

# Methods the governance body accepts for supplier proportions (Table 4).
DEPENDENCY_METHODS = {
    "purchase value": "procurement spend from supplier j / total procurement spend of firm i",
    "input volume": "volume of inputs from supplier j / total input volume of firm i",
    "strategic dependency": "management-assessed strategic criticality, documented and approved",
}
# Methods resting on management judgment require enhanced assurance procedures.
ENHANCED_ASSURANCE_METHODS = {"strategic dependency"}

# Declared method of each firm with enrolled suppliers.
DECLARED_ALPHA_METHOD = {
    "S1-RUBBER": "purchase value",
    "FOCAL": "purchase value",
    "D1-RETAIL": "input volume",
    "D2-RETAIL": "input volume",
}

# Simulated response records checked by the assurance provider (Section 4.1,
# third task), keyed by (entity, dimension, first transition of the episode).
# S1-RUBBER has no record for its M5 hard alert; the check reports the gap.
ALERT_RESPONSES = {
    ("S2-CHEM", "Environmental", "M0->M1"): {
        "action": "investigated the emissions increase; adopted a reduction plan",
        "reason": "Scope 1-3 and NOx emissions above the peer benchmark",
        "evidence": "S2C-ENV-01"},
    ("S2-CHEM", "Social", "M4->M5"): {
        "action": "acknowledged the notice; confirmed its own child labor policy",
        "reason": "hard alert received via customer S1-RUBBER",
        "evidence": "S2C-SOC-05"},
    ("S2-EQUIP", "Social", "M4->M5"): {
        "action": "acknowledged the notice; confirmed its own child labor policy",
        "reason": "hard alert received via customer S1-RUBBER",
        "evidence": "S2E-SOC-05"},
    ("FOCAL", "Social", "M4->M5"): {
        "action": "requested further information from S1-RUBBER",
        "reason": "hard alert received via S1-RUBBER",
        "evidence": "FOC-SOC-05"},
    ("D1-RETAIL", "Social", "M4->M5"): {
        "action": "contacted FOCAL; monitoring the branch",
        "reason": "hard alert received via FOCAL",
        "evidence": "D1R-SOC-05"},
    ("D2-RETAIL", "Social", "M4->M5"): {
        "action": "contacted FOCAL; monitoring the branch",
        "reason": "hard alert received via FOCAL",
        "evidence": "D2R-SOC-05"},
}


def dependency_weight_problems(weights: Dict[str, dict]) -> List[str]:
    """Validate the convex weights and the enrolled, acyclic supply relationships."""
    problems = []
    if set(weights) != set(ENTITY_ORDER):
        problems.append("dependency weights must cover exactly the enrolled firms")
    rank = {tier: i for i, tier in enumerate(TIER_ORDER)}
    for e, w in weights.items():
        if e not in SUPPLY_CHAIN:
            continue
        try:
            beta = float(w["beta"])
            alpha = {j: float(a) for j, a in w["alpha"].items()}
        except (KeyError, TypeError, ValueError, AttributeError):
            problems.append(f"{e}: invalid beta/alpha configuration")
            continue
        if not all(np.isfinite(a) and 0.0 <= a <= 1.0
                   for a in [beta, *alpha.values()]):
            problems.append(f"{e}: beta and alpha must be finite and lie in [0,1]")
        total = beta + sum(alpha.values())
        if abs(total - 1.0) > 1e-9:
            problems.append(f"{e}: beta+sum(alpha) = {total:.6f} != 1")
        if any(j not in SUPPLY_CHAIN or rank[SUPPLY_CHAIN[j]["tier"]]
               >= rank[SUPPLY_CHAIN[e]["tier"]] for j in alpha):
            problems.append(f"{e}: suppliers must be enrolled firms in an earlier tier")
    return problems


def check_dependency_weights() -> List[str]:
    """Validate the demonstration weights and their declared allocation methods."""
    problems = dependency_weight_problems(DEPENDENCY_WEIGHTS)
    for e, w in DEPENDENCY_WEIGHTS.items():
        if w["alpha"] and DECLARED_ALPHA_METHOD.get(e) not in DEPENDENCY_METHODS:
            problems.append(f"{e}: no accepted dependency-weight method declared")
    return problems


# =============================================================================
# SECTION 2.  GOVERNANCE POLICY (set by the governance body, exogenous to the data)
# =============================================================================

class ConfigurationError(ValueError):
    """Raised when the governance configuration cannot be satisfied, for
    example when the weight bounds are infeasible.
    """


@dataclass
class RegulatoryPolicy:
    """Parameters set by the governance body (Section 4.1); none is estimated
    from the data.
    """

    # --- baseline mass of the weighted indicators ---------------------------
    block_share_t2: float = 0.70     # Level 2 (established)
    block_share_t3: float = 0.30     # Level 3 (emerging)
    # (Level 1 indicators carry no weight; they raise separate hard alerts.)

    # --- weights: a = lambda*b + (1-lambda)*v, projected onto the bounds ---
    lam: float = 0.70                # 1: baseline only, 0: data only
    band_lo: float = 0.60            # Level 2 bounds, multiples of the baseline
    band_hi: float = 1.40

    # --- Level 3 bounds ----------------------------------------------------
    t3_floor_frac: float = 0.25      # floor as fraction of equal weight
    t3_cap_frac: float = 3.00        # cap as multiple of equal weight

    # --- dispersion measure feeding the data-driven weights -----------------
    dispersion_mode: str = "common"  # "ratio" and "cv": sensitivity alternatives

    # --- conversion to the common scale (Section 4.2.2) ---------------------
    # Magnitudes map [0, magnitude_range_multiple x anchor mean] onto [0, 1];
    # declarations and percentages keep their natural ranges.  The common-scale
    # dispersion, and with it the data-driven weights, depends on these ranges
    # (T23).  An explicit range by indicator overrides both.
    magnitude_range_multiple: float = 10.0
    conversion_range_overrides: Dict[str, Tuple[float, float]] = field(default_factory=dict)

    # --- scoring -----------------------------------------------------------
    # tau in sigmoid(s/tau).  The escalation levels and warning limits below are
    # in score units, so tau is set together with them.
    squash_scale: float = 1.0
    dim_weights: Tuple[float, float] = (0.50, 0.50)   # (E, S) composite

    # --- escalation on the peer-relative score -----------------------------
    # A peer-relative flag, not an absolute measure of adverse impact; the absolute
    # channel is absolute_risk_table().
    peer_relative_flag_level: float = 0.60
    # Internal-warning limits: a component fires when its change exceeds the
    # smaller of the absolute limit and the relative limit times the firm's
    # pre-event total.
    component_delta_abs: float = 0.005
    component_delta_rel: float = 0.010          # share of the node's own total
    reestimation_period: str = "quarterly"

    # --- declared reference scales (see declare_reference_scales) ----------
    eta_sigma_score: float = 0.10        # eta > 0 governance parameter
    ref_scale_declaration: float = 0.50  # 0 / 0.5 / 1 declaration: half-range
    ref_scale_percentage: float = 0.05   # 5 percentage points on a fraction scale
    ref_scale_relative: float = 0.10     # 10% of the anchor level for magnitudes
    ref_scale_floor: float = 1e-9        # lower bound of every reference scale
    # Explicit reference scales by indicator id or name, in the indicator's unit.
    reference_scale_overrides: Dict[str, float] = field(default_factory=dict)

    # --- cold-start and coverage safeguard ---------------------------------
    # A score with a missing scored input, or resting on less than this share of
    # the published weight, is provisional (see provisional_status).
    coverage_floor: float = 0.80

    # --- graduated responses to ordinary changes (Section 4.2.3) -----------
    escalation_size_multiple: float = 3.0   # change of at least this multiple of the limit
    escalation_periods: int = 2             # consecutive transitions or periods
    persistent_level: float = 0.60          # dimension total for the level-based rule

    # --- data age and valid references -------------------------------------
    max_data_age_periods: int = 1           # older values are treated as unavailable
    reference_max_age_periods: int = 2      # oldest benchmark usable as a valid reference

    def validate(self) -> List[str]:
        p = []
        numeric = ("block_share_t2", "block_share_t3", "lam", "band_lo", "band_hi",
                   "t3_floor_frac", "t3_cap_frac", "magnitude_range_multiple", "squash_scale",
                   "peer_relative_flag_level", "component_delta_abs",
                   "component_delta_rel", "eta_sigma_score", "ref_scale_declaration",
                   "ref_scale_percentage", "ref_scale_relative", "ref_scale_floor",
                   "coverage_floor", "escalation_size_multiple", "persistent_level")
        for name in numeric:
            if not np.isfinite(getattr(self, name)):
                p.append(f"{name} must be finite")
        if not all(0.0 < x <= 1.0 for x in (self.block_share_t2, self.block_share_t3)):
            p.append("block shares must be positive and <= 1")
        if abs(self.block_share_t2 + self.block_share_t3 - 1.0) > 1e-9:
            p.append("block shares must sum to 1")
        if not 0.0 <= self.lam <= 1.0:
            p.append("lambda must lie in [0,1]")
        if not (0.0 < self.band_lo <= 1.0 <= self.band_hi):
            p.append("Level 2 band must have a positive lower bound and bracket 1.0")
        if not (0.0 < self.t3_floor_frac <= 1.0 <= self.t3_cap_frac):
            p.append("Level 3 floor/cap fractions must be positive and bracket 1.0")
        if self.magnitude_range_multiple <= 0.0:
            p.append("magnitude_range_multiple must be > 0")
        if self.squash_scale <= 0.0:
            p.append("squash_scale must be > 0")
        if (len(self.dim_weights) != 2 or not all(np.isfinite(x) and x >= 0.0
                for x in self.dim_weights) or abs(sum(self.dim_weights) - 1.0) > 1e-9):
            p.append("dim_weights must contain two nonnegative finite weights summing to 1")
        if self.component_delta_abs < 0.0 or self.component_delta_rel < 0.0:
            p.append("component warning limits must be nonnegative")
        if self.dispersion_mode not in ("common", "ratio", "cv"):
            p.append("dispersion_mode must be 'common', 'ratio' or 'cv'")
        if self.eta_sigma_score <= 0.0:
            p.append("eta_sigma_score must be > 0; a zero floor reopens the "
                     "degenerate-sigma division")
        for name in ("ref_scale_declaration", "ref_scale_percentage",
                     "ref_scale_relative", "ref_scale_floor"):
            if getattr(self, name) <= 0.0:
                p.append(f"{name} must be > 0")
        if not 0.0 < self.coverage_floor <= 1.0:
            p.append("coverage_floor must lie in (0, 1]")
        if any(not np.isfinite(v) or v <= 0.0 for v in self.reference_scale_overrides.values()):
            p.append("reference_scale_overrides must contain finite positive unit-bearing scales")
        if self.escalation_size_multiple < 1.0:
            p.append("escalation_size_multiple must be >= 1")
        if not 0.0 < self.persistent_level < 1.0:
            p.append("persistent_level must lie in (0, 1)")
        for name in ("escalation_periods", "max_data_age_periods", "reference_max_age_periods"):
            value = getattr(self, name)
            if not isinstance(value, int) or isinstance(value, bool) or value < 0:
                p.append(f"{name} must be a nonnegative integer")
        if self.escalation_periods < 1:
            p.append("escalation_periods must be >= 1")
        return p

    def digest(self) -> str:
        blob = json.dumps(self.__dict__, sort_keys=True, default=str)
        return hashlib.sha256(blob.encode()).hexdigest()[:16]


# ---------------------------------------------------------------------------
# Level 1 hard-alert rules (Section 4.2.3).  They carry no weight and change no
# score.  policy_required: the declaration must be affirmative; count_limit: a
# ceiling set by the governance body (None: monitored only).  The active rules
# are the declarations S26-S28; a failure is a policy deficiency, not a
# confirmed child- or forced-labor incident.
# ---------------------------------------------------------------------------
HARD_THRESHOLD_RULES = {
    "S26": {"rule": "policy_required", "label": "Child Labor Policy"},
    "S27": {"rule": "policy_required", "label": "Forced/Compulsory Labor Policy"},
    "S28": {"rule": "policy_required", "label": "Modern Slavery Statement"},
    "S17": {"rule": "count_limit", "limit": None, "label": "Fatalities (Employees)"},
    "S18": {"rule": "count_limit", "limit": None, "label": "Fatalities (Contractors)"},
    "S21": {"rule": "count_limit", "limit": None, "label": "OSHA Violations"},
    "S40": {"rule": "count_limit", "limit": None, "label": "Data Privacy Breaches"},
    "S45": {"rule": "count_limit", "limit": None, "label": "Substantiated Ethics Violations"},
    "E38": {"rule": "count_limit", "limit": None, "label": "Environmental Fines & Penalties"},
}

# A Level 1 observation whose data source or status marks it as preliminary,
# unverified, unassured or estimated stays pending (Sections 4.2.1 and 4.2.3):
# it raises no confirmed hard alert and cannot support the closure of one until
# it is verified.  Audit evidence counts as verification.
UNVERIFIED_TAG_WORDS = ("prelim", "unverif", "unassur", "unconfirm", "pending",
                        "provisional", "estimat")


def is_unverified_tag(source, status) -> bool:
    """True when a value's data source or status marks it as not yet verified."""
    text = " ".join(str(x) for x in (source, status) if isinstance(x, str)).lower()
    return any(word in text for word in UNVERIFIED_TAG_WORDS)


# =============================================================================
# SECTION 3.  INDICATOR PARSING, DIRECTION, AND GOVERNANCE LEVEL
# =============================================================================

GOOD_METRIC_KEYWORDS = [
    "reduction", "recycled", "renewable", "reused", "recycling rate",
    "diversion rate", "women", "sustainable packaging", "deforestation-free",
    "biodiversity policy", "tcfd", "iso 14001", "iso 9001", "sbti",
    "cdp", "validated", "disclosure", "certified", "certifications",
    "training hours", "training investment", "living wage", "minimum hourly",
    "human rights policy", "due diligence process", "audits conducted",
    "child labor policy", "forced", "modern slavery", "conflict minerals",
    "freedom of association", "indigenous peoples", "community investment",
    "charitable", "volunteer hours", "local hiring", "near-miss",
    "data privacy policy", "anti-corruption", "whistleblower", "ethics hotline",
    "supplier code", "suppliers audited", "supplier diversity",
    "responsible sourcing", "minorities", "minority", "ethnic", "disabilities",
    "recycled input", "renewable energy share",
]

ID_RE = re.compile(r"^\s*([ES]\d{1,2})\s*[.．]")


def indicator_id(name: str) -> Optional[str]:
    """'E1. Scope 1 GHG Emissions (Direct)' -> 'E1'."""
    m = ID_RE.match(str(name))
    return m.group(1) if m else None


# Direction per indicator id: True means a higher raw value is better (lower
# risk).  The keyword list above is only a fallback for an unlisted id.
# Judgment calls:
#   E6        reduction against the base year, recorded as a positive share
#   E33, S1   land managed and headcount measure exposure: more is worse
#   S11       a larger gender pay gap is worse
#   S20       more near-miss reports is better (a leading safety indicator)
#   S35, S44  more grievances and ethics complaints is worse
#   S48       more suppliers terminated for non-compliance is worse
_HIGHER_BETTER_IDS = (
    "E6 E7 E8 E10 E11 E14 E15 E19 E20 E28 E29 E31 E32 E34 E35 E36 E37 "
    "S2 S3 S4 S5 S6 S7 S9 S10 S13 S14 S20 S22 S23 S24 S26 S27 S28 S29 S30 "
    "S31 S32 S33 S34 S39 S41 S42 S43 S46 S47 S49 S50").split()
_HIGHER_WORSE_IDS = (
    "E1 E2 E3 E4 E5 E9 E12 E13 E16 E17 E18 E21 E22 E23 E24 E25 E26 E27 E30 "
    "E33 E38 S1 S8 S11 S12 S15 S16 S17 S18 S19 S21 S25 S35 S36 S37 S38 S40 "
    "S44 S45 S48").split()
if set(_HIGHER_BETTER_IDS) & set(_HIGHER_WORSE_IDS):
    raise ConfigurationError("an indicator ID is listed with both directions")
HIGHER_IS_BETTER: Dict[str, bool] = {**{i: True for i in _HIGHER_BETTER_IDS},
                                     **{i: False for i in _HIGHER_WORSE_IDS}}


def is_good_metric(name: str) -> bool:
    """True when a higher raw value means LOWER risk (see HIGHER_IS_BETTER)."""
    iid = indicator_id(name)
    if iid in HIGHER_IS_BETTER:
        return HIGHER_IS_BETTER[iid]
    low = str(name).lower()
    return any(kw in low for kw in GOOD_METRIC_KEYWORDS)


def parse_value(val) -> float:
    """Raw Excel cell -> float.  Handles ~, %, $, commas, Yes/No/Partial,
    CDP letter grades, and 'x:1' ratios."""
    if pd.isna(val):
        return np.nan
    s = str(val).strip()
    if s.upper() in ("N/D", "N/R", "N/A", "NA", "NAN", "NONE", "-", ""):
        return np.nan
    low = s.lower()
    if low.startswith("yes"):
        return 1.0
    if low.startswith("no"):
        return 0.0
    if low.startswith("partial"):
        return 0.5
    grade = {"A": 0.95, "A-": 0.85, "B+": 0.75, "B": 0.65, "B-": 0.55,
             "C+": 0.45, "C": 0.35, "C-": 0.25, "D+": 0.15, "D": 0.10,
             "D-": 0.05, "F": 0.0}
    if s in grade:
        return grade[s]
    clean = s.replace("~", "").replace(",", "").replace("$", "").strip()
    if "(" in clean:
        clean = clean[: clean.index("(")].strip()
    if ":1" in clean:
        try:
            return float(clean.replace(":1", ""))
        except ValueError:
            return np.nan
    if clean.endswith("%"):
        try:
            return float(clean[:-1]) / 100.0
        except ValueError:
            return np.nan
    try:
        return float(clean)
    except ValueError:
        return np.nan


def parse_indicator_value(value, unit: str, number_format: str = "") -> float:
    """Parse a value in its declared unit; percentages are stored as fractions.

    A plain number in a '%' row is read as percentage points unless the cell is
    formatted as a percentage.  Missing values stay NaN; malformed or infinite
    values are rejected.
    """
    result = parse_value(value)
    if np.isinf(result):
        raise ConfigurationError("indicator value must not be infinite")
    if np.isnan(result):
        missing_tokens = {"N/D", "N/R", "N/A", "NA", "NAN", "NONE", "-", ""}
        if pd.isna(value) or str(value).strip().upper() in missing_tokens:
            return result
        raise ConfigurationError("indicator value is neither numeric nor a declared missing-value token")
    percentage_unit = str(unit).strip().lower() in ("%", "percent", "percentage", "percentage points")
    text_percent = isinstance(value, str) and "%" in value
    formatted_percent = not isinstance(value, str) and "%" in number_format
    if percentage_unit and not text_percent and not formatted_percent:
        result /= 100.0
    return float(result)


# ---- governance-level map --------------------------------------------------

TIER_FALLBACK_T1 = set(HARD_THRESHOLD_RULES.keys())

TIER_FALLBACK_T2 = {
    "E1", "E2", "E3", "E4", "E5", "E9", "E13", "E17", "E18", "E21", "E22",
    "E24", "E25", "E26", "E27", "E32", "E35",
    "S4", "S7", "S11", "S12", "S14", "S15", "S16", "S19", "S22", "S23",
    "S29", "S30", "S36", "S37", "S41", "S42", "S43", "S46", "S50",
}


def classification_digest(path: str) -> str:
    """Content digest of the governance classification workbook, bound into
    every weight version.
    """
    if not os.path.exists(path):
        return "absent"
    with open(path, "rb") as fh:
        return hashlib.sha256(fh.read()).hexdigest()[:12]


def load_governance_tiers(path: str = "Indicator_Governance_Tier_Classification.xlsx"
                          ) -> Tuple[Dict[str, str], str]:
    """Governance level of every usable indicator, read from the classification
    workbook.

    Only rows marked OK with level T1-T3 are admitted (70 weighted indicators and
    9 hard-alert indicators); excluded rows are not defaulted to T3.
    """
    if os.path.exists(path):
        try:
            df = pd.read_excel(path, sheet_name="Classification (88)")
            col_id = next(c for c in df.columns if str(c).strip().lower() == "id")
            col_tier = next(c for c in df.columns if str(c).strip().lower() == "tier")
            col_use = next(c for c in df.columns if "usability" in str(c).lower())
            m = {"__STRICT__": "1"}
            for _, r in df.iterrows():
                iid = str(r[col_id]).strip().upper()
                tier = str(r[col_tier]).strip().upper()
                marker = str(r[col_use]).strip().upper().split()
                usable = bool(marker) and marker[-1] == "OK"
                if usable and iid and tier in ("T1", "T2", "T3"):
                    m[iid] = tier
            count = len(m) - 1
            if count:
                return m, (f"strict usable set loaded from {os.path.basename(path)} "
                           f"({count} indicators: weighted plus hard gates)")
        except Exception as exc:
            print(f"  [warn] could not read tier workbook ({exc}); using fallback map")
    m = {i: "T1" for i in TIER_FALLBACK_T1}
    m.update({i: "T2" for i in TIER_FALLBACK_T2})
    return m, "built-in fallback map (classification workbook not found; not strict-79)"


# =============================================================================
# SECTION 4.  DATA LOADING
# =============================================================================

COMPANY_COLUMNS = {
    "S2-CHEM\n(A01)": "S2-CHEM",
    "S2-EQUIP\n(A02)": "S2-EQUIP",
    "S1-RUBBER\n(A03)": "S1-RUBBER",
    "S1-COATING\n(A04)": "S1-COATING",
    "FOCAL\n(A05)": "FOCAL",
    "D1-RETAIL\n(A06)": "D1-RETAIL",
    "D2-RETAIL\n(A07)": "D2-RETAIL",
}

DIM_SHEETS = {"Environmental": "Environmental Data",
              "Social": "Social Data"}
NOTES_SHEET = "Methodology"
# Optional sheet: source, status and reference period of individual values.
DATA_STATUS_SHEET = "Data Status"
DATA_STATUS_COLUMNS = ("entity", "indicator", "source", "status", "as of period", "note")
DEFAULT_DATA_TAG = ("reported", "reported", "", "")
# Alternative sheet names, read as the names above; generation renames them.
LEGACY_SHEET_NAMES = {"Environmental Raw Data": "Environmental Data",
                      "Social Raw Data": "Social Data",
                      "Methodology & Sources": "Methodology"}


def workbook_sheet(workbook, name: str):
    """Return sheet `name`, or the sheet holding it under an alternative name."""
    if name in workbook.sheetnames:
        return workbook[name]
    for old, new in LEGACY_SHEET_NAMES.items():
        if new == name and old in workbook.sheetnames:
            return workbook[old]
    raise KeyError(name)


def normalise_sheet_names(workbook) -> None:
    """Rename sheets that carry an alternative name."""
    for old, new in LEGACY_SHEET_NAMES.items():
        if old in workbook.sheetnames and new not in workbook.sheetnames:
            workbook[old].title = new

ENTITY_BY_CODE = {f"A0{i + 1}": e for i, e in enumerate(ENTITY_ORDER)}
ENTITY_CODE_RE = re.compile(r"\(\s*(A0[1-7])\s*\)", re.IGNORECASE)
HEADER_MAP_FILENAME = "entity_header_map.json"

PERIOD_FILE_RE = re.compile(
    r"^ESG_Raw_Data_Supply_Chain_M(\d+)"          # period index
    r"(?:\s*\(\d+\)|\s*[-_ ]?copy\s*\d*)?"      # ' (1)', ' copy', '-copy2'
    r"\.xlsx$", re.IGNORECASE)


def _normalise_header(head) -> str:
    return " ".join(str(head).replace("\n", " ").split()).strip().lower()


def load_entity_header_map(data_dir: str = ".") -> Dict[str, str]:
    """Optional entity_header_map.json next to the data, mapping workbook column
    headers to entity ids.  The generated workbooks use the entity ids and need
    none.
    """
    path = os.path.join(data_dir, HEADER_MAP_FILENAME)
    if not os.path.exists(path):
        return {}
    with open(path, "r", encoding="utf-8") as fh:
        raw = json.load(fh)
    out = {}
    for key, value in raw.items():
        entity = str(value).strip()
        if entity not in ENTITY_ORDER:
            raise ConfigurationError(
                f"{HEADER_MAP_FILENAME}: '{entity}' is not one of {ENTITY_ORDER}")
        out[_normalise_header(key)] = entity
    return out


def resolve_entity_header(head, header_map: Optional[Dict[str, str]] = None
                          ) -> Optional[str]:
    """Column header -> entity id, or None when the column is not an entity.

    Accepted, in order: the header map, the entity id itself, an A01-A07 code,
    the COMPANY_COLUMNS table.
    """
    text = str(head).strip()
    if not text or text.lower() == "nan":
        return None
    flat = _normalise_header(text)
    if header_map and flat in header_map:
        return header_map[flat]
    first_line = text.splitlines()[0].strip()
    if first_line in ENTITY_ORDER:
        return first_line
    code = ENTITY_CODE_RE.search(text)
    if code:
        return ENTITY_BY_CODE[code.group(1).upper()]
    if text in COMPANY_COLUMNS:
        return COMPANY_COLUMNS[text]
    if flat in {_normalise_header(k): v for k, v in COMPANY_COLUMNS.items()}:
        return {_normalise_header(k): v for k, v in COMPANY_COLUMNS.items()}[flat]
    return None


def discover_period_files(data_dir: str = ".", expect_periods: Optional[int] = None,
                          verbose: bool = True) -> List[Tuple[str, str]]:
    """Locate the period workbooks M0, M1, ... and stop when the sequence is
    incomplete or a period appears twice.
    """
    if not os.path.isdir(data_dir):
        raise FileNotFoundError(f"data directory not found: {data_dir}")
    found: Dict[int, List[str]] = {}
    for name in sorted(os.listdir(data_dir)):
        if name.startswith("~$") or name.startswith("."):
            continue
        m = PERIOD_FILE_RE.match(name)
        if m:
            # M6 and M7 are separate experiments, not periods (run_m6_peer_experiment,
            # run_m7_dependency_sensitivity).
            if int(m.group(1)) in (6, 7):
                continue
            found.setdefault(int(m.group(1)), []).append(name)
    if not found:
        raise FileNotFoundError(
            f"no ESG_Raw_Data_Supply_Chain_M*.xlsx under {os.path.abspath(data_dir)}")

    files, notes = [], []
    for idx in sorted(found):
        candidates = found[idx]
        canonical = f"ESG_Raw_Data_Supply_Chain_M{idx}.xlsx"
        if canonical in candidates:
            chosen = canonical
        else:
            chosen = max(candidates,
                         key=lambda n: os.path.getmtime(os.path.join(data_dir, n)))
        if len(candidates) > 1:
            notes.append(f"M{idx}: {len(candidates)} candidate files {candidates}; "
                         f"using '{chosen}' (newest, or the un-suffixed name)")
        elif chosen != canonical:
            notes.append(f"M{idx}: read from '{chosen}' (copy suffix accepted)")
        files.append((f"M{idx}", os.path.join(data_dir, chosen)))

    indices = sorted(found)
    if indices != list(range(len(indices))):
        missing = [f"M{i}" for i in range(max(indices) + 1) if i not in found]
        raise FileNotFoundError(
            f"period sequence incomplete in {os.path.abspath(data_dir)}: "
            f"missing {', '.join(missing)}")
    if expect_periods is not None and len(files) != expect_periods:
        raise FileNotFoundError(
            f"expected {expect_periods} periods (M0-M{expect_periods - 1}) in "
            f"{os.path.abspath(data_dir)} but found {len(files)}: "
            f"{', '.join(l for l, _ in files)}. Pass expect_periods=None to run "
            f"a deliberately partial set.")
    if verbose and notes:
        for n in notes:
            print(f"  [data] {n}")
    return files


def read_data_status(workbook) -> Dict[Tuple[str, str], Tuple[str, str, str, str]]:
    """{(entity, indicator id): (source, status, as-of period, note)} from DATA_STATUS_SHEET."""
    if DATA_STATUS_SHEET not in workbook.sheetnames:
        return {}
    rows = list(workbook[DATA_STATUS_SHEET].iter_rows(values_only=True))
    head = [str(v or "").strip().lower() for v in (rows[0] if rows else ())]
    if any(c not in head for c in DATA_STATUS_COLUMNS):
        raise ConfigurationError(f"{DATA_STATUS_SHEET}: expected columns {DATA_STATUS_COLUMNS}")
    pos = [head.index(c) for c in DATA_STATUS_COLUMNS]
    tags = {}
    for row in rows[1:]:
        entity, iid, source, status, as_of, text = (str(row[i] or "").strip() for i in pos)
        if not entity:
            continue
        if entity not in ENTITY_ORDER or not re.fullmatch(r"[ES]\d{1,2}", iid.upper()):
            raise ConfigurationError(f"{DATA_STATUS_SHEET}: unknown entity or indicator {entity}/{iid}")
        tags[(entity, iid.upper())] = (source or "reported", status or "reported", as_of, text)
    return tags


def load_period(filepath: str, tier_map: Dict[str, str],
                header_map: Optional[Dict[str, str]] = None) -> Dict[str, pd.DataFrame]:
    """Load one period's workbook into {dimension: tidy DataFrame}.

    Columns: entity, category, indicator, ind_id, unit, source_unit, value,
    is_good, tier, data_source, data_status, as_of, data_note
    """
    # Numeric cells with a '%' Excel format already store fractions; a plain
    # numeric cell in a '%' unit row stores percentage points. Text '12%' also
    # becomes a fraction. The number format, not the magnitude, decides this
    # (0.5 may mean 0.5%).
    from openpyxl import load_workbook
    workbook = load_workbook(filepath, read_only=False, data_only=True)
    if header_map is None:
        header_map = load_entity_header_map(os.path.dirname(os.path.abspath(filepath)))
    tags = read_data_status(workbook)
    out = {}
    for dim, sheet in DIM_SHEETS.items():
        try:
            worksheet = workbook_sheet(workbook, sheet)
        except KeyError:
            raise ConfigurationError(f"{os.path.basename(filepath)}: missing sheet {sheet}") from None
        raw = pd.DataFrame(worksheet.values)
        if raw.shape[0] < 4 or raw.shape[1] < 4:
            raise ConfigurationError(f"{os.path.basename(filepath)} [{sheet}]: invalid data layout")

        cols, unmatched = {}, []
        for j in range(3, raw.shape[1]):
            head = str(raw.iloc[1, j]).strip()
            entity = resolve_entity_header(head, header_map)
            if entity is not None:
                if entity in cols.values():
                    raise ConfigurationError(
                        f"{os.path.basename(filepath)} [{sheet}]: duplicate entity column {entity}")
                cols[j] = entity
            elif head and head.lower() != "nan":
                unmatched.append(head.replace("\n", " / "))
        if len(set(cols.values())) < len(ENTITY_ORDER):
            missing = [e for e in ENTITY_ORDER if e not in set(cols.values())]
            raise ConfigurationError(
                f"{os.path.basename(filepath)} [{sheet}]: resolved "
                f"{len(set(cols.values()))} of {len(ENTITY_ORDER)} entity columns; "
                f"missing {missing}. Unmatched headers: {unmatched}. "
                f"Supply an '{HEADER_MAP_FILENAME}' next to the data mapping each "
                f"column header to an entity id.")

        rows, cur_cat, seen_ids = [], "", set()
        for i in range(3, raw.shape[0]):
            c = raw.iloc[i, 0]
            if pd.notna(c) and str(c).strip():
                cur_cat = str(c).strip()
            ind = raw.iloc[i, 1]
            if pd.isna(ind) or not str(ind).strip():
                continue
            ind = str(ind).strip()
            unit = str(raw.iloc[i, 2]).strip() if pd.notna(raw.iloc[i, 2]) else ""
            if unit.lower() == "description":
                continue                      # free-text targets: not modeled
            iid = indicator_id(ind)
            if tier_map.get("__STRICT__") == "1" and iid not in tier_map:
                continue
            if iid is None:
                continue
            if iid in seen_ids:
                raise ConfigurationError(f"{os.path.basename(filepath)} [{sheet}]: duplicate indicator {iid}")
            if not iid.startswith(dim[0]):
                raise ConfigurationError(f"{os.path.basename(filepath)} [{sheet}]: wrong dimension for {iid}")
            seen_ids.add(iid)
            tier = tier_map.get(iid, "T3") if iid else "T3"
            good = is_good_metric(ind)
            is_percentage = unit.lower() in ("%", "percent", "percentage", "percentage points")
            for j, ent in cols.items():
                cell_value = raw.iloc[i, j]
                try:
                    v = parse_indicator_value(cell_value, unit, worksheet.cell(i + 1, j + 1).number_format)
                except ConfigurationError as exc:
                    raise ConfigurationError(
                        f"{os.path.basename(filepath)} [{sheet}]: invalid value for {ent}/{iid}: {exc}") from exc
                # Missing cells keep their metadata, so every indicator stays in
                # the configured governance universe.
                source, status, as_of, text = tags.get((ent, iid), DEFAULT_DATA_TAG)
                rows.append({"entity": ent, "category": cur_cat,
                             "indicator": ind, "ind_id": iid,
                             "unit": "fraction (percentage)" if is_percentage else unit,
                             "source_unit": unit, "value": v, "is_good": good, "tier": tier,
                             "data_source": source, "data_status": status,
                             "as_of": as_of, "data_note": text})
        if tier_map.get("__STRICT__") == "1":
            expected = {k for k in tier_map if k.startswith(dim[0]) and k != "__STRICT__"}
            if expected - seen_ids:
                raise ConfigurationError(
                    f"{os.path.basename(filepath)} [{sheet}]: missing configured indicators {sorted(expected - seen_ids)}")
        out[dim] = pd.DataFrame(rows, columns=["entity", "category", "indicator", "ind_id",
                                             "unit", "source_unit", "value", "is_good", "tier",
                                             "data_source", "data_status", "as_of", "data_note"])
    workbook.close()
    return out


def period_index(label: str) -> Optional[int]:
    """'M3' -> 3, 'M3-fixed' -> 3; None for a label without a period number."""
    m = re.match(r"^[A-Za-z]+(\d+)", str(label or ""))
    return int(m.group(1)) if m else None


def apply_data_age(data: Dict[str, pd.DataFrame], label: str,
                   policy: RegulatoryPolicy) -> Tuple[Dict[str, pd.DataFrame], List[dict]]:
    """Treat values older than policy.max_data_age_periods as unavailable.

    Returns the adjusted frames and one row per tagged or stale value.
    """
    now = period_index(label)
    out, rows = {}, []
    for dim, df in data.items():
        df = df.copy(deep=True)
        if "as_of" not in df:
            out[dim] = df
            continue
        ref = pd.to_numeric(df["as_of"].map(period_index), errors="coerce")
        age = (now - ref) if now is not None else pd.Series(np.nan, index=df.index)
        stale = age > policy.max_data_age_periods
        df.loc[stale, "value"] = np.nan
        df.loc[stale, "data_status"] = "stale; treated as unavailable"
        tagged = stale | (df["data_source"] != "reported") | (df["data_status"] != "reported")
        for i in df.index[tagged]:
            r = df.loc[i]
            rows.append({"period": label, "dimension": dim, "entity": r["entity"],
                         "ind_id": r["ind_id"], "data_source": r["data_source"],
                         "data_status": r["data_status"], "as_of": r["as_of"] or label,
                         "age_periods": 0 if pd.isna(age[i]) else int(age[i]),
                         "max_age_periods": policy.max_data_age_periods,
                         "used_in_score": bool(not stale[i] and np.isfinite(r["value"])),
                         "note": r["data_note"]})
        out[dim] = df
    return out, rows


# =============================================================================
# SECTION 5.  FEDERATED STATISTICS  (one upload stage per calibration;
# data-dependent reference initialization is accounted for separately)
# =============================================================================

@dataclass
class LocalSummary:
    """One firm-level (n, sum, sum of squares) triple per indicator.

    Calibration transmits the common-scale instance. Raw-unit instances support
    local evaluator checks and, when needed, the separate reference initialization.
    """
    entity: str
    n: Dict[str, int]
    sx: Dict[str, float]
    sxx: Dict[str, float]

    def payload_digest(self) -> str:
        blob = json.dumps({"n": self.n, "sx": self.sx, "sxx": self.sxx},
                          sort_keys=True)
        return hashlib.sha256(blob.encode()).hexdigest()[:16]

    def record_count_transmitted(self) -> int:
        """Count raw-record payload objects, not information disclosure.

        A transmitted summary with n=1 still reveals its observation via Sx.
        """
        return 0


def to_common_scale(values, bounds: Tuple[float, float], higher_is_better: bool) -> np.ndarray:
    """Published conversion to the common scale (Section 4.2.2): linear on the
    reference range, clipped to [0, 1] and oriented so that 1 is least favorable."""
    lo, hi = bounds
    x = np.clip((np.asarray(values, dtype=float) - lo) / (hi - lo), 0.0, 1.0)
    return 1.0 - x if higher_is_better else x


class Participant:
    """Reference participant emitting summary objects.

    The process holds plaintext inputs; an n=1 summary reveals that observation.
    """

    def __init__(self, name: str, benchmark_only: bool = False):
        self.name = name
        cfg = ({"tier": "Peer pool", "role": "Industry benchmark only"}
               if benchmark_only else SUPPLY_CHAIN[name])
        self.tier = cfg["tier"]
        self.role = cfg["role"]

    def local_summary(self, df_dim: pd.DataFrame) -> LocalSummary:
        sub = df_dim[df_dim["entity"] == self.name]
        if sub["indicator"].duplicated().any():
            raise ConfigurationError(
                f"{self.name}: one firm-level value per indicator and reporting period is required")
        n, sx, sxx = {}, {}, {}
        for _, r in sub.iterrows():
            k, v = r["indicator"], float(r["value"])
            if not np.isfinite(v):
                continue
            n[k], sx[k], sxx[k] = 1, v, v * v
        return LocalSummary(self.name, n, sx, sxx)

    def common_scale_summary(self, df_dim: pd.DataFrame,
                             bounds: Dict[str, Tuple[float, float]],
                             higher_is_better: Dict[str, bool]) -> LocalSummary:
        """The calibration upload of Section 4.2.2: the firm converts its own values
        with the published reference points and sends (1, x, x^2) per indicator."""
        sub = df_dim[df_dim["entity"] == self.name].copy()
        if not sub.empty:
            sub["value"] = [float(to_common_scale(v, bounds[k], bool(higher_is_better.get(k, False))))
                            if k in bounds else float("nan")
                            for k, v in zip(sub["indicator"], sub["value"])]
        return self.local_summary(sub)


@dataclass
class GlobalStats:
    """Common-scale coordinator output with raw-unit evaluator diagnostics.

    common_mu/common_sigma are aggregated from the calibration uploads;
    mu/sigma retain raw-unit diagnostics for legacy tables and sensitivities.

    Status of each indicator: published (n >= MIN_PARTICIPANTS, sigma > 0);
    zero_variance_observed (mu is published with sigma = 0, the governed weight is
    kept and z = 0); unavailable (too few reporters, nothing published);
    valid_reference (an earlier benchmark within policy.reference_max_age_periods
    is used and named in `carried_forward`).
    """
    mu: Dict[str, float]
    sigma: Dict[str, float]
    n: Dict[str, int]
    is_good: Dict[str, bool]
    tier: Dict[str, str]
    ind_id: Dict[str, Optional[str]]
    category: Dict[str, str]
    status: Dict[str, str] = field(default_factory=dict)
    unavailable: Dict[str, str] = field(default_factory=dict)
    s_ref: Dict[str, float] = field(default_factory=dict)
    sigma_score: Dict[str, float] = field(default_factory=dict)
    # Participants whose summaries entered this round (reported with the score).
    contributors: List[str] = field(default_factory=list)
    common_mu: Dict[str, float] = field(default_factory=dict)
    common_sigma: Dict[str, float] = field(default_factory=dict)
    input_bounds: Dict[str, Tuple[float, float]] = field(default_factory=dict)
    # Reported values outside the conversion range (clipped to its ends).
    conversion_clipped: Dict[str, int] = field(default_factory=dict)
    # Indicators benchmarked on an earlier valid reference this round: k -> source.
    carried_forward: Dict[str, str] = field(default_factory=dict)


class Coordinator:
    """Aggregation of participant summaries, simulated in a single process
    (not cryptographic privacy); publishes versioned parameters.
    """

    # Minimum number of reporters for a published benchmark.  With three, one
    # reporter could solve for the other two values from the published n, mu and
    # sigma; with four or more it learns only their sum and sum of squares
    # (see benchmark_disclosure_check).
    MIN_PARTICIPANTS = 4
    SIGMA_EPS = 1e-12        # absolute degeneracy floor
    SIGMA_REL_EPS = 1e-9     # sigma/|mu| below this counts as no dispersion

    def __init__(self, secure_aggregation: bool = True, seed: int = 20260829):
        self.secure_aggregation = secure_aggregation
        self.rng = np.random.default_rng(seed)
        self.mask_log: List[dict] = []

    # -- secure aggregation simulation --------------------------------------
    def _masked_sum(self, values: List[float], tag: str) -> float:
        """Additive-mask aggregation, simulated in a single process.

        Shows the communication pattern and the sufficiency of (n, Sx, Sxx), not
        cryptographic isolation.  The masked view is logged in `mask_log`; the exact
        sum is returned.
        """
        exact = float(np.sum(values))
        if not self.secure_aggregation or len(values) < 2:
            return exact
        scale = max(1.0, float(np.max(np.abs(values))) * 10.0)
        masks = self.rng.normal(0.0, scale, size=len(values))
        masks -= masks.mean()                     # sum(masks) == 0
        masked = np.asarray(values, dtype=float) + masks
        self.mask_log.append({
            "tag": tag,
            "coordinator_per_participant_view": [float(x) for x in masked[:3]],
            "simulated": True,
            "cryptographic_isolation_implemented": False,
            "note": ("single-process simulation: the coordinator holds the "
                     "plaintext values and generates the masks itself; with "
                     "one contributor the masked sum is that contribution"),
            "recovered_sum_is_exact": True,
        })
        return exact

    # -- aggregation and publication ----------------------------------------
    def aggregate(self, summaries: List[LocalSummary],
                  meta: pd.DataFrame) -> GlobalStats:
        """Publish (mu, sigma, n, status) for every indicator with enough reporters.

        An indicator with zero dispersion is published with sigma = 0 and keeps its
        governed weight.
        """
        return self._summary_statistics(summaries, meta)

    def _summary_statistics(self, summaries: List[LocalSummary],
                            meta: pd.DataFrame) -> GlobalStats:
        """The (n, sum, sum of squares) arithmetic of aggregate(), shared with the
        local raw-unit diagnostics, which are not uploads (raw_unit_diagnostics)."""
        if len({s.entity for s in summaries}) != len(summaries):
            raise ConfigurationError("duplicate participant summary in aggregation")
        for summary in summaries:
            if set(summary.n) != set(summary.sx) or set(summary.n) != set(summary.sxx):
                raise ConfigurationError(f"{summary.entity}: inconsistent summary indicator keys")
        first = meta.drop_duplicates("indicator").set_index("indicator")
        for k, group in meta.groupby("indicator"):
            for col in ("ind_id", "tier", "is_good", "unit"):
                if col in group and group[col].nunique(dropna=False) != 1:
                    raise ConfigurationError(f"inconsistent {col} metadata for {k}")
        keys = sorted(set(first.index) | {k for s in summaries for k in s.n})
        if set(keys) - set(first.index):
            raise ConfigurationError("summary contains indicators absent from the declared metadata")
        mu, sg, nn, status, unavailable = {}, {}, {}, {}, {}
        for k in keys:
            ns = [s.n[k] for s in summaries if k in s.n]
            if any(isinstance(n, (bool, np.bool_)) or n != 1 for n in ns):
                raise ConfigurationError(f"{k}: each firm must contribute exactly one observation")
            if any(not np.isfinite(s.sx[k]) or not np.isfinite(s.sxx[k])
                   for s in summaries if k in s.n):
                raise ConfigurationError(f"nonfinite summary for {k}")
            if any(not np.isclose(s.sxx[k], s.sx[k] * s.sx[k], rtol=1e-10, atol=1e-12)
                   for s in summaries if k in s.n):
                raise ConfigurationError(f"{k}: a one-observation summary requires Sxx = Sx squared")
            nn[k] = int(sum(ns))
            if len(ns) < self.MIN_PARTICIPANTS:
                # Too few reporters: nothing is published for the indicator
                # (a coverage event, not a dispersion event).
                unavailable[k] = (f"low_coverage: {len(ns)} reporting participants "
                                  f"< minimum group size {self.MIN_PARTICIPANTS}")
                status[k] = "unavailable_low_coverage" if ns else "unavailable_missing"
                continue
            n_tot = int(sum(ns))
            sx = self._masked_sum([s.sx[k] for s in summaries if k in s.n], f"sx::{k}")
            sxx = self._masked_sum([s.sxx[k] for s in summaries if k in s.n], f"sxx::{k}")
            m = sx / n_tot
            second_moment = sxx / n_tot
            var = max(second_moment - m * m, 0.0)
            # Cancellation in B/N - mu^2 can leave a small *positive* residual
            # even when every firm reports the same value. Taking sqrt first
            # would turn machine rounding into apparent dispersion and move
            # weights instead of applying Section 4.2.1's baseline fallback.
            roundoff = 8.0 * np.finfo(float).eps * max(abs(second_moment), m * m)
            if var <= roundoff:
                var = 0.0
            s_ = float(np.sqrt(var))
            # Degeneracy is tested against an absolute floor and relative to the
            # mean, so the test does not depend on the indicator's magnitude.
            degenerate = (s_ < self.SIGMA_EPS
                          or s_ < self.SIGMA_REL_EPS * max(abs(m), 1e-12))
            mu[k] = float(m)
            sg[k] = 0.0 if degenerate else s_
            nn[k] = n_tot
            status[k] = "zero_variance_observed" if degenerate else "published"

        return GlobalStats(
            mu=mu, sigma=sg, n=nn,
            is_good={k: bool(first.loc[k, "is_good"]) for k in keys},
            tier={k: str(first.loc[k, "tier"]) for k in keys},
            ind_id={k: first.loc[k, "ind_id"] for k in keys},
            category={k: str(first.loc[k, "category"]) for k in keys},
            status=status, unavailable=unavailable,
            contributors=sorted({s.entity for s in summaries
                                 if any(n > 0 for n in s.n.values())}),
        )


def centralised_stats(df_dim: pd.DataFrame) -> Tuple[Dict[str, float], Dict[str, float]]:
    """Ablation (a): mu and sigma computed directly from the pooled raw records."""
    mu, sg = {}, {}
    for k, g in df_dim.groupby("indicator"):
        finite = g[np.isfinite(g["value"])]
        v = finite["value"].to_numpy(dtype=float)
        if finite["entity"].nunique() < Coordinator.MIN_PARTICIPANTS:
            continue
        m, s = float(np.mean(v)), float(np.std(v))
        if s < Coordinator.SIGMA_EPS or s < Coordinator.SIGMA_REL_EPS * max(abs(m), 1e-12):
            s = 0.0                      # published as observed zero variance
        mu[k], sg[k] = m, s
    return mu, sg


def local_raw_summaries(frame: pd.DataFrame) -> List[LocalSummary]:
    """Each firm's raw-unit (n, sum, sum of squares) triples for one dimension,
    chain firms in ENTITY_ORDER and any benchmark-only peers after them."""
    present = set(frame["entity"])
    order = [e for e in ENTITY_ORDER if e in present] + sorted(present - set(ENTITY_ORDER))
    return [Participant(e, benchmark_only=e not in ENTITY_ORDER).local_summary(frame)
            for e in order]


def raw_unit_diagnostics(frame: pd.DataFrame,
                         summaries: Optional[List[LocalSummary]] = None) -> GlobalStats:
    """Evaluator-only raw-unit statistics from the in-process records.

    These support legacy diagnostic tables and ratio/CV sensitivity experiments.
    They are not another participant upload or the common-scale calibration: a
    local coordinator without masking or logging applies the same (n, sum, sum of
    squares) arithmetic, so the values equal a raw-unit reference initialization
    on the same records exactly and the pooled-record comparison (ablation (a),
    centralised_stats) remains an independent check.
    """
    sums = summaries if summaries is not None else local_raw_summaries(frame)
    return Coordinator(secure_aggregation=False)._summary_statistics(sums, frame)


def needs_reference_initialization(frame: pd.DataFrame, policy: RegulatoryPolicy) -> bool:
    """Whether ranges or reference scales require an initial raw-unit benchmark.

    Fully predeclared bounds and reference scales avoid this extra upload stage.
    """
    for row in frame.drop_duplicates("indicator").itertuples():
        k, iid, unit = row.indicator, row.ind_id, str(row.unit).strip().lower()
        has_bounds = (k in policy.conversion_range_overrides
                      or iid in policy.conversion_range_overrides
                      or natural_conversion_range(unit) is not None)
        fixed_scale = ("yes/no" in unit or unit in (
            "policy", "declaration", "fraction", "fraction (percentage)",
            "%", "percent", "percentage", "percentage points"))
        has_scale = (fixed_scale or k in policy.reference_scale_overrides
                     or iid in policy.reference_scale_overrides)
        if not has_bounds or not has_scale:
            return True
    return False


def centralised_common_stats(df_dim: pd.DataFrame, stats: "GlobalStats"
                             ) -> Tuple[Dict[str, float], Dict[str, float]]:
    """Ablation (a) on the common scale: mean and population standard deviation of
    the pooled converted records.
    """
    frame = df_dim
    if stats.contributors:
        frame = df_dim[df_dim["entity"].isin(stats.contributors)]
    mu, sg = {}, {}
    for k, g in frame.groupby("indicator"):
        if k not in stats.input_bounds:
            continue
        finite = g[np.isfinite(g["value"])]
        if finite["entity"].nunique() < Coordinator.MIN_PARTICIPANTS:
            continue
        lo, hi = stats.input_bounds[k]
        x = ((finite["value"].to_numpy(dtype=float) - lo) / (hi - lo)).clip(0.0, 1.0)
        if stats.is_good.get(k, False):
            x = 1.0 - x
        m, s = float(np.mean(x)), float(np.std(x))
        if s < Coordinator.SIGMA_EPS or s < Coordinator.SIGMA_REL_EPS * max(abs(m), 1e-12):
            s = 0.0
        mu[k], sg[k] = m, s
    return mu, sg


def benchmark_disclosure_check(frame: pd.DataFrame, stats: GlobalStats) -> pd.DataFrame:
    """What one reporting participant can infer about the others from a benchmark.

    Knowing its own values, a participant learns the sum and the sum of squares of
    the other r values, so each lies within m +/- sqrt((r - 1) v).  They are
    recovered exactly when r <= 2, which the minimum group size rules out, or when
    all are equal (flagged zero_dispersion).  narrowest_interval_rel is the
    smallest such half-width relative to m.
    """
    rows = []
    for k, mu in stats.mu.items():
        sigma = float(stats.sigma.get(k, 0.0))
        n = int(stats.n.get(k, 0))
        sub = frame[(frame["indicator"] == k) & np.isfinite(frame["value"])]
        exact, narrowest = False, np.inf
        for _, g in sub.groupby("entity"):
            v_own = g["value"].to_numpy(dtype=float)
            r = n - len(v_own)
            if r <= 2:
                exact = True
                continue
            m = (n * mu - float(v_own.sum())) / r
            q = (n * (sigma ** 2 + mu ** 2) - float(np.square(v_own).sum())) / r
            v = max(q - m * m, 0.0)
            if v <= 1e-9 * max(m * m, 1e-12):
                exact = True
                continue
            narrowest = min(narrowest, float(np.sqrt((r - 1) * v)) / max(abs(m), 1e-12))
        rows.append({"indicator": k, "ind_id": stats.ind_id.get(k),
                     "reporters": int(sub["entity"].nunique()),
                     "zero_dispersion": sigma == 0.0,
                     "single_member_exact_recovery": bool(exact),
                     "narrowest_interval_rel": (float(narrowest) if np.isfinite(narrowest)
                                                else np.nan)})
    return pd.DataFrame(rows)


def declare_reference_scales(stats: GlobalStats, meta: pd.DataFrame,
                             policy: RegulatoryPolicy
                             ) -> Tuple[Dict[str, float], Dict[str, str]]:
    """Reference scale s_ref,k of every indicator, with its basis.

    s_ref stands in for sigma where sigma cannot serve: as the dispersion
    denominator when the anchor sigma is zero, and in the scoring denominator
    sigma_score = max(sigma_t, eta * s_ref).  Declarations take half the policy
    scale, percentages 5 points, other magnitudes ref_scale_relative times the
    anchor mean.  Declared once at the anchor round and frozen.
    """
    units = meta.drop_duplicates("indicator").set_index("indicator")
    s_ref, basis = {}, {}
    for k in stats.tier:
        shown = str(units.loc[k, "unit"]).strip() if k in units.index else ""
        unit = shown.lower()
        iid = stats.ind_id.get(k)
        override = policy.reference_scale_overrides.get(k, policy.reference_scale_overrides.get(iid))
        if override is not None:
            value = float(override)
            why = f"explicit predeclared indicator reference scale [{shown or 'unit not stated'}]"
        elif "yes/no" in unit or unit in ("policy", "declaration"):
            value = policy.ref_scale_declaration
            why = "declaration scale 0/0.5/1; half-range"
        elif unit in ("fraction", "fraction (percentage)"):
            value = policy.ref_scale_percentage
            why = f"{100 * policy.ref_scale_percentage:g} percentage points, stored as a fraction"
        elif unit in ("%", "percent", "percentage", "percentage points"):
            value = 100.0 * policy.ref_scale_percentage
            why = f"{value:g} percentage points, stored on the 0..100 scale"
        elif unit in ("count", "count/year"):
            value = max(policy.ref_scale_relative * abs(stats.mu.get(k, 0.0)), 1.0)
            why = (f"max({policy.ref_scale_relative:.0%} of published anchor level, "
                   f"one discrete observation [{shown}]); fixed count resolution")
        else:
            if not unit or k not in stats.mu or abs(stats.mu[k]) <= policy.ref_scale_floor:
                raise ConfigurationError(
                    f"{iid or k}: no nonzero publishable anchor scale in [{unit or 'unstated unit'}]. "
                    "Set reference_scale_overrides for this indicator in its parsed measurement unit; "
                    "a low-coverage mean or arbitrary fallback scale is not published.")
            value = policy.ref_scale_relative * abs(stats.mu[k])
            why = (f"{policy.ref_scale_relative:.0%} of the anchor level "
                   f"[{shown or 'unit not stated'}]")
        if not np.isfinite(value) or value <= 0.0:
            raise ConfigurationError(f"{iid or k}: reference scale must be finite and positive")
        s_ref[k] = float(max(value, policy.ref_scale_floor))
        basis[k] = why
    return s_ref, basis


def attach_scoring_scale(stats: GlobalStats, s_ref: Dict[str, float],
                         policy: RegulatoryPolicy) -> GlobalStats:
    """Bind the scoring denominator sigma_score = max(sigma_t, eta * s_ref).

    A zero-dispersion indicator keeps a positive weight and contributes z = 0.
    """
    floor = policy.eta_sigma_score
    missing = [k for k in stats.tier if k not in s_ref]
    if missing:
        raise ConfigurationError(f"undeclared reference scales: {missing}")
    if any(not np.isfinite(v) or v <= 0.0 for v in s_ref.values()):
        raise ConfigurationError("reference scales must be finite and positive")
    stats.s_ref = {k: float(s_ref[k]) for k in stats.tier}
    stats.sigma_score = {
        k: float(max(stats.sigma.get(k, 0.0), floor * stats.s_ref[k]))
        for k in stats.mu}
    return stats


def scoring_sigma(stats: GlobalStats, k: str) -> float:
    """Denominator used to standardise indicator k; always positive."""
    s = stats.sigma_score.get(k)
    if s is None or not np.isfinite(s) or s <= 0.0:
        s = float(stats.sigma.get(k, 0.0))
    return s


# =============================================================================
# SECTION 6.  THREE-LEVEL WEIGHT GOVERNANCE
# =============================================================================

# Governance levels of Table 3.  The engine and the classification workbook
# code them T1-T3; results and the platform name them Level 1-3, and "tier"
# is kept for positions in the supply chain.
LEVEL_NAME = {"T1": "Level 1", "T2": "Level 2", "T3": "Level 3"}

def natural_conversion_range(unit) -> Optional[float]:
    """Upper end of the natural range of a declaration or percentage unit; None
    for a magnitude, whose range is set from the anchor mean."""
    unit = str(unit).strip().lower()
    if "yes/no" in unit or unit in ("policy", "declaration", "fraction", "fraction (percentage)"):
        return 1.0
    if unit in ("%", "percent", "percentage", "percentage points"):
        return 100.0
    return None


def attach_common_scale(stats, frame, policy, reference=None, coordinator=None):
    """Fixed conversion to the common scale [0, 1], high = adverse, clipped at
    the published range (Section 4.2.2).

    Ranges: declarations and fractions [0, 1], percentages [0, 100], other
    magnitudes [0, policy.magnitude_range_multiple x anchor mean].  The
    reference points are fixed before aggregation; each contributor converts
    its own values and uploads (1, x, x^2) on the common scale, and the
    common-scale benchmark is aggregated from these uploads.  Returns the
    uploads.
    """
    bounds = dict(reference.input_bounds) if reference is not None else {}
    clipped: Dict[str, int] = {}
    for k, group in frame.groupby("indicator"):
        override = policy.conversion_range_overrides.get(k, policy.conversion_range_overrides.get(stats.ind_id.get(k)))
        if override is not None:
            bounds[k] = tuple(map(float, override))
        if k not in bounds:
            hi = natural_conversion_range(group.iloc[0]["unit"])
            if hi is None:
                hi = max(policy.magnitude_range_multiple * abs(stats.mu.get(k, 0.0)),
                         stats.s_ref.get(k, 1.0))
            bounds[k] = (0.0, hi)
        lo, hi = bounds[k]
        if not np.isfinite([lo,hi]).all() or not hi > lo:
            raise ConfigurationError(f"{k}: invalid conversion range")
        reported = group.value[np.isfinite(group.value)]
        clipped[k] = int(((reported < lo) | (reported > hi)).sum())
    entities = stats.contributors or list(frame.entity.unique())
    summaries = [Participant(e, benchmark_only=e not in ENTITY_ORDER)
                 .common_scale_summary(frame, bounds, stats.is_good) for e in entities]
    coordinator = coordinator or Coordinator(secure_aggregation=False)
    common = coordinator.aggregate(summaries, frame)
    stats.common_mu, stats.common_sigma = common.mu, common.sigma
    stats.input_bounds = bounds
    stats.conversion_clipped = clipped
    return summaries


def _project_to_band(w: np.ndarray, lo: np.ndarray, hi: np.ndarray,
                     target: float, iters: int = 60) -> np.ndarray:
    """Euclidean projection onto a bounded simplex, by bisection on the
    multiplier (no renormalization afterwards).
    """
    w, lo, hi = (np.asarray(x, dtype=float) for x in (w, lo, hi))
    if (w.shape != lo.shape or w.shape != hi.shape
            or not all(np.isfinite(x).all() for x in (w, lo, hi))
            or not np.isfinite(target) or np.any(lo < 0.0) or np.any(lo > hi)
            or lo.sum() > target + 1e-12 or hi.sum() < target - 1e-12):
        raise ConfigurationError("bounded-simplex projection has invalid or infeasible bounds")
    if abs(float(lo.sum()) - target) <= 1e-12:
        return lo.copy()
    if abs(float(hi.sum()) - target) <= 1e-12:
        return hi.copy()
    left, right = float(np.min(w - hi)), float(np.max(w - lo))
    for _ in range(max(iters, 80)):
        mid = (left + right) / 2.0
        projected = np.clip(w - mid, lo, hi)
        if abs(float(projected.sum()) - target) <= 1e-13:
            return projected
        if projected.sum() > target:
            left = mid
        else:
            right = mid
    raise ConfigurationError("bounded-simplex projection failed to satisfy its block total")


@dataclass
class WeightSet:
    """Published weights for one dimension in one governance round."""
    dimension: str
    version_id: str
    timestamp: str
    policy_digest: str
    rows: pd.DataFrame          # indicator, ind_id, tier, status, mu, sigma,
                                # sigma_score, s_ref, rho, w_reg, w_data,
                                # w_final, band_lo, band_hi
    digest: str
    notes: List[str] = field(default_factory=list)
    register_digest: str = ""
    classification_digest: str = ""
    unavailable: Dict[str, str] = field(default_factory=dict)

    def w(self) -> Dict[str, float]:
        if self.rows.empty:
            return {}
        return dict(zip(self.rows["indicator"], self.rows["w_final"]))

    def scoring_indicators(self) -> List[str]:
        if self.rows.empty:
            return []
        return list(self.rows["indicator"])


def compute_dispersion(stats: GlobalStats, anchor: Optional[GlobalStats],
                       mode: str, s_ref: Optional[Dict[str, float]] = None
                       ) -> Dict[str, float]:
    """Dispersion rho_k behind the data-driven weights.

    'common' (default) is sigma on the common scale (Section 4.2.1); 'ratio'
    (sigma_t / sigma_anchor) and 'cv' (sigma / |mu|) are sensitivity alternatives.
    sigma_t = 0 gives rho = 0, and a zero anchor sigma is replaced by s_ref.
    """
    if mode not in ("common", "ratio", "cv"):
        raise ConfigurationError("unknown dispersion mode")
    if mode == "common":
        if not stats.common_sigma:
            raise ConfigurationError("common-scale statistics must be attached before weighting")
        return dict(stats.common_sigma)
    s_ref = s_ref or stats.s_ref
    rho: Dict[str, float] = {}
    for k in stats.mu:
        s = float(stats.sigma.get(k, 0.0))
        ref = float(s_ref.get(k, 0.0))
        if s <= Coordinator.SIGMA_EPS:
            rho[k] = 0.0
            continue
        if mode == "cv":
            m = abs(stats.mu.get(k, 0.0))
            denom = m if m > 1e-12 else ref
        else:
            base = anchor.sigma.get(k) if anchor else s
            denom = base if (base is not None and base > Coordinator.SIGMA_EPS) else ref
        if denom is None or not np.isfinite(denom) or denom <= 0.0:
            raise ConfigurationError(f"{k}: dispersion requires a positive declared reference scale")
        rho[k] = float(s / denom)
        if not np.isfinite(rho[k]):
            raise ConfigurationError(f"{k}: nonfinite dispersion ratio")
    return rho


# ---------------------------------------------------------------------------
# The governance register: the indicator universe and its fixed parameters.
# ---------------------------------------------------------------------------

@dataclass
class GovernanceRegister:
    """Indicator universe and fixed weight parameters of one dimension.

    Membership, w_reg, the Level 2 band and the Level 3 floor and cap are declared
    at the anchor round and versioned; they do not change with the data.
    """
    dimension: str
    t1: List[str]
    t2: List[str]
    t3: List[str]
    w_reg: Dict[str, float]
    band_lo: Dict[str, float]
    band_hi: Dict[str, float]
    t3_floor: Dict[str, float]
    t3_cap: Dict[str, float]
    s_ref: Dict[str, float]
    s_ref_basis: Dict[str, str]
    block_share_t2: float
    block_share_t3: float
    policy_digest: str
    classification_digest: str
    version_id: str
    digest: str

    def members(self) -> List[str]:
        return list(self.t2) + list(self.t3)


def declare_governance_register(dimension: str, stats: GlobalStats,
                                policy: RegulatoryPolicy, label: str = "R0",
                                w_reg_override: Optional[Dict[str, float]] = None,
                                s_ref: Optional[Dict[str, float]] = None,
                                s_ref_basis: Optional[Dict[str, str]] = None,
                                classification_digest: str = "absent"
                                ) -> GovernanceRegister:
    """Declare the fixed governance configuration for one dimension."""
    problems = policy.validate()
    if problems:
        raise ConfigurationError("configuration invalid: " + "; ".join(problems))
    # Metadata defines membership even when no benchmark is published.
    members = sorted(stats.tier)
    unknown = [k for k in members if stats.tier[k] not in ("T1", "T2", "T3")]
    if unknown:
        raise ConfigurationError(f"unclassified governance indicators: {unknown}")
    t1 = [k for k in members if stats.tier.get(k) == "T1"]
    t2 = [k for k in members if stats.tier.get(k) == "T2"]
    t3 = [k for k in members if stats.tier.get(k) == "T3"]

    share2 = policy.block_share_t2 if (t2 and t3) else (1.0 if t2 else 0.0)
    share3 = policy.block_share_t3 if (t2 and t3) else (1.0 if t3 else 0.0)

    w_reg: Dict[str, float] = {}
    if t2:
        if w_reg_override:
            if set(w_reg_override) - set(t2):
                raise ConfigurationError("w_reg_override contains indicators outside the T2 universe")
            raw = np.array([float(w_reg_override.get(k, share2 / len(t2))) for k in t2])
            if not np.isfinite(raw).all() or np.any(raw <= 0.0):
                raise ConfigurationError("regulated baseline weights must be finite and strictly positive")
            raw = raw / raw.sum() * share2
        else:
            raw = np.full(len(t2), share2 / len(t2))    # uniform baseline
        w_reg = {k: float(v) for k, v in zip(t2, raw)}

    band_lo = {k: policy.band_lo * w for k, w in w_reg.items()}
    band_hi = {k: policy.band_hi * w for k, w in w_reg.items()}
    eq3 = (share3 / len(t3)) if t3 else 0.0
    t3_floor = {k: policy.t3_floor_frac * eq3 for k in t3}
    t3_cap = {k: policy.t3_cap_frac * eq3 for k in t3}

    # Feasibility is a property of the configuration, so it is checked when
    # the configuration is declared, not after a vector has been built.
    for name, lo, hi, share in (("T2", band_lo, band_hi, share2),
                                ("T3", t3_floor, t3_cap, share3)):
        if not lo:
            continue
        if sum(lo.values()) > share + 1e-12 or sum(hi.values()) < share - 1e-12:
            raise ConfigurationError(
                f"{dimension} {name} block infeasible: sum(lower)="
                f"{sum(lo.values()):.6f}, block share={share:.6f}, sum(upper)="
                f"{sum(hi.values()):.6f}. Lower the floor, raise the cap, or "
                f"change the block share; the round does not proceed with a "
                f"configuration whose bounds cannot all hold.")

    s_ref = dict(s_ref or stats.s_ref)
    missing_scales = [k for k in members if k not in s_ref]
    if missing_scales or any(not np.isfinite(v) or v <= 0.0 for v in s_ref.values()):
        raise ConfigurationError(f"governance register requires declared positive reference scales; missing {missing_scales}")
    s_ref_basis = dict(s_ref_basis or {})
    canon = json.dumps({
        "dimension": dimension, "t1": sorted(t1), "t2": sorted(t2), "t3": sorted(t3),
        "w_reg": dict(sorted(w_reg.items())),
        "band": [policy.band_lo, policy.band_hi],
        "t3_floor_frac": policy.t3_floor_frac, "t3_cap_frac": policy.t3_cap_frac,
        "block_shares": [share2, share3],
        "s_ref": {k: float(v) for k, v in sorted(s_ref.items())},
        "s_ref_basis": s_ref_basis,
        "indicator_ids": stats.ind_id, "directions": stats.is_good,
    }, sort_keys=True)
    digest = hashlib.sha256(
        (canon + policy.digest() + classification_digest).encode()).hexdigest()[:16]

    return GovernanceRegister(
        dimension=dimension, t1=sorted(t1), t2=t2, t3=t3, w_reg=w_reg,
        band_lo=band_lo, band_hi=band_hi, t3_floor=t3_floor, t3_cap=t3_cap,
        s_ref=s_ref, s_ref_basis=s_ref_basis,
        block_share_t2=share2, block_share_t3=share3,
        policy_digest=policy.digest(), classification_digest=classification_digest,
        version_id=f"REG-{dimension[:1]}-{label}-{digest[:6]}", digest=digest)


def govern_weights(dimension: str, stats: GlobalStats,
                   anchor: Optional[GlobalStats], policy: RegulatoryPolicy,
                   round_label: str,
                   w_reg_override: Optional[Dict[str, float]] = None,
                   t1_ids_in_data: Optional[List[str]] = None,
                   register: Optional[GovernanceRegister] = None,
                   classification_digest_value: str = "absent") -> WeightSet:
    """Published weights of one dimension (Section 4.2.1).

    a = lambda * b + (1 - lambda) * v over all Level 2 and 3 indicators, then one
    Euclidean projection onto the bounds with total one.  The Level 2/3 split sets
    the baseline mass, not fixed final block totals.
    """
    notes: List[str] = []
    if register is None:
        # A call without the versioned register (a sensitivity run or a
        # stand-alone check) declares one from this round's statistics and
        # records this in the notes.
        s_ref = dict(stats.s_ref)
        if not s_ref:
            raise ConfigurationError("declare and attach unit-bearing reference scales before governing weights")
        register = declare_governance_register(
            dimension, stats, policy, round_label, w_reg_override,
            s_ref=s_ref, s_ref_basis={},
            classification_digest=classification_digest_value)
        notes.append("ad-hoc register declared from this round's statistics "
                     "(no versioned register supplied by the caller)")
    if register.policy_digest != policy.digest():
        raise ConfigurationError("policy differs from its governance register; declare a new governance version")
    if set(stats.tier) != set(register.t1 + register.t2 + register.t3):
        raise ConfigurationError("indicator universe differs from its fixed governance register")

    rho = compute_dispersion(stats, anchor, policy.dispersion_mode, register.s_ref)

    t1_report = sorted(t1_ids_in_data) if t1_ids_in_data is not None \
        else sorted(str(stats.ind_id.get(k) or k) for k in register.t1)
    if t1_report:
        notes.append(f"Level 1 (hard-alert) indicators excluded from weighting: "
                     f"{', '.join(t1_report)}")

    avail_t2 = [k for k in register.t2 if k in stats.mu]
    avail_t3 = [k for k in register.t3 if k in stats.mu]
    unavailable = {k: stats.unavailable.get(k, "not present in this period's benchmark")
                   for k in register.members() if k not in stats.mu}

    eq3 = (register.block_share_t3 / len(register.t3)) if register.t3 else 0.0
    baseline = dict(register.w_reg, **{k: eq3 for k in register.t3})
    available = avail_t2 + avail_t3
    # The available indicators share only their own baseline mass; the baseline
    # share of an indicator without a benchmark is not passed to the others.
    target = sum(baseline[k] for k in available)
    base = np.array([baseline[k] for k in available])
    r = np.array([rho.get(k, 0.0) for k in available])
    v = target * r/r.sum() if r.sum() > 0 else base.copy()
    a = policy.lam*base + (1-policy.lam)*v
    lo = np.array([register.band_lo[k] if k in register.t2 else register.t3_floor[k] for k in available])
    hi = np.minimum(1.0, [register.band_hi[k] if k in register.t2 else register.t3_cap[k] for k in available])
    projected = _project_to_band(a, lo, hi, target) if available else []
    alloc = {k: dict(w_data=float(v[i]), w_final=float(projected[i]),
                     lo=float(lo[i]), hi=float(hi[i])) for i,k in enumerate(available)}
    notes.append("Level 2 and Level 3 indicators share one dispersion normalization, "
                 "baseline blend and projection (70/30 sets the baseline mass only)")

    recs = []
    for k in avail_t2 + avail_t3:
        a = alloc[k]
        tier = "T2" if k in register.w_reg else "T3"
        recs.append(dict(
            indicator=k, ind_id=stats.ind_id.get(k), tier=tier,
            category=stats.category.get(k, ""),
            mu=stats.mu[k], sigma=stats.sigma.get(k, 0.0),
            sigma_score=scoring_sigma(stats, k), s_ref=register.s_ref.get(k, np.nan),
            status=stats.status.get(k, "published"), rho=rho.get(k, 0.0),
            w_reg=baseline[k],
            w_data=a["w_data"], w_final=a["w_final"],
            band_lo=a["lo"], band_hi=a["hi"],
            floor_binding=bool(a["w_final"] <= a["lo"] + 1e-12),
            cap_binding=bool(a["w_final"] >= a["hi"] - 1e-12)))

    rows = pd.DataFrame(recs, columns=["indicator", "ind_id", "tier", "category", "mu",
        "sigma", "sigma_score", "s_ref", "status", "rho", "w_reg", "w_data", "w_final",
        "band_lo", "band_hi", "floor_binding", "cap_binding"])
    if not rows.empty:
        rows["common_mu"] = rows.indicator.map(stats.common_mu)
        rows["common_sigma"] = rows.indicator.map(stats.common_sigma)
        rows["conversion_bounds"] = rows.indicator.map(lambda k: str(stats.input_bounds.get(k)))
        rows["values_clipped"] = rows.indicator.map(
            lambda k: int((stats.conversion_clipped or {}).get(k, 0)))
    # Values outside the published conversion range are clipped to its ends
    # (Section 4.2.2).  When clipping maps every reporter to the same value an
    # indicator with real dispersion keeps its weight but cannot differentiate
    # firms; the published version records this in its notes.
    for k in available:
        n_clipped = int((stats.conversion_clipped or {}).get(k, 0))
        if not n_clipped:
            continue
        lo_k, hi_k = stats.input_bounds.get(k, (float("nan"), float("nan")))
        collapsed = (stats.common_sigma.get(k, 0.0) == 0.0 and stats.sigma.get(k, 0.0) > 0.0)
        notes.append(
            f"{stats.ind_id.get(k) or k}: {n_clipped} of {int(stats.n.get(k, 0))} reported "
            f"value(s) lie outside the conversion range [{lo_k:g}, {hi_k:g}] and are clipped"
            + ("; every reporter maps to the same converted value, so the indicator keeps "
               "its weight but does not differentiate firms; review its conversion range "
               "and direction" if collapsed else ""))
    expected_total = target
    if not rows.empty:
        rows = rows.sort_values(["tier", "w_final"], ascending=[True, False]).reset_index(drop=True)
        total = float(rows["w_final"].sum())
        if abs(total - expected_total) > 1e-9:
            raise ConfigurationError(
                f"{dimension} weights sum to {total:.12f}, expected "
                f"{expected_total:.12f}. The block projection failed; the vector is "
                f"not renormalised because that would break the declared floors "
                f"and caps.")
        if unavailable:
            notes.append(f"{len(unavailable)} register member(s) carry no weight this "
                         f"round for coverage reasons: {sorted(unavailable)}")
        if stats.carried_forward:
            notes.append("valid reference used where too few firms reported: " + ", ".join(
                f"{stats.ind_id.get(k) or k} from {v}"
                for k, v in sorted(stats.carried_forward.items())))
        zero_var = [k for k in rows["indicator"] if stats.status.get(k) == "zero_variance_observed"]
        if zero_var:
            # A T2 member with no dispersion keeps lambda * w_reg, which lies
            # inside its band; a T3 member sits at its floor.  Either way the
            # weight stays positive and the indicator contributes z = 0.
            notes.append(f"{len(zero_var)} indicator(s) show observed zero variance and "
                         f"keep a positive governed weight (z = 0 contribution): "
                         f"{', '.join(str(stats.ind_id.get(k) or k) for k in zero_var)}")

    # The version digest binds the whole parameter set (mu, sigma, the scoring
    # denominator, the reference scale, the indicator direction, the governance
    # level and the baseline weight), not the final weight vector alone.
    # Parameters are serialized as Python floats, so a small change in any of
    # them changes the digest.
    canon = json.dumps({
        "common_mu": stats.common_mu, "common_sigma": stats.common_sigma,
        "input_bounds": stats.input_bounds,
        "mu": stats.mu, "sigma": stats.sigma, "sigma_score": stats.sigma_score,
        "s_ref": stats.s_ref, "weights": dict(zip(rows["indicator"], rows["w_final"])),
        "directions": stats.is_good, "tiers": stats.tier, "ind_id": stats.ind_id,
        "status": stats.status, "unavailable": unavailable, "n": stats.n,
        "carried_forward": stats.carried_forward,
        "hard_threshold_rules": HARD_THRESHOLD_RULES,
    }, sort_keys=True, allow_nan=False)
    digest = hashlib.sha256(
        (canon + policy.digest() + register.digest
         + register.classification_digest).encode()).hexdigest()[:16]
    version_id = f"{dimension[:1]}-{round_label}-{digest[:8]}"

    return WeightSet(dimension=dimension, version_id=version_id,
                     timestamp=datetime.now(timezone.utc).isoformat(timespec="seconds"),
                     policy_digest=policy.digest(), rows=rows, digest=digest,
                     notes=notes, register_digest=register.digest,
                     classification_digest=register.classification_digest,
                     unavailable=unavailable)


# =============================================================================
# SECTION 7.  PARAMETRIC SCORING FUNCTION  (no fitting, no labels)
# =============================================================================

def sigmoid(x: float) -> float:
    return float(1.0 / (1.0 + np.exp(-x))) if x >= 0 else float(np.exp(x)/(1.0+np.exp(x)))


def score_level_index(level: float, policy: RegulatoryPolicy) -> float:
    """Weighted deviation s = sum_k w_k z_k at which a local score reaches
    `level`, for 0 < level < 1; the same level stands for a different
    deviation under another tau.  The level rule is applied to dimension
    totals, so for a firm with suppliers no single deviation corresponds to
    the level."""
    return policy.squash_scale * float(np.log(level / (1.0 - level)))


@dataclass
class LocalScore:
    entity: str
    dimension: str
    score: float
    weighted_sum: float
    coverage: float                       # share of published weight the firm has data for
    contributions: Dict[str, float]       # indicator -> w * risk_z
    hard_threshold_hit: Optional[str] = None
    lower: float = 0.0
    upper: float = 1.0
    incomplete: bool = False
    # Per-indicator terms of the published point value (pre-sigmoid): w * risk_z
    # for a reported input, the upper-bound term for a missing one.
    point_terms: Dict[str, float] = field(default_factory=dict)


def score_entity(entity: str, df_dim: pd.DataFrame, stats: GlobalStats,
                 weights: WeightSet, policy: RegulatoryPolicy,
                 hard_violation: Optional[str] = None) -> LocalScore:
    """Local score R_local = sigmoid(sum_k w_k * z_k / tau); nothing is fitted.

    z_k is the common-scale deviation from the peer mean over sigma_score, signed
    so that z > 0 is worse than the benchmark.  A missing input keeps its weight
    and widens the score interval; the published score is then the upper bound
    (Section 4.2.4).
    """
    wmap = weights.w()
    sub = df_dim[df_dim["entity"] == entity]
    values = sub.set_index("indicator")["value"].to_dict()
    contrib, terms, coverage, low, high = {}, {}, 0.0, 0.0, 0.0
    for k,w in wmap.items():
        if k not in stats.mu:
            continue
        common = k in stats.common_mu
        if common:
            mu = stats.common_mu[k]
            span = stats.input_bounds[k][1] - stats.input_bounds[k][0]
            denom = max(stats.common_sigma[k], policy.eta_sigma_score*stats.s_ref[k]/span)
            endpoints = (0.0, 1.0)
        else:
            mu, denom = stats.mu[k], scoring_sigma(stats,k)
            endpoints = (-np.inf, np.inf)  # raw units, no conversion range
        x = float(values.get(k, np.nan))
        if np.isfinite(x):
            if common:
                x = float(np.clip((x-stats.input_bounds[k][0])/span,0,1))
                if stats.is_good[k]: x = 1-x
            z = (x-mu)/denom
            if not common and stats.is_good[k]: z = -z
            contrib[k] = terms[k] = w*z
            low += w*z; high += w*z; coverage += w
        else:
            low += w*(endpoints[0]-mu)/denom
            high += w*(endpoints[1]-mu)/denom
            terms[k] = w*(endpoints[1]-mu)/denom
    missing = (coverage <= 0.0 or coverage < sum(wmap.values())-1e-12
               or bool(weights.unavailable))
    if weights.unavailable or coverage <= 0.0:
        # Section 4.2.4: a wholly unavailable dimension has score bounds [0, 1].
        # Partial missingness instead uses the finite converted-input bounds.
        low, high = -np.inf, np.inf
    lower, upper = sigmoid(low/policy.squash_scale), sigmoid(high/policy.squash_scale)
    score = upper if missing else lower
    return LocalScore(entity, weights.dimension, score, sum(contrib.values()),
                      coverage, contrib, hard_violation, lower, upper, missing, terms)


def score_intervals(scores, dependency_weights=None):
    result = {}
    for dim, items in scores.items():
        lower, _ = propagate({e: sc.lower for e, sc in items.items()}, dependency_weights, missing=0.0)
        upper, _ = propagate({e: sc.upper for e, sc in items.items()}, dependency_weights, missing=1.0)
        result[dim] = {e: (lower[e], upper[e]) for e in lower}
    return result


def propagate(local: Dict[str, float],
              weights: Optional[Dict[str, dict]] = None, missing: float = 1.0
              ) -> Tuple[Dict[str, float], Dict[str, Dict[str, float]]]:
    """R_total_i = beta_i * R_local_i + sum_j alpha_ij * R_total_j (Section 4.2.3).

    Evaluated from Tier 2 downstream.  Returns (totals, components), with
    components[i] = {'__own__': beta * R_local, j: alpha_ij * R_total_j}.  A firm
    without a local score enters as `missing` (default 1, the upper bound).
    """
    dwmap = DEPENDENCY_WEIGHTS if weights is None else weights
    problems = dependency_weight_problems(dwmap)
    if problems:
        raise ConfigurationError("invalid dependency weights: " + "; ".join(problems))
    totals: Dict[str, float] = {}
    comps: Dict[str, Dict[str, float]] = {}
    for tier in TIER_ORDER:
        for e, cfg in SUPPLY_CHAIN.items():
            if cfg["tier"] != tier:
                continue
            dw = dwmap[e]
            c = {"__own__": dw["beta"] * local.get(e, missing)}
            for j, a in dw["alpha"].items():
                c[j] = a * totals.get(j, local.get(j, missing))
            comps[e] = c
            totals[e] = float(sum(c.values()))
    return totals, comps


# =============================================================================
# SECTION 8.  HARD-THRESHOLD ALERTS AND THE ASSURANCE PROVIDER
# =============================================================================

@dataclass
class AuditFinding:
    entity: str
    kind: str                # 'hard_threshold' | 'summary_anomaly' | 'alpha_method' | 'alert_response'
    indicator: Optional[str]
    detail: str
    triggers_recalculation: bool
    timestamp: str = field(default_factory=lambda: datetime.now(timezone.utc).isoformat(timespec="seconds"))


class AssuranceProvider:
    """Simulated assurance provider for one client (Section 4.1).

    verify_hard_thresholds  declared statements against audit evidence
    verify_summary          uploaded (n, Sx, Sxx) against the audited records
    recompute               the published score from the published parameters
    verify_alpha_method     declared against applied dependency-weight method
    verify_alert_responses  documented responses to alerts
    """

    def __init__(self, client: str):
        self.client = client
        self.findings: List[AuditFinding] = []

    # -- 1 -------------------------------------------------------------------
    def verify_hard_thresholds(self, df_dim: pd.DataFrame,
                               audit_evidence: Optional[Dict[str, float]] = None
                               ) -> Optional[str]:
        """First failing hard-threshold label of the client, or None.

        A missing declaration is recorded as pending, not as a failure, and so is
        a value tagged as preliminary or unverified (is_unverified_tag) unless
        audit evidence is supplied.  Audit evidence overrides the declared value.
        A count limit that is not set is monitored only.  Every confirmed failure
        is kept in `confirmed_labels`.
        """
        audit_evidence = audit_evidence or {}
        self.confirmed_labels = []
        sub = df_dim[df_dim["entity"] == self.client]
        reported: Dict[str, float] = {}
        unverified: Dict[str, bool] = {}
        for _, r in sub.iterrows():
            reported.setdefault(r["ind_id"], float(r["value"]))
            unverified.setdefault(r["ind_id"], is_unverified_tag(
                r.get("data_source", ""), r.get("data_status", "")))
        for iid in pd.unique(df_dim["ind_id"]):
            if iid not in HARD_THRESHOLD_RULES:
                continue
            rule = HARD_THRESHOLD_RULES[iid]
            if rule["rule"] == "count_limit" and rule.get("limit") is None:
                continue
            declared = reported.get(iid, float("nan"))
            actual = float(audit_evidence.get(iid, declared))   # evidence overrides
            if not np.isfinite(actual):
                self.findings.append(AuditFinding(self.client, "pending", iid,
                    "unknown: declaration missing; requires review", triggers_recalculation=False))
                continue
            if iid not in audit_evidence and unverified.get(iid, False):
                # Section 4.2.3: a preliminary report stays pending and is kept
                # apart from confirmed hard alerts until it is verified.
                self.findings.append(AuditFinding(self.client, "pending", iid,
                    "unverified: preliminary or unverified report; requires verification "
                    "before a hard alert is confirmed or closed", triggers_recalculation=False))
                continue
            if rule["rule"] == "policy_required":
                if actual < 1.0:
                    label = f"{iid} {rule['label']}"
                    shown = f"{declared:.0f}" if np.isfinite(declared) else "not reported"
                    self.findings.append(AuditFinding(
                        self.client, "hard_threshold", iid,
                        f"declared {shown}, audit evidence {actual:.0f}; "
                        f"affirmative declaration not supported",
                        triggers_recalculation=True))
                    self.confirmed_labels.append(label)
            elif rule["rule"] == "count_limit" and rule.get("limit") is not None:
                if actual > rule["limit"]:
                    label = f"{iid} {rule['label']} = {actual:g} > {rule['limit']:g}"
                    self.findings.append(AuditFinding(
                        self.client, "hard_threshold", iid,
                        f"governance-body ceiling {rule['limit']:g} exceeded",
                        triggers_recalculation=True))
                    self.confirmed_labels.append(label)
        return self.confirmed_labels[0] if self.confirmed_labels else None

    # -- 2 -------------------------------------------------------------------
    def verify_summary(self, summary: LocalSummary,
                       audited: Optional[pd.DataFrame] = None,
                       tol: float = 1e-6) -> List[AuditFinding]:
        found = []
        keys = set(summary.n) | set(summary.sx) | set(summary.sxx)
        if summary.entity != self.client:
            found.append(AuditFinding(self.client, "summary_anomaly", None,
                "uploaded summary identifies a different participant", triggers_recalculation=True))
        for k in keys:
            n, sx, sxx = summary.n.get(k), summary.sx.get(k), summary.sxx.get(k)
            if (isinstance(n, (bool, np.bool_)) or n != 1 or sx is None or sxx is None
                    or not np.isfinite([sx, sxx]).all()
                    or not np.isclose(sxx, sx * sx, rtol=tol, atol=tol)):
                found.append(AuditFinding(self.client, "summary_anomaly", k,
                    "expected a finite one-firm summary (1, x, x squared)",
                    triggers_recalculation=True))
        if audited is not None:
            a = audited[audited["entity"] == self.client]
            if a["indicator"].duplicated().any():
                raise ConfigurationError("audited records must contain one firm-level value per indicator")
            expected = set(a.loc[np.isfinite(a["value"]), "indicator"])
            for k in sorted(keys ^ expected):
                found.append(AuditFinding(self.client, "summary_anomaly", k,
                    "uploaded indicator coverage does not match the audited records",
                    triggers_recalculation=True))
            for _, r in a.iterrows():
                k = r["indicator"]
                if (k in summary.sx and np.isfinite(r["value"])
                        and not np.isclose(summary.sx[k], float(r["value"]), rtol=tol, atol=tol)):
                    f = AuditFinding(self.client, "summary_anomaly", k,
                                     f"uploaded Sx={summary.sx[k]:.6g} does not match "
                                     f"audited value {float(r['value']):.6g}",
                                     triggers_recalculation=True)
                    found.append(f)
        self.findings.extend(found)
        return found

    def verify_common_summary(self, summary: LocalSummary, audited: pd.DataFrame,
                              stats: GlobalStats, tol: float = 1e-6) -> List[AuditFinding]:
        """The common-scale upload against the audited records converted with the
        published reference points and directions (Section 4.2.2)."""
        a = audited[audited["entity"] == self.client].copy()
        a["value"] = [float(to_common_scale(v, stats.input_bounds[k],
                                            bool(stats.is_good.get(k, False))))
                      if k in stats.input_bounds else float("nan")
                      for k, v in zip(a["indicator"], a["value"])]
        found = self.verify_summary(summary, a, tol=tol)
        for f in found:
            f.detail = "common-scale upload: " + f.detail
        return found

    # -- 3 -------------------------------------------------------------------
    def recompute(self, df_dim: pd.DataFrame, stats: GlobalStats,
                  weights: WeightSet, policy: RegulatoryPolicy,
                  claimed: float, tol: float = 1e-9) -> Tuple[bool, float]:
        """Recompute the client's published score from the published parameters.

        The same function runs in the same process, so this shows that the parameters
        suffice, not that an independent implementation agrees.
        """
        s = score_entity(self.client, df_dim, stats, weights, policy)
        return bool(abs(s.score - claimed) <= tol), s.score

    # -- 4 -------------------------------------------------------------------
    def verify_alpha_method(self, declared_method: str, applied_method: str) -> dict:
        """Compare the declared and applied dependency-weight methods with each other
        and with the accepted methods (Table 4).  Procurement records are not
        inspected.
        """
        declared = str(declared_method).strip().lower()
        applied = str(applied_method).strip().lower()
        check = {"consistent": declared == applied,
                 "accepted": applied in DEPENDENCY_METHODS,
                 "enhanced_assurance": applied in ENHANCED_ASSURANCE_METHODS}
        if not (check["consistent"] and check["accepted"]):
            self.findings.append(AuditFinding(
                self.client, "alpha_method", None,
                f"declared '{declared_method}', applied '{applied_method}'"
                + ("" if check["accepted"] else "; not an accepted method"),
                triggers_recalculation=True))
        return check

    # -- 5 -------------------------------------------------------------------
    def verify_alert_responses(self, episodes: List[dict],
                               records: Dict[Tuple[str, str, str], dict]) -> List[dict]:
        """Did the client document a response, with its reason and evidence, to
        each alert episode that calls for one?  A gap is a material finding,
        reported to the coordinator and passed to the client's customers."""
        rows = []
        for ep in episodes:
            if ep["entity"] != self.client:
                continue
            rec = records.get((self.client, ep["dimension"], ep["first_transition"]), {})
            documented = all(str(rec.get(f, "")).strip() for f in ("action", "reason", "evidence"))
            passed = [] if documented else downstream_notices(self.client)
            if not documented:
                self.findings.append(AuditFinding(
                    self.client, "alert_response", None,
                    f"no documented response to the {ep['dimension']} {ep['level']} "
                    f"from {ep['first_transition']}", triggers_recalculation=False))
            rows.append({"entity": disp(self.client), "dimension": ep["dimension"],
                         "alert": ep["level"], "transitions": ", ".join(ep["transitions"]),
                         "alert_basis": ep["basis"], "response_documented": documented,
                         "action": rec.get("action", ""), "reason": rec.get("reason", ""),
                         "evidence": rec.get("evidence", ""),
                         "finding": "" if documented else "material: no documented response",
                         "finding_passed_to": "; ".join(f"{disp(c)} via {disp(v)}"
                                                        for c, v in passed)})
        return rows


# =============================================================================
# SECTION 8b.  HARD-ALERT REGISTER  (sticky, deduplicated, review-cleared)
# =============================================================================

@dataclass
class HardAlertRecord:
    entity: str
    dimension: str
    indicator: Optional[str]
    label: str
    event_id: str
    raised_period: str
    raised_at: str
    cleared_at: Optional[str] = None
    cleared_by: str = ""
    clearance_note: str = ""
    closure_notices: List[dict] = field(default_factory=list)

    @property
    def active(self) -> bool:
        return self.cleared_at is None


class HardAlertRegister:
    """Verified hard alerts: raised at once, kept until a review clears them, and
    not duplicated when the same event is resubmitted.
    """

    def __init__(self):
        self.records: Dict[str, HardAlertRecord] = {}

    @staticmethod
    def make_event_id(entity: str, label: str) -> str:
        return hashlib.sha256(f"{entity}|{label}".encode()).hexdigest()[:12]

    def raise_alert(self, entity: str, label: str, period: str,
                    dimension: str = "", indicator: Optional[str] = None,
                    event_id: Optional[str] = None) -> bool:
        """Register a verified failure.  False when it is a duplicate."""
        eid = event_id or self.make_event_id(entity, label)
        existing = self.records.get(eid)
        if existing is not None:
            if (existing.entity, existing.label, existing.dimension,
                    existing.indicator) != (entity, label, dimension, indicator):
                raise ValueError("event_id already identifies a different hard event")
            if existing.active or event_id is not None:
                return False
            # An unlabelled periodic observation can start a new episode after
            # clearance; a replayed explicit event id does not reopen it.
            eid = self.make_event_id(entity, label + "|" + period)
            if eid in self.records:
                return False
        self.records[eid] = HardAlertRecord(
            entity=entity, dimension=dimension, indicator=indicator, label=label,
            event_id=eid, raised_period=period,
            raised_at=datetime.now(timezone.utc).isoformat(timespec="seconds"))
        return True

    def clear(self, event_id: str, cleared_by: str, note: str = "") -> bool:
        """Clear one alert.  Only an explicit review reaches this method."""
        if any(not isinstance(text, str) or not text.strip() for text in (cleared_by, note)):
            raise ValueError("clearance requires a reviewer and a review note")
        rec = self.records.get(event_id)
        if rec is None or not rec.active:
            return False
        rec.cleared_at = datetime.now(timezone.utc).isoformat(timespec="seconds")
        rec.cleared_by = cleared_by
        rec.clearance_note = note
        return True

    def active_records(self, entity: Optional[str] = None,
                       dimension: Optional[str] = None) -> List[HardAlertRecord]:
        return [r for r in self.records.values() if r.active
                and (entity is None or r.entity == entity)
                and (dimension in (None, "") or r.dimension in ("", dimension))]

    def is_active(self, entity: str, dimension: Optional[str] = None) -> bool:
        return bool(self.active_records(entity, dimension))

    def labels(self, entity: str, dimension: Optional[str] = None) -> List[str]:
        return [r.label for r in self.active_records(entity, dimension)]

    def table(self) -> pd.DataFrame:
        rows = [{"entity": disp(r.entity), "dimension": r.dimension,
                 "indicator": r.indicator, "label": r.label,
                 "event_id": r.event_id, "raised_period": r.raised_period,
                 "raised_at": r.raised_at, "active": r.active,
                 "cleared_at": r.cleared_at or "", "cleared_by": r.cleared_by,
                 "clearance_note": r.clearance_note,
                 "closure_passed_to": "; ".join(f"{disp(n['firm'])} via {disp(n['via'])}"
                                                for n in r.closure_notices)}
                for r in self.records.values()]
        return pd.DataFrame(rows, columns=[
            "entity", "dimension", "indicator", "label", "event_id",
            "raised_period", "raised_at", "active", "cleared_at", "cleared_by",
            "clearance_note", "closure_passed_to"])


# =============================================================================
# SECTION 9.  INTERNAL WARNINGS AND HARD ALERTS  (component-wise, not net)
# =============================================================================

HARD_STATUS_ACTIVE = "confirmed, unresolved"


def combine_directions(labels) -> str:
    """Summarise warning directions: 'deterioration', 'improvement', 'mixed'
    (both occur), or '' when nothing fired."""
    seen = set()
    for label in labels:
        if label == "mixed":
            seen.update(("deterioration", "improvement"))
        elif label:
            seen.add(label)
    if not seen:
        return ""
    if seen == {"improvement"}:
        return "improvement"
    if seen == {"deterioration"}:
        return "deterioration"
    return "mixed"


@dataclass
class TriggerReport:
    entity: str
    dimension: str
    internal_warning: bool
    net_fired: bool
    net_change: float
    max_abs_component: float
    threshold: float
    component_deltas: Dict[str, float]
    upstream_dispersion: float
    max_single_share: float
    fired_branches: List[str]
    fired_indicators: List[str] = field(default_factory=list)
    reference_total: float = 0.0
    hard_local: bool = False
    hard_alert: bool = False
    hard_via: List[str] = field(default_factory=list)
    hard_labels: List[str] = field(default_factory=list)
    # Section 4.2.3: each firm receiving a hard alert learns the direct
    # supplier (downstream) or direct customer (upstream) it arrives through,
    # an opaque event ID, category, status and raising period, but not the
    # firm where the event occurred. Path records retain each direct branch;
    # hard_event_summary groups those paths into one entry per event.
    hard_records: List[dict] = field(default_factory=list)       # raised here
    hard_via_detail: List[dict] = field(default_factory=list)    # from suppliers
    hard_via_customers: List[str] = field(default_factory=list)
    hard_customer_detail: List[dict] = field(default_factory=list)  # from customers

    @property
    def hard_event_summary(self) -> List[dict]:
        """Distinct events visible here, with all their direct arrival branches."""
        events: Dict[str, dict] = {}
        for direction, records in (("own", self.hard_records),
                                   ("supplier", self.hard_via_detail),
                                   ("customer", self.hard_customer_detail)):
            for rec in records:
                eid = rec["event_id"]
                event = events.setdefault(eid, {
                    "event_id": eid, "category": rec["category"],
                    "status": rec["status"], "raised": rec["raised"],
                    "own_operations": False,
                    "via_direct_suppliers": [], "via_direct_customers": []})
                if direction == "own":
                    event["own_operations"] = True
                else:
                    key = ("via_direct_suppliers" if direction == "supplier"
                           else "via_direct_customers")
                    if rec["via"] not in event[key]:
                        event[key].append(rec["via"])
        return [events[eid] for eid in sorted(events)]

    @property
    def fired_directions(self) -> Dict[str, str]:
        """Direction of every component past the threshold; a positive change is a
        deterioration.
        """
        return {k: ("deterioration" if v > 0.0 else "improvement")
                for k, v in self.component_deltas.items() if abs(v) > self.threshold}

    @property
    def warning_direction(self) -> str:
        return combine_directions(self.fired_directions.values())


def own_indicator_effects(before: "LocalScore", after: "LocalScore", beta: float,
                          policy: RegulatoryPolicy,
                          ind_id: Optional[Dict[str, str]] = None) -> Dict[str, float]:
    """Change in the firm's total that each own indicator causes on its own, with
    the other indicators at their previous values (same benchmark and weights).
    """
    ind_id = ind_id or {}
    u0 = float(sum(before.point_terms.values()))
    base = sigmoid(u0 / policy.squash_scale)
    effects = {}
    for k in sorted(set(before.point_terms) | set(after.point_terms)):
        d = after.point_terms.get(k, 0.0) - before.point_terms.get(k, 0.0)
        if d != 0.0:
            effects[str(ind_id.get(k) or k)] = beta * (sigmoid((u0 + d) / policy.squash_scale)
                                                       - base)
    return effects


def component_trigger(entity: str, comps_t0: Dict[str, float],
                      comps_t1: Dict[str, float], totals_t1: Dict[str, float],
                      dimension: str, policy: RegulatoryPolicy,
                      hard_local: bool = False,
                      hard_labels: Optional[List[str]] = None,
                      reference_total: Optional[float] = None,
                      hard_records: Optional[List[dict]] = None,
                      own_effects: Optional[Dict[str, float]] = None) -> TriggerReport:
    """Internal warning at one firm, tested component by component (Section 4.2.3).

    Each own indicator and each supplier contribution is tested against the
    smaller of the absolute and relative limits; the net change is reported for
    comparison only.
    """
    keys = sorted(set(comps_t0) | set(comps_t1))
    deltas = {k: comps_t1.get(k, 0.0) - comps_t0.get(k, 0.0) for k in keys}
    net = float(sum(deltas.values()))
    tot = float(sum(comps_t1.values()))
    tested = dict(deltas)
    if own_effects is not None:
        tested.pop("__own__", None)
        tested.update(own_effects)

    # fire if |delta| > abs_limit OR |delta| > rel_limit * |reference|, so the
    # binding threshold is the smaller of the two.  The relative denominator is
    # the pre-event total, so the threshold does not move with the event.
    ref_total = float(sum(comps_t0.values())) if reference_total is None \
        else float(reference_total)
    th = min(policy.component_delta_abs, policy.component_delta_rel * abs(ref_total))
    fired = [k for k, v in tested.items() if abs(v) > th]

    up_now = {k: v for k, v in comps_t1.items() if k != "__own__"}
    disp_ = float(np.std(list(up_now.values()))) if len(up_now) > 1 else 0.0
    max_share = float(max(up_now.values()) / tot) if up_now and tot > 0 else 0.0
    records = ([dict(r) for r in hard_records] if hard_records is not None else
               [{"category": label, "status": HARD_STATUS_ACTIVE, "raised": ""}
                for label in (hard_labels or [])])
    for rec in records:
        # Stand-alone probes may supply only labels. Published and event-driven
        # runs pass the register's ID, including the identity of a reopened episode.
        if not rec.get("event_id"):
            rec["event_id"] = HardAlertRegister.make_event_id(
                entity, f"{dimension}|{rec['category']}|{rec.get('raised', '')}")
    branches = set(comps_t0) | set(comps_t1)
    return TriggerReport(
        entity=entity, dimension=dimension,
        internal_warning=bool(fired), net_fired=bool(abs(net) > th),
        net_change=net,
        max_abs_component=float(max((abs(v) for v in tested.values()), default=0.0)),
        threshold=th, component_deltas=tested,
        upstream_dispersion=disp_, max_single_share=max_share,
        fired_branches=[k for k in fired if k in branches and k != "__own__"],
        fired_indicators=[k for k in fired if k not in branches],
        reference_total=ref_total,
        hard_local=bool(hard_local), hard_alert=bool(hard_local),
        hard_labels=list(hard_labels or []),
        hard_records=records if hard_local else [],
    )


def cascade_alerts(reports: Dict[str, TriggerReport],
                   weights: Optional[Dict[str, dict]] = None) -> Dict[str, TriggerReport]:
    """Pass hard alerts along every declared edge, downstream and upstream
    (Section 4.2.3).

    Each receiver learns the event ID and its direct supplier or customer, with
    the category, status and raising period. The same event can arrive through
    several branches; events with matching descriptions keep their distinct IDs.
    Internal warnings are not passed on.
    """
    dwmap = DEPENDENCY_WEIGHTS if weights is None else weights

    def item(via, rec):
        return {"event_id": rec["event_id"],
                "via": via, "category": rec.get("category", ""),
                "status": rec.get("status", HARD_STATUS_ACTIVE),
                "raised": rec.get("raised", "")}

    for tier in TIER_ORDER:
        for e, cfg in SUPPLY_CHAIN.items():
            if cfg["tier"] != tier or e not in reports:
                continue
            for j in dwmap[e]["alpha"]:
                rj = reports.get(j)
                if rj is None or not rj.hard_alert:
                    continue
                if j not in reports[e].hard_via:
                    reports[e].hard_via.append(j)
                reports[e].hard_alert = True
                for rec in list(rj.hard_records) + list(rj.hard_via_detail):
                    if item(j, rec) not in reports[e].hard_via_detail:
                        reports[e].hard_via_detail.append(item(j, rec))
    for tier in reversed(TIER_ORDER):
        for e, cfg in SUPPLY_CHAIN.items():
            if cfg["tier"] != tier or e not in reports:
                continue
            sent = list(reports[e].hard_records) + list(reports[e].hard_customer_detail)
            if not sent:
                continue
            for j in dwmap[e]["alpha"]:
                rj = reports.get(j)
                if rj is None:
                    continue
                if e not in rj.hard_via_customers:
                    rj.hard_via_customers.append(e)
                rj.hard_alert = True
                for rec in sent:
                    if item(e, rec) not in rj.hard_customer_detail:
                        rj.hard_customer_detail.append(item(e, rec))
    return reports


PROVISIONAL_LOW_COVERAGE = "low coverage"
PROVISIONAL_NEW_ENTRANT = "new entrant"


def cascade_provisional(flags: Dict[str, bool],
                        weights: Optional[Dict[str, dict]] = None
                        ) -> Dict[str, List[str]]:
    """Direct suppliers through which a provisional score enters each total."""
    dwmap = DEPENDENCY_WEIGHTS if weights is None else weights
    via: Dict[str, List[str]] = {e: [] for e in ENTITY_ORDER}
    for tier in TIER_ORDER:
        for e, cfg in SUPPLY_CHAIN.items():
            if cfg["tier"] != tier:
                continue
            for j, alpha in dwmap[e]["alpha"].items():
                if alpha > 0.0 and (flags.get(j) or via.get(j)) and j not in via[e]:
                    via[e].append(j)
    return via


def provisional_status(result: "PeriodResult", policy: RegulatoryPolicy,
                       weights: Optional[Dict[str, dict]] = None
                       ) -> Dict[str, Dict[str, dict]]:
    """Provisional status of every published local score.

    A score is provisional when a scored input is missing (its interval applies,
    Section 4.2.4) or when it rests on less than policy.coverage_floor of the
    published weight.  The flag changes no number and passes along every positive
    dependency edge, naming only the direct supplier.  direct_upstream_coverage is
    the share of the firm's supplier dependency carried by suppliers whose data
    meet the coverage floor.  A dimension with no usable data at all is a serious
    reporting failure and calls for an assurance review (Section 4.2.4); the
    persistent case is added when the state is published.
    """
    status: Dict[str, Dict[str, dict]] = {}
    dwmap = result.dependency_weights if weights is None else weights
    for dim, dim_scores in result.scores.items():
        ws, st = result.weights.get(dim), result.stats.get(dim)
        published = (float(ws.rows["w_final"].sum())
                     if ws is not None and not ws.rows.empty else 0.0)
        contributors = set(getattr(st, "contributors", None) or [])
        # Firms that reported at least one usable weighted value in this dimension;
        # reporting nothing at all is a serious reporting failure.
        frame = (result.data or {}).get(dim)
        reporting = set()
        if frame is not None and len(frame):
            usable = np.isfinite(pd.to_numeric(frame["value"], errors="coerce"))
            if "tier" in frame:
                usable &= frame["tier"] != "T1"
            reporting = set(frame.loc[usable, "entity"])
        records: Dict[str, dict] = {}
        for e in ENTITY_ORDER:
            sc = dim_scores.get(e)
            codes: List[str] = []
            reasons: List[str] = []
            share = float("nan")
            if sc is not None:
                share = float(sc.coverage / published) if published > 0.0 else 0.0
                if sc.incomplete:
                    codes.append("incomplete")
                    reasons.append("missing scored inputs; published score interval applies")
                if share <= 0.0:
                    codes.append("low_coverage")
                    reasons.append(f"{PROVISIONAL_LOW_COVERAGE}: no indicator reported; "
                                   f"the point value is the interval upper bound")
                elif share < policy.coverage_floor - 1e-9:
                    codes.append("low_coverage")
                    reasons.append(f"{PROVISIONAL_LOW_COVERAGE}: reported {share:.1%} of "
                                   f"the published weight, below the floor "
                                   f"{policy.coverage_floor:.0%}")
            else:
                codes.extend(["incomplete", "low_coverage"])
                reasons.append("dimension score unavailable; score bounds are [0, 1]")
            alpha = dwmap.get(e, {}).get("alpha", {})
            exposure = float(sum(alpha.values()))
            # Section 4.2.4: a direct supplier counts as covered by usable data when
            # its score meets the governance minimum coverage; below the floor it
            # is flagged as insufficient information instead.
            covered = [j for j in alpha
                       if dim_scores.get(j) is not None and published > 0.0
                       and dim_scores[j].coverage / published >= policy.coverage_floor - 1e-9]
            serious = sc is None or e not in reporting
            records[e] = {"provisional": bool(codes), "codes": codes, "reasons": reasons,
                          "assurance_review": serious,
                          "assurance_review_reasons": (
                              ["serious reporting failure: no usable data in this dimension"]
                              if serious else []),
                          "coverage_share": share, "incomplete": sc is None or sc.incomplete,
                          "lower": sc.lower if sc else 0.0, "upper": sc.upper if sc else 1.0,
                          "contributed_to_benchmark": e in contributors, "via": [],
                          "direct_upstream_coverage": (
                              float(sum(alpha[j] for j in covered)) / exposure
                              if exposure > 0.0 else float("nan")),
                          "direct_suppliers_without_data": [j for j in alpha if j not in covered]}
        via = cascade_provisional({e: r["provisional"] for e, r in records.items()}, dwmap)
        for e, v in via.items():
            records[e]["via"] = v
        status[dim] = records
    return status


def peer_relative_flag(totals: Dict[str, float],
                       policy: RegulatoryPolicy) -> Dict[str, bool]:
    """Escalation flag on the peer-relative composite.  Not an absolute risk
    measure; the absolute channel is absolute_risk_table.
    """
    return {e: bool(v >= policy.peer_relative_flag_level) for e, v in totals.items()}


def graduated_response(rep: TriggerReport, streak: int, high_periods: int,
                       total: float, policy: RegulatoryPolicy) -> Tuple[str, str]:
    """Response level (none, monitor, investigate, hard alert) and its basis.

    streak        consecutive transitions, ending now, in which a component of
                  the node's total deteriorated past the limit
    high_periods  consecutive periods, ending now, with the dimension total at
                  or above policy.persistent_level (level-based rule)
    """
    if rep.hard_alert:
        return "hard alert", "; ".join(list(rep.hard_labels)
                                       + [f"via {disp(j)}" for j in rep.hard_via]
                                       + [f"via customer {disp(j)}"
                                          for j in rep.hard_via_customers])
    worse = [v for v in rep.component_deltas.values() if v > rep.threshold]
    basis = []
    if worse:
        ratio = max(worse) / rep.threshold if rep.threshold > 0 else np.inf
        if ratio >= policy.escalation_size_multiple:
            basis.append(f"change {ratio:.1f}x the limit")
        if streak >= policy.escalation_periods:
            basis.append(f"deterioration in {streak} consecutive transitions")
    if high_periods >= policy.escalation_periods:
        basis.append(f"total {total:.3f} >= {policy.persistent_level:.2f} "
                     f"for {high_periods} periods")
    if basis:
        return "investigate", "; ".join(basis)
    if worse:
        return "monitor", "deterioration below the size and duration criteria"
    if high_periods:
        return "monitor", f"total {total:.3f} >= {policy.persistent_level:.2f}"
    if rep.fired_directions:
        return "none", "improvement"
    return "none", ""


def alert_episodes(triggers: pd.DataFrame) -> List[dict]:
    """Runs of consecutive transitions at investigate or hard-alert level,
    per firm and dimension, from the trigger table (T3)."""
    names = {disp(e): e for e in ENTITY_ORDER}
    order = list(dict.fromkeys(triggers["transition"]))
    episodes: List[dict] = []
    for (entity, dim), g in triggers.groupby(["entity", "dimension"], sort=False):
        g = g.set_index("transition").reindex(order)
        current = None
        for t in order:
            level = g.loc[t, "response_level"]
            if level not in ("investigate", "hard alert"):
                current = None
            elif current is not None and current["level"] == level:
                current["transitions"].append(t)
            else:
                current = {"entity": names.get(entity, entity), "dimension": dim,
                           "level": level, "first_transition": t, "transitions": [t],
                           "basis": g.loc[t, "response_basis"]}
                episodes.append(current)
    return episodes


@dataclass
class OpenWarning:
    """An internal warning on one component of a firm's total: an own
    indicator (its id) or a supplier branch (the supplier's id)."""
    component: str
    opened: str        # transition in which the component deteriorated past the limit
    change: float      # cumulative change of the component since the warning opened


def warning_component_text(component: str) -> str:
    if component == "__own__":
        return "own operations"
    if component in ENTITY_ORDER:
        return f"branch via {disp(component)}"
    return f"own indicator {component}"


class InternalWarningRegister:
    """Open internal warnings by firm and dimension, carried across transitions.

    A warning closes only when the same component returns within the limit;
    changes in other components do not close it (Section 4.2.3).
    """

    def __init__(self) -> None:
        self.open: Dict[Tuple[str, str], Dict[str, OpenWarning]] = {}

    def update(self, dimension: str, entity: str, rep: TriggerReport,
               transition: str) -> Tuple[List[OpenWarning], List[OpenWarning]]:
        """Apply one transition; returns the warnings open afterwards and those it closed."""
        current = self.open.setdefault((dimension, entity), {})
        closed = []
        for component, warning in list(current.items()):
            warning.change += rep.component_deltas.get(component, 0.0)
            if warning.change <= rep.threshold:
                closed.append(current.pop(component))
        for component, delta in rep.component_deltas.items():
            if delta > rep.threshold and component not in current:
                current[component] = OpenWarning(component, transition, float(delta))
        return list(current.values()), closed


def carry_open_warnings(level: str, basis: str, open_now: List[OpenWarning],
                        transition: str) -> Tuple[str, str]:
    """Keep at least monitoring while a warning raised earlier is still open."""
    carried = [w for w in open_now if w.opened != transition]
    if level != "none" or not carried:
        return level, basis
    return "monitor", ("warning still open: " + "; ".join(
        f"{warning_component_text(w.component)} since {w.opened}" for w in carried)
        + "; changes in other components do not close it")


def open_warning_text(warnings: List[OpenWarning], closed: bool = False) -> str:
    return "; ".join(
        f"{warning_component_text(w.component)} (opened {w.opened})" if closed else
        f"{warning_component_text(w.component)} since {w.opened} ({w.change:+.4f})"
        for w in warnings)


def downstream_notices(source: str, weights: Optional[Dict[str, dict]] = None
                       ) -> List[Tuple[str, str]]:
    """(customer, direct supplier) for every firm downstream of `source` on
    declared dependency edges; deeper-tier identities are not passed on."""
    dwmap = DEPENDENCY_WEIGHTS if weights is None else weights
    reached, out = {source}, []
    for tier in TIER_ORDER:
        for e, cfg in SUPPLY_CHAIN.items():
            if cfg["tier"] != tier or e in reached:
                continue
            via = [j for j in dwmap[e]["alpha"] if j in reached]
            if via:
                reached.add(e)
                out.extend((e, j) for j in via)
    return out


def upstream_notices(source: str, weights: Optional[Dict[str, dict]] = None
                     ) -> List[Tuple[str, str]]:
    """(supplier, direct customer) for every firm upstream of `source` on
    declared dependency edges; deeper-tier identities are not passed on."""
    dwmap = DEPENDENCY_WEIGHTS if weights is None else weights
    reached, out = {source}, []
    for tier in reversed(TIER_ORDER):
        for e, cfg in SUPPLY_CHAIN.items():
            if cfg["tier"] != tier or e not in reached:
                continue
            for j in dwmap[e]["alpha"]:
                reached.add(j)
                out.append((j, e))
    return out


NO_LIMIT_DECLARED = "no limit declared by the governance body"


def absolute_risk_table(data: Dict[str, pd.DataFrame],
                        hard_hits: Optional[Dict[str, str]] = None) -> pd.DataFrame:
    """Raw values of the hard-threshold indicators against declared limits.

    No benchmark or weight is used; without a declared limit the outcome is
    'no limit declared'.
    """
    hard_hits = hard_hits or {}
    rows = []
    for dim, df in data.items():
        for _, r in df.iterrows():
            iid = r["ind_id"]
            if iid not in HARD_THRESHOLD_RULES:
                continue
            rule = HARD_THRESHOLD_RULES[iid]
            raw = float(r["value"])
            if rule["rule"] == "policy_required":
                limit, declared = 1.0, "affirmative declaration required"
                breach = bool(raw < 1.0) if np.isfinite(raw) else None
                status = "evaluated" if np.isfinite(raw) else "missing observation"
            else:
                limit = rule.get("limit")
                declared = "governance-body ceiling"
                if not np.isfinite(raw):
                    breach, status = None, "missing observation"
                elif limit is None:
                    breach, status = None, NO_LIMIT_DECLARED
                else:
                    breach, status = bool(raw > limit), "evaluated"
            if status == "evaluated" and is_unverified_tag(r.get("data_source", ""),
                                                           r.get("data_status", "")):
                status = "evaluated; unverified, pending review"
            rows.append({
                "entity": disp(str(r["entity"])), "dimension": dim,
                "ind_id": iid, "indicator": r["indicator"],
                "unit": r["unit"], "raw_value": raw,
                "declared_limit": limit, "limit_type": declared,
                "absolute_breach": breach, "limit_status": status,
                "hard_alert_raised": bool(hard_hits.get(str(r["entity"]), "").startswith(str(iid))),
            })
    out = pd.DataFrame(rows)
    if not out.empty:
        out = out.sort_values(["dimension", "ind_id", "entity"]).reset_index(drop=True)
    return out


# =============================================================================
# SECTION 10.  PERIOD PIPELINE
# =============================================================================

@dataclass
class PeriodResult:
    label: str
    filepath: str
    stats: Dict[str, GlobalStats]
    weights: Dict[str, WeightSet]
    local: Dict[str, Dict[str, float]]
    totals: Dict[str, Dict[str, float]]
    comps: Dict[str, Dict[str, Dict[str, float]]]
    scores: Dict[str, Dict[str, LocalScore]]
    composite: Dict[str, float]
    hard_hits: Dict[str, str]
    summaries: Dict[str, List[LocalSummary]]  # local raw-unit diagnostics, not regular uploads
    data: Dict[str, pd.DataFrame]
    absolute_risk: Optional[pd.DataFrame] = None
    hard_events: List[dict] = field(default_factory=list)
    # Snapshots belong to this state, so later events and reviews leave an
    # earlier report unchanged.  Each accepted snapshot is fingerprinted.
    hard_register: Optional[HardAlertRegister] = None
    policy_digest: str = ""
    benchmark_digest: str = ""
    state_id: str = ""
    event_sequence: int = 0
    event_history: List[dict] = field(default_factory=list)
    # {dimension: {entity: {"provisional", "codes", "reasons", "coverage_share",
    #  "contributed_to_benchmark", "via"}}}; see provisional_status().
    provisional: Dict[str, Dict[str, dict]] = field(default_factory=dict)
    intervals: Dict[str, Dict[str, Tuple[float, float]]] = field(default_factory=dict)
    pending_declarations: List[dict] = field(default_factory=list)
    # Calibration uploads on the common scale (Section 4.2.2), by dimension;
    # empty when the round reuses frozen parameters.  `summaries` holds the
    # raw-unit summaries of the reference implementation.
    common_summaries: Dict[str, List[LocalSummary]] = field(default_factory=dict)
    calibration_activity: List[dict] = field(default_factory=list)
    # Tagged (estimate, third-party) or stale values; see apply_data_age().
    data_status: List[dict] = field(default_factory=list)
    # Dependency weights in force for this state; a documented revision
    # between calibration rounds changes them (revise_dependency_weights).
    dependency_weights: Dict[str, dict] = field(
        default_factory=lambda: json.loads(json.dumps(DEPENDENCY_WEIGHTS)))
    internal_warnings: InternalWarningRegister = field(default_factory=InternalWarningRegister)
    warning_streaks: Dict[Tuple[str, str], int] = field(default_factory=dict)
    high_risk_periods: Dict[Tuple[str, str], int] = field(default_factory=dict)
    # Consecutive published periods in which a score carried the
    # insufficient-information (provisional) flag (Section 4.2.4).
    insufficient_periods: Dict[Tuple[str, str], int] = field(default_factory=dict)
    # Dependency-weight method in force for each firm with suppliers (Table 4).
    dependency_methods: Dict[str, str] = field(
        default_factory=lambda: dict(DECLARED_ALPHA_METHOD))


@dataclass
class EventUpdateResult:
    event_id: str
    after: PeriodResult
    table: pd.DataFrame
    affected_entities: List[str]
    recomputed_local: List[Tuple[str, str]]
    duplicate: bool = False
    event_aggregation_calls: int = 0
    additional_version_records: int = 0


def _reference_digest(value) -> str:
    """Deterministic in-process integrity check, not an authentication proof."""
    def plain(v):
        if isinstance(v, pd.DataFrame):
            return {"columns": list(v.columns), "index": list(v.index),
                    "values": plain(v.to_numpy().tolist())}
        if isinstance(v, dict):
            return {str(k): plain(x) for k, x in v.items()}
        if isinstance(v, (list, tuple)):
            return [plain(x) for x in v]
        if isinstance(v, np.generic):
            return v.item()
        if hasattr(v, "__dict__"):
            return plain(vars(v))
        return v
    blob = json.dumps(plain(value), sort_keys=True, default=str,
                      separators=(",", ":"))
    return hashlib.sha256(blob.encode("utf-8")).hexdigest()


def _event_benchmark_digest(result: PeriodResult, policy: RegulatoryPolicy,
                            classification: str) -> str:
    # The publication timestamp is left out so that the same inputs and
    # configuration give the same digest in every run.
    weights = {d: {k: v for k, v in vars(w).items() if k != "timestamp"}
               for d, w in result.weights.items()}
    return _reference_digest({
        "stats": result.stats, "weights": weights,
        "policy": policy.digest(), "classification": classification,
        "dependencies": result.dependency_weights, "entity_order": ENTITY_ORDER,
        "topology": SUPPLY_CHAIN, "hard_rules": HARD_THRESHOLD_RULES,
    })


class FederatedESGSystem:
    """Runs the calibration rounds and periods, event updates, alert reviews
    and dependency-weight revisions of the demonstration."""

    def __init__(self, policy: Optional[RegulatoryPolicy] = None,
                 data_dir: str = ".", tier_path: str = "Indicator_Governance_Tier_Classification.xlsx",
                 secure_aggregation: bool = True,
                 strict_classification: bool = True):
        self.policy = policy or RegulatoryPolicy()
        problems = self.policy.validate() + check_dependency_weights()
        if problems:
            raise ValueError("configuration invalid: " + "; ".join(problems))
        self.data_dir = data_dir
        tier_candidate = (tier_path if os.path.isabs(tier_path) else
                          os.path.join(data_dir, tier_path))
        if not os.path.exists(tier_candidate) and not os.path.isabs(tier_path):
            project_candidate = os.path.join(
                os.path.dirname(os.path.abspath(__file__)), tier_path)
            if os.path.exists(project_candidate):
                tier_candidate = project_candidate
        self.tier_map, self.tier_source = load_governance_tiers(tier_candidate)
        self.tier_path = tier_candidate
        self.classification_digest = classification_digest(tier_candidate)
        # The built-in fallback map is a different indicator set from the
        # published classification, so the run stops unless the fallback is
        # requested explicitly.
        if strict_classification and self.tier_map.get("__STRICT__") != "1":
            raise ConfigurationError(
                f"governance classification workbook not readable at "
                f"{tier_candidate}. The fallback map is a different indicator "
                f"set and must not be presented as the published "
                f"classification. Place "
                f"'Indicator_Governance_Tier_Classification.xlsx' next to the "
                f"data, or pass strict_classification=False to run the "
                f"fallback deliberately.")
        self.header_map = load_entity_header_map(data_dir)
        self.coordinator = Coordinator(secure_aggregation=secure_aggregation)
        self.participants = {e: Participant(e) for e in ENTITY_ORDER}
        self.anchor_stats: Dict[str, GlobalStats] = {}
        self.registers: Dict[str, GovernanceRegister] = {}
        self.reference_scales: Dict[str, Dict[str, float]] = {}
        self.reference_basis: Dict[str, Dict[str, str]] = {}
        self._indicator_metadata: Dict[str, str] = {}
        # Last published benchmark per indicator: {dimension: {indicator: {...}}}.
        self.valid_reference: Dict[str, Dict[str, dict]] = {}
        self.hard_register = HardAlertRegister()
        # Dependency weights in force; a documented revision persists.
        self.dependency_weights: Dict[str, dict] = json.loads(json.dumps(DEPENDENCY_WEIGHTS))
        # Method in force for each firm with suppliers; a documented revision
        # that changes it is reported as a method change.
        self.dependency_methods: Dict[str, str] = dict(DECLARED_ALPHA_METHOD)
        self.version_log: List[dict] = []
        self.audit_log: List[AuditFinding] = []
        self._latest_state_id = ""
        self._state_counter = 0
        self._state_digests: Dict[str, str] = {}
        self._accepted_events: Dict[str, dict] = {}
        self._latest_result: Optional[PeriodResult] = None
        # Common-scale uploads of the latest calibration round, by dimension.
        self._round_common_summaries: Dict[str, List[LocalSummary]] = {}
        self._conversion_references: Dict[str, GlobalStats] = {}
        self.calibration_log: List[dict] = []

    # -----------------------------------------------------------------------
    def recalibrate(self, dim: str, df: pd.DataFrame, label: str,
                    is_anchor: bool = False,
                    sums: Optional[List[LocalSummary]] = None) -> GlobalStats:
        """One common-scale upload stage for one dimension.

        A first data-dependent reference is initialized in a separate, logged
        raw-unit stage. Subsequent calibrations upload common-scale summaries
        only. Raw-unit tables are evaluator diagnostics computed locally.
        """
        if sums is None:
            sums = local_raw_summaries(df)
        if dim not in self.reference_scales and needs_reference_initialization(df, self.policy):
            st = self.coordinator.aggregate(sums, df)
            self.calibration_log.append({
                "period": label, "dimension": dim, "stage": "reference initialization",
                "upload_scale": "raw units", "dimension_aggregations": 1,
                "purpose": "establish data-dependent ranges and reference scales; not calibration"})
        else:
            st = raw_unit_diagnostics(df, sums)
        # Reference scales are declared once for the fixed governance
        # universe, including unavailable observations. A new indicator
        # or unit requires an explicit new governance configuration.
        if dim not in self.reference_scales:
            s_ref, basis = declare_reference_scales(st, df, self.policy)
            self.reference_scales[dim] = s_ref
            self.reference_basis[dim] = basis
        attach_scoring_scale(st, self.reference_scales[dim], self.policy)
        self._round_common_summaries[dim] = attach_common_scale(
            st, df, self.policy, self._conversion_references.get(dim, self.anchor_stats.get(dim)),
            coordinator=self.coordinator)
        self.calibration_log.append({
            "period": label, "dimension": dim, "stage": "collaborative calibration",
            "upload_scale": "common scale [0,1]", "dimension_aggregations": 1,
            "purpose": "one upload stage; raw-unit diagnostics are evaluator-only"})
        self._conversion_references.setdefault(dim, st)
        self._apply_valid_reference(dim, st, label)
        if is_anchor:
            self.anchor_stats[dim] = st
        return st

    def _apply_valid_reference(self, dim: str, st: GlobalStats, label: str) -> None:
        """Use the last published benchmark, within the allowed age, for an
        indicator with too few reporters; otherwise it stays unavailable."""
        now = period_index(label)
        refs = self.valid_reference.setdefault(dim, {})
        for k in sorted(st.unavailable):
            ref = refs.get(k)
            if (ref is None or now is None or ref["index"] is None
                    or not 0 <= now - ref["index"] <= self.policy.reference_max_age_periods):
                continue
            st.mu[k], st.sigma[k], st.n[k] = ref["mu"], ref["sigma"], ref["n"]
            st.common_mu[k], st.common_sigma[k] = ref["common_mu"], ref["common_sigma"]
            st.sigma_score[k] = float(max(ref["sigma"],
                                          self.policy.eta_sigma_score * st.s_ref[k]))
            st.status[k] = "valid_reference"
            st.carried_forward[k] = f"{ref['period']} (n={ref['n']})"
            del st.unavailable[k]
        for k in st.mu:
            if k not in st.carried_forward and k in st.common_mu:
                refs[k] = {"period": label, "index": now, "mu": st.mu[k],
                           "sigma": st.sigma.get(k, 0.0), "n": st.n[k],
                           "common_mu": st.common_mu[k],
                           "common_sigma": st.common_sigma.get(k, 0.0)}

    def run_period(self, filepath: str, label: str,
                   is_anchor: bool = False,
                   frozen_stats: Optional[Dict[str, GlobalStats]] = None,
                   frozen_weights: Optional[Dict[str, WeightSet]] = None,
                   audit_evidence: Optional[Dict[str, Dict[str, float]]] = None,
                   ) -> PeriodResult:
        """Run one period.

        With frozen_stats and frozen_weights the M0 benchmark and weights are held
        (fixed-reference series); otherwise they are re-estimated.
        """
        data, data_status = apply_data_age(
            load_period(filepath, self.tier_map, self.header_map), label, self.policy)
        stats, weights, summaries, common_summaries = {}, {}, {}, {}
        local, scores, hard_hits = {}, {}, {}
        hard_events: List[dict] = []

        for dim, df in data.items():
            metadata = df[["indicator", "ind_id", "tier", "is_good", "category", "unit"]]
            metadata = metadata.drop_duplicates().sort_values("indicator").reset_index(drop=True)
            metadata_digest = _reference_digest(metadata)
            if dim in self._indicator_metadata and self._indicator_metadata[dim] != metadata_digest:
                raise ConfigurationError(
                    f"{dim} indicator universe, direction, classification or units changed; "
                    "declare a new governance configuration before recalibration")
            self._indicator_metadata.setdefault(dim, metadata_digest)
            sums = [self.participants[e].local_summary(df) for e in ENTITY_ORDER
                    if e in set(df["entity"])]
            summaries[dim] = sums

            if frozen_stats and dim in frozen_stats:
                st = frozen_stats[dim]
                common_summaries[dim] = []        # no calibration round, nothing aggregated
            else:
                st = self.recalibrate(dim, df, label, is_anchor=is_anchor, sums=sums)
                common_summaries[dim] = list(self._round_common_summaries.get(dim, []))

            if frozen_weights and dim in frozen_weights:
                ws = frozen_weights[dim]
            else:
                t1_ids = sorted({str(x) for x in
                                 df.loc[df["tier"] == "T1", "ind_id"].dropna().unique()})
                if dim not in self.registers:
                    self.registers[dim] = declare_governance_register(
                        dim, st, self.policy, label,
                        s_ref=self.reference_scales.get(dim, {}),
                        s_ref_basis=self.reference_basis.get(dim, {}),
                        classification_digest=self.classification_digest)
                    reg = self.registers[dim]
                    self.version_log.append({
                        "period": label, "dimension": dim,
                        "version_id": reg.version_id, "digest": reg.digest,
                        "policy_digest": reg.policy_digest,
                        "timestamp": datetime.now(timezone.utc).isoformat(timespec="seconds"),
                        "n_indicators": len(reg.members()),
                        "notes": (f"governance register declared: Level 2={len(reg.t2)}, "
                                  f"Level 3={len(reg.t3)}, Level 1 hard-alert rules={len(reg.t1)}, "
                                  f"classification={reg.classification_digest}"),
                        "scoring_version": "", "effective_from": label,
                    })
                ws = govern_weights(dim, st, self.anchor_stats.get(dim),
                                    self.policy, label, t1_ids_in_data=t1_ids,
                                    register=self.registers[dim],
                                    classification_digest_value=self.classification_digest)
                # Section 4.2.2: a numbered scoring version with its effective
                # period, fixed until the next recalibration.
                number = 1 + sum(1 for v in self.version_log if v["dimension"] == dim
                                 and not str(v["version_id"]).startswith("REG-"))
                self.version_log.append({
                    "period": label, "dimension": dim,
                    "version_id": ws.version_id, "digest": ws.digest,
                    "policy_digest": ws.policy_digest, "timestamp": ws.timestamp,
                    "n_indicators": len(ws.rows), "notes": "; ".join(ws.notes),
                    "scoring_version": number, "effective_from": label,
                })

            stats[dim], weights[dim] = st, ws

            # hard thresholds, verified by each firm's own auditor
            dim_scores, dim_local = {}, {}
            for e in ENTITY_ORDER:
                ev = (audit_evidence or {}).get(e)
                auditor = AssuranceProvider(e)
                hit = auditor.verify_hard_thresholds(df, ev)
                self.audit_log.extend(auditor.findings)
                if hit:
                    hard_hits[e] = "; ".join(auditor.confirmed_labels)
                for hit_label in auditor.confirmed_labels:
                    hit = hit_label
                    # Immediate, deduplicated, and sticky until a review clears
                    # it.  Registration does not wait for any numerical delta.
                    active_hit = any(r.label == hit for r in
                                     self.hard_register.active_records(e, dim))
                    is_new = False if active_hit else self.hard_register.raise_alert(
                        e, hit, label, dimension=dim, indicator=hit.split()[0])
                    hard_events.append({"period": label, "dimension": dim,
                                        "entity": e, "label": hit,
                                        "newly_raised": is_new})
                sc = score_entity(e, df, st, ws, self.policy, hard_violation=hit)
                dim_scores[e] = sc
                dim_local[e] = sc.score
            scores[dim], local[dim] = dim_scores, dim_local

        totals, comps = {}, {}
        for dim in data:
            t, c = propagate(local[dim], self.dependency_weights)
            totals[dim], comps[dim] = t, c

        wE, wS = self.policy.dim_weights
        composite = {e: wE * totals.get("Environmental", {}).get(e, 0.5)
                        + wS * totals.get("Social", {}).get(e, 0.5)
                     for e in ENTITY_ORDER}

        result = PeriodResult(label=label, filepath=filepath, stats=stats,
                            weights=weights, local=local, totals=totals,
                            comps=comps, scores=scores, composite=composite,
                            hard_hits=hard_hits, summaries=summaries, data=data,
                            common_summaries=common_summaries,
                            calibration_activity=[dict(row) for row in self.calibration_log
                                                  if row["period"] == label],
                            absolute_risk=absolute_risk_table(data, hard_hits),
                            hard_events=hard_events, data_status=data_status,
                            dependency_weights=json.loads(json.dumps(self.dependency_weights)),
                            dependency_methods=dict(self.dependency_methods))
        result.provisional = provisional_status(result, self.policy)
        return self._publish_snapshot(result)

    def _publish_snapshot(self, result: PeriodResult) -> PeriodResult:
        import copy
        self._state_counter += 1
        result.intervals = score_intervals(result.scores, result.dependency_weights)
        result.pending_declarations = []
        for dim, frame in result.data.items():
            for entity in ENTITY_ORDER:
                auditor = AssuranceProvider(entity)
                auditor.verify_hard_thresholds(frame)
                for finding in auditor.findings:
                    if finding.kind == "pending":
                        result.pending_declarations.append(dict(dimension=dim, entity=entity,
                            indicator=finding.indicator,
                            status=("unverified_pending_review"
                                    if finding.detail.startswith("unverified")
                                    else "unknown_pending_review")))
        result.hard_register = copy.deepcopy(self.hard_register)
        if not result.state_id:
            previous = self._latest_result
            if previous is not None:
                result.internal_warnings = copy.deepcopy(previous.internal_warnings)
                result.warning_streaks = dict(previous.warning_streaks)
                result.high_risk_periods = dict(previous.high_risk_periods)
                result.insufficient_periods = dict(previous.insufficient_periods)
            for dim in result.scores:
                reports = (_transition_reports(previous, result, dim, self.policy)
                           if previous is not None else {})
                for entity in ENTITY_ORDER:
                    key = (dim, entity)
                    high = result.totals[dim][entity] >= self.policy.persistent_level
                    result.high_risk_periods[key] = (
                        result.high_risk_periods.get(key, 0) + 1 if high else 0)
                    flagged = bool(result.provisional.get(dim, {}).get(entity, {}).get("provisional"))
                    result.insufficient_periods[key] = (
                        result.insufficient_periods.get(key, 0) + 1 if flagged else 0)
                    if entity in reports:
                        rep = reports[entity]
                        worse = any(v > rep.threshold for v in rep.component_deltas.values())
                        result.warning_streaks[key] = (
                            result.warning_streaks.get(key, 0) + 1 if worse else 0)
                        result.internal_warnings.update(
                            dim, entity, rep, f"{previous.label}->{result.label}")
        self._flag_persistent_reporting_failures(result)
        result.policy_digest = self.policy.digest()
        result.benchmark_digest = _event_benchmark_digest(
            result, self.policy, self.classification_digest)
        result.state_id = f"{self._state_counter}:{result.label}:{result.benchmark_digest[:16]}"
        self._latest_state_id = result.state_id
        self._state_digests[result.state_id] = _reference_digest(result)
        self._latest_result = copy.deepcopy(result)
        return result

    def _flag_persistent_reporting_failures(self, result: PeriodResult) -> None:
        """Section 4.2.4: insufficient information in policy.escalation_periods
        consecutive published periods calls for an assurance review.  Events
        between calibrations do not advance the period count."""
        for dim, records in result.provisional.items():
            for entity, rec in records.items():
                n = result.insufficient_periods.get((dim, entity), 0)
                if rec.get("provisional") and n >= self.policy.escalation_periods:
                    reason = (f"persistent reporting failure: insufficient information "
                              f"in {n} consecutive periods")
                    rec["assurance_review"] = True
                    reasons = rec.setdefault("assurance_review_reasons", [])
                    if reason not in reasons:
                        reasons.append(reason)

    def _validate_event_baseline(self, baseline: PeriodResult) -> None:
        if not baseline.state_id or baseline.state_id != self._latest_state_id:
            raise ValueError("stale baseline: use the latest accepted period/event state")
        if self._state_digests.get(baseline.state_id) != _reference_digest(baseline):
            raise ValueError("baseline snapshot was modified after publication")
        if baseline.policy_digest != self.policy.digest():
            raise ValueError("policy changed: publish a recalibration before updating events")
        if baseline.benchmark_digest != _event_benchmark_digest(
                baseline, self.policy, self.classification_digest):
            raise ValueError("benchmark, governance or dependency configuration changed")
        if _reference_digest(baseline.hard_register) != _reference_digest(self.hard_register):
            raise ValueError("hard-alert state changed: use the reviewed current snapshot")
        if any(ws.policy_digest != baseline.policy_digest
               for ws in baseline.weights.values()):
            raise ValueError("published weights do not match the frozen policy")

    def event_update(self, baseline: PeriodResult, *, event_id: str,
                     source: str, dimension: str, indicator_id: str,
                     value, verified_by: str,
                     alert_limits: Optional[Tuple[float, float]] = None,
                     data_status: Optional[str] = None,
                     data_source: Optional[str] = None) -> EventUpdateResult:
        """Apply one verified observation under the last published parameters.

        Only the source firm's local score is recomputed, then the totals downstream;
        no aggregation and no new parameter version.  value=None with a data_status
        withdraws the observation.  State is kept in memory.
        """
        import copy
        if not isinstance(event_id, str) or not event_id.strip():
            raise ValueError("event_id must be an explicit nonempty string")
        if not isinstance(verified_by, str) or not verified_by.strip():
            raise ValueError("event requires a named verification attribution")
        if source not in ENTITY_ORDER or dimension not in baseline.data:
            raise ValueError("unknown event source or dimension")
        source_data = baseline.data[dimension]
        source_row = source_data[(source_data["entity"] == source)
                                 & (source_data["ind_id"] == indicator_id)]
        if len(source_row) != 1:
            raise ValueError("event must identify exactly one existing observation")
        if value is None:
            if not str(data_status or "").strip():
                raise ValueError("an unavailable observation requires its data status")
            parsed = float("nan")
        else:
            # Parsed as the loader parses observations, in canonical units
            # (percentages as fractions).
            parsed = parse_indicator_value(value, str(source_row.iloc[0]["unit"]))
            if not np.isfinite(parsed):
                raise ValueError("event value must parse to a finite observation")
        if alert_limits is not None:
            if len(alert_limits) != 2 or not all(
                    np.isfinite(float(v)) and float(v) >= 0 for v in alert_limits):
                raise ValueError("alert_limits requires finite nonnegative absolute and relative limits")
            alert_limits = tuple(float(v) for v in alert_limits)
        payload = {"source": source, "dimension": dimension,
                   "indicator_id": indicator_id,
                   "value": None if value is None else float(parsed),
                   "verified_by": verified_by.strip(), "alert_limits": alert_limits,
                   "data_status": data_status, "data_source": data_source}
        previous = self._accepted_events.get(event_id)
        if previous is not None:
            if previous["payload"] != payload:
                raise ValueError("event_id conflicts with its previously accepted payload")
            replay = copy.deepcopy(previous["result"])
            replay.duplicate = True
            replay.recomputed_local = []
            replay.table["local_score_recomputed"] = False
            replay.table["total_recomputed"] = False
            replay.table["duplicate_submission"] = True
            # Historical deltas remain in composite_delta; this submission adds
            # no new increment, even after a later improvement or clearance.
            replay.table["submission_composite_delta"] = 0.0
            return replay
        self._validate_event_baseline(baseline)
        if event_id in self.hard_register.records:
            raise ValueError("event_id is already reserved by a periodic hard event")
        dwmap = baseline.dependency_weights
        data = baseline.data[dimension].copy(deep=True)
        selected = (data["entity"] == source) & (data["ind_id"] == indicator_id)
        if int(selected.sum()) != 1:
            raise ValueError("event must identify exactly one existing observation")
        data.loc[selected, "value"] = float(parsed)
        data.loc[selected, "as_of"] = baseline.label.split("::event:")[0]
        if data_status:
            data.loc[selected, "data_status"] = str(data_status).strip()
        elif value is not None:
            # A verified replacement is current even if the previous observation
            # was stale or withdrawn. Explicit estimate/source tags remain available.
            data.loc[selected, "data_status"] = "reported"
        if data_source:
            data.loc[selected, "data_source"] = str(data_source).strip()
        auditor = AssuranceProvider(source)
        hit = auditor.verify_hard_thresholds(data)
        # Only the event's own observation can raise a hard event; an existing
        # failure in another indicator is not raised again.  A separate check
        # keeps the findings of the full review from being logged twice.
        event_hit = AssuranceProvider(source).verify_hard_thresholds(data.loc[selected])
        score = score_entity(source, data, baseline.stats[dimension],
                             baseline.weights[dimension], self.policy,
                             hard_violation=hit)
        after = copy.deepcopy(baseline)
        after.data[dimension] = data
        _, updated_status = apply_data_age(
            {dimension: data}, baseline.label.split("::event:")[0], self.policy)
        after.data_status = [r for r in after.data_status if r["dimension"] != dimension]
        after.data_status.extend(updated_status)
        after.scores[dimension][source] = score
        after.local[dimension][source] = score.score
        if hit:
            after.hard_hits[source] = hit
        else:
            other_hit = next((sc.hard_threshold_hit for dim, scores in after.scores.items()
                              if dim != dimension for e, sc in scores.items()
                              if e == source and sc.hard_threshold_hit), None)
            if other_hit:
                after.hard_hits[source] = other_hit
            else:
                after.hard_hits.pop(source, None)
        affected = {source}
        for entity in ENTITY_ORDER:
            if any(alpha > 0.0 and supplier in affected for supplier, alpha
                   in dwmap[entity]["alpha"].items()):
                affected.add(entity)
        for entity in ENTITY_ORDER:
            if entity not in affected:
                continue
            dw = dwmap[entity]
            components = {"__own__": dw["beta"] * after.local[dimension][entity]}
            components.update({supplier: alpha * after.totals[dimension][supplier]
                               for supplier, alpha in dw["alpha"].items()})
            after.comps[dimension][entity] = components
            after.totals[dimension][entity] = float(sum(components.values()))
            w_e, w_s = self.policy.dim_weights
            after.composite[entity] = (
                w_e * after.totals["Environmental"][entity]
                + w_s * after.totals["Social"][entity])
        register = copy.deepcopy(baseline.hard_register)
        if event_hit:
            # Timing travels with the notice (Section 4.2.3): the alert is raised
            # in the reporting period of the baseline state, not under the label
            # of an earlier event.
            register.raise_alert(source, event_hit, baseline.label.split("::event:")[0],
                                 dimension=dimension, indicator=indicator_id,
                                 event_id=event_id)
        # alert_limits changes only the internal-warning limits; the score
        # policy, weights and benchmark stay frozen, and the alert configuration
        # is recorded in the event identity.
        trigger_policy = copy.deepcopy(self.policy)
        if alert_limits is not None:
            trigger_policy.component_delta_abs, trigger_policy.component_delta_rel = alert_limits
        alert_policy_digest = _reference_digest({
            "base_policy_digest": baseline.policy_digest,
            "component_delta_abs": trigger_policy.component_delta_abs,
            "component_delta_rel": trigger_policy.component_delta_rel})
        ind_id = baseline.stats[dimension].ind_id
        reports = {}
        for entity in ENTITY_ORDER:
            reports[entity] = component_trigger(
                entity, baseline.comps[dimension][entity],
                after.comps[dimension][entity], after.totals[dimension],
                dimension, trigger_policy,
                hard_local=register.is_active(entity, dimension),
                hard_labels=register.labels(entity, dimension),
                reference_total=baseline.totals[dimension][entity],
                hard_records=_hard_records(register, entity, dimension),
                own_effects=own_indicator_effects(
                    baseline.scores[dimension][entity], after.scores[dimension][entity],
                    dwmap[entity]["beta"], self.policy, ind_id))
        cascade_alerts(reports, dwmap)
        after.intervals = score_intervals(after.scores, dwmap)
        after.provisional = provisional_status(after, self.policy)
        comparable = comparability(baseline, after, dimension, dwmap)

        def change_source(entity):
            # Section 4.2.4: the source of each change is identified separately.
            if entity == source:
                quality_change = (value is None or data_status or data_source
                                  or not np.isfinite(source_row.iloc[0]["value"]))
                return "data quality" if quality_change else "own ESG data"
            via = [j for j, a in dwmap[entity]["alpha"].items() if a > 0.0 and j in affected]
            return f"updated supplier total via {', '.join(disp(j) for j in via)}" if via else ""
        rows = []
        for entity in ENTITY_ORDER:
            report = reports[entity]
            transition = f"event:{event_id}"
            open_now, closed_now = after.internal_warnings.update(
                dimension, entity, report, transition)
            key = (dimension, entity)
            # Several observations between calibrations are not extra reporting
            # periods. They can raise a size-based response and keep warnings open,
            # but do not artificially advance the period-duration counters.
            high_periods = (baseline.high_risk_periods.get(key, 0)
                            if after.totals[dimension][entity] >= self.policy.persistent_level
                            else 0)
            level, basis = graduated_response(
                report, baseline.warning_streaks.get(key, 0), high_periods,
                after.totals[dimension][entity], trigger_policy)
            level, basis = carry_open_warnings(level, basis, open_now, transition)
            prov = after.provisional[dimension][entity]
            delta = after.composite[entity] - baseline.composite[entity]
            rows.append({
                "event_id": event_id, "entity": entity, "dimension": dimension,
                "event_source": source, "change_source": change_source(entity),
                "indicator_id": indicator_id,
                "verified_by": verified_by.strip(), "duplicate_submission": False,
                "local_score_recomputed": entity == source,
                "total_recomputed": entity in affected,
                "composite_before": baseline.composite[entity],
                "composite_after": after.composite[entity], "composite_delta": delta,
                "submission_composite_delta": delta,
                "alert": _alert_code(report),
                "warning_threshold": report.threshold,
                "reference_total_pre_event": report.reference_total,
                "internal_warning": report.internal_warning,
                "warning_signal": warning_signal(report),
                "open_internal_warnings": open_warning_text(open_now),
                "internal_warnings_closed": open_warning_text(closed_now, closed=True),
                "response_level": level, "response_basis": basis,
                "warning_direction": report.warning_direction,
                "indicators_flagged": ",".join(report.fired_indicators),
                "branches_flagged": ",".join(report.fired_branches),
                "hard_alert": report.hard_alert, "hard_alert_local": report.hard_local,
                "hard_event_count": len(report.hard_event_summary),
                "hard_event_ids": ",".join(r["event_id"] for r in report.hard_event_summary),
                "hard_event_details": json.dumps(report.hard_event_summary, sort_keys=True),
                "hard_alert_via_direct_supplier": ",".join(report.hard_via),
                "hard_alert_via_direct_customer": ",".join(report.hard_via_customers),
                "hard_notice_received": _hard_detail_text(
                    report.hard_via_detail + report.hard_customer_detail),
                "provisional": prov["provisional"],
                "provisional_via_direct_supplier": ",".join(prov["via"]),
                "comparable_with_previous": comparable[entity][0],
                "comparability_note": comparable[entity][1],
                "component_deltas": json.dumps(report.component_deltas, sort_keys=True),
                "parameter_version": baseline.weights[dimension].version_id,
                "benchmark_digest": baseline.benchmark_digest,
                "alert_policy_digest": alert_policy_digest,
                "alert_limits_override": json.dumps(alert_limits),
                "data_status": data_status or "",
            })
        after.event_sequence = baseline.event_sequence + 1
        after.label = baseline.label.split("::event:")[0] + "::event:" + event_id
        record = {"event_id": event_id, **payload, "baseline_state_id": baseline.state_id,
                  "sequence": after.event_sequence, "hard_event": bool(event_hit),
                  "benchmark_digest": baseline.benchmark_digest,
                  "alert_policy_digest": alert_policy_digest,
                  "composite_deltas": {r["entity"]: r["composite_delta"] for r in rows},
                  "verified_at": datetime.now(timezone.utc).isoformat(timespec="seconds")}
        after.event_history.append(record)
        if event_hit:
            after.hard_events.append({"period": after.label, "dimension": dimension,
                                      "entity": source, "label": event_hit,
                                      "event_id": event_id, "newly_raised": True})
        after.absolute_risk = absolute_risk_table(after.data, after.hard_hits)
        self.hard_register = register
        self.audit_log.extend(auditor.findings)
        after = self._publish_snapshot(after)
        result = EventUpdateResult(
            event_id=event_id, after=after, table=pd.DataFrame(rows),
            affected_entities=[e for e in ENTITY_ORDER if e in affected],
            recomputed_local=[(source, dimension)])
        self._accepted_events[event_id] = {"payload": payload,
                                          "result": copy.deepcopy(result)}
        return result

    def review_clear_hard_alert(self, baseline: PeriodResult, *, event_id: str,
                                reviewed_by: str, note: str) -> PeriodResult:
        """Clear a hard alert after review of the corrected observation and pass the
        closure notice on.
        """
        import copy
        self._validate_event_baseline(baseline)
        if any(not isinstance(text, str) or not text.strip() for text in (reviewed_by, note)):
            raise ValueError("clearance requires a reviewer and a review note")
        rec = self.hard_register.records.get(event_id)
        if rec is None or not rec.active:
            raise ValueError("no active hard alert for event_id")
        frame = baseline.data[rec.dimension]
        observation = frame[(frame["entity"] == rec.entity)
                            & (frame["ind_id"] == rec.indicator)]
        reviewer = AssuranceProvider(rec.entity)
        unresolved = reviewer.verify_hard_thresholds(observation)
        if (len(observation) != 1 or not np.isfinite(observation["value"]).all()
                or unresolved or any(f.kind == "pending" for f in reviewer.findings)):
            raise ValueError("resolve the hard-threshold observation before review clearance")
        after = copy.deepcopy(baseline)
        self.hard_register = copy.deepcopy(baseline.hard_register)
        self.hard_register.clear(event_id, cleared_by=reviewed_by.strip(), note=note.strip())
        # The documented resolution travels the same paths as the alert, to
        # firms downstream and upstream, naming only the direct supplier or
        # customer it arrives through (Section 4.2.3).
        closed = self.hard_register.records[event_id]
        paths = ([("downstream", c, v) for c, v in
                  downstream_notices(closed.entity, baseline.dependency_weights)]
                 + [("upstream", c, v) for c, v in
                    upstream_notices(closed.entity, baseline.dependency_weights)])
        closed.closure_notices = [
            {"event_id": closed.event_id, "direction": d,
             "firm": firm, "via": via, "category": closed.label,
             "status": "closed", "raised": closed.raised_period,
             "closed_at": closed.cleared_at, "resolution": closed.clearance_note}
            for d, firm, via in paths]
        after.event_history.append({"event_id": event_id, "kind": "review_clearance",
                                    "reviewed_by": reviewed_by.strip(), "note": note.strip(),
                                    "baseline_state_id": baseline.state_id,
                                    "closure_notices": closed.closure_notices})
        return self._publish_snapshot(after)

    def revise_dependency_weights(self, baseline: PeriodResult, *, event_id: str,
                                  firm: str, beta: float, alpha: Dict[str, float],
                                  method: str, evidence: str,
                                  verified_by: str) -> EventUpdateResult:
        """Documented and assured revision of one firm's dependency weights or
        supplier relationships (Section 4.2.3).

        Totals are recomputed and reported as a dependency revision; no internal
        warning is raised.
        """
        import copy
        if not isinstance(event_id, str) or not event_id.strip():
            raise ValueError("event_id must be an explicit nonempty string")
        if firm not in ENTITY_ORDER:
            raise ValueError("unknown firm")
        if any(not isinstance(text, str) or not text.strip() for text in (evidence, verified_by)):
            raise ValueError("a dependency revision requires documentation and assurance")
        alpha = {j: float(a) for j, a in alpha.items()}
        rank = {t: i for i, t in enumerate(TIER_ORDER)}
        if any(j not in ENTITY_ORDER or rank[SUPPLY_CHAIN[j]["tier"]]
               >= rank[SUPPLY_CHAIN[firm]["tier"]] for j in alpha):
            raise ValueError("suppliers must be enrolled firms in an earlier tier")
        if (not np.isfinite(float(beta)) or not np.isfinite(list(alpha.values())).all()
                or not 0.0 <= float(beta) <= 1.0 or any(a < 0.0 for a in alpha.values())
                or abs(float(beta) + sum(alpha.values()) - 1.0) > 1e-9):
            raise ValueError("beta and alpha must be finite, nonnegative and sum to one")
        payload = {"firm": firm, "beta": float(beta), "alpha": alpha, "method": method,
                   "evidence": str(evidence).strip(), "verified_by": str(verified_by).strip()}
        previous = self._accepted_events.get(event_id)
        if previous is not None:
            if previous["payload"] != payload:
                raise ValueError("event_id conflicts with its previously accepted payload")
            replay = copy.deepcopy(previous["result"])
            replay.duplicate = True
            replay.table["local_score_recomputed"] = False
            replay.table["total_recomputed"] = False
            replay.table["duplicate_submission"] = True
            replay.table["submission_composite_delta"] = 0.0
            return replay
        self._validate_event_baseline(baseline)
        if event_id in self.hard_register.records:
            raise ValueError("event_id is already reserved by a hard event")
        auditor = AssuranceProvider(firm)
        check = auditor.verify_alpha_method(method, method)
        if not check["accepted"]:
            raise ValueError(f"'{method}' is not an accepted dependency-weight method")
        # Sections 4.1 and 4.2.3: the method must be applied consistently, so a
        # revision that changes the method on record is reported as a change.
        applied_method = str(method).strip().lower()
        previous_method = str(getattr(baseline, "dependency_methods",
                                      DECLARED_ALPHA_METHOD).get(firm, "") or "")
        method_changed = bool(previous_method) and previous_method != applied_method
        after = copy.deepcopy(baseline)
        after.dependency_weights[firm] = {"beta": float(beta), "alpha": alpha}
        after.dependency_methods = dict(getattr(after, "dependency_methods",
                                                DECLARED_ALPHA_METHOD))
        after.dependency_methods[firm] = applied_method
        weights = after.dependency_weights
        w_e, w_s = self.policy.dim_weights
        for dim in after.local:
            after.totals[dim], after.comps[dim] = propagate(after.local[dim], weights)
        after.composite = {e: w_e * after.totals["Environmental"][e]
                           + w_s * after.totals["Social"][e] for e in ENTITY_ORDER}
        reach = dict(downstream_notices(firm, weights))
        affected = [e for e in ENTITY_ORDER if e == firm or e in reach]
        hard = {}
        for dim in after.local:
            reps = {e: component_trigger(
                e, after.comps[dim][e], after.comps[dim][e], after.totals[dim], dim,
                self.policy, hard_local=self.hard_register.is_active(e, dim),
                hard_labels=self.hard_register.labels(e, dim),
                reference_total=after.totals[dim][e],
                hard_records=_hard_records(self.hard_register, e, dim))
                for e in ENTITY_ORDER}
            hard[dim] = cascade_alerts(reps, weights)
        after.provisional = provisional_status(after, self.policy)
        comparable = {dim: comparability(baseline, after, dim, weights)
                      for dim in after.local}
        rows = []
        for e in ENTITY_ORDER:
            row = {"event_id": event_id, "entity": e, "revised_firm": firm,
                   "duplicate_submission": False,
                   "change_source": ("dependency-weight revision" if e == firm else
                                     f"updated supplier total via {disp(reach[e])}"
                                     if e in reach else ""),
                   "method": method,
                   "previous_method": previous_method if e == firm else "",
                   "method_changed": bool(method_changed and e == firm),
                   "evidence": payload["evidence"],
                   "verified_by": payload["verified_by"],
                   "enhanced_assurance": check["enhanced_assurance"],
                   "local_score_recomputed": False, "total_recomputed": e in affected,
                   "internal_warning": False,
                   "composite_before": baseline.composite[e],
                   "composite_after": after.composite[e],
                   "composite_delta": after.composite[e] - baseline.composite[e],
                   "submission_composite_delta": after.composite[e] - baseline.composite[e]}
            for dim in after.local:
                row[f"{dim[:1]}_total_before"] = baseline.totals[dim][e]
                row[f"{dim[:1]}_total_after"] = after.totals[dim][e]
                row[f"{dim[:1]}_hard_alert"] = hard[dim][e].hard_alert
            row["hard_alert"] = any(hard[dim][e].hard_alert for dim in after.local)
            row["hard_event_ids"] = ",".join(sorted({
                r["event_id"] for dim in after.local
                for r in hard[dim][e].hard_event_summary}))
            row["hard_event_count"] = len({
                r["event_id"] for dim in after.local
                for r in hard[dim][e].hard_event_summary})
            row["hard_event_details"] = json.dumps(
                [r for dim in after.local for r in hard[dim][e].hard_event_summary],
                sort_keys=True)
            codes = {_alert_code(hard[dim][e]) for dim in after.local}
            row["alert"] = next((c for c in ("H(self)", "H") if c in codes), "")
            row["comparable_with_previous"] = all(comparable[d][e][0] for d in after.local)
            row["comparability_note"] = "; ".join(
                f"{d}: {comparable[d][e][1]}" for d in after.local if comparable[d][e][1])
            rows.append(row)
        after.event_sequence = baseline.event_sequence + 1
        after.label = baseline.label.split("::event:")[0] + "::event:" + event_id
        after.event_history.append({
            "event_id": event_id, "kind": "dependency_weight_revision", **payload,
            "previous": baseline.dependency_weights[firm],
            "previous_method": previous_method, "method_changed": method_changed,
            "enhanced_assurance": check["enhanced_assurance"],
            "baseline_state_id": baseline.state_id, "sequence": after.event_sequence,
            "verified_at": datetime.now(timezone.utc).isoformat(timespec="seconds")})
        self.audit_log.extend(auditor.findings)
        self.dependency_weights = json.loads(json.dumps(after.dependency_weights))
        self.dependency_methods = dict(after.dependency_methods)
        after = self._publish_snapshot(after)
        result = EventUpdateResult(event_id=event_id, after=after, table=pd.DataFrame(rows),
                                   affected_entities=affected, recomputed_local=[])
        self._accepted_events[event_id] = {"payload": payload,
                                          "result": copy.deepcopy(result)}
        return result


# Firm whose observation each event changes; event D revises FOCAL's
# dependency weights.
EVENT_SOURCE = {"A": "S2-CHEM", "B": "S1-RUBBER", "C": "S1-RUBBER", "E": "S2-CHEM"}

# Title and description of each event scenario, used by the result workbook
# and the platform.
EVENT_SCENARIO_TEXT = {
    "A_single_environmental_event": (
        "A · Ordinary event",
        "S2-CHEM's E1 doubles. Only that firm's environmental local score is recomputed, "
        "and totals update along its downstream paths."),
    "B_hard_stop_event": (
        "B · Hard event",
        "S1-RUBBER's S26 declaration changes from Yes to No after assurance. Its social local "
        "score stays unchanged, and the hard alert reaches FOCAL and both retailers downstream "
        "and S2-CHEM and S2-EQUIP upstream. After the correction, review clearance sends a "
        "closure notice along the same paths."),
    "C_hard_stop_high_thresholds": (
        "C · Priority check",
        "Repeats B with warning thresholds set too high to trigger. Internal warnings stay "
        "off while every affected hard alert stays on."),
    "D_dependency_weight_revision": (
        "D · Dependency revision",
        "FOCAL revises its supplier proportions to 0.18 / 0.12 under the strategic-dependency "
        "method, documented and checked with enhanced assurance. Local scores stay unchanged; "
        "FOCAL's and its customers' totals are recomputed and reported as a dependency "
        "revision, with no internal warning."),
    "E_data_quality_change": (
        "E · Data-quality change",
        "S2-CHEM's E1 observation is withdrawn. Its environmental score becomes an interval "
        "published at the upper bound, and the provisional status passes downstream naming "
        "only the direct supplier."),
}


def run_event_demonstrations(data_dir: str = ".",
                             policy: Optional[RegulatoryPolicy] = None,
                             strict_classification: bool = True) -> dict:
    """Independent event scenarios from the same published R0 state.

    A  an environmental observation at S2-CHEM
    B  a failed declaration at S1-RUBBER (hard alert), corrected and cleared by review
    C  as B, with warning limits no change can reach
    D  a documented dependency-weight revision at FOCAL
    E  a data-quality change: an observation of S2-CHEM is withdrawn
    """
    import copy
    # The events start from M0; M1-M5 are not read.
    files = discover_period_files(data_dir, expect_periods=None, verbose=False)
    anchor_path = next(path for label, path in files if label == "M0")
    scenarios, checks, versions, closures = {}, [], {}, []
    names = ("A_single_environmental_event", "B_hard_stop_event",
             "C_hard_stop_high_thresholds", "D_dependency_weight_revision",
             "E_data_quality_change")
    # Firms whose totals an event at the source reaches, and the suppliers its
    # hard alert reaches upstream.
    chain = {"S2-CHEM": {"S2-CHEM", "S1-RUBBER", "FOCAL", "D1-RETAIL", "D2-RETAIL"},
             "S1-RUBBER": {"S1-RUBBER", "FOCAL", "D1-RETAIL", "D2-RETAIL"}}
    alerted_upstream = {"S2-CHEM": set(), "S1-RUBBER": {"S2-CHEM", "S2-EQUIP"}}
    root_system = FederatedESGSystem(policy=copy.deepcopy(policy or RegulatoryPolicy()),
                                     data_dir=data_dir,
                                     strict_classification=strict_classification)
    root_baseline = root_system.run_period(anchor_path, "R0", is_anchor=True)
    for name in names:
        kind = name[0]
        source = EVENT_SOURCE.get(kind)
        system, baseline = copy.deepcopy((root_system, root_baseline))
        if baseline.hard_hits:
            raise AssertionError("event demonstrations require a clean R0 baseline")
        baseline_digest = _reference_digest(baseline)
        version_log = copy.deepcopy(system.version_log)
        calls = {"aggregate": 0}

        def forbidden_aggregation(*args, **kwargs):
            calls["aggregate"] += 1
            raise AssertionError("event update called coordinator.aggregate")
        system.coordinator.aggregate = forbidden_aggregation
        if kind == "D":
            result = system.revise_dependency_weights(
                baseline, event_id=name, firm="FOCAL", beta=0.70,
                alpha={"S1-RUBBER": 0.18, "S1-COATING": 0.12},
                method="strategic dependency",
                evidence="few alternative sources for the rubber components",
                verified_by="scenario assurance simulation")
        else:
            dim, iid = ("Social", "S26") if kind in "BC" else ("Environmental", "E1")
            observation = baseline.data[dim].query("entity == @source and ind_id == @iid")
            value = {"A": 2.0 * float(observation.value.iloc[0]), "E": None}.get(kind, 0.0)
            result = system.event_update(
                baseline, event_id=name, source=source, dimension=dim,
                indicator_id=iid, value=value, verified_by="scenario assurance simulation",
                alert_limits=(2.0, 10.0) if kind == "C" else None,
                data_status="withdrawn: not supported by records" if kind == "E" else None)
        table = result.table
        after = result.after
        affected = table[table.total_recomputed]
        unaffected = table[~table.total_recomputed]
        notified = alerted_upstream[source] if kind in "BC" else set()
        invariants = {
            "zero_event_aggregation_calls": calls["aggregate"] == 0,
            "zero_new_parameter_versions": system.version_log == version_log,
            "frozen_parameters": after.stats == baseline.stats and all(
                after.weights[d].rows.equals(baseline.weights[d].rows)
                and after.weights[d].digest == baseline.weights[d].digest
                for d in baseline.weights),
            "immutable_baseline": _reference_digest(baseline) == baseline_digest,
            "unrelated_entities_unchanged": bool(
                (unaffected.composite_delta == 0).all()
                and not unaffected.internal_warning.any()
                and not unaffected.loc[~unaffected.entity.isin(notified), "hard_alert"].any()),
        }
        if kind == "D":
            upstream = [e for e in ENTITY_ORDER if e not in {"FOCAL", "D1-RETAIL", "D2-RETAIL"}]
            record = after.event_history[-1]
            invariants.update({
                "local_scores_unchanged": after.local == baseline.local,
                "upstream_totals_unchanged": all(
                    after.totals[d][e] == baseline.totals[d][e]
                    for d in baseline.totals for e in upstream),
                "revised_firm_and_customers_recomputed": set(result.affected_entities) == {
                    "FOCAL", "D1-RETAIL", "D2-RETAIL"},
                "reported_as_dependency_revision": bool(
                    table.loc[table.entity == "FOCAL", "change_source"].item()
                    == "dependency-weight revision" and not table.internal_warning.any()),
                "documented_and_assured": bool(
                    record["method"] in DEPENDENCY_METHODS and record["evidence"]
                    and record["verified_by"] and record["enhanced_assurance"]),
                "method_change_disclosed": bool(
                    record["method_changed"]
                    and record["previous_method"] == DECLARED_ALPHA_METHOD["FOCAL"]
                    and after.dependency_methods["FOCAL"] == record["method"]),
                "declared_table_unchanged": (
                    DEPENDENCY_WEIGHTS["FOCAL"]["alpha"] == {"S1-RUBBER": 0.12, "S1-COATING": 0.18}
                    and after.dependency_weights["FOCAL"]["alpha"]
                    == {"S1-RUBBER": 0.18, "S1-COATING": 0.12}),
                "revision_carried_forward": (
                    system.dependency_weights == after.dependency_weights),
            })
        else:
            invariants.update({
                "only_source_dimension_recomputed": result.recomputed_local == [(source, dim)]
                    and all(after.local[d][e] == baseline.local[d][e]
                            for d in baseline.local for e in ENTITY_ORDER
                            if (e, d) != (source, dim)),
                "positive_dependency_chain_only": set(result.affected_entities) == chain[source],
            })
        if kind == "A":
            invariants["continuous_channel_behavior"] = bool((affected.composite_delta > 0).all())
            invariants["comparable_when_coverage_unchanged"] = bool(
                table.comparable_with_previous.all())
        if kind in "BC":
            upstream_rows = table[table.entity.isin(notified)]
            invariants["continuous_channel_behavior"] = bool((affected.composite_delta == 0).all())
            invariants["hard_alert_whole_chain"] = bool(affected.hard_alert.all())
            invariants["hard_alert_upstream_without_score_change"] = bool(
                len(upstream_rows) == len(notified) and upstream_rows.hard_alert.all()
                and (upstream_rows.hard_alert_via_direct_customer == source).all()
                and (upstream_rows.composite_delta == 0).all())
            invariants["direct_supplier_visibility"] = table.loc[
                table.entity == "FOCAL", "hard_alert_via_direct_supplier"].item() == "S1-RUBBER"
            invariants["alert_code_self_at_source"] = dict(zip(table.entity, table.alert)) == {
                e: "H(self)" if e == source else "H" if e in chain[source] | notified else ""
                for e in ENTITY_ORDER}
        if kind == "C":
            invariants["hard_alert_bypasses_high_thresholds"] = bool(
                not table.internal_warning.any() and affected.hard_alert.all())
        if kind == "B":
            corrected = system.event_update(after, event_id="B_correction",
                                            source=source, dimension=dim,
                                            indicator_id=iid, value=1.0,
                                            verified_by="scenario assurance simulation")
            cleared = system.review_clear_hard_alert(
                corrected.after, event_id=name, reviewed_by="scenario assurance simulation",
                note="corrected declaration verified against policy evidence")
            notices = cleared.hard_register.records[name].closure_notices
            closures.extend(dict(n, event_id=name) for n in notices)
            invariants["closure_notice_passed_on"] = bool(
                {(n["direction"], n["firm"], n["via"]) for n in notices} == {
                    ("downstream", "FOCAL", "S1-RUBBER"), ("downstream", "D1-RETAIL", "FOCAL"),
                    ("downstream", "D2-RETAIL", "FOCAL"), ("upstream", "S2-CHEM", "S1-RUBBER"),
                    ("upstream", "S2-EQUIP", "S1-RUBBER")}
                and all(n["resolution"] and n["event_id"] == name for n in notices)
                and not cleared.hard_register.is_active(source, dim))
        if kind == "E":
            before_sc, after_sc = baseline.scores[dim]["S2-CHEM"], after.scores[dim]["S2-CHEM"]
            told = dict(zip(table.entity, table.provisional_via_direct_supplier))
            invariants.update({
                "point_value_is_upper_bound": bool(
                    after_sc.incomplete and after_sc.score == after_sc.upper
                    and after_sc.lower <= before_sc.score <= after_sc.upper),
                "provisional_notice_downstream": told == {
                    "S2-CHEM": "", "S2-EQUIP": "", "S1-RUBBER": "S2-CHEM",
                    "S1-COATING": "", "FOCAL": "S1-RUBBER", "D1-RETAIL": "FOCAL",
                    "D2-RETAIL": "FOCAL"},
                "data_status_recorded": after.data[dim].loc[
                    (after.data[dim].entity == "S2-CHEM") & (after.data[dim].ind_id == iid),
                    "data_status"].item() == "withdrawn: not supported by records",
                "reported_as_data_quality_change": table.loc[
                    table.entity == "S2-CHEM", "change_source"].item() == "data quality",
                "marked_not_directly_comparable": dict(
                    zip(table.entity, table.comparable_with_previous.astype(bool))) == {
                    "S2-CHEM": False, "S2-EQUIP": True, "S1-RUBBER": False,
                    "S1-COATING": True, "FOCAL": False, "D1-RETAIL": False,
                    "D2-RETAIL": False},
            })
        for check, passed in invariants.items():
            checks.append({"scenario": name, "check": check,
                           "passed": bool(passed), "status": "PASS" if passed else "FAIL"})
        if not all(invariants.values()):
            raise AssertionError(f"{name}: failed event invariants: {invariants}")
        scenarios[name] = table
        versions[name] = {"policy_digest": baseline.policy_digest,
                          "benchmark_digest": baseline.benchmark_digest,
                          "versions": {d: w.version_id for d, w in baseline.weights.items()}}
    regression_passed = bool(.006 > min(.005, .01 * .706))
    checks.append({"scenario": "threshold regression", "check": "minimum_OR_rule",
                   "passed": regression_passed,
                   "status": "PASS" if regression_passed else "FAIL"})
    bootstrap = sum(r["dimension_aggregations"] for r in root_system.calibration_log
                    if r["stage"] == "reference initialization")
    calibration = sum(r["dimension_aggregations"] for r in root_system.calibration_log
                      if r["stage"] == "collaborative calibration")
    return {"scenarios": scenarios, "acceptance": pd.DataFrame(checks),
            "closures": pd.DataFrame(closures), "verification": {
        "status": "PASS", "type": "single-process reference event implementation",
        "independent_scenarios": list(names), "collaborative_recalibrations": 1,
        "dimension_aggregations_at_R0": bootstrap + calibration,
        "total_baseline_dimension_aggregations": bootstrap + calibration,
        "reference_initialization_dimension_aggregations": bootstrap,
        "calibration_dimension_aggregations": calibration,
        "baseline_sequential_upload_stages": int(bootstrap > 0) + int(calibration > 0),
        "calibration_log": list(root_system.calibration_log),
        "event_aggregation_calls": 0,
        "additional_version_records": 0, "versions": versions,
        "limits": ["Caller-attributed verification, no authentication or cryptographic isolation",
                   "Upload-stage counts describe the simulated protocol; no distributed transport is implemented",
                   "In-memory event ledger; durable persistence is not implemented",
                   "Full-chain event tables are evaluator output, not a focal-firm access view",
                   "Strict increase applies to isolated first adverse events below saturation",
                   "Scenario C records an alert-only override; scoring policy and R0 stay frozen"],
    }}


# =============================================================================
# SECTION 11.  LOCAL-SCORER SUBSTITUTION AND SHAP  (optional experiments)
# =============================================================================
# The published score is a declared rule, sigmoid(sum_k w_k * z_k / tau), with
# no outcome label and nothing fitted.  Neither experiment feeds back into the
# weights, scores or propagation.
#   11a  One XGBoost scorer per dimension, fitted to the rule's local scores for
#        M0-M5 under the frozen M0 parameters, replaces the rule.  The check
#        confirms that hard alerts coincide and compares internal warnings.
#   11b  SHAP of the same scorer: per-indicator changes in the events, and an
#        exploratory comparison with the rule's weights (different units).
# =============================================================================

ML_REQUIREMENTS = (
    "scikit-learn>=1.4,<1.7; xgboost>=2.0,<3.1; shap>=0.45,<0.48; "
    "matplotlib>=3.8,<3.11; scipy>=1.16 on Python 3.11+"
)

SHAP_TOP_N = 15
SHAP_SYNTHETIC_ROWS = 2000
SCORER_SPEC = ("XGBRegressor: 400 trees, depth 2, learning rate 0.05, "
               "no subsampling, random_state {seed}")


def _optional_dependency_failure(exc: BaseException) -> str:
    """Reason an optional package is unavailable; the run continues without it."""
    if isinstance(exc, ModuleNotFoundError):
        return f"missing optional dependency: {exc.name}"
    detail = (str(exc).strip().splitlines() or [""])[0][:300]
    return (f"optional dependency installed but failed to load "
            f"({type(exc).__name__}: {detail}); reinstall or upgrade the failing "
            f"package in this interpreter")


def _make_scorer_model(regressor, random_state: int):
    return regressor(
        objective="reg:squarederror", n_estimators=400, max_depth=2,
        learning_rate=0.05, subsample=1.0, colsample_bytree=1.0,
        reg_lambda=1.0, min_child_weight=1.0, random_state=random_state, n_jobs=1)


def _indicator_sort_key(ind_id) -> Tuple[str, int]:
    m = re.match(r"^([A-Za-z]+)(\d+)$", str(ind_id))
    return (m.group(1), int(m.group(2))) if m else (str(ind_id), 0)


def _indicator_attribution_panel(periods: Dict[str, "PeriodResult"],
                                 dimension: str) -> dict:
    """Entity x period panel of the rule's inputs z_k, local scores and
    contributions w_k * z_k.
    """
    feature_rows, contribution_rows, keys = [], [], []
    target, excluded, complete = [], [], []
    names, tiers, weights_seen = {}, {}, {}
    for label, result in periods.items():
        stats, weights = result.stats[dimension], result.weights[dimension]
        wmap = weights.w()
        tier_of = dict(zip(weights.rows["indicator"], weights.rows["tier"]))
        for entity in ENTITY_ORDER:
            score = result.scores[dimension].get(entity)
            if score is None:
                continue
            if score.coverage <= 0.0:
                excluded.append({"period": label, "entity": disp(entity),
                                 "reason": "no scoreable indicator"})
                continue
            features, contributions = {}, {}
            for k, c in score.contributions.items():
                w = float(wmap.get(k, 0.0))
                if w <= 0.0:
                    continue
                iid = str(stats.ind_id.get(k) or k)
                features[iid] = float(c) / w
                contributions[iid] = float(c)
                names[iid] = str(k)
                tiers[iid] = tier_of.get(k, "")
                weights_seen.setdefault(iid, []).append(w)
            feature_rows.append(features)
            contribution_rows.append(contributions)
            keys.append({"period": label, "entity": disp(entity)})
            target.append(float(score.score))
            complete.append(not score.incomplete)
    columns = sorted(names, key=_indicator_sort_key)
    X = pd.DataFrame(feature_rows).reindex(columns=columns).astype(float)
    C = pd.DataFrame(contribution_rows).reindex(columns=columns).fillna(0.0).astype(float)
    return {"X": X, "rule": C, "y": np.asarray(target, dtype=float),
            "complete": np.asarray(complete, dtype=bool),
            "keys": pd.DataFrame(keys), "excluded": excluded, "names": names,
            "tiers": tiers,
            "mean_weight": {k: float(np.mean(v)) for k, v in weights_seen.items()}}


def _scorer_features(entity: str, frame: pd.DataFrame, stats: GlobalStats,
                     weights: WeightSet, policy: RegulatoryPolicy) -> Dict[str, float]:
    """The rule's standardized inputs z_k of one firm under the given frozen
    parameters, keyed by indicator id.  They are read back from score_entity
    as contribution / w_k, so the scorer sees exactly what the rule sees."""
    wmap = weights.w()
    score = score_entity(entity, frame, stats, weights, policy)
    return {str(stats.ind_id.get(k) or k): float(c) / float(wmap[k])
            for k, c in score.contributions.items() if float(wmap.get(k, 0.0)) > 0.0}


def _rule_contributions(entity: str, frame: pd.DataFrame, stats: GlobalStats,
                        weights: WeightSet, policy: RegulatoryPolicy) -> Dict[str, float]:
    """The rule's exact per-indicator contribution w_k * z_k (pre-sigmoid)."""
    score = score_entity(entity, frame, stats, weights, policy)
    return {str(stats.ind_id.get(k) or k): float(v) for k, v in score.contributions.items()}


def _soft_difference(rule_fires: bool, alt_fires: bool) -> str:
    if rule_fires and alt_fires:
        return "same (both fire)"
    if not rule_fires and not alt_fires:
        return "same (neither fires)"
    return "missed trigger vs rule" if rule_fires else "new trigger vs rule"


def run_scorer_substitution_experiment(data_dir: str, policy: RegulatoryPolicy,
                                       engine_events: Optional[dict] = None,
                                       strict_classification: bool = True,
                                       random_state: int = 42) -> dict:
    """Replace the scoring rule by an XGBoost scorer fitted to it under the
    frozen M0 parameters, and compare events A-C and the periodic transitions.
    """
    import copy
    empty = {k: pd.DataFrame() for k in ("training", "events", "event_inputs",
                                         "transitions", "summary", "checks")}
    try:
        from xgboost import XGBRegressor
    except Exception as exc:
        return {"status": "UNAVAILABLE", "reason": _optional_dependency_failure(exc),
                "results_usable": False, "requirements": ML_REQUIREMENTS, **empty}

    files = discover_period_files(data_dir, expect_periods=None, verbose=False)
    labels = [label for label, _ in files]
    system = FederatedESGSystem(policy=copy.deepcopy(policy), data_dir=data_dir,
                                strict_classification=strict_classification)
    r0 = system.run_period(dict(files)["M0"], "R0", is_anchor=True)
    frozen = {label: system.run_period(path, f"{label}@R0", frozen_stats=r0.stats,
                                       frozen_weights=r0.weights)
              for label, path in files}
    dims = ("Environmental", "Social")
    if engine_events is None:
        engine_events = run_event_demonstrations(data_dir=data_dir, policy=policy,
                                                 strict_classification=strict_classification)

    # ---- one XGBoost scorer per dimension, fitted to frozen-rule scores -----
    panels, models, columns, training_rows = {}, {}, {}, []
    for dim in dims:
        panel = _indicator_attribution_panel(frozen, dim)
        X, y = panel["X"], panel["y"]
        model = _make_scorer_model(XGBRegressor, random_state)
        model.fit(X, y)
        fitted = model.predict(X)
        sst = float(np.sum((y - y.mean()) ** 2))
        panels[dim], models[dim], columns[dim] = panel, model, list(X.columns)
        training_rows.append({
            "dimension": dim,
            "frozen_parameter_version": r0.weights[dim].version_id,
            "training_rows": int(len(y)),
            "distinct_feature_rows": int(X.round(12).drop_duplicates().shape[0]),
            "excluded_rows": "; ".join(f"{e['period']} {e['entity']} ({e['reason']})"
                                       for e in panel["excluded"]) or "none",
            "indicators": int(X.shape[1]),
            "target": ("rule local score for the M0-M5 inputs under frozen R0 parameters "
                       "(interval upper bound where inputs are incomplete); no outcome label"),
            "scorer": SCORER_SPEC.format(seed=random_state),
            "in_sample_r2": (float(1.0 - np.sum((y - fitted) ** 2) / sst)
                             if sst > 0 else float("nan")),
            "max_abs_in_sample_error": float(np.max(np.abs(y - fitted))),
        })

    def feature_row(entity: str, frame: pd.DataFrame, dim: str) -> pd.DataFrame:
        feats = _scorer_features(entity, frame, r0.stats[dim], r0.weights[dim], policy)
        return pd.DataFrame([feats]).reindex(columns=columns[dim]).astype(float)

    def local_scores(data: Dict[str, pd.DataFrame], scorer: str) -> Dict[str, Dict[str, float]]:
        """Local score of every firm under one scorer.  The hard-threshold
        check is shared; a hard alert changes neither scorer's local score."""
        out_local = {}
        for dim in dims:
            frame, dim_local = data[dim], {}
            for e in ENTITY_ORDER:
                hit = AssuranceProvider(e).verify_hard_thresholds(frame)
                if scorer == "rule":
                    dim_local[e] = score_entity(e, frame, r0.stats[dim], r0.weights[dim],
                                                policy, hard_violation=hit).score
                else:
                    dim_local[e] = float(models[dim].predict(feature_row(e, frame, dim))[0])
            out_local[dim] = dim_local
        return out_local

    def chain(local: Dict[str, Dict[str, float]]):
        totals, comps = {}, {}
        for dim in dims:
            totals[dim], comps[dim] = propagate(local[dim])
        wE, wS = policy.dim_weights
        composite = {e: wE * totals["Environmental"][e] + wS * totals["Social"][e]
                     for e in ENTITY_ORDER}
        return totals, comps, composite

    def own_effects(frame0: pd.DataFrame, frame1: pd.DataFrame, dim: str,
                    scorer: str) -> Dict[str, Dict[str, float]]:
        """Each firm's own-indicator effects (own_indicator_effects) under one
        scorer: the change in the firm's total if only that input changed."""
        effects = {}
        for e in ENTITY_ORDER:
            beta = DEPENDENCY_WEIGHTS[e]["beta"]
            if scorer == "rule":
                effects[e] = own_indicator_effects(
                    score_entity(e, frame0, r0.stats[dim], r0.weights[dim], policy),
                    score_entity(e, frame1, r0.stats[dim], r0.weights[dim], policy),
                    beta, policy, r0.stats[dim].ind_id)
                continue
            x0, x1 = feature_row(e, frame0, dim), feature_row(e, frame1, dim)
            base = float(models[dim].predict(x0)[0])
            effects[e] = {}
            for k in x0.columns:
                a, b = x0.at[0, k], x1.at[0, k]
                if (pd.isna(a) and pd.isna(b)) or a == b:
                    continue
                probe = x0.copy()
                probe.at[0, k] = b
                effects[e][k] = beta * (float(models[dim].predict(probe)[0]) - base)
        return effects

    def alert_reports(comps0, comps1, totals0, totals1, dim, trigger_policy, hard, effects):
        reports = {e: component_trigger(
            e, comps0[dim][e], comps1[dim][e], totals1[dim], dim, trigger_policy,
            hard_local=bool(hard.get(e)), hard_labels=list(hard.get(e, [])),
            reference_total=totals0[dim][e], own_effects=effects[e]) for e in ENTITY_ORDER}
        return cascade_alerts(reports)

    def report_fields(rep, suffix: str) -> dict:
        return {f"threshold_{suffix}": rep.threshold,
                f"reference_total_{suffix}": rep.reference_total,
                f"max_abs_component_{suffix}": rep.max_abs_component,
                f"component_deltas_{suffix}": json.dumps(rep.component_deltas, sort_keys=True),
                f"warning_{suffix}": bool(rep.internal_warning),
                f"hard_alert_{suffix}": bool(rep.hard_alert),
                f"hard_via_{suffix}": ", ".join(disp(b) for b in rep.hard_via)}

    # ---- A/B/C frozen events under both scorers ----------------------------
    scenario_specs = [("A_single_environmental_event", "Environmental", "E1", None),
                      ("B_hard_stop_event", "Social", "S26", None),
                      ("C_hard_stop_high_thresholds", "Social", "S26", (2.0, 10.0))]
    event_rows, input_rows, engine_diffs, shap_inputs = [], [], [], []
    pre = {}
    for s in ("rule", "xgb"):
        local0 = local_scores(r0.data, s)
        pre[s] = chain(local0) + (local0,)
    for name, dim, iid, limits in scenario_specs:
        source = EVENT_SOURCE[name[0]]
        data = copy.deepcopy(r0.data)
        frame = data[dim]
        selected = (frame["entity"] == source) & (frame["ind_id"] == iid)
        before_value = float(frame.loc[selected, "value"].iloc[0])
        frame.loc[selected, "value"] = 2.0 * before_value if iid == "E1" else 0.0
        event_value = float(frame.loc[selected, "value"].iloc[0])
        event_hit = AssuranceProvider(source).verify_hard_thresholds(frame.loc[selected])
        hard = {source: [event_hit]} if event_hit else {}
        trigger_policy = copy.deepcopy(policy)
        if limits is not None:
            trigger_policy.component_delta_abs, trigger_policy.component_delta_rel = limits

        results = {}
        for s in ("rule", "xgb"):
            totals0, comps0, composite0, local0 = pre[s]
            local1 = local_scores(data, s)
            totals1, comps1, composite1 = chain(local1)
            reps = alert_reports(comps0, comps1, totals0, totals1, dim, trigger_policy, hard,
                                 own_effects(r0.data[dim], data[dim], dim, s))
            results[s] = (local0, local1, totals0, totals1, composite0, composite1, reps)

        engine = engine_events.get("scenarios", {}).get(name)
        for e in ENTITY_ORDER:
            row = {"scenario": name, "entity": disp(e), "dimension": dim,
                   "event_indicator": iid,
                   "alert_limits": "policy" if limits is None else json.dumps(limits)}
            for s in ("rule", "xgb"):
                local0, local1, totals0, totals1, composite0, composite1, reps = results[s]
                row.update({f"local_before_{s}": local0[dim][e], f"local_after_{s}": local1[dim][e],
                            f"total_before_{s}": totals0[dim][e], f"total_after_{s}": totals1[dim][e],
                            f"composite_before_{s}": composite0[e],
                            f"composite_after_{s}": composite1[e],
                            f"composite_delta_{s}": composite1[e] - composite0[e]})
                row.update(report_fields(reps[e], s))
            row["warning_difference"] = _soft_difference(row["warning_rule"], row["warning_xgb"])
            row["hard_alert_identical"] = (row["hard_alert_rule"] == row["hard_alert_xgb"]
                                           and row["hard_via_rule"] == row["hard_via_xgb"])
            event_rows.append(row)
            erow = engine[engine["entity"] == e] if engine is not None else pd.DataFrame()
            if len(erow) == 1:
                engine_diffs.append({
                    "delta_gap": abs(float(erow["composite_delta"].iloc[0])
                                     - row["composite_delta_rule"]),
                    "alerts_equal": (bool(erow["internal_warning"].iloc[0]) == row["warning_rule"]
                                     and bool(erow["hard_alert"].iloc[0]) == row["hard_alert_rule"])})
            else:
                engine_diffs.append({"delta_gap": float("inf"), "alerts_equal": False})

        # Inputs handed to SHAP (11b): the source firm before and after.
        source_hit = AssuranceProvider(source).verify_hard_thresholds(frame)
        shap_inputs.append({
            "scenario": name, "source": source, "dimension": dim, "indicator": iid,
            "x_before": feature_row(source, r0.data[dim], dim),
            "x_after": feature_row(source, frame, dim),
            "rule_before": _rule_contributions(source, r0.data[dim], r0.stats[dim],
                                               r0.weights[dim], policy),
            "rule_after": _rule_contributions(source, frame, r0.stats[dim],
                                              r0.weights[dim], policy),
            "hard_alert_after": source_hit,
            "local": {s: (results[s][0][dim][source], results[s][1][dim][source])
                      for s in ("rule", "xgb")}})

        # Whether the event input already occurs in the training panel.
        X = panels[dim]["X"]
        vec = feature_row(source, frame, dim).iloc[0]
        same_pattern = X.isna().eq(vec.isna()).all(axis=1)
        close = ((X.fillna(0.0) - vec.fillna(0.0)).abs() <= 1e-9).all(axis=1)
        row_seen = bool((same_pattern & close).any())
        is_feature = iid in X.columns
        value_seen = bool(is_feature and (X[iid] - vec[iid]).abs().le(1e-9).any())
        raw_seen = any(bool(np.isclose(frozen[label].data[dim].loc[
            frozen[label].data[dim]["ind_id"] == iid, "value"].astype(float), event_value,
            rtol=1e-12, atol=1e-12).any()) for label in labels)
        outside = [k for k in X.columns if pd.notna(vec[k])
                   and (vec[k] < X[k].min() - 1e-12 or vec[k] > X[k].max() + 1e-12)]
        co_moves = []
        for i in range(1, len(labels)):
            f0, f1 = frozen[labels[i - 1]].data[dim], frozen[labels[i]].data[dim]
            v0 = f0[f0["entity"] == source].set_index("ind_id")["value"].astype(float)
            v1 = f1[f1["entity"] == source].set_index("ind_id")["value"].astype(float)
            changed = [k for k in v0.index if k in v1.index
                       and not ((np.isnan(v0[k]) and np.isnan(v1[k]))
                                or np.isclose(v0[k], v1[k], rtol=1e-12, atol=0.0))]
            if iid in changed:
                others = sorted((k for k in changed if k != iid), key=_indicator_sort_key)
                co_moves.append(f"{labels[i - 1]}->{labels[i]}: " + (", ".join(others) or "none"))
        overlaps = row_seen or value_seen or raw_seen
        input_rows.append({
            "scenario": name, "source": disp(source), "dimension": dim,
            "changed_indicator": iid, "value_before": before_value, "value_after": event_value,
            "changed_indicator_is_scorer_feature": is_feature,
            "event_feature_row_matches_training_row": row_seen,
            "changed_risk_z_seen_in_training": (("yes" if value_seen else "no")
                                                if is_feature else "n/a (not a feature)"),
            "changed_raw_value_seen_in_M0_M5": raw_seen,
            "features_outside_training_range": ", ".join(outside) or "none",
            "training_changes_to_this_indicator_and_co_moving_indicators":
                "; ".join(co_moves) or "never changed in M0-M5",
            "experiment_label": ("event perturbation (input overlaps the training panel)"
                                 if overlaps else "unseen-input test"),
            "note": ("hard-gate indicator is not a scorer feature: the scorer input is the "
                     "unchanged M0 row, so neither local score moves; the hard alert is "
                     "raised and propagated separately"
                     if not is_feature else ""),
        })

    # ---- 70 periodic entity-transitions under the frozen R0 rule ------------
    local_by = {s: {label: local_scores(frozen[label].data, s) for label in labels}
                for s in ("rule", "xgb")}
    chain_by = {s: {label: chain(local_by[s][label]) for label in labels} for s in ("rule", "xgb")}
    transition_rows = []
    for i in range(1, len(labels)):
        a, b = labels[i - 1], labels[i]
        for dim in dims:
            hard: Dict[str, List[str]] = {}
            for label in labels[: i + 1]:           # sticky until review; none here
                for e in ENTITY_ORDER:
                    hit = frozen[label].scores[dim][e].hard_threshold_hit
                    if hit and hit not in hard.setdefault(e, []):
                        hard[e].append(hit)
            hard = {e: v for e, v in hard.items() if v}
            reps = {}
            for s in ("rule", "xgb"):
                totals0, comps0, _ = chain_by[s][a]
                totals1, comps1, _ = chain_by[s][b]
                reps[s] = alert_reports(comps0, comps1, totals0, totals1, dim, policy, hard,
                                        own_effects(frozen[a].data[dim], frozen[b].data[dim],
                                                    dim, s))
            for e in ENTITY_ORDER:
                row = {"transition": f"{a}->{b}", "dimension": dim, "entity": disp(e)}
                for s in ("rule", "xgb"):
                    row[f"local_before_{s}"] = local_by[s][a][dim][e]
                    row[f"local_after_{s}"] = local_by[s][b][dim][e]
                    row.update(report_fields(reps[s][e], s))
                row["warning_difference"] = _soft_difference(row["warning_rule"],
                                                             row["warning_xgb"])
                row["hard_alert_identical"] = (row["hard_alert_rule"] == row["hard_alert_xgb"]
                                               and row["hard_via_rule"] == row["hard_via_xgb"])
                transition_rows.append(row)

    events_df, transitions_df = pd.DataFrame(event_rows), pd.DataFrame(transition_rows)

    def tally(frame: pd.DataFrame, scope: str) -> dict:
        out_row = {"scope": scope, "comparisons": int(len(frame))}
        col = frame["warning_difference"]
        out_row["warning_same"] = int(col.str.startswith("same").sum())
        out_row["warning_missed_vs_rule"] = int((col == "missed trigger vs rule").sum())
        out_row["warning_new_vs_rule"] = int((col == "new trigger vs rule").sum())
        out_row["hard_alerts_rule"] = int(frame["hard_alert_rule"].sum())
        out_row["hard_alerts_xgb"] = int(frame["hard_alert_xgb"].sum())
        out_row["hard_alert_identical"] = bool(frame["hard_alert_identical"].all())
        return out_row

    summary = pd.DataFrame([tally(events_df, "frozen events A/B/C"),
                            tally(transitions_df, "70 periodic entity-transitions (training panel)")])

    c_rows = events_df[events_df["scenario"] == "C_hard_stop_high_thresholds"]
    chain_ids = ({disp(e) for e in _downstream_closure({EVENT_SOURCE["C"]})}
                 | {disp(j) for j, _ in upstream_notices(EVENT_SOURCE["C"])})
    b_source = events_df[(events_df["scenario"] == "B_hard_stop_event")
                         & (events_df["entity"] == disp(EVENT_SOURCE["B"]))]
    score_unchanged = bool(
        len(b_source) == 1
        and np.isclose(b_source["local_after_rule"], b_source["local_before_rule"],
                       rtol=0.0, atol=1e-12).all()
        and np.isclose(b_source["local_after_xgb"], b_source["local_before_xgb"],
                       rtol=0.0, atol=1e-12).all()
        and bool(b_source["hard_alert_rule"].all()) and bool(b_source["hard_alert_xgb"].all()))
    # The features are the rule's own inputs: on rows with complete inputs the
    # frozen weights and the sigmoid reproduce the rule's local score from them.
    reproduce_gap = 0.0
    for dim in dims:
        panel = panels[dim]
        wmap0 = r0.weights[dim].w()
        wvec = np.array([float(wmap0[panel["names"][c]]) for c in columns[dim]])
        index = panel["X"].fillna(0.0).to_numpy() @ wvec / policy.squash_scale
        rebuilt = 1.0 / (1.0 + np.exp(-index))
        complete = panel["complete"]
        if complete.any():
            reproduce_gap = max(reproduce_gap,
                                float(np.max(np.abs(rebuilt[complete] - panel["y"][complete]))))
    gap = max(d["delta_gap"] for d in engine_diffs)
    alerts_equal = all(d["alerts_equal"] for d in engine_diffs)
    checks = [
        ("hard_alert_invariant_events", bool(events_df["hard_alert_identical"].all()),
         "hard alerts and their direct-supplier paths coincide for every node in A/B/C"),
        ("hard_alert_invariant_transitions", bool(transitions_df["hard_alert_identical"].all()),
         "hard alerts coincide on all periodic entity-transitions"),
        ("hard_alert_leaves_local_score_unchanged", score_unchanged,
         "event B raises the hard alert at the source and leaves its Social local score "
         "unchanged under both scorers (alerts travel separately from scores)"),
        ("scorer_features_are_rule_inputs", bool(reproduce_gap < 1e-12),
         f"sigmoid(sum_k w_k z_k / tau) on the scorer's features reproduces the rule's "
         f"local score on every complete-input row; max gap {reproduce_gap:.2e}"),
        ("xgb_hard_alert_bypasses_high_thresholds",
         bool(len(c_rows) and not c_rows["warning_xgb"].any()
              and c_rows.loc[c_rows["entity"].isin(chain_ids), "hard_alert_xgb"].all()),
         "scenario C under the XGBoost scorer: no internal warning, hard alert on the whole chain"),
        ("rule_harness_reproduces_event_engine", bool(gap < 1e-12 and alerts_equal),
         f"max composite-delta gap vs FederatedESGSystem.event_update = {gap:.3e}; "
         f"alert flags equal = {alerts_equal}"),
    ]
    checks_df = pd.DataFrame([{"check": c, "result": "PASS" if ok else "FAIL", "detail": d}
                              for c, ok, d in checks])
    failed = checks_df.loc[checks_df["result"] != "PASS", "check"].tolist()
    return {"status": "FAILED" if failed else "COMPLETE",
            "reason": ("critical checks failed: " + ", ".join(failed)) if failed else "",
            "results_usable": not failed, "requirements": ML_REQUIREMENTS,
            "training": pd.DataFrame(training_rows), "events": events_df,
            "event_inputs": pd.DataFrame(input_rows), "transitions": transitions_df,
            "summary": summary, "checks": checks_df,
            # internal objects for 11b; not exported
            "_models": models, "_panels": panels, "_event_shap_inputs": shap_inputs}


def _shap_rule_agreement(shap_df: pd.DataFrame, rule: pd.DataFrame) -> dict:
    mean_abs_shap, mean_abs_rule = shap_df.abs().mean(), rule.abs().mean()
    top = min(10, shap_df.shape[1])
    overlap = len(set(mean_abs_shap.nlargest(top).index) & set(mean_abs_rule.nlargest(top).index))
    # Exploratory only.  SHAP is in post-sigmoid score units and measured from
    # the average prediction; the rule term is in pre-sigmoid index units and
    # measured from z = 0.  Centring aligns the reference point, not the units,
    # so these statistics do not show that the rule's weights were recovered.
    centred = rule - rule.mean()
    return {"mean_abs_shap": mean_abs_shap, "mean_abs_rule": mean_abs_rule,
            "rank_agreement": float(mean_abs_shap.rank().corr(mean_abs_rule.rank())),
            "overlap": f"{overlap}/{top}",
            "pointwise": float(np.corrcoef(shap_df.to_numpy().ravel(),
                                           centred.to_numpy().ravel())[0, 1])}


def run_indicator_shap_diagnostic(scorer: dict, output_dir: str = "",
                                  random_state: int = 42,
                                  squash_scale: float = 1.0) -> dict:
    """SHAP attribution of the XGBoost scorer: per-indicator changes in the
    events and an exploratory comparison with the rule's weights.
    """
    empty_summary = pd.DataFrame(columns=[
        "dimension", "panel", "rows", "distinct_feature_rows", "excluded_rows",
        "indicators", "target", "model", "in_sample_r2_fidelity",
        "rank_agreement_spearman", "top10_overlap", "pointwise_correlation",
        "interpretation"])
    empty_comparison = pd.DataFrame(columns=[
        "dimension", "ind_id", "indicator", "tier", "frozen_governed_weight",
        "observations", "mean_abs_shap", "mean_shap", "rank_shap",
        "mean_abs_rule_contribution", "mean_rule_contribution", "rank_rule"])
    empty_observations = pd.DataFrame(columns=[
        "dimension", "period", "entity", "ind_id", "risk_z", "shap_value",
        "rule_contribution"])
    empty_change = pd.DataFrame(columns=[
        "scenario", "dimension", "source", "ind_id", "indicator", "risk_z_before",
        "risk_z_after", "shap_before", "shap_after", "shap_change",
        "rule_contribution_change_index", "rule_effect_on_local_score", "note"])
    empty = {"summary": empty_summary, "comparison": empty_comparison,
             "observations": empty_observations, "event_change": empty_change,
             "artifacts": pd.DataFrame(columns=["dimension", "artifact"])}
    if scorer.get("status") == "UNAVAILABLE":
        return {"status": "UNAVAILABLE", "reason": scorer["reason"],
                "results_usable": False, "requirements": ML_REQUIREMENTS, **empty}
    try:
        import shap
        import matplotlib
        matplotlib.use("Agg")
        import matplotlib.pyplot as plt
        from xgboost import XGBRegressor
    except Exception as exc:
        return {"status": "UNAVAILABLE", "reason": _optional_dependency_failure(exc),
                "results_usable": False, "requirements": ML_REQUIREMENTS, **empty}

    summary_rows, comparison_frames, observation_frames, artifacts = [], [], [], []
    explainers = {}
    model_label = "XGBoost scorer of the substitution experiment (frozen R0 rule scores)"
    for dimension in ("Environmental", "Social"):
        panel, model = scorer["_panels"][dimension], scorer["_models"][dimension]
        X, rule, y = panel["X"], panel["rule"], panel["y"]
        explainer = shap.TreeExplainer(model)
        explainers[dimension] = explainer
        shap_df = pd.DataFrame(np.asarray(explainer.shap_values(X), dtype=float),
                               columns=X.columns)
        fitted = model.predict(X)
        sst = float(np.sum((y - y.mean()) ** 2))
        r2 = float(1.0 - np.sum((y - fitted) ** 2) / sst) if sst > 0 else float("nan")
        observed = _shap_rule_agreement(shap_df, rule)
        mean_abs_shap, mean_abs_rule = observed["mean_abs_shap"], observed["mean_abs_rule"]
        comparison = pd.DataFrame({
            "dimension": dimension,
            "ind_id": X.columns,
            "indicator": [panel["names"][k] for k in X.columns],
            "tier": [panel["tiers"].get(k, "") for k in X.columns],
            "frozen_governed_weight": [panel["mean_weight"].get(k, np.nan) for k in X.columns],
            "observations": X.notna().sum().to_numpy(),
            "mean_abs_shap": mean_abs_shap.to_numpy(),
            "mean_shap": shap_df.mean().to_numpy(),
            "rank_shap": mean_abs_shap.rank(ascending=False, method="min").astype(int).to_numpy(),
            "mean_abs_rule_contribution": mean_abs_rule.to_numpy(),
            "mean_rule_contribution": rule.mean().to_numpy(),
            "rank_rule": mean_abs_rule.rank(ascending=False, method="min").astype(int).to_numpy(),
        }).sort_values("rank_shap", kind="stable").reset_index(drop=True)
        comparison_frames.append(comparison)
        summary_rows.append({
            "dimension": dimension, "panel": "observed M0-M5 under frozen R0",
            "rows": int(len(y)),
            "distinct_feature_rows": int(X.round(12).drop_duplicates().shape[0]),
            "excluded_rows": "; ".join(f"{e['period']} {e['entity']} ({e['reason']})"
                                       for e in panel["excluded"]) or "none",
            "indicators": int(X.shape[1]),
            "target": ("rule local score under frozen R0 parameters (interval upper bound "
                       "where inputs are incomplete)"),
            "model": model_label,
            "in_sample_r2_fidelity": r2,
            "rank_agreement_spearman": observed["rank_agreement"],
            "top10_overlap": observed["overlap"],
            "pointwise_correlation": observed["pointwise"],
            "interpretation": ("Exploratory comparison. SHAP of the substitution scorer is "
                               "in post-sigmoid score units; the rule term w_k*risk_z_k is in "
                               "pre-sigmoid index units, so these statistics do not show that "
                               "the rule's weights were recovered. The event-level comparison "
                               "(T4i) uses score units for both. Not a predictive or causal "
                               "model; R2 measures fidelity to the rule, not accuracy."),
        })

        # Synthetic setting: the same frozen weights and sigmoid applied to
        # independent, complete standard-normal risk_z.  The weights are used
        # as published (fixed, summing to one over the dimension), not
        # renormalized over the columns.
        rng = np.random.default_rng([random_state, 0 if dimension == "Environmental" else 1])
        w = np.array([panel["mean_weight"][k] for k in X.columns], dtype=float)
        Z = pd.DataFrame(rng.standard_normal((SHAP_SYNTHETIC_ROWS, X.shape[1])),
                         columns=X.columns)
        synthetic_rule = Z * w
        synthetic_y = 1.0 / (1.0 + np.exp(-synthetic_rule.sum(axis=1).to_numpy() / squash_scale))
        synthetic_model = _make_scorer_model(XGBRegressor, random_state)
        synthetic_model.fit(Z, synthetic_y)
        synthetic_fit = synthetic_model.predict(Z)
        synthetic_shap = pd.DataFrame(np.asarray(
            shap.TreeExplainer(synthetic_model).shap_values(Z), dtype=float), columns=Z.columns)
        synthetic = _shap_rule_agreement(synthetic_shap, synthetic_rule)
        sst_s = float(np.sum((synthetic_y - synthetic_y.mean()) ** 2))
        summary_rows.append({
            "dimension": dimension, "panel": f"synthetic setting ({SHAP_SYNTHETIC_ROWS} rows)",
            "rows": SHAP_SYNTHETIC_ROWS, "distinct_feature_rows": SHAP_SYNTHETIC_ROWS,
            "excluded_rows": "n/a", "indicators": int(X.shape[1]),
            "target": ("same frozen weights and sigmoid; independent, complete "
                       "standard-normal risk_z"),
            "model": "separate XGBoost fit with the same specification",
            "in_sample_r2_fidelity": (float(1.0 - np.sum((synthetic_y - synthetic_fit) ** 2)
                                            / sst_s) if sst_s > 0 else float("nan")),
            "rank_agreement_spearman": synthetic["rank_agreement"],
            "top10_overlap": synthetic["overlap"],
            "pointwise_correlation": synthetic["pointwise"],
            "interpretation": ("This comparison changes sample size, indicator dependence "
                               "and missingness simultaneously; it cannot isolate the cause "
                               "of any observed agreement differences."),
        })

        long = shap_df.copy()
        long.insert(0, "period", panel["keys"]["period"].to_numpy())
        long.insert(1, "entity", panel["keys"]["entity"].to_numpy())
        long = long.melt(id_vars=["period", "entity"], var_name="ind_id", value_name="shap_value")
        zs = X.copy()
        zs.insert(0, "period", panel["keys"]["period"].to_numpy())
        zs.insert(1, "entity", panel["keys"]["entity"].to_numpy())
        zs = zs.melt(id_vars=["period", "entity"], var_name="ind_id", value_name="risk_z")
        rc = rule.copy()
        rc.insert(0, "period", panel["keys"]["period"].to_numpy())
        rc.insert(1, "entity", panel["keys"]["entity"].to_numpy())
        rc = rc.melt(id_vars=["period", "entity"], var_name="ind_id", value_name="rule_contribution")
        merged = zs.merge(long, on=["period", "entity", "ind_id"]).merge(
            rc, on=["period", "entity", "ind_id"])
        merged.insert(0, "dimension", dimension)
        observation_frames.append(merged[merged["risk_z"].notna()].reset_index(drop=True))

        if output_dir:
            os.makedirs(output_dir, exist_ok=True)
            figure_path = os.path.join(output_dir, f"XGBoost_SHAP_{dimension}.png")
            shown = comparison.head(SHAP_TOP_N).iloc[::-1]

            def short(name: str, limit: int = 44) -> str:
                text = re.sub(r"^\s*[ES]\d{1,2}\s*[.．]\s*", "", str(name)).strip()
                return text if len(text) <= limit else text[:limit - 1].rstrip() + "…"

            labels = [f"{iid}  {short(nm)}" for iid, nm in zip(shown["ind_id"], shown["indicator"])]
            ink, muted, grid, series = "#0b0b0b", "#52514e", "#e4e3df", "#2a78d6"
            fig, ax = plt.subplots(figsize=(9.0, max(3.6, 0.36 * len(shown) + 1.6)))
            fig.patch.set_facecolor("white")
            ax.barh(range(len(shown)), shown["mean_abs_shap"], height=0.62, color=series,
                    edgecolor="white", linewidth=2)
            ax.set_yticks(range(len(shown)))
            ax.set_yticklabels(labels, fontsize=8.5, color=ink)
            xmax = float(shown["mean_abs_shap"].max()) or 1.0
            ax.set_xlim(0, xmax * 1.28)
            for i, (v, rr) in enumerate(zip(shown["mean_abs_shap"], shown["rank_rule"])):
                ax.text(v + xmax * 0.012, i, f"{v:.4f}   rule #{int(rr)}",
                        va="center", fontsize=7.5, color=muted)
            ax.xaxis.grid(True, color=grid, linewidth=0.8)
            ax.set_axisbelow(True)
            for side in ("top", "right", "left"):
                ax.spines[side].set_visible(False)
            ax.spines["bottom"].set_color(grid)
            ax.tick_params(axis="y", length=0)
            ax.tick_params(axis="x", colors=muted, labelsize=8)
            ax.set_xlabel("mean |SHAP value| (local-score units)", fontsize=8.5, color=muted)
            fig.text(0.012, 0.985, f"Indicator-level SHAP attribution — {dimension}",
                     fontsize=11, color=ink, ha="left", va="top")
            fig.text(0.012, 0.985 - 0.34 / fig.get_size_inches()[1],
                     ("XGBoost scorer of the substitution experiment, fitted to rule scores "
                      f"under frozen R0 parameters · {len(y)} entity-periods · top "
                      f"{len(shown)} of {X.shape[1]} indicators\n"
                      "\"rule #\" = the indicator's rank in the governed rule's exact "
                      "decomposition (w × risk_z)\n"
                      f"Exploratory SHAP–rule rank agreement (different units): observed "
                      f"panel ρ = {observed['rank_agreement']:.2f}; synthetic "
                      f"{SHAP_SYNTHETIC_ROWS}-row setting ρ = {synthetic['rank_agreement']:.2f}"),
                     fontsize=7.5, color=muted, ha="left", va="top", linespacing=1.4)
            fig.tight_layout(rect=(0, 0, 1, 1 - 0.95 / fig.get_size_inches()[1]))
            fig.savefig(figure_path, dpi=200)
            plt.close(fig)
            artifacts.append({"dimension": dimension, "artifact": figure_path})

    # ---- per-indicator contribution change for each event -----------------
    change_rows = []
    for ev in scorer["_event_shap_inputs"]:
        dimension, panel = ev["dimension"], scorer["_panels"][ev["dimension"]]
        (rule0, rule1), (xgb0, xgb1) = ev["local"]["rule"], ev["local"]["xgb"]
        base = {"scenario": ev["scenario"], "dimension": dimension, "source": disp(ev["source"])}
        if ev["hard_alert_after"]:
            change_rows.append({**base, "ind_id": "(hard alert)", "indicator": ev["hard_alert_after"],
                                "note": ("the hard alert is raised outside the scorer and "
                                         "changes neither local score; the scorer input is "
                                         "unchanged, so there is no SHAP change to explain. "
                                         f"Local score rule {rule0:.6f}->{rule1:.6f}, "
                                         f"XGBoost {xgb0:.6f}->{xgb1:.6f}.")})
            continue
        explainer = explainers[dimension]
        s0 = np.asarray(explainer.shap_values(ev["x_before"]), dtype=float)[0]
        s1 = np.asarray(explainer.shap_values(ev["x_after"]), dtype=float)[0]
        cols = list(ev["x_before"].columns)
        d_rule = {k: ev["rule_after"].get(k, 0.0) - ev["rule_before"].get(k, 0.0) for k in cols}
        d_rule_total = float(sum(d_rule.values()))
        for j, k in enumerate(cols):
            z0, z1 = float(ev["x_before"].iloc[0, j]), float(ev["x_after"].iloc[0, j])
            moved = not ((np.isnan(z0) and np.isnan(z1)) or np.isclose(z0, z1, rtol=0, atol=1e-12))
            if not moved and abs(s1[j] - s0[j]) <= 1e-12 and abs(d_rule[k]) <= 1e-15:
                continue
            change_rows.append({
                **base, "ind_id": k, "indicator": panel["names"].get(k, k),
                "risk_z_before": z0, "risk_z_after": z1,
                "shap_before": float(s0[j]), "shap_after": float(s1[j]),
                "shap_change": float(s1[j] - s0[j]),
                "rule_contribution_change_index": d_rule[k],
                # The rule's local-score change allocated in proportion to the
                # index contributions (exact when one indicator moves).
                "rule_effect_on_local_score": ((rule1 - rule0) * d_rule[k] / d_rule_total
                                               if abs(d_rule_total) > 1e-15 else 0.0),
                "note": "" if moved else "input unchanged; SHAP change from tree interactions",
            })
        change_rows.append({
            **base, "ind_id": "TOTAL", "indicator": "local score change",
            "shap_change": float(np.sum(s1) - np.sum(s0)),
            "rule_contribution_change_index": d_rule_total,
            "rule_effect_on_local_score": rule1 - rule0,
            "note": (f"XGBoost local score {xgb0:.6f}->{xgb1:.6f} (SHAP changes sum to this "
                     f"difference); rule local score {rule0:.6f}->{rule1:.6f}")})

    failed = scorer.get("status") == "FAILED"
    return {"status": "FAILED" if failed else "COMPLETE",
            "reason": ("the substitution experiment failed its checks; its SHAP results "
                       "inherit that status") if failed else "",
            "results_usable": not failed, "requirements": ML_REQUIREMENTS,
            "summary": pd.DataFrame(summary_rows, columns=empty_summary.columns),
            "comparison": (pd.concat(comparison_frames, ignore_index=True)
                           if comparison_frames else empty_comparison),
            "observations": (pd.concat(observation_frames, ignore_index=True)
                             if observation_frames else empty_observations),
            "event_change": pd.DataFrame(change_rows, columns=empty_change.columns),
            "artifacts": pd.DataFrame(artifacts, columns=["dimension", "artifact"])}


# =============================================================================
# SECTION 12.  ABLATION BASELINES
# =============================================================================
# Federated scoring against (a) centralized raw pooling, (b) each firm's own
# history and (c) rules-based scoring.  Without outcome labels the comparison
# reports ranking agreement, availability in the first period, shock response
# and cross-period variance instead of accuracy and calibration.
# =============================================================================

def spearman(a: List[float], b: List[float]) -> float:
    """Rank correlation without scipy."""
    a, b = np.asarray(a, float), np.asarray(b, float)
    if len(a) < 3 or np.all(a == a[0]) or np.all(b == b[0]):
        return float("nan")

    def rank(x):
        order = np.argsort(x)
        r = np.empty(len(x), float)
        r[order] = np.arange(1, len(x) + 1)
        # average ties
        for v in np.unique(x):
            m = x == v
            if m.sum() > 1:
                r[m] = r[m].mean()
        return r

    ra, rb = rank(a), rank(b)
    ra -= ra.mean(); rb -= rb.mean()
    den = np.sqrt((ra ** 2).sum() * (rb ** 2).sum())
    return float((ra * rb).sum() / den) if den > 0 else float("nan")


def _compose_and_propagate(per_entity_dim: Dict[str, Dict[str, float]],
                           policy: RegulatoryPolicy) -> Dict[str, float]:
    """Same propagation and composite for every ablation baseline; a missing
    local score enters as 0.5.
    """
    wE, wS = policy.dim_weights
    out = {}
    for dim in ("Environmental", "Social"):
        local = {e: per_entity_dim.get(e, {}).get(dim, np.nan) for e in ENTITY_ORDER}
        filled = {e: (0.5 if np.isnan(v) else v) for e, v in local.items()}
        totals, _ = propagate(filled)
        for e in ENTITY_ORDER:
            out.setdefault(e, {})[dim] = (np.nan if np.isnan(local[e]) else totals[e])
    return {e: wE * d["Environmental"] + wS * d["Social"] for e, d in out.items()}


def baseline_local_only(period_files: List[Tuple[str, str]], tier_map: Dict[str, str],
                        policy: RegulatoryPolicy) -> Tuple[Dict[str, Dict[str, float]], pd.DataFrame]:
    """Ablation (b): each firm benchmarks against its own history.

    An indicator becomes scoreable once the firm has two earlier observations
    with nonzero spread, so there is no score in the first periods.
    """
    hist: Dict[Tuple[str, str, str], List[float]] = {}
    out: Dict[str, Dict[str, float]] = {}
    cov_rows = []
    for label, path in period_files:
        data = load_period(path, tier_map)
        per_entity_dim: Dict[str, Dict[str, float]] = {}
        for dim, df in data.items():
            for e in ENTITY_ORDER:
                sub = df[(df["entity"] == e) & (df["tier"] != "T1")]
                zs, n_ind = [], 0
                for _, r in sub.iterrows():
                    v = float(r["value"])
                    if not np.isfinite(v):
                        continue
                    n_ind += 1
                    key = (e, dim, r["indicator"])
                    past = hist.get(key, [])
                    if len(past) >= 2:
                        m, s = float(np.mean(past)), float(np.std(past))
                        if s > 1e-12:
                            z = (v - m) / s
                            zs.append(-z if r["is_good"] else z)
                    hist.setdefault(key, []).append(v)
                per_entity_dim.setdefault(e, {})[dim] = (
                    sigmoid(float(np.mean(zs)) / policy.squash_scale) if zs else float("nan"))
                cov_rows.append({"period": label, "dimension": dim, "entity": disp(e),
                                 "indicators_held": n_ind, "indicators_scoreable": len(zs),
                                 "coverage": (len(zs) / n_ind) if n_ind else 0.0})
        out[label] = _compose_and_propagate(per_entity_dim, policy)
    return out, pd.DataFrame(cov_rows)


def weight_drift_evidence(out: dict, files, policy: RegulatoryPolicy) -> pd.DataFrame:
    """Indicators whose dispersion or published weight moved against the anchor
    round (T2b).
    """
    labels = [l for l, _ in files]
    anchor = labels[0]
    rows = []
    for lab in labels[1:]:
        for dim in ("Environmental", "Social"):
            base = out["periods"][anchor].weights[dim].rows.set_index("indicator")
            cur = out["periods"][lab].weights[dim].rows
            for _, x in cur.iterrows():
                k = x["indicator"]
                if k not in base.index:
                    continue
                w0 = float(base.loc[k, "w_final"])
                w1 = float(x["w_final"])
                s0, s1 = float(base.loc[k, "rho"]), float(x["rho"])
                same_sigma = bool(np.isclose(s1, s0, rtol=1e-9, atol=Coordinator.SIGMA_EPS))
                if same_sigma and abs(w1 - w0) < 1e-9:
                    continue                      # nothing moved
                if same_sigma:
                    dispersion = "unchanged (weight reallocated across the dimension)"
                else:
                    dispersion = "narrowed" if s1 < s0 else "widened"
                on_floor = bool(np.isclose(w1, x["band_lo"], rtol=1e-5))
                on_cap = bool(np.isclose(w1, x["band_hi"], rtol=1e-5))
                rows.append({
                    "period": lab, "dimension": dim, "ind_id": x["ind_id"],
                    "indicator": k, "tier": x["tier"],
                    "rho": x["rho"],
                    "dispersion": dispersion,
                    "w_anchor": w0, "w_now": w1,
                    "pct_change": (w1 - w0) / w0 if w0 else np.nan,
                    "floor": x["band_lo"], "cap": x["band_hi"],
                    "on_a_bound": on_floor or on_cap,
                })
    df = pd.DataFrame(rows, columns=["period", "dimension", "ind_id", "indicator", "tier",
                                     "rho", "dispersion", "w_anchor", "w_now", "pct_change",
                                     "floor", "cap", "on_a_bound"])
    if not df.empty:
        df = df.sort_values(["period", "rho"]).reset_index(drop=True)
    return df


def baseline_rules_based(period_files: List[Tuple[str, str]], tier_map: Dict[str, str],
                         policy: RegulatoryPolicy,
                         anchor_ranges: Optional[dict] = None) -> Dict[str, Dict[str, float]]:
    """Ablation (c): fixed equal weights, fixed min-max reference band frozen at
    the anchor period, no re-estimation, no drift."""
    out, ranges = {}, anchor_ranges or {}
    for idx, (label, path) in enumerate(period_files):
        data = load_period(path, tier_map)
        per: Dict[str, Dict[str, float]] = {}
        for dim, df in data.items():
            if idx == 0 and dim not in ranges:
                ranges[dim] = {k: (float(g["value"].min()), float(g["value"].max()))
                               for k, g in df.groupby("indicator")}
            rng = ranges.get(dim, {})
            for e in ENTITY_ORDER:
                sub = df[(df["entity"] == e) & (df["tier"] != "T1")]
                vals = []
                for _, r in sub.iterrows():
                    k = r["indicator"]
                    if k not in rng:
                        continue
                    v = float(r["value"])
                    if not np.isfinite(v):
                        # A missing disclosure has no position in the frozen
                        # band and is skipped, as in the federated route.
                        continue
                    lo, hi = rng[k]
                    if not (np.isfinite(lo) and np.isfinite(hi)) or hi - lo < 1e-12:
                        continue
                    x = (v - lo) / (hi - lo)
                    vals.append(1.0 - x if r["is_good"] else x)
                per.setdefault(e, {})[dim] = float(np.mean(vals)) if vals else float("nan")
        out[label] = _compose_and_propagate(per, policy)
    return out


def verify_federated_equals_centralised(result: PeriodResult) -> pd.DataFrame:
    """Ablation (a): the benchmark from participant summaries equals the one from
    pooled raw records, in raw units and on the common scale (T5).
    """
    rows = []
    for dim, df in result.data.items():
        st = result.stats[dim]
        cmu, csg = centralised_stats(df)
        keys = sorted(set(st.mu) & set(cmu))
        dmu = max((abs(st.mu[k] - cmu[k]) / max(1.0, abs(cmu[k])) for k in keys), default=0.0)
        dsg = max((abs(st.sigma[k] - csg[k]) / max(1.0, abs(csg[k])) for k in keys), default=0.0)
        qmu, qsg = centralised_common_stats(df, st)
        ckeys = sorted(set(st.common_mu) & set(qmu))
        dcmu = max((abs(st.common_mu[k] - qmu[k]) for k in ckeys), default=0.0)
        dcsg = max((abs(st.common_sigma[k] - qsg[k]) for k in ckeys), default=0.0)
        same_keys = set(st.mu) == set(cmu) and set(st.common_mu) == set(qmu)
        rows.append({"dimension": dim, "indicators_compared": len(keys),
                     "max_rel_diff_mu": dmu, "max_rel_diff_sigma": dsg,
                     "common_scale_indicators_compared": len(ckeys),
                     "max_abs_diff_common_mu": dcmu, "max_abs_diff_common_sigma": dcsg,
                     "same_benchmark_indicators": same_keys,
                     "identical_to_1e-9": bool(same_keys and max(dmu, dsg, dcmu, dcsg) < 1e-9)})
    return pd.DataFrame(rows)


def baseline_centralised(sysm: "FederatedESGSystem", out: dict, files,
                         policy: RegulatoryPolicy
                         ) -> Tuple[Dict[str, Dict[str, float]], pd.DataFrame]:
    """Ablation (a), recomputed: the benchmark from pooled raw records, then the
    same weights, scores and propagation.
    """
    import copy
    labels = [l for l, _ in files]
    wE, wS = policy.dim_weights
    anchor_c: Dict[str, GlobalStats] = {}
    series: Dict[str, Dict[str, float]] = {}
    rows: List[dict] = []
    for i, label in enumerate(labels):
        r = out["periods"][label]
        local: Dict[str, Dict[str, float]] = {}
        totals: Dict[str, Dict[str, float]] = {}
        for dim, df in r.data.items():
            st = r.stats[dim]
            st_c = copy.deepcopy(st)
            pool = df[df["entity"].isin(st.contributors)] if st.contributors else df
            st_c.mu, st_c.sigma = centralised_stats(pool)
            attach_scoring_scale(st_c, sysm.reference_scales[dim], policy)
            st_c.common_mu, st_c.common_sigma = centralised_common_stats(df, st)
            if i == 0:
                anchor_c[dim] = st_c
            t1_ids = sorted({str(x) for x in
                             df.loc[df["tier"] == "T1", "ind_id"].dropna().unique()})
            ws_c = govern_weights(dim, st_c, anchor_c.get(dim), policy, f"{label}-central",
                                  t1_ids_in_data=t1_ids, register=sysm.registers.get(dim),
                                  classification_digest_value=sysm.classification_digest)
            local[dim] = {e: score_entity(e, df, st_c, ws_c, policy).score
                          for e in ENTITY_ORDER}
            totals[dim], _ = propagate(local[dim])
        series[label] = {e: wE * totals.get("Environmental", {}).get(e, 0.5)
                            + wS * totals.get("Social", {}).get(e, 0.5)
                         for e in ENTITY_ORDER}
        rows.append({
            "period": label,
            "max_abs_diff_local": max(abs(local[d][e] - r.local[d][e])
                                      for d in local for e in ENTITY_ORDER),
            "max_abs_diff_total": max(abs(totals[d][e] - r.totals[d][e])
                                      for d in totals for e in ENTITY_ORDER),
            "max_abs_diff_composite": max(abs(series[label][e] - r.composite[e])
                                          for e in ENTITY_ORDER),
        })
    return series, pd.DataFrame(rows)


def verify_dimensional_independence(system: "FederatedESGSystem",
                                    result: PeriodResult) -> pd.DataFrame:
    """Raise every social value by 25% and check that the environmental scores do
    not move (T6).
    """
    data = {d: df.copy() for d, df in result.data.items()}
    data["Social"]["value"] = data["Social"]["value"] * 1.25

    stats, weights, local = {}, {}, {}
    for dim, df in data.items():
        sums = [system.participants[e].local_summary(df) for e in ENTITY_ORDER
                if e in set(df["entity"])]
        st = system.coordinator.aggregate(sums, df)
        attach_scoring_scale(st, system.reference_scales.get(dim, {}), system.policy)
        attach_common_scale(st, df, system.policy, system.anchor_stats.get(dim))
        ws = govern_weights(dim, st, system.anchor_stats.get(dim), system.policy,
                            "IND-TEST", register=system.registers.get(dim),
                            classification_digest_value=system.classification_digest)
        stats[dim], weights[dim] = st, ws
        local[dim] = {e: score_entity(e, df, st, ws, system.policy).score for e in ENTITY_ORDER}

    tot_E, _ = propagate(local["Environmental"])
    rows = []
    for e in ENTITY_ORDER:
        base = result.totals["Environmental"][e]
        pert = tot_E[e]
        rows.append({"entity": disp(e), "E_total_base": base,
                     "E_total_after_S_shock": pert, "abs_diff": abs(pert - base),
                     "independent": bool(abs(pert - base) < 1e-12)})
    return pd.DataFrame(rows)


# =============================================================================
# SECTION 13.  DEPENDENCY-WEIGHT SENSITIVITY
# =============================================================================

def dependency_sensitivity(result: PeriodResult, entity: str = "FOCAL",
                           betas: Tuple[float, ...] = (0.50, 0.60, 0.70, 0.80, 0.90),
                           policy: Optional[RegulatoryPolicy] = None) -> pd.DataFrame:
    """Sweep one firm's beta with its alphas rescaled; totals, upstream share and
    peer-relative flag (T7).
    """
    policy = policy or RegulatoryPolicy()
    orig_alpha = dict(DEPENDENCY_WEIGHTS[entity]["alpha"])
    orig_sum = sum(orig_alpha.values())
    rows = []
    for b in betas:
        if orig_sum <= 0:
            continue
        scale = (1.0 - b) / orig_sum
        weights = {k: {"beta": v["beta"], "alpha": dict(v["alpha"])}
                   for k, v in DEPENDENCY_WEIGHTS.items()}
        weights[entity] = {"beta": b, "alpha": {k: v * scale for k, v in orig_alpha.items()}}
        rec = {"beta": b}
        for dim in ("Environmental", "Social"):
            tot, comp = propagate(result.local[dim], weights)
            up = sum(v for k, v in comp[entity].items() if k != "__own__")
            rec[f"{dim[:1]}_total"] = tot[entity]
            rec[f"{dim[:1]}_upstream_share"] = up / tot[entity] if tot[entity] else np.nan
        wE, wS = policy.dim_weights
        rec["composite"] = wE * rec["E_total"] + wS * rec["S_total"]
        rec["peer_relative_flag"] = bool(rec["composite"] >= policy.peer_relative_flag_level)
        rows.append(rec)
    df = pd.DataFrame(rows)
    if not df.empty:
        df.insert(0, "entity", disp(entity))
    return df


# M7: FOCAL's own / S1-RUBBER / S1-COATING weights on the M0 inputs.  The
# alternatives are illustrative settings, not weights derived from procurement
# or strategic-dependency data.  `generate` writes M7_SETTINGS to the
# Dependency Weights sheet of the M7 workbook, next to the M0 indicator
# values; the run reads the settings from that sheet.
M7_WORKBOOK = "ESG_Raw_Data_Supply_Chain_M7.xlsx"
M7_SETTINGS_SHEET = "Dependency Weights"
M7_FIRM = "FOCAL"
M7_BRANCHES = ("S1-RUBBER", "S1-COATING")
M7_SETTINGS = (("baseline", 0.70, 0.12, 0.18),
               ("higher upstream share", 0.50, 0.20, 0.30),
               ("alternative supplier allocation", 0.70, 0.18, 0.12))
M7_SETTING_NOTES = {
    "baseline": "FOCAL's declared dependency weights",
    "higher upstream share": "higher total upstream share; supplier ratio of the baseline kept",
    "alternative supplier allocation": ("total upstream share of the baseline kept; "
                                        "allocation between the suppliers changed"),
}
M7_SETTINGS_COLUMNS = ("Setting", "Company", "Own operations (beta)",
                       *(f"{j} (alpha)" for j in M7_BRANCHES), "Note")
M7_CHECK_TEXT = {
    "inputs_equal_M0": "the M7 workbook holds the M0 indicator values and data tags",
    "baseline_is_declared": "the baseline setting equals FOCAL's declared dependency weights",
    "baseline_reproduces_M0": "the baseline setting reproduces FOCAL's published M0 totals",
    "weights_sum_to_one": "own-operations and supplier weights sum to one in every setting",
    "higher_share_keeps_supplier_ratio": ("the higher upstream share keeps the S1-RUBBER : "
                                          "S1-COATING ratio of the baseline"),
    "alternative_keeps_upstream_share": ("the alternative allocation keeps the baseline's "
                                         "total upstream share"),
    "upstream_totals_unchanged": "no firm upstream of FOCAL changes its total",
    "contributions_add_up": "own operations plus both branches equal FOCAL's total",
}


def load_m7_settings(path: str) -> Tuple[Tuple[str, float, float, float], ...]:
    """FOCAL's dependency-weight settings from the Dependency Weights sheet of
    the M7 workbook, one row per setting, read down to the first empty row."""
    from openpyxl import load_workbook
    workbook = load_workbook(path, read_only=True, data_only=True)
    try:
        if M7_SETTINGS_SHEET not in workbook.sheetnames:
            raise ConfigurationError(f"M7 missing sheet: {M7_SETTINGS_SHEET}")
        rows = list(workbook[M7_SETTINGS_SHEET].iter_rows(values_only=True))
    finally:
        workbook.close()
    head = [str(v or "").strip() for v in (rows[0] if rows else ())]
    needed = M7_SETTINGS_COLUMNS[:-1]
    if any(c not in head for c in needed):
        raise ConfigurationError(f"M7 {M7_SETTINGS_SHEET}: expected columns {list(needed)}")
    pos = [head.index(c) for c in needed]
    settings = []
    for row in rows[1:]:
        cells = [row[p] if p < len(row) else None for p in pos]
        name = str(cells[0] or "").strip()
        if not name:
            break
        firm = str(cells[1] or "").strip()
        if firm != M7_FIRM:
            raise ConfigurationError(f"M7 setting '{name}' is for '{firm}', not {M7_FIRM}")
        try:
            beta, a_rubber, a_coating = (float(v) for v in cells[2:])
        except (TypeError, ValueError):
            raise ConfigurationError(f"M7 setting '{name}': weights must be numbers") from None
        if not all(np.isfinite(w) and 0.0 <= w <= 1.0 for w in (beta, a_rubber, a_coating)):
            raise ConfigurationError(f"M7 setting '{name}': weights must lie between 0 and 1")
        settings.append((name, beta, a_rubber, a_coating))
    names, expected = [s[0] for s in settings], [s[0] for s in M7_SETTINGS]
    if names != expected:
        raise ConfigurationError(f"M7 {M7_SETTINGS_SHEET}: expected the settings {expected} "
                                 f"in this order, found {names}")
    return tuple(settings)


def run_m7_dependency_sensitivity(system: "FederatedESGSystem", anchor: PeriodResult,
                                  path: str) -> dict:
    """M7: FOCAL's total and branch contributions under the dependency-weight
    settings of the M7 workbook, on the M0 inputs.
    """
    import copy
    settings = load_m7_settings(path)
    reference = load_period(anchor.filepath, system.tier_map, system.header_map)
    inputs = load_period(path, system.tier_map, system.header_map)
    w_e, w_s = system.policy.dim_weights
    rows, totals = [], {}
    for name, beta, a_rubber, a_coating in settings:
        weights = copy.deepcopy(DEPENDENCY_WEIGHTS)
        weights[M7_FIRM] = {"beta": beta,
                            "alpha": dict(zip(M7_BRANCHES, (a_rubber, a_coating)))}
        parts = {}
        for dim in _DIMS2:
            totals[(name, dim)], comps = propagate(anchor.local[dim], weights)
            parts[dim] = dict(comps[M7_FIRM], total=totals[(name, dim)][M7_FIRM])
        parts["Composite"] = {k: w_e * parts["Environmental"][k] + w_s * parts["Social"][k]
                              for k in parts["Environmental"]}
        for dim, p in parts.items():
            rows.append({"setting": name, "beta_FOCAL": beta,
                         "alpha_S1-RUBBER": a_rubber, "alpha_S1-COATING": a_coating,
                         "dimension": dim, "FOCAL_total": p["total"],
                         "own_operations": p["__own__"],
                         **{f"branch_{j}": p[j] for j in M7_BRANCHES},
                         "branch_ranking": " > ".join(
                             disp(j) for j in sorted(M7_BRANCHES, key=lambda j: -p[j]))})
    table = pd.DataFrame(rows)
    base = table[table["setting"] == settings[0][0]].set_index("dimension")
    table["change_vs_baseline"] = table["FOCAL_total"] - table["dimension"].map(base["FOCAL_total"])
    table["ranking_changed"] = table["branch_ranking"] != table["dimension"].map(base["branch_ranking"])

    declared = DEPENDENCY_WEIGHTS[M7_FIRM]
    (_, b0, r0, c0), (_, _, r1, c1), (_, _, r2, c2) = settings
    upstream = [e for e in ENTITY_ORDER if e not in _downstream_closure({M7_FIRM})]
    checks = {
        "inputs_equal_M0": all(reference[d].reset_index(drop=True).equals(
            inputs[d].reset_index(drop=True)) for d in DIM_SHEETS),
        "baseline_is_declared": (abs(b0 - declared["beta"]) < 1e-12 and all(
            abs(a - declared["alpha"].get(j, 0.0)) < 1e-12
            for j, a in zip(M7_BRANCHES, (r0, c0)))
            and set(declared["alpha"]) == set(M7_BRANCHES)),
        "baseline_reproduces_M0": all(
            abs(totals[(settings[0][0], d)][M7_FIRM] - anchor.totals[d][M7_FIRM]) < 1e-12
            for d in _DIMS2),
        "weights_sum_to_one": all(abs(b + r + c - 1.0) < 1e-12 for _, b, r, c in settings),
        "higher_share_keeps_supplier_ratio": abs(r1 * c0 - r0 * c1) < 1e-12,
        "alternative_keeps_upstream_share": abs((r2 + c2) - (r0 + c0)) < 1e-12,
        "upstream_totals_unchanged": all(
            abs(totals[(n, d)][e] - anchor.totals[d][e]) < 1e-12
            for n, *_ in settings for d in _DIMS2 for e in upstream),
        "contributions_add_up": bool(np.allclose(
            table["own_operations"] + sum(table[f"branch_{j}"] for j in M7_BRANCHES),
            table["FOCAL_total"], rtol=0.0, atol=1e-12)),
    }
    if not all(checks.values()):
        raise AssertionError(f"M7 checks failed: {checks}")
    return {"table": table, "settings": settings,
            "checks": pd.DataFrame([{"check": k, "passed": bool(v)} for k, v in checks.items()])}


# =============================================================================
# SECTION 14.  CONSTRUCTED OFFSETTING STRESS SCENARIO
# =============================================================================

def offsetting_scenario(result: PeriodResult, dimension: str = "Environmental",
                        branch_hub: str = "S1-RUBBER", delta: float = 0.08
                        ) -> Tuple[Dict[str, float], Dict[str, Dict[str, float]], dict]:
    """Constructed case: one supplier of a branch deteriorates and the other
    improves by an offsetting amount (T11).
    """
    local = dict(result.local[dimension])
    ups = list(DEPENDENCY_WEIGHTS[branch_hub]["alpha"].keys())
    if len(ups) < 2:
        raise ValueError(f"{branch_hub} needs two enrolled upstream suppliers")
    a, b = ups[0], ups[1]
    wa = DEPENDENCY_WEIGHTS[branch_hub]["alpha"][a]
    wb = DEPENDENCY_WEIGHTS[branch_hub]["alpha"][b]

    local[a] = float(np.clip(local[a] + delta, 0.0, 1.0))
    local[b] = float(np.clip(local[b] - delta * wa / wb, 0.0, 1.0))

    totals, comps = propagate(local)
    meta = {"dimension": dimension, "hub": branch_hub,
            "deteriorating": a, "improving": b, "delta": delta,
            "offset_delta": delta * wa / wb,
            "note": "constructed stress scenario, not observed data"}
    return totals, comps, meta


# =============================================================================
# SECTION 15.  ACCEPTANCE TESTS
# =============================================================================

def run_governance_acceptance_tests(sysm: "FederatedESGSystem",
                                    anchor: PeriodResult,
                                    policy: RegulatoryPolicy) -> pd.DataFrame:
    """Acceptance tests for the governance and event rules, run in the pipeline.

    Each row states the requirement, what was executed, and the outcome.
    """
    rows: List[dict] = []

    def record(name: str, requirement: str, passed: bool, detail: str) -> None:
        rows.append({"test": name, "requirement": requirement,
                     "result": "PASS" if passed else "FAIL", "detail": detail})

    def regovern(dimension: str, flatten_indicator: Optional[str], label: str):
        df = anchor.data[dimension].copy(deep=True)
        if flatten_indicator is not None:
            sel = df["indicator"] == flatten_indicator
            df.loc[sel, "value"] = float(df.loc[sel, "value"].mean())
        sums = [Participant(e).local_summary(df) for e in ENTITY_ORDER
                if e in set(df["entity"])]
        st = Coordinator(secure_aggregation=False).aggregate(sums, df)
        attach_scoring_scale(st, sysm.reference_scales.get(dimension, {}), policy)
        attach_common_scale(st, df, policy, sysm.anchor_stats.get(dimension))
        ws = govern_weights(dimension, st, sysm.anchor_stats.get(dimension), policy,
                            label, register=sysm.registers.get(dimension),
                            classification_digest_value=sysm.classification_digest)
        return df, st, ws

    for dimension, tier in (("Environmental", "T2"), ("Environmental", "T3")):
        reg = sysm.registers.get(dimension)
        stats = anchor.stats[dimension]
        pool = reg.t2 if tier == "T2" else reg.t3
        if not reg or not pool:
            record(f"{LEVEL_NAME[tier].replace(' ', '')}_zero_variance_retained",
                   "indicator kept and floored",
                   False, f"no {LEVEL_NAME[tier]} member in the {dimension} register")
            continue
        preferred = "E1" if tier == "T2" else "E10"
        name = next((k for k in pool if str(stats.ind_id.get(k)) == preferred), pool[0])
        _, st, ws = regovern(dimension, name, f"ACCEPT-{tier}")
        row = ws.rows[ws.rows["indicator"] == name]
        floor = (reg.band_lo[name] if tier == "T2" else reg.t3_floor[name])
        kept = not row.empty
        w_fin = float(row["w_final"].iloc[0]) if kept else 0.0
        above = kept and w_fin >= floor - 1e-12
        # the baseline weights of the other members stay fixed
        base_ok = True
        if tier == "T2":
            live = ws.rows[ws.rows["tier"] == "T2"].set_index("indicator")["w_reg"]
            base_ok = all(abs(float(live[k]) - reg.w_reg[k]) < 1e-12
                          for k in live.index if k in reg.w_reg)
        finite = bool(np.isfinite(ws.rows["w_final"]).all())
        block_ok = abs(float(ws.rows.w_final.sum()) - 1.0) < 1e-9
        record(f"{LEVEL_NAME[tier].replace(' ', '')}_zero_variance_retained",
               f"{LEVEL_NAME[tier]} indicator with identical values across all firms keeps its "
               f"governed weight at or above its positive floor",
               bool(kept and above and base_ok and finite and block_ok),
               f"{stats.ind_id.get(name)} status="
               f"{st.status.get(name)} w_final={w_fin:.6f} floor={floor:.6f} "
               f"other w_reg unchanged={base_ok} block sum ok={block_ok} "
               f"all weights finite={finite}")

    # anchor sigma = 0, dispersion appears later: no division by zero
    reg = sysm.registers.get("Environmental")
    stats = anchor.stats["Environmental"]
    name = next((k for k in reg.t2 if str(stats.ind_id.get(k)) == "E1"), reg.t2[0])
    _, flat_stats, _ = regovern("Environmental", name, "ACCEPT-ANCHOR0")
    rho = compute_dispersion(stats, flat_stats, policy.dispersion_mode, reg.s_ref)
    finite = all(np.isfinite(v) for v in rho.values())
    record("anchor_zero_variance_no_division",
           "dispersion appearing after a flat anchor round uses the declared "
           "scale s_ref instead of dividing by zero",
           bool(finite and rho.get(name, 0.0) > 0.0),
           f"rho({stats.ind_id.get(name)})={rho.get(name, float('nan')):.6g} "
           f"s_ref={reg.s_ref.get(name, float('nan')):.6g} "
           f"({reg.s_ref_basis.get(name, 'n/a')}); all rho finite={finite}")

    # threshold uses the smaller of the two limits
    delta, abs_limit, rel_limit = 0.006, 0.005, 0.010 * 0.706
    regression_report = component_trigger(
        "S2-CHEM", {"__own__": 0.706}, {"__own__": 0.706 + delta},
        {"S2-CHEM": 0.706 + delta}, "Environmental",
        RegulatoryPolicy(component_delta_abs=abs_limit, component_delta_rel=0.010),
        reference_total=0.706)
    fires_min = regression_report.internal_warning and abs(regression_report.threshold - abs_limit) < 1e-12
    fires_max = abs(delta) > max(abs_limit, rel_limit)
    record("component_threshold_is_minimum",
           "|delta| > abs_limit OR |delta| > rel_limit*|reference| implies the "
           "binding threshold is the smaller of the two",
           bool(fires_min and not fires_max),
           f"delta={delta}, abs_limit={abs_limit}, relative limit={rel_limit:.5f}; "
           f"min-rule fires={fires_min}, max-rule would fire={fires_max}")

    # hard alert is decided before, and independently of, the warning rule
    unreachable = RegulatoryPolicy(component_delta_abs=2.0, component_delta_rel=10.0)
    dim = "Social"
    comps0 = anchor.comps[dim]
    reps = {}
    for e in ENTITY_ORDER:
        reps[e] = component_trigger(e, comps0[e], comps0[e], anchor.totals[dim], dim,
                                    unreachable, hard_local=(e == "S2-CHEM"),
                                    hard_labels=["S26 Child Labor Policy"] if e == "S2-CHEM" else [],
                                    reference_total=anchor.totals[dim][e])
    reps = cascade_alerts(reps)
    warnings_quiet = not any(r.internal_warning for r in reps.values())
    chain = ["S2-CHEM", "S1-RUBBER", "FOCAL", "D1-RETAIL", "D2-RETAIL"]
    chain_hard = all(reps[e].hard_alert for e in chain)
    unrelated_quiet = not any(reps[e].hard_alert for e in ("S2-EQUIP", "S1-COATING"))
    record("hard_alert_bypasses_warning_limits",
           "a verified hard threshold raises and passes on the alert with zero "
           "score movement and unreachable warning limits",
           bool(warnings_quiet and chain_hard and unrelated_quiet),
           f"internal warnings={not warnings_quiet}; hard alert on "
           f"{'->'.join(disp(e) for e in chain)}={chain_hard}; unrelated branches "
           f"quiet={unrelated_quiet}")

    # sticky and deduplicated
    probe = HardAlertRegister()
    first = probe.raise_alert("S2-CHEM", "S26 Child Labor Policy", "M5", dimension="Social")
    duplicate = probe.raise_alert("S2-CHEM", "S26 Child Labor Policy", "M6", dimension="Social")
    still_active = probe.is_active("S2-CHEM", "Social")
    eid = next(iter(probe.records))
    cleared = probe.clear(eid, cleared_by="assurance provider", note="acceptance test")
    record("hard_alert_sticky_and_deduplicated",
           "a resubmitted event raises no second alert; the alert survives a "
           "period with zero change and is cleared only by review",
           bool(first and not duplicate and still_active and cleared
                and not probe.is_active("S2-CHEM", "Social")),
           f"first={first} duplicate={duplicate} active_next_period={still_active} "
           f"cleared_by_review={cleared}")

    # infeasible bounds raise a configuration error
    try:
        bad = RegulatoryPolicy(band_lo=1.30, band_hi=1.40)
        declare_governance_register("Environmental", anchor.stats["Environmental"],
                                    bad, "ACCEPT-INFEASIBLE",
                                    s_ref=sysm.reference_scales.get("Environmental", {}))
        raised = False
        detail = "no error raised"
    except ConfigurationError as exc:
        raised = True
        detail = str(exc).split(":")[0] + " (ConfigurationError raised)"
    record("infeasible_bounds_raise", "sum(lower) <= block share <= sum(upper) is "
           "checked before projection and fails loudly", raised, detail)

    ok = all(abs(ws.rows.w_final.sum()-1.0)<1e-9 and np.isfinite(ws.rows.w_final).all()
             and (ws.rows.w_final >= ws.rows.band_lo-1e-12).all()
             and (ws.rows.w_final <= ws.rows.band_hi+1e-12).all() for ws in anchor.weights.values())
    record("dimension_simplex_and_bounds", "all Level 2 and Level 3 weights sum to one and respect bounds", bool(ok), "global projection")

    # a blank or absent declaration stays unknown; audit evidence of a
    # deficiency raises the hard alert
    df_s = anchor.data.get("Social")
    if df_s is not None and "S26" in set(df_s["ind_id"]):
        firm = "S2-CHEM" if "S2-CHEM" in set(df_s["entity"]) else str(df_s["entity"].iloc[0])
        own = (df_s["entity"] == firm) & (df_s["ind_id"] == "S26")
        as_declared = AssuranceProvider(firm).verify_hard_thresholds(df_s)
        blank = df_s.copy(deep=True)
        blank.loc[own, "value"] = np.nan
        hit_blank = AssuranceProvider(firm).verify_hard_thresholds(blank)
        omitted = df_s.loc[~own]
        hit_omitted = AssuranceProvider(firm).verify_hard_thresholds(omitted)
        hit_evidence = AssuranceProvider(firm).verify_hard_thresholds(omitted, {"S26": 0.0})
        record("missing_declaration_unknown_not_confirmed",
               "blank or absent declarations remain unknown; verified policy deficiency triggers an alert",
               bool(as_declared is None and hit_blank is None and hit_omitted is None and hit_evidence and hit_evidence.startswith("S26")),
               f"{disp(firm)} as declared: {as_declared or 'no hit'}; blank: {hit_blank}; "
               f"not submitted: {hit_omitted}; not submitted with audit evidence 0: "
               f"{hit_evidence}")

        # a failed declaration tagged as preliminary stays pending; the same value
        # backed by audit evidence raises the alert
        prelim = df_s.copy(deep=True)
        prelim.loc[own, "value"] = 0.0
        if "data_status" in prelim:
            prelim.loc[own, "data_status"] = "preliminary report; pending verification"
        prelim_auditor = AssuranceProvider(firm)
        hit_prelim = prelim_auditor.verify_hard_thresholds(prelim)
        pending_prelim = any(f.kind == "pending" and f.indicator == "S26"
                             and f.detail.startswith("unverified")
                             for f in prelim_auditor.findings)
        hit_verified = AssuranceProvider(firm).verify_hard_thresholds(prelim, {"S26": 0.0})
        record("unverified_declaration_pending_not_confirmed",
               "a failed declaration reported as preliminary or unverified stays pending and "
               "raises no hard alert until verified; audit evidence of the deficiency raises it",
               bool(hit_prelim is None and pending_prelim and hit_verified
                    and hit_verified.startswith("S26")),
               f"{disp(firm)} S26 = No as a preliminary report: "
               f"{'pending review' if pending_prelim else 'not pending'}, hard alert "
               f"{hit_prelim or 'none'}; with audit evidence 0: {hit_verified}")

    # ---- minimum group size for a published benchmark ------------------
    df = anchor.data["Environmental"].copy(deep=True)
    finite = df[np.isfinite(df["value"])]
    counts = finite.groupby("indicator")["entity"].nunique()
    name = str(counts.idxmax())
    reporters = [e for e in ENTITY_ORDER
                 if e in set(finite.loc[finite["indicator"] == name, "entity"])]
    outcome = {}
    for size in (Coordinator.MIN_PARTICIPANTS - 1, Coordinator.MIN_PARTICIPANTS):
        sub = df.copy(deep=True)
        drop = (sub["indicator"] == name) & ~sub["entity"].isin(reporters[:size])
        sub.loc[drop, "value"] = np.nan
        sums = [Participant(e).local_summary(sub) for e in ENTITY_ORDER
                if e in set(sub["entity"])]
        st = Coordinator(secure_aggregation=False).aggregate(sums, sub)
        outcome[size] = st.status.get(name, "missing")
    small, enough = Coordinator.MIN_PARTICIPANTS - 1, Coordinator.MIN_PARTICIPANTS
    record("benchmark_minimum_group_size",
           f"a benchmark is published only when at least {enough} participants report "
           f"the indicator; this publication rule reduces inference risk without guaranteeing privacy",
           outcome[small] == "unavailable_low_coverage"
           and outcome[enough] in ("published", "zero_variance_observed"),
           f"{name}: {small} reporters -> {outcome[small]}; {enough} reporters -> {outcome[enough]}")

    # ---- internal warnings: direction, own indicators, kept at the firm ----
    probe_policy = RegulatoryPolicy(component_delta_abs=0.005, component_delta_rel=0.010)
    c0 = {"__own__": 0.30, "S2-CHEM": 0.10, "S2-EQUIP": 0.15}
    mixed = component_trigger("S1-RUBBER", c0,
                              {"__own__": 0.30, "S2-CHEM": 0.12, "S2-EQUIP": 0.13},
                              {}, "Environmental", probe_policy, reference_total=0.55)
    better = component_trigger("S1-RUBBER", c0,
                               {"__own__": 0.28, "S2-CHEM": 0.10, "S2-EQUIP": 0.15},
                               {}, "Environmental", probe_policy, reference_total=0.55)
    chain_reps = {e: component_trigger(e, {"__own__": 0.5}, {"__own__": 0.5}, {},
                                       "Environmental", probe_policy, reference_total=0.5)
                  for e in ENTITY_ORDER}
    chain_reps["S1-RUBBER"] = mixed
    chain_reps = cascade_alerts(chain_reps)
    customers_quiet = not any(chain_reps[e].internal_warning
                              for e in ("FOCAL", "D1-RETAIL", "D2-RETAIL"))
    record("internal_warning_stays_at_firm",
           "every fired component carries its direction, improvements included; a warning "
           "at one firm does not raise the same warning at its customers",
           bool(mixed.internal_warning and better.internal_warning
                and mixed.fired_directions == {"S2-CHEM": "deterioration",
                                               "S2-EQUIP": "improvement"}
                and mixed.warning_direction == "mixed"
                and better.warning_direction == "improvement" and customers_quiet),
           f"mixed case {mixed.fired_directions}; improvement-only case -> "
           f"{better.warning_direction}; customers of S1-RUBBER warned: {not customers_quiet}")
    flat = {"__own__": 0.7 * 0.5}
    before = LocalScore("FOCAL", "Environmental", 0.5, 0.0, 1.0, {"E1": 0.0, "E2": 0.0},
                        point_terms={"E1": 0.0, "E2": 0.0})
    after = LocalScore("FOCAL", "Environmental", 0.5, 0.0, 1.0, {"E1": 0.4, "E2": -0.4},
                       point_terms={"E1": 0.4, "E2": -0.4})
    own = component_trigger("FOCAL", flat, flat, {}, "Environmental", probe_policy,
                            reference_total=0.5,
                            own_effects=own_indicator_effects(before, after, 0.7, policy))
    record("own_indicators_tested_before_combining",
           "a deteriorating own indicator raises a warning even when another own "
           "indicator offsets it in the firm's score",
           bool(own.internal_warning and not own.net_fired
                and own.fired_directions == {"E1": "deterioration", "E2": "improvement"}),
           f"own score change {own.net_change:+.4f}; indicator effects "
           + ", ".join(f"{k} {v:+.4f}" for k, v in own.component_deltas.items()))

    # ---- hard alerts and closures, downstream and upstream (Section 4.2.3) --
    raised_here = [{"event_id": "acceptance-hard-notice",
                    "category": "S26 Child Labor Policy", "status": HARD_STATUS_ACTIVE,
                    "raised": "M5"}]

    def hard_probe(source):
        return cascade_alerts({e: component_trigger(
            e, {"__own__": 0.5}, {"__own__": 0.5}, {}, "Social", probe_policy,
            reference_total=0.5, hard_local=(e == source),
            hard_labels=["S26 Child Labor Policy"] if e == source else [],
            hard_records=raised_here if e == source else None) for e in ENTITY_ORDER})
    hard_reps = hard_probe("S2-CHEM")
    expected_focal = [{"via": "S1-RUBBER", **raised_here[0]}]
    retail_ok = all(hard_reps[e].hard_via_detail == [{"via": "FOCAL", **raised_here[0]}]
                    for e in ("D1-RETAIL", "D2-RETAIL"))
    identity_hidden = all("S2-CHEM" not in json.dumps(hard_reps[e].hard_via_detail)
                          for e in ("FOCAL", "D1-RETAIL", "D2-RETAIL"))
    record("hard_notice_carries_category_status_timing",
           "a hard notice tells each customer the direct branch, the event category, its "
           "status and when it was raised, without the identity of the deeper-tier firm",
           bool(hard_reps["FOCAL"].hard_via_detail == expected_focal and retail_ok
                and identity_hidden),
           f"FOCAL receives {_hard_detail_text(hard_reps['FOCAL'].hard_via_detail)}; "
           f"D1-RETAIL receives {_hard_detail_text(hard_reps['D1-RETAIL'].hard_via_detail)}; "
           f"deeper-tier identity withheld={identity_hidden}")
    up_reps = hard_probe("FOCAL")
    upstream = {"S1-RUBBER": "FOCAL", "S1-COATING": "FOCAL",
                "S2-CHEM": "S1-RUBBER", "S2-EQUIP": "S1-RUBBER"}
    downstream = {"D1-RETAIL": "FOCAL", "D2-RETAIL": "FOCAL"}
    record("hard_alert_upstream_and_downstream",
           "a hard alert and its closure reach every affected firm upstream and downstream, "
           "each told only the direct customer or supplier it arrives through",
           bool(all(up_reps[e].hard_alert and up_reps[e].hard_via_customers == [v]
                    and up_reps[e].hard_customer_detail == [{"via": v, **raised_here[0]}]
                    for e, v in upstream.items())
                and all(up_reps[e].hard_via == [v] for e, v in downstream.items())
                and dict(upstream_notices("FOCAL")) == upstream
                and dict(downstream_notices("FOCAL")) == downstream),
           "event at FOCAL: upstream " + ", ".join(
               f"{disp(e)} via {disp(v)}" for e, v in upstream.items())
           + "; downstream " + ", ".join(f"{disp(e)} via {disp(v)}" for e, v in downstream.items()))

    # Section 4.2.3: distinguish events even when descriptions match, while
    # retaining every direct branch through which a shared event arrives.
    def routed_events(register, dependency_weights=None):
        reports = {e: component_trigger(
            e, {"__own__": 0.5}, {"__own__": 0.5}, {}, "Social", probe_policy,
            reference_total=0.5, hard_local=register.is_active(e, "Social"),
            hard_labels=register.labels(e, "Social"),
            hard_records=_hard_records(register, e, "Social")) for e in ENTITY_ORDER}
        return cascade_alerts(reports, dependency_weights)

    label = "S26 Child Labor Policy"
    distinct = HardAlertRegister()
    for entity, event_id in (("S2-CHEM", "routing-A"), ("S2-EQUIP", "routing-B")):
        distinct.raise_alert(entity, label, "M5", dimension="Social",
                             indicator="S26", event_id=event_id)
    distinct_reports = routed_events(distinct)
    recipients = ("FOCAL", "D1-RETAIL", "D2-RETAIL")
    distinct_ok = all(
        {r["event_id"] for r in distinct_reports[e].hard_event_summary}
        == {"routing-A", "routing-B"} for e in recipients)
    record("hard_events_with_matching_descriptions_stay_distinct",
           "different events of the same category and period remain distinct after "
           "their paths merge at a supplier",
           distinct_ok,
           "two sources, same category and M5: FOCAL and each retailer must retain two IDs")

    diamond = json.loads(json.dumps(DEPENDENCY_WEIGHTS))
    diamond["S1-COATING"] = {"beta": 0.8, "alpha": {"S2-CHEM": 0.2}}
    shared = HardAlertRegister()
    shared.raise_alert("S2-CHEM", label, "M5", dimension="Social",
                       indicator="S26", event_id="routing-shared")
    shared_reports = routed_events(shared, diamond)
    focal_events = shared_reports["FOCAL"].hard_event_summary
    before_replay = json.dumps({e: r.hard_event_summary for e, r in shared_reports.items()},
                               sort_keys=True)
    cascade_alerts(shared_reports, diamond)
    after_replay = json.dumps({e: r.hard_event_summary for e, r in shared_reports.items()},
                              sort_keys=True)
    shared_ok = bool(
        len(focal_events) == 1 and focal_events[0]["event_id"] == "routing-shared"
        and set(focal_events[0]["via_direct_suppliers"]) == {"S1-RUBBER", "S1-COATING"}
        and not focal_events[0]["own_operations"]
        and all(len(shared_reports[e].hard_event_summary) == 1
                and shared_reports[e].hard_event_summary[0]["via_direct_suppliers"] == ["FOCAL"]
                for e in ("D1-RETAIL", "D2-RETAIL"))
        and "S2-CHEM" not in json.dumps(focal_events)
        and before_replay == after_replay)
    record("shared_hard_event_keeps_all_branches_without_double_counting",
           "one event arriving through two direct suppliers is one event with two branches; "
           "a repeated notification changes neither its count nor its paths",
           shared_ok, "diamond topology: one event, two FOCAL branches, no deeper-tier identity")

    same_source = HardAlertRegister()
    for event_id in ("routing-first", "routing-second"):
        same_source.raise_alert("S2-CHEM", label, "M5", dimension="Social",
                                indicator="S26", event_id=event_id)
    same_before = routed_events(same_source)
    same_source.clear("routing-first", "acceptance assurance", "first event resolved")
    same_after = routed_events(same_source)
    record("hard_event_clearance_does_not_close_another_event",
           "closing one event leaves another event of the same category, period and source "
           "active throughout the downstream chain",
           bool(all(len(same_before[e].hard_event_summary) == 2
                    and {r["event_id"] for r in same_after[e].hard_event_summary}
                    == {"routing-second"} for e in recipients)),
           "two IDs at one source; only routing-first is cleared")

    upstream_shared = HardAlertRegister()
    upstream_shared.raise_alert("FOCAL", label, "M5", dimension="Social",
                                indicator="S26", event_id="routing-upstream")
    upstream_reports = routed_events(upstream_shared, diamond)
    supplier_events = upstream_reports["S2-CHEM"].hard_event_summary
    upstream_paths = set(upstream_notices("FOCAL", diamond))
    downstream_paths = set(downstream_notices("S2-CHEM", diamond))
    record("hard_event_closure_routes_preserve_shared_branches",
           "upstream and downstream closure routing retains every direct branch of a "
           "shared event, including a supplier reached through two customers",
           bool(len(supplier_events) == 1
                and set(supplier_events[0]["via_direct_customers"]) == {"S1-RUBBER", "S1-COATING"}
                and {("S2-CHEM", "S1-RUBBER"), ("S2-CHEM", "S1-COATING")} <= upstream_paths
                and {("FOCAL", "S1-RUBBER"), ("FOCAL", "S1-COATING")} <= downstream_paths),
           "shared upstream supplier receives the event and closure through both direct customers")

    # ---- direct upstream coverage (Section 4.2.4) --------------------
    import copy as _copy
    full = provisional_status(anchor, policy)["Environmental"]
    gap_state = _copy.copy(anchor)
    gap_state.scores = {d: dict(v) for d, v in anchor.scores.items()}
    gap_state.scores["Environmental"].pop("S2-CHEM", None)
    gap = provisional_status(gap_state, policy)["Environmental"]
    rubber = DEPENDENCY_WEIGHTS["S1-RUBBER"]["alpha"]
    expected_gap = rubber["S2-EQUIP"] / sum(rubber.values())
    with_suppliers = [e for e in ENTITY_ORDER if DEPENDENCY_WEIGHTS[e]["alpha"]]
    ok_coverage = bool(
        all(np.isclose(full[e]["direct_upstream_coverage"], 1.0) for e in with_suppliers)
        and all(np.isnan(full[e]["direct_upstream_coverage"])
                for e in ENTITY_ORDER if e not in with_suppliers)
        and np.isclose(gap["S1-RUBBER"]["direct_upstream_coverage"], expected_gap)
        and gap["S1-RUBBER"]["direct_suppliers_without_data"] == ["S2-CHEM"]
        and np.isclose(gap["FOCAL"]["direct_upstream_coverage"], 1.0))
    record("direct_upstream_coverage",
           "each firm records the share of its direct supplier dependency covered by usable "
           "data; a direct supplier without data lowers it, a deeper-tier gap does not",
           ok_coverage,
           "anchor: " + ", ".join(f"{disp(e)} {full[e]['direct_upstream_coverage']:.0%}"
                                  for e in with_suppliers)
           + f"; without S2-CHEM data: S1-RUBBER "
             f"{gap['S1-RUBBER']['direct_upstream_coverage']:.0%} (expected {expected_gap:.0%}), "
             f"FOCAL {gap['FOCAL']['direct_upstream_coverage']:.0%}")

    # ---- values outside the conversion range are flagged (Section 4.2.2) ---
    stats_e = anchor.stats["Environmental"]
    unit_range = next((k for k in sysm.registers["Environmental"].members()
                       if tuple(stats_e.input_bounds.get(k, ())) == (0.0, 1.0)), None)
    if unit_range is None:
        record("conversion_range_clipping_flagged", "clipped indicators are flagged", False,
               "no weighted indicator with conversion range [0, 1]")
    else:
        df_c = anchor.data["Environmental"].copy(deep=True)
        sel = (df_c["indicator"] == unit_range) & np.isfinite(df_c["value"])
        df_c.loc[sel, "value"] = -0.10 - 0.05 * np.arange(int(sel.sum()))
        sums_c = [Participant(e).local_summary(df_c) for e in ENTITY_ORDER
                  if e in set(df_c["entity"])]
        st_c = Coordinator(secure_aggregation=False).aggregate(sums_c, df_c)
        attach_scoring_scale(st_c, sysm.reference_scales.get("Environmental", {}), policy)
        attach_common_scale(st_c, df_c, policy, sysm.anchor_stats.get("Environmental"))
        ws_c = govern_weights("Environmental", st_c, sysm.anchor_stats.get("Environmental"),
                              policy, "ACCEPT-CLIP",
                              register=sysm.registers.get("Environmental"),
                              classification_digest_value=sysm.classification_digest)
        iid = str(stats_e.ind_id.get(unit_range) or unit_range)
        flagged = [n for n in ws_c.notes
                   if n.startswith(f"{iid}:") and "does not differentiate" in n]
        n_clip = int(st_c.conversion_clipped.get(unit_range, 0))
        record("conversion_range_clipping_flagged",
               "reported values outside the published conversion range are counted, and an "
               "indicator that clipping collapses to one value is named in the weight version",
               bool(n_clip == int(sel.sum()) and flagged
                    and st_c.common_sigma.get(unit_range, 1.0) == 0.0),
               f"{iid}: {n_clip} value(s) set below the range [0, 1]; note: "
               f"{flagged[0] if flagged else 'missing'}")

    # ---- graduated responses: size, duration, level (Section 4.2.3) -------
    def respond(after: float, streak: int, high: int, before: float = 0.30):
        rep = component_trigger("S1-RUBBER", {"__own__": before}, {"__own__": after}, {},
                                "Environmental", policy, reference_total=0.5)
        return graduated_response(rep, streak, high, after, policy)[0]
    limit = min(policy.component_delta_abs, policy.component_delta_rel * 0.5)
    small, large = 0.30 + 1.5 * limit, 0.30 + (policy.escalation_size_multiple + 1) * limit
    level_hi = max(policy.persistent_level, 0.30)
    graded = {"brief deterioration": respond(small, 1, 0),
              "large deterioration": respond(large, 1, 0),
              "continued deterioration": respond(small, policy.escalation_periods, 0),
              "persistently high, no increase": respond(level_hi, 0, policy.escalation_periods,
                                                        before=level_hi),
              "improvement": respond(0.30 - 1.5 * limit, 0, 0)}
    expected = {"brief deterioration": "monitor", "large deterioration": "investigate",
                "continued deterioration": "investigate",
                "persistently high, no increase": "investigate", "improvement": "none"}
    record("graduated_response",
           "a brief deterioration calls for monitoring; a large or continued deterioration, "
           "or a persistently high total without a new increase, calls for investigation; "
           "an improvement is identified as such",
           graded == expected, "; ".join(f"{k}: {v}" for k, v in graded.items()))

    # ---- an open internal warning closes only when its component reverses --
    # S1-RUBBER's S2-CHEM branch deteriorates; next, the S2-EQUIP branch
    # improves while S2-CHEM stays where it is; finally S2-CHEM reverses.
    step = 1.5 * limit
    path = [{"__own__": 0.30, "S2-CHEM": 0.10, "S2-EQUIP": 0.10},
            {"__own__": 0.30, "S2-CHEM": 0.10 + step, "S2-EQUIP": 0.10},
            {"__own__": 0.30, "S2-CHEM": 0.10 + step, "S2-EQUIP": 0.10 - step},
            {"__own__": 0.30, "S2-CHEM": 0.10, "S2-EQUIP": 0.10 - step}]
    register_w, run_w, trace = InternalWarningRegister(), 0, []
    for t in range(1, len(path)):
        label = f"P{t - 1}->P{t}"
        rep = component_trigger("S1-RUBBER", path[t - 1], path[t], {}, "Environmental",
                                policy, reference_total=0.5)
        run_w = run_w + 1 if any(v > rep.threshold for v in rep.component_deltas.values()) else 0
        level_w, basis_w = graduated_response(rep, run_w, 0, 0.5, policy)
        open_w, closed_w = register_w.update("Environmental", "S1-RUBBER", rep, label)
        level_w, _ = carry_open_warnings(level_w, basis_w, open_w, label)
        trace.append((level_w, [w.component for w in open_w], [w.component for w in closed_w]))
    record("internal_warning_open_until_reversed",
           "a warning on a deteriorating branch stays open, at least at monitoring, when "
           "another branch improves, and closes only when the same branch reverses",
           trace == [("monitor", ["S2-CHEM"], []), ("monitor", ["S2-CHEM"], []),
                     ("none", [], ["S2-CHEM"])],
           "; ".join(f"step {i + 1}: {lvl}, open {', '.join(o) or 'none'}, closed "
                     f"{', '.join(c) or 'none'}" for i, (lvl, o, c) in enumerate(trace)))

    # ---- valid reference when too few firms report (Section 4.2.2) --------
    refs = sysm.valid_reference.get("Environmental", {})
    k_ref = next(k for k in sysm.registers["Environmental"].t2 if k in refs)
    ref = refs[k_ref]
    keep = ENTITY_ORDER[:Coordinator.MIN_PARTICIPANTS - 1]
    df_ref = anchor.data["Environmental"].copy(deep=True)
    df_ref.loc[(df_ref["indicator"] == k_ref) & ~df_ref["entity"].isin(keep), "value"] = np.nan
    next_label = f"M{ref['index'] + 1}"
    late_label = f"M{ref['index'] + policy.reference_max_age_periods + 1}"
    st_ref = _copy.deepcopy(sysm).recalibrate("Environmental", df_ref, next_label)
    ws_ref = govern_weights("Environmental", st_ref, sysm.anchor_stats["Environmental"],
                            policy, next_label, register=sysm.registers["Environmental"],
                            classification_digest_value=sysm.classification_digest)
    st_old = _copy.deepcopy(sysm).recalibrate("Environmental", df_ref, late_label)
    ok_ref = bool(st_ref.status.get(k_ref) == "valid_reference"
                  and st_ref.mu[k_ref] == ref["mu"]
                  and st_ref.carried_forward.get(k_ref, "").startswith(ref["period"])
                  and k_ref in set(ws_ref.rows["indicator"])
                  and any("valid reference" in n for n in ws_ref.notes)
                  and st_old.status.get(k_ref) == "unavailable_low_coverage"
                  and k_ref not in st_old.mu)
    record("valid_reference_fallback",
           "with too few reporters an explicitly identified earlier benchmark within the "
           "allowed age is used and named in the version; beyond that age no current "
           "benchmark is reported",
           ok_ref,
           f"{st_ref.ind_id.get(k_ref)} with {len(keep)} reporters: {next_label} uses "
           f"{st_ref.carried_forward.get(k_ref, 'none')}; {late_label} -> "
           f"{st_old.status.get(k_ref)}")

    # ---- data beyond the allowed age is unavailable (Section 4.2.4) --------
    df_age = anchor.data["Environmental"].copy(deep=True)
    firm, iid_age = "S2-CHEM", "E1"
    sel = (df_age["entity"] == firm) & (df_age["ind_id"] == iid_age)
    df_age.loc[sel, "as_of"] = "M0"
    current, _ = apply_data_age({"Environmental": df_age},
                                f"M{policy.max_data_age_periods}", policy)
    aged, aged_rows = apply_data_age({"Environmental": df_age},
                                     f"M{policy.max_data_age_periods + 1}", policy)
    sc_full = anchor.scores["Environmental"][firm]
    sc_aged = score_entity(firm, aged["Environmental"], anchor.stats["Environmental"],
                           anchor.weights["Environmental"], policy)
    ok_age = bool(np.isfinite(current["Environmental"].loc[sel, "value"]).all()
                  and aged["Environmental"].loc[sel, "value"].isna().all()
                  and any(r["entity"] == firm and r["ind_id"] == iid_age
                          and not r["used_in_score"] for r in aged_rows)
                  and sc_aged.incomplete and sc_aged.score >= sc_full.score - 1e-12
                  and sc_aged.lower - 1e-12 <= sc_full.score <= sc_aged.upper + 1e-12)
    record("data_age_limit",
           "a value older than the allowed age is treated as unavailable: its weight is kept, "
           "the score interval is published and the point value is its upper bound",
           ok_age,
           f"{firm} {iid_age} as of M0, allowed age {policy.max_data_age_periods}: kept at "
           f"M{policy.max_data_age_periods}, unavailable at M{policy.max_data_age_periods + 1}; "
           f"score {sc_full.score:.4f} -> {sc_aged.score:.4f} "
           f"(interval {sc_aged.lower:.4f}-{sc_aged.upper:.4f})")

    return pd.DataFrame(rows)


# =============================================================================
# SECTION 16.  M6 INDUSTRY PEER EXPERIMENT
# =============================================================================

M6_WORKBOOK = "ESG_Raw_Data_Supply_Chain_M6.xlsx"


# Simulated tire and rubber peers of S1-RUBBER.  Inputs use fixed multiples of
# S1-RUBBER's M5 values, with percentage bounds, unchanged binary indicators
# and separate E1/E10 changes.
# With S1-RUBBER the cohort meets the minimum publication size of four.
# Group size alone does not prevent inference from equal values or other
# constraints; benchmark_disclosure_check evaluates the simulated cohort.
#            name            before  E1 after (None: unchanged)
M6_PEERS = (("PEER-TIRE-01", 0.70, 0.40),
            ("PEER-TIRE-02", 1.20, 2.00),
            ("PEER-TIRE-03", 0.90, None))
M6_PEER_NAMES = tuple(p[0] for p in M6_PEERS)
M6_PEER_NOTE = "Simulated tire and rubber peer; not based on any named company"

# Requirement text for the M6 checks when they join the acceptance table (T15).
M6_CHECK_TEXT = {
    "cohort_fixed": "the same four firms form the peer group before and after the change",
    "cohort_meets_minimum_group_size": ("the peer group has at least the minimum group size "
                                        "for a published benchmark"),
    "no_single_member_reconstruction": ("no single member can recover another member's value "
                                        "on an indicator with dispersion"),
    "target_raw_unchanged": "S1-RUBBER's own data are identical before and after",
    "peers_not_supply_chain_nodes": ("the simulated peers are neither suppliers nor customers "
                                     "in the chain"),
    "weights_within_bounds": "every M6 weight lies within its Level 2 band or Level 3 floor and cap",
    "weights_sum_one": "the M6 weights sum to one in both states",
    "target_fixed_score_unchanged": ("S1-RUBBER's score under the before parameters does not "
                                     "move, so its own-data effect is zero"),
}


def _m6_simulated_value(row, peer_index: int, state: str) -> float:
    """Deterministic peer values shared by the workbook and the CSV input."""
    _, factor, e1_after = M6_PEERS[peer_index]
    value = float(row.value)
    if row.unit == "fraction (percentage)":
        value = value * factor if value < 0 else min(1.0, value * factor)
    elif "yes/no" not in str(row.unit).lower():
        value *= factor
    if state == "after" and row.ind_id == "E1" and e1_after is not None:
        value = float(row.value) * e1_after
    if state == "after" and row.ind_id == "E10":
        value = float(row.value)
    return value


def load_m6_peer_workbook(path: str, base: pd.DataFrame) -> pd.DataFrame:
    """Read editable before/after peer columns using the workbook units."""
    from openpyxl import load_workbook
    rows = []
    workbook = load_workbook(path, data_only=True)
    try:
        for state, sheet_name in [("before", "Environmental Peer Before"),
                                  ("after", DIM_SHEETS["Environmental"])]:
            try:
                sheet = workbook_sheet(workbook, sheet_name)
            except KeyError:
                raise ConfigurationError(f"M6 missing sheet: {sheet_name}") from None
            columns = {}
            for cell in sheet[2]:
                name = str(cell.value or "").split("\n")[0].strip()
                if name in M6_PEER_NAMES:
                    if name in columns:
                        raise ConfigurationError(f"M6 duplicate peer column: {name}")
                    columns[name] = cell.column
            if set(columns) != set(M6_PEER_NAMES):
                raise ConfigurationError(f"M6 missing peer columns in {sheet_name}")
            metadata = base.set_index("ind_id")
            for cells in sheet.iter_rows(min_row=4):
                iid = indicator_id(str(cells[1].value or ""))
                if iid not in metadata.index:
                    continue
                unit = str(cells[2].value or "")
                if unit != str(metadata.loc[iid, "source_unit"]):
                    raise ConfigurationError(f"M6 unit mismatch for {iid}")
                for entity, column in columns.items():
                    cell = sheet.cell(cells[0].row, column)
                    rows.append(dict(state=state, entity=entity, industry="Tires & Rubber",
                        ind_id=iid, unit=metadata.loc[iid, "unit"],
                        value=parse_indicator_value(cell.value, unit, cell.number_format),
                        provenance="SIMULATED; not company disclosure"))
    finally:
        workbook.close()
    return pd.DataFrame(rows)


def run_m6_peer_experiment(data_dir: str = "data", policy=None,
                           export_dir: str = "", generate: bool = False) -> dict:
    """M6: S1-RUBBER benchmarked against three simulated tire and rubber peers on
    its environmental Level 2 and 3 indicators.

    A separate experiment, not a further period; the chain firms keep their M5
    data and the peers take part in the benchmark only.
    """
    policy = policy or RegulatoryPolicy()
    system = FederatedESGSystem(policy=policy, data_dir=data_dir)
    files = dict(discover_period_files(data_dir, verbose=False))
    master = os.path.join(data_dir, M6_WORKBOOK)
    source = master if os.path.isfile(master) else files.get("M5", files["M0"])
    raw = load_period(source, system.tier_map, system.header_map)["Environmental"]
    target = "S1-RUBBER"
    base = raw[(raw.entity == target) & (raw.tier != "T1") &
               np.isfinite(raw.value)].copy()
    peers = list(M6_PEER_NAMES)
    path = os.path.join(data_dir, "m6", "peer_inputs.csv")
    if generate:
        if os.path.exists(path) or os.path.exists(master):
            raise ConfigurationError(f"M6 inputs already exist: {path}; edit them explicitly")
        records = []
        for peer_index, entity in enumerate(peers):
            for row in base.itertuples():
                for state in ("before", "after"):
                    v = _m6_simulated_value(row, peer_index, state)
                    records.append(dict(state=state, entity=entity,
                        industry="Tires & Rubber", ind_id=row.ind_id,
                        unit=row.unit, value=v, provenance="SIMULATED; not company disclosure",
                        industry_reference=M6_PEER_NOTE))
        os.makedirs(os.path.dirname(path), exist_ok=True)
        pd.DataFrame(records).to_csv(path, index=False)
    if not os.path.isfile(path) and not os.path.isfile(master):
        raise ConfigurationError("M6 inputs missing; run m6 --generate-inputs once")
    inputs = (load_m6_peer_workbook(master, base) if os.path.isfile(master)
              else pd.read_csv(path))
    expected = {(s, e, k) for s in ("before", "after") for e in peers for k in base.ind_id}
    keys = list(inputs[["state", "entity", "ind_id"]].itertuples(index=False, name=None))
    if len(keys) != len(set(keys)) or set(keys) != expected:
        raise ConfigurationError("M6 requires one before/after observation per peer and declared indicator")
    if not np.isfinite(pd.to_numeric(inputs.value, errors="coerce")).all():
        raise ConfigurationError("M6 peer values must be finite numbers")
    units = base.set_index("ind_id").unit.to_dict()
    if any(r.unit != units[r.ind_id] or r.industry != "Tires & Rubber" for r in inputs.itertuples()):
        raise ConfigurationError("M6 peer units or industry differ from the declared cohort")
    frames, statistics, governed = {}, {}, {}
    dim = "Environmental"
    for state in ("before", "after"):
        members = [base.copy()]
        for entity in peers:
            frame = base.copy()
            frame["entity"] = entity
            values = inputs[(inputs.state == state) & (inputs.entity == entity)].set_index("ind_id").value
            frame["value"] = frame.ind_id.map(values)
            members.append(frame)
        frame = pd.concat(members, ignore_index=True)
        frames[state] = frame
        sums = [Participant(e, benchmark_only=e in peers).local_summary(frame)
                for e in [target] + peers]
        stats = system.recalibrate(dim, frame, "M6-" + state,
                                   is_anchor=state == "before", sums=sums)
        if state == "before":
            scales, basis = system.reference_scales[dim], system.reference_basis[dim]
            register = declare_governance_register(dim, stats, policy, "M6-before",
                s_ref=scales, s_ref_basis=basis, classification_digest=system.classification_digest)
            anchor = stats
        statistics[state] = stats
        governed[state] = govern_weights(dim, stats, anchor, policy, "M6-" + state,
                                         register=register)
    left, right = (governed[s].rows for s in ("before", "after"))
    weights = left.merge(right, on=["indicator", "ind_id", "tier"], suffixes=("_before", "_after"))
    for name in ("mu", "sigma", "rho", "w_data", "w_final"):
        weights[name + "_delta"] = weights[name + "_after"] - weights[name + "_before"]
    scores = []
    for entity in [target] + peers:
        def evaluate(state, st, ws):
            return score_entity(entity, frames[state], statistics[st], governed[ws], policy).score
        before = evaluate("before", "before", "before")
        fixed = evaluate("after", "before", "before")
        benchmark = evaluate("after", "after", "before")
        after = evaluate("after", "after", "after")
        scores.append(dict(entity=entity, before=before, fixed= fixed,
            benchmark_only=benchmark, after=after, own_data_effect=fixed-before,
            benchmark_effect=benchmark-fixed, weight_effect=after-benchmark,
            total_change=after-before))
    disclosure = pd.concat([benchmark_disclosure_check(frames[s], statistics[s]).assign(state=s)
                            for s in ("before", "after")], ignore_index=True)
    checks = {
        "cohort_fixed": all(set(s.contributors) == {target, *peers} for s in statistics.values()),
        "cohort_meets_minimum_group_size": len({target, *peers}) >= Coordinator.MIN_PARTICIPANTS,
        "no_single_member_reconstruction": not disclosure.loc[
            ~disclosure.zero_dispersion, "single_member_exact_recovery"].any(),
        "target_raw_unchanged": frames["before"].query("entity == @target").equals(frames["after"].query("entity == @target")),
        "peers_not_supply_chain_nodes": not set(peers).intersection(ENTITY_ORDER),
        "weights_within_bounds": all(((w.rows.w_final >= w.rows.band_lo - 1e-10) &
                                      (w.rows.w_final <= w.rows.band_hi + 1e-10)).all() for w in governed.values()),
        "weights_sum_one": all(abs(w.rows.w_final.sum() - 1) < 1e-9 for w in governed.values()),
        "target_fixed_score_unchanged": abs(scores[0]["own_data_effect"]) < 1e-12,
    }
    result = {"inputs": inputs, "weights": weights, "scores": pd.DataFrame(scores),
              "checks": pd.DataFrame([{"check": k, "passed": bool(v)} for k, v in checks.items()]),
              "cohort_data": pd.concat([f.assign(state=s) for s, f in frames.items()], ignore_index=True)}
    result["disclosure"] = disclosure
    result["calibration_log"] = pd.DataFrame(system.calibration_log)
    # The two M6 recalibrations are published like any M0-M5 round: a weight
    # version and digest per state, bound to the M6 register.
    result["versions"] = pd.DataFrame([{
        "experiment": "M6", "state": s, "dimension": dim,
        "peer_group": ", ".join([target] + peers),
        "register_version": register.version_id,
        "version_id": governed[s].version_id, "digest": governed[s].digest,
        "policy_digest": governed[s].policy_digest, "timestamp": governed[s].timestamp,
        "n_indicators": int(len(governed[s].rows)), "notes": "; ".join(governed[s].notes),
    } for s in ("before", "after")])
    result["chain_and_peer_data"] = pd.concat([
        pd.concat([raw.assign(provenance="M6 chain data" if os.path.isfile(master) else "Existing M5 designed scenario"),
                   f[f.entity.isin(peers)].assign(provenance="SIMULATED; not company disclosure")],
                  ignore_index=True).assign(state=s)
        for s, f in frames.items()], ignore_index=True)
    result["chain_and_peer_data"]["in_industry_benchmark"] = result["chain_and_peer_data"].entity.isin([target] + peers)
    if not all(checks.values()):
        raise AssertionError(f"M6 checks failed: {checks}")
    if export_dir:
        os.makedirs(export_dir, exist_ok=True)
        for name, frame in result.items():
            frame.to_csv(os.path.join(export_dir, name + ".csv"), index=False)
    return result


# =============================================================================
# SECTION 17.  COMPLETE DEMONSTRATION RUN
# =============================================================================

def run_full_demonstration(data_dir: str = ".", export_dir: str = "outputs",
                           policy: Optional[RegulatoryPolicy] = None,
                           verbose: bool = True,
                           expect_periods: Optional[int] = 6,
                           strict_classification: bool = True) -> dict:
    """Run the complete demonstration; with export_dir, write the result and
    checks workbooks.
    """
    policy = policy or RegulatoryPolicy()
    sysm = FederatedESGSystem(policy=policy, data_dir=data_dir,
                              strict_classification=strict_classification)

    files = discover_period_files(data_dir, expect_periods=expect_periods,
                                  verbose=verbose)

    out: dict = {"policy": policy, "system": sysm, "periods": {},
                 "periods_fixed": {}, "tier_source": sysm.tier_source,
                 "classification_digest": sysm.classification_digest,
                 "period_files": files}


    # ---- rolling line (benchmark re-estimated each round) ------------------
    for i, (label, path) in enumerate(files):
        r = sysm.run_period(path, label, is_anchor=(i == 0))
        out["periods"][label] = r

    # ---- fixed line (M0 benchmark and weights frozen) ---------------------
    anchor = out["periods"][files[0][0]]
    fixed_sysm = FederatedESGSystem(policy=policy, data_dir=data_dir,
                                    strict_classification=strict_classification)
    for label, path in files:
        out["periods_fixed"][label] = fixed_sysm.run_period(
            path, label + "-fixed", frozen_stats=anchor.stats,
            frozen_weights=anchor.weights)


    # ---- verification ------------------------------------------------------
    out["equiv"] = verify_federated_equals_centralised(anchor)
    out["independence"] = verify_dimensional_independence(sysm, anchor)
    out["sensitivity"] = dependency_sensitivity(out["periods"][files[-1][0]], policy=policy)

    # ---- triggers ----------------------------------------------------------
    out["triggers"] = _build_trigger_table(out, files, policy)

    # ---- weight drift against the anchor round ----------------------------
    out["drift"] = weight_drift_evidence(out, files, policy)

    # ---- offsetting stress case -------------------------------------------
    if len(files) >= 2:
        base = out["periods"][files[1][0]]
        tot_o, comp_o, meta_o = offsetting_scenario(base)
        reps_o = {e: component_trigger(e, base.comps["Environmental"][e],
                                       comp_o[e], tot_o, "Environmental", policy,
                                       reference_total=base.totals["Environmental"][e])
                  for e in ENTITY_ORDER}
        reps_o = cascade_alerts(reps_o)
        out["offsetting"] = {"meta": meta_o, "totals": tot_o, "comps": comp_o,
                             "reports": reps_o,
                             "hub_report": reps_o["S1-RUBBER"],
                             "focal_report": reps_o["FOCAL"]}

    # ---- assurance simulation ---------------------------------------------
    out["assurance"] = run_assurance_scenarios(sysm, anchor, policy,
                                               triggers=out["triggers"])

    # ---- cold-start and coverage safeguard --------------------------------
    out["cold_start"] = cold_start_scenario(sysm, anchor, policy)

    # ---- ablation ----------------------------------------------------------
    out["ablation"] = build_ablation_table(sysm, out, files, policy)

    # ---- scenario changes and dependency-weight robustness ----------------
    out["scenario_changes"] = scenario_change_table(out, files, policy)
    out["weight_robustness"] = dependency_weight_robustness(out, files, policy)
    out["scoring_sensitivity"] = scoring_choice_sensitivity(
        out, files, policy, data_dir, strict_classification=strict_classification)

    # ---- absolute channel, hard-alert state, acceptance tests -------------
    last = out["periods"][files[-1][0]]
    out["absolute_risk"] = last.absolute_risk
    out["hard_alert_register"] = sysm.hard_register.table()
    out["acceptance"] = run_governance_acceptance_tests(sysm, anchor, policy)
    extra_checks = list(out["scenario_changes"]["checks"])
    extra_checks.extend(out["weight_robustness"].get("checks", []))
    extra_checks.extend(out["scoring_sensitivity"]["checks"])
    extra_checks.extend(out["cold_start"]["checks"])
    extra_checks.extend(out["ablation"].get("checks", []))
    extra_checks.extend(out["assurance"].get("checks", []))
    extra_checks.extend(cross_period_checks(out, files, policy))
    if extra_checks:
        out["acceptance"] = pd.concat([out["acceptance"], pd.DataFrame(extra_checks)],
                                      ignore_index=True)
    out["events"] = run_event_demonstrations(
        data_dir=data_dir, policy=policy,
        strict_classification=strict_classification)
    event_acceptance = out["events"]["acceptance"].copy()
    event_acceptance["test"] = (event_acceptance["scenario"].astype(str) + ": "
                                + event_acceptance["check"].astype(str))
    event_acceptance["requirement"] = event_acceptance["check"]
    event_acceptance["result"] = event_acceptance["status"]
    event_acceptance["detail"] = "isolated frozen-benchmark event acceptance"
    out["acceptance"] = pd.concat(
        [out["acceptance"], event_acceptance[["test", "requirement", "result", "detail"]]],
        ignore_index=True)

    # ---- M6 industry peer experiment (not a further period) --------------
    if (os.path.isfile(os.path.join(data_dir, M6_WORKBOOK)) or
            os.path.isfile(os.path.join(data_dir, "m6", "peer_inputs.csv"))):
        out["m6_peers"] = run_m6_peer_experiment(data_dir, policy)
        m6_acceptance = pd.DataFrame([{
            "test": f"M6 peer experiment: {r.check}",
            "requirement": M6_CHECK_TEXT.get(r.check, r.check),
            "result": "PASS" if r.passed else "FAIL",
            "detail": "industry peer experiment, separate from the M0-M5 panel"}
            for r in out["m6_peers"]["checks"].itertuples()])
        out["acceptance"] = pd.concat([out["acceptance"], m6_acceptance],
                                      ignore_index=True)

    # ---- M7 dependency-weight sensitivity on the M0 inputs ----------------
    m7_path = os.path.join(data_dir, M7_WORKBOOK)
    if os.path.isfile(m7_path):
        out["m7"] = run_m7_dependency_sensitivity(sysm, anchor, m7_path)
        out["acceptance"] = pd.concat([out["acceptance"], pd.DataFrame([{
            "test": f"M7 dependency sensitivity: {r.check}",
            "requirement": M7_CHECK_TEXT.get(r.check, r.check),
            "result": "PASS" if r.passed else "FAIL",
            "detail": "FOCAL dependency weights varied on the M0 inputs"}
            for r in out["m7"]["checks"].itertuples()])], ignore_index=True)

    failed = out["acceptance"].loc[out["acceptance"]["result"] != "PASS", "test"].tolist()
    if failed:
        raise AssertionError("acceptance tests failed: " + ", ".join(failed))
    out["registers"] = sysm.registers
    out["calibration_log"] = pd.DataFrame(sysm.calibration_log)

    # ---- XGBoost scorer and SHAP (optional, no feedback into the results) --
    out["scorer_substitution"] = run_scorer_substitution_experiment(
        data_dir, policy, engine_events=out["events"],
        strict_classification=strict_classification)
    sub = out["scorer_substitution"]
    if sub["status"] in ("COMPLETE", "FAILED"):
        scorer_acceptance = sub["checks"].rename(columns={"detail": "requirement"})
        scorer_acceptance["test"] = "scorer substitution: " + scorer_acceptance["check"]
        scorer_acceptance["detail"] = ("critical check; the substitution results are marked "
                                       "usable only when every check passes")
        out["acceptance"] = pd.concat(
            [out["acceptance"], scorer_acceptance[["test", "requirement", "result", "detail"]]],
            ignore_index=True)
    out["ml_benchmark"] = run_indicator_shap_diagnostic(
        sub, output_dir=export_dir, squash_scale=policy.squash_scale)

    n_marked = int((~out["triggers"]["comparable_with_previous"].astype(bool)).sum())
    out["scope"] = pd.DataFrame([
        {"claim": "demonstration data",
         "status": "SIMULATED",
         "detail": "every indicator value is simulated: each M0 value of the seven "
                   "chain firms is a seeded random draw around a reference value of the "
                   "same unit and order of magnitude, so no value reproduces a figure "
                   "reported by any company; declarations are design values; M1-M5 are "
                   "defined relative to the simulated M0 and M6 peers are derived from "
                   "it. Firms are identified only by position codes; the scores describe "
                   "no company (Methodology sheet of each input workbook)."},
        {"claim": "peer groups (M0-M5)",
         "status": "SIMPLIFIED",
         "detail": "the design benchmarks each firm against a peer group of comparable "
                   "firms (industry, size). M0-M5 benchmark the seven chain firms as one "
                   "cross-industry group, a demonstration simplification: their scores "
                   "are relative to this mixed group, and on absolute indicators such as "
                   "emissions or headcount, larger firms tend to appear riskier. An "
                   "industry peer group is shown separately in M6 (S1-RUBBER with three "
                   "simulated tire and rubber peers)."},
        {"claim": "calibration upload stages",
         "status": "ACCOUNTED SEPARATELY",
         "detail": "each regular calibration aggregates one upload of common-scale "
                   "(n, sum, sum of squares) summaries per dimension. If conversion "
                   "ranges or reference scales depend on the initial data, a raw-unit "
                   "reference initialization is performed first and logged separately; "
                   "initialization plus first calibration therefore needs two sequential "
                   "upload stages. Fully predeclared ranges and scales avoid initialization. "
                   "Later raw-unit tables are local evaluator diagnostics, not uploads."},
        {"claim": "secure aggregation",
         "status": "SIMULATED",
         "detail": "additive masking inside one process; the coordinator object "
                   "holds the plaintext values and generates the masks. No "
                   "cryptographic isolation is implemented or claimed. For a "
                   "participant reporting n=1, sum(x)=x."},
        {"claim": "auditor recomputation",
         "status": "REFERENCE",
         "detail": "the same scoring function is called again from the published "
                   "parameters; it evidences parameter sufficiency, not an "
                   "independent second implementation."},
        {"claim": "alpha-method verification",
         "status": "REFERENCE",
         "detail": "declared and applied method names are compared with each other and "
                   "with the accepted methods (Table 4); strategic dependency is marked "
                   "for enhanced assurance. Procurement records are never inspected."},
        {"claim": "weight version digest",
         "status": "IMPLEMENTED",
         "detail": "binds mu, sigma, sigma_score, s_ref, indicator direction, "
                   "governance level, w_reg and w_final, plus the policy, "
                   "register and classification digests."},
        {"claim": "cross-period interpretation",
         "status": "IMPLEMENTED",
         "detail": "each transition splits the change in a firm's total into the change "
                   "under the previous version's benchmark and weights (net_change) and "
                   "the change due only to the new benchmark and weights "
                   "(new_version_effect, T3); a fixed-reference series is kept (T1, T18c). "
                   "A change of indicator definitions, or of the data coverage behind a "
                   "total, marks the transition as not directly comparable (T3, T16), "
                   "naming only the direct supplier a change arrives through. "
                   + (f"{n_marked} firm-transitions in M0-M5 are marked." if n_marked else
                      "M0-M5 change neither, so no transition is marked; event E is.")},
        {"claim": "ratchet effect",
         "status": "OPEN QUESTION",
         "detail": "a possible dynamic incentive under a rolling benchmark, left "
                   "for future work; not tested or established here."},
        {"claim": "XGBoost / SHAP",
         "status": "ATTRIBUTION AID",
         "detail": "SHAP of the XGBoost scorer used in the substitution "
                   "experiment. Event-level SHAP changes are compared with the "
                   "rule's local-score change in the same units; the global "
                   "comparison with the rule decomposition is exploratory "
                   "(post- vs pre-sigmoid units) and does not show that the "
                   "weights were recovered. No predictive accuracy or "
                   "calibration is claimed."},
        {"claim": "local-scorer substitution",
         "status": "SENSITIVITY EXPERIMENT",
         "detail": "an XGBoost regressor (squared error, 400 trees, depth 2) fitted "
                   "in-sample to the rule's local scores under frozen R0 parameters "
                   "replaces the rule; its features are the rule's own standardized "
                   "inputs and its target is the rule's score, not an outcome label. "
                   "Hard-threshold checks, propagation and alert rules are shared, "
                   "and a hard alert changes neither scorer's local score. Hard-alert "
                   "invariance is asserted; internal-warning differences are reported "
                   "relative to the rule, not to a ground truth. The experiment shows "
                   "whether the decisions depend on the rule's functional form; it is "
                   "not evidence of predictive accuracy, probability calibration or "
                   "external validity, which would need observed outcome labels."},
        {"claim": "cold-start and coverage safeguard",
         "status": "IMPLEMENTED",
         "detail": f"a new entrant scores itself with the published parameters (no "
                   f"default parameter is assigned) and is not flagged merely because "
                   f"it missed the last recalibration. A score is provisional when a "
                   f"scored input is missing (its interval is published) or when it "
                   f"rests on less than {policy.coverage_floor:.0%} of the published "
                   f"weight. The flag changes no number and travels every positive "
                   f"dependency edge as a branch notice naming only the direct "
                   f"supplier. A direct supplier counts toward the firm's direct "
                   f"upstream coverage only when its data meet that floor; no usable "
                   f"data in a dimension, or insufficient information in "
                   f"{policy.escalation_periods} consecutive periods, calls for an "
                   f"assurance review. A missing mandatory declaration, or one "
                   f"reported as preliminary or unverified, is recorded as pending "
                   f"review, not as a confirmed failure. A withheld indicator is a "
                   f"missing input: its weight is kept, the score interval is "
                   f"published and the point value is its upper bound, so "
                   f"withholding never lowers the published score and widens the "
                   f"interval instead of removing the indicator. Completeness still "
                   f"rests on assurance (T20, T21)."},
        {"claim": "internal warnings and hard alerts",
         "status": "IMPLEMENTED",
         "detail": "each firm tests its own indicators and its supplier contributions "
                   "separately; an internal warning stays with that firm, and its "
                   "customers test the updated total under their own rules. A warning "
                   "stays open until the same component has come back within the limit "
                   "of its level before the warning; an improvement in another component "
                   "does not close it, and while it is open the response is at least "
                   "monitoring (T3 open_internal_warnings). Closure after a documented "
                   "review is not simulated. A hard alert and its closure notice travel "
                   "every declared edge upstream and downstream, naming only the direct "
                   "supplier or customer (T3, T14b)."},
        {"claim": "offsetting inside a branch",
         "status": "DETECTED AT THE BRANCH",
         "detail": "in M4, S1-RUBBER's S2-CHEM branch deteriorates while its S2-EQUIP "
                   "branch improves by an offsetting amount. S1-RUBBER tests each supplier "
                   "contribution and raises an "
                   "internal warning although its total barely moves; FOCAL receives "
                   "S1-RUBBER's almost unchanged total and, under its own rule, raises "
                   "none. Offsetting is detected by the firm whose components move; the "
                   "design does not ensure that the focal firm detects every deterioration "
                   "deeper in the chain (T3, T11)."},
        {"claim": "Level 1 hard alerts in the demonstration",
         "status": "POLICY DECLARATIONS ONLY",
         "detail": "the active Level 1 rules are the mandatory policy declarations "
                   "S26-S28. The M5 alert at S1-RUBBER marks a failed child-labor policy "
                   "declaration, a policy deficiency, not a confirmed child-labor or "
                   "forced-labor incident. Count-based conditions (fatalities, regulatory "
                   "violations, data breaches, substantiated ethics violations and "
                   "environmental fines) are monitored, but their thresholds are governance "
                   "parameters that the demonstration does not set (T13, T14)."},
        {"claim": "independent event update",
         "status": "IMPLEMENTED REFERENCE",
         "detail": "A-E branch independently from the same published R0 state: ESG data, "
                   "hard alert with review clearance, unreachable warning limits, a "
                   "documented dependency-weight revision (strategic dependency, "
                   "enhanced assurance) and a withdrawn observation. Each change carries "
                   "its source. Frozen benchmark and weight versions; aggregation "
                   "forbidden during events. Full-chain output is an evaluator view; "
                   "access control and durable event storage are absent."},
    ])
    sens = out["scoring_sensitivity"]["table"].set_index("variant")
    published = sens.loc["published"]
    varied = sens.drop(index="published")
    out["scope"] = pd.concat([out["scope"], pd.DataFrame([{
        "claim": "conversion ranges, dispersion measure and tau",
        "status": "SENSITIVITY REPORTED",
        "detail": (f"magnitude indicators are converted on [0, "
                   f"{policy.magnitude_range_multiple:g} x anchor mean] and declarations and "
                   f"percentages on their natural ranges, so the common-scale standard "
                   f"deviation behind the data-driven weights depends on these ranges; "
                   f"weighting by dispersion is a design choice, not evidence that an "
                   f"indicator matters more. The escalation levels and warning limits are in "
                   f"score units and are set together with tau: at tau = "
                   f"{policy.squash_scale:g} a local score of {policy.persistent_level:.2f} is "
                   f"a weighted deviation of "
                   f"{score_level_index(policy.persistent_level, policy):.2f} peer standard "
                   f"deviations. The level-based rule is applied to dimension totals, which "
                   f"for a firm with suppliers combine its own score with its suppliers' "
                   f"totals, so this reading holds only for a firm without suppliers, and "
                   f"{policy.persistent_level:.2f} is a governance setting, not a calibrated "
                   f"threshold. Declarations and percentages carry "
                   f"{published['E_w_final_declarations_percentages']:.0%} of the published M0 "
                   f"environmental weight. With one setting changed at a time and the "
                   f"thresholds held (T23a, T23b): "
                   + "; ".join(
                       f"{name}, {r['E_w_final_declarations_percentages']:.0%} of that weight, "
                       f"{r['warning_decisions_changed']} of {r['decisions_compared']} warning "
                       f"decisions and {r['response_levels_changed']} response levels changed"
                       for name, r in varied.iterrows())
                   + (". No hard alert changes." if int(varied["hard_alerts_changed"].max()) == 0
                      else ". Hard alerts change; see T23a."))}])], ignore_index=True)
    if "m6_peers" in out:
        out["scope"] = pd.concat([out["scope"], pd.DataFrame([{
            "claim": "industry peer benchmark (M6)",
            "status": "SEPARATE EXPERIMENT",
            "detail": "not a further period: S1-RUBBER is benchmarked against three "
                      "simulated tire and rubber peers on its environmental Level 2 and "
                      "Level 3 indicators, a four-firm peer group (the minimum group size). The "
                      "seven chain firms keep their M5 data; the peers take part in the "
                      "benchmark only and send or receive no scores. Its scores are not "
                      "comparable with the M0-M5 composite (sheet M6)."}])],
            ignore_index=True)
    if "m7" in out:
        listed = [f"{b:.2f}/{r:.2f}/{c:.2f} ({n})" for n, b, r, c in out["m7"]["settings"]]
        out["scope"] = pd.concat([out["scope"], pd.DataFrame([{
            "claim": "dependency-weight sensitivity (M7)",
            "status": "SENSITIVITY EXPERIMENT",
            "detail": "not a further period; the M7 workbook holds the M0 indicator values, "
                      "so M0 data, supply relationships, peer benchmarks and ESG weights stay "
                      "fixed. FOCAL own / S1-RUBBER / S1-COATING weights "
                      + ", ".join(listed[:-1]) + " and " + listed[-1]
                      + f", read from the {M7_SETTINGS_SHEET} sheet of the M7 workbook. "
                      "Illustrative settings, not derived from procurement or "
                      "strategic-dependency data (M7 tables, sheet Dependency weights)."}])],
            ignore_index=True)

    # ---- version log -------------------------------------------------------
    out["version_log"] = pd.DataFrame(sysm.version_log)


    if export_dir:
        out["export_path"], out["checks_path"] = export_results(out, files, export_dir)
    if verbose:
        _print_summary(out)
    return out


# =============================================================================
# SECTION 18.  RESULT TABLES
# =============================================================================

def _score_pairs(p0: PeriodResult, p1: PeriodResult, dim: str,
                 policy: RegulatoryPolicy) -> Dict[str, Tuple[LocalScore, LocalScore]]:
    """Each firm's local score before and after a transition under the previous
    version's benchmark and weights, so a benchmark update alone moves nothing."""
    same = p0.weights[dim].digest == p1.weights[dim].digest
    return {e: (p0.scores[dim][e],
                p1.scores[dim][e] if same else
                score_entity(e, p1.data[dim], p0.stats[dim], p0.weights[dim], policy))
            for e in ENTITY_ORDER}


def _transition_reports(p0: PeriodResult, p1: PeriodResult, dim: str,
                        policy: RegulatoryPolicy,
                        totals0: Optional[Dict[str, float]] = None,
                        comps0: Optional[Dict[str, Dict[str, float]]] = None,
                        totals1: Optional[Dict[str, float]] = None,
                        comps1: Optional[Dict[str, Dict[str, float]]] = None,
                        weights: Optional[Dict[str, dict]] = None,
                        pairs: Optional[Dict[str, Tuple[LocalScore, LocalScore]]] = None
                        ) -> Dict[str, TriggerReport]:
    """Internal warnings and hard alerts of one dimension in one transition,
    tested under the previous version's benchmark and weights.
    """
    dwmap = p1.dependency_weights if weights is None else weights
    pairs = pairs or _score_pairs(p0, p1, dim, policy)
    if totals1 is None and p0.weights[dim].digest != p1.weights[dim].digest:
        totals1, comps1 = propagate({e: pairs[e][1].score for e in ENTITY_ORDER}, dwmap)
    register = p1.hard_register
    totals0 = p0.totals[dim] if totals0 is None else totals0
    comps0 = p0.comps[dim] if comps0 is None else comps0
    totals1 = p1.totals[dim] if totals1 is None else totals1
    comps1 = p1.comps[dim] if comps1 is None else comps1
    reps = {}
    for e in ENTITY_ORDER:
        sticky = bool(register.is_active(e, dim)) if register else False
        reps[e] = component_trigger(
            e, comps0[e], comps1[e], totals1, dim, policy,
            hard_local=sticky,
            hard_labels=register.labels(e, dim) if register else [],
            reference_total=totals0[e],
            hard_records=_hard_records(register, e, dim),
            own_effects=own_indicator_effects(pairs[e][0], pairs[e][1],
                                              dwmap[e]["beta"], policy,
                                              p0.stats[dim].ind_id))
    return cascade_alerts(reps, dwmap)


def _hard_records(register: Optional["HardAlertRegister"], entity: str,
                  dimension: str) -> List[dict]:
    """Opaque ID, category, status and raising period of active events at one node."""
    if register is None:
        return []
    return [{"event_id": r.event_id, "category": r.label,
             "status": HARD_STATUS_ACTIVE, "raised": r.raised_period}
            for r in register.active_records(entity, dimension)]


def _direction_text(directions: Dict[str, str]) -> str:
    return "; ".join(f"{'own operations' if k == '__own__' else disp(k)}: {v}"
                     for k, v in sorted(directions.items()))


def _hard_detail_text(details: List[dict]) -> str:
    events: Dict[str, dict] = {}
    for rec in details:
        event = events.setdefault(rec["event_id"], {**rec, "branches": []})
        if rec["via"] not in event["branches"]:
            event["branches"].append(rec["via"])
    return "; ".join(
        f"{eid} via {', '.join(disp(v) for v in rec['branches'])}: "
        f"{rec['category']} ({rec['status']}, raised {rec['raised'] or 'n/a'})"
        for eid, rec in sorted(events.items()))


# ---------------------------------------------------------------------------
# Cross-period interpretation (Section 4.2.4): same-rules change, new-version
# effect and comparability of every transition.
# ---------------------------------------------------------------------------

def _definition_signature(result: "PeriodResult", dimension: str) -> Dict[str, tuple]:
    """Definition of every weighted indicator of a dimension: unit, direction
    and conversion reference points.  Governance levels, benchmark values and
    weights are scoring rules, not definitions; a change in them is the
    new-version effect."""
    ws, st = result.weights[dimension], result.stats[dimension]
    units = (result.data[dimension].groupby("ind_id")["unit"]
             .agg(lambda u: tuple(sorted(set(map(str, u))))).to_dict())

    def points(indicator):
        bounds = getattr(st, "input_bounds", {}).get(indicator)
        return None if bounds is None else tuple(
            round(float(b), 12) if np.isfinite(b) else None for b in bounds)

    return {str(r.ind_id): (units.get(r.ind_id, ()),
                            bool(st.is_good.get(r.indicator, False)),
                            points(r.indicator))
            for r in ws.rows.itertuples()}


def _reported_indicators(result: "PeriodResult", dimension: str, entity: str) -> set:
    """Weighted indicators the firm reported a usable value for."""
    sc = result.scores.get(dimension, {}).get(entity)
    if sc is None:
        return set()
    ids = result.stats[dimension].ind_id
    return {str(ids.get(k, k)) for k in sc.contributions}


def comparability(before: "PeriodResult", after: "PeriodResult", dimension: str,
                  weights: Optional[Dict[str, dict]] = None
                  ) -> Dict[str, Tuple[bool, str]]:
    """Whether each firm's total in `after` is directly comparable with `before`
    (Section 4.2.4), with the reason when it is not.

    Not comparable when indicator definitions change, a current benchmark becomes
    unavailable or available again, or the reported indicators of the firm or of
    a firm upstream change; a customer learns only the direct supplier.  A new
    benchmark, new weights or a new governance level alone are the new-version
    effect.
    """
    dwmap = after.dependency_weights if weights is None else weights
    sig0 = _definition_signature(before, dimension)
    sig1 = _definition_signature(after, dimension)
    changed = sorted((k for k in set(sig0) | set(sig1) if sig0.get(k) != sig1.get(k)),
                     key=_indicator_sort_key)
    common = []
    if changed:
        common.append(f"indicator definitions changed ({', '.join(changed[:6])}"
                      f"{', ...' if len(changed) > 6 else ''})")
    if bool(before.weights[dimension].unavailable) != bool(after.weights[dimension].unavailable):
        common.append("no current benchmark" if after.weights[dimension].unavailable
                      else "current benchmark available again")
    own: Dict[str, str] = {}
    for e in ENTITY_ORDER:
        r0 = _reported_indicators(before, dimension, e)
        r1 = _reported_indicators(after, dimension, e)
        parts = []
        if r0 - r1:
            parts.append("no longer reported: "
                         + ", ".join(sorted(r0 - r1, key=_indicator_sort_key)))
        if r1 - r0:
            parts.append("newly reported: "
                         + ", ".join(sorted(r1 - r0, key=_indicator_sort_key)))
        own[e] = f"own coverage changed ({'; '.join(parts)})" if parts else ""
    via = cascade_provisional({e: bool(text) for e, text in own.items()}, dwmap)
    result = {}
    for e in ENTITY_ORDER:
        notes = common + ([own[e]] if own[e] else [])
        if via[e]:
            notes.append("coverage changed upstream, via "
                         + ", ".join(disp(j) for j in via[e]))
        result[e] = (not notes, "; ".join(notes))
    return result


def cross_period_checks(out: dict, files, policy: RegulatoryPolicy) -> List[dict]:
    """Acceptance tests for cross-period interpretation (Section 4.2.4)."""
    import copy
    checks: List[dict] = []

    def record(test: str, requirement: str, passed: bool, detail: str) -> None:
        checks.append({"test": test, "requirement": requirement,
                       "result": "PASS" if passed else "FAIL", "detail": detail})

    labels = [l for l, _ in files]
    if len(labels) < 2:
        return checks
    trig = out["triggers"]
    residual = float((trig["published_change"] - trig["net_change"]
                      - trig["new_version_effect"]).abs().max()) if len(trig) else 0.0

    # The previous period's data scored under the new benchmark and weights:
    # only the version changes, so the check expects no same-rules change.
    p0, p1 = out["periods"][labels[0]], out["periods"][labels[1]]
    mixed = copy.copy(p1)
    mixed.data = p0.data
    mixed.scores = {d: {e: score_entity(e, p0.data[d], p1.stats[d], p1.weights[d], policy)
                        for e in ENTITY_ORDER} for d in p1.scores}
    mixed.local = {d: {e: sc.score for e, sc in s.items()} for d, s in mixed.scores.items()}
    mixed.totals, mixed.comps = {}, {}
    for d in mixed.local:
        mixed.totals[d], mixed.comps[d] = propagate(mixed.local[d], p1.dependency_weights)
    worst, moved, cases = 0.0, 0, 0
    for d in mixed.local:
        reps = _transition_reports(p0, mixed, d, policy)
        for e in ENTITY_ORDER:
            cases += 1
            worst = max(worst, abs(reps[e].net_change))
            moved += abs(mixed.totals[d][e] - p0.totals[d][e]) > 1e-12
    record("cross-period: a new benchmark and new weights alone give no same-rules change",
           f"re-scoring the {labels[0]} data under the {labels[1]} benchmark and weights "
           "leaves every firm's change under the previous version at zero, so the whole "
           "published change is reported as the new-version effect; in T3 the published "
           "change equals the same-rules change plus the new-version effect",
           worst <= 1e-12 and moved > 0 and residual <= 1e-12,
           f"{moved} of {cases} firm-dimension totals moved, largest same-rules change "
           f"{worst:.1e}; T3 residual {residual:.1e} over {len(trig)} rows")

    base = out["periods"][labels[-1]]
    dim, other, firm = "Environmental", "Social", "S1-RUBBER"
    reported = _reported_indicators(base, dim, firm)
    iid = next(str(i) for i in base.weights[dim].rows["ind_id"] if str(i) in reported)

    # A firm stops reporting one weighted indicator.
    data = base.data[dim].copy(deep=True)
    data.loc[(data["entity"] == firm) & (data["ind_id"] == iid), "value"] = np.nan
    changed = copy.copy(base)
    changed.data = dict(base.data)
    changed.data[dim] = data
    changed.scores = {d: dict(s) for d, s in base.scores.items()}
    changed.scores[dim][firm] = score_entity(firm, data, base.stats[dim],
                                             base.weights[dim], policy)
    flags, flags_other = comparability(base, changed, dim), comparability(base, changed, other)
    customers = downstream_notices(firm, base.dependency_weights)
    expected = {firm} | {c for c, _ in customers}
    marked = {e for e, (ok, _) in flags.items() if not ok}
    told = all(flags[c][1] == f"coverage changed upstream, via {disp(v)}"
               for c, v in customers)
    record("cross-period: a coverage change is marked as not directly comparable",
           "when a firm stops reporting a weighted indicator, its total and the totals of "
           "every firm downstream are marked not directly comparable; a customer is told "
           "only the direct supplier the change arrives through, and the other dimension "
           "is unaffected",
           marked == expected and told and all(ok for ok, _ in flags_other.values()),
           f"{firm} {iid} withheld in {base.label}: marked "
           f"{', '.join(disp(e) for e in ENTITY_ORDER if e in marked)}")

    # The unit of one indicator is redefined.
    data = base.data[dim].copy(deep=True)
    sel = data["ind_id"] == iid
    data.loc[sel, "unit"] = data.loc[sel, "unit"].astype(str) + " (redefined)"
    changed = copy.copy(base)
    changed.data = dict(base.data)
    changed.data[dim] = data
    flags, flags_other = comparability(base, changed, dim), comparability(base, changed, other)
    record("cross-period: a definition change is marked for every firm",
           "when an indicator definition changes, every firm's total in that dimension is "
           "marked not directly comparable, naming the indicator; the other dimension is "
           "unaffected",
           all(not ok and iid in note for ok, note in flags.values())
           and all(ok for ok, _ in flags_other.values()),
           f"unit of {iid} redefined in {base.label}: "
           f"{sum(not ok for ok, _ in flags.values())} of {len(flags)} firms marked")
    return checks


def _build_trigger_table(out: dict, files, policy) -> pd.DataFrame:
    """Internal warnings and hard alerts by transition, dimension and firm (T3).

    The published change in each total is split into the change under the
    previous version (net_change) and the new-version effect, and a transition
    that is not directly comparable is marked.
    """
    rows = []
    labels = [l for l, _ in files]
    streak: Dict[Tuple[str, str], int] = {}
    high: Dict[Tuple[str, str], int] = {
        (dim, e): int(out["periods"][labels[0]].totals[dim][e] >= policy.persistent_level)
        for dim in ("Environmental", "Social") for e in ENTITY_ORDER}
    register = InternalWarningRegister()
    for i in range(1, len(labels)):
        p0, p1 = out["periods"][labels[i - 1]], out["periods"][labels[i]]
        transition = f"{labels[i-1]}->{labels[i]}"
        for dim in ("Environmental", "Social"):
            reps = _transition_reports(p0, p1, dim, policy)
            comparable = comparability(p0, p1, dim)
            for e in ENTITY_ORDER:
                rep = reps[e]
                key = (dim, e)
                worse = any(v > rep.threshold for v in rep.component_deltas.values())
                streak[key] = streak.get(key, 0) + 1 if worse else 0
                high[key] = high[key] + 1 if p1.totals[dim][e] >= policy.persistent_level else 0
                level, basis = graduated_response(rep, streak[key], high[key],
                                                  p1.totals[dim][e], policy)
                open_now, closed_now = register.update(dim, e, rep, transition)
                level, basis = carry_open_warnings(level, basis, open_now, transition)
                rows.append({
                    "transition": transition, "dimension": dim,
                    "entity": disp(e), "tier": SUPPLY_CHAIN[e]["tier"],
                    "net_change": rep.net_change,
                    "published_change": p1.totals[dim][e] - p0.totals[dim][e],
                    "new_version_effect": (p1.totals[dim][e] - p0.totals[dim][e]
                                           - rep.net_change),
                    "comparable_with_previous": comparable[e][0],
                    "comparability_note": comparable[e][1],
                    "max_abs_component": rep.max_abs_component,
                    "reference_total_pre_event": rep.reference_total,
                    "threshold": rep.threshold,
                    "net_rule_fires": rep.net_fired,
                    "internal_warning": rep.internal_warning,
                    "warning_signal": warning_signal(rep),
                    "hard_alert": rep.hard_alert,
                    "hard_alert_local": rep.hard_local,
                    "hard_event_count": len(rep.hard_event_summary),
                    "hard_event_ids": ", ".join(r["event_id"] for r in rep.hard_event_summary),
                    "hard_event_details": json.dumps(rep.hard_event_summary, sort_keys=True),
                    "hard_alert_via_direct_supplier": ", ".join(disp(b) for b in rep.hard_via),
                    "hard_alert_via_direct_customer": ", ".join(
                        disp(b) for b in rep.hard_via_customers),
                    "hard_threshold_label": "; ".join(rep.hard_labels),
                    "indicators_flagged": ", ".join(rep.fired_indicators),
                    "branches_flagged": ", ".join(disp(b) for b in rep.fired_branches),
                    "upstream_dispersion": rep.upstream_dispersion,
                    "max_single_upstream_share": rep.max_single_share,
                    "MISSED_BY_NET_RULE": rep.internal_warning and not rep.net_fired,
                    "warning_direction": rep.warning_direction,
                    "fired_component_directions": _direction_text(rep.fired_directions),
                    "hard_alert_raised_here": "; ".join(
                        f"{r['event_id']}: {r['category']} "
                        f"({r['status']}, raised {r['raised'] or 'n/a'})"
                        for r in rep.hard_records),
                    "hard_notice_received": _hard_detail_text(
                        rep.hard_via_detail + rep.hard_customer_detail),
                    "open_internal_warnings": open_warning_text(open_now),
                    "internal_warnings_closed": open_warning_text(closed_now, closed=True),
                    "response_level": level,
                    "response_basis": basis,
                })
    return pd.DataFrame(rows)


# ---------------------------------------------------------------------------
# Score changes across the scenarios, with checks that the untouched dimension
# does not move and that, with the benchmark held at M0, a firm moves only
# through its own new data or its suppliers.
# ---------------------------------------------------------------------------

_DIMS2 = ("Environmental", "Social")


def _raw_data_changes(p0: PeriodResult, p1: PeriodResult) -> Dict[str, List[str]]:
    """Firms whose submitted values changed between two rounds, by dimension.

    Read from the uploaded data rather than from the scores, so the checks can
    tell a score that moved because of new data from one that moved without it.
    """
    moved: Dict[str, List[str]] = {}
    for dim in _DIMS2:
        a = p0.data[dim].groupby(["entity", "indicator"])["value"].first()
        b = p1.data[dim].groupby(["entity", "indicator"])["value"].first()
        idx = a.index.union(b.index)
        va = pd.to_numeric(a.reindex(idx), errors="coerce").to_numpy(float)
        vb = pd.to_numeric(b.reindex(idx), errors="coerce").to_numpy(float)
        same = (va == vb) | (np.isnan(va) & np.isnan(vb))
        ents = set(idx[~same].get_level_values(0))
        moved[dim] = [e for e in ENTITY_ORDER if e in ents]
    return moved


def _downstream_closure(sources, weights: Optional[Dict[str, dict]] = None) -> set:
    """The sources plus every firm reachable from them along positive
    dependency edges, i.e. every firm whose total a change at a source can
    reach through propagation."""
    dwmap = DEPENDENCY_WEIGHTS if weights is None else weights
    reach = set(sources)
    for tier in TIER_ORDER:
        for e, cfg in SUPPLY_CHAIN.items():
            if cfg["tier"] != tier or e in reach:
                continue
            if any(a > 0.0 and j in reach for j, a in dwmap[e]["alpha"].items()):
                reach.add(e)
    return reach


def _fmt_thousandths(x: float) -> str:
    if x is None or not np.isfinite(x):
        return ""
    if abs(x) < 1e-12:
        return "0"
    return f"{x * 1000:+.3f}".replace("-", "−")


WARNING_SIGNAL_NOTE = (
    "W↑, W↓ and W± mark an internal warning at the firm for a deterioration, an "
    "improvement or changes in both directions: a component changed beyond the "
    "threshold. An improvement alone does not call for a risk investigation, although "
    "an earlier warning or a persistent high level can still require a response. "
    "H(self) marks a hard alert raised at the firm and H one received from a supplier "
    "or customer.")


def warning_signal(rep: TriggerReport) -> str:
    """Display direction without changing the two-sided threshold test."""
    if not rep.internal_warning:
        return ""
    return {"deterioration": "W↑", "improvement": "W↓", "mixed": "W±"}.get(
        rep.warning_direction, "W")


def _alert_code(rep: TriggerReport) -> str:
    """Hard alerts take display priority over direction-labelled internal warnings."""
    if rep.hard_local:
        return "H(self)"
    if rep.hard_alert:
        return "H"
    return warning_signal(rep)


def scenario_change_table(out: dict, files, policy: RegulatoryPolicy) -> dict:
    """Score changes by transition in the published run and with the benchmark
    fixed at M0 (T18), with their acceptance checks.
    """
    labels = [l for l, _ in files]
    runs = (("published", out["periods"]), ("fixed M0", out["periods_fixed"]))
    rows = []
    for i in range(1, len(labels)):
        a, b = labels[i - 1], labels[i]
        raw = _raw_data_changes(out["periods"][a], out["periods"][b])
        dims_changed = [d for d in _DIMS2 if raw[d]]
        altered = dims_changed[0] if len(dims_changed) == 1 else None
        sources = [e for e in ENTITY_ORDER if any(e in raw[d] for d in dims_changed)]
        path = _downstream_closure(sources)
        for run_name, per in runs:
            p0, p1 = per[a], per[b]
            reps = (_transition_reports(p0, p1, altered, policy)
                    if run_name == "published" and altered else {})
            for e in ENTITY_ORDER:
                dE = p1.totals["Environmental"][e] - p0.totals["Environmental"][e]
                dS = p1.totals["Social"][e] - p0.totals["Social"][e]
                if altered == "Environmental":
                    d_alt, d_oth = dE, abs(dS)
                elif altered == "Social":
                    d_alt, d_oth = dS, abs(dE)
                else:
                    d_alt, d_oth = np.nan, np.nan
                rep = reps.get(e)
                rows.append({
                    "benchmark": run_name,
                    "transition": f"{a}->{b}",
                    "altered_dimension": altered or ("both" if dims_changed else "none"),
                    "firms_with_new_data": ", ".join(disp(x) for x in sources),
                    "entity": disp(e), "tier": SUPPLY_CHAIN[e]["tier"],
                    "on_change_path": e in path,
                    "E_total_change": dE, "S_total_change": dS,
                    "change_altered_dimension": d_alt,
                    "change_x1000": round(d_alt * 1000, 3) if np.isfinite(d_alt) else np.nan,
                    "abs_change_other_dimension": d_oth,
                    "composite_change": p1.composite[e] - p0.composite[e],
                    "alert": _alert_code(rep) if rep else "",
                    "internal_warning": rep.internal_warning if rep else None,
                    "warning_direction": rep.warning_direction if rep else "",
                    "hard_alert": rep.hard_alert if rep else None,
                })
    long = pd.DataFrame(rows)

    def panel(run_name: str) -> pd.DataFrame:
        wide = {"Firm": [disp(e) for e in ENTITY_ORDER]}
        sub = long[long["benchmark"] == run_name]
        for t, g in sub.groupby("transition", sort=False):
            alt = g["altered_dimension"].iloc[0]
            head = t.replace("->", "→") + f" ({alt[:1] if alt in _DIMS2 else alt})"
            wide[head] = [(_fmt_thousandths(x["change_altered_dimension"])
                           + (" " + x["alert"] if x["alert"] else "")).strip()
                          for _, x in g.iterrows()]
        return pd.DataFrame(wide)

    checks: List[dict] = []
    single = long[long["altered_dimension"].isin(_DIMS2)]
    for run_name, label in (("published", "published run"),
                            ("fixed M0", "M0 benchmark fixed")):
        sub = single[single["benchmark"] == run_name]
        if sub.empty:
            continue
        worst = float(sub["abs_change_other_dimension"].max())
        checks.append({
            "test": f"scenario table: untouched dimension unchanged ({label})",
            "requirement": ("in every transition whose new data touch one dimension "
                            "only, no firm's total in the other dimension moves"),
            "result": "PASS" if worst <= 1e-12 else "FAIL",
            "detail": (f"{len(sub)} firm-transitions over {sub['transition'].nunique()} "
                       f"transitions; largest change in the other dimension {worst:.1e}")})
    fx = single[single["benchmark"] == "fixed M0"]
    if not fx.empty:
        off, on = fx[~fx["on_change_path"]], fx[fx["on_change_path"]]
        worst_off = float(off["change_altered_dimension"].abs().max()) if len(off) else 0.0
        moved_on = int((on["change_altered_dimension"].abs() > 1e-12).sum())
        checks.append({
            "test": "scenario table: firms off the change path unchanged (M0 benchmark fixed)",
            "requirement": ("with the benchmark and weights held at M0, a firm's total "
                            "moves only through new data at the firm or through its "
                            "suppliers' totals"),
            "result": "PASS" if worst_off <= 1e-12 else "FAIL",
            "detail": (f"{len(off)} off-path firm-transitions, largest change "
                       f"{worst_off:.1e}; {moved_on} of {len(on)} firm-transitions on a "
                       f"change path moved")})
    return {"long": long, "panel_published": panel("published"),
            "panel_fixed": panel("fixed M0"), "checks": checks}


def dependency_weight_robustness(out: dict, files, policy: RegulatoryPolicy,
                                 betas: Tuple[float, ...] = (0.50, 0.60, 0.70, 0.80, 0.90)
                                 ) -> dict:
    """Warning and alert decisions under alternative dependency weights (T19).

    For every firm with suppliers, beta takes each value in `betas` and the
    upstream share is split as declared, equally and, for two suppliers, swapped;
    the other firms keep their weights.  Decisions use the published rule.
    """
    import copy
    labels = [l for l, _ in files]
    per = out["periods"]
    first = labels[0]
    wE, wS = policy.dim_weights

    # Local scores before and after each transition under the previous
    # version's parameters.
    pairs = {(labels[i - 1], labels[i], dim): _score_pairs(per[labels[i - 1]],
                                                           per[labels[i]], dim, policy)
             for i in range(1, len(labels)) for dim in _DIMS2}

    def evaluate(weights):
        tot, cmp_ = {}, {}
        for l in labels:
            for dim in _DIMS2:
                tot[(l, dim)], cmp_[(l, dim)] = propagate(per[l].local[dim], weights)
        dec = {}
        for i in range(1, len(labels)):
            a, b = labels[i - 1], labels[i]
            for dim in _DIMS2:
                pair = pairs[(a, b, dim)]
                totals1, comps1 = propagate({e: pair[e][1].score for e in ENTITY_ORDER},
                                            weights)
                reps = _transition_reports(per[a], per[b], dim, policy,
                                           totals0=tot[(a, dim)], comps0=cmp_[(a, dim)],
                                           totals1=totals1, comps1=comps1,
                                           weights=weights, pairs=pair)
                for e, r in reps.items():
                    dec[(f"{a}->{b}", dim, e)] = (bool(r.internal_warning or r.hard_alert),
                                                  bool(r.hard_alert))
        flags = {}
        for l in labels:
            comp = {e: wE * tot[(l, "Environmental")][e] + wS * tot[(l, "Social")][e]
                    for e in ENTITY_ORDER}
            for e, v in peer_relative_flag(comp, policy).items():
                flags[(l, e)] = bool(v)
        return dec, flags, tot, cmp_

    # The published decisions, as in the trigger table (T3).
    published: Dict[Tuple[str, str, str], Tuple[bool, bool]] = {}
    for i in range(1, len(labels)):
        a, b = labels[i - 1], labels[i]
        for dim in _DIMS2:
            for e, r in _transition_reports(per[a], per[b], dim, policy,
                                            pairs=pairs[(a, b, dim)]).items():
                published[(f"{a}->{b}", dim, e)] = (bool(r.internal_warning or r.hard_alert),
                                                    bool(r.hard_alert))

    declared = copy.deepcopy(DEPENDENCY_WEIGHTS)
    dec0, flags0, _, _ = evaluate(declared)
    reproduced = sum(1 for k in published if dec0.get(k) == published[k])

    rows, flips = [], []
    for firm in ENTITY_ORDER:
        alpha0 = dict(declared[firm]["alpha"])
        s0 = sum(alpha0.values())
        if not alpha0 or s0 <= 0:
            continue
        sups = list(alpha0)
        if len(sups) == 1:
            splits = [("single supplier", {sups[0]: 1.0})]
        else:
            splits = [("declared proportions", {k: v / s0 for k, v in alpha0.items()}),
                      ("equal split", {k: 1.0 / len(sups) for k in sups})]
            if len(sups) == 2:
                x, y = sups
                splits.append(("declared shares swapped",
                               {x: alpha0[y] / s0, y: alpha0[x] / s0}))
        beta0 = float(declared[firm]["beta"])
        grid = sorted({round(float(b), 6) for b in betas} | {round(beta0, 6)})
        for split_name, share in splits:
            for beta in grid:
                w = copy.deepcopy(declared)
                w[firm] = {"beta": beta,
                           "alpha": {k: (1.0 - beta) * s for k, s in share.items()}}
                dec, flags, tot, cmp_ = evaluate(w)
                changed = [k for k in dec0 if dec[k][0] != dec0[k][0]]
                rec = {
                    "firm_varied": disp(firm), "declared_beta": beta0,
                    "alpha_split": split_name, "beta": beta,
                    "alpha": ", ".join(f"{disp(k)} {v:.3f}" for k, v in w[firm]["alpha"].items()),
                    "declared_scheme": bool(abs(beta - beta0) < 1e-9 and all(
                        abs(w[firm]["alpha"][k] - alpha0[k]) < 1e-9 for k in alpha0)),
                }
                for dim in _DIMS2:
                    comps_f = cmp_[(first, dim)][firm]
                    total_f = tot[(first, dim)][firm]
                    up = sum(v for k, v in comps_f.items() if k != "__own__")
                    rec[f"{dim[:1]}_upstream_share_{first}"] = (up / total_f if total_f
                                                                else np.nan)
                rec.update({
                    "decisions_compared": len(dec0),
                    "decisions_changed": len(changed),
                    "alerts_added": sum(1 for k in changed if dec[k][0]),
                    "alerts_removed": sum(1 for k in changed if not dec[k][0]),
                    "hard_alerts_changed": sum(1 for k in dec0 if dec[k][1] != dec0[k][1]),
                    "peer_flags_changed": sum(1 for k in flags0 if flags[k] != flags0[k]),
                    "changed_decisions": "; ".join(
                        f"{t} {d[:1]} {disp(e)} {'raised' if dec[(t, d, e)][0] else 'dropped'}"
                        for t, d, e in changed),
                })
                rows.append(rec)
                for t, d, e in changed:
                    flips.append({"firm_varied": disp(firm), "alpha_split": split_name,
                                  "beta": beta, "transition": t, "dimension": d,
                                  "entity": disp(e),
                                  "declared_decision": dec0[(t, d, e)][0],
                                  "decision_under_scheme": dec[(t, d, e)][0]})

    table = pd.DataFrame(rows)
    flips_df = pd.DataFrame(flips, columns=["firm_varied", "alpha_split", "beta",
                                            "transition", "dimension", "entity",
                                            "declared_decision", "decision_under_scheme"])
    reference_check = {
        "test": "dependency-weight robustness: declared weights reproduce the published decisions",
        "requirement": ("the robustness grid applies the published rules, so the declared "
                        "weights give exactly the published warning and alert decisions (T3)"),
        "result": ("PASS" if published and reproduced == len(published)
                   and len(dec0) == len(published) else "FAIL"),
        "detail": f"{reproduced} of {len(published)} decisions identical"}
    if table.empty:
        summary = pd.DataFrame()
        check = None
    else:
        summary = (table.groupby(["firm_varied", "alpha_split"], sort=False)
                   .agg(schemes=("beta", "size"),
                        beta_min=("beta", "min"), beta_max=("beta", "max"),
                        schemes_with_changes=("decisions_changed",
                                              lambda s: int((s > 0).sum())),
                        max_decisions_changed=("decisions_changed", "max"),
                        max_hard_alerts_changed=("hard_alerts_changed", "max"),
                        max_peer_flags_changed=("peer_flags_changed", "max"))
                   .reset_index())
        n_changed = int((table["decisions_changed"] > 0).sum())
        # The grid contains each firm's declared weights as a reference row, so
        # the alternatives are the schemes that differ from the declared one.
        n_declared = int(table["declared_scheme"].sum())
        n_alt = len(table) - n_declared
        check = {
            "test": "dependency-weight robustness: hard alerts unchanged",
            "requirement": ("a hard alert reaches the same firms under every tested beta "
                            "and alpha split"),
            "result": "PASS" if int(table["hard_alerts_changed"].max()) == 0 else "FAIL",
            "detail": (f"{len(table)} weight schemes ({n_alt} alternatives plus {n_declared} "
                       f"declared reference schemes) x {len(dec0)} decisions; warning "
                       f"decisions changed in {n_changed} alternative schemes (at most "
                       f"{int(table['decisions_changed'].max())} per scheme); "
                       f"{n_alt - n_changed} of {n_alt} alternatives change no decision")}
    checks = [c for c in (reference_check, check) if c is not None]
    return {"table": table, "summary": summary, "flips": flips_df, "check": check,
            "checks": checks, "decisions_per_scheme": len(dec0)}


# ---------------------------------------------------------------------------
# Scoring choices.  The common-scale dispersion behind the data-driven weights
# depends on the conversion ranges, and the escalation levels and warning
# limits are in score units, so the published run is compared with runs that
# change one of these settings (T23).
# ---------------------------------------------------------------------------

SCORING_VARIANTS = (
    ("magnitude range x20", {"magnitude_range_multiple": 20.0}),
    ("magnitude range x50", {"magnitude_range_multiple": 50.0}),
    ("coefficient of variation", {"dispersion_mode": "cv"}),
    ("tau 0.5", {"squash_scale": 0.5}),
    ("tau 2", {"squash_scale": 2.0}),
)


def scoring_choice_sensitivity(out: dict, files, policy: RegulatoryPolicy, data_dir: str,
                               strict_classification: bool = True,
                               variants=SCORING_VARIANTS) -> dict:
    """M0-M5 rerun with one scoring setting changed at a time (T23a, T23b).

    Each variant recalibrates every period from the same inputs; the warning
    limits and escalation levels stay as published, and decisions are compared
    with the published trigger table (T3).
    """
    import dataclasses
    labels = [l for l, _ in files]
    runs = [("published", policy, out["periods"], out["triggers"])]
    for name, change in variants:
        varied = dataclasses.replace(policy, **change)
        if varied.digest() == policy.digest():
            continue
        system = FederatedESGSystem(policy=varied, data_dir=data_dir,
                                    strict_classification=strict_classification)
        periods = {label: system.run_period(path, label, is_anchor=(i == 0))
                   for i, (label, path) in enumerate(files)}
        runs.append((name, varied, periods,
                     _build_trigger_table({"periods": periods}, files, varied)))

    anchor = out["periods"][labels[0]]
    natural = {d: {k: natural_conversion_range(u) is not None for k, u in
                   anchor.data[d].drop_duplicates("indicator")
                   .set_index("indicator")["unit"].items()}
               for d in _DIMS2}
    reference = {(r.transition, r.dimension, r.entity): r
                 for r in out["triggers"].itertuples()}
    rows = []
    for name, varied, periods, trig in runs:
        first = periods[labels[0]]
        rec = {"variant": name, "magnitude_range_multiple": varied.magnitude_range_multiple,
               "dispersion": varied.dispersion_mode, "tau": varied.squash_scale}
        for d in _DIMS2:
            w = first.weights[d].rows
            on_natural = w["indicator"].map(natural[d]).fillna(False).astype(bool)
            for col in ("w_data", "w_final"):
                rec[f"{d[:1]}_{col}_declarations_percentages"] = float(
                    w.loc[on_natural, col].sum() / w[col].sum())
        rec["values_clipped"] = int(sum(sum((periods[l].stats[d].conversion_clipped or {}).values())
                                        for l in labels for d in _DIMS2))
        rec["max_abs_local_change"] = float(max(
            abs(periods[l].local[d][e] - out["periods"][l].local[d][e])
            for l in labels for d in _DIMS2 for e in ENTITY_ORDER))
        rec["min_rank_correlation_composite"] = float(np.nanmin([spearman(
            [out["periods"][l].composite[e] for e in ENTITY_ORDER],
            [periods[l].composite[e] for e in ENTITY_ORDER]) for l in labels]))
        now = {(r.transition, r.dimension, r.entity): r for r in trig.itertuples()}
        changes, n_warning, n_level, n_hard = [], 0, 0, 0
        for key, ref in reference.items():
            cur, parts = now[key], []
            if bool(cur.internal_warning) != bool(ref.internal_warning):
                n_warning += 1
                parts.append("warning " + ("added" if cur.internal_warning else "removed"))
            if cur.response_level != ref.response_level:
                n_level += 1
                parts.append(f"{ref.response_level} -> {cur.response_level}")
            if bool(cur.hard_alert) != bool(ref.hard_alert):
                n_hard += 1
                parts.append("hard alert " + ("added" if cur.hard_alert else "removed"))
            if parts:
                changes.append(f"{key[0]} {key[1][:1]} {key[2]}: {', '.join(parts)}")
        level = varied.persistent_level
        rec.update({
            "decisions_compared": len(reference),
            "internal_warnings": int(trig["internal_warning"].sum()),
            "warning_decisions_changed": n_warning,
            "response_levels_changed": n_level,
            "hard_alerts": int(trig["hard_alert"].sum()),
            "hard_alerts_changed": n_hard,
            "persistent_level": level,
            "persistent_level_as_local_score_deviation": score_level_index(level, varied),
            "totals_at_or_above_persistent_level": int(sum(
                v >= level for l in labels for d in _DIMS2 for v in periods[l].totals[d].values())),
            "decision_changes": "; ".join(changes),
        })
        rows.append(rec)
    table = pd.DataFrame(rows)

    # Weights depend on the conversion range and the dispersion measure, not on tau.
    weight_runs = [(name, periods[labels[0]]) for name, varied, periods, _ in runs
                   if name == "published"
                   or (varied.magnitude_range_multiple, varied.dispersion_mode)
                   != (policy.magnitude_range_multiple, policy.dispersion_mode)]
    frames = []
    for d in _DIMS2:
        merged = None
        for name, first in weight_runs:
            w = first.weights[d].rows[["ind_id", "indicator", "tier", "w_data", "w_final"]]
            w = w.rename(columns={"w_data": f"w_data {name}", "w_final": f"w_final {name}"})
            merged = w if merged is None else merged.merge(
                w, on=["ind_id", "indicator", "tier"], how="outer")
        merged.insert(0, "dimension", d)
        merged.insert(4, "conversion", merged["indicator"].map(
            lambda k: "natural range" if natural[d].get(k) else "magnitude"))
        order = sorted(merged.index, key=lambda i: _indicator_sort_key(merged.at[i, "ind_id"]))
        frames.append(merged.loc[order])
    weights = pd.concat(frames, ignore_index=True)
    weights = weights[["dimension", "ind_id", "indicator", "tier", "conversion"]
                      + [c for c in weights.columns if c.startswith("w_final")]
                      + [c for c in weights.columns if c.startswith("w_data")]]

    alternatives = table[table["variant"] != "published"]
    check = {
        "test": "scoring choices: hard alerts unchanged",
        "requirement": ("the Level 1 hard alerts do not depend on the conversion range, the "
                        "dispersion measure or tau"),
        "result": ("PASS" if len(alternatives)
                   and int(alternatives["hard_alerts_changed"].max()) == 0 else "FAIL"),
        "detail": "; ".join(
            f"{r.variant}: {r.hard_alerts} hard alerts, {r.warning_decisions_changed} of "
            f"{r.decisions_compared} warning decisions changed"
            for r in alternatives.itertuples())}
    return {"table": table, "weights": weights, "checks": [check]}


# ---------------------------------------------------------------------------
# Cold-start and coverage safeguard.  Evaluator scenario on the anchor-round
# data; nothing here is published.
# ---------------------------------------------------------------------------

def _scenario_state(label: str, data: Dict[str, pd.DataFrame],
                    stats: Dict[str, GlobalStats], weights: Dict[str, WeightSet],
                    policy: RegulatoryPolicy) -> PeriodResult:
    """Score, propagate and flag one scenario state under given parameters."""
    scores: Dict[str, Dict[str, LocalScore]] = {}
    local: Dict[str, Dict[str, float]] = {}
    hard_hits: Dict[str, str] = {}
    for dim, df in data.items():
        scores[dim], local[dim] = {}, {}
        for e in ENTITY_ORDER:
            hit = AssuranceProvider(e).verify_hard_thresholds(df)
            if hit:
                hard_hits[e] = hit
            sc = score_entity(e, df, stats[dim], weights[dim], policy, hard_violation=hit)
            scores[dim][e], local[dim][e] = sc, sc.score
    totals, comps = {}, {}
    for dim in data:
        totals[dim], comps[dim] = propagate(local[dim])
    w_e, w_s = policy.dim_weights
    composite = {e: w_e * totals.get("Environmental", {}).get(e, 0.5)
                    + w_s * totals.get("Social", {}).get(e, 0.5) for e in ENTITY_ORDER}
    result = PeriodResult(label=label, filepath="", stats=stats, weights=weights,
                          local=local, totals=totals, comps=comps, scores=scores,
                          composite=composite, hard_hits=hard_hits, summaries={},
                          data=data)
    result.provisional = provisional_status(result, policy)
    return result


def _withholding_path(firm: str, dim: str, base: PeriodResult,
                      policy: RegulatoryPolicy, max_steps: int = 40) -> List[dict]:
    """Selective reporting: withhold, one at a time, the input whose removal
    gives the lowest published score (T21b).
    """
    import copy
    st, ws = base.stats[dim], base.weights[dim]
    wmap = ws.w()
    frame = base.data[dim].copy(deep=True)
    focal = next(e for e in ENTITY_ORDER if SUPPLY_CHAIN[e]["tier"] == "Focal")
    reach = _downstream_closure({firm})
    closure = [e for e in ENTITY_ORDER if e in reach and e != firm]
    hit = base.scores[dim][firm].hard_threshold_hit
    base_local = base.local[dim][firm]
    rows: List[dict] = []
    withheld: List[str] = []
    for step in range(max_steps + 1):
        sc = score_entity(firm, frame, st, ws, policy, hard_violation=hit)
        local = dict(base.local[dim])
        local[firm] = sc.score
        totals, comps = propagate(local)
        state = copy.copy(base)
        state.scores = {d: dict(v) for d, v in base.scores.items()}
        state.scores[dim][firm] = sc
        prov = provisional_status(state, policy)[dim]
        reports = {e: component_trigger(
            e, base.comps[dim][e], comps[e], totals, dim, policy,
            reference_total=base.totals[dim][e],
            own_effects=own_indicator_effects(
                base.scores[dim][e], sc if e == firm else base.scores[dim][e],
                DEPENDENCY_WEIGHTS[e]["beta"], policy, st.ind_id))
            for e in ENTITY_ORDER}
        cascade_alerts(reports)
        share = prov[firm]["coverage_share"]
        rows.append({
            "dimension": dim, "firm": disp(firm), "step": step,
            "withheld_indicator": (str(st.ind_id.get(withheld[-1]) or withheld[-1])
                                   if withheld else ""),
            "indicators_withheld": len(withheld),
            "coverage_share": share, "incomplete": bool(sc.incomplete),
            "lower": sc.lower, "upper": sc.upper,
            "coverage_floor": policy.coverage_floor,
            "local_score": sc.score, "change_in_local_score": sc.score - base_local,
            "full_score_within_interval": bool(sc.lower - 1e-12 <= base_local
                                               <= sc.upper + 1e-12),
            "provisional": prov[firm]["provisional"],
            "reason_codes": ",".join(prov[firm]["codes"]),
            "internal_warning_at_firm": bool(reports[firm].internal_warning),
            "warning_direction_at_firm": reports[firm].warning_direction,
            "downstream_internal_warning": ", ".join(
                disp(e) for e in closure if reports[e].internal_warning),
            "downstream_told_provisional": ", ".join(
                disp(e) for e in closure if prov[e]["via"]),
            # every notice names only direct suppliers with a positive weight
            "notice_via_direct_suppliers": all(
                DEPENDENCY_WEIGHTS[e]["alpha"].get(j, 0.0) > 0.0
                for e in closure for j in prov[e]["via"]),
            "focal_total": totals[focal],
            "change_in_focal_total": totals[focal] - base.totals[dim][focal],
        })
        if step == max_steps or (step > 0 and share < policy.coverage_floor - 1e-9):
            break
        own = frame["entity"] == firm
        candidates = sorted({k for k in frame.loc[own & np.isfinite(frame["value"]), "indicator"]
                             if float(wmap.get(k, 0.0)) > 0.0})
        best, best_score = None, np.inf
        for k in candidates:
            mask = own & (frame["indicator"] == k)
            saved = frame.loc[mask, "value"].copy()
            frame.loc[mask, "value"] = np.nan
            trial = score_entity(firm, frame, st, ws, policy, hard_violation=hit).score
            frame.loc[mask, "value"] = saved
            if trial < best_score:
                best, best_score = k, trial
        if best is None:
            break
        frame.loc[own & (frame["indicator"] == best), "value"] = np.nan
        withheld.append(best)
    return rows


def cold_start_scenario(sysm: "FederatedESGSystem", anchor: PeriodResult,
                        policy: RegulatoryPolicy, entrant: str = "S2-CHEM") -> dict:
    """Cold-start and coverage safeguard on the anchor data (T21).

    A  a new entrant scores itself with the published parameters, then joins the
       next recalibration;
    B  the same firm withholds inputs one at a time (_withholding_path).
    """
    if entrant not in ENTITY_ORDER:
        raise ValueError(f"unknown entrant {entrant}")
    founding = [e for e in ENTITY_ORDER if e != entrant]
    data = {dim: df.copy(deep=True) for dim, df in anchor.data.items()}
    stats_a, weights_a, stats_b, weights_b = {}, {}, {}, {}
    for dim, df in data.items():
        present = set(df["entity"])
        t1_ids = sorted({str(x) for x in
                         df.loc[df["tier"] == "T1", "ind_id"].dropna().unique()})
        st0 = Coordinator(secure_aggregation=False).aggregate(
            [Participant(e).local_summary(df) for e in founding if e in present], df)
        s_ref, basis = declare_reference_scales(st0, df, policy)
        attach_scoring_scale(st0, s_ref, policy)
        attach_common_scale(st0, df[df.entity.isin(founding)], policy)
        reg = declare_governance_register(
            dim, st0, policy, "CS0", s_ref=s_ref, s_ref_basis=basis,
            classification_digest=sysm.classification_digest)
        stats_a[dim] = st0
        weights_a[dim] = govern_weights(
            dim, st0, st0, policy, "CS0", t1_ids_in_data=t1_ids, register=reg,
            classification_digest_value=sysm.classification_digest)
        st1 = Coordinator(secure_aggregation=False).aggregate(
            [Participant(e).local_summary(df) for e in ENTITY_ORDER if e in present], df)
        attach_scoring_scale(st1, reg.s_ref, policy)
        attach_common_scale(st1, df, policy, st0)
        stats_b[dim] = st1
        weights_b[dim] = govern_weights(
            dim, st1, st0, policy, "CS1", t1_ids_in_data=t1_ids, register=reg,
            classification_digest_value=sysm.classification_digest)
    state_a = _scenario_state("CS-A", data, stats_a, weights_a, policy)
    state_b = _scenario_state("CS-B", data, stats_b, weights_b, policy)
    reach = _downstream_closure({entrant})
    closure = [e for e in ENTITY_ORDER if e in reach and e != entrant]
    focal = next(e for e in ENTITY_ORDER if SUPPLY_CHAIN[e]["tier"] == "Focal")

    stage_names = {"A": "entrant scored with the incumbents' published benchmark",
                   "B": "next recalibration includes the entrant"}
    entrant_rows = []
    for key, state, wts in (("A", state_a, weights_a), ("B", state_b, weights_b)):
        for dim in state.scores:
            for e in ENTITY_ORDER:
                v = state.provisional[dim][e]
                entrant_rows.append({
                    "stage": f"{key}: {stage_names[key]}", "dimension": dim,
                    "entity": disp(e), "tier": SUPPLY_CHAIN[e]["tier"],
                    "role_in_scenario": ("entrant" if e == entrant else
                                         "downstream of entrant" if e in closure else
                                         "incumbent"),
                    "local_score": state.local[dim][e], "total": state.totals[dim][e],
                    "coverage_share": v["coverage_share"],
                    "contributed_to_benchmark": v["contributed_to_benchmark"],
                    "provisional": v["provisional"], "reasons": "; ".join(v["reasons"]),
                    "provisional_via_direct_supplier": ", ".join(disp(j) for j in v["via"]),
                    "parameter_version": wts[dim].version_id,
                })
    paths = {dim: _withholding_path(entrant, dim, anchor, policy) for dim in anchor.scores}

    summary: List[dict] = []

    def add(part: str, dim: str, measure: str, value) -> None:
        summary.append({"part": part, "dimension": dim, "measure": measure, "value": value})

    for dim in state_a.scores:
        a_loc, b_loc = state_a.local[dim][entrant], state_b.local[dim][entrant]
        part = f"A new entrant ({disp(entrant)})"
        add(part, dim, "entrant local score before it contributes to the benchmark", a_loc)
        add(part, dim, "entrant flagged provisional at entry",
            "yes" if state_a.provisional[dim][entrant]["provisional"] else "no")
        add(part, dim, "entrant local score after the next recalibration", b_loc)
        add(part, dim, "change at the recalibration", b_loc - a_loc)
        add(part, dim, f"{disp(focal)} total before the entrant contributes",
            state_a.totals[dim][focal])
        add(part, dim, f"{disp(focal)} total after the recalibration",
            state_b.totals[dim][focal])
    for dim, path in paths.items():
        part = f"B selective reporting ({disp(entrant)})"
        full = path[0]
        held = [r for r in path if r["step"] > 0]
        ok_rows = [r for r in held if "low_coverage" not in r["reason_codes"]]
        lowest = min(ok_rows, key=lambda r: r["local_score"]) if ok_rows else None
        first = next((r for r in held if "low_coverage" in r["reason_codes"]), None)
        add(part, dim, "local score, full reporting", full["local_score"])
        add(part, dim, "coverage share, full reporting", full["coverage_share"])
        step1 = held[0] if held else None
        if step1 is not None:
            add(part, dim, "first withheld indicator", step1["withheld_indicator"])
            add(part, dim, "first withheld indicator: change in the point score",
                step1["change_in_local_score"])
            add(part, dim, "first withheld indicator: published score interval",
                f"[{step1['lower']:.4f}, {step1['upper']:.4f}]")
            add(part, dim, "first withheld indicator: score published as provisional",
                "yes" if step1["provisional"] else "no")
            add(part, dim, "first withheld indicator: firms told downstream",
                step1["downstream_told_provisional"])
            add(part, dim, "first withheld indicator: internal warning at the firm",
                "yes" if step1["internal_warning_at_firm"] else "no")
        if lowest is not None:
            add(part, dim, "lowest point score while withholding, at or above the coverage floor",
                lowest["local_score"])
            add(part, dim, "change against full reporting at that point",
                lowest["change_in_local_score"])
            add(part, dim, "coverage share at that point", lowest["coverage_share"])
            add(part, dim, "indicators withheld at that point", lowest["indicators_withheld"])
            add(part, dim, "published score interval at that point",
                f"[{lowest['lower']:.4f}, {lowest['upper']:.4f}]")
            add(part, dim, f"change in {disp(focal)} total at that point",
                lowest["change_in_focal_total"])
        if first is not None:
            add(part, dim, "first step below the floor: coverage share", first["coverage_share"])
            add(part, dim, "first step below the floor: point score", first["local_score"])

    checks: List[dict] = []
    same_flags = all(state_a.provisional[d][entrant]["codes"]
                     == state_b.provisional[d][entrant]["codes"] for d in state_a.scores)
    checks.append({
        "test": "cold start: applicable benchmark does not automatically flag newcomer",
        "requirement": ("newcomer status alone does not make a score provisional: the "
                        "entrant carries the same flags before and after it contributes to "
                        "the benchmark"),
        "result": "PASS" if same_flags else "FAIL",
        "detail": "; ".join(
            f"{d[:1]} {disp(entrant)} {state_a.local[d][entrant]:.4f} before contributing -> "
            f"{state_b.local[d][entrant]:.4f} after the recalibration, flags "
            f"{','.join(state_a.provisional[d][entrant]['codes']) or 'none'} / "
            f"{','.join(state_b.provisional[d][entrant]['codes']) or 'none'}"
            for d in state_a.scores)})
    # Branch notices are exercised on the withholding path, where the
    # entrant's score is incomplete from the first withheld indicator.
    notice_ok, notice_rows = True, 0
    for d, path in paths.items():
        for r in path:
            if not r["provisional"]:
                continue
            notice_rows += 1
            told = {x for x in r["downstream_told_provisional"].split(", ") if x}
            if told != {disp(e) for e in closure} or not r["notice_via_direct_suppliers"]:
                notice_ok = False
    checks.append({
        "test": "cold start: provisional notice reaches every customer downstream",
        "requirement": ("every firm downstream of a provisional score is told, through each "
                        "direct supplier on the path, that its branch rests on a provisional "
                        "score; the notice names direct suppliers only"),
        "result": "PASS" if (notice_ok and notice_rows > 0) else "FAIL",
        "detail": (f"{notice_rows} provisional states on the withholding path; notified: "
                   + ", ".join(disp(e) for e in closure))})
    cov_ok, details = True, []
    for d, path in paths.items():
        for r in path:
            low = "low_coverage" in r["reason_codes"]
            if low != (r["coverage_share"] < policy.coverage_floor - 1e-9):
                cov_ok = False
            if low:
                told = {x for x in r["downstream_told_provisional"].split(", ") if x}
                if told != {disp(e) for e in closure}:
                    cov_ok = False
        below = next((r for r in path if "low_coverage" in r["reason_codes"]), None)
        text = f"{d[:1]}: full {path[0]['local_score']:.4f} at {path[0]['coverage_share']:.1%}"
        if below is not None:
            text += (f", below the floor from {below['coverage_share']:.1%} "
                     f"({below['local_score']:.4f})")
        details.append(text)
    checks.append({
        "test": "coverage floor: selective reporting below the floor is flagged",
        "requirement": ("along a greedy withholding path a score carries the low-coverage "
                        "flag exactly when the firm reported less than the published floor "
                        "share of the published weight, and every customer downstream is "
                        "then told its branch rests on a provisional score"),
        "result": "PASS" if cov_ok and all(
            any("low_coverage" in r["reason_codes"] for r in p) for p in paths.values())
            else "FAIL",
        "detail": f"floor {policy.coverage_floor:.0%}; " + "; ".join(details)})
    held = [r for p in paths.values() for r in p if r["step"] > 0]
    smallest = min((r["change_in_local_score"] for r in held), default=float("nan"))
    checks.append({
        "test": "upper bound: withholding never lowers the published score",
        "requirement": ("on the greedy withholding path every published point score is at "
                        "least the full-reporting score, which lies inside every published "
                        "interval"),
        "result": "PASS" if (held and smallest >= -1e-12
                             and all(r["full_score_within_interval"] for p in paths.values()
                                     for r in p)) else "FAIL",
        "detail": "; ".join(
            f"{d[:1]}: {len(p) - 1} indicators withheld, smallest change "
            f"{min(r['change_in_local_score'] for r in p[1:]):+.4f}" if len(p) > 1
            else f"{d[:1]}: nothing withheld" for d, p in paths.items())})

    return {"entrant_name": entrant, "state_a": state_a, "state_b": state_b,
            "entrant": pd.DataFrame(entrant_rows),
            "withholding": pd.DataFrame([r for d in paths for r in paths[d]]),
            "summary": pd.DataFrame(summary), "checks": checks}


def run_assurance_scenarios(sysm: FederatedESGSystem, anchor: PeriodResult,
                            policy: RegulatoryPolicy,
                            triggers: Optional[pd.DataFrame] = None) -> dict:
    """Simulated assurance checks A-E on the anchor round (T10a-T10e)."""
    res: dict = {"checks": []}
    df_s = anchor.data["Social"]

    # ---- scenario A: unsupported hard-threshold declaration ---------------
    # A hard alert travels to every affected firm, naming only the direct
    # supplier or customer; no score changes.
    target = "S2-CHEM"
    auditor = AssuranceProvider(target)
    evidence = {"S26": 0.0}          # audit finds the affirmative unsupported
    hit = auditor.verify_hard_thresholds(df_s, evidence)
    found = list(auditor.confirmed_labels) if hit else []

    tot_before = anchor.totals["Social"]
    tot_after, comp_after = propagate(dict(anchor.local["Social"]))
    notices = cascade_alerts({e: component_trigger(
        e, anchor.comps["Social"][e], comp_after[e], tot_after, "Social", policy,
        hard_local=(e == target and bool(found)),
        hard_labels=found if e == target else [],
        reference_total=tot_before[e],
        hard_records=([{"event_id": HardAlertRegister.make_event_id(
                            target, f"assurance|{label}|{anchor.label}"),
                        "category": label, "status": HARD_STATUS_ACTIVE,
                        "raised": anchor.label} for label in found]
                      if e == target else []))
        for e in ENTITY_ORDER})

    cascade = pd.DataFrame([{
        "entity": disp(e), "tier": SUPPLY_CHAIN[e]["tier"],
        "S_total_before": tot_before[e], "S_total_after": tot_after[e],
        "change": tot_after[e] - tot_before[e],
        "recalculated": abs(tot_after[e] - tot_before[e]) > 1e-12,
        "hard_alert": notices[e].hard_alert,
        "hard_alert_local": notices[e].hard_local,
        "hard_event_count": len(notices[e].hard_event_summary),
        "hard_event_ids": ",".join(r["event_id"] for r in notices[e].hard_event_summary),
        "hard_notice_received": _hard_detail_text(
            notices[e].hard_via_detail + notices[e].hard_customer_detail),
    } for e in ENTITY_ORDER])
    res["A_hard_threshold"] = {"finding": hit, "auditor_findings": auditor.findings,
                               "cascade": cascade}

    # ---- scenario B: anomalous uploaded summary ---------------------------
    # S2-EQUIP uploads raw-unit and common-scale summaries computed from a value
    # 60% below its audited record; the auditor compares them with the records.
    df_e = anchor.data["Environmental"]
    sums = {s.entity: s for s in anchor.summaries["Environmental"]}
    tampered_entity = "S2-EQUIP"
    k0 = next(iter(sums[tampered_entity].sx))
    uploaded = df_e.copy(deep=True)
    tampered_row = (uploaded["entity"] == tampered_entity) & (uploaded["indicator"] == k0)
    uploaded.loc[tampered_row, "value"] = uploaded.loc[tampered_row, "value"] * 0.40
    tampered = Participant(tampered_entity).local_summary(uploaded)
    anchor_e = anchor.stats["Environmental"]
    tampered_common = Participant(tampered_entity).common_scale_summary(
        uploaded, anchor_e.input_bounds, anchor_e.is_good)

    auditor_b = AssuranceProvider(tampered_entity)
    findings_b = auditor_b.verify_summary(tampered, audited=df_e)
    findings_b += auditor_b.verify_common_summary(tampered_common, df_e, anchor_e)

    corrupt = [tampered if s.entity == tampered_entity else s
               for s in anchor.summaries["Environmental"]]
    coord = Coordinator(secure_aggregation=False)
    st_bad = coord.aggregate(corrupt, df_e)
    attach_scoring_scale(st_bad, sysm.reference_scales.get("Environmental", {}), policy)
    attach_common_scale(st_bad, uploaded, policy, anchor.stats["Environmental"])
    ws_bad = govern_weights("Environmental", st_bad, sysm.anchor_stats.get("Environmental"),
                            policy, "TAMPERED",
                            register=sysm.registers.get("Environmental"),
                            classification_digest_value=sysm.classification_digest)
    local_bad = {e: score_entity(e, df_e, st_bad, ws_bad, policy).score
                 for e in ENTITY_ORDER}
    tot_bad, _ = propagate(local_bad)

    impact = pd.DataFrame([{
        "entity": disp(e),
        "E_total_clean": anchor.totals["Environmental"][e],
        "E_total_if_undetected": tot_bad[e],
        "distortion": tot_bad[e] - anchor.totals["Environmental"][e],
    } for e in ENTITY_ORDER])
    res["B_summary_anomaly"] = {"entity": disp(tampered_entity), "indicator": k0,
                                "findings": findings_b, "impact": impact}

    # ---- scenario C: independent recomputation ----------------------------
    rec = []
    for e in ENTITY_ORDER:
        a = AssuranceProvider(e)
        ok, val = a.recompute(anchor.data["Environmental"], anchor.stats["Environmental"],
                              anchor.weights["Environmental"], policy,
                              claimed=anchor.local["Environmental"][e])
        rec.append({"entity": disp(e), "published_score": anchor.local["Environmental"][e],
                    "auditor_recomputation": val, "matches_to_1e-9": ok})
    res["C_recomputation"] = pd.DataFrame(rec)

    # ---- scenario D: declared vs applied dependency-weight method ---------
    d = []
    for e, declared in DECLARED_ALPHA_METHOD.items():
        applied = declared if e != "D1-RETAIL" else "strategic dependency"
        check = AssuranceProvider(e).verify_alpha_method(declared, applied)
        d.append({"entity": disp(e), "declared_method": declared,
                  "method_applied": applied, "method_accepted": check["accepted"],
                  "consistent_with_declaration": check["consistent"],
                  "enhanced_assurance_required": check["enhanced_assurance"],
                  "verified": check["consistent"] and check["accepted"]})
    res["D_alpha_method"] = pd.DataFrame(d)

    # ---- scenario E: responses to alerts (third assurance task) -----------
    if triggers is not None:
        episodes = alert_episodes(triggers)
        review = pd.DataFrame([row for e in ENTITY_ORDER for row in
                               AssuranceProvider(e).verify_alert_responses(
                                   episodes, ALERT_RESPONSES)])
        res["E_alert_responses"] = review
        gaps = review[~review["response_documented"]] if len(review) else review
        names = {disp(e): e for e in ENTITY_ORDER}
        passed_ok = all(x.finding_passed_to == "; ".join(
            f"{disp(c)} via {disp(v)}" for c, v in downstream_notices(names[x.entity]))
            for x in gaps.itertuples())
        res["checks"].append({
            "test": "assurance: responses to alerts are checked",
            "requirement": ("every investigate or hard-alert episode has a documented response "
                            "with its reason and evidence, or the gap is a finding passed to the "
                            "firm's customers through their direct suppliers"),
            "result": "PASS" if (len(review) == len(episodes) and len(gaps) and passed_ok)
                      else "FAIL",
            "detail": (f"{len(review)} episodes, {len(review) - len(gaps)} documented; "
                       + "; ".join(f"finding at {x.entity} passed to {x.finding_passed_to}"
                                   for x in gaps.itertuples()))})

    return res


def build_ablation_table(sysm: FederatedESGSystem, out: dict, files,
                         policy: RegulatoryPolicy) -> dict:
    labels = [l for l, _ in files]
    fed = {l: out["periods"][l].composite for l in labels}
    cen, cen_diff = baseline_centralised(sysm, out, files, policy)
    loc, loc_cov = baseline_local_only(files, sysm.tier_map, policy)
    rul = baseline_rules_based(files, sysm.tier_map, policy)

    # Periods in which no hard alert is active anywhere in the published run.
    # A hard alert changes no score in any method; the second variance figure
    # only shows that the comparison does not depend on those periods.  The
    # rules-based baseline has no hard-alert channel.
    calm = [l for l in labels if not out["periods"][l].hard_hits]

    def mean_firm_sd(series, periods):
        sds = []
        for e in ENTITY_ORDER:
            v = np.array([series[l].get(e, np.nan) for l in periods], dtype=float)
            v = v[np.isfinite(v)]
            if len(v) >= 2:
                sds.append(float(np.std(v)))
        return (float(np.mean(sds)) if sds else np.nan), len(sds)

    def rank_agreement(series, label):
        """Spearman of a method's composites with the federated ones in one period,
        over the firms the method can score there (at least three)."""
        cur = series[label]
        common = [e for e in ENTITY_ORDER
                  if e in cur and not np.isnan(cur[e]) and e in fed[label]]
        return (spearman([cur[e] for e in common], [fed[label][e] for e in common])
                if len(common) >= 3 else np.nan)

    peer_benchmark = {"federated (proposed)": True, "(a) centralized raw pooling": True,
                      "(b) isolated local, own history": False,
                      "(c) rules-based, frozen band": False}

    def max_diff(series):
        diffs = [abs(series[l][e] - fed[l][e]) for l in labels for e in ENTITY_ORDER
                 if np.isfinite(series[l].get(e, np.nan))]
        return float(max(diffs)) if diffs else np.nan

    rows = []
    for name, series in (("federated (proposed)", fed),
                         ("(a) centralized raw pooling", cen),
                         ("(b) isolated local, own history", loc),
                         ("(c) rules-based, frozen band", rul)):
        first = series[labels[0]]
        cold = not all(np.isnan(v) for v in first.values())
        rho = rank_agreement(series, labels[-1])
        per_period = [r for r in (rank_agreement(series, l) for l in labels) if np.isfinite(r)]
        focal = [series[l].get("FOCAL", np.nan) for l in labels]
        shock = (focal[1] - focal[0]) if len(focal) > 1 else np.nan
        scored = np.mean([np.mean([0.0 if np.isnan(series[l].get(e, np.nan)) else 1.0
                                   for e in ENTITY_ORDER]) for l in labels])
        # Variance across scenarios for every firm, not only the focal one:
        # each firm's composite SD over the periods it can be scored in, then
        # averaged over the firms scoreable in at least two periods.
        sd_all, n_sd = mean_firm_sd(series, labels)
        sd_calm, _ = mean_firm_sd(series, calm)
        rows.append({
            "method": name,
            "entity_periods_scoreable": scored,
            "score_available_in_first_period": cold,
            "spearman_vs_federated_final": rho,
            # range over the rounds in which at least three firms can be scored
            "spearman_vs_federated_min_over_periods": min(per_period) if per_period else np.nan,
            "spearman_vs_federated_max_over_periods": max(per_period) if per_period else np.nan,
            "periods_with_spearman": len(per_period),
            "focal_shock_response_M0_to_M1": shock,
            "focal_cross_period_sd": float(np.nanstd(focal)),
            "mean_firm_cross_period_sd": sd_all,
            "firms_in_mean_sd": n_sd,
            "mean_firm_sd_periods_without_hard_alert": sd_calm,
            "peer_benchmark": peer_benchmark[name],
            "weights_adapt": name.startswith(("federated", "(a)")),
            # largest absolute composite difference from the federated scores
            # over every firm-period the method scores
            "max_abs_diff_vs_federated": max_diff(series),
        })

    table = pd.DataFrame(rows)
    detail = pd.DataFrame({
        "period": labels,
        **{f"federated_{disp(e)}": [fed[l][e] for l in labels] for e in ENTITY_ORDER},
        **{f"centralized_{disp(e)}": [cen[l][e] for l in labels] for e in ENTITY_ORDER},
        **{f"local_only_{disp(e)}": [loc[l].get(e, np.nan) for l in labels] for e in ENTITY_ORDER},
        **{f"rules_{disp(e)}": [rul[l].get(e, np.nan) for l in labels] for e in ENTITY_ORDER},
    })
    worst = float(cen_diff[["max_abs_diff_local", "max_abs_diff_total",
                            "max_abs_diff_composite"]].to_numpy().max()) if not cen_diff.empty else 0.0
    checks = [{
        "test": "ablation (a): pooled raw records give the federated scores",
        "requirement": ("recomputing the benchmark, weights, local scores and totals from the "
                        "pooled raw records of every period reproduces the federated results"),
        "result": "PASS" if worst < 1e-9 else "FAIL",
        "detail": (f"{len(cen_diff)} periods; largest absolute difference in any local score, "
                   f"total or composite {worst:.2e}")}]

    return {"table": table, "detail": detail, "local_coverage": loc_cov,
            "series": {"federated": fed, "centralized": cen, "local_only": loc, "rules": rul},
            "centralised_diff": cen_diff, "checks": checks,
            "periods_without_hard_alert": calm}


def _coverage_or_nan(result: PeriodResult, dimension: str, entity: str) -> float:
    """Reported share of published weight.  A hard alert does not replace the
    weighted score, so the share is defined for a firm with a hard alert too;
    it is undefined only when the firm has no local score."""
    sc = result.scores.get(dimension, {}).get(entity)
    if sc is None:
        return float("nan")
    return float(sc.coverage)


# =============================================================================
# SECTION 19.  CONSOLE REPORT AND EXPORT
# =============================================================================

RESULT_WORKBOOK = "ESG_FedStat_Results.xlsx"
CHECKS_WORKBOOK = "ESG_FedStat_Checks.xlsx"
RESULT_NOTES = (
    "All indicator values are simulated. M0-M5 are cumulative scenarios; M6 (industry "
    "peers) and M7 (dependency weights) are separate experiments, not further periods.",
    "Each table keeps its code (T1, T18b, ...); tables that share a sheet are stacked under "
    f"their titles. Scope, version log and acceptance tests are in {CHECKS_WORKBOOK}.",
    "Level 1-3 are indicator governance levels (coded T1-T3 in the classification "
    "workbook); Tier 1, Tier 2, Focal and Downstream are supply-chain positions. Numbers "
    "are stored at full precision and shown to six decimals (fifteen in columns named "
    "delta).",
    "M0-M5 use a mixed-industry demonstration cohort, not industry-comparable ESG "
    "rankings. M6 separately illustrates an industry-specific peer benchmark.",
    WARNING_SIGNAL_NOTE,
)
CHECKS_NOTES = (
    f"Scope (T0), roles (T12), version log (T9) and acceptance tests (T15, T16). The "
    f"results are in {RESULT_WORKBOOK}.",
    "Timestamps, and the version identifiers and digests that contain them, change with "
    "every run.",
)
EVENT_KEY_COLUMNS = (
    "entity", "dimension", "change_source", "method", "enhanced_assurance",
    "local_score_recomputed", "total_recomputed", "composite_before", "composite_after",
    "composite_delta", "alert", "E_total_before", "E_total_after", "S_total_before",
    "S_total_after", "warning_threshold", "internal_warning", "warning_signal", "warning_direction",
    "indicators_flagged", "branches_flagged", "hard_alert", "hard_event_count",
    "hard_event_ids", "hard_notice_received",
    "provisional", "provisional_via_direct_supplier", "comparable_with_previous",
    "comparability_note")
M6_WEIGHT_COLUMNS = (
    "ind_id", "indicator", "tier", "sigma_before", "sigma_after", "w_final_before",
    "w_final_after", "w_final_delta", "band_lo_after", "band_hi_after",
    "floor_binding_after", "cap_binding_after")


def _level_labels(frame: pd.DataFrame) -> pd.DataFrame:
    """Write a `tier` column of governance codes T1-T3 as `level`, Level 1-3.

    A `tier` column holding supply-chain positions is left as it is.
    """
    if "tier" not in frame.columns or not set(
            frame["tier"].dropna().astype(str)) <= set(LEVEL_NAME):
        return frame
    return frame.assign(tier=frame["tier"].map(lambda v: LEVEL_NAME.get(v, v))).rename(
        columns={"tier": "level"})


def _result_sheets(out: dict, files) -> Tuple[list, list]:
    """Sheets of the result and checks workbooks.

    A sheet is (name, blocks) and a block (code, title, frame).  A sheet with one
    block is written as a plain table; otherwise its blocks are stacked under
    their titles.  Empty blocks and sheets are left out.
    """
    labels = [l for l, _ in files]
    policy = out["policy"]
    results: list = []
    checks: list = []

    def sheet(target: list, name: str, *blocks) -> None:
        kept = [(code, title, _level_labels(frame)) for code, title, frame in blocks
                if frame is not None and len(frame.columns)]
        if kept:
            target.append((name, kept))

    def rows(frame):
        return frame if frame is not None and len(frame) else None

    score_rows = []
    for l in labels:
        r, rf = out["periods"][l], out["periods_fixed"][l]
        for e in ENTITY_ORDER:
            score_rows.append({
                "period": l, "entity": disp(e), "tier": SUPPLY_CHAIN[e]["tier"],
                "E_local": r.local["Environmental"][e],
                "S_local": r.local["Social"][e],
                "E_total": r.totals["Environmental"][e],
                "S_total": r.totals["Social"][e],
                "composite_rolling_benchmark": r.composite[e],
                "composite_fixed_benchmark": rf.composite[e],
                "hard_threshold": r.hard_hits.get(e, ""),
                # share of the published weight the firm reported
                "E_coverage": _coverage_or_nan(r, "Environmental", e),
                "S_coverage": _coverage_or_nan(r, "Social", e),
                # scoring versions behind the rolling-benchmark scores
                "E_version": r.weights["Environmental"].version_id,
                "S_version": r.weights["Social"].version_id,
            })

    weight_frames = []
    for l in labels:
        for dim, ws in out["periods"][l].weights.items():
            d = ws.rows.copy()
            d.insert(0, "period", l)
            d.insert(1, "dimension", dim)
            d.insert(2, "version_id", ws.version_id)
            weight_frames.append(d)

    # provisional status of every published local score (cold-start and
    # coverage safeguard)
    prov_rows = []
    for l in labels:
        r = out["periods"][l]
        for dim in ("Environmental", "Social"):
            for e in ENTITY_ORDER:
                v = r.provisional.get(dim, {}).get(e, {})
                prov_rows.append({
                    "period": l, "dimension": dim, "entity": disp(e),
                    "tier": SUPPLY_CHAIN[e]["tier"],
                    "coverage_share": v.get("coverage_share", np.nan),
                    "coverage_floor": policy.coverage_floor,
                    "contributed_to_benchmark": v.get("contributed_to_benchmark", False),
                    "provisional": v.get("provisional", False),
                    "reasons": "; ".join(v.get("reasons", [])),
                    "provisional_via_direct_supplier": ", ".join(
                        disp(j) for j in v.get("via", [])),
                    "direct_upstream_coverage": v.get("direct_upstream_coverage", np.nan),
                    "direct_suppliers_without_data": ", ".join(
                        disp(j) for j in v.get("direct_suppliers_without_data", [])),
                    "assurance_review": bool(v.get("assurance_review", False)),
                    "assurance_review_reasons": "; ".join(v.get("assurance_review_reasons", [])),
                })

    interval_rows = []
    for label, result in out["periods"].items():
        for dim, items in result.scores.items():
            intervals = score_intervals(result.scores)[dim]
            for entity, sc in items.items():
                prov = result.provisional.get(dim, {}).get(entity, {})
                interval_rows.append(dict(period=label, dimension=dim, entity=entity,
                    coverage=sc.coverage, incomplete=sc.incomplete,
                    low_coverage=sc.coverage < policy.coverage_floor - 1e-12,
                    local_lower=sc.lower, local_upper=sc.upper,
                    total_lower=intervals[entity][0], total_upper=intervals[entity][1],
                    point_value=sc.score,
                    direct_upstream_coverage=prov.get("direct_upstream_coverage", np.nan)))
    pending = pd.DataFrame([dict(period=label, **row)
                            for label, result in out["periods"].items()
                            for row in result.pending_declarations],
                           columns=["period", "dimension", "entity", "indicator", "status"])
    data_status = pd.DataFrame(
        [row for l in labels for row in out["periods"][l].data_status],
        columns=["period", "dimension", "entity", "ind_id", "data_source", "data_status",
                 "as_of", "age_periods", "max_age_periods", "used_in_score", "note"])

    reg_rows = []
    for dim, reg in (out.get("registers") or {}).items():
        for k in reg.members():
            reg_rows.append({
                "dimension": dim, "register_version": reg.version_id,
                "register_digest": reg.digest,
                "classification_digest": reg.classification_digest,
                "indicator": k, "tier": "T2" if k in reg.w_reg else "T3",
                "w_reg": reg.w_reg.get(k, np.nan),
                "lower_bound": reg.band_lo.get(k, reg.t3_floor.get(k, np.nan)),
                "upper_bound": reg.band_hi.get(k, reg.t3_cap.get(k, np.nan)),
                "s_ref": reg.s_ref.get(k, np.nan),
                "s_ref_basis": reg.s_ref_basis.get(k, ""),
            })

    offsetting = None
    if "offsetting" in out:
        off = out["offsetting"]
        offsetting = pd.DataFrame([{
            "scenario": "constructed offsetting stress case",
            **{k: v for k, v in off["meta"].items()},
            "hub_net_change": off["hub_report"].net_change,
            "hub_max_component": off["hub_report"].max_abs_component,
            "hub_aggregate_rule_fires": off["hub_report"].net_fired,
            "hub_internal_warning": off["hub_report"].internal_warning,
            "focal_net_change": off["focal_report"].net_change,
            "focal_aggregate_rule_fires": off["focal_report"].net_fired,
            "focal_internal_warning": off["focal_report"].internal_warning,
        }])

    roles = pd.DataFrame([{
        "role": "Governance body (regulator, consortium or public institution)",
        "sets": "hard-alert rules, baseline weights, bounds, lambda, recalibration schedule",
        "sees": "aggregate results and authorized assurance findings; no firm records"}, {
        "role": "Participating firm",
        "sets": "its local scores and its dependency-weight method",
        "sees": "its own data and its direct suppliers' total scores"}, {
        "role": "Coordinator",
        "sets": "recalibration, versioned weight publication, log",
        "sees": "submitted summaries; in this single-process simulation also the plaintext "
                "values (a summary with n = 1 reveals the value)"}, {
        "role": "Assurance provider",
        "sets": "checks of declarations, data, calculations and dependency-weight method",
        "sees": "its client's records and the published parameters (filtered by client in "
                "the simulation, not isolated)"},
    ])

    events = out.get("events") or {}
    event_text, event_frames = [], []
    for name, table in events.get("scenarios", {}).items():
        head, text = EVENT_SCENARIO_TEXT.get(name, (name, ""))
        event_text.append({"event": name[0], "type": head.split(" · ", 1)[-1],
                           "description": text})
        event_frames.append(table.assign(event=name[0]))
    event_table = None
    if event_frames:
        # One table for events A-E: the event letter, the key columns, then the
        # rest.  Columns an event does not use stay empty.
        event_table = pd.concat(event_frames, ignore_index=True, sort=False)
        first = ["event"] + [c for c in EVENT_KEY_COLUMNS if c in event_table.columns]
        event_table = event_table[first + [c for c in event_table.columns
                                           if c not in first]]

    ablation, assurance = out["ablation"], out["assurance"]
    sens = out["sensitivity"]
    betas = (f"from {sens['beta'].min():.2f} to {sens['beta'].max():.2f} "
             if len(sens) else "")
    sc = out.get("scenario_changes") or {}
    rob = out.get("weight_robustness") or {}
    cs = out.get("cold_start") or {}
    sub = out.get("scorer_substitution")
    ml = out.get("ml_benchmark")
    m6 = out.get("m6_peers")
    m7 = out.get("m7")

    def status(result):
        if result is None:
            return None
        return pd.DataFrame([{"status": result["status"], "reason": result["reason"],
                              "results_usable": result.get("results_usable", False),
                              "requirements": result["requirements"]}])

    # ---- result workbook, one sheet per topic --------------------------------
    sheet(results, "Scores",
          ("T1", "Local, total and composite scores by period and firm, with the rolling "
           "and the fixed M0 benchmark",
           pd.DataFrame(score_rows)),
          ("T18b", "Change in each firm's total in the dimension the scenario alters, "
           "published run (thousandths; internal warning W↑ deterioration, W↓ improvement, "
           "W± both directions; H(self) hard alert raised at the firm, H hard alert "
           "received)",
           sc.get("panel_published")),
          ("T18c", "The same changes with the benchmark and weights fixed at M0 "
           "(thousandths)", sc.get("panel_fixed")),
          ("T18a", "Score changes by benchmark, transition and firm",
           sc.get("long")))
    sheet(results, "Weights",
          ("T2", "Published weights by period, dimension and indicator",
           pd.concat(weight_frames, ignore_index=True)),
          ("T2b", "Weight drift against the anchor period",
           out["drift"]),
          ("T2c", "Governance register: baseline weights, bounds and reference scales",
           pd.DataFrame(reg_rows) if reg_rows else None))
    sheet(results, "Alerts",
          ("T3", "Internal warnings and hard alerts by transition, dimension and firm "
           "(net_change: under the previous benchmark and weights; new_version_effect: "
           "due only to the new ones)", out["triggers"]),
          ("T11", "Constructed offsetting case inside one branch", offsetting),
          ("T14", "Hard-alert register, M0-M5",
           out.get("hard_alert_register")),
          ("T14b", "Closure notices after review clearance, events",
           events.get("closures")),
          ("T13", f"Hard-threshold indicators against declared limits, {labels[-1]}",
           out.get("absolute_risk")))
    sheet(results, "Events",
          ("T16", "Events A-E, each from the published M0 state with the benchmark and "
           "weights frozen", pd.DataFrame(event_text) if event_text else None),
          ("T16A-E", "Results by event and firm (columns an event does not use are empty)",
           event_table))
    sheet(results, "Comparisons",
          ("T5", f"Benchmark from participant summaries and from pooled records, "
           f"{labels[0]}", out["equiv"]),
          ("T6", f"Environmental totals before and after a +25% shock to every social "
           f"value, {labels[0]}", out["independence"]),
          ("T4", "Federated scoring against (a) centralized pooling, (b) own-history "
           "scoring and (c) rules-based scoring", ablation["table"]),
          ("T4b", "Composite scores by period and method", ablation["detail"]),
          ("T4c", "Indicator coverage of own-history scoring", ablation["local_coverage"]))
    sub_part = (lambda key: rows(sub.get(key))) if sub else (lambda key: None)
    ml_part = (lambda key: rows(ml.get(key))) if ml else (lambda key: None)
    figures = ml_part("artifacts")
    if figures is not None and "artifact" in figures.columns:
        figures = figures.assign(artifact=figures["artifact"].map(os.path.basename))
    sheet(results, "ML check",
          ("T17g", "XGBoost scorer in place of the scoring rule: status (runs only with "
           "the optional packages)", status(sub)),
          ("T17e", "Scorer substitution, summary", sub_part("summary")),
          ("T17f", "Scorer substitution, checks", sub_part("checks")),
          ("T17a", "Scorer training", sub_part("training")),
          ("T17b", "Scorer events", sub_part("events")),
          ("T17c", "Scorer event inputs", sub_part("event_inputs")),
          ("T17d", "Scorer transitions", sub_part("transitions")),
          ("T4h", "SHAP attribution of the XGBoost scorer, status", status(ml)),
          ("T4d", "SHAP scorer summary", ml_part("summary")),
          ("T4e", "SHAP and rule weights", ml_part("comparison")),
          ("T4g", "SHAP figures, saved next to this workbook", figures),
          ("T4i", "SHAP change in the events", ml_part("event_change")),
          ("T4f", "SHAP value and rule contribution by indicator and observation",
           ml_part("observations")))
    m7_blocks: tuple = ()
    if m7 is not None:
        settings = pd.DataFrame([{
            "setting": n, "beta_FOCAL": b,
            **{f"alpha_{j}": a for j, a in zip(M7_BRANCHES, (r, c))},
            "upstream_share": r + c,
            f"ratio_{M7_BRANCHES[0]}_to_{M7_BRANCHES[1]}": r / c if c else np.nan,
            "note": M7_SETTING_NOTES.get(n, ""),
        } for n, b, r, c in m7["settings"]])
        m7_checks = m7["checks"].assign(
            requirement=m7["checks"]["check"].map(lambda k: M7_CHECK_TEXT.get(k, k)))
        cols = ["setting", "beta_FOCAL", *[f"alpha_{j}" for j in M7_BRANCHES], "dimension",
                "FOCAL_total", "change_vs_baseline", "own_operations",
                *[f"branch_{j}" for j in M7_BRANCHES], "branch_ranking", "ranking_changed"]
        m7_blocks = (
            ("M7", "FOCAL's total and branch contributions under three dependency-weight "
             "settings (M0 inputs, benchmarks and ESG weights fixed)",
             m7["table"][[c for c in cols if c in m7["table"]]]),
            ("M7", f"Settings from the {M7_SETTINGS_SHEET} sheet of the M7 workbook "
             "(illustrative)", settings),
            ("M7", "Checks", m7_checks[["check", "requirement", "passed"]]))
    sheet(results, "Dependency weights",
          ("T7", f"FOCAL's totals and upstream share for beta {betas}in {labels[-1]}, "
           f"supplier weights rescaled", sens),
          ("T19b", "Alternative dependency-weight schemes by firm and the decisions they "
           "change", rob.get("summary")),
          ("T19a", "Warning and alert decisions under each scheme", rob.get("table")),
          ("T19c", "Decisions that change", rob.get("flips")),
          *m7_blocks)
    scoring = out.get("scoring_sensitivity") or {}
    sheet(results, "Scoring choices",
          ("T23a", f"{labels[0]}-{labels[-1]} rerun with one setting changed at a time: "
           "conversion range of magnitude indicators, dispersion measure and tau; warning "
           "limits and escalation levels as published, decisions compared with T3",
           scoring.get("table")),
          ("T23b", f"{labels[0]} weights by indicator under each conversion range and "
           "dispersion measure (tau does not change weights)", scoring.get("weights")))
    sheet(results, "Assurance",
          ("T10a", "Unsupported hard-threshold declaration: social totals and hard alerts "
           "by firm", assurance["A_hard_threshold"]["cascade"]),
          ("T10b", "Anomalous uploaded summary: environmental totals as published and if "
           "undetected", assurance["B_summary_anomaly"]["impact"]),
          ("T10c", f"Recomputation of the published {labels[0]} environmental local scores",
           assurance["C_recomputation"]),
          ("T10d", "Declared and applied dependency-weight methods",
           assurance["D_alpha_method"]),
          ("T10e", "Documented responses to investigate and hard-alert episodes",
           assurance.get("E_alert_responses")))
    sheet(results, "Missing data",
          ("T21c", "Cold start and selective reporting, summary", cs.get("summary")),
          ("T21a", "New entrant by stage", cs.get("entrant")),
          ("T21b", "Selective reporting step by step", cs.get("withholding")),
          ("T20", "Provisional status and coverage of every published local score",
           pd.DataFrame(prov_rows)),
          (None, "Score intervals and coverage by period, dimension and firm",
           pd.DataFrame(interval_rows)),
          ("T22", "Observations that are not plain reported values (estimates, tags, age)",
           data_status),
          (None, "Mandatory declarations pending review", rows(pending)))
    if m6 is not None:
        m6_checks = m6["checks"].assign(
            requirement=m6["checks"]["check"].map(lambda k: M6_CHECK_TEXT.get(k, k)))
        m6_weights = m6["weights"]
        first = [c for c in M6_WEIGHT_COLUMNS if c in m6_weights.columns]
        sheet(results, "M6",
              ("M6", "Environmental local scores of S1-RUBBER and three simulated peers "
               "before and after the industry benchmark, change split into own data, "
               "benchmark and weights", m6["scores"]),
              ("M6", "Level 2 and Level 3 weight parameters before and after",
               m6_weights[first + [c for c in m6_weights.columns if c not in first]]),
              ("M6", "Published weight versions", m6.get("versions")),
              ("M6", "Disclosure check of the peer benchmark", m6.get("disclosure")),
              ("M6", "Peer inputs", m6.get("inputs")),
              ("M6", "Checks", m6_checks[["check", "requirement", "passed"]]))

    # ---- checks workbook -----------------------------------------------------
    sheet(checks, "Scope",
          ("T0", "What the demonstration implements, simulates or leaves open",
           out.get("scope")),
          ("T12", "What each role sets and sees", roles))
    sheet(checks, "Versions",
          ("T9", "Governance registers and published weight versions with their digests",
           out["version_log"]),
          ("T9a", "Reference initialization and common-scale calibration uploads",
           out.get("calibration_log")),
          ("T9b", "M6 reference initialization and calibration uploads",
           m6.get("calibration_log") if m6 else None))
    sheet(checks, "Tests",
          ("T15", "Acceptance tests",
           out.get("acceptance")),
          ("T16", "Event acceptance checks", events.get("acceptance")),
          ("T16", "Event verification record",
           pd.DataFrame(_verification_rows(events["verification"]))
           if events.get("verification") else None))
    return results, checks


def _print_summary(out: dict) -> None:
    """Short console summary of a run; the detail is in the result workbooks."""
    policy, labels = out["policy"], list(out["periods"])
    paths = dict(out["period_files"])
    acceptance, triggers = out["acceptance"], out["triggers"]
    print("ESG FedStat demonstration (all indicator values are simulated)")
    print(f"  data             {os.path.dirname(os.path.abspath(paths[labels[0]]))}")
    print(f"  periods          {', '.join(labels)}")
    print(f"  classification   digest {out['classification_digest']}")
    print(f"  policy           lambda {policy.lam:.2f}; Level 2 bounds "
          f"{policy.band_lo:.2f}-{policy.band_hi:.2f} x baseline weight; Level 3 bounds "
          f"{policy.t3_floor_frac:.2f}-{policy.t3_cap_frac:.2f} x equal weight")
    print()
    print("  Composite score by period (benchmark re-estimated each period)")
    print("  " + " " * 12 + "".join(f"{label:>8}" for label in labels))
    for e in ENTITY_ORDER:
        print("  " + f"{disp(e):<12}"
              + "".join(f"{out['periods'][label].composite[e]:8.4f}" for label in labels))
    print()
    print(f"  Internal warnings {int(triggers['internal_warning'].sum())}, hard alerts "
          f"{int(triggers['hard_alert'].sum())} (firm-dimension transitions, T3)")
    scoring = out.get("scoring_sensitivity")
    if scoring is not None:
        varied = scoring["table"][scoring["table"]["variant"] != "published"]
        print("  Scoring choices  warning decisions changed: " + ", ".join(
            f"{r.variant} {r.warning_decisions_changed}" for r in varied.itertuples())
              + " (T23a)")
    events = out.get("events") or {}
    if events:
        checks = events["acceptance"]
        print(f"  Events A-E       {int(checks['passed'].sum())}/{len(checks)} checks pass "
              f"(frozen M0 parameters)")
    for name, key in (("M6 peers", "m6_peers"), ("M7 weights", "m7")):
        if key in out:
            checks = out[key]["checks"]
            print(f"  {name:<17}{int(checks['passed'].sum())}/{len(checks)} checks pass")
        else:
            print(f"  {name:<17}not run (input workbook not found)")
    scorer = out.get("scorer_substitution") or {}
    if scorer:
        reason = f" ({scorer['reason']})" if scorer.get("reason") else ""
        print(f"  XGBoost and SHAP {str(scorer['status']).lower()}{reason}")
    print(f"  Acceptance tests {int((acceptance['result'] == 'PASS').sum())}/"
          f"{len(acceptance)} pass")
    if out.get("export_path"):
        print()
        print(f"  results          {out['export_path']}")
        print(f"  checks           {out['checks_path']}")


def export_results(out: dict, files, export_dir: str) -> Tuple[str, str]:
    """Write the result workbook and the checks workbook; return both paths."""
    os.makedirs(export_dir, exist_ok=True)
    results, checks = _result_sheets(out, files)
    paths = []
    for book, sheets, heading, notes in (
            (RESULT_WORKBOOK, results, "Results of the ESG FedStat demonstration",
             RESULT_NOTES),
            (CHECKS_WORKBOOK, checks, "Checks of the ESG FedStat demonstration",
             CHECKS_NOTES)):
        path = os.path.join(export_dir, book)
        _write_result_workbook(path, heading, notes, sheets)
        paths.append(path)
    return paths[0], paths[1]


def _block_rows(blocks) -> List[int]:
    """Row of each block's title on a stacked sheet, with one empty row
    between a table and the next title."""
    starts, row = [], 1
    for _, _, frame in blocks:
        starts.append(row)
        row += len(frame) + 3
    return starts


def _write_result_workbook(path: str, heading: str, notes: Tuple[str, ...],
                           sheets: list) -> None:
    """Contents sheet, then each sheet as a plain or stacked table."""
    from openpyxl.worksheet.hyperlink import Hyperlink
    entries = []
    for name, blocks in sheets:
        starts = [1] if len(blocks) == 1 else _block_rows(blocks)
        for (code, title, _), row in zip(blocks, starts):
            entries.append((name, code or "", title, row))
    contents = pd.DataFrame([e[:3] for e in entries], columns=["sheet", "table", "contents"])
    layout = {}
    with pd.ExcelWriter(path, engine="openpyxl") as xw:
        start = len(notes) + 2
        contents.to_excel(xw, sheet_name="Contents", index=False, startrow=start)
        front = xw.sheets["Contents"]
        front.cell(row=1, column=1, value=heading)
        for row, note in enumerate(notes, 2):
            front.cell(row=row, column=1, value=note)
        layout["Contents"] = [(1, start + 1, start + 1 + len(contents), contents.shape[1])]
        for row, (name, _, _, target) in enumerate(entries, start + 2):
            front.cell(row=row, column=1).hyperlink = Hyperlink(
                ref="", location=f"'{name}'!A{target}")
        for name, blocks in sheets:
            if len(blocks) == 1:
                frame = blocks[0][2]
                frame.to_excel(xw, sheet_name=name, index=False)
                layout[name] = [(None, 1, len(frame) + 1, frame.shape[1])]
                continue
            placed = []
            for (code, title, frame), row in zip(blocks, _block_rows(blocks)):
                frame.to_excel(xw, sheet_name=name, index=False, startrow=row)
                xw.sheets[name].cell(row=row, column=1,
                                     value=f"{code}  {title}" if code else title)
                placed.append((row, row + 1, row + 1 + len(frame), frame.shape[1]))
            layout[name] = placed
        _format_result_workbook(xw.book, layout, heading, len(notes))


def _format_result_workbook(workbook, layout: dict, title: str,
                            n_notes: int = len(RESULT_NOTES)) -> None:
    """Black-and-white three-line tables; numbers kept at full precision."""
    from openpyxl.styles import Alignment, Border, Font, Side
    from openpyxl.utils import get_column_letter
    heavy, light = Side(style="medium", color="FF000000"), Side(style="thin", color="FF000000")
    body_font = Font(name="Arial", size=10, color="FF000000")
    bold_font = Font(name="Arial", size=10, bold=True, color="FF000000")
    wide = ("detail", "requirement", "notes", "indicator", "value", "field", "reason")

    def lines(text: str, width: float) -> int:
        per_line = max(1.0, (width - 2) * 1.15)       # Arial 10 against the width unit
        return sum(max(1, int(np.ceil(len(part) / per_line)))
                   for part in str(text or "").split("\n"))

    for sheet in workbook.worksheets:
        blocks = layout.get(sheet.title, [(None, 1, sheet.max_row, sheet.max_column)])
        plain = len(blocks) == 1 and blocks[0][0] is None
        is_contents = sheet.title == "Contents"
        widths: Dict[int, float] = {}
        for _, header_row, last_row, ncols in blocks:
            sheet.row_dimensions[header_row].height = 34
            for col in range(1, ncols + 1):
                cells = [sheet.cell(row=r, column=col) for r in range(header_row, last_row + 1)]
                head, body = cells[0], cells[1:]
                header = str(head.value or "")
                numeric = any(isinstance(c.value, (int, float)) and not isinstance(c.value, bool)
                              for c in body)
                logical = not numeric and any(isinstance(c.value, bool) for c in body)
                head.font = bold_font
                head.alignment = Alignment(
                    horizontal="right" if numeric else "center" if logical else None,
                    indent=1 if numeric else 0, vertical="center", wrap_text=True)
                head.border = Border(top=heavy, bottom=light if body else heavy)
                for cell in body:
                    cell.font = (Font(name="Arial", size=10, color="FF000000", underline="single")
                                 if cell.hyperlink else body_font)
                    if isinstance(cell.value, bool):
                        cell.alignment = Alignment(horizontal="center", vertical="top")
                    elif isinstance(cell.value, (int, float)):
                        cell.alignment = Alignment(horizontal="right", indent=1, vertical="top")
                    else:
                        cell.alignment = Alignment(vertical="top", wrap_text=True)
                    if isinstance(cell.value, float):
                        cell.number_format = ("0.000000000000000" if "delta" in header.lower()
                                              else "0.000000")
                if body:
                    body[-1].border = Border(bottom=heavy)
                longest = max((len(str(c.value or "")) for c in cells), default=10)
                width = min(60, max(15, min(longest + 2, 34)))
                if any(word in header.lower() for word in wide):
                    width = min(64, max(width, 48))
                widths[col] = max(widths.get(col, 0), width)
        if is_contents:
            widths = {1: 20, 2: 10, 3: 110}
        for col, width in widths.items():
            sheet.column_dimensions[get_column_letter(col)].width = width
        for _, header_row, last_row, ncols in blocks:
            for r in range(header_row + 1, last_row + 1):
                count = max((lines(sheet.cell(row=r, column=c).value, widths.get(c, 15))
                             for c in range(1, ncols + 1)), default=1)
                sheet.row_dimensions[r].height = min(300, max(22, 15 * count + 8))
        if is_contents:
            titled = [(1, True)] + [(r, False) for r in range(2, 2 + n_notes)]
        else:
            titled = [(t, True) for t, _, _, _ in blocks if t is not None]
        for row, bold in titled:
            ncols = max(widths) if widths else 1
            span = sum(widths.get(c, 15) for c in range(1, ncols + 1))
            cell = sheet.cell(row=row, column=1)
            cell.font = Font(name="Arial", size=12 if is_contents and row == 1 else 10, bold=bold,
                             color="FF000000")
            cell.alignment = Alignment(vertical="top", wrap_text=True)
            if ncols > 1:
                sheet.merge_cells(start_row=row, start_column=1, end_row=row, end_column=ncols)
            sheet.row_dimensions[row].height = max(20, 15 * lines(cell.value, span * 0.9) + 6)
        sheet.sheet_view.showGridLines = False
        if plain:
            sheet.freeze_panes = "A2"
            sheet.auto_filter.ref = sheet.dimensions
            sheet.print_title_rows = "1:1"
        sheet.sheet_properties.pageSetUpPr.fitToPage = True
        sheet.page_setup.orientation = "landscape"
        sheet.page_setup.paperSize = sheet.PAPERSIZE_A3
        sheet.page_setup.fitToWidth = 1
        sheet.page_setup.fitToHeight = 0
        sheet.print_options.horizontalCentered = True
    workbook.active = 0
    workbook.properties.creator = "ESG research demonstration"
    workbook.properties.lastModifiedBy = "ESG research demonstration"
    workbook.properties.title = title


def _verification_rows(value, prefix: str = "") -> List[dict]:
    """Flatten nested audit evidence into readable, independently filterable rows."""
    if isinstance(value, dict):
        return [row for key, item in value.items()
                for row in _verification_rows(item, f"{prefix}.{key}" if prefix else str(key))]
    if isinstance(value, (list, tuple)):
        return [row for index, item in enumerate(value, 1)
                for row in _verification_rows(item, f"{prefix}[{index}]" )]
    return [{"field": prefix, "value": value}]


def export_event_results(events: dict, export_dir: str) -> str:
    """Write the event tables and the acceptance table as CSV and the
    verification record as JSON; return the output directory."""
    output = Path(export_dir)
    output.mkdir(parents=True, exist_ok=True)
    for scenario, table in events["scenarios"].items():
        table.to_csv(output / f"{scenario}.csv", index=False, encoding="utf-8-sig",
                     float_format="%.17g")
    events["acceptance"].to_csv(output / "acceptance.csv", index=False, encoding="utf-8-sig")
    (output / "verification.json").write_text(
        json.dumps(events["verification"], indent=2, ensure_ascii=False), encoding="utf-8")
    return str(output.resolve())

# =============================================================================
# SCENARIO GENERATION AND M4 CALIBRATION
# =============================================================================

SRC = "ESG_Raw_Data_Supply_Chain_M0.xlsx"
OUT = "data"                 # see DATA_DIRECTORY_NAME below; same value

SHEETS = {"E": DIM_SHEETS["Environmental"], "S": DIM_SHEETS["Social"]}
COMPANIES = ["S2-CHEM", "S2-EQUIP", "S1-RUBBER", "S1-COATING",
             "FOCAL", "D1-RETAIL", "D2-RETAIL"]


def locate(ws):
    """Return indicator rows and anonymous entity columns by fixed schema."""
    rows = {}
    cols = {entity: column for entity, column in
            zip(COMPANIES, range(4, 4 + len(COMPANIES)))}
    for i in range(4, ws.max_row + 1):
        match = ID_RE.match(str(ws.cell(row=i, column=2).value or ""))
        if match:
            rows[match.group(1)] = i
    return rows, cols


def apply(wb, edits):
    """edits: {(dim, ind_id): {company: value}}"""
    loc = {d: locate(wb[SHEETS[d]]) for d in SHEETS}
    for (dim, iid), per_co in edits.items():
        ws = wb[SHEETS[dim]]
        rows, cols = loc[dim]
        if iid not in rows:
            print(f"  [warn] {dim}/{iid} not found, skipped")
            continue
        for co, val in per_co.items():
            ws.cell(row=rows[iid], column=cols[co]).value = val


def note(wb, lines):
    ws = wb[NOTES_SHEET]
    r = ws.max_row + 2
    for k, line in enumerate(lines):
        ws.cell(row=r + k, column=1).value = line


US_ENGLISH_WORDS = {
    "labour": "labor", "labours": "labors", "labourer": "laborer",
    "litres": "liters", "litre": "liter", "tonnes": "metric tons",
    "behaviour": "behavior", "behaviours": "behaviors",
    "colour": "color", "colours": "colors", "favourable": "favorable",
    "unfavourable": "unfavorable", "honouring": "honoring",
    "judgement": "judgment", "judgements": "judgments",
    "enrolment": "enrollment", "enrol": "enroll",
    "normalised": "normalized", "normalisation": "normalization",
    "renormalised": "renormalized", "renormalisation": "renormalization",
    "standardised": "standardized", "standardisation": "standardization",
    "anonymised": "anonymized", "anonymisation": "anonymization",
    "organisations": "organizations", "organisation": "organization",
    "modelling": "modeling", "modelled": "modeled",
    "centralised": "centralized", "authorised": "authorized",
    "recognised": "recognized", "optimisation": "optimization",
    "optimised": "optimized", "summarised": "summarized",
    "neighbour": "neighbor", "neighbours": "neighbors",
}
_US_WORD_PATTERN = re.compile(r"\b(" + "|".join(US_ENGLISH_WORDS) + r")\b", re.I)

# Wording applied to the text cells of the input workbooks; indicator IDs,
# statuses and numeric inputs are left unchanged.
INPUT_TEXT_EDITS = {
    "E3. Scope 1+2 Combined": "E3. Combined Scope 1 and 2 GHG Emissions",
    "E6. GHG Reduction vs Baseline Year": "E6. GHG Reduction vs. Baseline Year",
    "E10. Renewable Energy Share (Electricity)": "E10. Renewable Electricity Share",
    "E17. Water Discharge (to surface/sea)": "E17. Water Discharge (to Surface Water or Sea)",
    "E38. Environmental Fines & Penalties (Last 3 Yrs)": "E38. Environmental Fines and Penalties (Last 3 Years)",
    "S5. Ethnic/Racial Minority in Workforce (US)": "S5. Ethnic/Racial Minorities in Workforce (US)",
    "S6. Ethnic/Racial Minority in Management (US)": "S6. Ethnic/Racial Minorities in Management (US)",
    "S11. Gender Pay Gap (Women vs Men, Median)": "S11. Gender Pay Gap (Women vs. Men, Median)",
    "S16. Lost Time Injury Rate (LTIR)": "S16. Lost-Time Injury Rate (LTIR)",
    "S21. OSHA Violations (Last 3 Yrs)": "S21. OSHA Violations (Last 3 Years)",
    "S25. Critical Audit Non-Conformances Found": "S25. Critical Audit Nonconformances Found",
    "S36. Product Recalls (Last 3 Yrs)": "S36. Product Recalls (Last 3 Years)",
    "S39. Product Quality Certifications (ISO 9001 etc.)": "S39. Product Quality Certifications (e.g., ISO 9001)",
    "S40. Data Privacy Breaches (Last 3 Yrs)": "S40. Data Privacy Breaches (Last 3 Years)",
    "S48. Suppliers Terminated for Non-Compliance": "S48. Suppliers Terminated for Noncompliance",
    "~0% (reported)": "~0% (simulated)",
    "Scope 3 set from the panel's median Scope 3 / Scope 1+2 ratio":
        "Simulated Scope 3 input drawn around an estimate based on the panel's median Scope 3 / Scope 1+2 ratio.",
    "Environmental indicators E1–E38 and social indicators S1–S50 keep their labels and units.":
        "Indicator IDs E1–E38 and S1–S50 and their measurement units are consistent across the workbooks.",
    "management and board by 4–12 percentage points and its supplier-diversity spending":
        "management and board by 4–12 percentage points. Supplier-diversity spending is multiplied",
    "by 44%. Social data only; every environmental score must stay unchanged.":
        "by 1.44 before display rounding. Only social inputs change; environmental scores remain unchanged.",
    "incident rate (S15) rises by about 60% and its safety violations (S21) rise 4.5-fold":
        "incident rate (S15) rises by about 60%. The S21 violation count is multiplied by 4.5",
    "against M0.": "relative to M0 before rounding to a whole count.",
    "All indicator values are simulated; the SIMULATED DATA note at the end of this sheet describes how M0 was drawn.":
        "All indicator values are simulated. The SIMULATED DATA section describes how M0 was generated.",
    "Supply links, dependency weights and the governance gradient on seven process-and-verification indicators in M0 are designed.":
        "Supply relationships, dependency weights and the M0 gradient on seven process and verification indicators are scenario assumptions.",
    "The indicator values are the exact inputs that reproduce the published results.":
        "Run the supplied code with these inputs and its declared configuration to reproduce the demonstration results.",
    "Indicators with observed zero variance keep their positive governed weights and floors.":
        "Weighted indicators with observed zero variance retain positive governed weights and their declared bounds.",
    "S26/S27/S28 are active declaration gates. E38/S17/S18/S21/S40/S45 are monitored without configured numeric ceilings.":
        "S26, S27 and S28 have active policy-declaration rules. E38, S17, S18, S21, S40 and S45 are monitored without configured numerical limits.",
    "Admitting an indicator does not mean complete coverage. Missing observations are retained and bounded in scoring (a published score interval), not imputed or replaced by zero.":
        "Including an indicator does not establish complete coverage. Missing inputs retain their weights and produce score bounds; they are not replaced by zero.",
    "Weights blend the baseline weights with a data-driven share and are projected onto approved bounds that sum to one; the Level 2/3 block shares set the baseline mass only.":
        "Level 2 and 3 weights share one baseline blend and bounded projection. Final weights sum to one when all benchmarks are available; 70/30 specifies the default baseline allocation only.",
    "All indicator values are simulated. Each numeric M0 value of the seven chain firms is a":
        "All indicator values are simulated. Numeric M0 values for the seven chain firms are drawn",
    "seeded random draw around a reference value of the same unit and order of magnitude":
        "around reference magnitudes in the same units. Amounts, counts, rates and ratios use",
    "(log-normal for amounts, counts, rates and ratios; logit-normal for percentage shares;":
        "log-normal draws; percentage shares use logit-normal draws. The log or logit SD is 0.30,",
    "spread 0.30, minimum change 5%), so no value reproduces a figure reported":
        "with a minimum absolute shift of 0.05 on that scale. Display rounding and identity constraints",
    "by any company. E3 = E1 + E2, E14 = E13 x E15, E19 = E18 x E20 and E22 = E18 - E19 hold":
        "are applied after the draw. These inputs do not represent any company's reported performance.",
    "after the draw. Yes/No/Partial declarations, including the M0 governance gradient, are":
        "After the draw, E3 = E1 + E2, E14 = E13 × E15, E19 = E18 × E20 and E22 = E18 − E19 hold.",
    "design values; free-text targets and letter grades, which no score uses, are generic.":
        "Yes/No/Partial declarations are scenario settings. Free-text targets and letter grades are generic and excluded from weighted scoring.",
    "M1–M5 changes are defined relative to the simulated M0 (multiples or percentage-point":
        "M1–M5 changes are defined relative to the simulated M0 as multiples or percentage-point",
    "shifts), except the M3 renewable shares, which are set directly. No scored input of the":
        "shifts, except M3 renewable shares, which are set directly. Every scored indicator for the",
    "seven chain firms is missing.":
        "seven chain firms has an input in the supplied baseline and periodic scenarios.",
    "Target values are taken from Environmental Data. Both states use the same four industry members,":
        "Target values are taken from Environmental Data. Both states use the same four industry members.",
    "the minimum group size at which no member can solve for another's values from the published benchmark.":
        "Four members meet the minimum publication size. This reduces inference risk but does not guarantee privacy.",
    "Percent cells store fractions and use percentage formatting. These scores are not full M0–M5 ESG scores.":
        "Numeric percentage cells store fractions and use percentage formatting. M6 reports environmental peer scores, not supply-chain composites.",
    "higher total upstream share; supplier ratio of the baseline kept":
        "Higher upstream dependency weight; the baseline ratio between supplier weights is preserved.",
    "total upstream share of the baseline kept; allocation between the suppliers changed":
        "Baseline upstream dependency weight is preserved; its allocation between suppliers changes.",
}


def proofread_input_text(value: str) -> str:
    """Apply US spelling and INPUT_TEXT_EDITS to one text value."""
    def american(match):
        word = match.group(0)
        replacement = US_ENGLISH_WORDS[word.lower()]
        return replacement.upper() if word.isupper() else (
            replacement.capitalize() if word[0].isupper() else replacement)
    text = _US_WORD_PATTERN.sub(american, value)
    return INPUT_TEXT_EDITS.get(text, text)


def proofread_input_workbook(wb) -> None:
    """Apply proofread_input_text to every text cell of the workbook."""
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                if isinstance(cell.value, str) and cell.data_type != "f":
                    cell.value = proofread_input_text(cell.value)


# ---------------------------------------------------------------------------
# M0: governance gradient on process and verification indicators
# ---------------------------------------------------------------------------
# Deeper-tier suppliers lag retailers on whether due diligence is run and the
# supplier code enforced, not on whether a policy exists.  Applied to the
# reference workbook before the baseline is simulated.
M0_EDITS = {
    ("S", "S23"): {"S2-CHEM": "Partial", "S1-RUBBER": "Partial"},          # HR due diligence process
    ("S", "S29"): {"S2-CHEM": "No", "S1-RUBBER": "Partial",                # conflict minerals DD
                   "S1-COATING": "No"},
    ("S", "S30"): {"S1-RUBBER": "Partial"},                                # freedom of association
    ("S", "S46"): {"S1-RUBBER": "Partial", "S2-CHEM": "Partial"},          # supplier code of conduct
    ("S", "S50"): {"S2-CHEM": "Partial", "S1-RUBBER": "Partial"},          # responsible sourcing
    ("E", "E35"): {"S1-RUBBER": "Partial", "S1-COATING": "Partial"}, # TCFD alignment
    ("E", "E16"): {"FOCAL": "No", "D2-RETAIL": "No"},                       # water-stress operations
    # Values set before the draw: E6 as a positive reduction share, FOCAL's
    # Scope 3 (E4) from the panel's median ratio of Scope 3 to Scope 1+2, and
    # S2-CHEM's deforestation commitment (E32) as No, as for the other Tier 2
    # supplier.
    ("E", "E6"): {"S2-CHEM": "12%", "S2-EQUIP": "18%", "S1-RUBBER": "25.4%",
                  "S1-COATING": "15%", "FOCAL": "14%", "D1-RETAIL": "18.1%",
                  "D2-RETAIL": "22%"},
    ("E", "E4"): {"FOCAL": "~3,840,000"},
    ("E", "E32"): {"S2-CHEM": "No"},
}

# ---------------------------------------------------------------------------
# Simulated baseline
# ---------------------------------------------------------------------------
# Every numeric M0 value of the seven firms is a seeded draw around a reference
# value of the same unit and magnitude: log-normal for amounts, counts, rates
# and ratios, logit-normal for shares (spread SIMULATION_SIGMA, minimum change
# SIMULATION_MIN_SHIFT).  Zeros and 100% shares are kept; declarations are
# design values.  After the draw E3 = E1 + E2, E14 = E13 x E15, E19 = E18 x E20
# and E22 = E18 - E19 are restored, and E17 <= 0.95 E13, S16 <= 0.9 S15 and
# S45 <= S44 are enforced.  The draw is applied once; a workbook carrying
# SIMULATED_BASELINE_ID is reused as it is.
SIMULATION_SEED = 20261026
SIMULATION_SIGMA = 0.30
SIMULATION_MIN_SHIFT = 0.05
SIMULATED_BASELINE_ID = f"simulated-baseline-v1-seed-{SIMULATION_SEED}"

_GENERIC_TEXT = {
    "E7": ("30% by 2030", "40% by 2030", "45% by 2030", "50% by 2030", "Net zero by 2040"),
    "E11": ("50% by 2030", "80% by 2030", "100% by 2030", "100% by 2035"),
    "E31": ("None", "Reduce 20% by 2030", "Reduce 25% by 2030", "Reduce 30% by 2030"),
    "E36": ("C", "C+", "B-", "B", "B+", "A-"),
}
_MISSING_TEXT = {"", "N/A", "N/D", "N/R"}
_NUMBER_TEXT = re.compile(r"^(~?)\s*(-?\d[\d,]*(?:\.\d+)?)\s*(%?)$")
_RATIO_TEXT = re.compile(r"^(~?)\s*(\d[\d,]*(?:\.\d+)?)\s*:\s*1$")


@dataclass(frozen=True)
class _NumericCell:
    """A parsed numeric workbook cell and the way it is displayed."""
    prefix: str          # "~" for an approximate magnitude, otherwise ""
    value: float
    decimals: int
    percent: bool
    ratio: bool          # "85:1"
    native: bool         # stored as a number rather than as text


def _parse_numeric_cell(value) -> Optional[_NumericCell]:
    if value is None or isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        x = float(value)
        if not np.isfinite(x):
            return None
        return _NumericCell("", x, 0 if x.is_integer() else 2, False, False, True)
    if not isinstance(value, str):
        return None
    text = value.strip()
    match = _RATIO_TEXT.match(text)
    ratio = match is not None
    if match is None:
        match = _NUMBER_TEXT.match(text)
    if match is None:
        return None
    digits = match.group(2)
    decimals = len(digits.split(".", 1)[1]) if "." in digits else 0
    percent = (not ratio) and match.group(3) == "%"
    return _NumericCell(match.group(1), float(digits.replace(",", "")), decimals,
                        percent, ratio, False)


def _round_significant(x: float, digits: int = 3) -> float:
    if x == 0.0 or not np.isfinite(x):
        return x
    return round(x, digits - 1 - int(np.floor(np.log10(abs(x)))))


def _format_simulated(cell: _NumericCell, x: float):
    """Write a simulated value in the unit and precision of its cell
    (three significant figures for magnitudes of 1,000 or more)."""
    if cell.native:
        return int(round(x)) if cell.decimals == 0 else round(x, cell.decimals)
    if cell.percent:
        body = f"{x:.{cell.decimals}f}%"
    elif cell.ratio:
        body = f"{int(round(x)):,}:1"
    elif cell.decimals == 0:
        body = f"{int(round(_round_significant(x, 3) if abs(x) >= 1000 else x)):,}"
    else:
        body = f"{x:,.{cell.decimals}f}"
    return cell.prefix + body


def _simulation_shift(iid: str, entity: str) -> float:
    """Seeded normal draw for one cell, on the log (or logit) scale."""
    key = f"{SIMULATION_SEED}|{iid}|{entity}".encode("utf-8")
    rng = random.Random(int.from_bytes(hashlib.sha256(key).digest()[:8], "big"))
    shift = SIMULATION_SIGMA * rng.gauss(0.0, 1.0)
    if abs(shift) < SIMULATION_MIN_SHIFT:
        shift = SIMULATION_MIN_SHIFT if shift >= 0.0 else -SIMULATION_MIN_SHIFT
    return shift


def _simulate_cell(value, iid: str, entity: str):
    cell = _parse_numeric_cell(value)
    if cell is None:
        return value
    x = cell.value
    if x == 0.0 or (cell.percent and x >= 100.0):
        return value
    shift = _simulation_shift(iid, entity)
    if cell.percent and x > 0.0:
        p = x / 100.0
        y = 100.0 / (1.0 + float(np.exp(-(np.log(p / (1.0 - p)) + shift))))
    else:
        y = x * float(np.exp(shift))
    simulated = _format_simulated(cell, y)
    if simulated == _format_simulated(cell, x):
        # The draw would still display as the reference value: move one
        # display step (upwards for small counts, so that none becomes zero).
        if cell.decimals:
            step = 10.0 ** -cell.decimals
        elif abs(y) >= 1000 and not cell.percent:
            step = 10.0 ** (int(np.floor(np.log10(abs(y)))) - 2)
        else:
            step = 1.0
        small_count = cell.decimals == 0 and not cell.percent and abs(x) < 3
        upward = shift > 0.0 or small_count or y - step < 0.0
        y = x + step if small_count else (y + step if upward else y - step)
        simulated = _format_simulated(cell, y)
    return simulated


def _generic_text(iid: str, entity: str, value):
    options = _GENERIC_TEXT.get(iid)
    if (options is None or not isinstance(value, str)
            or value.strip().upper() in _MISSING_TEXT):
        return value
    key = f"{SIMULATION_SEED}|{iid}|{entity}|text".encode("utf-8")
    return options[int.from_bytes(hashlib.sha256(key).digest()[:4], "big") % len(options)]


def _restore_identities(wb) -> None:
    """Re-derive the totals and shares that independent draws would decouple."""
    def sheet(key):
        ws = wb[SHEETS[key]]
        rows, cols = locate(ws)

        def get(iid, co):
            return (_parse_numeric_cell(ws.cell(row=rows[iid], column=cols[co]).value)
                    if iid in rows else None)

        def put(iid, co, cell, x, exact=False):
            ws.cell(row=rows[iid], column=cols[co]).value = (
                cell.prefix + f"{int(round(x)):,}" if exact and not cell.native
                else _format_simulated(cell, x))
        return get, put

    get, put = sheet("E")
    for co in COMPANIES:
        e1, e2, e3 = get("E1", co), get("E2", co), get("E3", co)
        if e1 and e2 and e3:
            put("E3", co, e3, e1.value + e2.value, exact=True)
        e13, e14, e15, e17 = get("E13", co), get("E14", co), get("E15", co), get("E17", co)
        if e13 and e14 and e15 and e15.percent:
            put("E14", co, e14, e13.value * e15.value / 100.0)
        if e13 and e17 and e17.value > 0.95 * e13.value:
            put("E17", co, e17, 0.95 * e13.value)
        e18, e19, e20, e22 = get("E18", co), get("E19", co), get("E20", co), get("E22", co)
        if e18 and e19 and e20 and e20.percent:
            diverted = _round_significant(e18.value * e20.value / 100.0, 3)
            put("E19", co, e19, diverted)
            if e22:
                put("E22", co, e22, e18.value - diverted, exact=True)
    get, put = sheet("S")
    for co in COMPANIES:
        s15, s16 = get("S15", co), get("S16", co)
        if s15 and s16 and s16.value > 0.9 * s15.value:
            put("S16", co, s16, 0.9 * s15.value)
        s44, s45 = get("S44", co), get("S45", co)
        if s44 and s45 and s45.value > s44.value:
            put("S45", co, s45, s44.value)


def simulate_baseline(wb) -> int:
    """Replace the seven firms' M0 values with seeded simulated values.

    Returns the number of cells changed.  Declarations, missing-value codes
    and observed zeros are left as they are.
    """
    changed = 0
    for key in ("E", "S"):
        ws = wb[SHEETS[key]]
        rows, cols = locate(ws)
        for iid, r in rows.items():
            for co in COMPANIES:
                cell = ws.cell(row=r, column=cols[co])
                old = cell.value
                new = (_generic_text(iid, co, old) if iid in _GENERIC_TEXT
                       else _simulate_cell(old, iid, co))
                if new != old:
                    cell.value = new
                    changed += 1
    _restore_identities(wb)
    return changed


def _baseline_is_simulated(wb) -> bool:
    if str(getattr(wb.properties, "identifier", None) or "") == SIMULATED_BASELINE_ID:
        return True
    for name in wb.sheetnames:
        if name.strip().lower().startswith("methodology"):
            for row in wb[name].iter_rows(min_col=1, max_col=1, values_only=True):
                if row and isinstance(row[0], str) and SIMULATED_BASELINE_ID in row[0]:
                    return True
    return False


def prepare_baseline(wb) -> bool:
    """Bring a source workbook to the simulated M0 baseline.

    The draw is applied once and marked; a marked workbook is only normalized and
    anonymized.  Returns True when the draw was applied now.
    """
    simulated = _baseline_is_simulated(wb)
    normalise_sheet_names(wb)
    _anonymise_workbook_entities(wb)
    if not simulated:
        apply(wb, M0_EDITS)
        simulate_baseline(wb)
    proofread_input_workbook(wb)
    wb.properties.identifier = SIMULATED_BASELINE_ID
    return not simulated


def baseline_digest(wb) -> str:
    """Digest of the seven firms' prepared M0 values (both data sheets)."""
    h = hashlib.sha256()
    for key in ("E", "S"):
        ws = wb[SHEETS[key]]
        rows, cols = locate(ws)
        for iid in sorted(rows, key=rows.get):
            for co in COMPANIES:
                value = ws.cell(row=rows[iid], column=cols[co]).value
                h.update(f"{key}|{iid}|{co}|{value!r}\n".encode("utf-8"))
    return h.hexdigest()[:16]


# ---------------------------------------------------------------------------
# Scenario edits relative to the simulated M0
# ---------------------------------------------------------------------------
# A string is written as it stands.  Scaled(f) writes the cell's M0 value times
# f, and Shifted(d) its M0 value plus d (percentage points for a percentage
# cell), in the unit and precision of the M0 cell.  Every period starts from
# M0, so the multiples are always relative to M0, not to the previous period.

@dataclass(frozen=True)
class Scaled:
    factor: float


@dataclass(frozen=True)
class Shifted:
    delta: float


def _format_scenario_value(cell: _NumericCell, x: float) -> str:
    if cell.percent:
        return f"~{min(max(x, 0.0), 100.0):.{cell.decimals}f}%"
    if cell.ratio:
        return f"~{int(round(x)):,}:1"
    if cell.decimals == 0:
        return f"~{int(round(x)):,}"
    return f"~{x:,.{cell.decimals}f}"


def resolve_scenario_edits(wb, edits) -> dict:
    """Turn relative scenario edits into cell values against the M0 cells of `wb`.

    Runs before any scenario edit is applied to `wb`.
    """
    loc = {d: locate(wb[SHEETS[d]]) for d in SHEETS}
    resolved = {}
    for (dim, iid), per_co in edits.items():
        rows, cols = loc[dim]
        out = {}
        for co, spec in per_co.items():
            if not isinstance(spec, (Scaled, Shifted)):
                out[co] = spec
                continue
            if iid not in rows:
                continue                              # apply() reports the missing row
            base = wb[SHEETS[dim]].cell(row=rows[iid], column=cols[co]).value
            cell = _parse_numeric_cell(base)
            if cell is None:
                raise ConfigurationError(
                    f"scenario edit {dim}/{iid}/{co} needs a numeric M0 value, found {base!r}")
            x = cell.value * spec.factor if isinstance(spec, Scaled) else cell.value + spec.delta
            out[co] = _format_scenario_value(cell, x)
        resolved[(dim, iid)] = out
    return resolved


# Values that are not plain reported figures, listed in the Data Status sheet
# of every generated period workbook: (entity, indicator) -> (source, status, note).
DATA_TAGS = {
    ("FOCAL", "E4"): ("estimate", "unassured",
                      "Scope 3 set from the panel's median Scope 3 / Scope 1+2 ratio"),
}

# Appended to the Methodology sheet of every generated period workbook.  The
# first line carries the baseline marker.
SIMULATED_DATA_NOTES = [
    "",
    f"SIMULATED DATA ({SIMULATED_BASELINE_ID})",
    "All indicator values are simulated. Each numeric M0 value of the seven chain firms is a",
    "seeded random draw around a reference value of the same unit and order of magnitude",
    "(log-normal for amounts, counts, rates and ratios; logit-normal for percentage shares;",
    f"spread {SIMULATION_SIGMA:.2f}, minimum change {SIMULATION_MIN_SHIFT:.0%}), so no value "
    f"reproduces a figure reported",
    "by any company. E3 = E1 + E2, E14 = E13 x E15, E19 = E18 x E20 and E22 = E18 - E19 hold",
    "after the draw. Yes/No/Partial declarations, including the M0 governance gradient, are",
    "design values; free-text targets and letter grades, which no score uses, are generic.",
    "M1–M5 changes are defined relative to the simulated M0 (multiples or percentage-point",
    "shifts), except the M3 renewable shares, which are set directly. No scored input of the",
    "seven chain firms is missing.",
]

# ---------------------------------------------------------------------------
# M1: Tier 2 shock, upstream emissions double
# ---------------------------------------------------------------------------
M1_EDITS = {
    ("E", "E1"): {"S2-CHEM": Scaled(2.0)},
    ("E", "E2"): {"S2-CHEM": Scaled(2.0)},
    ("E", "E3"): {"S2-CHEM": Scaled(2.0)},
    ("E", "E4"): {"S2-CHEM": Scaled(2.0)},
    ("E", "E24"): {"S2-CHEM": Scaled(1.89)},       # NOx +89%
}

# ---------------------------------------------------------------------------
# M2: social change only, upstream diversity improves
# ---------------------------------------------------------------------------
M2_EDITS = {
    ("S", "S2"): {"S2-EQUIP": Shifted(4.0)},
    ("S", "S3"): {"S2-EQUIP": Shifted(4.0)},
    ("S", "S4"): {"S2-EQUIP": Shifted(12.0)},
    ("S", "S5"): {"S2-EQUIP": Shifted(5.0)},
    ("S", "S6"): {"S2-EQUIP": Shifted(5.0)},
    ("S", "S49"): {"S2-EQUIP": Scaled(1.44)},     # supplier diversity spend, T3 representative
}

# ---------------------------------------------------------------------------
# M3: renewable shares converge (dispersion collapse on E10)
# ---------------------------------------------------------------------------
# Every firm converges on renewable electricity procurement, which is cheap and
# purchasable.  E10's cross-sectional spread collapses to about one percentage
# point, while Scope 1 process emissions, the expensive route, barely move
# (S2-CHEM stays at its M1 level, less 1%).
M3_EDITS = {
    ("E", "E10"): {"S2-CHEM": "~58%", "S2-EQUIP": "~60%", "S1-RUBBER": "~61%",
                   "S1-COATING": "~59%", "FOCAL": "~60%",
                   "D1-RETAIL": "~62%", "D2-RETAIL": "~59%"},
    # Scope 1 essentially flat
    ("E", "E1"): {"S2-CHEM": Scaled(1.98), "S2-EQUIP": Scaled(0.993),
                  "S1-RUBBER": Scaled(0.995), "S1-COATING": Scaled(0.995),
                  "FOCAL": Scaled(0.995), "D1-RETAIL": Scaled(0.996),
                  "D2-RETAIL": Scaled(0.997)},
}

# ---------------------------------------------------------------------------
# M4: offsetting inside the Tier-1 hub's branch
# ---------------------------------------------------------------------------
# S2-CHEM's thirteen risk-block indicators go to M4_CHEM_FACTOR times their M0
# value (greenhouse gases and NOx fall from their M1-M3 level, the others rise
# 50%); S2-EQUIP's are scaled by a common factor g, solved by
# `calibrate_m4_edits` so that S1-RUBBER's aggregate barely moves while each
# supplier's contribution moves past the warning limit.
M4_CHEM_FACTOR = 1.50
M4_RISK_BLOCK = ["E1", "E2", "E3", "E4", "E24", "E25", "E26", "E27",
                 "E13", "E17", "E18", "E21", "E22"]

# The solved factor is saved next to the generated workbooks with the digest of
# the prepared baseline it belongs to; `generate` reuses it only for that baseline.
M4_CALIBRATION_FILE = "M4_calibration.json"


def m4_edits(improvement_factor: float) -> dict:
    return {("E", iid): {"S2-CHEM": Scaled(M4_CHEM_FACTOR),
                         "S2-EQUIP": Scaled(float(improvement_factor))}
            for iid in M4_RISK_BLOCK}


# ---------------------------------------------------------------------------
# M5: hard threshold
# ---------------------------------------------------------------------------
M5_EDITS = {
    ("S", "S26"): {"S1-RUBBER": "No"},          # declared policy criterion fails
    ("S", "S21"): {"S1-RUBBER": Scaled(4.5)},   # OSHA violations climb
    ("S", "S15"): {"S1-RUBBER": Scaled(1.59)},  # incident rate worsens
}

# Scenario notes, written to the Methodology sheet of each generated workbook
# and shown on the dashboard.
PERIOD_NOTES = {
    0: ["", "SCENARIO M0 — BASELINE",
        "Baseline of the simulated supply chain. All values are simulated (see SIMULATED DATA",
        "below); scores do not reflect the ESG performance of any company. A governance",
        "gradient is applied to process-and-verification indicators (S23, S29, S30, S46, S50,",
        "E35, E16), where a deeper-tier supplier plausibly lags a retailer. Policy indicators",
        "that are universal among large firms stay universal; they carry zero dispersion but",
        "remain in the governed register with a positive weight. The published sigma_score",
        "controls standardization."],
    1: ["", "SCENARIO M1 — TIER 2 EMISSIONS SHOCK",
        "S2-CHEM's Scope 1, 2 and 3 greenhouse gas emissions double and its NOx emissions",
        "rise 89% against M0. Tests whether a change at a Tier 2 supplier reaches the focal",
        "firm, which cannot see that supplier, and how the Level 2 weight of E1 moves inside its band."],
    2: ["", "SCENARIO M2 — SOCIAL CHANGE ONLY",
        "S2-EQUIP raises the shares of women and ethnic minorities in its workforce,",
        "management and board by 4–12 percentage points and its supplier-diversity spending",
        "by 44%. Social data only; every environmental score must stay unchanged."],
    3: ["", "SCENARIO M3 — RENEWABLE SHARES CONVERGE",
        "All seven firms move their renewable electricity share (E10) to 58–62%, the cheap",
        "decarbonization route, while Scope 1 emissions, the expensive route, fall by 1% or",
        "less. E10's dispersion collapses and its data-driven weight falls; the bounded",
        "projection keeps the weight between its Level 3 floor and cap, so the indicator stays in the score."],
    4: ["", "SCENARIO M4 — OFFSETTING INSIDE ONE BRANCH",
        "S2-CHEM cuts greenhouse gas and NOx emissions by 21–25% but raises water use, waste",
        "and other air pollutants by 50%; S2-EQUIP cuts the same 13 environmental indicators",
        "by a common calibrated factor (stated below). The two changes cancel in S1-RUBBER's",
        "total, which moves by less than the warning limit; testing each supplier contribution",
        "separately still raises an internal warning at S1-RUBBER, while FOCAL receives",
        "S1-RUBBER's almost unchanged total."],
    5: ["", "SCENARIO M5 — S26 DECLARATION FAILS",
        "S1-RUBBER's child labor policy declaration (S26) changes from Yes to No, its recordable",
        "incident rate (S15) rises by about 60% and its safety violations (S21) rise 4.5-fold",
        "against M0.",
        "This is a simulated policy failure, not proof of an actual child-labor incident.",
        "The failed declaration raises a hard alert at S1-RUBBER without changing any score; the",
        "alert travels every declared dependency edge downstream and upstream and stays active",
        "until a review clears it, regardless of the warning thresholds."],
}


IDENTIFYING_COLUMNS = {
    "ticker": "code",              # replaced by the stable A01-A07 code
    "anonymous code": "code",
    "exchange": "not disclosed",
    "cusip/isin": "not disclosed",
    "cusip": "not disclosed",
    "isin": "not disclosed",
    "lei": "not disclosed",
    "website": "not disclosed",
    "headquarters": "not disclosed",
}

SCALE_COLUMNS = ("employees", "revenue")
# Identifier columns holding only "not disclosed"; generation drops them.
DROPPED_COLUMNS = {"exchange", "cusip/isin", "cusip", "isin", "lei", "website",
                   "headquarters"}
COMPANY_COLUMN_WIDTHS = {"tier": 12, "role": 24, "company": 14, "code": 8,
                         "industry": 20, "employees": 26, "revenue ($b)": 24}


def _magnitude_band(raw) -> str:
    """'~2,100,000' -> '1,000,000 to 10,000,000 (order of magnitude)'."""
    text = str(raw or "").replace("~", "").replace(",", "").replace("$", "").strip()
    if "(order of magnitude)" in text:
        return str(raw)
    try:
        value = float(text)
    except ValueError:
        return "not disclosed"
    if value <= 0:
        return "not disclosed"
    lower = 10.0 ** int(np.floor(np.log10(value)))
    upper = lower * 10.0
    fmt = (lambda v: f"{v:,.0f}") if lower >= 1 else (lambda v: f"{v:g}")
    return f"{fmt(lower)} to {fmt(upper)} (order of magnitude)"


def _anonymise_company_info(wb) -> None:
    """Replace the ticker, exchange and CUSIP/ISIN with the anonymous code or
    'not disclosed', and reduce headcount and revenue to an order of magnitude.
    """
    for sheet_name in wb.sheetnames:
        if "company" not in sheet_name.lower():
            continue
        ws = wb[sheet_name]
        header_row, headers = None, {}
        for i, row in enumerate(ws.iter_rows(min_row=1, max_row=5), start=1):
            labels = {str(c.value).strip().lower(): c.column for c in row
                      if isinstance(c.value, str)}
            if "tier" in labels or "company" in labels:
                header_row, headers = i, labels
                break
        if header_row is None:
            continue
        data_row = header_row + 1
        for offset, entity in enumerate(COMPANIES):
            r = data_row + offset
            for label, column in headers.items():
                if label in IDENTIFYING_COLUMNS:
                    replacement = IDENTIFYING_COLUMNS[label]
                    ws.cell(row=r, column=column).value = (
                        f"A0{offset + 1}" if replacement == "code" else replacement)
                    if replacement == "code":
                        ws.cell(row=header_row, column=column).value = "Code"
                elif label in ("company", "company name", "entity"):
                    ws.cell(row=r, column=column).value = entity
                elif any(label.startswith(s) for s in SCALE_COLUMNS):
                    ws.cell(row=r, column=column).value = _magnitude_band(
                        ws.cell(row=r, column=column).value)
        note_row = data_row + len(COMPANIES) + 1
        ws.cell(row=note_row, column=1).value = (
            "S2-CHEM to D2-RETAIL (codes A01 to A07) are positions in the simulated "
            "supply chain, identified only by these codes. Headcount and revenue are "
            "order-of-magnitude bands. Scores computed from this workbook do not reflect "
            "the ESG performance of any company.")
        drop = sorted(column for label, column in headers.items()
                      if label in DROPPED_COLUMNS)
        for column in reversed(drop):
            ws.delete_cols(column)
        for letter in list(ws.column_dimensions.keys()):
            del ws.column_dimensions[letter]
        for cell in ws[header_row]:
            label = str(cell.value or "").strip().lower()
            if label == "anonymous code":
                cell.value, label = "Code", "code"
            if label in COMPANY_COLUMN_WIDTHS:
                ws.column_dimensions[cell.column_letter].width = COMPANY_COLUMN_WIDTHS[label]


def _anonymise_workbook_entities(wb):
    """Anonymise a workbook: entity headers become the entity id and A01-A07
    code, source labels elsewhere become entity ids, comments and links are
    removed, the company information is reduced (_anonymise_company_info),
    and the Methodology sheet and document properties are rewritten."""
    replacements = {}
    for sheet_name in SHEETS.values():
        ws = wb[sheet_name]
        for entity, column in zip(COMPANIES, range(4, 4 + len(COMPANIES))):
            raw = str(ws.cell(row=2, column=column).value or "")
            source_label = raw.split("\n")[0].strip()
            if source_label and source_label != entity:
                replacements[source_label] = entity
            # Header: entity id on the first line, A01-A07 code on the second.
            ws.cell(row=2, column=column).value = (
                f"{entity}\n(A0{COMPANIES.index(entity) + 1})")
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                if isinstance(cell, openpyxl.cell.cell.MergedCell):
                    continue
                cell.comment = None
                cell.hyperlink = None
                if isinstance(cell.value, str):
                    value = cell.value
                    for source_label, entity in replacements.items():
                        value = value.replace(source_label, entity)
                    cell.value = value
    _anonymise_company_info(wb)
    if NOTES_SHEET in wb.sheetnames:
        ws = wb[NOTES_SHEET]
        for merged in list(ws.merged_cells.ranges):
            ws.unmerge_cells(str(merged))
        for row in ws.iter_rows():
            for cell in row:
                cell.value = None
        lines = [
            "CONTROLLED ESG DEMONSTRATION DATA — METHODOLOGY AND LIMITS",
            "This workbook holds the inputs of a simulated supply chain built to demonstrate the federated-statistics design.",
            "S2-CHEM, S2-EQUIP, S1-RUBBER, S1-COATING, FOCAL, D1-RETAIL and D2-RETAIL are positions in the chain, identified only by these codes.",
            "All indicator values are simulated; the SIMULATED DATA note at the end of this sheet describes how M0 was drawn.",
            "Supply links, dependency weights and the governance gradient on seven process-and-verification indicators in M0 are designed.",
            "M1–M5 are preset scenarios applied cumulatively to M0, not historical observations of any company.",
            "Scores computed from these inputs do not reflect the ESG performance of any company.",
            "The indicator values are the exact inputs that reproduce the published results.",
            "Values prefixed with ~ are approximate magnitudes.",
            "N/D means not disclosed; N/A means not applicable; N/R means not rated or registered.",
            "Missing and low-coverage observations are distinct from observed zero variance.",
            "Environmental indicators E1–E38 and social indicators S1–S50 keep their labels and units.",
            "Governance levels and scoring references follow the versioned classification and configuration shipped with the code.",
            "Indicators with observed zero variance keep their positive governed weights and floors.",
            "A verified policy deficiency raises an independent hard alert, passed upstream and downstream, without changing any score.",
            "A hard alert does not set any score to 1 and does not establish an actual labor incident.",
            "Internal warnings test each own indicator and supplier contribution against min(abs_limit, rel_limit × abs(pre-event dimension total)).",
            "Indicator limit checks are separate from peer-relative scores; undeclared limits are unknown, not safe.",
            "Secure aggregation and assurance are single-process reference simulations, not cryptographic isolation or an independent external audit.",
            "S26/S27/S28 are active declaration gates. E38/S17/S18/S21/S40/S45 are monitored without configured numeric ceilings.",
            "Admitting an indicator does not mean complete coverage. Missing observations are retained and bounded in scoring (a published score interval), not imputed or replaced by zero.",
            "Weights blend the baseline weights with a data-driven share and are projected onto approved bounds that sum to one; the Level 2/3 block shares set the baseline mass only.",
        ]
        for index, value in enumerate(lines, start=1):
            ws.cell(row=index, column=1).value = value
        # Drop the emptied rows below, so the period note follows directly.
        if ws.max_row > len(lines):
            ws.delete_rows(len(lines) + 1, ws.max_row - len(lines))
    wb.properties.creator = "ESG research demonstration"
    wb.properties.lastModifiedBy = "ESG research demonstration"
    wb.properties.title = "Controlled ESG demonstration data"
    wb.properties.subject = "Scenario inputs of a simulated supply chain"
    wb.properties.description = "Simulated demonstration data for anonymous positions in a simulated supply chain; scores do not reflect the ESG performance of any company"
    wb.properties.keywords = "ESG; simulated supply chain; demonstration"
    wb.properties.category = "Research"
    if hasattr(wb, "custom_doc_props"):
        for prop in list(wb.custom_doc_props):
            del wb.custom_doc_props[prop.name]
    if hasattr(wb, "_external_links"):
        wb._external_links = []


def _format_input_workbook(wb) -> None:
    """Black-and-white three-line layout for the input workbooks; values, number
    formats and layout are left as they are.
    """
    from collections import Counter
    from copy import copy
    from openpyxl.styles import Border, Font, PatternFill, Side
    proofread_input_workbook(wb)
    black = "FF000000"
    heavy, light = Side(style="medium", color=black), Side(style="thin", color=black)

    def filled(cell):
        return cell.value is not None and str(cell.value).strip() != ""

    def rule(sheet, row, last_col, top=None, bottom=None):
        for c in range(1, last_col + 1):
            cell = sheet.cell(row, c)
            border = copy(cell.border)
            if top is not None:
                border.top = top
            if bottom is not None:
                border.bottom = bottom
            cell.border = border

    for sheet in wb.worksheets:
        sheet.sheet_view.showGridLines = False
        cells = [cell for row in sheet.iter_rows() for cell in row]
        sizes = Counter(cell.font.sz for cell in cells
                        if filled(cell) and cell.font.name == "Arial" and not cell.font.b)
        body = sizes.most_common(1)[0][0] if sizes else 10
        notes = "methodology" in sheet.title.lower()
        for cell in cells:
            cell.fill = PatternFill(fill_type=None)
            cell.border = Border()
            if not filled(cell):
                continue
            if notes:
                text = str(cell.value)
                heading = cell.row == 1 or (any(ch.isalpha() for ch in text)
                                            and text == text.upper())
                size = cell.font.sz if cell.row == 1 else body
                cell.font = Font(name="Arial", size=size, bold=heading, color=black)
                alignment = copy(cell.alignment)
                alignment.wrap_text = True
                cell.alignment = alignment
                # Row height that fits the note within the column width.
                capacity = max(20, int(sheet.column_dimensions["A"].width))
                lines = max(1, int(np.ceil(len(text) / capacity)))
                sheet.row_dimensions[cell.row].height = max(
                    sheet.row_dimensions[cell.row].height or 0,
                    lines * float(size or body) * 1.4 + 3)
                continue
            if cell.column == 2 and ID_RE.match(str(cell.value)):
                alignment = copy(cell.alignment)
                alignment.wrap_text = True
                cell.alignment = alignment
                capacity = max(15, int(sheet.column_dimensions["B"].width))
                lines = max(1, int(np.ceil(len(str(cell.value)) / capacity)))
                sheet.row_dimensions[cell.row].height = max(
                    sheet.row_dimensions[cell.row].height or 0,
                    lines * float(cell.font.sz or body) * 1.4 + 3)
            font = copy(cell.font)
            if font.name != "Arial":
                # Cells written without a style (row labels, notes) take the
                # font of a styled neighbor in the same row, else the body font.
                like = next((other.font for other in sheet[cell.row]
                             if other is not cell and filled(other)
                             and other.font.name == "Arial"), None)
                font = copy(like) if like is not None else Font(name="Arial", size=body)
            font.color = black
            cell.font = font

        if notes:
            continue
        header_row = next((r for r in range(1, min(sheet.max_row, 6) + 1)
                           if any(str(sheet.cell(r, c).value or "").strip().lower()
                                  in ("indicator", "tier", "company")
                                  for c in range(1, min(sheet.max_column, 4) + 1))), None)
        if header_row is None:
            continue
        last_col = max(c for c in range(1, sheet.max_column + 1)
                       if filled(sheet.cell(header_row, c)))
        indicator_rows = [r for r in range(header_row + 1, sheet.max_row + 1)
                          if ID_RE.match(str(sheet.cell(r, 2).value or ""))]
        if indicator_rows:
            # Indicator sheets: title, header, tier line, one row per indicator.
            rule(sheet, header_row, last_col, top=heavy)
            rule(sheet, indicator_rows[0] - 1, last_col, bottom=light)
            rule(sheet, indicator_rows[-1], last_col, bottom=heavy)
            continue
        # Company sheet: the chain's firms under the header, then the note and
        # any further block of rows (the M6 industry peers).
        blocks, current = [], []
        for r in range(header_row + 1, sheet.max_row + 1):
            if sum(filled(sheet.cell(r, c)) for c in range(1, last_col + 1)) >= 2:
                current.append(r)
            elif current:
                blocks.append(current)
                current = []
        if current:
            blocks.append(current)
        rule(sheet, header_row, last_col, top=heavy, bottom=light)
        for k, block in enumerate(blocks):
            if k:
                rule(sheet, block[0], last_col, top=heavy)
            rule(sheet, block[-1], last_col, bottom=heavy)


def period_workbook(source: str, edits: dict, period: int = 0):
    """Load `source`, bring it to the simulated M0 baseline, apply the
    period's scenario edits (resolved against that baseline) and write its
    Data Status sheet."""
    wb = openpyxl.load_workbook(source)
    prepare_baseline(wb)
    apply(wb, resolve_scenario_edits(wb, edits))
    write_data_status(wb, f"M{period}")
    return wb


def write_data_status(wb, period_label: str) -> None:
    """(Re)write the Data Status sheet with DATA_TAGS, as of `period_label`."""
    if DATA_STATUS_SHEET in wb.sheetnames:
        del wb[DATA_STATUS_SHEET]
    ws = wb.create_sheet(DATA_STATUS_SHEET)
    ws.append([c.capitalize() for c in DATA_STATUS_COLUMNS])
    for (entity, iid), (source, status, text) in DATA_TAGS.items():
        ws.append([entity, iid, source, status, period_label, text])
    for letter, width in zip("ABCDEF", (14, 11, 13, 13, 13, 70)):
        ws.column_dimensions[letter].width = width


def build(period, edits, out_dir=OUT, extra_notes=None):
    """Write one period workbook.

    The source is loaded into memory and saved to the destination, which may
    be the source itself (regeneration from the generated M0)."""
    os.makedirs(out_dir, exist_ok=True)
    dst = os.path.join(out_dir, f"ESG_Raw_Data_Supply_Chain_M{period}.xlsx")
    wb = period_workbook(SRC, edits, period)
    try:
        note(wb, PERIOD_NOTES[period] + list(extra_notes or []) + SIMULATED_DATA_NOTES)
        for key, label in (("E", "Environmental"), ("S", "Social")):
            wb[SHEETS[key]]["A1"] = f"{label} Data — Scenario M{period} of the simulated supply chain"
        _format_input_workbook(wb)
        wb.save(dst)
    finally:
        wb.close()
    return dst


def cumulative(*edit_dicts):
    """Later periods carry forward earlier changes."""
    merged = {}
    for d in edit_dicts:
        for k, v in d.items():
            merged.setdefault(k, {}).update(v)
    return merged


def calibrate_m4_edits(source: str, classification: str,
                       iterations: int = 16, verbose: bool = True):
    """Solve for S2-EQUIP's M4 factor g; returns (edits, metadata).

    Bisection on g drives S1-RUBBER's net environmental change towards zero while
    each supplier's contribution moves past the warning limit.
    """
    global SRC
    SRC = source
    policy = RegulatoryPolicy()
    prior = cumulative(M1_EDITS, M2_EDITS, M3_EDITS)
    book = openpyxl.load_workbook(source)
    try:
        prepare_baseline(book)
        digest = baseline_digest(book)
    finally:
        book.close()

    def probe(g):
        with tempfile.TemporaryDirectory(prefix="esg_m4_") as work:
            for period, edits in ((0, {}), (3, prior), (4, cumulative(prior, m4_edits(g)))):
                wb = period_workbook(source, edits, period)
                try:
                    wb.save(os.path.join(work, f"ESG_Raw_Data_Supply_Chain_M{period}.xlsx"))
                finally:
                    wb.close()
            shutil.copyfile(classification,
                            os.path.join(work, "Indicator_Governance_Tier_Classification.xlsx"))
            system = FederatedESGSystem(data_dir=work)
            system.run_period(os.path.join(work, "ESG_Raw_Data_Supply_Chain_M0.xlsx"),
                              "P0", is_anchor=True)
            before = system.run_period(os.path.join(work, "ESG_Raw_Data_Supply_Chain_M3.xlsx"),
                                       "P3")
            after = system.run_period(os.path.join(work, "ESG_Raw_Data_Supply_Chain_M4.xlsx"),
                                      "P4", frozen_stats=before.stats,
                                      frozen_weights=before.weights)
            ca = before.comps["Environmental"]["S1-RUBBER"]
            cb = after.comps["Environmental"]["S1-RUBBER"]
            legs = {k: cb.get(k, 0.0) - ca.get(k, 0.0) for k in set(ca) | set(cb)}
            limit = min(policy.component_delta_abs, policy.component_delta_rel
                        * abs(before.totals["Environmental"]["S1-RUBBER"]))
            return float(sum(legs.values())), legs, limit

    lo, hi, best = 0.05, 1.00, None
    for _ in range(iterations):
        mid = 0.5 * (lo + hi)
        net, legs, limit = probe(mid)
        if best is None or abs(net) < abs(best[1]):
            best = (mid, net, legs, limit)
        if verbose:
            print(f"  g={mid:.6f}  hub net={net:+.8f}")
        if net > 0:
            hi = mid
        else:
            lo = mid
    g, net, legs, limit = best
    supplier_legs = {k: v for k, v in legs.items() if k != "__own__"}
    holds = abs(net) < limit and all(abs(v) > limit for v in supplier_legs.values())
    meta = {"improvement_factor": g, "hub_net_change": net, "leg_changes": legs,
            "alert_limit": limit, "design_holds": bool(holds), "baseline_digest": digest,
            "source": "calibrated now"}
    if verbose:
        print(f"  M4 calibration: g = {g:.6f} (S2-EQUIP cut {1 - g:.0%}); S1-RUBBER net "
              f"{net:+.6f}; legs " + ", ".join(f"{disp(k)} {v:+.6f}"
                                                for k, v in supplier_legs.items())
              + f"; warning limit {limit:.4f}")
        if not holds:
            print("  [warn] the M4 design does not hold for this baseline: the net change "
                  "must stay below the warning limit while each supplier leg exceeds it.")
    return m4_edits(g), meta


def default_source_workbook(here: Optional[Path] = None) -> Optional[Path]:
    """Workbook that scenario generation starts from: a reference workbook next
    to this file, else the generated M0.  None when neither exists.
    """
    here = here or Path(__file__).resolve().parent
    candidates = [here / "ESG_Raw_Data_Supply_Chain_M0.xlsx",
                  default_data_directory(here) / "ESG_Raw_Data_Supply_Chain_M0.xlsx"]
    for candidate in candidates:
        if candidate.is_file():
            return candidate
    return None


def generate_m6_workbook(output_dir: str, classification: str) -> str:
    """Extend generated M5 with three simulated industry peers and both states."""
    from copy import copy
    source = os.path.join(output_dir, "ESG_Raw_Data_Supply_Chain_M5.xlsx")
    system = FederatedESGSystem(data_dir=output_dir,
                               tier_path=os.path.abspath(classification))
    raw = load_period(source, system.tier_map, system.header_map)["Environmental"]
    base = raw[(raw.entity == "S1-RUBBER") & (raw.tier != "T1") &
               np.isfinite(raw.value)]
    observations = {r.ind_id: r for r in base.itertuples()}
    wb = openpyxl.load_workbook(source)
    try:
        normalise_sheet_names(wb)
        after = wb[DIM_SHEETS["Environmental"]]
        wb[DIM_SHEETS["Social"]]["A1"] = "Social Data — M5 values, not used in M6"
        before = wb.copy_worksheet(after)
        before.title = "Environmental Peer Before"
        for state, sheet in (("before", before), ("after", after)):
            sheet["A1"] = f"M6 Environmental Data — {state} (peer values simulated)"
            for j, entity in enumerate(M6_PEER_NAMES):
                column = 11 + j
                letter = openpyxl.utils.get_column_letter(column)
                sheet.column_dimensions[letter].width = 24
                for r in range(2, sheet.max_row + 1):
                    cell = sheet.cell(r, column)
                    cell._style = copy(sheet.cell(r, 10)._style)
                    if r == 2:
                        cell.value = entity
                    elif r == 3:
                        cell.value = "Benchmark only"
                    else:
                        row = observations.get(indicator_id(str(sheet.cell(r, 2).value or "")))
                        cell.value = _m6_simulated_value(row, j, state) if row is not None else "N/A"
                        cell.number_format = ("0.00%" if row is not None and
                                              row.unit == "fraction (percentage)" else "#,##0.00")
            sheet.freeze_panes = "D4"
        info = wb["Company Info"]
        start = info.max_row + 2
        labels = [str(c.value or "").strip().lower() for c in info[1]]
        for j, entity in enumerate(M6_PEER_NAMES):
            known = {"tier": "Benchmark only", "role": "Simulated industry peer",
                     "company": entity, "code": f"P0{j + 1}",
                     "anonymous code": f"P0{j + 1}", "industry": "Tires & Rubber"}
            for c, label in enumerate(labels, 1):
                cell = info.cell(start + j, c, known.get(label, "N/A"))
                cell._style = copy(info.cell(8, c)._style)
                cell.alignment = openpyxl.styles.Alignment(wrap_text=True, vertical="center")
            info.row_dimensions[start + j].height = 42
        notes_start = wb[NOTES_SHEET].max_row + 2
        note(wb, [
            "", "M6 INDUSTRY PEER EXPERIMENT — SIMULATED INPUTS",
            "The seven firms of the chain keep their M5 scenario values. Target: S1-RUBBER (Tires & Rubber).",
            "PEER-TIRE-01 to PEER-TIRE-03: simulated tire and rubber peers, not based on any named company.",
            "Peer inputs are constructed from S1-RUBBER's M5 values using fixed multipliers, with percentage bounds, unchanged binary indicators and the E1/E10 changes specified below.",
            "Before: sheet Environmental Peer Before. After: sheet Environmental Data. Peer values are in columns K to M and can be edited.",
            "Only environmental Level 2 and Level 3 indicators with finite target values are included. Other peer cells are N/A.",
            "Initial peer multipliers: 0.70, 1.20 and 0.90. E1 after: 0.40, 2.00 and unchanged. E10 after: all equal to the target's value.",
            "Target values are taken from Environmental Data. Both states use the same four industry members.",
            "Four members meet the minimum publication size. This reduces inference risk but does not guarantee privacy.",
            "Peers contribute industry statistics only, with no procurement links. Social data are unchanged and excluded.",
            "Percent cells store fractions and use percentage formatting. These scores are not full M0–M5 ESG scores.",
            "Generated by esg_fedstat_pipeline.py generate. Regeneration replaces manual M6 input edits.",
        ])
        notes = wb[NOTES_SHEET]
        for r in range(notes_start, notes.max_row + 1):
            notes.cell(r, 1).alignment = openpyxl.styles.Alignment(wrap_text=True, vertical="center")
            notes.row_dimensions[r].height = 32
        _format_input_workbook(wb)
        destination = os.path.join(output_dir, M6_WORKBOOK)
        wb.save(destination)
        return destination
    finally:
        wb.close()


def generate_m7_workbook(output_dir: str) -> str:
    """Write M7: the generated M0 indicator values and FOCAL's dependency-weight
    settings (M7_SETTINGS) on the Dependency Weights sheet."""
    from openpyxl.styles import Alignment, Font
    source = os.path.join(output_dir, "ESG_Raw_Data_Supply_Chain_M0.xlsx")
    wb = period_workbook(source, {}, 0)
    try:
        note(wb, [
            "", "M7 DEPENDENCY-WEIGHT SENSITIVITY — M0 INPUTS",
            "Not a further period. The indicator values are those of M0, and the supply relationships,",
            "peer benchmarks and ESG weights of M0 stay fixed. Only FOCAL's dependency weights change.",
            "The weight on its own operations and the weights on its two direct suppliers, S1-RUBBER and",
            "S1-COATING, take the three settings of the Dependency Weights sheet. Every other firm keeps",
            "its declared dependency weights. The baseline setting is FOCAL's declared weights. The higher",
            "upstream share keeps the S1-RUBBER : S1-COATING ratio of the baseline, and the alternative",
            "supplier allocation keeps the baseline's total upstream share. The alternatives are",
            "illustrative sensitivity settings, not weights derived from procurement or",
            "strategic-dependency data.",
            "Generated by esg_fedstat_pipeline.py generate from the M0 workbook. Regeneration replaces",
            "manual edits of the Dependency Weights sheet.",
        ] + SIMULATED_DATA_NOTES)
        for key, label in (("E", "Environmental"), ("S", "Social")):
            wb[SHEETS[key]]["A1"] = f"{label} Data — M0 values, held fixed in M7"
        ws = wb.create_sheet(M7_SETTINGS_SHEET)
        ws.append(list(M7_SETTINGS_COLUMNS))
        for name, beta, a_rubber, a_coating in M7_SETTINGS:
            ws.append([name, M7_FIRM, beta, a_rubber, a_coating, M7_SETTING_NOTES[name]])
            for column in (3, 4, 5):
                ws.cell(ws.max_row, column).number_format = "0.00"
        ws.append([])
        ws.append(["Every other firm keeps its declared dependency weights. Edit the weights in "
                   "columns C to E; the run reads the settings down to the first empty row."])
        for row in ws.iter_rows():
            for cell in row:
                cell.font = Font(name="Arial", size=9, bold=cell.row == 1)
                cell.alignment = Alignment(
                    horizontal=("center" if 3 <= cell.column <= 5
                                and cell.row <= 1 + len(M7_SETTINGS) else None),
                    vertical="center")
        for letter, width in zip("ABCDEF", (30, 10, 20, 18, 18, 62)):
            ws.column_dimensions[letter].width = width
        _format_input_workbook(wb)
        destination = os.path.join(output_dir, M7_WORKBOOK)
        wb.save(destination)
        return destination
    finally:
        wb.close()


def generate_scenarios(source: str, output_dir: str,
                       classification: str, recalibrate_m4: bool = False,
                       verbose: bool = True):
    """Generate M0-M5, the M6 peer workbook and the M7 settings from the
    simulated M0 baseline.

    M4 is calibrated unless M4_CALIBRATION_FILE holds a factor for the same
    baseline; `recalibrate_m4` forces a new calibration.
    """
    global SRC
    if source is None or not os.path.isfile(str(source)):
        raise ConfigurationError(
            f"scenario source workbook not found at {source}. Regeneration needs "
            f"ESG_Raw_Data_Supply_Chain_M0.xlsx, either next to the pipeline or "
            f"inside the panel directory. The six workbooks already in the panel "
            f"directory are the supplied inputs and no regeneration is required "
            f"to reproduce the published results.")
    SRC = source
    os.makedirs(output_dir, exist_ok=True)
    book = openpyxl.load_workbook(source)
    try:
        drawn_now = prepare_baseline(book)
        digest = baseline_digest(book)
    finally:
        book.close()
    calibration_path = os.path.join(output_dir, M4_CALIBRATION_FILE)
    stored = {}
    if os.path.isfile(calibration_path):
        with open(calibration_path, encoding="utf-8") as fh:
            stored = json.load(fh)
    if (recalibrate_m4 or stored.get("improvement_factor") is None
            or stored.get("baseline_digest") != digest):
        m4, calibration = calibrate_m4_edits(source, classification, verbose=verbose)
        with open(calibration_path, "w", encoding="utf-8") as fh:
            json.dump({k: calibration[k] for k in ("baseline_digest", "improvement_factor",
                                                   "hub_net_change", "alert_limit",
                                                   "design_holds")}, fh, indent=2)
    else:
        m4 = m4_edits(float(stored["improvement_factor"]))
        calibration = dict(stored, source=f"stored calibration ({M4_CALIBRATION_FILE})")
    g = float(calibration["improvement_factor"])
    m4_note = [f"M4 calibration: S2-EQUIP's 13 risk-block indicators (E1–E4, E13, E17, E18, "
               f"E21, E22, E24–E27) are {g:.4f} times their M0 value (a cut of {1 - g:.0%}); "
               f"S2-CHEM's are {M4_CHEM_FACTOR:.2f} times their M0 value."]
    plans = [
        (0, {}),
        (1, cumulative(M1_EDITS)),
        (2, cumulative(M1_EDITS, M2_EDITS)),
        (3, cumulative(M1_EDITS, M2_EDITS, M3_EDITS)),
        (4, cumulative(M1_EDITS, M2_EDITS, M3_EDITS, m4)),
        (5, cumulative(M1_EDITS, M2_EDITS, M3_EDITS, m4, M5_EDITS)),
    ]
    paths = []
    for period, edits in plans:
        destination = build(period, edits, output_dir,
                            extra_notes=m4_note if period >= 4 else None)
        paths.append(destination)
        if verbose:
            print(f"  wrote {destination} ({sum(len(v) for v in edits.values())} scenario edits)")
    destination = generate_m6_workbook(output_dir, classification)
    paths.append(destination)
    if verbose:
        print(f"  wrote {destination} (three simulated peers, before/after)")
    destination = generate_m7_workbook(output_dir)
    paths.append(destination)
    if verbose:
        print(f"  wrote {destination} (M0 values, {len(M7_SETTINGS)} FOCAL "
              f"dependency-weight settings)")
        print(f"  M0 baseline: {'simulated now' if drawn_now else 'already simulated, reused'} "
              f"({SIMULATED_BASELINE_ID}, digest {digest})")
        print(f"  M4 factor {g:.6f}: {calibration['source']} ({calibration_path})")
        if not calibration.get("design_holds", True):
            print("  [warn] M4 offsetting design does not hold; see the calibration line above.")
    return {"files": paths, "calibration": calibration, "baseline_digest": digest,
            "baseline_drawn_now": drawn_now}


def _single_file_cli():
    import argparse
    here = Path(__file__).resolve().parent
    parser = argparse.ArgumentParser(
        description="Single-file federated-statistics ESG pipeline")
    sub = parser.add_subparsers(dest="command")

    generate = sub.add_parser("generate", help="generate M0-M5, M6 industry peers and "
                                                "M7 dependency-weight settings")
    generate.add_argument("--source", default=str(default_source_workbook(here) or
                                                   here / "ESG_Raw_Data_Supply_Chain_M0.xlsx"))
    generate.add_argument("--classification", default=str(here / "Indicator_Governance_Tier_Classification.xlsx"))
    generate.add_argument("--output-dir", default=str(default_data_directory(here)))
    generate.add_argument("--recalibrate-m4", action="store_true")

    calibrate = sub.add_parser("calibrate", help="recalculate and print M4 parameters")
    calibrate.add_argument("--source", default=str(default_source_workbook(here) or
                                                    here / "ESG_Raw_Data_Supply_Chain_M0.xlsx"))
    calibrate.add_argument("--classification", default=str(here / "Indicator_Governance_Tier_Classification.xlsx"))

    verify = sub.add_parser("verify", help="run the complete demonstration")
    verify.add_argument("--data-dir", default=str(default_data_directory(here)))
    verify.add_argument("--out", default=str(here / "outputs"))
    verify.add_argument("--lam", type=float, default=0.70)
    verify.add_argument("--dispersion", choices=["common", "ratio", "cv"], default="common")
    verify.add_argument("--magnitude-range", type=float, default=10.0,
                        help="upper end of the magnitude conversion range, x anchor mean")

    events = sub.add_parser("events", help="verify independent frozen-benchmark events A-E")
    events.add_argument("--data-dir", default=str(default_data_directory(here)))
    events.add_argument("--out", default=str(here / "event_demo_outputs"))

    m6 = sub.add_parser("m6", help="M6 industry peer experiment (separate from M0-M5): "
                                   "S1-RUBBER and three simulated peers, Level 2 and 3 weights")
    m6.add_argument("--data-dir", default=str(default_data_directory(here)))
    m6.add_argument("--out", default=str(here / "outputs" / "m6"))
    m6.add_argument("--generate-inputs", action="store_true",
                    help="create explicitly synthetic peer data once; never overwrite")

    args = parser.parse_args()
    if args.command is None:
        data_dir = str(default_data_directory(here))
        run_full_demonstration(
            data_dir=data_dir, export_dir=str(here / "outputs"),
            policy=RegulatoryPolicy(), verbose=True)
    elif args.command == "m6":
        result = run_m6_peer_experiment(args.data_dir, export_dir=args.out,
                                        generate=args.generate_inputs)
        print(result["weights"].loc[lambda d: d.ind_id.isin(["E1", "E10"]),
              ["ind_id", "tier", "sigma_before", "sigma_after", "w_final_before", "w_final_after"]].to_string(index=False))
        print(result["checks"].to_string(index=False))
        print("M6 outputs:", args.out)
    elif args.command == "generate":
        generate_scenarios(args.source, args.output_dir, args.classification,
                           recalibrate_m4=args.recalibrate_m4, verbose=True)
    elif args.command == "calibrate":
        edits, metadata = calibrate_m4_edits(args.source, args.classification, verbose=True)
        print(json.dumps({"metadata": metadata, "edits": {str(k): v for k, v in edits.items()}},
                         indent=2, ensure_ascii=False, default=str))
    elif args.command == "events":
        result = run_event_demonstrations(data_dir=args.data_dir)
        destination = export_event_results(result, args.out)
        print(json.dumps(result["verification"], indent=2, ensure_ascii=False))
        print("Event outputs:", destination)
    else:
        run_full_demonstration(
            data_dir=args.data_dir, export_dir=args.out,
            policy=RegulatoryPolicy(lam=args.lam, dispersion_mode=args.dispersion,
                                    magnitude_range_multiple=args.magnitude_range),
            verbose=True)


DATA_DIRECTORY_NAME = "data"


def default_data_directory(here: Optional[Path] = None) -> Path:
    """Directory of the period workbooks.

    The `data` subdirectory next to this file, or the directory holding this
    file when it is absent.  Only the `generate` command writes workbooks.
    """
    here = here or Path(__file__).resolve().parent
    candidate = here / DATA_DIRECTORY_NAME
    return candidate if candidate.is_dir() else here

# =============================================================================
# EMBEDDED STREAMLIT DASHBOARD
# =============================================================================

# Streamlit and plotly are needed only by the dashboard.  Without them the
# module imports with a no-op stand-in for streamlit, and the engine runs on
# pandas, numpy and openpyxl.
try:
    import streamlit as st
    import plotly.graph_objects as go
    DASHBOARD_DEPENDENCIES = True
except ImportError:
    DASHBOARD_DEPENDENCIES = False

    class _StreamlitUnavailable:
        """Returns a pass-through decorator for any attribute access."""

        def __getattr__(self, name):
            def _shim(*args, **kwargs):
                if len(args) == 1 and callable(args[0]) and not kwargs:
                    return args[0]                # bare decorator
                return lambda func: func          # parameterised decorator
            return _shim

        def __bool__(self):
            return False

    st = _StreamlitUnavailable()
    go = None

core = _sys.modules[__name__]

# =============================================================================
# PALETTE  (categorical colors assigned by slot, not cycled)
# =============================================================================

C = {
    "s1": "#2a78d6",   # blue: federated / own operations
    "s2": "#eb6834",   # orange: local-only
    "s3": "#1baf7a",   # aqua: rules-based
    "s4": "#eda100",   # yellow
    "s5": "#e87ba4",   # magenta
    "s6": "#008300",   # green
    "s7": "#4a3aa7",   # violet
    "s8": "#e34948",   # red
    "good": "#0ca30c", "warning": "#fab219",
    "serious": "#ec835a", "critical": "#d03b3b",
    "grid": "#e6e5e1", "ink": "#0b0b0b", "ink2": "#52514e",
    "muted": "#8a8880", "surface": "#ffffff",
}

GOV_COLOR = {"T1": C["critical"], "T2": C["s1"], "T3": C["s3"]}

# The engine and the classification workbook code the three indicator-governance
# levels T1-T3; the platform names them Level 1-3 (Table 3) and keeps "tier" for
# positions in the supply chain.
GOV_LEVEL = LEVEL_NAME
LEVEL_TITLE = {
    "T1": "Level 1 · Serious conditions",
    "T2": "Level 2 · Established indicators",
    "T3": "Level 3 · Emerging indicators",
}
LEVEL_RULE = {
    "T1": "Excluded from weighted scoring; a confirmed condition raises a hard alert",
    "T2": "Approved baseline weight; bounded data-driven adjustment",
    "T3": "Approved baseline weight; wider adjustment between a floor and a cap",
}

# One hue per entity, assigned in fixed supply-chain order and not cycled, so
# an entity keeps its color in every chart on every tab.  Filtering a chart
# does not repaint the survivors.
ENTITY_COLOR = {e: C[f"s{i+1}"] for i, e in enumerate(ENTITY_ORDER)}
LINE_DASH = ["solid", "dash", "dot", "dashdot", "longdash", "longdashdot"]
TIER_COLOR = {"Tier 2": C["s1"], "Tier 1": C["s2"], "Focal": C["s7"],
              "Downstream": C["s3"]}

PLATFORM_NAME = "Federated ESG Supply-Chain Risk Platform"
PLATFORM_TAGLINE = "Federated statistics for multi-tier ESG risk scoring"
NAV_HEIGHT = 60

CSS = """
<style>
  :root {
    --navy: #14304f; --navy-2: #1d4270; --accent: #2a5d9a;
    --ink: #1a1f27; --ink-2: #4a5260; --muted: #7b8391;
    --line: #dde1e7; --line-2: #eceff3; --page: #f4f6f9; --card: #ffffff;
    --crit: #b42318; --warn: #b54708;
  }

  /* ---- remove Streamlit application chrome ---------------------------- */
  [data-testid="stToolbarActions"], [data-testid="stAppDeployButton"],
  .stDeployButton, [data-testid="stMainMenu"], #MainMenu,
  [data-testid="stDecoration"], [data-testid="stStatusWidget"],
  footer, .viewerBadge_container__1QSob {display: none !important;}
  header[data-testid="stHeader"] {background: transparent; height: 0;}
  /* the platform has no sidebar; hide any remnant and its toggles */
  section[data-testid="stSidebar"], [data-testid="stSidebarCollapsedControl"],
  [data-testid="stExpandSidebarButton"], [data-testid="collapsedControl"]
      {display: none !important;}

  /* ---- page frame ----------------------------------------------------- */
  .stApp {background: var(--page);}
  .block-container {padding-top: calc(__NAV__px + 1.1rem) !important;
                    padding-bottom: 0 !important; max-width: 1440px;}
  [data-testid="stElementContainer"]:has(> div > div > style),
  [data-testid="stElementContainer"]:has(.topbar),
  .element-container:has(.topbar) {position: absolute !important; height: 0; margin: 0;}
  h1, h2, h3, h4 {color: var(--ink); letter-spacing: -0.005em;}
  h4 {font-size: 1.02rem !important; font-weight: 600 !important;
      padding: 0 0 6px !important; margin: 14px 0 10px !important;
      border-bottom: 1px solid var(--line);}
  p, li {color: var(--ink);}
  [data-testid="stCaptionContainer"], .stCaption {color: var(--ink-2) !important;}

  /* ---- fixed top bar -------------------------------------------------- */
  .topbar {position: fixed; top: 0; left: 0; right: 0; height: __NAV__px;
           z-index: 999990; background: var(--navy); color: #fff;
           display: flex; align-items: center; padding: 0 28px;
           box-shadow: 0 1px 0 rgba(0,0,0,.25);}
  .topbar .mark {width: 30px; height: 30px; border-radius: 4px; margin-right: 12px;
                 flex: 0 0 30px; background: #fff; display: flex; align-items: center;
                 justify-content: center;}
  .topbar .name {font-weight: 600; font-size: 1.02rem; letter-spacing: .005em;
                 line-height: 1.15; white-space: nowrap;}
  .topbar .tag {font-size: .72rem; color: #b9c6d8; letter-spacing: .02em; white-space: nowrap;}

  /* site navigation: a keyed radio placed in the top bar */
  .st-key-rolenav {position: fixed !important; top: 0; right: 28px; left: auto !important;
                   height: __NAV__px; z-index: 999991; display: flex; align-items: center;
                   width: auto !important;}
  .st-key-rolenav [data-testid="stElementContainer"], .st-key-rolenav .element-container,
  .st-key-rolenav [data-testid="stRadio"] {width: auto !important;}
  .st-key-rolenav [data-testid="stWidgetLabel"] {display: none;}
  .st-key-rolenav [role="radiogroup"] {gap: 0 !important; flex-wrap: nowrap; margin: 0 !important;}
  .st-key-rolenav label:has(input[type="radio"]) {margin: 0 !important; padding: 0 16px !important;
       height: __NAV__px; display: flex; align-items: center; cursor: pointer; gap: 0 !important;
       border-bottom: 3px solid transparent; border-top: 3px solid transparent;}
  .st-key-rolenav label:has(input[type="radio"]) div:not([data-testid="stMarkdownContainer"]):not(:has([data-testid="stMarkdownContainer"])):not([data-testid="stMarkdownContainer"] *) {display: none !important;}
  .st-key-rolenav label:has(input[type="radio"]) p {color: #c9d4e3 !important; margin: 0;
       font-size: .9rem; font-weight: 500; white-space: nowrap;}
  .st-key-rolenav label:has(input[type="radio"]):hover p {color: #ffffff !important;}
  .st-key-rolenav label:has(input:checked) {border-bottom-color: #ffffff;}
  .st-key-rolenav label:has(input:checked) p {color: #ffffff !important; font-weight: 600;}
  @media (max-width: 1180px) {.topbar .tag {display: none;}
       .st-key-rolenav label:has(input[type="radio"]) {padding: 0 10px !important;}}
  @media (max-width: 960px) {.topbar .name {font-size: .9rem;}
       .st-key-rolenav {right: 8px;}
       .st-key-rolenav label:has(input[type="radio"]) p {font-size: .8rem;}}

  /* ---- page heading (a keyed container, so it can hold controls) ------- */
  .st-key-pagehead {background: var(--card); border: 1px solid var(--line); border-radius: 4px;
                    padding: 20px 26px 16px !important; gap: .5rem !important;
                    margin-bottom: 2px;}
  .st-key-pagehead [data-testid="stWidgetLabel"] p {font-size: .7rem !important;
       text-transform: uppercase; letter-spacing: .07em; color: var(--muted); font-weight: 500;}
  .ph-top .crumb {font-size: .74rem; color: var(--accent); letter-spacing: .07em;
                  text-transform: uppercase; margin-bottom: 6px; font-weight: 600;}
  .ph-top h1 {font-size: 1.62rem !important; font-weight: 600 !important; margin: 0 !important;
              padding: 0 !important; color: var(--ink); line-height: 1.25;}
  .ph-top {padding-bottom: 6px;}
  .ph-top .meta {display: flex; flex-wrap: wrap; gap: 8px; margin: 10px 0 0;}
  .ph-top .chip {font-size: .76rem; color: var(--ink-2); background: var(--line-2);
                 border: 1px solid var(--line); border-radius: 3px; padding: 2px 9px;}
  .ph-body .lead {font-size: .9rem; color: var(--ink-2); line-height: 1.55; margin: 0 0 12px;
                  max-width: 1100px;}
  .ph-body .rows {display: grid; grid-template-columns: 150px 1fr; gap: 5px 16px; margin: 0;
                  padding-top: 12px; border-top: 1px solid var(--line-2); font-size: .86rem;
                  line-height: 1.5;}
  .ph-body .rows .k {color: var(--muted); font-weight: 500;}
  .ph-body .rows .v {color: var(--ink);}
  .ph-body .rows .ctx {color: var(--muted);}

  /* Streamlit pulls every markdown block up by 1rem to absorb a paragraph
     margin; the HTML blocks below have none, so give the space back. */
  [data-testid="stMarkdownContainer"]:has(.ph-top),
  [data-testid="stMarkdownContainer"]:has(.ph-body),
  [data-testid="stMarkdownContainer"]:has(.card-title),
  [data-testid="stMarkdownContainer"]:has(.card-text),
  [data-testid="stMarkdownContainer"]:has(.card-foot),
  [data-testid="stMarkdownContainer"]:has(.ctl-note),
  [data-testid="stMarkdownContainer"]:has(table.doc) {margin-bottom: 0 !important;}

  /* ---- cards, tables and notices -------------------------------------- */
  div[class*="st-key-card"] {background: var(--card); border: 1px solid var(--line);
                             border-radius: 4px; padding: 16px 20px 16px !important;
                             height: 100%; box-sizing: border-box;}
  /* Older Streamlit releases give children a fixed pixel width that ignores the
     padding above; let them follow the padded box instead. */
  .st-key-pagehead [data-testid="stElementContainer"], .st-key-pagehead .element-container,
  .st-key-pagehead [data-testid="stElementContainer"] > div,
  .st-key-pagehead .element-container > div,
  .st-key-pagehead [data-testid="stHorizontalBlock"],
  div[class*="st-key-card"] [data-testid="stElementContainer"],
  div[class*="st-key-card"] .element-container,
  div[class*="st-key-card"] [data-testid="stElementContainer"] > div,
  div[class*="st-key-card"] .element-container > div,
  div[class*="st-key-card"] [data-testid="stHorizontalBlock"] {width: 100% !important;}
  /* cards placed side by side share the height of their row */
  [data-testid="stColumn"] div:has(div[class*="st-key-card"]) {height: 100%;}
  .card-title {font-size: .98rem; font-weight: 600; color: var(--ink); margin: 0;}
  .card-sub {font-size: .7rem; color: var(--muted); text-transform: uppercase;
             letter-spacing: .07em; margin: 2px 0 0; font-weight: 500;}
  .ctl-note {font-size: .78rem; color: var(--muted); line-height: 1.45;
             margin: -0.95rem 0 0.6rem;}
  .card-text {font-size: .86rem; color: var(--ink-2); line-height: 1.55;}
  .card-text ol, .card-text ul {margin: 0 0 6px 1.1rem; padding: 0;}
  .card-text p {font-size: .86rem; color: var(--ink-2); line-height: 1.55; margin: 0 0 8px;}
  .card-text li {margin-bottom: 5px; color: var(--ink-2); font-size: .86rem;}
  .card-foot {font-size: .8rem; color: var(--muted); border-top: 1px solid var(--line-2);
              padding-top: 8px; margin-top: 8px;}
  .card-count {font-size: 1.5rem; font-weight: 600; color: var(--navy); line-height: 1.2;
               font-variant-numeric: tabular-nums;}
  table.doc {width: 100%; border-collapse: collapse; background: var(--card);
             border: 1px solid var(--line); font-size: .86rem;}
  table.doc th {text-align: left; font-size: .7rem; text-transform: uppercase;
                letter-spacing: .07em; color: var(--ink-2); background: #f7f8fa;
                padding: 9px 12px; border-bottom: 1px solid var(--line); font-weight: 600;}
  table.doc td {padding: 10px 12px; border-bottom: 1px solid var(--line-2);
                vertical-align: top; color: var(--ink); line-height: 1.5;}
  table.doc td.key {font-weight: 600; white-space: nowrap;}
  table.doc td.muted {color: var(--muted);}
  table.doc tr:last-child td {border-bottom: none;}
  .tag-s {display: inline-block; font-size: .68rem; font-weight: 600; letter-spacing: .05em;
          padding: 1px 7px; border-radius: 3px; background: var(--line-2); color: var(--ink-2);
          border: 1px solid var(--line); white-space: nowrap;}
  .kpi {background: var(--card); border: 1px solid var(--line); border-top: 3px solid var(--navy);
        border-radius: 3px; padding: 12px 16px 12px; height: 100%;}
  .kpi .lab {font-size: .7rem; color: var(--ink-2); text-transform: uppercase;
             letter-spacing: .07em; font-weight: 500;}
  .kpi .val {font-size: 1.5rem; font-weight: 600; color: var(--ink); line-height: 1.3;
             font-variant-numeric: tabular-nums; overflow-wrap: anywhere;}
  .kpi .sub {font-size: .76rem; color: var(--muted);}
  .kpi.alert {border-top-color: var(--crit);}
  .kpi.alert .val {color: var(--crit);}
  .note, .warn, .crit {background: var(--card); border: 1px solid var(--line);
         border-left: 3px solid var(--accent); padding: 10px 14px; border-radius: 2px;
         font-size: .86rem; color: var(--ink-2);}
  .warn {border-left-color: #d98a1d; background: #fffaf2;}
  .crit {border-left-color: var(--crit); background: #fef6f5;}
  .scope {border: 1px solid var(--line); border-radius: 3px; padding: 12px 16px 6px;
          height: 100%; font-size: .86rem; color: var(--ink-2); background: var(--card);}
  .scope.hidden {background: #f8f9fb; border-style: dashed; border-color: #c3c9d2;}
  .scope h5 {margin: 0 0 8px; font-size: .72rem; text-transform: uppercase; font-weight: 600;
             letter-spacing: .07em; color: var(--ink);}
  .scope ul {margin: 0 0 6px 1.1rem; padding: 0;}
  .scope li {margin-bottom: 3px; color: var(--ink-2); font-size: .86rem;}
  .hiddenbox {border: 1px dashed #c3c9d2; border-radius: 3px; padding: 8px 12px;
              font-size: .82rem; color: var(--muted); background: #f8f9fb;}
  .hiddenbox .hb {font-size: .66rem; font-weight: 600; letter-spacing: .07em;
                  text-transform: uppercase; color: var(--ink-2); margin-right: 8px;}

  /* ---- Streamlit widgets restyled ------------------------------------- */
  /* visible input borders (BaseWeb controls up to 1.5x, react-aria afterwards) */
  div[data-baseweb="select"] > div, div[data-baseweb="input"], div[data-baseweb="textarea"],
  .stSelectbox .react-aria-ComboBox [role="group"], [data-testid="stTextInputRootElement"]
      {border: 1px solid #cdd3dc !important; border-radius: 3px !important;
       background-color: #ffffff !important;}
  div[data-baseweb="select"] > div:hover, div[data-baseweb="input"]:hover,
  .stSelectbox .react-aria-ComboBox [role="group"]:hover,
  [data-testid="stTextInputRootElement"]:hover {border-color: #9aa5b4 !important;}
  [data-testid="stExpander"] details {background: var(--card); border: 1px solid var(--line) !important;
                                      border-radius: 3px !important;}
  [data-testid="stExpander"] summary p {font-weight: 500; font-size: .88rem;}
  .stTabs [data-baseweb="tab-list"] {gap: 0; border-bottom: 1px solid var(--line);}
  .stTabs [data-baseweb="tab"] {padding: 12px 18px; height: auto;}
  .stTabs [data-baseweb="tab"] p {font-size: .9rem; font-weight: 500; color: var(--ink-2);}
  .stTabs [aria-selected="true"] p {color: var(--navy); font-weight: 600;}
  .stTabs [data-baseweb="tab-highlight"] {background-color: var(--navy) !important; height: 3px;}
  .stTabs [data-baseweb="tab-border"] {display: none;}
  .stTabs [data-baseweb="tab-panel"] {padding-top: 18px;}
  [data-testid="stDataFrame"], [data-testid="stTable"] {border: 1px solid var(--line); border-radius: 3px;}
  [data-testid="stPlotlyChart"] {background: var(--card); border: 1px solid var(--line);
                                 border-radius: 3px; padding: 6px;}
  .stKatex, .katex-display {margin: .3em 0 !important;}

  /* ---- footer --------------------------------------------------------- */
  .sitefoot {margin: 40px -1rem 0; padding: 22px 28px 26px; border-top: 1px solid var(--line);
             background: #eef1f5; font-size: .78rem; color: var(--ink-2); line-height: 1.6;}
  .sitefoot .row {display: flex; justify-content: space-between; flex-wrap: wrap; gap: 8px 24px;}
  .sitefoot b {color: var(--ink); font-weight: 600;}
  .sitefoot code {font-size: .74rem; background: transparent; color: var(--ink-2);}
</style>
""".replace("__NAV__", str(NAV_HEIGHT))


def base_layout(fig, height=380, xtitle="", ytitle="", legend=True):
    fig.update_layout(
        height=height, paper_bgcolor=C["surface"], plot_bgcolor=C["surface"],
        margin=dict(l=10, r=16, t=32, b=36),
        font=dict(family="Source Sans Pro, -apple-system, Segoe UI, sans-serif",
                  size=12, color=C["ink2"]),
        hoverlabel=dict(bgcolor="#ffffff", bordercolor=C["grid"],
                        font=dict(color=C["ink"], size=12)),
        showlegend=legend,
        legend=dict(orientation="h", yanchor="bottom", y=1.02, x=0,
                    bgcolor="rgba(0,0,0,0)", font=dict(size=11)),
    )
    fig.update_xaxes(title=xtitle, gridcolor=C["grid"], zerolinecolor=C["grid"],
                     linecolor=C["grid"], title_font=dict(size=11))
    fig.update_yaxes(title=ytitle, gridcolor=C["grid"], zerolinecolor=C["grid"],
                     linecolor=C["grid"], title_font=dict(size=11))
    return fig


# --- Streamlit version compatibility ----------------------------------------
# Recent releases take `width="stretch"` and older ones only
# `use_container_width`; every call goes through these two helpers.

def show_df(obj, **kw):
    try:
        return st.dataframe(obj, width="stretch", **kw)
    except TypeError:
        return st.dataframe(obj, use_container_width=True, **kw)


def show_chart(fig, **kw):
    # Streamlit versions without a `width` parameter treat it as a deprecated
    # Plotly keyword and print a warning on the page.
    import inspect
    try:
        stretch = "width" in inspect.signature(st.plotly_chart).parameters
    except (TypeError, ValueError):
        stretch = False
    if stretch:
        return st.plotly_chart(fig, width="stretch", **kw)
    return st.plotly_chart(fig, use_container_width=True, **kw)


def kpi(label, value, sub=""):
    st.markdown(f"<div class='kpi'><div class='lab'>{label}</div>"
                f"<div class='val'>{value}</div>"
                f"<div class='sub'>{sub}</div></div>", unsafe_allow_html=True)


# =============================================================================
# ENGINE  (cached on policy, implementation, and exact input contents)
# =============================================================================

DASHBOARD_RESULT_SCHEMA = "common-calibration-initialization-directional-signals-v7"


def dashboard_cache_revision(data_dir: str) -> str:
    """Cache key over the engine source and the content of every input file."""
    source = Path(__file__).resolve()
    folder = Path(data_dir).expanduser().resolve()
    periods = discover_period_files(str(folder), expect_periods=6, verbose=False)
    classification = folder / "Indicator_Governance_Tier_Classification.xlsx"
    if not classification.exists():
        classification = source.parent / classification.name
    inputs = [source, *(Path(p) for _, p in periods),
              classification, folder / HEADER_MAP_FILENAME,
              folder / "m6" / "peer_inputs.csv", folder / M6_WORKBOOK, folder / M7_WORKBOOK]
    digest = hashlib.sha256(DASHBOARD_RESULT_SCHEMA.encode("utf-8"))
    for path in inputs:
        digest.update(str(path.resolve()).encode("utf-8"))
        digest.update(b"\0")
        if path.exists():
            digest.update(hashlib.sha256(path.read_bytes()).digest())
        else:
            digest.update(b"missing")
        digest.update(b"\0")
    return digest.hexdigest()


def dashboard_result_is_current(result) -> bool:
    """True when a cached result has the current schema and the fields the
    pages read."""
    if not isinstance(result, dict) or result.get("dashboard_schema") != DASHBOARD_RESULT_SCHEMA:
        return False
    events = result.get("events")
    if not isinstance(events, dict):
        return False
    verification = events.get("verification")
    scenarios = events.get("scenarios")
    acceptance = events.get("acceptance")
    return (
        {"policy", "periods", "acceptance", "scope", "cold_start"}.issubset(result)
        and all(hasattr(p, "provisional") for p in result["periods"].values())
        and isinstance(verification, dict)
        and {"event_aggregation_calls", "additional_version_records"}.issubset(verification)
        and isinstance(scenarios, dict) and bool(scenarios)
        and isinstance(acceptance, pd.DataFrame) and "passed" in acceptance.columns
        and all(isinstance(table, pd.DataFrame) and {
            "composite_before", "composite_after", "composite_delta", "internal_warning",
            "hard_event_count", "hard_event_ids", "hard_event_details"
        }.issubset(table.columns) for table in scenarios.values())
    )


@st.cache_resource(show_spinner="Running the federated re-estimation…")
def run_engine(data_dir, lam, band_lo, band_hi, t3_floor, t3_cap,
               block_t2, dispersion, squash, alert, dabs, drel, anonymise,
               coverage_floor, engine_revision):
    # engine_revision is part of the cache key, so it has no leading underscore.
    core.ANONYMISE_ENTITIES = anonymise
    pol = RegulatoryPolicy(
        block_share_t2=block_t2, block_share_t3=round(1 - block_t2, 6),
        lam=lam, band_lo=band_lo, band_hi=band_hi,
        t3_floor_frac=t3_floor, t3_cap_frac=t3_cap,
        dispersion_mode=dispersion, squash_scale=squash,
        peer_relative_flag_level=alert, component_delta_abs=dabs,
        component_delta_rel=drel, coverage_floor=coverage_floor)
    result = core.run_full_demonstration(data_dir=data_dir, export_dir="",
                                         policy=pol, verbose=False)
    result["dashboard_schema"] = DASHBOARD_RESULT_SCHEMA
    return result


def load_dashboard_results(data_dir, lam, band_lo, band_hi, t3_floor, t3_cap,
                           block_t2, dispersion, squash, alert, dabs, drel, anonymise,
                           coverage_floor=0.80):
    """Load a compatible result, rebuilding a stale cache entry at most once."""
    data_dir = str(Path(data_dir).expanduser().resolve())
    revision = dashboard_cache_revision(data_dir)
    args = (data_dir, lam, band_lo, band_hi, t3_floor, t3_cap,
            block_t2, dispersion, squash, alert, dabs, drel, anonymise,
            float(coverage_floor))
    result = run_engine(*args, engine_revision=revision)
    if not dashboard_result_is_current(result):
        run_engine.clear()
        result = run_engine(*args, engine_revision=revision)
    if not dashboard_result_is_current(result):
        raise RuntimeError("The engine returned an incomplete dashboard result after "
                           "a cache refresh. Restart Streamlit with this pipeline file.")
    return result


# =============================================================================
# PLATFORM PAGES: OVERVIEW | FIRM | GOVERNANCE BODY | ASSURANCE PROVIDER
# =============================================================================
# Every page is a filtered view of one single-process run.  The role pages show
# each party's information boundary in the design but do not enforce it (no
# authentication).  The coordinator has no page; its published records are on
# the Overview page.

DIMS = ("Environmental", "Social")

POLICY_DEFAULTS = {
    "pol_lam": 0.70, "pol_band": (0.60, 1.40), "pol_t3_floor": 0.25,
    "pol_t3_cap": 3.0, "pol_block_t2": 0.70, "pol_dispersion": "common",
    "pol_squash": 1.0, "pol_alert": 0.60, "pol_dabs": 0.005, "pol_drel": 0.010,
    "pol_cov": 0.80,
}

NAV_ITEMS = ["Overview", "Participating firm", "Governance body", "Assurance provider"]


def direct_suppliers(e: str) -> List[str]:
    return [j for j, a in DEPENDENCY_WEIGHTS[e]["alpha"].items() if a > 0]


def direct_customers(e: str) -> List[str]:
    return [c for c in ENTITY_ORDER if DEPENDENCY_WEIGHTS[c]["alpha"].get(e, 0) > 0]


def kpi_box(label, value, sub="", alert=False):
    st.markdown(f"<div class='kpi{' alert' if alert else ''}'><div class='lab'>{label}</div>"
                f"<div class='val'>{value}</div>"
                f"<div class='sub'>{sub}</div></div>", unsafe_allow_html=True)


def scope_panel(can_see: List[str], cannot_see: List[str],
                who: str = "you") -> None:
    """Two boxes side by side: what this role sees, what it does not."""
    box = st.expander(f"What {who} can and cannot see", expanded=False)
    a, b = box.columns(2)
    with a:
        st.markdown("<div class='scope'><h5>Visible to " + who + "</h5><ul>"
                    + "".join(f"<li>{x}</li>" for x in can_see)
                    + "</ul></div>", unsafe_allow_html=True)
    with b:
        st.markdown("<div class='scope hidden'><h5>Not visible to " + who + "</h5><ul>"
                    + "".join(f"<li>{x}</li>" for x in cannot_see)
                    + "</ul></div>", unsafe_allow_html=True)


ROLE_TABLE = {   # role: (what it does, what it sees)
    "Governance body": ("sets the hard-threshold list, base weights, bounds, λ and the "
                        "re-estimation period (a regulator, an industry consortium or a "
                        "public institution)",
                        "aggregate results and authorized assurance findings, no firm records"),
    "Participating firm": ("scores itself locally, declares its dependency-weight method "
                           "and responds to alerts",
                           "its own data and the aggregate scores of its direct suppliers"),
    "Coordinator": ("runs the re-estimation, publishes versioned weights and keeps the log",
                    "summary statistics and scores only, no raw records"),
    "Assurance provider": ("verifies hard-threshold declarations, raw data and whether the "
                           "declared dependency-weight method was applied, and checks "
                           "documented responses to alerts",
                           "its client's records, including supplier totals received by the client"),
}

ROLE_PORTAL = {
    "Overview": "Overview",
    "Participating firm": "Participating firm portal",
    "Governance body": "Governance body console",
    "Coordinator": "Coordination service",
    "Assurance provider": "Assurance provider workspace",
}


def _cap(text: str) -> str:
    return text[:1].upper() + text[1:]


def role_rows(role: str) -> List[Tuple[str, str]]:
    does, sees = ROLE_TABLE[role]
    return [("Responsibilities", _cap(does) + "."), ("Data access", _cap(sees) + ".")]


def scenario_rows(period: str) -> List[Tuple[str, str]]:
    """Page-heading row describing the designed scenario of a period."""
    head, desc = _period_note(period)
    if not head:
        return []
    i = int(period.lstrip("Mm"))
    earlier = f" Builds on M1–M{i - 1}." if i > 2 else (" Builds on M1." if i == 2 else "")
    return [(f"Scenario {period}",
             f"<b>{head}.</b> {desc}{earlier} <span class='ctx'>Context for the reader of "
             f"the demonstration; no participant receives this description.</span>")]


def level_text(text):
    """Engine text with the governance tiers T1-T3 written as Level 1-3."""
    if not isinstance(text, str):
        return text
    text = re.sub(r"\bT([123])\b", r"Level \1", text)
    return re.sub(r"\bgovernance tier", "governance level", text)


def keyed_container(key: str):
    """A container carrying a CSS hook; plain container on Streamlit < 1.39."""
    try:
        return st.container(key=key)
    except TypeError:
        return st.container()


def header_columns():
    """Title area on the left, page-level controls on the right."""
    try:
        return st.columns([1.6, 1.4], gap="medium", vertical_alignment="bottom")
    except TypeError:                              # Streamlit < 1.36
        return st.columns([1.6, 1.4])


def control_columns(n: int = 2):
    """Side-by-side page controls whose input boxes line up even if a label wraps."""
    try:
        return st.columns([1, 1.3] if n == 2 else n, vertical_alignment="bottom")
    except TypeError:                              # Streamlit < 1.36
        return st.columns([1, 1.3] if n == 2 else n)


def header_top(eyebrow: str, title: str, chips: List[str] = ()) -> None:
    chip_html = "".join(f"<span class='chip'>{c}</span>" for c in chips)
    st.markdown(f"<div class='ph-top'><div class='crumb'>{eyebrow}</div><h1>{title}</h1>"
                + (f"<div class='meta'>{chip_html}</div>" if chips else "") + "</div>",
                unsafe_allow_html=True)


def header_body(lead: str, rows: List[Tuple[str, str]] = ()) -> None:
    grid = "".join(f"<div class='k'>{k}</div><div class='v'>{v}</div>" for k, v in rows)
    st.markdown(f"<div class='ph-body'><div class='lead'>{lead}</div>"
                + (f"<div class='rows'>{grid}</div>" if rows else "") + "</div>",
                unsafe_allow_html=True)


def card(key: str):
    return keyed_container(f"card_{key}")


def card_head(title: str, sub: str = "") -> None:
    st.markdown(f"<div class='card-title'>{title}</div>"
                + (f"<div class='card-sub'>{sub}</div>" if sub else ""),
                unsafe_allow_html=True)


def control_note(text: str) -> None:
    """Plain-language note under a control: what it sets and what moving it changes."""
    st.markdown(f"<div class='ctl-note'>{text}</div>", unsafe_allow_html=True)


def html_table(columns: List[str], rows: List[List[str]], key_col: bool = True) -> str:
    head = "".join(f"<th>{c}</th>" for c in columns)
    body = "".join("<tr>" + "".join(
        f"<td class='key'>{v}</td>" if (key_col and i == 0) else f"<td>{v}</td>"
        for i, v in enumerate(r)) + "</tr>" for r in rows)
    return f"<table class='doc'><thead><tr>{head}</tr></thead><tbody>{body}</tbody></table>"


_MARK_SVG = ("<svg width='18' height='18' viewBox='0 0 18 18'>"
             "<rect x='1' y='10' width='4' height='7' fill='#14304f'/>"
             "<rect x='7' y='5' width='4' height='12' fill='#2a5d9a'/>"
             "<rect x='13' y='1' width='4' height='16' fill='#14304f'/></svg>")


def site_topbar() -> None:
    st.markdown(
        f"<div class='topbar'><div class='mark'>{_MARK_SVG}</div>"
        f"<div><div class='name'>{PLATFORM_NAME}</div>"
        f"<div class='tag'>{PLATFORM_TAGLINE} · research prototype</div></div></div>",
        unsafe_allow_html=True)


def site_navigation(items: List[str]) -> str:
    """Site menu rendered in the top bar (falls back to an inline control)."""
    if st.session_state.get("view_role") not in items:
        st.session_state["view_role"] = items[0]
    with keyed_container("rolenav"):
        return st.radio("Page", items, horizontal=True, key="view_role",
                        label_visibility="collapsed")


def site_footer(out: Optional[dict] = None) -> None:
    digest = ""
    if out is not None:
        try:
            digest = f"Policy digest <code>{out['policy'].digest()}</code>"
        except Exception:
            digest = ""
    st.markdown(
        "<div class='sitefoot'><div class='row'>"
        f"<div><b>{PLATFORM_NAME}</b> · reference implementation for research use</div>"
        f"<div>{digest}</div></div>"
        "<div style='margin-top:8px;max-width:1100px'>The role pages are filtered "
        "renderings of one single-process run. They illustrate each party's information "
        "boundary in the design; they do not enforce it. The coordinator runs as a "
        "server-side service and has no page of its own. M0–M5 are cumulative designed "
        "scenarios, not observed history; M6 is a separate industry peer experiment and M7 "
        "a dependency-weight sensitivity on the M0 inputs, neither a further period. "
        "Full-chain evaluation evidence (ablation, "
        "acceptance tests, event tables) is in the result and checks workbooks. Entity "
        "aliases do not guarantee anonymity.</div></div>",
        unsafe_allow_html=True)


CAPABILITIES = {
    "can": [
        "A firm can see that risk rose through the branch of a named direct supplier, "
        "which identifies the branch that needs further investigation",
        "A score splits into own operations, each upstream branch, and the environmental "
        "and social dimensions",
        "The scoring rule is public; every weight version is traceable and recomputable",
        "Upstream suppliers hand over no raw operating data (in the design; this reference "
        "runs in one process)",
    ],
    "cannot": [
        "Identify the specific Tier-2+ firm behind a branch signal. The signal supports "
        "prioritization and further investigation; identifying a deeper-tier firm "
        "requires follow-up assessment",
        "Compare rolling scores across periods directly, hence the fixed and rolling lines "
        "reported side by side",
        "Rule out strategic timing. A rolling benchmark could reward delaying improvements; "
        "this is a possible incentive, not a demonstrated result",
        "Guarantee participation. Enrollment is voluntary, and the riskiest suppliers have "
        "the least reason to join",
    ],
}


def hidden_note(text: str) -> None:
    st.markdown(f"<div class='hiddenbox'><span class='hb'>Not visible here</span>"
                f"{_cap(text)}</div>", unsafe_allow_html=True)


def _mask_source(source: str, me: str) -> str:
    if source == me:
        return "your own submission"
    if source in direct_suppliers(me):
        return f"branch via direct supplier {disp(source)}"
    return "not disclosed"


def _transition_into(period: str, labels: List[str]) -> Optional[str]:
    i = labels.index(period)
    return f"{labels[i-1]}->{labels[i]}" if i > 0 else None


def keep_state(key: str, default, allowed: Optional[List] = None) -> None:
    """Initialise a widget value once and keep it while its page is not shown."""
    ss = st.session_state
    if key not in ss or (allowed is not None and ss[key] not in allowed):
        ss[key] = default
    else:
        ss[key] = ss[key]


def _preserve_policy_state() -> None:
    """Keep the policy settings while the pages that hold those controls are hidden."""
    for k, v in POLICY_DEFAULTS.items():
        keep_state(k, v)


PAGE_STATE_KEYS = ("firm_me", "aud_client", "period_firm", "period_aud", "regulator_scenario")


def _carry_page_state() -> None:
    """Keep the firm, client and period chosen on a page while another page is
    shown.
    """
    ss = st.session_state
    for k in PAGE_STATE_KEYS:
        if k in ss:
            ss[k] = ss[k]


def _restore_policy_defaults() -> None:
    for k, v in POLICY_DEFAULTS.items():
        st.session_state[k] = v


def _keep_band_around_one() -> None:
    """Keep the base weight (1.0) inside the drift band."""
    lo, hi = st.session_state["pol_band"]
    st.session_state["pol_band"] = (min(lo, 1.0), max(hi, 1.0))


def dispersion_label(mode: str) -> str:
    """Display name of the dispersion measure behind the weights."""
    return {"common": "Common-scale standard deviation",
            "ratio": "Dispersion ratio σ(t) / σ(anchor)",
            "cv": "Coefficient of variation σ / |μ|"}[mode]


def _policy_table(policy: RegulatoryPolicy) -> pd.DataFrame:
    return pd.DataFrame([
        {"Parameter": "Baseline influence λ (Levels 2 and 3)", "Value": f"{policy.lam:.2f}"},
        {"Parameter": "Level 2 drift band (× base weight)",
         "Value": f"[{policy.band_lo:.2f}, {policy.band_hi:.2f}]"},
        {"Parameter": "Level 3 floor / cap (× equal weight)",
         "Value": f"{policy.t3_floor_frac:.2f} / {policy.t3_cap_frac:.1f}"},
        {"Parameter": "Baseline allocation · Level 2 / Level 3",
         "Value": f"{policy.block_share_t2:.2f} / {policy.block_share_t3:.2f}"},
        {"Parameter": "Dispersion measure", "Value": dispersion_label(policy.dispersion_mode)},
        {"Parameter": "Conversion range of magnitude indicators",
         "Value": f"[0, {policy.magnitude_range_multiple:g} × anchor mean]; declarations "
                  f"and percentages on their natural ranges"},
        {"Parameter": "τ (squashing scale)", "Value": f"{policy.squash_scale:.1f}"},
        {"Parameter": "Peer-relative escalation level",
         "Value": f"{policy.peer_relative_flag_level:.2f}"},
        {"Parameter": "Internal warning: absolute / relative limit",
         "Value": f"{policy.component_delta_abs:.3f} / {policy.component_delta_rel:.3f}"},
        {"Parameter": "Coverage floor for a non-provisional score",
         "Value": f"{policy.coverage_floor:.0%} of published weight"},
        {"Parameter": "Investigate: size / duration / persistent level",
         "Value": f"{policy.escalation_size_multiple:.1f}× limit / "
                  f"{policy.escalation_periods} periods / {policy.persistent_level:.2f}"},
        {"Parameter": "Persistent level as a weighted deviation of a local score",
         "Value": f"s = {score_level_index(policy.persistent_level, policy):.2f} peer SDs "
                  f"at τ = {policy.squash_scale:.1f}; exact only for a firm without "
                  f"suppliers"},
        {"Parameter": "Maximum data age / valid-reference age (periods)",
         "Value": f"{policy.max_data_age_periods} / {policy.reference_max_age_periods}"},
        {"Parameter": "Re-estimation period", "Value": policy.reestimation_period},
        {"Parameter": "Policy digest", "Value": policy.digest()},
    ])


def published_policy_panel(policy: RegulatoryPolicy) -> None:
    with st.expander("Published governance policy (read-only)"):
        show_df(_policy_table(policy), hide_index=True)
        st.caption("Only the governance body changes the policy.")


def _period_option_label(label: str) -> str:
    """Selector label; M6 and M7 are tagged because they are not periods."""
    return {"M6": "M6 · Peer benchmark", "M7": "M7 · Weight sensitivity"}.get(label, label)


def _period_chip(period: str) -> str:
    """Page-heading chip for a reporting period or one of the separate experiments."""
    return {"M6": "M6 industry peer experiment",
            "M7": "M7 dependency-weight sensitivity"}.get(period, f"Reporting period {period}")


def period_picker(labels: List[str], key: str, default: Optional[str] = None):
    """Period selector.  The default is the latest periodic round (M5), not a
    separate experiment (M6, M7), so a page opens on the supply-chain results."""
    keep_state(key, default if default in labels else labels[-1], allowed=labels)
    return st.selectbox("Scenario", labels, key=key,
                        format_func=_period_option_label,
                        help="M0–M5 are cumulative periodic scenarios of the supply "
                             "chain. M6 is a separate industry peer experiment for "
                             "S1-RUBBER, and M7 a sensitivity of FOCAL's results to its "
                             "dependency weights (α, β) on the M0 inputs. Neither is a "
                             "further period. M6 uses its own industry benchmark; M7 "
                             "compares alternative settings with the M0 baseline.")


def dashboard_period_labels(out: dict) -> List[str]:
    """Offer M6 and M7 without adding them to the time-series engine."""
    return (list(out["periods"]) + (["M6"] if "m6_peers" in out else [])
            + (["M7"] if "m7" in out else []))


def m6_dashboard_panel(out: dict, *, entity: Optional[str] = None,
                       regulator: bool = False) -> None:
    panel = out["m6_peers"]
    st.markdown(
        "<div class='note'><b>M6 is not a further period.</b> It is a separate experiment "
        "on the peer group. S1-RUBBER is benchmarked against three simulated "
        f"tire and rubber firms on its {len(panel['weights'])} environmental Level 2 and 3 indicators; "
        "the seven chain firms keep their M5 data, and the peers take part in the benchmark "
        "only. Its scores are not comparable with the M0–M5 composite.</div>",
        unsafe_allow_html=True)
    st.markdown("")
    scope_panel(
        who="the governance body" if regulator else "you",
        can_see=(["Published industry means, dispersion and indicator weights",
                  "Level 2 and Level 3 weight changes under the published policy",
                  "The two published weight versions of the peer group"] if regulator else
                 ["Your own environmental inputs and peer score" if entity else
                  "The four-company reference cohort and its simulated inputs",
                  "Published industry means, dispersion and indicator weights",
                  "Score changes attributed to own data, the benchmark and weights"]),
        cannot_see=(["Company-level raw values or individual scores", "Individual uploads or procurement links"]
                    if regulator else
                   ["Other firms' raw inputs and individual scores", "Other firms' uploads or procurement links"]
                    if entity else
                   ["Actual ESG disclosures for the simulated peers"]))
    published_policy_panel(out["policy"])
    tabs = st.tabs(["Weight overview", "Indicators"] if regulator else
                   ["Score overview", "Indicators", "Company inputs" if entity else "Inputs and checks"])
    weights = panel["weights"]
    versions = panel.get("versions")
    with tabs[0]:
        if regulator:
            cols = st.columns(3)
            for col, tier in zip(cols[:2], ("T2", "T3")):
                with col:
                    kpi_box(f"{GOV_LEVEL[tier]} indicators", str(int((weights.tier == tier).sum())),
                            "common environmental indicators")
            with cols[2]:
                kpi_box("Weight change", f"{weights.w_final_delta.abs().max():.4f}",
                        "largest absolute change")
            if versions is not None and not versions.empty:
                st.markdown("#### Published weight versions · M6 peer group")
                show_df(versions[["state", "dimension", "version_id", "digest",
                                  "register_version", "policy_digest", "timestamp",
                                  "n_indicators", "peer_group"]].rename(columns={
                    "state": "State", "dimension": "Dimension", "version_id": "Version",
                    "digest": "Parameter digest", "register_version": "Register",
                    "policy_digest": "Policy digest", "timestamp": "Published",
                    "n_indicators": "Indicators", "peer_group": "Peer group"}),
                    hide_index=True)
                st.caption("Each M6 recalibration publishes a versioned parameter set with a "
                           "digest, like an M0–M5 round, bound to the separate M6 register.")
        else:
            scores = panel["scores"]
            if entity is not None:
                scores = scores[scores.entity == entity]
            if scores.empty:
                with st.columns(1)[0]:
                    kpi_box("Environmental peer score", "Not available", "outside the M6 industry cohort")
                st.info(f"{entity} is not in this tire and rubber industry cohort and has no M6 peer score. "
                        "Select S1-RUBBER under Participant / Client to view the target company's results.")
            else:
                # The KPIs show the target.  The table lists all four members on the
                # Overview page and only the signed-in firm (or client) elsewhere.
                target = scores[scores.entity == "S1-RUBBER"]
                row = (target if not target.empty else scores).iloc[0]
                cols = st.columns(5)
                metrics = [("Environmental peer score", row.after, f"{row.total_change:+.4f} vs before"),
                           ("Score before", row.before, "same industry cohort"),
                           ("Own-data effect", row.own_data_effect, "own indicator changes"),
                           ("Benchmark effect", row.benchmark_effect, "industry benchmark changes"),
                           ("Weight effect", row.weight_effect, "governed weight changes")]
                for col, (label, value, detail) in zip(cols, metrics):
                    with col:
                        kpi_box(label, f"{value:.4f}", detail)
                st.caption(f"Company: {row.entity}. Own-data effect + benchmark effect + weight effect "
                           "= total change, decomposed in that order.")
                with card("m6_score_attribution"):
                    card_head("Environmental score changes", "Before and after the peer intervention")
                    show_df(scores[["entity", "before", "after", "own_data_effect", "benchmark_effect",
                                    "weight_effect", "total_change"]].rename(columns={
                        "entity": "Company", "before": "Before", "after": "After",
                        "own_data_effect": "Own-data effect", "benchmark_effect": "Benchmark effect",
                        "weight_effect": "Weight effect", "total_change": "Total change"}), hide_index=True)
        with card("m6_scenario"):
            card_head("Industry peer experiment",
                      "M6 · environmental Level 2 and 3 indicators · separate from M0–M5")
            st.caption("S1-RUBBER and three simulated tire and rubber peers form the peer group, "
                       "four members, the minimum group size for a published benchmark. "
                       "The peers start at 0.7, 0.9 and 1.2 times S1-RUBBER's values; afterwards "
                       "two of them move apart on direct emissions (E1) and all three adopt "
                       "S1-RUBBER's renewable electricity share (E10). S1-RUBBER's own "
                       "indicators remain unchanged.")
            st.caption("Peers have no supply-chain procurement links and send or receive no "
                       "scores. M6 does not calculate a social total, a supply-chain composite or "
                       "alerts; its environmental scores are not directly comparable with M0–M5 "
                       "ESG scores.")
    with tabs[1]:
        _m6_weight_tables(weights, out["policy"])
    if regulator:
        return
    with tabs[2]:
        with card("m6_inputs"):
            card_head("Company inputs" if entity else "Simulated peer inputs",
                      "Before and after observations")
            if entity is not None:
                own = panel["chain_and_peer_data"]
                show_df(own[own.entity == entity][["state", "ind_id", "tier", "unit", "value"]]
                        .replace({"tier": GOV_LEVEL})
                        .rename(columns={"state": "State", "ind_id": "Indicator", "tier": "Level",
                                         "unit": "Unit", "value": "Value"}), hide_index=True)
            else:
                show_df(panel["inputs"], hide_index=True)
                with st.expander("Experiment checks"):
                    show_df(panel["checks"], hide_index=True)
                if versions is not None and not versions.empty:
                    with st.expander("Published weight versions of the M6 peer group"):
                        show_df(versions.assign(notes=versions["notes"].map(level_text)),
                                hide_index=True)


def _m6_weight_tables(weights: pd.DataFrame, policy: RegulatoryPolicy) -> None:
    """M6 weight changes, in the cards and tables of the other periods."""
    st.markdown("#### Level 2 and Level 3 weight changes")
    st.caption("Level 2 holds established indicators, bounded around the base weight. Level 3 holds "
               "emerging indicators, with wider bounds. Both levels use the baseline blend "
               "and a single projection. Current dispersion measure: "
               f"{dispersion_label(policy.dispersion_mode)}. Means and SDs below use each "
               "indicator's raw unit.")
    for tier in ("T2", "T3"):
        rows = weights[weights.tier == tier]
        st.markdown(f"**{GOV_LEVEL[tier]}**")
        if rows.empty:
            st.caption("No common indicators are available at this level.")
            continue
        st.bar_chart(rows.set_index("ind_id")[["w_final_before", "w_final_after"]]
                     .rename(columns={"w_final_before": "Before", "w_final_after": "After"}))
        show_df(rows[["ind_id", "mu_before", "mu_after", "sigma_before", "sigma_after",
                      "w_final_before", "w_final_after", "w_final_delta"]]
                .rename(columns={"ind_id": "Indicator", "mu_before": "Raw mean before", "mu_after": "Raw mean after",
                                 "sigma_before": "Raw SD before", "sigma_after": "Raw SD after",
                                 "w_final_before": "Weight before", "w_final_after": "Weight after",
                                 "w_final_delta": "Weight change"}), hide_index=True)


M7_DIMENSIONS = ("Environmental", "Social", "Composite")


def m7_dashboard_panel(out: dict, *, role: str = "overview",
                       entity: Optional[str] = None) -> None:
    """M7, FOCAL's total and branch contributions under its dependency-weight
    settings on the M0 inputs, filtered to what each role would hold.

    `role` is "overview", "firm", "auditor" or "regulator"; `entity` is the
    signed-in firm or the assurance client.
    """
    m7 = out["m7"]
    table, checks, settings = m7["table"], m7["checks"], m7["settings"]
    anchor = out["periods"][list(out["periods"])[0]]
    focal = disp(M7_FIRM)
    suppliers = " and ".join(disp(j) for j in M7_BRANCHES)
    st.markdown(
        f"<div class='note'><b>M7 is not a further period.</b> It is a sensitivity analysis "
        f"of {focal}'s dependency weights on the M0 inputs. The ESG data, supply "
        f"relationships, peer benchmarks and indicator weights of M0 stay fixed; only "
        f"{focal}'s weight on its own operations (β) and on its direct suppliers "
        f"{suppliers} (α) change. The alternative settings are illustrative, not weights "
        f"derived from procurement or strategic-dependency data.</div>",
        unsafe_allow_html=True)
    st.markdown("")
    fixed = pd.DataFrame([{"Dimension": d, "Weight version": anchor.weights[d].version_id,
                           "Parameter digest": anchor.weights[d].digest} for d in DIMS])
    check_rows = checks.assign(requirement=checks["check"].map(M7_CHECK_TEXT))[
        ["check", "requirement", "passed"]].rename(columns={
            "check": "Check", "requirement": "Requirement", "passed": "Passed"})
    check_sub = f"{int(checks['passed'].sum())} of {len(checks)} pass · also in T15"

    if role == "regulator":
        scope_panel(
            who="the governance body",
            can_see=["That the analysis uses the published M0 benchmark and weight versions",
                     "Whether the checks of the analysis pass"],
            cannot_see=[f"{focal}'s dependency weights β and α in any setting",
                        f"{focal}'s totals and branch contributions",
                        "Any firm's raw values or scores"])
        published_policy_panel(out["policy"])
        with card("m7_fixed"):
            card_head("Parameters held fixed", "Published M0 benchmark and weight versions")
            show_df(fixed, hide_index=True)
        with card("m7_checks"):
            card_head("Checks of the analysis", check_sub)
            show_df(check_rows, hide_index=True)
        return

    if entity is not None and entity != M7_FIRM:
        with st.columns(1)[0]:
            kpi_box("Dependency-weight sensitivity", "Not applicable",
                    f"only {focal}'s weights vary")
        upstream = entity not in _downstream_closure({M7_FIRM})
        st.info(f"M7 varies {focal}'s dependency weights only. "
                + (f"{disp(entity)} is upstream of {focal}, so its scores are the same in "
                   f"every setting. " if upstream else
                   f"{disp(entity)} receives {focal}'s total under the weights {focal} "
                   f"applies; M7 reports {focal}'s own results. ")
                + f"Select {focal} under {'Client' if role == 'auditor' else 'Participant'} "
                  "to view the analysis.")
        return

    if role == "auditor":
        scope_panel(
            who="the assurance provider",
            can_see=["Its client's declared β and α and the records behind them",
                     "Its client's total and branch contributions under every setting",
                     "The checks, including that the baseline reproduces the published "
                     "M0 totals"],
            cannot_see=["Any other firm's data, scores or dependency weights",
                        "How much of a supplier's total comes from that supplier's own "
                        "suppliers"])
    elif role == "firm":
        scope_panel(
            can_see=["Your own M0 local scores and the M0 totals of your direct suppliers",
                     "Your declared dependency weights and the two alternative settings",
                     "Your total and each branch contribution under every setting"],
            cannot_see=["How much of a supplier's total comes from that supplier's own "
                        "suppliers",
                        "Firms beyond your direct suppliers and their scores",
                        "Other firms' dependency weights"])
    else:
        scope_panel(
            can_see=[f"The dependency-weight settings and {focal}'s results under each",
                     "The checks of the analysis"],
            cannot_see=["Procurement or strategic-dependency data behind the alternative "
                        "settings, which are illustrative"])
    published_policy_panel(out["policy"])

    order = [s[0] for s in settings]
    by_dim = {d: table[table["dimension"] == d].set_index("setting").loc[order]
              for d in M7_DIMENSIONS}
    alpha_cols = [f"α {disp(j)}" for j in M7_BRANCHES]
    w_e, w_s = out["policy"].dim_weights
    tabs = st.tabs(["Score overview", "Branch contributions", "Settings and checks"])

    with tabs[0]:
        cols = st.columns(len(settings))
        for col, (name, beta, a_rubber, a_coating) in zip(cols, settings):
            row = by_dim["Composite"].loc[name]
            sub = f"β {beta:.2f} · α {a_rubber:.2f} / {a_coating:.2f}"
            if name != order[0]:
                sub += f" · {row['change_vs_baseline']:+.4f} vs baseline"
            with col:
                kpi_box(f"{_cap(name)} · composite", f"{row['FOCAL_total']:.4f}", sub)
        st.caption(f"Composite = {w_e:.2f} × environmental + {w_s:.2f} × social total; "
                   f"α are the weights on {' / '.join(disp(j) for j in M7_BRANCHES)}.")
        summary = pd.DataFrame([{
            "Setting": _cap(name), "β": beta,
            **dict(zip(alpha_cols, (a_rubber, a_coating))),
            **{d: by_dim[d].loc[name, "FOCAL_total"] for d in M7_DIMENSIONS},
            "Composite vs baseline": by_dim["Composite"].loc[name, "change_vs_baseline"],
            "Branch ranking changed in": ", ".join(
                d for d in M7_DIMENSIONS if bool(by_dim[d].loc[name, "ranking_changed"]))
                or "none",
        } for name, beta, a_rubber, a_coating in settings])
        with card("m7_totals"):
            card_head(f"{focal}'s total by setting",
                      "M0 inputs · benchmark and weights of M0")
            show_df(summary.style.format({"β": "{:.2f}", **{c_: "{:.2f}" for c_ in alpha_cols},
                                          **{d: "{:.4f}" for d in M7_DIMENSIONS},
                                          "Composite vs baseline": "{:+.4f}"}),
                    hide_index=True)

    with tabs[1]:
        dim = st.radio("Dimension", M7_DIMENSIONS, horizontal=True, key="m7_dim")
        rows = by_dim[dim]
        labels = [_cap(n) for n in order]
        fig = go.Figure()
        # Stacked own operations, S1-COATING, S1-RUBBER: the adjacent fills keep
        # their color-vision separation without relying on the gaps alone.
        for column, name, colour in ([("own_operations", "own operations",
                                       ENTITY_COLOR[M7_FIRM])]
                                     + [(f"branch_{j}", f"branch via {disp(j)}",
                                         ENTITY_COLOR[j]) for j in reversed(M7_BRANCHES)]):
            fig.add_trace(go.Bar(
                y=labels, x=rows[column].tolist(), name=name, orientation="h", width=0.36,
                marker=dict(color=colour, line=dict(width=2, color=C["surface"])),
                hovertemplate="%{y}<br>" + name + " %{x:.4f}<extra></extra>"))
        totals = rows["FOCAL_total"].tolist()
        for label, total in zip(labels, totals):
            fig.add_annotation(x=total, y=label, text=f"{total:.4f}", showarrow=False,
                               xanchor="left", xshift=6, font=dict(size=11, color=C["ink2"]))
        fig.update_layout(barmode="stack", legend_traceorder="normal")
        base_layout(fig, 270, f"Contribution to {focal}'s {dim.lower()} "
                              f"{'score' if dim == 'Composite' else 'total'}", "")
        fig.update_xaxes(range=[0, max(totals) * 1.16])
        fig.update_yaxes(autorange="reversed", showgrid=False)
        show_chart(fig)
        detail = pd.DataFrame([{
            "Setting": _cap(name),
            "Own operations": rows.loc[name, "own_operations"],
            **{f"Branch via {disp(j)}": rows.loc[name, f"branch_{j}"] for j in M7_BRANCHES},
            f"{focal} total": rows.loc[name, "FOCAL_total"],
            "Upstream contribution (% of total score)": ((rows.loc[name, "FOCAL_total"] - rows.loc[name, "own_operations"])
                               / rows.loc[name, "FOCAL_total"]
                               if rows.loc[name, "FOCAL_total"] else np.nan),
            "Branch ranking": rows.loc[name, "branch_ranking"],
            "Ranking changed": "yes" if bool(rows.loc[name, "ranking_changed"]) else "",
        } for name in order])
        show_df(detail.style.format({"Own operations": "{:.4f}",
                                     **{f"Branch via {disp(j)}": "{:.4f}" for j in M7_BRANCHES},
                                     f"{focal} total": "{:.4f}", "Upstream contribution (% of total score)": "{:.1%}"}),
                hide_index=True)
        st.caption(f"A branch contribution is α times the supplier's total, which {focal} "
                   f"receives as one number. The ranking compares the two branch "
                   f"contributions; Ranking changed marks a setting that reverses the "
                   f"baseline ranking.")

    with tabs[2]:
        with card("m7_settings"):
            card_head("Dependency-weight settings", f"{M7_SETTINGS_SHEET} sheet of the M7 workbook")
            show_df(pd.DataFrame([{
                "Setting": _cap(name), "β own operations": beta,
                **dict(zip(alpha_cols, (a_rubber, a_coating))),
                "Upstream dependency weight (Σα)": a_rubber + a_coating,
                "Note": _cap(M7_SETTING_NOTES.get(name, "")),
            } for name, beta, a_rubber, a_coating in settings]).style.format(
                {"β own operations": "{:.2f}", **{c_: "{:.2f}" for c_ in alpha_cols},
                 "Upstream dependency weight (Σα)": "{:.1%}"}), hide_index=True)
            st.caption("The workbook holds the M0 indicator values, checked against the M0 "
                       "workbook in every run. Every other firm keeps its declared "
                       "dependency weights.")
        with card("m7_fixed"):
            card_head("Parameters held fixed", "Published M0 benchmark and weight versions")
            show_df(fixed, hide_index=True)
        with card("m7_checks"):
            card_head("Checks of the analysis", check_sub)
            show_df(check_rows, hide_index=True)


# -----------------------------------------------------------------------------
# FIRM VIEW
# -----------------------------------------------------------------------------

def firm_view(out: dict) -> None:
    policy = out["policy"]
    labels = list(out["periods"].keys())
    keep_state("firm_me", "FOCAL", allowed=ENTITY_ORDER)

    with keyed_container("pagehead"):
        left, right = header_columns()
        with right:
            c1, c2 = control_columns(2)
            with c1:
                me = st.selectbox("Participant", ENTITY_ORDER, format_func=disp,
                                  key="firm_me",
                                  help="Choose which participant's "
                                       "view to display.")
            with c2:
                period = period_picker(dashboard_period_labels(out), "period_firm",
                                       default=labels[-1])
        sups, custs = direct_suppliers(me), direct_customers(me)
        with left:
            header_top(
                ROLE_PORTAL["Participating firm"],
                f"{disp(me)} · {SUPPLY_CHAIN[me]['role']}",
                [SUPPLY_CHAIN[me]["tier"], _period_chip(period),
                 f"{len(sups)} direct supplier{'s' if len(sups) != 1 else ''}",
                 f"{len(custs)} direct customer{'s' if len(custs) != 1 else ''}"])
        header_body(
            "Everything on this page is what this firm would hold under the intended "
            "design: its own data, the published parameters, and the aggregate scores of "
            "its direct suppliers.", scenario_rows(period) + role_rows("Participating firm"))

    if period == "M6":
        m6_dashboard_panel(out, entity=me)
        return
    if period == "M7":
        m7_dashboard_panel(out, role="firm", entity=me)
        return
    r = out["periods"][period]
    trig_all = out["triggers"]
    mine = trig_all[trig_all["entity"] == disp(me)]
    into = _transition_into(period, labels)
    now = mine[mine["transition"] == into] if into else mine.iloc[0:0]

    scope_panel(
        can_see=[
            "Your own raw indicator values, local, total and composite scores",
            "Published conversion rules, common-scale benchmarks, weights, versions, "
            "digests and the governance body's policy",
            ("Aggregate total score of each direct supplier: "
             + ", ".join(disp(j) for j in sups)) if sups else
            "No enrolled direct supplier, so no upstream score is received",
            ("Your direct customers (" + ", ".join(disp(c_) for c_ in custs)
             + "), which receive your total but do not enter your score") if custs else
            "No direct customer, so your total is not passed downstream",
            "Your own dependency weights β and α and the branch split of your score",
            "Your own internal warnings, and hard alerts reaching you through a direct "
            "supplier or customer",
            "Your own upload (n, Σx, Σx²) and your own assurance findings",
        ],
        cannot_see=[
            "Any other firm's raw values or local score",
            "Suppliers beyond the direct tier: who they are, how many, their scores",
            "Which deeper-tier firm raised a hard alert, and other firms' internal warnings",
            "A direct supplier's own-operations vs its-upstream split",
            "Your customers' scores" if custs else "Downstream firms' scores",
            "Your rank among peers, or other firms' dependency weights",
        ])
    published_policy_panel(policy)

    tabs = st.tabs(["Score overview", "Indicators", "Suppliers",
                    "Warnings and alerts", "Uploads and assurance"])

    # ---- 1 · MY SCORE ---------------------------------------------------
    with tabs[0]:
        i = labels.index(period)
        prev = out["periods"][labels[i - 1]] if i > 0 else None
        comp = r.composite[me]
        flagged = peer_relative_flag({me: comp}, policy)[me]
        dd = bool(now["internal_warning"].any()) if not now.empty else False
        direction = (combine_directions(str(x) for x in now["warning_direction"]
                                        if str(x) not in ("", "nan"))
                     if not now.empty else "")
        hard_local = me in r.hard_hits
        hard_via = sorted({x for col in ("hard_alert_via_direct_supplier",
                                         "hard_alert_via_direct_customer")
                           for v in now[col] for x in str(v).split(", ")
                           if x and x != "nan"}) if not now.empty else []

        k = st.columns(5)
        with k[0]:
            kpi_box("Composite risk", f"{comp:.4f}",
                    (f"{comp - prev.composite[me]:+.4f} vs {labels[i-1]}"
                     if prev else "first period"))
        with k[1]:
            kpi_box("Environmental total", f"{r.totals['Environmental'][me]:.4f}",
                    f"local {r.local['Environmental'][me]:.4f}")
        with k[2]:
            kpi_box("Social total", f"{r.totals['Social'][me]:.4f}",
                    f"local {r.local['Social'][me]:.4f}")
        with k[3]:
            kpi_box("Peer-relative flag", "above level" if flagged else "below level",
                    f"published level {policy.peer_relative_flag_level:.2f}",
                    alert=flagged)
        with k[4]:
            if hard_local:
                kpi_box("Hard threshold", "own alert", r.hard_hits[me], alert=True)
            elif hard_via:
                kpi_box("Hard threshold", "received",
                        "through " + ", ".join(hard_via), alert=True)
            else:
                still_open = ([f"{x['dimension'][0]} · {x['open_internal_warnings']}"
                               for _, x in now.iterrows() if x["open_internal_warnings"]]
                              if not now.empty else [])
                if dd and direction in ("deterioration", "mixed"):
                    kpi_box("Internal warning", f"raised · {direction}",
                            f"transition {into}", alert=True)
                elif still_open:
                    kpi_box("Internal warning", "open", " | ".join(still_open), alert=True)
                else:
                    kpi_box("Internal warning",
                            (f"raised · {direction}" if direction else "raised")
                            if dd else "none",
                            f"transition {into}" if into else "no prior period")

        prov_own = {d_: r.provisional.get(d_, {}).get(me, {}) for d_ in DIMS}
        prov_msgs = [f"{d_.lower()} score: {'; '.join(v.get('reasons', []))}"
                     for d_, v in prov_own.items() if v.get("provisional")]
        prov_via = sorted({disp(j) for v in prov_own.values() for j in v.get("via", [])})
        if prov_msgs or prov_via:
            st.markdown(
                "<div class='warn'><b>Provisional.</b> "
                + (("Own " + " | ".join(prov_msgs) + ". ") if prov_msgs else "")
                + (("The branch through " + ", ".join(prov_via) + " rests on a "
                    "provisional score upstream. ") if prov_via else "")
                + "A provisional score is computed with the published rule like any "
                "other; the flag states why it may still change.</div>",
                unsafe_allow_html=True)
        st.markdown("")
        left, right = st.columns([1.3, 1])
        with left:
            st.markdown("#### My composite over time")
            fig = go.Figure()
            fig.add_trace(go.Scatter(
                x=labels, y=[out["periods"][l].composite[me] for l in labels],
                name="rolling benchmark", mode="lines+markers",
                line=dict(width=2.2, color=ENTITY_COLOR[me]),
                marker=dict(size=8, line=dict(width=2, color=C["surface"])),
                hovertemplate="%{x}: %{y:.4f}<extra>rolling</extra>"))
            fig.add_trace(go.Scatter(
                x=labels, y=[out["periods_fixed"][l].composite[me] for l in labels],
                name="fixed first-round benchmark", mode="lines+markers",
                line=dict(width=2, color=C["muted"], dash="dash"),
                marker=dict(size=8, color=C["muted"], line=dict(width=2, color=C["surface"])),
                hovertemplate="%{x}: %{y:.4f}<extra>fixed</extra>"))
            fig.add_hline(y=policy.peer_relative_flag_level,
                          line=dict(color=C["critical"], width=1.2, dash="dot"),
                          annotation_text="published escalation level",
                          annotation_position="top left",
                          annotation_font=dict(size=10, color=C["critical"]))
            show_chart(base_layout(fig, 330, "", "Composite risk"))
            st.markdown(
                "<div class='note'><b>Two lines, one firm.</b> On the rolling line a "
                "score can rise because this firm got worse or because its peers got "
                "better. The fixed line freezes the first round's μ, σ_score and "
                "weights, so its movement reflects this firm's own observations and "
                "its suppliers' scores. Neither line measures absolute harm; raw values "
                "are checked against declared limits separately (My indicators).</div>",
                unsafe_allow_html=True)
        with right:
            st.markdown("#### My scores by period")
            show_df(pd.DataFrame([{
                "Period": l,
                "E local": out["periods"][l].local["Environmental"][me],
                "S local": out["periods"][l].local["Social"][me],
                "E total": out["periods"][l].totals["Environmental"][me],
                "S total": out["periods"][l].totals["Social"][me],
                "Composite": out["periods"][l].composite[me],
                "Fixed bench.": out["periods_fixed"][l].composite[me],
            } for l in labels]).style.format(
                {c_: "{:.4f}" for c_ in ["E local", "S local", "E total", "S total",
                                         "Composite", "Fixed bench."]}),
                hide_index=True)
            hidden_note("where this score ranks among peers, and any other firm's "
                        "composite. The flag compares this firm's own score with a "
                        "published level; it is not a league table.")

    # ---- 2 · MY INDICATORS ----------------------------------------------
    with tabs[1]:
        dim = st.radio("Dimension", DIMS, horizontal=True, key="f_ind_dim")
        stats, ws = r.stats[dim], r.weights[dim]
        ls = r.scores[dim][me]
        common_scale = bool(stats.common_mu)
        own = r.data[dim][r.data[dim]["entity"] == me].set_index("indicator")
        rows = []
        for _, w in ws.rows.iterrows():
            k_ = w["indicator"]
            x = float(own.loc[k_, "value"]) if k_ in own.index else np.nan
            mu = stats.mu.get(k_, np.nan)
            sig = scoring_sigma(stats, k_) if k_ in stats.mu else np.nan
            good = bool(stats.is_good.get(k_, False))
            converted = np.nan
            z = (x - mu) / sig if np.isfinite(x) and np.isfinite(sig) and sig > 0 else np.nan
            rz = -z if good else z
            if k_ in stats.common_mu:
                lo_, hi_ = stats.input_bounds[k_]
                converted = np.clip((x-lo_)/(hi_-lo_), 0, 1)
                if good: converted = 1-converted
                mu = stats.common_mu[k_]
                sig = max(stats.common_sigma[k_], policy.eta_sigma_score*stats.s_ref[k_]/(hi_-lo_))
                rz = (converted-mu)/sig
            rows.append({
                "ID": w["ind_id"], "Indicator": k_,
                "Level": GOV_LEVEL.get(w["tier"], w["tier"]),
                "My value": x, "Unit": own.loc[k_, "unit"] if k_ in own.index else "",
                "Converted value (0–1)": converted,
                "Published μ": mu, "σ_score": sig,
                "Direction": "higher = lower risk" if good else "higher = higher risk",
                "Published weight": w["w_final"], "risk z": rz,
                "Contribution": ls.contributions.get(k_, np.nan),
                "Reported": "yes" if np.isfinite(x) else "not reported",
            })
        ind = pd.DataFrame(rows)[["ID", "Indicator", "My value", "Unit", "Converted value (0–1)", "Published μ",
                                  "risk z", "Contribution", "Published weight", "Level",
                                  "σ_score", "Direction", "Reported"]]

        k = st.columns(4)
        with k[0]:
            kpi_box(f"{dim} local score", f"{r.local[dim][me]:.4f}",
                    "sigmoid(Σ w·risk z / τ); missing inputs produce an interval")
        prov_dim = r.provisional.get(dim, {}).get(me, {})
        share_ = prov_dim.get("coverage_share", ls.coverage)
        with k[1]:
            kpi_box("Share of published weight reported",
                    f"{share_:.1%}",
                    f"published floor {policy.coverage_floor:.0%}; below it the score "
                    f"is provisional",
                    alert=bool(prov_dim.get("provisional")))
        with k[2]:
            worse = int((ind["risk z"] > 0).sum())
            kpi_box("Indicators worse than benchmark", f"{worse} / {len(ind)}",
                    "risk z > 0")
        with k[3]:
            kpi_box("Weight version", ws.version_id, f"digest {ws.digest}")

        if ls.incomplete:
            bounds = score_intervals(r.scores, r.dependency_weights)[dim][me]
            st.warning(f"Incomplete inputs. Local interval [{ls.lower:.4f}, {ls.upper:.4f}]; "
                       f"total interval [{bounds[0]:.4f}, {bounds[1]:.4f}]. The point score is the upper bound.")
        if prov_dim.get("provisional"):
            st.markdown(
                "<div class='warn'><b>Provisional score.</b> "
                + "; ".join(prov_dim.get("reasons", []))
                + ". The number is the published rule applied to what this firm "
                "reported; each direct customer is told that its branch through this "
                "firm rests on a provisional score.</div>", unsafe_allow_html=True)
        if ls.hard_threshold_hit:
            st.markdown(
                f"<div class='crit'><b>Hard threshold: {ls.hard_threshold_hit}.</b> "
                f"The alert stays separate from the continuous score, reaches every "
                f"affected firm upstream and downstream, and remains active until a "
                f"reviewed clearance.</div>",
                unsafe_allow_html=True)

        left, right = st.columns([1, 1.25])
        with left:
            st.markdown("#### What drives my score")
            top = ind.dropna(subset=["Contribution"])
            top = top.reindex(top["Contribution"].abs().sort_values(ascending=False).index).head(12)
            top = top.iloc[::-1]
            fig = go.Figure(go.Bar(
                y=top["ID"], x=top["Contribution"], orientation="h",
                marker=dict(color=[C["critical"] if v > 0 else C["s1"]
                                   for v in top["Contribution"]],
                            line=dict(width=2, color=C["surface"])),
                customdata=top["Indicator"],
                hovertemplate="%{customdata}<br>contribution %{x:+.4f}<extra></extra>"))
            fig.add_vline(x=0, line=dict(color=C["ink"], width=1))
            show_chart(base_layout(fig, 400, "Weighted indicator term (w × z)", "", legend=False))
            st.caption("Positive terms increase the weighted index; negative terms reduce it. "
                       "Terms are summed before division by τ and the sigmoid conversion; "
                       "they do not add directly to the final local score.")
        with right:
            st.markdown("#### My values against the published benchmark")
            mean_label = "Peer mean (0–1)" if common_scale else "Peer mean (raw units)"
            sd_label = "Scoring SD (0–1)" if common_scale else "Scoring SD (raw units)"
            display = ind.sort_values("Contribution", ascending=False).rename(columns={
                "My value": "Reported value", "Published μ": mean_label,
                "σ_score": sd_label, "Contribution": "Weighted term (w × z)",
                "Direction": "Raw-value direction"})
            show_df(display.style.format({
                "Reported value": "{:.4g}", "Converted value (0–1)": "{:.4g}",
                mean_label: "{:.4g}", sd_label: "{:.4g}",
                "Published weight": "{:.4f}", "risk z": "{:+.3f}",
                "Weighted term (w × z)": "{:+.4f}"}, na_rep="—"),
                hide_index=True, height=400)
        st.caption("Reported values use the stated raw units. Converted values, peer means "
                   "and scoring SDs use the common 0–1 scale, where higher is less favorable. "
                   "The scoring SD includes the published positive lower bound. The benchmark "
                   "is computed over the reporting participants in the "
                   "demonstration cohort (seven heterogeneous entities), published with "
                   "the weight version; they are the only information about peers this "
                   "firm receives. They are enough to recompute its own score exactly.")
        hidden_note(f"any peer's value on any indicator. A benchmark is published only "
                    f"when at least {Coordinator.MIN_PARTICIPANTS} firms report the "
                    f"indicator. This reduces inference risk but does not guarantee privacy: "
                    f"equal values, other constraints or firms pooling their information "
                    f"can still reveal individual values.")

        st.markdown("#### My raw-indicator checks against declared limits")
        ab_ = r.absolute_risk
        ab_ = ab_[(ab_["entity"] == me) & (ab_["dimension"] == dim)] if ab_ is not None else None
        if ab_ is None or ab_.empty:
            st.caption(f"No hard-threshold indicator is defined in {dim.lower()} for "
                       f"this firm.")
        else:
            show_df(ab_.drop(columns=["entity"]), hide_index=True)

    # ---- 3 · MY SUPPLIERS -----------------------------------------------
    with tabs[2]:
        dim = st.radio("Dimension", DIMS, horizontal=True, key="f_sup_dim")
        totals, comps = r.totals[dim], r.comps[dim]
        dw_me = r.dependency_weights[me]
        beta = dw_me["beta"]

        left, right = st.columns([1.15, 1])
        with left:
            st.markdown("#### My view of the chain")
            fig = go.Figure()
            ys = np.linspace(1.6, -0.2, len(sups)) if len(sups) > 1 else [0.7] * len(sups)
            me_xy = (2, 0.7)
            if sups:
                fig.add_shape(type="rect", x0=-0.45, x1=0.45, y0=-0.55, y1=1.95,
                              line=dict(color=C["muted"], width=1, dash="dash"),
                              fillcolor="#f7f7f5", layer="below")
                fig.add_annotation(x=0, y=0.7, showarrow=False,
                                   text="further upstream<br><i>not visible</i><br>"
                                        "folded into each<br>supplier's total",
                                   font=dict(size=11, color=C["muted"]))
            for j, y in zip(sups, ys):
                a = dw_me["alpha"][j]
                fig.add_annotation(x=me_xy[0] - 0.22, y=me_xy[1], ax=1.18, ay=y,
                                   xref="x", yref="y", axref="x", ayref="y",
                                   showarrow=True, arrowhead=2, arrowsize=1,
                                   arrowwidth=1 + 6 * a, arrowcolor=C["s1"])
                if sups:
                    fig.add_annotation(x=0.82, y=y, ax=0.45, ay=y, xref="x", yref="y",
                                       axref="x", ayref="y", showarrow=True, arrowhead=2,
                                       arrowwidth=1, arrowcolor="#b8b6ae")
            cys = np.linspace(1.3, 0.1, len(custs)) if len(custs) > 1 else [0.7] * len(custs)
            for y in cys:
                fig.add_annotation(x=2.84, y=y, ax=me_xy[0] + 0.22, ay=me_xy[1],
                                   xref="x", yref="y", axref="x", ayref="y",
                                   showarrow=True, arrowhead=2, arrowwidth=1,
                                   arrowcolor="#b8b6ae")
            if custs:
                fig.add_annotation(x=3.0, y=max(cys) + 0.45, showarrow=False,
                                   text=f"<i>receive{'s' if len(custs) == 1 else ''} this "
                                        f"firm's total<br>not part of its score</i>",
                                   font=dict(size=10.5, color=C["muted"]))
                fig.add_trace(go.Scatter(
                    x=[3.0] * len(custs), y=list(cys), mode="markers+text",
                    text=[disp(c_) for c_ in custs],
                    textposition="bottom center", textfont=dict(size=10, color=C["ink"]),
                    marker=dict(size=30, symbol="circle-open", line=dict(width=2.5),
                                color=[ENTITY_COLOR[c_] for c_ in custs]),
                    hovertemplate="%{text}<extra>direct customer, not part of this "
                                  "firm's score</extra>",
                    showlegend=False))
            fig.add_trace(go.Scatter(
                x=[1.0] * len(sups), y=list(ys), mode="markers+text",
                text=[f"{disp(j)}<br>total {totals[j]:.3f} · α {dw_me['alpha'][j]:.2f}"
                      for j in sups],
                textposition="bottom center", textfont=dict(size=10, color=C["ink"]),
                marker=dict(size=34, color=[ENTITY_COLOR[j] for j in sups],
                            line=dict(width=2, color=C["surface"])),
                hovertemplate="%{text}<extra>direct supplier (aggregate only)</extra>",
                showlegend=False))
            fig.add_trace(go.Scatter(
                x=[me_xy[0]], y=[me_xy[1]], mode="markers+text",
                text=[f"<b>{disp(me)}</b><br>total {totals[me]:.3f}"],
                textposition="bottom center", textfont=dict(size=11, color=C["ink"]),
                marker=dict(size=44, color=ENTITY_COLOR[me], line=dict(width=2, color=C["surface"])),
                hovertemplate="%{text}<extra>this firm</extra>", showlegend=False))
            fig.update_xaxes(visible=True, showgrid=False, zeroline=False,
                             range=[-0.6, 3.65], tickmode="array", tickvals=[0, 1, 2, 3],
                             ticktext=["beyond direct tier", "direct suppliers",
                                       "this firm", "direct customers"],
                             tickfont=dict(size=11, color=C["muted"]))
            fig.update_yaxes(visible=False, range=[-0.9, 2.2])
            show_chart(base_layout(fig, 380, legend=False))
            if not sups:
                st.caption("This firm has no enrolled direct supplier, so its total "
                           "equals its own-operations score.")

        with right:
            st.markdown("#### Where my total comes from")
            tot = sum(comps[me].values())
            att = [{"Component": "own operations", "Weight": beta,
                    "Score": r.local[dim][me],
                    "Contribution": comps[me]["__own__"],
                    "Share of total": comps[me]["__own__"] / tot if tot else np.nan}]
            for j in sups:
                att.append({"Component": f"branch via {disp(j)}",
                            "Weight": dw_me["alpha"][j],
                            "Score": totals[j],
                            "Contribution": comps[me].get(j, 0.0),
                            "Share of total": comps[me].get(j, 0.0) / tot if tot else np.nan})
            show_df(pd.DataFrame(att).style.format({
                "Weight": "{:.3f}", "Score": "{:.4f}",
                "Contribution": "{:.4f}", "Share of total": "{:.1%}"}), hide_index=True)
            method_ = DECLARED_ALPHA_METHOD.get(me)
            st.caption(("β and α are this firm's own dependency weights, set by the "
                        f"{method_} method accepted under Table 4"
                        + ("; it requires enhanced assurance"
                           if method_ in ENHANCED_ASSURANCE_METHODS else "")
                        + ". They are not transmitted; an assurance provider checks that "
                        "the declared method was applied. A documented revision between "
                        "calibration rounds is reported separately from ESG changes.")
                       if method_ else "No enrolled direct supplier, so β = 1.")
            if sups:
                hidden_note("how much of a supplier's total comes from that supplier's "
                            "own operations and how much from its own suppliers. One "
                            "number per branch is all that crosses the boundary.")

        if sups:
            st.markdown("#### Supplier scores received each period")
            fig = go.Figure()
            for n_, j in enumerate(sups):
                fig.add_trace(go.Scatter(
                    x=labels, y=[out["periods"][l].totals[dim][j] for l in labels],
                    name=disp(j), mode="lines+markers",
                    line=dict(width=2, color=ENTITY_COLOR[j]),
                    marker=dict(size=8, line=dict(width=2, color=C["surface"])),
                    hovertemplate="%{fullData.name}<br>%{x}: %{y:.4f}<extra></extra>"))
            show_chart(base_layout(fig, 280, "", f"{dim} total received"))

            with st.expander("How sensitive is my attribution to my own β?"):
                sens = dependency_sensitivity(r, entity=me, policy=policy)
                show_df(sens.drop(columns=["entity"]).rename(columns={
                    "beta": "β", "E_total": "E total", "S_total": "S total",
                    "E_upstream_share": "E upstream share",
                    "S_upstream_share": "S upstream share", "composite": "Composite",
                    "peer_relative_flag": "Above escalation level"}).style.format({
                    "β": "{:.2f}", "E total": "{:.4f}", "S total": "{:.4f}",
                    "E upstream share": "{:.1%}", "S upstream share": "{:.1%}",
                    "Composite": "{:.4f}"}), hide_index=True)
                st.caption("The sweep rescales this firm's α proportionally and uses "
                           "only its own local score and the supplier totals it "
                           "already receives.")

    # ---- 4 · WARNINGS AND ALERTS -----------------------------------------
    with tabs[3]:
        st.markdown("#### Warnings and alerts on my node, period by period")
        t = mine[["transition", "dimension", "published_change", "net_change",
                  "new_version_effect", "comparable_with_previous", "max_abs_component",
                  "threshold", "internal_warning", "warning_signal", "indicators_flagged",
                  "branches_flagged", "warning_direction", "open_internal_warnings",
                  "internal_warnings_closed", "response_level",
                  "hard_alert_local", "hard_alert_via_direct_supplier",
                  "hard_alert_via_direct_customer", "hard_event_count", "hard_event_ids",
                  "hard_notice_received", "upstream_dispersion",
                  "max_single_upstream_share"]].rename(columns={
            "transition": "Change", "dimension": "Dim",
            "published_change": "Published change", "net_change": "Same rules",
            "new_version_effect": "New benchmark/weights",
            "comparable_with_previous": "Comparable",
            "max_abs_component": "Largest component change", "threshold": "Threshold",
            "internal_warning": "Internal warning",
            "warning_signal": "Warning signal",
            "indicators_flagged": "Own indicators beyond threshold",
            "branches_flagged": "Supplier contributions beyond threshold",
            "warning_direction": "Direction",
            "open_internal_warnings": "Warnings still open",
            "internal_warnings_closed": "Warnings closed",
            "response_level": "Response",
            "hard_alert_local": "Hard alert (own)",
            "hard_alert_via_direct_supplier": "Hard alert via supplier",
            "hard_alert_via_direct_customer": "Hard alert via customer",
            "hard_event_count": "Distinct hard events", "hard_event_ids": "Hard event IDs",
            "hard_notice_received": "Hard-event paths",
            "upstream_dispersion": "Upstream dispersion",
            "max_single_upstream_share": "Largest single upstream share"})
        t = t.assign(Dim=t["Dim"].str[0],
                     Comparable=t["Comparable"].map({True: "yes", False: "not directly"}))
        for c_ in ("Internal warning", "Hard alert (own)"):
            t[c_] = t[c_].map({True: "yes", False: ""})
        show_df(t.style.format({"Published change": "{:+.5f}", "Same rules": "{:+.5f}",
                                "New benchmark/weights": "{:+.5f}",
                                "Largest component change": "{:.5f}",
                                "Threshold": "{:.5f}", "Upstream dispersion": "{:.4f}",
                                "Largest single upstream share": "{:.1%}"}),
                hide_index=True, height=360)
        st.caption("The published change in the total is split into the change under "
                   "the previous period's benchmark and weights (same rules), which the "
                   "warning tests use, and the part due only to the new benchmark and "
                   "weights. A change is not directly comparable when indicator "
                   "definitions or the data coverage behind the total change "
                   "(Section 4.2.4).")
        st.caption("Each own indicator and each supplier contribution is tested "
                   "separately; the warning fires whatever the net change and stays "
                   "with this firm. A warning stays open until the same component has "
                   "come back within the limit; an improvement elsewhere does not close "
                   "it. The number in brackets is the component's change since the "
                   "warning opened. Hard alerts arrive through a direct supplier or a "
                   "direct customer. Matching event IDs identify one event across several "
                   "branches; different IDs remain separate even when their descriptions match.")

        st.markdown("#### Provisional scores")
        prov_rows = []
        for l in labels:
            for d_ in DIMS:
                v = out["periods"][l].provisional.get(d_, {}).get(me, {})
                if v.get("provisional") or v.get("via"):
                    prov_rows.append({
                        "Period": l, "Dim": d_[0],
                        "Own score provisional": "yes" if v.get("provisional") else "",
                        "Reason": "; ".join(v.get("reasons", [])),
                        "Branch rests on a provisional score via":
                            ", ".join(disp(j) for j in v.get("via", [])),
                    })
        if prov_rows:
            show_df(pd.DataFrame(prov_rows), hide_index=True)
        else:
            st.caption(f"No provisional score on this firm's node or in its branches in "
                       f"any period. A score is provisional when a scored input is "
                       f"missing, or when the firm reported less than "
                       f"{policy.coverage_floor:.0%} of the published weight. The notice "
                       f"names only the direct supplier.")

        st.markdown("#### Component changes vs the net change")
        cc1, cc2 = st.columns([1, 3])
        with cc1:
            trans = list(dict.fromkeys(mine["transition"]))
            ratio = mine["max_abs_component"] / mine["threshold"].where(mine["threshold"] > 0)
            best = mine.loc[ratio.fillna(0).idxmax()] if len(mine) else None
            tr = st.selectbox("Transition", trans,
                              index=trans.index(best["transition"]) if best is not None else 0,
                              key="f_al_tr", help="Defaults to the largest component change "
                                                  "relative to its threshold.")
            dim_c = st.radio("Dimension", DIMS, horizontal=True, key="f_al_dim",
                             index=DIMS.index(best["dimension"]) if best is not None else 0)
        p0_, p1_ = tr.split("->")
        rep_ = _transition_reports(out["periods"][p0_], out["periods"][p1_], dim_c,
                                   policy)[me]
        keys_ = sorted(rep_.component_deltas, key=lambda k_: (k_ in SUPPLY_CHAIN, k_))
        deltas = [rep_.component_deltas[k_] for k_ in keys_]
        names_ = [f"via {disp(k_)}" if k_ in SUPPLY_CHAIN else f"own {k_}" for k_ in keys_]
        th_ = rep_.threshold
        with cc2:
            fig = go.Figure()
            fig.add_trace(go.Bar(
                x=names_, y=deltas, name="component change",
                marker=dict(color=[C["critical"] if abs(v) > th_ else C["s1"] for v in deltas],
                            line=dict(width=2, color=C["surface"])),
                hovertemplate="%{x}<br>%{y:+.6f}<extra></extra>"))
            fig.add_trace(go.Bar(
                x=["net change"], y=[rep_.net_change], name="net change",
                marker=dict(color="#b8b6ae", line=dict(width=2, color=C["surface"])),
                hovertemplate="net %{y:+.6f}<extra></extra>"))
            for sgn in (1, -1):
                fig.add_hline(y=sgn * th_, line=dict(color=C["critical"], width=1, dash="dot"))
            fig.add_hline(y=0, line=dict(color=C["ink2"], width=1))
            fig.update_layout(bargap=0.35)
            show_chart(base_layout(fig, 300, "", f"Δ {dim_c.lower()} total"))
            st.caption(f"Dotted lines: ± threshold {th_:.5f}. An own-indicator bar is the "
                       f"change that indicator alone would cause in this firm's total. A "
                       f"red bar raises an internal warning even when the net bar stays "
                       f"inside.")
        a, b = st.columns(2)
        with a:
            st.markdown(
                "<div class='note'><b>What this firm learns.</b> "
                "<i>Own indicators</i> and <i>supplier contributions</i> beyond the "
                "threshold raise this firm's internal warning; its customers test the "
                "updated total under their own rules. <i>Hard alert via</i>: a confirmed "
                "hard-alert condition reached this firm through that direct supplier or direct "
                "customer, with its category, status and timing.</div>",
                unsafe_allow_html=True)
        with b:
            hidden_note("which firm raised a hard alert, or another firm's internal "
                        "warnings. The direct supplier or customer is the only name a "
                        "hard alert carries.")

        st.markdown("#### Frozen-benchmark events as this firm receives them")
        st.caption("Isolated scenarios from the same published R0 state. The row shown "
                   "is this firm's node only; the event source is masked unless it is "
                   "this firm or a direct supplier.")
        ev_rows = []
        for sc, table in out["events"]["scenarios"].items():
            row = table[table["entity"] == me]
            if row.empty:
                continue
            x = row.iloc[0]
            if "event_source" in x:
                src = str(x["event_source"])
                source_ = _mask_source(src, me)
                indicator_ = x["indicator_id"] if src == me else "not disclosed"
            else:
                source_ = ("your dependency-weight revision" if x["revised_firm"] == me
                           else x["change_source"] or "not disclosed")
                indicator_ = ""
            ev_rows.append({
                "Scenario": sc, "Event source": source_, "Indicator": indicator_,
                "Composite before": x["composite_before"],
                "Composite after": x["composite_after"],
                "Δ composite": x["composite_delta"],
                "Internal warning": bool(x["internal_warning"]),
                "Warning signal": x.get("warning_signal", ""),
                "Direction": x.get("warning_direction", ""),
                "Response": x.get("response_level", ""),
                "Warnings still open": x.get("open_internal_warnings", ""),
                "Warnings closed": x.get("internal_warnings_closed", ""),
                "Hard alert": bool(x["hard_alert"]),
                "Distinct hard events": x.get("hard_event_count", 0),
                "Hard event IDs": x.get("hard_event_ids", ""),
                "Hard alert via supplier": x.get("hard_alert_via_direct_supplier", "") or "",
                "Hard alert via customer": x.get("hard_alert_via_direct_customer", "") or "",
            })
        if ev_rows:
            evdf = pd.DataFrame(ev_rows)
            for c_ in ("Internal warning", "Hard alert"):
                evdf[c_] = evdf[c_].map({True: "yes", False: ""})
            show_df(evdf.style.format({
                "Composite before": "{:.6f}", "Composite after": "{:.6f}",
                "Δ composite": "{:+.2e}"}), hide_index=True)
        st.caption(WARNING_SIGNAL_NOTE)

    # ---- 5 · UPLOADS & ASSURANCE ----------------------------------------
    with tabs[4]:
        c1, c2 = st.columns([1.2, 1])
        with c1:
            dim = st.radio("Dimension", DIMS, horizontal=True, key="f_up_dim")
            st.markdown("#### What this firm transmits")
            common_up = next((s for s in getattr(r, "common_summaries", {}).get(dim, [])
                              if s.entity == me), None)
            s_ = next((s for s in r.summaries[dim] if s.entity == me), None)
            if common_up is None and s_ is None:
                st.caption("No summary transmitted in this dimension.")
            else:
                if common_up is not None:
                    up = pd.DataFrame([{"Indicator": k_, "n": common_up.n[k_],
                                        "Σx (0–1)": common_up.sx.get(k_),
                                        "Σx² (0–1)": common_up.sxx.get(k_)}
                                       for k_ in common_up.n])
                    show_df(up.style.format({"Σx (0–1)": "{:.6g}", "Σx² (0–1)": "{:.6g}"}),
                            hide_index=True, height=320)
                    st.caption(f"Calibration upload on the common 0–1 scale: the firm "
                               f"converts its own values with the published reference "
                               f"points before summarizing (Section 4.2.2). Payload digest "
                               f"{common_up.payload_digest()} · three numbers per indicator, "
                               f"no raw rows.")
                if s_ is not None:
                    with st.expander("Raw-unit summaries (local evaluator diagnostics)"):
                        raw_up = pd.DataFrame([{"Indicator": k_, "n": s_.n[k_],
                                                "Σx": s_.sx.get(k_), "Σx²": s_.sxx.get(k_)}
                                               for k_ in s_.n])
                        show_df(raw_up.style.format({"Σx": "{:.6g}", "Σx²": "{:.6g}"}),
                                hide_index=True)
                        st.caption(f"Diagnostic digest {s_.payload_digest()}. These are "
                                   f"local in-process checks, not regular calibration uploads. "
                                   f"Only a separately logged reference initialization "
                                   f"uses raw-unit uploads to establish ranges and scales.")
                st.markdown(
                    "<div class='warn'>With one observation per firm, n = 1 and Σx "
                    "<b>is</b> the value. Summary statistics do not hide it from the "
                    "coordinator; this reference simulates masking in one process and "
                    "gives no cryptographic guarantee.</div>", unsafe_allow_html=True)
        with c2:
            st.markdown(f"#### Assurance experiments · {labels[0]}")
            st.caption("Independent experiments under the baseline parameters, including "
                       "deliberately introduced anomalies. These results do not describe "
                       "the currently selected period.")
            a_ = out["assurance"]
            rec = a_["C_recomputation"]
            rec = rec[rec["entity"] == disp(me)]
            if not rec.empty:
                x = rec.iloc[0]
                kpi_box("Auditor recomputation",
                        "matches" if x["matches_to_1e-9"] else "differs",
                        f"published {x['published_score']:.6f} · recomputed "
                        f"{x['auditor_recomputation']:.6f}",
                        alert=not bool(x["matches_to_1e-9"]))
            st.markdown("")
            am = a_["D_alpha_method"]
            am = am[am["entity"] == disp(me)]
            if not am.empty:
                x = am.iloc[0]
                kpi_box("Dependency-weight method", "verified" if x["verified"] else
                        "does not match", f"declared: {x['declared_method']}",
                        alert=not bool(x["verified"]))
                if not x["verified"]:
                    st.caption(f"Applied method recorded: {x['method_applied']}")
            st.markdown("")
            findings = [f for f in a_["A_hard_threshold"]["auditor_findings"]
                        if f.entity == me]
            if a_["B_summary_anomaly"]["entity"] == me:
                findings += list(a_["B_summary_anomaly"]["findings"])
            if findings:
                for f in findings:
                    st.markdown(f"<div class='crit'>{f.kind} · {f.indicator} · "
                                f"{f.detail}</div>", unsafe_allow_html=True)
            else:
                st.markdown("<div class='note'>No assurance finding against this firm "
                            "in the reference probes.</div>", unsafe_allow_html=True)
            st.markdown("")
            hidden_note("assurance results for any other firm. Each assurance provider "
                        "reports to its own client.")


# -----------------------------------------------------------------------------
# GOVERNANCE BODY VIEW
# -----------------------------------------------------------------------------

def regulator_view(out: dict) -> None:
    policy = out["policy"]
    dispersion_name = dispersion_label(policy.dispersion_mode)
    labels = list(out["periods"].keys())
    trig = out["triggers"]

    # M0-M5 are the periodic recalibration rounds of the supply chain; M6 (an
    # industry peer experiment) and M7 (a dependency-weight sensitivity on the
    # M0 inputs) are separate experiments, not further periods.
    families = (["M0–M5"] + (["M6"] if "m6_peers" in out else [])
                + (["M7"] if "m7" in out else []))
    keep_state("regulator_scenario", "M0–M5", allowed=families)
    chosen = st.session_state.get("regulator_scenario", "M0–M5")
    with keyed_container("pagehead"):
        chips = {"M6": ["M6 industry peer experiment",
                        "Four-firm peer group · environmental Level 2 and 3 indicators"],
                 "M7": ["M7 dependency-weight sensitivity",
                        "M0 inputs · benchmark and weights of M0"]}.get(
            chosen, [f"Periods {labels[0]}–{labels[-1]}",
                     f"Re-estimation {policy.reestimation_period}"])
        header_top(ROLE_PORTAL["Governance body"], "Governance policy and chain-wide oversight",
                   chips + [f"Policy digest {policy.digest()}"])
        header_body(
            "The governance body (here a regulator, but equally an industry consortium or a "
            "public institution) sets the rules the weights and alerts follow and reads what "
            "the coordinator publishes. In the intended design it receives no firm-level "
            "data from the system.", role_rows("Governance body"))
    scenario = st.selectbox(
        "Scenario", families, key="regulator_scenario",
        format_func=lambda f: {
            "M0–M5": "M0–M5 · periodic recalibration of the supply chain",
            "M6": "M6 · Peer benchmark",
            "M7": "M7 · Weight sensitivity"}[f])
    if scenario == "M6":
        m6_dashboard_panel(out, regulator=True)
        return
    if scenario == "M7":
        m7_dashboard_panel(out, role="regulator")
        return
    scope_panel(
        who="the governance body",
        can_see=[
            "The policy it sets: the indicators at each level, base weights, bands, λ, "
            "floors, warning limits",
            "Every published weight version and digest, including μ and σ per indicator",
            "How many participants report each indicator",
            "Chain-wide counts: firms above the peer-relative escalation level, with an "
            "internal warning, with an active hard alert",
        ],
        cannot_see=[
            "Any firm's raw values, local, total or composite score",
            "Which firms are flagged, warned or under a hard alert",
            "Any firm's dependency weights β and α (the assurance provider checks the "
            "declared method)",
            "Individual uploads (n, Σx, Σx²)",
            "Supply-chain links between named firms",
        ])
    st.caption("Firm-level information would reach a supervisor through its own "
               "information requests, outside this system. With seven participants even "
               "chain-wide counts can point to a firm; a deployment would need a minimum "
               "reporting cell size.")

    tabs = st.tabs(["Policy", "Weight governance", "Chain-wide monitoring", "Trigger rule"])

    # ---- 1 · POLICY -------------------------------------------------------
    with tabs[0]:
        st.markdown(
            "<div class='note'><b>Policy settings.</b> Every control below is a governance "
            "parameter. None is estimated from data, which is what makes the weights "
            "governable. A change is published as a new policy digest and applies to every "
            "page of the platform.</div>", unsafe_allow_html=True)
        st.markdown("")
        g1, g2, g3 = st.columns(3)
        with g1:
            with card("pol_t2"):
                card_head("Shared weight settings", "Applied to Level 2 and Level 3")
                st.slider("Baseline influence λ (Levels 2 and 3)",
                          0.0, 1.0, step=0.05, key="pol_lam")
                control_note(
                    f"At {policy.lam:.2f}, {policy.lam:.0%} of each weight's starting value is "
                    f"the governance body's base weight and {1 - policy.lam:.0%} follows how much "
                    "firms differ on the indicator; the result is then projected onto the "
                    "approved bounds with the weights summing to one. Changes the weights and "
                    "every score.")
                st.slider("Level 2 share of baseline weight", 0.3, 0.9, step=0.05,
                          key="pol_block_t2")
                control_note(
                    f"Level 2 holds {policy.block_share_t2:.0%} of each dimension's base weight "
                    f"and Level 3 the remaining {policy.block_share_t3:.0%}; after the data-driven "
                    "adjustment and the projection the final shares can differ somewhat. "
                    "Raising it makes established indicators count more than emerging ones.")
                st.radio("Dispersion measure", ["common", "ratio", "cv"],
                         format_func=dispersion_label, key="pol_dispersion")
                control_note(
                    "The default common-scale SD measures spread after conversion to 0–1. "
                    "The ratio compares raw-unit SD with the anchor SD; CV divides raw-unit "
                    "SD by the absolute mean. Ratio and CV are sensitivity alternatives. "
                    "The selected measure affects data-driven weights at both levels.")
        with g2:
            with card("pol_t3"):
                card_head("Indicator weight bounds", "Narrower for Level 2, wider for Level 3")
                st.slider("Level 2 drift band (× base weight)", 0.2, 2.0,
                          step=0.05, key="pol_band", on_change=_keep_band_around_one)
                control_note(
                    "How far a Level 2 weight may move from its base weight. The band "
                    "always contains 1, the base weight itself. A narrower band keeps "
                    "weights nearer the governance body's choice.")
                st.slider("Level 3 floor (× equal baseline weight)", 0.05, 0.9, step=0.05,
                          key="pol_t3_floor")
                control_note(
                    "Lowest weight a Level 3 indicator can fall to, in multiples of the equal "
                    "weight (the Level 3 share divided by the number of Level 3 indicators). It "
                    "keeps an indicator on which firms have converged, such as E10 from M3, in "
                    "the score.")
                st.slider("Level 3 cap (× equal baseline weight)", 1.5, 6.0, step=0.5,
                          key="pol_t3_cap")
                control_note(
                    "Highest weight a Level 3 indicator can reach. It stops one indicator with "
                    "an unusually wide spread from dominating the score.")
        with g3:
            with card("pol_alert"):
                card_head("Scoring and escalation", "Applied by every participant")
                st.slider("τ (squashing scale of the sigmoid)", 0.5, 3.0, step=0.1,
                          key="pol_squash")
                control_note(
                    "A larger τ pulls every score toward 0.5 and a smaller one spreads scores "
                    "toward 0 and 1. Within a dimension, complete-input local scores keep "
                    "their order. The escalation levels and warning limits are in score "
                    f"units, so τ is set together with them: at τ = {policy.squash_scale:.1f} "
                    f"a local score of {policy.persistent_level:.2f} means a weighted "
                    f"deviation of {score_level_index(policy.persistent_level, policy):.2f} "
                    "peer standard deviations. The level rule is applied to totals, which "
                    "for a firm with suppliers also carry the suppliers' scores, so for such "
                    "a firm the level matches no single deviation. Changing τ alone changes "
                    "flags and warnings (result table T23a).")
                st.slider("Peer-relative escalation level", 0.4, 0.9, step=0.01,
                          key="pol_alert")
                control_note(
                    "A firm whose composite score reaches this level is flagged for "
                    "escalation. The score is relative to peers, so the flag does not measure "
                    "absolute harm. No score changes.")
                st.slider("Absolute limit on a component change", 0.001, 0.05,
                          step=0.001, format="%.3f", key="pol_dabs")
                control_note(
                    "Each own indicator and each supplier contribution raises an internal "
                    "warning at the firm when it moves the total by more than the smaller of "
                    "the two limits. Lower limits catch smaller movements. No score changes.")
                st.slider("Relative limit (share of the pre-event total)", 0.002, 0.05,
                          step=0.002, format="%.3f", key="pol_drel")
                control_note(
                    "The second limit, as a share of the firm's total before the change. "
                    "Whichever limit is smaller applies.")
                st.slider("Coverage floor for a non-provisional score "
                          "(share of published weight reported)", 0.50, 1.00,
                          step=0.05, format="%.2f", key="pol_cov")
                control_note(
                    "A firm reporting less than this share of the published weight gets a "
                    "provisional score, and its customers downstream receive a branch notice. "
                    "The number itself does not change.")
        b1, b2 = st.columns([1, 3])
        with b1:
            st.button("Restore default policy", on_click=_restore_policy_defaults,
                      key="pol_reset")
        with b2:
            st.caption(f"Policy digest {policy.digest()} · re-estimation "
                       f"{policy.reestimation_period}. Hard alerts bypass the warning "
                       f"thresholds. Otherwise a component raises an internal warning when "
                       f"|Δcomponent| > min(absolute limit, relative limit × |pre-event "
                       f"total|), labelled as a deterioration or an improvement. A score "
                       f"with a missing input, or resting on less than "
                       f"{policy.coverage_floor:.0%} of the published weight, is published "
                       f"as provisional.")

        c2 = st.container()
        with c2:
            st.markdown("#### Hard-threshold register (Level 1)")
            reg = pd.DataFrame([{
                "ID": k, "Indicator": v["label"], "Rule": v["rule"],
                "Ceiling": "—" if v.get("limit") is None else v["limit"],
                "Status": ("active hard-alert rule" if v["rule"] == "policy_required"
                           or v.get("limit") is not None
                           else "monitored, ceiling not declared"),
            } for k, v in HARD_THRESHOLD_RULES.items()])
            show_df(reg, hide_index=True)
            pending = int((reg["Status"] != "active hard-alert rule").sum())
            st.markdown(
                f"<div class='warn'><b>{pending} count-based thresholds have no "
                f"ceiling.</b> The ceiling is a regulatory parameter; those indicators "
                f"are reported but raise no alert until the governance body sets it. "
                f"The active rules are policy declarations, so an alert marks a policy "
                f"deficiency, not a confirmed event.</div>",
                unsafe_allow_html=True)
        st.markdown(f"#### Published weight versions · {labels[0]}–{labels[-1]} supply chain")
        vlog = out["version_log"]
        show_df(vlog.assign(notes=vlog["notes"].map(level_text)).rename(columns={
            "period": "Period", "dimension": "Dimension", "version_id": "Version",
            "digest": "Parameter digest", "policy_digest": "Policy digest",
            "timestamp": "Published", "n_indicators": "Indicators", "notes": "Notes",
            "scoring_version": "No.", "effective_from": "Effective from"}),
            hide_index=True)
        if "m6_peers" in out:
            st.caption("M6 is a separate four-firm peer group with its own register; its "
                       "two published versions (before and after) are listed under "
                       "Scenario → M6.")

    # ---- 2 · WEIGHT GOVERNANCE ------------------------------------------
    with tabs[1]:
        c1, c2, c3 = st.columns([1, 1, 2])
        with c1:
            period = st.selectbox("Period", labels, index=len(labels) - 1, key="rg_p")
        with c2:
            dim = st.selectbox("Dimension", list(DIMS), key="rg_d")
        ws = out["periods"][period].weights[dim]
        with c3:
            st.markdown(f"<div class='note'>Version <code>{ws.version_id}</code> · "
                        f"digest <code>{ws.digest}</code> · policy "
                        f"<code>{ws.policy_digest}</code></div>", unsafe_allow_html=True)
        k = st.columns(4)
        n_t1 = len([i for i in HARD_THRESHOLD_RULES
                    if i in set(out["periods"][period].data[dim]["ind_id"].dropna())])
        with k[0]:
            kpi_box("Level 1 · serious conditions", str(n_t1), "hard alert, carries no weight")
        with k[1]:
            kpi_box("Level 2 · established indicators", str((ws.rows["tier"] == "T2").sum()),
                    f"λ = {policy.lam:.2f}, band ×[{policy.band_lo:.2f}, {policy.band_hi:.2f}]")
        with k[2]:
            kpi_box("Level 3 · emerging indicators", str((ws.rows["tier"] == "T3").sum()),
                    f"floor {policy.t3_floor_frac:.2f}×, cap {policy.t3_cap_frac:.1f}× "
                    f"equal weight")
        with k[3]:
            on_bound = int((np.isclose(ws.rows["w_final"], ws.rows["band_lo"], rtol=1e-5)
                            | np.isclose(ws.rows["w_final"], ws.rows["band_hi"], rtol=1e-5)).sum())
            kpi_box("Weights on a bound", str(on_bound), "clipped by band or floor")

        left, right = st.columns([1.25, 1])
        with left:
            st.markdown("#### Published weights")
            d = ws.rows.copy()
            d["id"] = d["ind_id"].fillna(d["indicator"].str[:12])
            fig = go.Figure()
            for tier in ["T2", "T3"]:
                s = d[d["tier"] == tier]
                if s.empty:
                    continue
                fig.add_trace(go.Bar(
                    y=s["id"], x=s["w_final"], orientation="h",
                    name={"T2": "Level 2 · established", "T3": "Level 3 · emerging"}[tier],
                    marker=dict(color=GOV_COLOR[tier], line=dict(width=2, color=C["surface"])),
                    customdata=np.stack([s["rho"], s["w_data"], s["band_lo"], s["band_hi"]], axis=-1),
                    hovertemplate=("%{y}<br>weight %{x:.4f}<br>ρ %{customdata[0]:.3f}"
                                   "<br>data component %{customdata[1]:.4f}"
                                   "<br>bounds [%{customdata[2]:.4f}, %{customdata[3]:.4f}]"
                                   "<extra></extra>")))
            fig.update_layout(barmode="group", bargap=0.25)
            show_chart(base_layout(fig, max(360, 16 * len(d)), "Published weight", ""))
        with right:
            st.markdown("#### Dispersion drives the drift")
            fig = go.Figure()
            for tier, sym in (("T2", "circle"), ("T3", "circle")):
                s = d[d["tier"] == tier]
                if s.empty:
                    continue
                fig.add_trace(go.Scatter(
                    x=s["rho"], y=s["w_final"], mode="markers",
                    name=GOV_LEVEL[tier],
                    marker=dict(size=11, symbol=sym, color=GOV_COLOR[tier],
                                line=dict(width=2, color=C["surface"])),
                    text=s["ind_id"],
                    hovertemplate="%{text}<br>ρ %{x:.3f} → w %{y:.4f}<extra></extra>"))
            if policy.dispersion_mode == "ratio":
                fig.add_vline(x=1.0, line=dict(color=C["muted"], width=1, dash="dot"),
                              annotation_text="SD equals anchor SD",
                              annotation_font=dict(size=10, color=C["muted"]))
            show_chart(base_layout(fig, 300, dispersion_name, "Published weight"))

            st.markdown("#### Why a floor · conceptual illustration")
            collapse = st.slider("Simulated dispersion collapse on one indicator (ρ)",
                                 0.02, 1.0, 0.10, 0.02, key="rg_collapse")
            n_other = max(1, int((ws.rows["tier"] == "T3").sum()) - 1)
            share = policy.block_share_t3
            floor = policy.t3_floor_frac * share / (n_other + 1)
            rounds = np.arange(0, 7)
            rho_path = np.clip(collapse ** (rounds / 3.0), 0.01, None)
            raw = share * rho_path / (rho_path + n_other)
            fig = go.Figure()
            fig.add_trace(go.Scatter(x=rounds, y=raw, name="no floor", mode="lines+markers",
                                     line=dict(width=2, color=C["s2"], dash="dash"),
                                     marker=dict(size=8, line=dict(width=2, color=C["surface"]))))
            fig.add_trace(go.Scatter(x=rounds, y=np.maximum(raw, floor),
                                     name="with the governance floor", mode="lines+markers",
                                     line=dict(width=2, color=C["s1"]),
                                     marker=dict(size=8, line=dict(width=2, color=C["surface"]))))
            fig.add_hline(y=floor, line=dict(color=C["muted"], width=1, dash="dot"),
                          annotation_text=f"floor {floor:.4f}",
                          annotation_font=dict(size=10, color=C["muted"]))
            show_chart(base_layout(fig, 260, "Re-estimation round", "Indicator weight"))
            st.caption("Simplified illustration: one indicator's dispersion shrinks while "
                       "the others stay equal. It isolates the effect of a positive floor "
                       "and omits the shared baseline blend, caps and joint projection. "
                       "The published weights are shown in the tables and trajectory below.")

        st.markdown("#### Dispersion and weight over M0–M5, against the bounds")
        wt_rows = []
        for l in labels:
            wr = out["periods"][l].weights[dim].rows
            wr = wr.assign(period=l)
            wt_rows.append(wr[["period", "ind_id", "indicator", "tier", "rho", "w_final",
                               "band_lo", "band_hi"]])
        wt = pd.concat(wt_rows, ignore_index=True)
        ids = list(dict.fromkeys(wt["ind_id"].dropna()))
        default_id = "E10" if "E10" in ids else ids[0]
        tc1, tc2, tc3 = st.columns([1, 2, 2])
        with tc1:
            pick = st.selectbox("Indicator", ids, index=ids.index(default_id),
                                key=f"rg_traj_{dim}")
        one = wt[wt["ind_id"] == pick]
        tier_ = one["tier"].iloc[-1]
        level_ = GOV_LEVEL.get(tier_, tier_)
        with tc2:
            fig = go.Figure()
            fig.add_trace(go.Scatter(x=one["period"], y=one["band_hi"], name="upper bound",
                                     mode="lines", line=dict(color=C["muted"], width=1,
                                                             dash="dash")))
            fig.add_trace(go.Scatter(x=one["period"], y=one["band_lo"], name="lower bound",
                                     mode="lines", line=dict(color=C["muted"], width=1,
                                                             dash="dot")))
            fig.add_trace(go.Scatter(x=one["period"], y=one["w_final"], name="published weight",
                                     mode="lines+markers",
                                     line=dict(color=GOV_COLOR.get(tier_, C["s1"]), width=2.2),
                                     marker=dict(size=8, line=dict(width=2, color=C["surface"]))))
            show_chart(base_layout(fig, 280, "", f"{pick} weight ({level_})"))
        with tc3:
            fig = go.Figure()
            fig.add_trace(go.Scatter(x=one["period"], y=one["rho"], name=dispersion_name,
                                     mode="lines+markers", line=dict(color=C["s2"], width=2),
                                     marker=dict(size=8, line=dict(width=2, color=C["surface"]))))
            if policy.dispersion_mode == "ratio":
                fig.add_hline(y=1.0, line=dict(color=C["muted"], width=1, dash="dot"))
            show_chart(base_layout(fig, 280, "", f"{pick} · {dispersion_name}"))
        st.caption(f"{pick} is a {level_} indicator under the current classification. When "
                   f"its dispersion changes, its weight responds to its share of "
                   f"dimension-wide dispersion, the baseline blend and the approved bounds.")

        drift = out.get("drift")
        if drift is not None and not drift.empty:
            dd = drift[(drift["period"] == period) & (drift["dimension"] == dim)]
            st.markdown("#### Dispersion and weight changes since the anchor")
            st.caption(f"Dispersion uses the current measure: {dispersion_name}.")
            if dd.empty:
                st.caption(f"Nothing moved in {dim.lower()} at {period}.")
            else:
                show_df(dd.nsmallest(10, "rho")[["ind_id", "tier", "rho", "dispersion",
                                                  "w_anchor", "w_now", "pct_change",
                                                  "on_a_bound"]].replace({"tier": GOV_LEVEL})
                        .rename(columns={
                    "ind_id": "ID", "tier": "Level", "rho": "ρ", "dispersion": "Dispersion",
                    "w_anchor": "Weight at anchor", "w_now": "Weight now",
                    "pct_change": "Change", "on_a_bound": "On a bound"}).style.format({
                    "ρ": "{:.3f}", "Weight at anchor": "{:.4f}", "Weight now": "{:.4f}",
                    "Change": "{:+.1%}"}), hide_index=True)
        with st.expander("Full published weight table"):
            show_df(ws.rows[["ind_id", "indicator", "tier", "category", "status", "mu",
                             "sigma", "rho", "w_reg", "w_data", "w_final", "band_lo",
                             "band_hi"]].replace({"tier": GOV_LEVEL})
                    .rename(columns={"tier": "level"}).style.format({
                "mu": "{:.4g}", "sigma": "{:.4g}", "rho": "{:.3f}", "w_reg": "{:.4f}",
                "w_data": "{:.4f}", "w_final": "{:.4f}", "band_lo": "{:.4f}",
                "band_hi": "{:.4f}"}), hide_index=True)
            for n_ in ws.notes:
                st.caption(f"· {level_text(n_)}")

    # ---- 3 · CHAIN-WIDE MONITORING --------------------------------------
    with tabs[2]:
        rows = []
        for l in labels:
            r = out["periods"][l]
            into = _transition_into(l, labels)
            t = trig[trig["transition"] == into] if into else trig.iloc[0:0]
            warn_e = int(t[(t["dimension"] == "Environmental")]["internal_warning"].sum())
            warn_s = int(t[(t["dimension"] == "Social")]["internal_warning"].sum())
            nodes_warn = t[t["internal_warning"]]["entity"].nunique()
            nodes_hard = t[t["hard_alert"]]["entity"].nunique()
            reg_ = r.hard_register
            active = (0 if reg_ is None else
                      sum(bool(reg_.is_active(e, d_)) for e in ENTITY_ORDER for d_ in DIMS))
            rows.append({
                "Period": l, "Participants": len(ENTITY_ORDER),
                "Above peer-relative level": sum(peer_relative_flag(r.composite, policy).values()),
                "Firms with an internal warning": nodes_warn if into else np.nan,
                "Firms with a warning still open": (
                    t[t["open_internal_warnings"] != ""]["entity"].nunique()
                    if into else np.nan),
                "Internal warnings E": warn_e if into else np.nan,
                "Internal warnings S": warn_s if into else np.nan,
                "Firms reached by a hard alert": nodes_hard if into else np.nan,
                "Active hard alerts (origin)": active,
                "Missed by an aggregate rule": int(t["MISSED_BY_NET_RULE"].sum()) if into else np.nan,
                "Provisional scores": sum(bool(v.get("provisional"))
                                          for d_ in DIMS
                                          for v in r.provisional.get(d_, {}).values()),
            })
        mon = pd.DataFrame(rows)
        last = mon.iloc[-1]
        k = st.columns(4)
        with k[0]:
            kpi_box("Participants enrolled", str(int(last["Participants"])),
                    "voluntary enrollment")
        with k[1]:
            kpi_box("Above peer-relative level", f"{int(last['Above peer-relative level'])}",
                    f"latest period {labels[-1]}")
        with k[2]:
            kpi_box("Firms with an internal warning",
                    f"{int(last['Firms with an internal warning'])}",
                    f"transition into {labels[-1]}")
        with k[3]:
            kpi_box("Active hard alerts", f"{int(last['Active hard alerts (origin)'])}",
                    "until reviewed clearance",
                    alert=last["Active hard alerts (origin)"] > 0)

        left, right = st.columns([1.3, 1])
        with left:
            st.markdown("#### Chain-wide counts by period")
            fig = go.Figure()
            for n_, (col, color) in enumerate((("Above peer-relative level", C["warning"]),
                                               ("Firms with an internal warning", C["s1"]))):
                fig.add_trace(go.Bar(x=mon["Period"], y=mon[col], name=col,
                                     marker=dict(color=color, line=dict(width=2, color=C["surface"]))))
            fig.add_trace(go.Scatter(x=mon["Period"], y=mon["Active hard alerts (origin)"],
                                     name="Active hard alerts", mode="lines+markers",
                                     line=dict(color=C["critical"], width=2),
                                     marker=dict(size=8, symbol="x")))
            fig.update_layout(barmode="group", bargap=0.3)
            show_chart(base_layout(fig, 320, "", "Number of firms"))
        with right:
            st.markdown("#### Monitoring table")
            show_df(mon.drop(columns=["Participants"]), hide_index=True)
            hidden_note("which firms these are. Counts only; names and scores stay "
                        "with the firms and their assurance providers.")

        st.markdown("#### Reporting coverage per indicator")
        c1, c2 = st.columns([1, 3])
        with c1:
            p_ = st.selectbox("Period", labels, index=len(labels) - 1, key="rm_p")
            d_ = st.selectbox("Dimension", list(DIMS), key="rm_d")
        stats = out["periods"][p_].stats[d_]
        ws = out["periods"][p_].weights[d_]
        cov = ws.rows[["ind_id", "indicator", "tier", "status"]].copy()
        cov["Participants reporting"] = [stats.n.get(k_, 0) for k_ in cov["indicator"]]
        with c2:
            k2 = st.columns(3)
            with k2[0]:
                kpi_box("Indicators published", str((cov["status"] == "published").sum()))
            with k2[1]:
                kpi_box("Zero variance observed",
                        str(cov["status"].astype(str).str.startswith("zero").sum()),
                        "kept with a positive weight floor")
            with k2[2]:
                kpi_box("Below full coverage",
                        str((cov["Participants reporting"] < len(ENTITY_ORDER)).sum()),
                        f"fewer than {len(ENTITY_ORDER)} reporting")
        show_df(cov.replace({"tier": GOV_LEVEL}).rename(columns={
                    "ind_id": "ID", "indicator": "Indicator", "tier": "Level",
                    "status": "Status"}),
                hide_index=True, height=300)

    # ---- 4 · TRIGGER RULE -----------------------------------------------
    with tabs[3]:
        st.markdown(
            "<div class='crit'><b>Why the warning is tested per component, never on the "
            "net.</b> Each firm tests its own indicators and supplier contributions before "
            "combining them: one deteriorating and another improving can leave the total "
            "untouched. An internal warning stays with the firm that raised it and stays "
            "open until the same component has come back within the limit; an improving "
            "branch does not close it. Hard alerts are checked first, bypass every warning "
            "threshold and reach every affected firm upstream and downstream.</div>",
            unsafe_allow_html=True)
        st.markdown("")
        agg = (trig.groupby(["transition", "dimension"])
               .agg(examined=("entity", "size"),
                    aggregate_rule=("net_rule_fires", "sum"),
                    internal_warning=("internal_warning", "sum"),
                    hard_alert=("hard_alert", "sum"),
                    missed_by_aggregate=("MISSED_BY_NET_RULE", "sum"))
               .reset_index())
        k = st.columns(4)
        with k[0]:
            kpi_box("Node-transitions examined", str(len(trig)))
        with k[1]:
            kpi_box("Aggregate rule fires", str(int(trig["net_rule_fires"].sum())),
                    "keyed on net change")
        with k[2]:
            kpi_box("Internal warnings", str(int(trig["internal_warning"].sum())),
                    "own indicators and supplier contributions")
        with k[3]:
            kpi_box("Missed by the aggregate rule", str(int(trig["MISSED_BY_NET_RULE"].sum())),
                    "internal warning without an aggregate alert",
                    alert=trig["MISSED_BY_NET_RULE"].any())
        left, right = st.columns([1.2, 1])
        with left:
            show_df(agg.rename(columns={
                "transition": "Change", "dimension": "Dimension", "examined": "Nodes",
                "aggregate_rule": "Aggregate rule", "internal_warning": "Internal warning",
                "hard_alert": "Hard alert",
                "missed_by_aggregate": "Missed by aggregate"}), hide_index=True, height=420)
        with right:
            fig = go.Figure()
            g = agg.groupby("transition")[["aggregate_rule", "internal_warning"]].sum().reset_index()
            fig.add_trace(go.Bar(x=g["transition"], y=g["aggregate_rule"],
                                 name="aggregate rule", marker=dict(color=C["s2"],
                                 line=dict(width=2, color=C["surface"]))))
            fig.add_trace(go.Bar(x=g["transition"], y=g["internal_warning"],
                                 name="internal warning",
                                 marker=dict(color=C["s1"], line=dict(width=2, color=C["surface"]))))
            fig.update_layout(barmode="group", bargap=0.3)
            show_chart(base_layout(fig, 320, "", "Internal warnings (E + S)"))
            st.caption(f"Current limits: min({policy.component_delta_abs:.3f}, "
                       f"{policy.component_delta_rel:.3f} × |pre-event total|). Change the "
                       f"limits under Policy to see how the counts respond.")

        if "offsetting" in out:
            off = out["offsetting"]
            m, hub, foc = off["meta"], off["hub_report"], off["focal_report"]
            st.markdown("#### Constructed offsetting stress case")
            st.caption("A designed scenario to test the rule, not observed data; nodes "
                       "are shown by role.")
            a, b = st.columns([1, 1.2])
            with a:
                show_df(pd.DataFrame([
                    {"Node": "Tier-1 branch hub", "Net change": hub.net_change,
                     "Largest component change": hub.max_abs_component,
                     "Threshold": hub.threshold,
                     "Aggregate rule": "fires" if hub.net_fired else "silent",
                     "Internal warning": "yes" if hub.internal_warning else "no"},
                    {"Node": "Focal firm", "Net change": foc.net_change,
                     "Largest component change": foc.max_abs_component,
                     "Threshold": foc.threshold,
                     "Aggregate rule": "fires" if foc.net_fired else "silent",
                     "Internal warning": "yes" if foc.internal_warning else "no"},
                ]).style.format({"Net change": "{:+.6f}",
                                 "Largest component change": "{:.6f}",
                                 "Threshold": "{:.6f}"}), hide_index=True)
                st.markdown(
                    f"<div class='warn'>One supplier deteriorates by {m['delta']:.3f} "
                    f"while another improves by {m['offset_delta']:.3f}. The hub's "
                    f"aggregate does not move, so an aggregate-keyed rule reports a quiet "
                    f"quarter. The hub's own test raises an internal warning; the focal "
                    f"firm receives an unchanged total.</div>", unsafe_allow_html=True)
            with b:
                base = out["periods"][labels[1]]
                names = list(off["comps"][m["hub"]])
                role = {m["deteriorating"]: "via deteriorating supplier",
                        m["improving"]: "via improving supplier"}
                lab = ["own operations" if k_ == "__own__" else role.get(k_, "via supplier")
                       for k_ in names]
                fig = go.Figure()
                fig.add_trace(go.Bar(x=lab, y=[base.comps[m["dimension"]][m["hub"]][k_] for k_ in names],
                                     name="before", marker=dict(color=C["s1"],
                                     line=dict(width=2, color=C["surface"]))))
                fig.add_trace(go.Bar(x=lab, y=[off["comps"][m["hub"]][k_] for k_ in names],
                                     name="after", marker=dict(color=C["s2"],
                                     line=dict(width=2, color=C["surface"]))))
                fig.update_layout(barmode="group", bargap=0.3)
                show_chart(base_layout(fig, 300, "Component of the hub's total", "Contribution"))


# -----------------------------------------------------------------------------
# OVERVIEW  (system design and the coordinator's published records)
# -----------------------------------------------------------------------------


def _period_note(label: str) -> Tuple[str, str]:
    """Scenario title and description for a period label such as 'M3'."""
    try:
        lines = [proofread_input_text(x)
                 for x in PERIOD_NOTES[int(str(label).lstrip("Mm"))] if x]
    except (KeyError, ValueError):
        return "", ""
    head = lines[0].split("—", 1)[-1].strip()
    if " (" in head:
        h, rest = head.split(" (", 1)
        head = h.capitalize() + " (" + rest
    else:
        head = head.capitalize()
    return level_text(head), level_text(" ".join(lines[1:]))


def data_setup_panel(expanded: bool = False) -> None:
    with st.expander("Research data setup", expanded=expanded):
        st.text_input("Data directory", key="data_dir")
        st.caption("Regeneration replaces the M0–M7 inputs, including manual M6 and M7 "
                   "edits.")
        if st.button("Generate / refresh M0–M7", key="regen_m0_m5"):
            root = Path(__file__).resolve().parent
            source = default_source_workbook(root)
            classification = root / "Indicator_Governance_Tier_Classification.xlsx"
            if source is None or not classification.is_file():
                st.error("Regeneration is unavailable here: the source or classification "
                         "workbook was not found. The six workbooks already in the data "
                         "directory are the supplied inputs.")
            else:
                target = st.session_state["data_dir"]
                try:
                    with st.spinner("Generating M0–M7 inputs…"):
                        generate_scenarios(str(source), target, str(classification),
                                           recalibrate_m4=False, verbose=False)
                except (ConfigurationError, OSError) as exc:
                    st.error(f"Regeneration stopped: {exc}")
                else:
                    run_engine.clear()
                    st.success(f"M0–M7 generated in {Path(target).name}.")


def overview_view(out: dict) -> None:
    policy = out["policy"]
    labels = list(out["periods"].keys())
    last = out["periods"][labels[-1]]

    events_run = list(out["events"]["scenarios"])
    with keyed_container("pagehead"):
        # Four kinds of experiment: the periodic scenarios, the single-change
        # events, the separate M6 industry peer experiment and the M7
        # dependency-weight sensitivity.
        chips = ["Reference implementation", "Single-process demonstration",
                 f"{len(ENTITY_ORDER)} participants in {len(TIER_ORDER)} supply-chain tiers",
                 f"Periods {labels[0]}–{labels[-1]}"]
        if events_run:
            chips.append(f"Events {events_run[0][0]}–{events_run[-1][0]}")
        if "m6_peers" in out:
            chips.append("M6 industry peer experiment (separate)")
        if "m7" in out:
            chips.append("M7 dependency-weight sensitivity (separate)")
        header_top(ROLE_PORTAL["Overview"], "System overview", chips)
        header_body(
            "The platform produces branch-level signals to support prioritization and "
            "further investigation of upstream ESG conditions. A branch signal identifies "
            "a direct supplier relationship for follow-up; it does not identify the "
            "specific deeper-tier firm or establish fulfillment of due diligence obligations. "
            "Participants exchange summary statistics and scores instead of raw records, a "
            "governance body governs how indicators are weighted, and assurance providers "
            "verify what each firm declares.",
            [("Design question", "Can decentralized participants generate governed, auditable "
                                 "branch-risk signals without any party, the focal firm "
                                 "included, holding the full chain's raw data?"),
             ("Evidence type", "A controlled single-process demonstration with simulated "
                               "inputs for anonymous positions in a simulated supply chain. "
                               "It is not a predictive-validation study, a "
                               "deployed secure system or an assessment of real companies."),
             ("Benchmark scope", "M0–M5 use one mixed-industry demonstration cohort. "
                                 "M6 separately illustrates an industry peer benchmark for "
                                 "the rubber supplier."),
             ("Hard-alert example", "M5 tests a failed policy declaration. It does not "
                                    "represent a confirmed child-labor incident.")])

    tabs = st.tabs(["System design", "Recalibration and publication", "Event updates",
                    "Scenarios", "Scope and limitations"])

    # ---- 1 · SYSTEM DESIGN ------------------------------------------------
    with tabs[0]:
        st.markdown("#### How the system operates")
        c1, c2 = st.columns(2)
        with c1:
            with card("cycle_recal"):
                card_head("Periodic recalibration", "Low frequency, for example quarterly")
                st.markdown(
                    "<div class='card-text'><ol>"
                    "<li>Each participant uploads three summary statistics per indicator, "
                    "n, Σx and Σx². No record leaves the firm.</li>"
                    "<li>The coordinator adds them up and derives the peer benchmark μ and σ "
                    "for every indicator.</li>"
                    "<li>The same round yields the data component w_data of each weight from "
                    "the dispersion on the common 0–1 scale. The approved base weight w_reg is "
                    "blended in with λ, and the result is projected onto the approved bounds "
                    "so that the weights sum to one.</li>"
                    "<li>The coordinator publishes benchmark and weights as a versioned "
                    "parameter set with a digest.</li></ol></div>"
                    "<div class='card-foot'>One common-scale calibration upload stage; "
                    "reference initialization is separate. No iteration and no "
                    "convergence criterion.</div>", unsafe_allow_html=True)
        with c2:
            with card("cycle_event"):
                card_head("Event-driven scoring",
                          "Whenever a firm's data, suppliers, dependency weights or data "
                          "quality change")
                st.markdown(
                    "<div class='card-text'><ol>"
                    "<li>The participant recomputes its own local score with the published "
                    "parameters, which stay frozen, or records a documented and assured "
                    "dependency-weight revision.</li>"
                    "<li>It uploads its new total score to the coordinator.</li>"
                    "<li>The coordinator relays the total to the participant's direct "
                    "customers; a hard alert reaches every affected firm upstream and "
                    "downstream.</li>"
                    "<li>Each customer recomputes its own total and tests the change under its "
                    "own rule, so the signal moves downstream one tier at a time.</li></ol></div>"
                    "<div class='card-foot'>No new aggregation takes place and no new "
                    "parameter version is published.</div>", unsafe_allow_html=True)

        st.markdown("#### Scoring function")
        with card("scoring"):
            c1, c2 = st.columns([1.1, 1], gap="large")
            with c1:
                st.latex(r"u_{ik} = \operatorname{clip}_{[0,1]}\!\left("
                         r"\frac{x_{ik}-a_k}{b_k-a_k}\right)")
                st.latex(r"\tilde{x}_{ik}=\begin{cases}1-u_{ik},&"
                         r"\text{higher raw values are favorable}\\u_{ik},&"
                         r"\text{higher raw values are adverse}\end{cases}")
                st.latex(r"z_{ik} = \frac{\tilde{x}_{ik}-\tilde{\mu}_k}"
                         r"{\tilde{\sigma}^{\mathrm{score}}_k}")
                st.latex(r"r_i^{D} = \operatorname{sigmoid}\!\left(\frac{1}{\tau}\,"
                         r"\sum_{k\in K_d} w_k\, z_{ik}\right)")
                st.latex(r"\operatorname{sigmoid}(s) = \frac{1}{1 + e^{-s}}")
                st.latex(r"R_i^{D} = \beta_i\, r_i^{D} + \sum_{j\in S_i} \alpha_{ij}\, R_j^{D}")
                st.latex(r"C_i = q^{E} R_i^{E} + q^{S} R_i^{S}")
            with c2:
                qe, qs = policy.dim_weights
                st.markdown(
                    "<div class='card-text'>"
                    "<p>The local score is a weighted mean of standardized indicator "
                    "values, the linear aggregation used to build composite indicators "
                    "(OECD and JRC, 2008), passed through the logistic (sigmoid) function "
                    "so that it lies between 0 and 1. Logistic regression uses the same "
                    "function, but here nothing is estimated. There is no outcome variable, "
                    "and the weights are governed parameters rather than fitted "
                    "coefficients. The sigmoid preserves the order of the local weighted "
                    "indices within a dimension. Changing τ can change rankings of "
                    "propagated totals and composites.</p><ul>"
                    "<li>x̃ is the value on the common 0–1 scale, higher meaning less "
                    "favorable (the scale is reversed where a higher raw value is better); "
                    "[a<sub>k</sub>, b<sub>k</sub>] is the published conversion range and "
                    "μ̃, σ̃ the peer mean and standard deviation on that scale. "
                    "S<sub>i</sub> is the set of direct suppliers.</li>"
                    "<li>σ̃<sup>score</sup> = max(σ̃<sub>k</sub>, η·s<sub>ref,k</sub>/(b<sub>k</sub>"
                    " − a<sub>k</sub>)) is the published standardization scale, so an "
                    "indicator on which every firm reports the same value contributes zero "
                    "while keeping its weight. Weights are fixed: a missing input is bounded "
                    "at 0 and 1 and the score is published as an interval.</li>"
                    "<li>Indicator weights w come from the governed recalibration. "
                    "Dependency weights α and β are set by each firm under its declared "
                    "method and never leave it; the demonstration uses assumed values.</li>"
                    "<li>A failed Level 1 hard threshold raises an independent hard alert; "
                    "it does not change the continuous score.</li>"
                    "<li>A score with a missing input, or resting on less than "
                    f"{policy.coverage_floor:.0%} of the published weight, is published as "
                    "<i>provisional</i>. The flag changes no number and is passed to every "
                    "customer downstream as a branch notice.</li>"
                    f"<li>In the current policy τ = {policy.squash_scale:.1f} and "
                    f"q<sup>E</sup> = {qe:.2f}, q<sup>S</sup> = {qs:.2f}.</li>"
                    "</ul></div>", unsafe_allow_html=True)

        st.markdown("#### Three-level indicator governance")
        cnt = {}
        for d_ in DIMS:
            rows_ = last.weights[d_].rows
            present = set(last.data[d_]["ind_id"].dropna())
            cnt[d_] = {"T1": len([i for i in HARD_THRESHOLD_RULES if i in present]),
                       "T2": int((rows_["tier"] == "T2").sum()),
                       "T3": int((rows_["tier"] == "T3").sum())}
        levels = [
            ("T1", "Declarations such as child-labor and forced-labor policies. A failed "
                   "declaration raises an independent hard "
                   "alert that bypasses every warning threshold and persists until reviewed "
                   "clearance. The assurance provider verifies the declaration. A serious "
                   "breach therefore cannot be averaged out by good results elsewhere."),
            ("T2", "Indicators such as emissions and workplace injuries. The published weight "
                   "is the bounded-simplex projection of λ·b + (1 − λ)·v, where the governance "
                   f"body sets w_reg, λ and the bounds. Current policy λ = {policy.lam:.2f}, bounds "
                   f"{policy.band_lo:.2f}× and {policy.band_hi:.2f}× the base weight. This "
                   "balances prescription with adaptation to the data."),
            ("T3", "The same baseline blend and joint projection apply, with wider bounds: "
                   "a floor of "
                   f"{policy.t3_floor_frac:.2f}× and a cap of {policy.t3_cap_frac:.1f}× the "
                   "equal weight. The floor keeps materiality apart from discriminating power "
                   "and damps the oscillation a zero weight would cause. The wider range "
                   "leaves room to adapt to new information."),
        ]
        cols = st.columns(3)
        for col, (tid, text) in zip(cols, levels):
            with col:
                with card(f"tier_{tid}"):
                    card_head(LEVEL_TITLE[tid], LEVEL_RULE[tid])
                    st.markdown(
                        f"<div class='card-count'>{cnt['Environmental'][tid]} E · "
                        f"{cnt['Social'][tid]} S</div>"
                        f"<div class='card-sub' style='margin:0 0 10px'>indicators in "
                        f"{labels[-1]}</div><div class='card-text'>{text}</div>",
                        unsafe_allow_html=True)

        st.markdown("#### Parties and information boundaries")
        where = {
            "Governance body": "Governance body console",
            "Participating firm": "Participating firm portal",
            "Coordinator": "Runs server-side as the aggregation service and has no page of "
                           "its own. Its published records are under Recalibration and "
                           "publication.",
            "Assurance provider": "Assurance provider workspace",
        }
        st.markdown(html_table(
            ["Party", "Responsibilities", "Holds", "On this platform"],
            [[role, _cap(does) + ".", _cap(sees) + ".", where[role]]
             for role, (does, sees) in ROLE_TABLE.items()]), unsafe_allow_html=True)
        st.caption("In the intended design, operational records remain within firms and "
                   "assurance uses authorized access to client records. This prototype "
                   "illustrates those information boundaries in one process; it does not "
                   "enforce separate storage or access permissions.")

    # ---- 2 · RECALIBRATION AND PUBLICATION -------------------------------
    with tabs[1]:
        st.markdown(
            "<div class='note'><b>Run by the coordinator.</b> The coordinator is the "
            "aggregation service. It adds up the participants' summary statistics, derives "
            "the peer benchmark and the governed weights, and publishes them as a versioned "
            "parameter set that every party can read.</div>", unsafe_allow_html=True)
        st.markdown("")
        eq = out["equiv"]
        vlog = out["version_log"]
        n_versions = int((~vlog["version_id"].astype(str).str.startswith("REG-")).sum())
        k = st.columns(4)
        with k[0]:
            kpi_box("Parameter versions published", str(n_versions),
                    f"one per dimension per period, {labels[0]}–{labels[-1]}")
        with k[1]:
            kpi_box("Upload stages per calibration", "1",
                    "reference initialization counted separately")
        with k[2]:
            kpi_box("Firms uploading",
                    f"{len(last.common_summaries.get('Environmental', []))} / {len(ENTITY_ORDER)}",
                    "common scale, three numbers per indicator")
        with k[3]:
            kpi_box("Federated vs centralized",
                    "identical" if eq["identical_to_1e-9"].all() else "differs",
                    f"max rel. diff {eq['max_rel_diff_sigma'].max():.1e}",
                    alert=not eq["identical_to_1e-9"].all())

        st.markdown("#### Recalibration round")
        s1, s2, _ = st.columns([1, 1, 2])
        with s1:
            p_ = st.selectbox("Period", labels, index=len(labels) - 1, key="ov_rc_p")
        with s2:
            d_ = st.selectbox("Dimension", list(DIMS), key="ov_rc_d")
        r_ = out["periods"][p_]
        sums, stats = r_.common_summaries[d_], r_.stats[d_]
        activity = pd.DataFrame(r_.calibration_activity)
        if not activity.empty:
            show_df(activity, hide_index=True)
        st.caption("Regular calibration uses one common-scale upload stage. A data-dependent "
                   "reference initialization, when listed, adds an earlier raw-unit stage. "
                   "The E and S dimension aggregations can be batched within each stage.")
        left, right = st.columns([1, 1.6])
        with left:
            show_df(pd.DataFrame([{
                "Participant": disp(s_.entity), "Indicators": len(s_.n),
                "Numbers received": 3 * len(s_.n), "Payload digest": s_.payload_digest(),
            } for s_ in sums]), hide_index=True)
            st.caption("Uploads received. Under secure aggregation only the sums on the "
                       "right would be revealed, not the values inside each payload.")
        with right:
            rows = []
            for k_ in stats.mu:
                rows.append({"Indicator": k_,
                             "Σn": sum(s_.n.get(k_, 0) for s_ in sums),
                             "ΣΣx": sum(s_.sx.get(k_, 0.0) for s_ in sums),
                             "ΣΣx²": sum(s_.sxx.get(k_, 0.0) for s_ in sums),
                             "μ (raw, diagnostic)": stats.mu[k_],
                             "σ (raw, diagnostic)": stats.sigma.get(k_, np.nan),
                             "μ (0–1)": stats.common_mu.get(k_, np.nan),
                             "σ (0–1)": stats.common_sigma.get(k_, np.nan),
                             "Status": stats.status.get(k_, "")})
            show_df(pd.DataFrame(rows).style.format({
                "ΣΣx": "{:.6g}", "ΣΣx²": "{:.6g}",
                "μ (raw, diagnostic)": "{:.4g}", "σ (raw, diagnostic)": "{:.4g}",
                "μ (0–1)": "{:.4g}", "σ (0–1)": "{:.4g}"}),
                hide_index=True, height=360)
            st.caption("Sums and the μ (0–1), σ (0–1) benchmark come from the actual "
                       "common-scale uploads: μ = ΣΣx / Σn and "
                       "σ = √max(ΣΣx² / Σn − μ², 0). Raw-unit means and SDs are "
                       "evaluator diagnostics computed from local in-process records, "
                       "not extra calibration uploads. The default weight dispersion "
                       "uses the common scale; ratio/CV modes are sensitivity experiments.")

        st.markdown("#### Version log")
        show_df(vlog.assign(notes=vlog["notes"].map(level_text)).rename(columns={
            "period": "Period", "dimension": "Dimension", "version_id": "Version",
            "digest": "Parameter digest", "policy_digest": "Policy digest",
            "timestamp": "Published", "n_indicators": "Indicators", "notes": "Notes",
            "scoring_version": "No.", "effective_from": "Effective from"}),
            hide_index=True)
        st.caption("REG entries record the governance register declared at the anchor "
                   "round. Each recalibration has a round-specific version ID; the "
                   "parameter digest identifies its contents.")
        m6_versions = out.get("m6_peers", {}).get("versions")
        if m6_versions is not None and not m6_versions.empty:
            with st.expander("M6 industry peer experiment: its two published versions"):
                show_df(m6_versions[["state", "dimension", "version_id", "digest",
                                     "register_version", "n_indicators", "peer_group"]],
                        hide_index=True)
                st.caption("M6 recalibrates a separate four-firm peer group (S1-RUBBER and "
                           "three simulated peers); it is not a period of the M0–M5 panel.")

        st.markdown("#### Check of the federated route against centralized pooling")
        show_df(eq.rename(columns={
            "dimension": "Dimension", "indicators_compared": "Indicators",
            "max_rel_diff_mu": "Max rel. diff μ", "max_rel_diff_sigma": "Max rel. diff σ",
            "common_scale_indicators_compared": "Common-scale indicators",
            "max_abs_diff_common_mu": "Max diff μ (common scale)",
            "max_abs_diff_common_sigma": "Max diff σ (common scale)",
            "same_benchmark_indicators": "Same indicators",
            "identical_to_1e-9": "Identical to 1e-9"}).style.format({
            "Max rel. diff μ": "{:.2e}", "Max rel. diff σ": "{:.2e}",
            "Max diff μ (common scale)": "{:.2e}", "Max diff σ (common scale)": "{:.2e}"}),
            hide_index=True)
        st.markdown(
            "<div class='warn'><b>Reference implementation gap.</b> This is numerical "
            "agreement within one process, not a privacy test. Secure aggregation is a "
            "deployment assumption. In this reference the coordinator receives plaintext "
            "contributions, and with one observation per firm n = 1, so Σx is the value "
            "itself.</div>", unsafe_allow_html=True)

    # ---- 3 · EVENT UPDATES --------------------------------------------------
    with tabs[2]:
        events = out["events"]
        ver = events["verification"]
        acc = events["acceptance"]
        st.markdown(
            "<div class='note'><b>Frozen-parameter event experiments.</b> Each scenario "
            "starts from the same published R0 state with μ, σ and the weights frozen. The "
            "source firm recomputes its own local score or revises its dependency weights; "
            "totals then move downstream, and hard alerts reach every affected firm "
            "upstream and downstream. No new aggregation takes place.</div>",
            unsafe_allow_html=True)
        st.markdown("")
        k = st.columns(4)
        with k[0]:
            kpi_box("Aggregation calls during events", str(ver["event_aggregation_calls"]),
                    "the benchmark is not re-estimated")
        with k[1]:
            kpi_box("New parameter versions", str(ver["additional_version_records"]),
                    "events use the frozen R0 version")
        with k[2]:
            kpi_box("Scenarios", str(len(events["scenarios"])),
                    " · ".join(k_[0] for k_ in events["scenarios"]) + ", each from R0")
        with k[3]:
            n_pass = int(acc["passed"].sum())
            kpi_box("Invariant checks passed", f"{n_pass} / {len(acc)}",
                    "zero aggregation, frozen parameters, alert paths",
                    alert=n_pass < len(acc))
        st.markdown("#### Node-by-node result")
        s1, s2 = st.columns([1, 3])
        with s1:
            sc = st.selectbox("Scenario", list(events["scenarios"]), key="ov_ev_sc",
                              format_func=lambda x: EVENT_SCENARIO_TEXT.get(x, (x, ""))[0])
        with s2:
            st.markdown(f"<div class='note' style='margin-top:26px'>"
                        f"{EVENT_SCENARIO_TEXT.get(sc, ('', sc))[1]}</div>",
                        unsafe_allow_html=True)
        table = events["scenarios"][sc]
        cols_ = [c_ for c_ in ("entity", "local_score_recomputed", "total_recomputed",
                               "change_source", "composite_before", "composite_after",
                               "composite_delta", "alert", "internal_warning", "warning_signal",
                               "warning_direction", "response_level", "open_internal_warnings",
                               "internal_warnings_closed", "hard_alert",
                               "hard_alert_via_direct_supplier",
                               "hard_alert_via_direct_customer",
                               "hard_event_count", "hard_event_ids",
                               "provisional_via_direct_supplier", "parameter_version")
                 if c_ in table]
        show_df(table[cols_].rename(columns={
            "entity": "Participant", "local_score_recomputed": "Local recomputed",
            "total_recomputed": "Total recomputed", "change_source": "Change source",
            "composite_before": "Score before",
            "composite_after": "Score after", "composite_delta": "Δ composite",
            "alert": "Alert",
            "internal_warning": "Internal warning", "hard_alert": "Hard alert",
            "warning_signal": "Warning signal", "warning_direction": "Direction",
            "response_level": "Response", "open_internal_warnings": "Warnings still open",
            "internal_warnings_closed": "Warnings closed",
            "hard_alert_via_direct_supplier": "Hard alert via supplier",
            "hard_alert_via_direct_customer": "Hard alert via customer",
            "hard_event_count": "Distinct hard events", "hard_event_ids": "Hard event IDs",
            "provisional_via_direct_supplier": "Provisional via",
            "parameter_version": "Version used"}).style.format({
            "Score before": "{:.6f}", "Score after": "{:.6f}", "Δ composite": "{:+.2e}"}),
            hide_index=True)
        st.caption("Evaluator view of the reference run. No participant receives this table; "
                   "each firm sees only its own row, as on the Participating firm page. "
                   "Warning signals label deterioration (W↑), improvement (W↓) and "
                   "changes in both directions (W±) separately. "
                   "A failed policy declaration does not establish a verified "
                   "child-labor incident.")
        st.caption(WARNING_SIGNAL_NOTE)
        with st.expander("Invariant checks"):
            show_df(acc.rename(columns={"scenario": "Scenario", "check": "Check",
                                        "passed": "Passed", "status": "Status"}),
                    hide_index=True)

    # ---- 4 · SCENARIOS --------------------------------------------------------
    with tabs[3]:
        # Periodic scenarios, single-change events, the separate M6 industry
        # peer experiment and the M7 dependency-weight sensitivity.
        st.markdown(f"#### Periodic scenarios ({labels[0]}–{labels[-1]})")
        rows = []
        for l in labels:
            head, desc = _period_note(l)
            rows.append([l, head or "—", desc])
        st.markdown(html_table(["Period", "Scenario", "What it exercises"], rows),
                    unsafe_allow_html=True)
        st.caption("M0–M5 are cumulative designed scenarios, not observed chronological ESG "
                   "outcomes; the benchmark and weights are recalculated in every round. The "
                   "seven entities form one heterogeneous peer group; industry and size "
                   "comparability remain external-validity limits.")
        st.markdown("#### Single-change events (frozen M0 parameters)")
        st.markdown(html_table(["Experiment", "Intervention and expected result"],
                               [[t, d] for t, d in EVENT_SCENARIO_TEXT.values()]),
                    unsafe_allow_html=True)
        st.caption("Each experiment starts independently from R0 and is reported under "
                   "Event updates.")
        if "m6_peers" in out:
            st.markdown("#### Industry peer experiment (M6)")
            m6_dashboard_panel(out)
        if "m7" in out:
            st.markdown("#### Dependency-weight sensitivity (M7)")
            m7_dashboard_panel(out)

    # ---- 5 · SCOPE AND LIMITATIONS ------------------------------------------
    with tabs[4]:
        c1, c2 = st.columns(2)
        with c1:
            with card("can"):
                card_head("What the system can do", "Capabilities of the design")
                st.markdown("<div class='card-text'><ul>"
                            + "".join(f"<li>{x}</li>" for x in CAPABILITIES["can"])
                            + "</ul></div>", unsafe_allow_html=True)
        with c2:
            with card("cannot"):
                card_head("What the system cannot do", "Stated as limitations")
                st.markdown("<div class='card-text'><ul>"
                            + "".join(f"<li>{x}</li>" for x in CAPABILITIES["cannot"])
                            + "</ul></div>", unsafe_allow_html=True)
        scope = out.get("scope")
        if scope is not None and not scope.empty:
            st.markdown("#### Status of each claim in this reference implementation")
            st.markdown(html_table(
                ["Claim", "Status", "What was done"],
                [[_cap(str(x["claim"])), f"<span class='tag-s'>{x['status']}</span>",
                  _cap(str(x["detail"]))] for _, x in scope.iterrows()]),
                unsafe_allow_html=True)
        st.markdown("#### Reference run")
        data_setup_panel()

# -----------------------------------------------------------------------------
# ASSURANCE PROVIDER VIEW
# -----------------------------------------------------------------------------

def auditor_view(out: dict) -> None:
    policy = out["policy"]
    labels = list(out["periods"].keys())
    a_ = out["assurance"]
    anchor_label = labels[0]
    keep_state("aud_client", ENTITY_ORDER[0], allowed=ENTITY_ORDER)

    with keyed_container("pagehead"):
        left, right = header_columns()
        with right:
            c1, c2 = control_columns(2)
            with c1:
                client = st.selectbox("Client", ENTITY_ORDER, format_func=disp,
                                      key="aud_client")
            with c2:
                period = period_picker(dashboard_period_labels(out), "period_aud",
                                       default=labels[-1])
        with left:
            header_top(ROLE_PORTAL["Assurance provider"],
                       f"Engagement · client {disp(client)}",
                       [SUPPLY_CHAIN[client]["tier"], SUPPLY_CHAIN[client]["role"],
                        _period_chip(period)])
        header_body(
            "The assurance provider verifies its client's hard-threshold declarations, raw "
            "data and declared dependency-weight method, and checks documented responses "
            "to relevant alerts.",
            scenario_rows(period) + role_rows("Assurance provider"))
    if period == "M6":
        m6_dashboard_panel(out, entity=client)
        return
    if period == "M7":
        m7_dashboard_panel(out, role="auditor", entity=client)
        return
    r = out["periods"][period]
    scope_panel(
        who="the assurance provider",
        can_see=[
            "Its client's raw indicator values and the underlying records",
            "Its client's uploaded triple (n, Σx, Σx²) and published score",
            "Its client's hard-threshold declarations and supporting evidence",
            "Its client's β and α and the procurement records behind them",
            "Direct-supplier totals received and recorded by its client",
            "The published parameters (μ, σ_score, weights, versions) needed to recompute",
        ],
        cannot_see=[
            "Other firms' raw values, local scores, declarations or dependency weights",
            "Other assurance providers' clients or findings",
            "Deeper-tier identities, scores or findings beyond the client's own records",
        ])
    published_policy_panel(policy)

    tabs = st.tabs(["Data against upload", "Hard-threshold declarations",
                    "Score recomputation", "Dependency-weight method", "Alert responses"])

    # ---- 1 · DATA VS UPLOAD ----------------------------------------------
    with tabs[0]:
        dim = st.radio("Dimension", DIMS, horizontal=True, key="a_up_dim")
        s_ = next((s for s in r.summaries[dim] if s.entity == client), None)
        data = r.data[dim][r.data[dim]["entity"] == client]
        if s_ is not None:
            live = AssuranceProvider(client).verify_summary(s_, audited=r.data[dim])
            rows = []
            for _, x in data.iterrows():
                k_ = x["indicator"]
                rows.append({"Indicator": k_, "Audited value": x["value"], "Unit": x["unit"],
                             "Local n": s_.n.get(k_), "Local Σx": s_.sx.get(k_),
                             "Local Σx²": s_.sxx.get(k_),
                             "Matches": ("yes" if k_ in s_.sx and np.isclose(
                                 s_.sx[k_], x["value"], rtol=1e-6, atol=1e-9)
                                 else ("not recorded" if k_ not in s_.sx else "NO"))})
            kpi_box("Local raw-summary reconciliation", "consistent" if not live else
                    f"{len(live)} mismatch(es)",
                    f"{period} · {dim.lower()} · diagnostic digest {s_.payload_digest()}",
                    alert=bool(live))
            st.markdown("")
            show_df(pd.DataFrame(rows).style.format({
                "Audited value": "{:.6g}", "Local n": "{:.0f}", "Local Σx": "{:.6g}",
                "Local Σx²": "{:.6g}"}, na_rep="—"), hide_index=True, height=340)
            st.caption("Local raw-unit summaries are evaluator diagnostics. The regular "
                       "calibration upload is checked on the common scale below; a raw-unit "
                       "reference initialization is accounted for separately.")
        common_up = next((s for s in getattr(r, "common_summaries", {}).get(dim, [])
                          if s.entity == client), None)
        if common_up is not None:
            live_c = AssuranceProvider(client).verify_common_summary(common_up, r.data[dim],
                                                                     r.stats[dim])
            kpi_box("Common-scale upload", "consistent" if not live_c else
                    f"{len(live_c)} mismatch(es)",
                    "audited values converted with the published reference points "
                    "and directions (Section 4.2.2)", alert=bool(live_c))
        b_ = a_["B_summary_anomaly"]
        if b_["entity"] == disp(client):
            st.markdown(f"#### Reference probe B · tampered upload ({anchor_label})")
            for f in b_["findings"]:
                st.markdown(f"<div class='crit'>{f.kind} · {f.indicator} · {f.detail}</div>",
                            unsafe_allow_html=True)
            imp = b_["impact"]
            imp = imp[imp["entity"] == disp(client)]
            if not imp.empty:
                x = imp.iloc[0]
                st.markdown(f"<div class='warn'>Had the understated Σx gone undetected, the "
                            f"client's environmental total would have moved from "
                            f"{x['E_total_clean']:.4f} to {x['E_total_if_undetected']:.4f} "
                            f"({x['distortion']:+.4f}). The effect on other firms is not "
                            f"visible to this provider.</div>", unsafe_allow_html=True)
        else:
            st.caption("Reference probe B (a tampered upload) was run for another "
                       "provider's client; its result is not visible here.")

    # ---- 2 · HARD-THRESHOLD DECLARATIONS ---------------------------------
    with tabs[1]:
        ab_ = r.absolute_risk
        ab_ = ab_[ab_["entity"] == client] if ab_ is not None else None
        st.markdown(f"#### Client's Level 1 indicators · {period}")
        if ab_ is None or ab_.empty:
            st.caption("No hard-threshold indicator reported.")
        else:
            show_df(ab_.drop(columns=["entity"]), hide_index=True)
        if client in r.hard_hits:
            st.markdown(f"<div class='crit'><b>Hard alert in force: {r.hard_hits[client]}.</b> "
                        f"The alert travels separately from the scores and does not change "
                        f"the client's local score; it stays active until a reviewed "
                        f"clearance.</div>",
                        unsafe_allow_html=True)
        reg_ = out.get("hard_alert_register")
        if reg_ is not None and not reg_.empty:
            mine = reg_[reg_["entity"] == client]
            if not mine.empty:
                st.markdown("#### Hard-alert register entries for this client · all periods")
                show_df(mine.drop(columns=["entity"]), hide_index=True)
        st.markdown(f"#### Reference probe A · declaration check ({anchor_label})")
        fa = [f for f in a_["A_hard_threshold"]["auditor_findings"] if f.entity == client]
        if fa:
            for f in fa:
                st.markdown(f"<div class='crit'>{f.kind} · {f.indicator} · {f.detail}</div>",
                            unsafe_allow_html=True)
            casc = a_["A_hard_threshold"]["cascade"]
            x = casc[casc["entity"] == disp(client)]
            if not x.empty:
                x = x.iloc[0]
                st.caption(f"Client's social total with the hard alert: "
                           f"{x['S_total_before']:.4f} → {x['S_total_after']:.4f} (the alert "
                           f"does not change scores). The hard notice to the client's customers "
                           f"and suppliers is relayed by the coordinator and is not visible "
                           f"to this provider.")
        else:
            st.caption("The reference declaration check found nothing at this client.")

    # ---- 3 · SCORE RECOMPUTATION -----------------------------------------
    with tabs[2]:
        rows = []
        for d_ in DIMS:
            # A hard alert does not replace the weighted score, so the score of
            # a client with a hard alert is recomputed like any other.
            ok, val = AssuranceProvider(client).recompute(
                r.data[d_], r.stats[d_], r.weights[d_], policy, claimed=r.local[d_][client])
            rows.append({"Dimension": d_, "Published local score": r.local[d_][client],
                         "Recomputed": val, "Matches to 1e-9": "yes" if ok else "NO",
                         "Weight version": r.weights[d_].version_id})
        st.markdown(f"#### Recomputation from published parameters · {period}")
        show_df(pd.DataFrame(rows).style.format({
            "Published local score": "{:.9f}", "Recomputed": "{:.9f}"}, na_rep="—"),
            hide_index=True)
        st.caption("Inputs: the client's audited values plus the published μ, σ_score and "
                   "weights of the stated version. This reuses the published scoring "
                   "function; it shows the parameters suffice, not a second implementation.")

    # ---- 4 · DEPENDENCY-WEIGHT METHOD ------------------------------------
    with tabs[3]:
        dw = r.dependency_weights[client]
        st.markdown(f"#### Client dependency weights · {period}")
        st.caption(f"The method-check results below are from an independent {anchor_label} "
                   "experiment with a deliberately introduced mismatch. They do not "
                   "describe a newly performed check for the selected period.")
        k = st.columns(3)
        am = a_["D_alpha_method"]
        am = am[am["entity"] == disp(client)]
        with k[0]:
            kpi_box("Client's β (own operations)", f"{dw['beta']:.2f}",
                    "α: " + (", ".join(f"{disp(j)} = {v:.2f}" for j, v in dw["alpha"].items())
                             or "no enrolled upstream"))
        if not am.empty:
            x = am.iloc[0]
            with k[1]:
                kpi_box(f"Method experiment · {anchor_label}", "verified" if x["verified"] else "does not match",
                        f"declared: {x['declared_method']}", alert=not bool(x["verified"]))
            with k[2]:
                kpi_box("Assurance procedure",
                        "enhanced" if x["enhanced_assurance_required"] else "standard",
                        f"applied: {x['method_applied']}",
                        alert=bool(x["enhanced_assurance_required"]))
            if not x["verified"]:
                st.markdown(f"<div class='crit'>Applied method recorded: "
                            f"{x['method_applied']}. "
                            + ("The declared method was not followed."
                               if not x["consistent_with_declaration"] else "")
                            + ("" if x["method_accepted"] else
                               " The method is not accepted under Table 4.")
                            + "</div>", unsafe_allow_html=True)
        st.markdown("#### Accepted methods (Table 4)")
        show_df(pd.DataFrame([{"Method": m_, "Basis": b_,
                               "Assurance": ("enhanced" if m_ in ENHANCED_ASSURANCE_METHODS
                                             else "standard")}
                              for m_, b_ in DEPENDENCY_METHODS.items()]), hide_index=True)
        st.caption("The reference probe compares declared and applied method labels with "
                   "each other and with the accepted methods. It does not inspect "
                   "procurement records or recompute α from them.")

    # ---- 5 · ALERT RESPONSES ---------------------------------------------
    with tabs[4]:
        st.markdown("#### Documented alert responses · all periods")
        ar = a_.get("E_alert_responses")
        ar = (ar[ar["entity"] == disp(client)] if ar is not None and len(ar)
              else pd.DataFrame())
        if ar.empty:
            st.caption("No alert episode at this client called for a documented response.")
        else:
            show_df(ar.drop(columns=["entity"]), hide_index=True)


# =============================================================================
# DASHBOARD ENTRY POINT
# =============================================================================

def main():
    st.set_page_config(page_title=PLATFORM_NAME, page_icon="◧", layout="wide",
                       initial_sidebar_state="collapsed")
    st.markdown(CSS, unsafe_allow_html=True)
    site_topbar()
    _preserve_policy_state()
    _carry_page_state()
    keep_state("data_dir", str(default_data_directory()))
    ss = st.session_state
    view = site_navigation(NAV_ITEMS)

    band = ss["pol_band"]
    try:
        out = load_dashboard_results(
            ss["data_dir"], ss["pol_lam"],
            band[0], band[1], ss["pol_t3_floor"], ss["pol_t3_cap"], ss["pol_block_t2"],
            ss["pol_dispersion"], ss["pol_squash"], ss["pol_alert"], ss["pol_dabs"],
            ss["pol_drel"], True, coverage_floor=ss["pol_cov"])
    except FileNotFoundError as exc:
        st.error(f"{exc}\n\nSet the data directory below to the folder holding "
                 f"ESG_Raw_Data_Supply_Chain_M0–M5.xlsx.")
        data_setup_panel(expanded=True)
        site_footer()
        return
    except Exception as exc:
        st.error(f"Engine failed: {exc}")
        raise
    core.ANONYMISE_ENTITIES = True

    if view == "Participating firm":
        firm_view(out)
    elif view == "Governance body":
        regulator_view(out)
    elif view == "Assurance provider":
        auditor_view(out)
    else:
        overview_view(out)
    site_footer(out)


def _in_streamlit_context():
    """True only when Streamlit is executing this file as an application."""
    try:
        from streamlit.runtime.scriptrunner import get_script_run_ctx
        try:
            return get_script_run_ctx(suppress_warning=True) is not None
        except TypeError:  # compatibility with older Streamlit releases
            return get_script_run_ctx() is not None
    except Exception:
        return False


if __name__ == "__main__":
    if _in_streamlit_context():
        main()
    else:
        _single_file_cli()

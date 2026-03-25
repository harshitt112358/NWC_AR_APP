# ============================================================
# AR METRICS APP (Standalone)
# Updated with:
# 1) Comment box for every metric block
# 2) Second Excel export with components, component values, metric value, and comment
# ============================================================

import re
import io
from dataclasses import dataclass
from typing import Any, Dict, List, Optional, Tuple
from collections import defaultdict

import pandas as pd
import streamlit as st

st.set_page_config(page_title="AR Metrics App", layout="wide")
st.title("NWC — AR Metrics (Auto-calc)")

# ============================================================
# Metric Model
# ============================================================

@dataclass
class MetricSpec:
    lever: str
    kpi: str
    components: List[str]
    calc_type: str  # identity | ratio_pct | multi_ratio_pct | days_ratio
    numerator: Optional[str] = None
    denominator: Optional[str] = None
    numerators: Optional[List[str]] = None
    multi_denominator: Optional[str] = None
    notes: Optional[str] = None
    metric_name: Optional[str] = None
    region: Optional[str] = None


AR_METRICS: List[MetricSpec] = [

    # ============================================================
    # 1) COLLECTION OF OVERDUE RECEIVABLES
    # ============================================================

    # ---- Avg AR as % of Sales ----
    MetricSpec("Collection of overdue receivables", "Avg AR as % of Sales (Industry/ Region): Overall",
               ["AR (Overall)", "Sales (Overall)"], "ratio_pct", "AR (Overall)", "Sales (Overall)",
               metric_name="Avg AR as % of Sales", region="Overall"),
    MetricSpec("Collection of overdue receivables", "Avg AR as % of Sales (Industry/ Region): Americas",
               ["AR (Americas)", "Sales (Americas)"], "ratio_pct", "AR (Americas)", "Sales (Americas)",
               metric_name="Avg AR as % of Sales", region="Americas"),
    MetricSpec("Collection of overdue receivables", "Avg AR as % of Sales (Industry/ Region): APAC",
               ["AR (APAC)", "Sales (APAC)"], "ratio_pct", "AR (APAC)", "Sales (APAC)",
               metric_name="Avg AR as % of Sales", region="APAC"),
    MetricSpec("Collection of overdue receivables", "Avg AR as % of Sales (Industry/ Region): EMEA",
               ["AR (EMEA)", "Sales (EMEA)"], "ratio_pct", "AR (EMEA)", "Sales (EMEA)",
               metric_name="Avg AR as % of Sales", region="EMEA"),

    # ---- Avg Overdues as % of AR ----
    MetricSpec("Collection of overdue receivables", "Avg Overdues as % of AR (Industry/ Region): Overall",
               ["Overdues (Overall)", "AR (Overall)"], "ratio_pct", "Overdues (Overall)", "AR (Overall)",
               metric_name="Avg Overdues as % of AR", region="Overall"),
    MetricSpec("Collection of overdue receivables", "Avg Overdues as % of AR (Industry/ Region): Americas",
               ["Overdues (Americas)", "AR (Americas)"], "ratio_pct", "Overdues (Americas)", "AR (Americas)",
               metric_name="Avg Overdues as % of AR", region="Americas"),
    MetricSpec("Collection of overdue receivables", "Avg Overdues as % of AR (Industry/ Region): APAC",
               ["Overdues (APAC)", "AR (APAC)"], "ratio_pct", "Overdues (APAC)", "AR (APAC)",
               metric_name="Avg Overdues as % of AR", region="APAC"),
    MetricSpec("Collection of overdue receivables", "Avg Overdues as % of AR (Industry/ Region): EMEA",
               ["Overdues (EMEA)", "AR (EMEA)"], "ratio_pct", "Overdues (EMEA)", "AR (EMEA)",
               metric_name="Avg Overdues as % of AR", region="EMEA"),

    # ---- Aging profile of overdues buckets — Overall ----
    MetricSpec("Collection of overdue receivables",
               "Aging profile of overdues buckets (Industry/ Region): Overall — 1-15 Days Overdues (Overall) (%)",
               ["1-15 Days Overdues (Overall)", "Overdues (Overall)"], "ratio_pct",
               "1-15 Days Overdues (Overall)", "Overdues (Overall)",
               metric_name="Aging profile — 1-15 Days Overdues (%)", region="Overall"),
    MetricSpec("Collection of overdue receivables",
               "Aging profile of overdues buckets (Industry/ Region): Overall — 16-30 Days Overdues (Overall) (%)",
               ["16-30 Days Overdues (Overall)", "Overdues (Overall)"], "ratio_pct",
               "16-30 Days Overdues (Overall)", "Overdues (Overall)",
               metric_name="Aging profile — 16-30 Days Overdues (%)", region="Overall"),
    MetricSpec("Collection of overdue receivables",
               "Aging profile of overdues buckets (Industry/ Region): Overall — 31-45 Days Overdues (Overall) (%)",
               ["31-45 Days Overdues (Overall)", "Overdues (Overall)"], "ratio_pct",
               "31-45 Days Overdues (Overall)", "Overdues (Overall)",
               metric_name="Aging profile — 31-45 Days Overdues (%)", region="Overall"),
    MetricSpec("Collection of overdue receivables",
               "Aging profile of overdues buckets (Industry/ Region): Overall — 46-60 Days Overdues (Overall) (%)",
               ["46-60 Days Overdues (Overall)", "Overdues (Overall)"], "ratio_pct",
               "46-60 Days Overdues (Overall)", "Overdues (Overall)",
               metric_name="Aging profile — 46-60 Days Overdues (%)", region="Overall"),
    MetricSpec("Collection of overdue receivables",
               "Aging profile of overdues buckets (Industry/ Region): Overall — 61-90 Days Overdues (Overall) (%)",
               ["61-90 Days Overdues (Overall)", "Overdues (Overall)"], "ratio_pct",
               "61-90 Days Overdues (Overall)", "Overdues (Overall)",
               metric_name="Aging profile — 61-90 Days Overdues (%)", region="Overall"),
    MetricSpec("Collection of overdue receivables",
               "Aging profile of overdues buckets (Industry/ Region): Overall — 91-120 Days Overdues (Overall) (%)",
               ["91-120 Days Overdues (Overall)", "Overdues (Overall)"], "ratio_pct",
               "91-120 Days Overdues (Overall)", "Overdues (Overall)",
               metric_name="Aging profile — 91-120 Days Overdues (%)", region="Overall"),
    MetricSpec("Collection of overdue receivables",
               "Aging profile of overdues buckets (Industry/ Region): Overall — 120+ Days Overdues (Overall) (%)",
               ["120+ Days Overdues (Overall)", "Overdues (Overall)"], "ratio_pct",
               "120+ Days Overdues (Overall)", "Overdues (Overall)",
               metric_name="Aging profile — 120+ Days Overdues (%)", region="Overall"),

    # Americas aging buckets
    MetricSpec("Collection of overdue receivables",
               "Aging profile of overdues buckets (Industry/ Region): Americas — 1-15 Days Overdues (Americas) (%)",
               ["1-15 Days Overdues (Americas)", "Overdues (Americas)"], "ratio_pct",
               "1-15 Days Overdues (Americas)", "Overdues (Americas)",
               metric_name="Aging profile — 1-15 Days Overdues (%)", region="Americas"),
    MetricSpec("Collection of overdue receivables",
               "Aging profile of overdues buckets (Industry/ Region): Americas — 16-30 Days Overdues (Americas) (%)",
               ["16-30 Days Overdues (Americas)", "Overdues (Americas)"], "ratio_pct",
               "16-30 Days Overdues (Americas)", "Overdues (Americas)",
               metric_name="Aging profile — 16-30 Days Overdues (%)", region="Americas"),
    MetricSpec("Collection of overdue receivables",
               "Aging profile of overdues buckets (Industry/ Region): Americas — 31-45 Days Overdues (Americas) (%)",
               ["31-45 Days Overdues (Americas)", "Overdues (Americas)"], "ratio_pct",
               "31-45 Days Overdues (Americas)", "Overdues (Americas)",
               metric_name="Aging profile — 31-45 Days Overdues (%)", region="Americas"),
    MetricSpec("Collection of overdue receivables",
               "Aging profile of overdues buckets (Industry/ Region): Americas — 46-60 Days Overdues (Americas) (%)",
               ["46-60 Days Overdues (Americas)", "Overdues (Americas)"], "ratio_pct",
               "46-60 Days Overdues (Americas)", "Overdues (Americas)",
               metric_name="Aging profile — 46-60 Days Overdues (%)", region="Americas"),
    MetricSpec("Collection of overdue receivables",
               "Aging profile of overdues buckets (Industry/ Region): Americas — 61-90 Days Overdues (Americas) (%)",
               ["61-90 Days Overdues (Americas)", "Overdues (Americas)"], "ratio_pct",
               "61-90 Days Overdues (Americas)", "Overdues (Americas)",
               metric_name="Aging profile — 61-90 Days Overdues (%)", region="Americas"),
    MetricSpec("Collection of overdue receivables",
               "Aging profile of overdues buckets (Industry/ Region): Americas — 91-120 Days Overdues (Americas) (%)",
               ["91-120 Days Overdues (Americas)", "Overdues (Americas)"], "ratio_pct",
               "91-120 Days Overdues (Americas)", "Overdues (Americas)",
               metric_name="Aging profile — 91-120 Days Overdues (%)", region="Americas"),
    MetricSpec("Collection of overdue receivables",
               "Aging profile of overdues buckets (Industry/ Region): Americas — 120+ Days Overdues (Americas) (%)",
               ["120+ Days Overdues (Americas)", "Overdues (Americas)"], "ratio_pct",
               "120+ Days Overdues (Americas)", "Overdues (Americas)",
               metric_name="Aging profile — 120+ Days Overdues (%)", region="Americas"),

    # APAC aging buckets
    MetricSpec("Collection of overdue receivables",
               "Aging profile of overdues buckets (Industry/ Region): APAC — 1-15 Days Overdues (APAC) (%)",
               ["1-15 Days Overdues (APAC)", "Overdues (APAC)"], "ratio_pct",
               "1-15 Days Overdues (APAC)", "Overdues (APAC)",
               metric_name="Aging profile — 1-15 Days Overdues (%)", region="APAC"),
    MetricSpec("Collection of overdue receivables",
               "Aging profile of overdues buckets (Industry/ Region): APAC — 16-30 Days Overdues (APAC) (%)",
               ["16-30 Days Overdues (APAC)", "Overdues (APAC)"], "ratio_pct",
               "16-30 Days Overdues (APAC)", "Overdues (APAC)",
               metric_name="Aging profile — 16-30 Days Overdues (%)", region="APAC"),
    MetricSpec("Collection of overdue receivables",
               "Aging profile of overdues buckets (Industry/ Region): APAC — 31-45 Days Overdues (APAC) (%)",
               ["31-45 Days Overdues (APAC)", "Overdues (APAC)"], "ratio_pct",
               "31-45 Days Overdues (APAC)", "Overdues (APAC)",
               metric_name="Aging profile — 31-45 Days Overdues (%)", region="APAC"),
    MetricSpec("Collection of overdue receivables",
               "Aging profile of overdues buckets (Industry/ Region): APAC — 46-60 Days Overdues (APAC) (%)",
               ["46-60 Days Overdues (APAC)", "Overdues (APAC)"], "ratio_pct",
               "46-60 Days Overdues (APAC)", "Overdues (APAC)",
               metric_name="Aging profile — 46-60 Days Overdues (%)", region="APAC"),
    MetricSpec("Collection of overdue receivables",
               "Aging profile of overdues buckets (Industry/ Region): APAC — 61-90 Days Overdues (APAC) (%)",
               ["61-90 Days Overdues (APAC)", "Overdues (APAC)"], "ratio_pct",
               "61-90 Days Overdues (APAC)", "Overdues (APAC)",
               metric_name="Aging profile — 61-90 Days Overdues (%)", region="APAC"),
    MetricSpec("Collection of overdue receivables",
               "Aging profile of overdues buckets (Industry/ Region): APAC — 91-120 Days Overdues (APAC) (%)",
               ["91-120 Days Overdues (APAC)", "Overdues (APAC)"], "ratio_pct",
               "91-120 Days Overdues (APAC)", "Overdues (APAC)",
               metric_name="Aging profile — 91-120 Days Overdues (%)", region="APAC"),
    MetricSpec("Collection of overdue receivables",
               "Aging profile of overdues buckets (Industry/ Region): APAC — 120+ Days Overdues (APAC) (%)",
               ["120+ Days Overdues (APAC)", "Overdues (APAC)"], "ratio_pct",
               "120+ Days Overdues (APAC)", "Overdues (APAC)",
               metric_name="Aging profile — 120+ Days Overdues (%)", region="APAC"),

    # EMEA aging buckets
    MetricSpec("Collection of overdue receivables",
               "Aging profile of overdues buckets (Industry/ Region): EMEA — 1-15 Days Overdues (EMEA) (%)",
               ["1-15 Days Overdues (EMEA)", "Overdues (EMEA)"], "ratio_pct",
               "1-15 Days Overdues (EMEA)", "Overdues (EMEA)",
               metric_name="Aging profile — 1-15 Days Overdues (%)", region="EMEA"),
    MetricSpec("Collection of overdue receivables",
               "Aging profile of overdues buckets (Industry/ Region): EMEA — 16-30 Days Overdues (EMEA) (%)",
               ["16-30 Days Overdues (EMEA)", "Overdues (EMEA)"], "ratio_pct",
               "16-30 Days Overdues (EMEA)", "Overdues (EMEA)",
               metric_name="Aging profile — 16-30 Days Overdues (%)", region="EMEA"),
    MetricSpec("Collection of overdue receivables",
               "Aging profile of overdues buckets (Industry/ Region): EMEA — 31-45 Days Overdues (EMEA) (%)",
               ["31-45 Days Overdues (EMEA)", "Overdues (EMEA)"], "ratio_pct",
               "31-45 Days Overdues (EMEA)", "Overdues (EMEA)",
               metric_name="Aging profile — 31-45 Days Overdues (%)", region="EMEA"),
    MetricSpec("Collection of overdue receivables",
               "Aging profile of overdues buckets (Industry/ Region): EMEA — 46-60 Days Overdues (EMEA) (%)",
               ["46-60 Days Overdues (EMEA)", "Overdues (EMEA)"], "ratio_pct",
               "46-60 Days Overdues (EMEA)", "Overdues (EMEA)",
               metric_name="Aging profile — 46-60 Days Overdues (%)", region="EMEA"),
    MetricSpec("Collection of overdue receivables",
               "Aging profile of overdues buckets (Industry/ Region): EMEA — 61-90 Days Overdues (EMEA) (%)",
               ["61-90 Days Overdues (EMEA)", "Overdues (EMEA)"], "ratio_pct",
               "61-90 Days Overdues (EMEA)", "Overdues (EMEA)",
               metric_name="Aging profile — 61-90 Days Overdues (%)", region="EMEA"),
    MetricSpec("Collection of overdue receivables",
               "Aging profile of overdues buckets (Industry/ Region): EMEA — 91-120 Days Overdues (EMEA) (%)",
               ["91-120 Days Overdues (EMEA)", "Overdues (EMEA)"], "ratio_pct",
               "91-120 Days Overdues (EMEA)", "Overdues (EMEA)",
               metric_name="Aging profile — 91-120 Days Overdues (%)", region="EMEA"),
    MetricSpec("Collection of overdue receivables",
               "Aging profile of overdues buckets (Industry/ Region): EMEA — 120+ Days Overdues (EMEA) (%)",
               ["120+ Days Overdues (EMEA)", "Overdues (EMEA)"], "ratio_pct",
               "120+ Days Overdues (EMEA)", "Overdues (EMEA)",
               metric_name="Aging profile — 120+ Days Overdues (%)", region="EMEA"),

    # ---- DSO (Days) ----
    MetricSpec("Collection of overdue receivables", "DSO (Industry/ Region): Overall",
               ["AR (Overall)", "Sales (Overall)"], "days_ratio", "AR (Overall)", "Sales (Overall)",
               metric_name="DSO (Days)", region="Overall"),
    MetricSpec("Collection of overdue receivables", "DSO (Industry/ Region): Americas",
               ["AR (Americas)", "Sales (Americas)"], "days_ratio", "AR (Americas)", "Sales (Americas)",
               metric_name="DSO (Days)", region="Americas"),
    MetricSpec("Collection of overdue receivables", "DSO (Industry/ Region): APAC",
               ["AR (APAC)", "Sales (APAC)"], "days_ratio", "AR (APAC)", "Sales (APAC)",
               metric_name="DSO (Days)", region="APAC"),
    MetricSpec("Collection of overdue receivables", "DSO (Industry/ Region): EMEA",
               ["AR (EMEA)", "Sales (EMEA)"], "days_ratio", "AR (EMEA)", "Sales (EMEA)",
               metric_name="DSO (Days)", region="EMEA"),

    # ---- Cash benefit as % of avg overdues ----
    MetricSpec("Collection of overdue receivables", "Cash benefit as % of avg overdues (Industry/ Region): Overall",
               ["Cash Benefit (Overall)", "Overdues (Overall)"], "ratio_pct", "Cash Benefit (Overall)", "Overdues (Overall)",
               metric_name="Cash benefit as % of avg overdues", region="Overall"),
    MetricSpec("Collection of overdue receivables", "Cash benefit as % of avg overdues (Industry/ Region): Americas",
               ["Cash Benefit (Americas)", "Overdues (Americas)"], "ratio_pct", "Cash Benefit (Americas)", "Overdues (Americas)",
               metric_name="Cash benefit as % of avg overdues", region="Americas"),
    MetricSpec("Collection of overdue receivables", "Cash benefit as % of avg overdues (Industry/ Region): APAC",
               ["Cash Benefit (APAC)", "Overdues (APAC)"], "ratio_pct", "Cash Benefit (APAC)", "Overdues (APAC)",
               metric_name="Cash benefit as % of avg overdues", region="APAC"),
    MetricSpec("Collection of overdue receivables", "Cash benefit as % of avg overdues (Industry/ Region): EMEA",
               ["Cash Benefit (EMEA)", "Overdues (EMEA)"], "ratio_pct", "Cash Benefit (EMEA)", "Overdues (EMEA)",
               metric_name="Cash benefit as % of avg overdues", region="EMEA"),

    # ---- Cash Benefit from eliminating overdues as a % of sales ----
    MetricSpec("Collection of overdue receivables",
               "Cash Benefit from eliminating overdues as a % of sales (Industry / Region): Overall",
               ["Cash Benefit (Overall)", "Sales (Overall)"], "ratio_pct", "Cash Benefit (Overall)", "Sales (Overall)",
               metric_name="Cash Benefit from eliminating overdues as a % of sales", region="Overall"),
    MetricSpec("Collection of overdue receivables",
               "Cash Benefit from eliminating overdues as a % of sales (Industry / Region): Americas",
               ["Cash Benefit (Americas)", "Sales (Americas)"], "ratio_pct", "Cash Benefit (Americas)", "Sales (Americas)",
               metric_name="Cash Benefit from eliminating overdues as a % of sales", region="Americas"),
    MetricSpec("Collection of overdue receivables",
               "Cash Benefit from eliminating overdues as a % of sales (Industry / Region): APAC",
               ["Cash Benefit (APAC)", "Sales (APAC)"], "ratio_pct", "Cash Benefit (APAC)", "Sales (APAC)",
               metric_name="Cash Benefit from eliminating overdues as a % of sales", region="APAC"),
    MetricSpec("Collection of overdue receivables",
               "Cash Benefit from eliminating overdues as a % of sales (Industry / Region): EMEA",
               ["Cash Benefit (EMEA)", "Sales (EMEA)"], "ratio_pct", "Cash Benefit (EMEA)", "Sales (EMEA)",
               metric_name="Cash Benefit from eliminating overdues as a % of sales", region="EMEA"),

    # ============================================================
    # 2) ELIMINATE DELAYED PAYMENTS
    # ============================================================

    # ---- Delayed Payments as % Sales ----
    MetricSpec("Eliminate delayed Payments", "Delayed Payments as % Sales (Industry/ Region): Overall",
               ["Delayed Payments (Overall)", "Sales (Overall)"], "ratio_pct", "Delayed Payments (Overall)", "Sales (Overall)",
               metric_name="Delayed Payments as % Sales", region="Overall"),
    MetricSpec("Eliminate delayed Payments", "Delayed Payments as % Sales (Industry/ Region): Americas",
               ["Delayed Payments (Americas)", "Sales (Americas)"], "ratio_pct", "Delayed Payments (Americas)", "Sales (Americas)",
               metric_name="Delayed Payments as % Sales", region="Americas"),
    MetricSpec("Eliminate delayed Payments", "Delayed Payments as % Sales (Industry/ Region): APAC",
               ["Delayed Payments (APAC)", "Sales (APAC)"], "ratio_pct", "Delayed Payments (APAC)", "Sales (APAC)",
               metric_name="Delayed Payments as % Sales", region="APAC"),
    MetricSpec("Eliminate delayed Payments", "Delayed Payments as % Sales (Industry/ Region): EMEA",
               ["Delayed Payments (EMEA)", "Sales (EMEA)"], "ratio_pct", "Delayed Payments (EMEA)", "Sales (EMEA)",
               metric_name="Delayed Payments as % Sales", region="EMEA"),

    # ---- Profile of delayed payments by # days delayed — Overall ----
    MetricSpec("Eliminate delayed Payments",
               "Profile of dealyed payments by # days delayed (Industry/ Region): Overall — 1-5 Days Delayed Payments (Overall) (%)",
               ["1-5 Days Delayed Payments (Overall)", "Sales (Overall)"], "ratio_pct",
               "1-5 Days Delayed Payments (Overall)", "Sales (Overall)",
               metric_name="Profile of delayed payments — 1-5 Days (%)", region="Overall"),
    MetricSpec("Eliminate delayed Payments",
               "Profile of dealyed payments by # days delayed (Industry/ Region): Overall — 6-15 Days Delayed Payments (Overall) (%)",
               ["6-15 Days Delayed Payments (Overall)", "Sales (Overall)"], "ratio_pct",
               "6-15 Days Delayed Payments (Overall)", "Sales (Overall)",
               metric_name="Profile of delayed payments — 6-15 Days (%)", region="Overall"),
    MetricSpec("Eliminate delayed Payments",
               "Profile of dealyed payments by # days delayed (Industry/ Region): Overall — 16-30 Days Delayed Payments (Overall) (%)",
               ["16-30 Days Delayed Payments (Overall)", "Sales (Overall)"], "ratio_pct",
               "16-30 Days Delayed Payments (Overall)", "Sales (Overall)",
               metric_name="Profile of delayed payments — 16-30 Days (%)", region="Overall"),
    MetricSpec("Eliminate delayed Payments",
               "Profile of dealyed payments by # days delayed (Industry/ Region): Overall — 31-60 Days Delayed Payments (Overall) (%)",
               ["31-60 Days Delayed Payments (Overall)", "Sales (Overall)"], "ratio_pct",
               "31-60 Days Delayed Payments (Overall)", "Sales (Overall)",
               metric_name="Profile of delayed payments — 31-60 Days (%)", region="Overall"),
    MetricSpec("Eliminate delayed Payments",
               "Profile of dealyed payments by # days delayed (Industry/ Region): Overall — 61-90 Days Delayed Payments (Overall) (%)",
               ["61-90 Days Delayed Payments (Overall)", "Sales (Overall)"], "ratio_pct",
               "61-90 Days Delayed Payments (Overall)", "Sales (Overall)",
               metric_name="Profile of delayed payments — 61-90 Days (%)", region="Overall"),
    MetricSpec("Eliminate delayed Payments",
               "Profile of dealyed payments by # days delayed (Industry/ Region): Overall — 90+ Days Delayed Payments (Overall) (%)",
               ["90+ Days Delayed Payments (Overall)", "Sales (Overall)"], "ratio_pct",
               "90+ Days Delayed Payments (Overall)", "Sales (Overall)",
               metric_name="Profile of delayed payments — 90+ Days (%)", region="Overall"),

    # Americas
    MetricSpec("Eliminate delayed Payments",
               "Profile of dealyed payments by # days delayed (Industry/ Region): Americas — 1-5 Days Delayed Payments (Americas) (%)",
               ["1-5 Days Delayed Payments (Americas)", "Sales (Americas)"], "ratio_pct",
               "1-5 Days Delayed Payments (Americas)", "Sales (Americas)",
               metric_name="Profile of delayed payments — 1-5 Days (%)", region="Americas"),
    MetricSpec("Eliminate delayed Payments",
               "Profile of dealyed payments by # days delayed (Industry/ Region): Americas — 6-15 Days Delayed Payments (Americas) (%)",
               ["6-15 Days Delayed Payments (Americas)", "Sales (Americas)"], "ratio_pct",
               "6-15 Days Delayed Payments (Americas)", "Sales (Americas)",
               metric_name="Profile of delayed payments — 6-15 Days (%)", region="Americas"),
    MetricSpec("Eliminate delayed Payments",
               "Profile of dealyed payments by # days delayed (Industry/ Region): Americas — 16-30 Days Delayed Payments (Americas) (%)",
               ["16-30 Days Delayed Payments (Americas)", "Sales (Americas)"], "ratio_pct",
               "16-30 Days Delayed Payments (Americas)", "Sales (Americas)",
               metric_name="Profile of delayed payments — 16-30 Days (%)", region="Americas"),
    MetricSpec("Eliminate delayed Payments",
               "Profile of dealyed payments by # days delayed (Industry/ Region): Americas — 31-60 Days Delayed Payments (Americas) (%)",
               ["31-60 Days Delayed Payments (Americas)", "Sales (Americas)"], "ratio_pct",
               "31-60 Days Delayed Payments (Americas)", "Sales (Americas)",
               metric_name="Profile of delayed payments — 31-60 Days (%)", region="Americas"),
    MetricSpec("Eliminate delayed Payments",
               "Profile of dealyed payments by # days delayed (Industry/ Region): Americas — 61-90 Days Delayed Payments (Americas) (%)",
               ["61-90 Days Delayed Payments (Americas)", "Sales (Americas)"], "ratio_pct",
               "61-90 Days Delayed Payments (Americas)", "Sales (Americas)",
               metric_name="Profile of delayed payments — 61-90 Days (%)", region="Americas"),
    MetricSpec("Eliminate delayed Payments",
               "Profile of dealyed payments by # days delayed (Industry/ Region): Americas — 90+ Days Delayed Payments (Americas) (%)",
               ["90+ Days Delayed Payments (Americas)", "Sales (Americas)"], "ratio_pct",
               "90+ Days Delayed Payments (Americas)", "Sales (Americas)",
               metric_name="Profile of delayed payments — 90+ Days (%)", region="Americas"),

    # APAC
    MetricSpec("Eliminate delayed Payments",
               "Profile of dealyed payments by # days delayed (Industry/ Region): APAC — 1-5 Days Delayed Payments (APAC) (%)",
               ["1-5 Days Delayed Payments (APAC)", "Sales (APAC)"], "ratio_pct",
               "1-5 Days Delayed Payments (APAC)", "Sales (APAC)",
               metric_name="Profile of delayed payments — 1-5 Days (%)", region="APAC"),
    MetricSpec("Eliminate delayed Payments",
               "Profile of dealyed payments by # days delayed (Industry/ Region): APAC — 6-15 Days Delayed Payments (APAC) (%)",
               ["6-15 Days Delayed Payments (APAC)", "Sales (APAC)"], "ratio_pct",
               "6-15 Days Delayed Payments (APAC)", "Sales (APAC)",
               metric_name="Profile of delayed payments — 6-15 Days (%)", region="APAC"),
    MetricSpec("Eliminate delayed Payments",
               "Profile of dealyed payments by # days delayed (Industry/ Region): APAC — 16-30 Days Delayed Payments (APAC) (%)",
               ["16-30 Days Delayed Payments (APAC)", "Sales (APAC)"], "ratio_pct",
               "16-30 Days Delayed Payments (APAC)", "Sales (APAC)",
               metric_name="Profile of delayed payments — 16-30 Days (%)", region="APAC"),
    MetricSpec("Eliminate delayed Payments",
               "Profile of dealyed payments by # days delayed (Industry/ Region): APAC — 31-60 Days Delayed Payments (APAC) (%)",
               ["31-60 Days Delayed Payments (APAC)", "Sales (APAC)"], "ratio_pct",
               "31-60 Days Delayed Payments (APAC)", "Sales (APAC)",
               metric_name="Profile of delayed payments — 31-60 Days (%)", region="APAC"),
    MetricSpec("Eliminate delayed Payments",
               "Profile of dealyed payments by # days delayed (Industry/ Region): APAC — 61-90 Days Delayed Payments (APAC) (%)",
               ["61-90 Days Delayed Payments (APAC)", "Sales (APAC)"], "ratio_pct",
               "61-90 Days Delayed Payments (APAC)", "Sales (APAC)",
               metric_name="Profile of delayed payments — 61-90 Days (%)", region="APAC"),
    MetricSpec("Eliminate delayed Payments",
               "Profile of dealyed payments by # days delayed (Industry/ Region): APAC — 90+ Days Delayed Payments (APAC) (%)",
               ["90+ Days Delayed Payments (APAC)", "Sales (APAC)"], "ratio_pct",
               "90+ Days Delayed Payments (APAC)", "Sales (APAC)",
               metric_name="Profile of delayed payments — 90+ Days (%)", region="APAC"),

    # EMEA
    MetricSpec("Eliminate delayed Payments",
               "Profile of dealyed payments by # days delayed (Industry/ Region): EMEA — 1-5 Days Delayed Payments (EMEA) (%)",
               ["1-5 Days Delayed Payments (EMEA)", "Sales (EMEA)"], "ratio_pct",
               "1-5 Days Delayed Payments (EMEA)", "Sales (EMEA)",
               metric_name="Profile of delayed payments — 1-5 Days (%)", region="EMEA"),
    MetricSpec("Eliminate delayed Payments",
               "Profile of dealyed payments by # days delayed (Industry/ Region): EMEA — 6-15 Days Delayed Payments (EMEA) (%)",
               ["6-15 Days Delayed Payments (EMEA)", "Sales (EMEA)"], "ratio_pct",
               "6-15 Days Delayed Payments (EMEA)", "Sales (EMEA)",
               metric_name="Profile of delayed payments — 6-15 Days (%)", region="EMEA"),
    MetricSpec("Eliminate delayed Payments",
               "Profile of dealyed payments by # days delayed (Industry/ Region): EMEA — 16-30 Days Delayed Payments (EMEA) (%)",
               ["16-30 Days Delayed Payments (EMEA)", "Sales (EMEA)"], "ratio_pct",
               "16-30 Days Delayed Payments (EMEA)", "Sales (EMEA)",
               metric_name="Profile of delayed payments — 16-30 Days (%)", region="EMEA"),
    MetricSpec("Eliminate delayed Payments",
               "Profile of dealyed payments by # days delayed (Industry/ Region): EMEA — 31-60 Days Delayed Payments (EMEA) (%)",
               ["31-60 Days Delayed Payments (EMEA)", "Sales (EMEA)"], "ratio_pct",
               "31-60 Days Delayed Payments (EMEA)", "Sales (EMEA)",
               metric_name="Profile of delayed payments — 31-60 Days (%)", region="EMEA"),
    MetricSpec("Eliminate delayed Payments",
               "Profile of dealyed payments by # days delayed (Industry/ Region): EMEA — 61-90 Days Delayed Payments (EMEA) (%)",
               ["61-90 Days Delayed Payments (EMEA)", "Sales (EMEA)"], "ratio_pct",
               "61-90 Days Delayed Payments (EMEA)", "Sales (EMEA)",
               metric_name="Profile of delayed payments — 61-90 Days (%)", region="EMEA"),
    MetricSpec("Eliminate delayed Payments",
               "Profile of dealyed payments by # days delayed (Industry/ Region): EMEA — 90+ Days Delayed Payments (EMEA) (%)",
               ["90+ Days Delayed Payments (EMEA)", "Sales (EMEA)"], "ratio_pct",
               "90+ Days Delayed Payments (EMEA)", "Sales (EMEA)",
               metric_name="Profile of delayed payments — 90+ Days (%)", region="EMEA"),

    # ---- Cash benefit as % of Annual delayed Payments ----
    MetricSpec("Eliminate delayed Payments",
               "Cash benefit as % of Annual delayed Payments (Industry/ Region): Overall",
               ["Cash Benefit (Overall)", "Delayed Payments (Overall)"], "ratio_pct",
               "Cash Benefit (Overall)", "Delayed Payments (Overall)",
               metric_name="Cash benefit as % of Annual delayed Payments", region="Overall"),
    MetricSpec("Eliminate delayed Payments",
               "Cash benefit as % of Annual delayed Payments (Industry/ Region): Americas",
               ["Cash Benefit (Americas)", "Delayed Payments (Americas)"], "ratio_pct",
               "Cash Benefit (Americas)", "Delayed Payments (Americas)",
               metric_name="Cash benefit as % of Annual delayed Payments", region="Americas"),
    MetricSpec("Eliminate delayed Payments",
               "Cash benefit as % of Annual delayed Payments (Industry/ Region): APAC",
               ["Cash Benefit (APAC)", "Delayed Payments (APAC)"], "ratio_pct",
               "Cash Benefit (APAC)", "Delayed Payments (APAC)",
               metric_name="Cash benefit as % of Annual delayed Payments", region="APAC"),
    MetricSpec("Eliminate delayed Payments",
               "Cash benefit as % of Annual delayed Payments (Industry/ Region): EMEA",
               ["Cash Benefit (EMEA)", "Delayed Payments (EMEA)"], "ratio_pct",
               "Cash Benefit (EMEA)", "Delayed Payments (EMEA)",
               metric_name="Cash benefit as % of Annual delayed Payments", region="EMEA"),

    # ---- Cash benefit from eliminating delayed payments as % of Sales ----
    MetricSpec("Eliminate delayed Payments",
               "Cash benefit from eliminating delayed payments as % of Sales (Industry/ Region): Overall",
               ["Cash Benefit (Overall)", "Sales (Overall)"], "ratio_pct", "Cash Benefit (Overall)", "Sales (Overall)",
               metric_name="Cash benefit from eliminating delayed payments as % of Sales", region="Overall"),
    MetricSpec("Eliminate delayed Payments",
               "Cash benefit from eliminating delayed payments as % of Sales (Industry/ Region): Americas",
               ["Cash Benefit (Americas)", "Sales (Americas)"], "ratio_pct", "Cash Benefit (Americas)", "Sales (Americas)",
               metric_name="Cash benefit from eliminating delayed payments as % of Sales", region="Americas"),
    MetricSpec("Eliminate delayed Payments",
               "Cash benefit from eliminating delayed payments as % of Sales (Industry/ Region): APAC",
               ["Cash Benefit (APAC)", "Sales (APAC)"], "ratio_pct", "Cash Benefit (APAC)", "Sales (APAC)",
               metric_name="Cash benefit from eliminating delayed payments as % of Sales", region="APAC"),
    MetricSpec("Eliminate delayed Payments",
               "Cash benefit from eliminating delayed payments as % of Sales (Industry/ Region): EMEA",
               ["Cash Benefit (EMEA)", "Sales (EMEA)"], "ratio_pct", "Cash Benefit (EMEA)", "Sales (EMEA)",
               metric_name="Cash benefit from eliminating delayed payments as % of Sales", region="EMEA"),

    # ============================================================
    # 3) SHORTEN PAYMENT TERMS
    # ============================================================

    MetricSpec("Shorten Payment Terms", "Avg PT days (Industry / Region / Customer size1): Overall",
               ["Avg PT Days (Overall)"], "identity", numerator="Avg PT Days (Overall)",
               metric_name="Avg PT days", region="Overall"),
    MetricSpec("Shorten Payment Terms", "Avg PT days (Industry / Region / Customer size1): Americas",
               ["Avg PT Days (Americas)"], "identity", numerator="Avg PT Days (Americas)",
               metric_name="Avg PT days", region="Americas"),
    MetricSpec("Shorten Payment Terms", "Avg PT days (Industry / Region / Customer size1): APAC",
               ["Avg PT Days (APAC)"], "identity", numerator="Avg PT Days (APAC)",
               metric_name="Avg PT days", region="APAC"),
    MetricSpec("Shorten Payment Terms", "Avg PT days (Industry / Region / Customer size1): EMEA",
               ["Avg PT Days (EMEA)"], "identity", numerator="Avg PT Days (EMEA)",
               metric_name="Avg PT days", region="EMEA"),

    # % sales by Payment buckets — Overall
    MetricSpec("Shorten Payment Terms", "% sales by Payment buckets (Industry / Region): Overall — Immediate (%)",
               ["Immediate Sales (Overall)", "Sales (Overall)"], "ratio_pct", "Immediate Sales (Overall)", "Sales (Overall)",
               metric_name="% sales by Payment bucket — Immediate", region="Overall"),
    MetricSpec("Shorten Payment Terms", "% sales by Payment buckets (Industry / Region): Overall — 0-15 Days Sales (Overall) (%)",
               ["0-15 Days Sales (Overall)", "Sales (Overall)"], "ratio_pct", "0-15 Days Sales (Overall)", "Sales (Overall)",
               metric_name="% sales by Payment bucket — 0-15 Days", region="Overall"),
    MetricSpec("Shorten Payment Terms", "% sales by Payment buckets (Industry / Region): Overall — 16-30 Days Sales (Overall) (%)",
               ["16-30 Days Sales (Overall)", "Sales (Overall)"], "ratio_pct", "16-30 Days Sales (Overall)", "Sales (Overall)",
               metric_name="% sales by Payment bucket — 16-30 Days", region="Overall"),
    MetricSpec("Shorten Payment Terms", "% sales by Payment buckets (Industry / Region): Overall — 31-45 Days Sales (Overall) (%)",
               ["31-45 Days Sales (Overall)", "Sales (Overall)"], "ratio_pct", "31-45 Days Sales (Overall)", "Sales (Overall)",
               metric_name="% sales by Payment bucket — 31-45 Days", region="Overall"),
    MetricSpec("Shorten Payment Terms", "% sales by Payment buckets (Industry / Region): Overall — 46-60 Days Sales (Overall) (%)",
               ["46-60 Days Sales (Overall)", "Sales (Overall)"], "ratio_pct", "46-60 Days Sales (Overall)", "Sales (Overall)",
               metric_name="% sales by Payment bucket — 46-60 Days", region="Overall"),
    MetricSpec("Shorten Payment Terms", "% sales by Payment buckets (Industry / Region): Overall — 61-90 Days Sales (Overall) (%)",
               ["61-90 Days Sales (Overall)", "Sales (Overall)"], "ratio_pct", "61-90 Days Sales (Overall)", "Sales (Overall)",
               metric_name="% sales by Payment bucket — 61-90 Days", region="Overall"),
    MetricSpec("Shorten Payment Terms", "% sales by Payment buckets (Industry / Region): Overall — 90+ Days Sales (Overall) (%)",
               ["90+ Days Sales (Overall)", "Sales (Overall)"], "ratio_pct", "90+ Days Sales (Overall)", "Sales (Overall)",
               metric_name="% sales by Payment bucket — 90+ Days", region="Overall"),

    # Americas
    MetricSpec("Shorten Payment Terms", "% sales by Payment buckets (Industry / Region): Americas — Immediate (%)",
               ["Immediate Sales (Americas)", "Sales (Americas)"], "ratio_pct", "Immediate Sales (Americas)", "Sales (Americas)",
               metric_name="% sales by Payment bucket — Immediate", region="Americas"),
    MetricSpec("Shorten Payment Terms", "% sales by Payment buckets (Industry / Region): Americas — 0-15 Days Sales (Americas) (%)",
               ["0-15 Days Sales (Americas)", "Sales (Americas)"], "ratio_pct", "0-15 Days Sales (Americas)", "Sales (Americas)",
               metric_name="% sales by Payment bucket — 0-15 Days", region="Americas"),
    MetricSpec("Shorten Payment Terms", "% sales by Payment buckets (Industry / Region): Americas — 16-30 Days Sales (Americas) (%)",
               ["16-30 Days Sales (Americas)", "Sales (Americas)"], "ratio_pct", "16-30 Days Sales (Americas)", "Sales (Americas)",
               metric_name="% sales by Payment bucket — 16-30 Days", region="Americas"),
    MetricSpec("Shorten Payment Terms", "% sales by Payment buckets (Industry / Region): Americas — 31-45 Days Sales (Americas) (%)",
               ["31-45 Days Sales (Americas)", "Sales (Americas)"], "ratio_pct", "31-45 Days Sales (Americas)", "Sales (Americas)",
               metric_name="% sales by Payment bucket — 31-45 Days", region="Americas"),
    MetricSpec("Shorten Payment Terms", "% sales by Payment buckets (Industry / Region): Americas — 46-60 Days Sales (Americas) (%)",
               ["46-60 Days Sales (Americas)", "Sales (Americas)"], "ratio_pct", "46-60 Days Sales (Americas)", "Sales (Americas)",
               metric_name="% sales by Payment bucket — 46-60 Days", region="Americas"),
    MetricSpec("Shorten Payment Terms", "% sales by Payment buckets (Industry / Region): Americas — 61-90 Days Sales (Americas) (%)",
               ["61-90 Days Sales (Americas)", "Sales (Americas)"], "ratio_pct", "61-90 Days Sales (Americas)", "Sales (Americas)",
               metric_name="% sales by Payment bucket — 61-90 Days", region="Americas"),
    MetricSpec("Shorten Payment Terms", "% sales by Payment buckets (Industry / Region): Americas — 90+ Days Sales (Americas) (%)",
               ["90+ Days Sales (Americas)", "Sales (Americas)"], "ratio_pct", "90+ Days Sales (Americas)", "Sales (Americas)",
               metric_name="% sales by Payment bucket — 90+ Days", region="Americas"),

    # APAC
    MetricSpec("Shorten Payment Terms", "% sales by Payment buckets (Industry / Region): APAC — Immediate (%)",
               ["Immediate Sales (APAC)", "Sales (APAC)"], "ratio_pct", "Immediate Sales (APAC)", "Sales (APAC)",
               metric_name="% sales by Payment bucket — Immediate", region="APAC"),
    MetricSpec("Shorten Payment Terms", "% sales by Payment buckets (Industry / Region): APAC — 0-15 Days Sales (APAC) (%)",
               ["0-15 Days Sales (APAC)", "Sales (APAC)"], "ratio_pct", "0-15 Days Sales (APAC)", "Sales (APAC)",
               metric_name="% sales by Payment bucket — 0-15 Days", region="APAC"),
    MetricSpec("Shorten Payment Terms", "% sales by Payment buckets (Industry / Region): APAC — 16-30 Days Sales (APAC) (%)",
               ["16-30 Days Sales (APAC)", "Sales (APAC)"], "ratio_pct", "16-30 Days Sales (APAC)", "Sales (APAC)",
               metric_name="% sales by Payment bucket — 16-30 Days", region="APAC"),
    MetricSpec("Shorten Payment Terms", "% sales by Payment buckets (Industry / Region): APAC — 31-45 Days Sales (APAC) (%)",
               ["31-45 Days Sales (APAC)", "Sales (APAC)"], "ratio_pct", "31-45 Days Sales (APAC)", "Sales (APAC)",
               metric_name="% sales by Payment bucket — 31-45 Days", region="APAC"),
    MetricSpec("Shorten Payment Terms", "% sales by Payment buckets (Industry / Region): APAC — 46-60 Days Sales (APAC) (%)",
               ["46-60 Days Sales (APAC)", "Sales (APAC)"], "ratio_pct", "46-60 Days Sales (APAC)", "Sales (APAC)",
               metric_name="% sales by Payment bucket — 46-60 Days", region="APAC"),
    MetricSpec("Shorten Payment Terms", "% sales by Payment buckets (Industry / Region): APAC — 61-90 Days Sales (APAC) (%)",
               ["61-90 Days Sales (APAC)", "Sales (APAC)"], "ratio_pct", "61-90 Days Sales (APAC)", "Sales (APAC)",
               metric_name="% sales by Payment bucket — 61-90 Days", region="APAC"),
    MetricSpec("Shorten Payment Terms", "% sales by Payment buckets (Industry / Region): APAC — 90+ Days Sales (APAC) (%)",
               ["90+ Days Sales (APAC)", "Sales (APAC)"], "ratio_pct", "90+ Days Sales (APAC)", "Sales (APAC)",
               metric_name="% sales by Payment bucket — 90+ Days", region="APAC"),

    # EMEA
    MetricSpec("Shorten Payment Terms", "% sales by Payment buckets (Industry / Region): EMEA — Immediate (%)",
               ["Immediate Sales (EMEA)", "Sales (EMEA)"], "ratio_pct", "Immediate Sales (EMEA)", "Sales (EMEA)",
               metric_name="% sales by Payment bucket — Immediate", region="EMEA"),
    MetricSpec("Shorten Payment Terms", "% sales by Payment buckets (Industry / Region): EMEA — 0-15 Days Sales (EMEA) (%)",
               ["0-15 Days Sales (EMEA)", "Sales (EMEA)"], "ratio_pct", "0-15 Days Sales (EMEA)", "Sales (EMEA)",
               metric_name="% sales by Payment bucket — 0-15 Days", region="EMEA"),
    MetricSpec("Shorten Payment Terms", "% sales by Payment buckets (Industry / Region): EMEA — 16-30 Days Sales (EMEA) (%)",
               ["16-30 Days Sales (EMEA)", "Sales (EMEA)"], "ratio_pct", "16-30 Days Sales (EMEA)", "Sales (EMEA)",
               metric_name="% sales by Payment bucket — 16-30 Days", region="EMEA"),
    MetricSpec("Shorten Payment Terms", "% sales by Payment buckets (Industry / Region): EMEA — 31-45 Days Sales (EMEA) (%)",
               ["31-45 Days Sales (EMEA)", "Sales (EMEA)"], "ratio_pct", "31-45 Days Sales (EMEA)", "Sales (EMEA)",
               metric_name="% sales by Payment bucket — 31-45 Days", region="EMEA"),
    MetricSpec("Shorten Payment Terms", "% sales by Payment buckets (Industry / Region): EMEA — 46-60 Days Sales (EMEA) (%)",
               ["46-60 Days Sales (EMEA)", "Sales (EMEA)"], "ratio_pct", "46-60 Days Sales (EMEA)", "Sales (EMEA)",
               metric_name="% sales by Payment bucket — 46-60 Days", region="EMEA"),
    MetricSpec("Shorten Payment Terms", "% sales by Payment buckets (Industry / Region): EMEA — 61-90 Days Sales (EMEA) (%)",
               ["61-90 Days Sales (EMEA)", "Sales (EMEA)"], "ratio_pct", "61-90 Days Sales (EMEA)", "Sales (EMEA)",
               metric_name="% sales by Payment bucket — 61-90 Days", region="EMEA"),
    MetricSpec("Shorten Payment Terms", "% sales by Payment buckets (Industry / Region): EMEA — 90+ Days Sales (EMEA) (%)",
               ["90+ Days Sales (EMEA)", "Sales (EMEA)"], "ratio_pct", "90+ Days Sales (EMEA)", "Sales (EMEA)",
               metric_name="% sales by Payment bucket — 90+ Days", region="EMEA"),

    # ---- Avg PT days top X Vendor / others (A/B/C split) ----
    MetricSpec("Shorten Payment Terms", "Avg PT days top X Vendor / others (A: Top 80%) | Overall",
               ["Avg PT Days Top Vendors A (Overall)"], "identity", numerator="Avg PT Days Top Vendors A (Overall)",
               metric_name="Avg PT days top X Vendor / others (A: Top 80%)", region="Overall"),
    MetricSpec("Shorten Payment Terms", "Avg PT days top X Vendor / others (B: Next 15%) | Overall",
               ["Avg PT Days Top Vendors B (Overall)"], "identity", numerator="Avg PT Days Top Vendors B (Overall)",
               metric_name="Avg PT days top X Vendor / others (B: Next 15%)", region="Overall"),
    MetricSpec("Shorten Payment Terms", "Avg PT days top X Vendor / others (C: Last 5%) | Overall",
               ["Avg PT Days Top Vendors C (Overall)"], "identity", numerator="Avg PT Days Top Vendors C (Overall)",
               metric_name="Avg PT days top X Vendor / others (C: Last 5%)", region="Overall"),

    MetricSpec("Shorten Payment Terms", "Avg PT days top X Vendor / others (A: Top 80%) | Americas",
               ["Avg PT Days Top Vendors A (Americas)"], "identity", numerator="Avg PT Days Top Vendors A (Americas)",
               metric_name="Avg PT days top X Vendor / others (A: Top 80%)", region="Americas"),
    MetricSpec("Shorten Payment Terms", "Avg PT days top X Vendor / others (B: Next 15%) | Americas",
               ["Avg PT Days Top Vendors B (Americas)"], "identity", numerator="Avg PT Days Top Vendors B (Americas)",
               metric_name="Avg PT days top X Vendor / others (B: Next 15%)", region="Americas"),
    MetricSpec("Shorten Payment Terms", "Avg PT days top X Vendor / others (C: Last 5%) | Americas",
               ["Avg PT Days Top Vendors C (Americas)"], "identity", numerator="Avg PT Days Top Vendors C (Americas)",
               metric_name="Avg PT days top X Vendor / others (C: Last 5%)", region="Americas"),

    MetricSpec("Shorten Payment Terms", "Avg PT days top X Vendor / others (A: Top 80%) | APAC",
               ["Avg PT Days Top Vendors A (APAC)"], "identity", numerator="Avg PT Days Top Vendors A (APAC)",
               metric_name="Avg PT days top X Vendor / others (A: Top 80%)", region="APAC"),
    MetricSpec("Shorten Payment Terms", "Avg PT days top X Vendor / others (B: Next 15%) | APAC",
               ["Avg PT Days Top Vendors B (APAC)"], "identity", numerator="Avg PT Days Top Vendors B (APAC)",
               metric_name="Avg PT days top X Vendor / others (B: Next 15%)", region="APAC"),
    MetricSpec("Shorten Payment Terms", "Avg PT days top X Vendor / others (C: Last 5%) | APAC",
               ["Avg PT Days Top Vendors C (APAC)"], "identity", numerator="Avg PT Days Top Vendors C (APAC)",
               metric_name="Avg PT days top X Vendor / others (C: Last 5%)", region="APAC"),

    MetricSpec("Shorten Payment Terms", "Avg PT days top X Vendor / others (A: Top 80%) | EMEA",
               ["Avg PT Days Top Vendors A (EMEA)"], "identity", numerator="Avg PT Days Top Vendors A (EMEA)",
               metric_name="Avg PT days top X Vendor / others (A: Top 80%)", region="EMEA"),
    MetricSpec("Shorten Payment Terms", "Avg PT days top X Vendor / others (B: Next 15%) | EMEA",
               ["Avg PT Days Top Vendors B (EMEA)"], "identity", numerator="Avg PT Days Top Vendors B (EMEA)",
               metric_name="Avg PT days top X Vendor / others (B: Next 15%)", region="EMEA"),
    MetricSpec("Shorten Payment Terms", "Avg PT days top X Vendor / others (C: Last 5%) | EMEA",
               ["Avg PT Days Top Vendors C (EMEA)"], "identity", numerator="Avg PT Days Top Vendors C (EMEA)",
               metric_name="Avg PT days top X Vendor / others (C: Last 5%)", region="EMEA"),

    # ---- Avg PT days top X customers ----
    MetricSpec("Shorten Payment Terms", "Avg PT days top 10 customers / others (Industry / Region): Overall",
               ["Avg PT Days Top 10 (Overall)"], "identity", numerator="Avg PT Days Top 10 (Overall)",
               metric_name="Avg PT days top 10 customers", region="Overall"),
    MetricSpec("Shorten Payment Terms", "Avg PT days top 15 customers / others (Industry / Region): Overall",
               ["Avg PT Days Top 15 (Overall)"], "identity", numerator="Avg PT Days Top 15 (Overall)",
               metric_name="Avg PT days top 15 customers", region="Overall"),
    MetricSpec("Shorten Payment Terms", "Avg PT days top 20 customers / others (Industry / Region): Overall",
               ["Avg PT Days Top 20 (Overall)"], "identity", numerator="Avg PT Days Top 20 (Overall)",
               metric_name="Avg PT days top 20 customers", region="Overall"),

    MetricSpec("Shorten Payment Terms", "Avg PT days top 10 customers / others (Industry / Region): Americas",
               ["Avg PT Days Top 10 (Americas)"], "identity", numerator="Avg PT Days Top 10 (Americas)",
               metric_name="Avg PT days top 10 customers", region="Americas"),
    MetricSpec("Shorten Payment Terms", "Avg PT days top 15 customers / others (Industry / Region): Americas",
               ["Avg PT Days Top 15 (Americas)"], "identity", numerator="Avg PT Days Top 15 (Americas)",
               metric_name="Avg PT days top 15 customers", region="Americas"),
    MetricSpec("Shorten Payment Terms", "Avg PT days top 20 customers / others (Industry / Region): Americas",
               ["Avg PT Days Top 20 (Americas)"], "identity", numerator="Avg PT Days Top 20 (Americas)",
               metric_name="Avg PT days top 20 customers", region="Americas"),

    MetricSpec("Shorten Payment Terms", "Avg PT days top 10 customers / others (Industry / Region): APAC",
               ["Avg PT Days Top 10 (APAC)"], "identity", numerator="Avg PT Days Top 10 (APAC)",
               metric_name="Avg PT days top 10 customers", region="APAC"),
    MetricSpec("Shorten Payment Terms", "Avg PT days top 15 customers / others (Industry / Region): APAC",
               ["Avg PT Days Top 15 (APAC)"], "identity", numerator="Avg PT Days Top 15 (APAC)",
               metric_name="Avg PT days top 15 customers", region="APAC"),
    MetricSpec("Shorten Payment Terms", "Avg PT days top 20 customers / others (Industry / Region): APAC",
               ["Avg PT Days Top 20 (APAC)"], "identity", numerator="Avg PT Days Top 20 (APAC)",
               metric_name="Avg PT days top 20 customers", region="APAC"),

    MetricSpec("Shorten Payment Terms", "Avg PT days top 10 customers / others (Industry / Region): EMEA",
               ["Avg PT Days Top 10 (EMEA)"], "identity", numerator="Avg PT Days Top 10 (EMEA)",
               metric_name="Avg PT days top 10 customers", region="EMEA"),
    MetricSpec("Shorten Payment Terms", "Avg PT days top 15 customers / others (Industry / Region): EMEA",
               ["Avg PT Days Top 15 (EMEA)"], "identity", numerator="Avg PT Days Top 15 (EMEA)",
               metric_name="Avg PT days top 15 customers", region="EMEA"),
    MetricSpec("Shorten Payment Terms", "Avg PT days top 20 customers / others (Industry / Region): EMEA",
               ["Avg PT Days Top 20 (EMEA)"], "identity", numerator="Avg PT Days Top 20 (EMEA)",
               metric_name="Avg PT days top 20 customers", region="EMEA"),

    # ---- Cash Benefit as % sales in >30 PT days bucket ----
    MetricSpec("Shorten Payment Terms", "Cash Benefit as % sales in >30 PT days bucket (Industry / Region): Overall",
               ["Cash Benefit (Overall)", "Sales >30 PT Days (Overall)"], "ratio_pct",
               "Cash Benefit (Overall)", "Sales >30 PT Days (Overall)",
               metric_name="Cash Benefit as % sales in >30 PT days bucket", region="Overall"),
    MetricSpec("Shorten Payment Terms", "Cash Benefit as % sales in >30 PT days bucket (Industry / Region): Americas",
               ["Cash Benefit (Americas)", "Sales >30 PT Days (Americas)"], "ratio_pct",
               "Cash Benefit (Americas)", "Sales >30 PT Days (Americas)",
               metric_name="Cash Benefit as % sales in >30 PT days bucket", region="Americas"),
    MetricSpec("Shorten Payment Terms", "Cash Benefit as % sales in >30 PT days bucket (Industry / Region): APAC",
               ["Cash Benefit (APAC)", "Sales >30 PT Days (APAC)"], "ratio_pct",
               "Cash Benefit (APAC)", "Sales >30 PT Days (APAC)",
               metric_name="Cash Benefit as % sales in >30 PT days bucket", region="APAC"),
    MetricSpec("Shorten Payment Terms", "Cash Benefit as % sales in >30 PT days bucket (Industry / Region): EMEA",
               ["Cash Benefit (EMEA)", "Sales >30 PT Days (EMEA)"], "ratio_pct",
               "Cash Benefit (EMEA)", "Sales >30 PT Days (EMEA)",
               metric_name="Cash Benefit as % sales in >30 PT days bucket", region="EMEA"),

    # ---- Cash Benefit by Shortening PT days as a % of Sales ----
    MetricSpec("Shorten Payment Terms",
               "Cash Benefit by Shortening PT days as a % of Sales (Industry / Region): Overall",
               ["Cash Benefit (Overall)", "Sales (Overall)"], "ratio_pct", "Cash Benefit (Overall)", "Sales (Overall)",
               metric_name="Cash Benefit by Shortening PT days as a % of Sales", region="Overall"),
    MetricSpec("Shorten Payment Terms",
               "Cash Benefit by Shortening PT days as a % of Sales (Industry / Region): Americas",
               ["Cash Benefit (Americas)", "Sales (Americas)"], "ratio_pct", "Cash Benefit (Americas)", "Sales (Americas)",
               metric_name="Cash Benefit by Shortening PT days as a % of Sales", region="Americas"),
    MetricSpec("Shorten Payment Terms",
               "Cash Benefit by Shortening PT days as a % of Sales (Industry / Region): APAC",
               ["Cash Benefit (APAC)", "Sales (APAC)"], "ratio_pct", "Cash Benefit (APAC)", "Sales (APAC)",
               metric_name="Cash Benefit by Shortening PT days as a % of Sales", region="APAC"),
    MetricSpec("Shorten Payment Terms",
               "Cash Benefit by Shortening PT days as a % of Sales (Industry / Region): EMEA",
               ["Cash Benefit (EMEA)", "Sales (EMEA)"], "ratio_pct", "Cash Benefit (EMEA)", "Sales (EMEA)",
               metric_name="Cash Benefit by Shortening PT days as a % of Sales", region="EMEA"),

    # ============================================================
    # 4) HARMONIZE PAYMENT TERMS
    # ============================================================

    MetricSpec("Harmonize Payment Terms", "# distinct PT days offered to top 10 customers (Industry / Region): Overall",
               ["Distinct PT Days Top 10 (Overall)"], "identity", numerator="Distinct PT Days Top 10 (Overall)",
               metric_name="# distinct PT days offered to top 10 customers", region="Overall"),
    MetricSpec("Harmonize Payment Terms", "# distinct PT days offered to top 15 customers (Industry / Region): Overall",
               ["Distinct PT Days Top 15 (Overall)"], "identity", numerator="Distinct PT Days Top 15 (Overall)",
               metric_name="# distinct PT days offered to top 15 customers", region="Overall"),
    MetricSpec("Harmonize Payment Terms", "# distinct PT days offered to top 20 customers (Industry / Region): Overall",
               ["Distinct PT Days Top 20 (Overall)"], "identity", numerator="Distinct PT Days Top 20 (Overall)",
               metric_name="# distinct PT days offered to top 20 customers", region="Overall"),

    MetricSpec("Harmonize Payment Terms", "# distinct PT days offered to top 10 customers (Industry / Region): Americas",
               ["Distinct PT Days Top 10 (Americas)"], "identity", numerator="Distinct PT Days Top 10 (Americas)",
               metric_name="# distinct PT days offered to top 10 customers", region="Americas"),
    MetricSpec("Harmonize Payment Terms", "# distinct PT days offered to top 15 customers (Industry / Region): Americas",
               ["Distinct PT Days Top 15 (Americas)"], "identity", numerator="Distinct PT Days Top 15 (Americas)",
               metric_name="# distinct PT days offered to top 15 customers", region="Americas"),
    MetricSpec("Harmonize Payment Terms", "# distinct PT days offered to top 20 customers (Industry / Region): Americas",
               ["Distinct PT Days Top 20 (Americas)"], "identity", numerator="Distinct PT Days Top 20 (Americas)",
               metric_name="# distinct PT days offered to top 20 customers", region="Americas"),

    MetricSpec("Harmonize Payment Terms", "# distinct PT days offered to top 10 customers (Industry / Region): APAC",
               ["Distinct PT Days Top 10 (APAC)"], "identity", numerator="Distinct PT Days Top 10 (APAC)",
               metric_name="# distinct PT days offered to top 10 customers", region="APAC"),
    MetricSpec("Harmonize Payment Terms", "# distinct PT days offered to top 15 customers (Industry / Region): APAC",
               ["Distinct PT Days Top 15 (APAC)"], "identity", numerator="Distinct PT Days Top 15 (APAC)",
               metric_name="# distinct PT days offered to top 15 customers", region="APAC"),
    MetricSpec("Harmonize Payment Terms", "# distinct PT days offered to top 20 customers (Industry / Region): APAC",
               ["Distinct PT Days Top 20 (APAC)"], "identity", numerator="Distinct PT Days Top 20 (APAC)",
               metric_name="# distinct PT days offered to top 20 customers", region="APAC"),

    MetricSpec("Harmonize Payment Terms", "# distinct PT days offered to top 10 customers (Industry / Region): EMEA",
               ["Distinct PT Days Top 10 (EMEA)"], "identity", numerator="Distinct PT Days Top 10 (EMEA)",
               metric_name="# distinct PT days offered to top 10 customers", region="EMEA"),
    MetricSpec("Harmonize Payment Terms", "# distinct PT days offered to top 15 customers (Industry / Region): EMEA",
               ["Distinct PT Days Top 15 (EMEA)"], "identity", numerator="Distinct PT Days Top 15 (EMEA)",
               metric_name="# distinct PT days offered to top 15 customers", region="EMEA"),
    MetricSpec("Harmonize Payment Terms", "# distinct PT days offered to top 20 customers (Industry / Region): EMEA",
               ["Distinct PT Days Top 20 (EMEA)"], "identity", numerator="Distinct PT Days Top 20 (EMEA)",
               metric_name="# distinct PT days offered to top 20 customers", region="EMEA"),

    # ---- distinct PT days offered by top X vendors (A/B/C split) ----
    MetricSpec("Harmonize Payment Terms", "distinct PT days offered by top X vendors (A: Top 80%) | Overall",
               ["Distinct PT Days Top Vendors A (Overall)"], "identity", numerator="Distinct PT Days Top Vendors A (Overall)",
               metric_name="distinct PT days offered by top X vendors (A: Top 80%)", region="Overall"),
    MetricSpec("Harmonize Payment Terms", "distinct PT days offered by top X vendors (B: Next 15%) | Overall",
               ["Distinct PT Days Top Vendors B (Overall)"], "identity", numerator="Distinct PT Days Top Vendors B (Overall)",
               metric_name="distinct PT days offered by top X vendors (B: Next 15%)", region="Overall"),
    MetricSpec("Harmonize Payment Terms", "distinct PT days offered by top X vendors (C: Last 5%) | Overall",
               ["Distinct PT Days Top Vendors C (Overall)"], "identity", numerator="Distinct PT Days Top Vendors C (Overall)",
               metric_name="distinct PT days offered by top X vendors (C: Last 5%)", region="Overall"),

    MetricSpec("Harmonize Payment Terms", "distinct PT days offered by top X vendors (A: Top 80%) | Americas",
               ["Distinct PT Days Top Vendors A (Americas)"], "identity", numerator="Distinct PT Days Top Vendors A (Americas)",
               metric_name="distinct PT days offered by top X vendors (A: Top 80%)", region="Americas"),
    MetricSpec("Harmonize Payment Terms", "distinct PT days offered by top X vendors (B: Next 15%) | Americas",
               ["Distinct PT Days Top Vendors B (Americas)"], "identity", numerator="Distinct PT Days Top Vendors B (Americas)",
               metric_name="distinct PT days offered by top X vendors (B: Next 15%)", region="Americas"),
    MetricSpec("Harmonize Payment Terms", "distinct PT days offered by top X vendors (C: Last 5%) | Americas",
               ["Distinct PT Days Top Vendors C (Americas)"], "identity", numerator="Distinct PT Days Top Vendors C (Americas)",
               metric_name="distinct PT days offered by top X vendors (C: Last 5%)", region="Americas"),

    MetricSpec("Harmonize Payment Terms", "distinct PT days offered by top X vendors (A: Top 80%) | APAC",
               ["Distinct PT Days Top Vendors A (APAC)"], "identity", numerator="Distinct PT Days Top Vendors A (APAC)",
               metric_name="distinct PT days offered by top X vendors (A: Top 80%)", region="APAC"),
    MetricSpec("Harmonize Payment Terms", "distinct PT days offered by top X vendors (B: Next 15%) | APAC",
               ["Distinct PT Days Top Vendors B (APAC)"], "identity", numerator="Distinct PT Days Top Vendors B (APAC)",
               metric_name="distinct PT days offered by top X vendors (B: Next 15%)", region="APAC"),
    MetricSpec("Harmonize Payment Terms", "distinct PT days offered by top X vendors (C: Last 5%) | APAC",
               ["Distinct PT Days Top Vendors C (APAC)"], "identity", numerator="Distinct PT Days Top Vendors C (APAC)",
               metric_name="distinct PT days offered by top X vendors (C: Last 5%)", region="APAC"),

    MetricSpec("Harmonize Payment Terms", "distinct PT days offered by top X vendors (A: Top 80%) | EMEA",
               ["Distinct PT Days Top Vendors A (EMEA)"], "identity", numerator="Distinct PT Days Top Vendors A (EMEA)",
               metric_name="distinct PT days offered by top X vendors (A: Top 80%)", region="EMEA"),
    MetricSpec("Harmonize Payment Terms", "distinct PT days offered by top X vendors (B: Next 15%) | EMEA",
               ["Distinct PT Days Top Vendors B (EMEA)"], "identity", numerator="Distinct PT Days Top Vendors B (EMEA)",
               metric_name="distinct PT days offered by top X vendors (B: Next 15%)", region="EMEA"),
    MetricSpec("Harmonize Payment Terms", "distinct PT days offered by top X vendors (C: Last 5%) | EMEA",
               ["Distinct PT Days Top Vendors C (EMEA)"], "identity", numerator="Distinct PT Days Top Vendors C (EMEA)",
               metric_name="distinct PT days offered by top X vendors (C: Last 5%)", region="EMEA"),

    # ---- Std devn of PT days offered ----
    MetricSpec("Harmonize Payment Terms", "Std devn of PT days offered (Industry / Region / Top customers3): Overall",
               ["Std Dev PT Days (Overall)"], "identity", numerator="Std Dev PT Days (Overall)",
               metric_name="Std devn of PT days offered", region="Overall"),
    MetricSpec("Harmonize Payment Terms", "Std devn of PT days offered (Industry / Region / Top customers3): Americas",
               ["Std Dev PT Days (Americas)"], "identity", numerator="Std Dev PT Days (Americas)",
               metric_name="Std devn of PT days offered", region="Americas"),
    MetricSpec("Harmonize Payment Terms", "Std devn of PT days offered (Industry / Region / Top customers3): APAC",
               ["Std Dev PT Days (APAC)"], "identity", numerator="Std Dev PT Days (APAC)",
               metric_name="Std devn of PT days offered", region="APAC"),
    MetricSpec("Harmonize Payment Terms", "Std devn of PT days offered (Industry / Region / Top customers3): EMEA",
               ["Std Dev PT Days (EMEA)"], "identity", numerator="Std Dev PT Days (EMEA)",
               metric_name="Std devn of PT days offered", region="EMEA"),

    # ---- Cash Benefit as % of sales ----
    MetricSpec("Harmonize Payment Terms", "Cash Benefit as % of sales2 (Industry / Region): Overall",
               ["Cash Benefit (Overall)", "Sales (Overall)"], "ratio_pct", "Cash Benefit (Overall)", "Sales (Overall)",
               metric_name="Cash Benefit as % of sales", region="Overall"),
    MetricSpec("Harmonize Payment Terms", "Cash Benefit as % of sales2 (Industry / Region): Americas",
               ["Cash Benefit (Americas)", "Sales (Americas)"], "ratio_pct", "Cash Benefit (Americas)", "Sales (Americas)",
               metric_name="Cash Benefit as % of sales", region="Americas"),
    MetricSpec("Harmonize Payment Terms", "Cash Benefit as % of sales2 (Industry / Region): APAC",
               ["Cash Benefit (APAC)", "Sales (APAC)"], "ratio_pct", "Cash Benefit (APAC)", "Sales (APAC)",
               metric_name="Cash Benefit as % of sales", region="APAC"),
    MetricSpec("Harmonize Payment Terms", "Cash Benefit as % of sales2 (Industry / Region): EMEA",
               ["Cash Benefit (EMEA)", "Sales (EMEA)"], "ratio_pct", "Cash Benefit (EMEA)", "Sales (EMEA)",
               metric_name="Cash Benefit as % of sales", region="EMEA"),

    # ============================================================
    # 5) HARMONIZE DISCOUNT TERMS
    # ============================================================

    MetricSpec("Harmonize Discount Terms", "# distinct Discount terms offered to top 10 customers (Industry / Region): Overall",
               ["Distinct Discount Terms Top 10 (Overall)"], "identity", numerator="Distinct Discount Terms Top 10 (Overall)",
               metric_name="# distinct Discount terms offered to top 10 customers", region="Overall"),
    MetricSpec("Harmonize Discount Terms", "# distinct Discount terms offered to top 10 customers (Industry / Region): Americas",
               ["Distinct Discount Terms Top 10 (Americas)"], "identity", numerator="Distinct Discount Terms Top 10 (Americas)",
               metric_name="# distinct Discount terms offered to top 10 customers", region="Americas"),
    MetricSpec("Harmonize Discount Terms", "# distinct Discount terms offered to top 10 customers (Industry / Region): APAC",
               ["Distinct Discount Terms Top 10 (APAC)"], "identity", numerator="Distinct Discount Terms Top 10 (APAC)",
               metric_name="# distinct Discount terms offered to top 10 customers", region="APAC"),
    MetricSpec("Harmonize Discount Terms", "# distinct Discount terms offered to top 10 customers (Industry / Region): EMEA",
               ["Distinct Discount Terms Top 10 (EMEA)"], "identity", numerator="Distinct Discount Terms Top 10 (EMEA)",
               metric_name="# distinct Discount terms offered to top 10 customers", region="EMEA"),

    MetricSpec("Harmonize Discount Terms", "Avg discount % offered (Top customers / Industry / Region): Overall",
               ["Avg Discount % (Overall)"], "identity", numerator="Avg Discount % (Overall)",
               metric_name="Avg discount % offered", region="Overall"),
    MetricSpec("Harmonize Discount Terms", "Avg discount % offered (Top customers / Industry / Region): Americas",
               ["Avg Discount % (Americas)"], "identity", numerator="Avg Discount % (Americas)",
               metric_name="Avg discount % offered", region="Americas"),
    MetricSpec("Harmonize Discount Terms", "Avg discount % offered (Top customers / Industry / Region): APAC",
               ["Avg Discount % (APAC)"], "identity", numerator="Avg Discount % (APAC)",
               metric_name="Avg discount % offered", region="APAC"),
    MetricSpec("Harmonize Discount Terms", "Avg discount % offered (Top customers / Industry / Region): EMEA",
               ["Avg Discount % (EMEA)"], "identity", numerator="Avg Discount % (EMEA)",
               metric_name="Avg discount % offered", region="EMEA"),

    MetricSpec("Harmonize Discount Terms", "Avg discount terms as % of WACC (Industry / Region): Overall",
               ["Avg Discount % (Overall)", "WACC (Overall)"], "ratio_pct", "Avg Discount % (Overall)", "WACC (Overall)",
               metric_name="Avg discount terms as % of WACC", region="Overall"),
    MetricSpec("Harmonize Discount Terms", "Avg discount terms as % of WACC (Industry / Region): Americas",
               ["Avg Discount % (Americas)", "WACC (Americas)"], "ratio_pct", "Avg Discount % (Americas)", "WACC (Americas)",
               metric_name="Avg discount terms as % of WACC", region="Americas"),
    MetricSpec("Harmonize Discount Terms", "Avg discount terms as % of WACC (Industry / Region): APAC",
               ["Avg Discount % (APAC)", "WACC (APAC)"], "ratio_pct", "Avg Discount % (APAC)", "WACC (APAC)",
               metric_name="Avg discount terms as % of WACC", region="APAC"),
    MetricSpec("Harmonize Discount Terms", "Avg discount terms as % of WACC (Industry / Region): EMEA",
               ["Avg Discount % (EMEA)", "WACC (EMEA)"], "ratio_pct", "Avg Discount % (EMEA)", "WACC (EMEA)",
               metric_name="Avg discount terms as % of WACC", region="EMEA"),

    MetricSpec("Harmonize Discount Terms", "Cash Benefit as % of sales4 (Industry / Region): Overall",
               ["Cash Benefit (Overall)", "Sales with Discounts (Overall)"], "ratio_pct", "Cash Benefit (Overall)", "Sales with Discounts (Overall)",
               metric_name="Cash Benefit as % of sales", region="Overall"),
    MetricSpec("Harmonize Discount Terms", "Cash Benefit as % of sales4 (Industry / Region): Americas",
               ["Cash Benefit (Americas)", "Sales with Discounts (Americas)"], "ratio_pct", "Cash Benefit (Americas)", "Sales with Discounts (Americas)",
               metric_name="Cash Benefit as % of sales", region="Americas"),
    MetricSpec("Harmonize Discount Terms", "Cash Benefit as % of sales4 (Industry / Region): APAC",
               ["Cash Benefit (APAC)", "Sales with Discounts (APAC)"], "ratio_pct", "Cash Benefit (APAC)", "Sales with Discounts (APAC)",
               metric_name="Cash Benefit as % of sales", region="APAC"),
    MetricSpec("Harmonize Discount Terms", "Cash Benefit as % of sales4 (Industry / Region): EMEA",
               ["Cash Benefit (EMEA)", "Sales with Discounts (EMEA)"], "ratio_pct", "Cash Benefit (EMEA)", "Sales with Discounts (EMEA)",
               metric_name="Cash Benefit as % of sales", region="EMEA"),

    # ============================================================
    # 6) ELIMINATE INCORRECT DISCOUNTS
    # ============================================================

    MetricSpec("Eliminate incorrect Discounts",
               "Discounts eligible for vs Discounts given (Top customers / Region / Industry): Overall",
               ["Discounts Eligible (Overall)", "Discounts Given (Overall)"], "ratio_pct",
               "Discounts Eligible (Overall)", "Discounts Given (Overall)",
               metric_name="Discounts eligible for vs Discounts given", region="Overall"),
    MetricSpec("Eliminate incorrect Discounts",
               "Discounts eligible for vs Discounts given (Top customers / Region / Industry): Americas",
               ["Discounts Eligible (Americas)", "Discounts Given (Americas)"], "ratio_pct",
               "Discounts Eligible (Americas)", "Discounts Given (Americas)",
               metric_name="Discounts eligible for vs Discounts given", region="Americas"),
    MetricSpec("Eliminate incorrect Discounts",
               "Discounts eligible for vs Discounts given (Top customers / Region / Industry): APAC",
               ["Discounts Eligible (APAC)", "Discounts Given (APAC)"], "ratio_pct",
               "Discounts Eligible (APAC)", "Discounts Given (APAC)",
               metric_name="Discounts eligible for vs Discounts given", region="APAC"),
    MetricSpec("Eliminate incorrect Discounts",
               "Discounts eligible for vs Discounts given (Top customers / Region / Industry): EMEA",
               ["Discounts Eligible (EMEA)", "Discounts Given (EMEA)"], "ratio_pct",
               "Discounts Eligible (EMEA)", "Discounts Given (EMEA)",
               metric_name="Discounts eligible for vs Discounts given", region="EMEA"),

    # ============================================================
    # 7) TBD
    # ============================================================

    MetricSpec("TBD", "Revised DSO",
               ["Revised DSO"], "identity", numerator="Revised DSO",
               metric_name="Revised DSO", region="—"),
    MetricSpec("TBD", "Change in DSO",
               ["Change in DSO"], "identity", numerator="Change in DSO",
               metric_name="Change in DSO", region="—"),
    MetricSpec("TBD", "Cash benefit as % of inscope spend",
               ["Cash Benefit", "Inscope Spend"], "ratio_pct", "Cash Benefit", "Inscope Spend",
               metric_name="Cash benefit as % of inscope spend", region="—"),
]

# ============================================================
# CALCULATION ENGINE
# ============================================================

def safe_float(x: Any):
    try:
        return float(str(x).replace(",", ""))
    except:
        return None


def calc_metric(spec: MetricSpec, inputs: Dict[str, Any]):
    outputs = {}

    if spec.calc_type == "identity":
        outputs["Value"] = safe_float(inputs.get(spec.numerator))
        return outputs

    if spec.calc_type == "ratio_pct":
        num = safe_float(inputs.get(spec.numerator))
        den = safe_float(inputs.get(spec.denominator))
        outputs["Value (%)"] = None if num is None or den is None or den == 0 else (num / den) * 100
        return outputs

    if spec.calc_type == "multi_ratio_pct":
        den = safe_float(inputs.get(spec.multi_denominator))
        for n in spec.numerators or []:
            num = safe_float(inputs.get(n))
            outputs[f"{n} (%)"] = None if num is None or den is None or den == 0 else (num / den) * 100
        return outputs

    if spec.calc_type == "days_ratio":
        num = safe_float(inputs.get(spec.numerator))
        den = safe_float(inputs.get(spec.denominator))
        outputs["Value (Days)"] = None if num is None or den is None or den == 0 else (num / den) * 365
        return outputs

    return outputs


# ============================================================
# SESSION STATE
# ============================================================

if "demographics" not in st.session_state:
    st.session_state.demographics = {
        "Company": "",
        "Industry": "",
        "Industry L2": "",
        "Primary Region": "",
        "Currency": "",
        "FY / Period": "",
    }

if "kpi_inputs" not in st.session_state:
    st.session_state.kpi_inputs = {}

if "kpi_comments" not in st.session_state:
    st.session_state.kpi_comments = {}

# ============================================================
# SIDEBAR
# ============================================================

menu = st.sidebar.radio("Menu", ["Demographics", "KPI Components & Value"])

# ============================================================
# DEMOGRAPHICS PAGE
# ============================================================

if menu == "Demographics":

    st.subheader("Demographics")
    st.caption("Fill in the client details below. These will be included in every row of the exported Excel.")

    col1, col2 = st.columns(2)

    with col1:
        st.session_state.demographics["Company"] = st.text_input(
            "Company", st.session_state.demographics["Company"],
            placeholder="e.g. American Airlines"
        )
        st.session_state.demographics["Industry"] = st.text_input(
            "Industry", st.session_state.demographics["Industry"],
            placeholder="e.g. Advanced Manufacturing & Services"
        )
        st.session_state.demographics["Industry L2"] = st.text_input(
            "Industry L2", st.session_state.demographics["Industry L2"],
            placeholder="e.g. Airlines, Logistics & Transport"
        )

    with col2:
        region_options = ["", "AMER", "APAC", "EMEA", "Global"]
        current_region = st.session_state.demographics["Primary Region"]
        st.session_state.demographics["Primary Region"] = st.selectbox(
            "Primary Region",
            options=region_options,
            index=region_options.index(current_region) if current_region in region_options else 0
        )
        st.session_state.demographics["Currency"] = st.text_input(
            "Currency", st.session_state.demographics["Currency"],
            placeholder="e.g. USD"
        )
        st.session_state.demographics["FY / Period"] = st.text_input(
            "FY / Period", st.session_state.demographics["FY / Period"],
            placeholder="e.g. FY2024"
        )

    st.divider()
    st.markdown("**Preview**")
    preview_df = pd.DataFrame([{
        "Function": "AR",
        "Company": st.session_state.demographics["Company"],
        "Industry": st.session_state.demographics["Industry"],
        "Industry L2": st.session_state.demographics["Industry L2"],
        "Primary Region": st.session_state.demographics["Primary Region"],
        "Currency": st.session_state.demographics["Currency"],
        "FY / Period": st.session_state.demographics["FY / Period"],
    }])
    st.dataframe(preview_df, use_container_width=True, hide_index=True)

# ============================================================
# KPI PAGE
# ============================================================

else:
    export_rows = []
    component_export_rows = []

    metrics_by_lever = defaultdict(list)
    for spec in AR_METRICS:
        metrics_by_lever[spec.lever].append(spec)

    for lever, specs in metrics_by_lever.items():

        st.markdown(f"## {lever}")

        for spec in specs:
            display_label = f"{spec.metric_name}  |  {spec.region}"

            with st.expander(display_label):
                col1, col2, col3 = st.columns(3)
                col1.markdown(f"**Lever:** {spec.lever}")
                col2.markdown(f"**Metric:** {spec.metric_name}")
                col3.markdown(f"**Region:** {spec.region}")

                st.divider()

                kpi_key = f"{spec.lever}|||{spec.kpi}"

                if kpi_key not in st.session_state.kpi_inputs:
                    st.session_state.kpi_inputs[kpi_key] = {comp: "" for comp in spec.components}

                if kpi_key not in st.session_state.kpi_comments:
                    st.session_state.kpi_comments[kpi_key] = ""

                inputs = {}

                # Component Inputs
                for comp in spec.components:
                    st.session_state.kpi_inputs[kpi_key][comp] = st.text_input(
                        comp,
                        value=st.session_state.kpi_inputs[kpi_key].get(comp, ""),
                        key=f"in::{kpi_key}::{comp}"
                    )
                    inputs[comp] = st.session_state.kpi_inputs[kpi_key][comp]

                # Comment box
                st.session_state.kpi_comments[kpi_key] = st.text_area(
                    "Comment",
                    value=st.session_state.kpi_comments.get(kpi_key, ""),
                    key=f"comment::{kpi_key}",
                    placeholder="Describe how the component values were extracted..."
                )

                # Calculate
                outputs = calc_metric(spec, inputs)

                # Output Fields
                for out_label, v in outputs.items():
                    st.text_input(
                        out_label,
                        value="" if v is None else f"{v:,.6f}",
                        disabled=True,
                        key=f"out::{kpi_key}::{out_label}"
                    )

                # Export Rows
                for out_label, v in outputs.items():
                    if "(%)" in out_label or out_label == "Value (%)":
                        unit = "%"
                    elif "Days" in out_label:
                        unit = "Days"
                    else:
                        unit = "Abs."

                    export_rows.append({
                        "Function": "AR",
                        "Lever": spec.lever,
                        "Company": st.session_state.demographics.get("Company", ""),
                        "Industry": st.session_state.demographics.get("Industry", ""),
                        "Industry L2": st.session_state.demographics.get("Industry L2", ""),
                        "Primary Region": st.session_state.demographics.get("Primary Region", ""),
                        "Currency": st.session_state.demographics.get("Currency", ""),
                        "FY / Period": st.session_state.demographics.get("FY / Period", ""),
                        "Region": spec.region,
                        "KPI": spec.metric_name,
                        "Value": "" if v is None else v,
                        "Unit": unit,
                    })

                    for comp in spec.components:
                        component_export_rows.append({
                            "Function": "AR",
                            "Lever": spec.lever,
                            "Company": st.session_state.demographics.get("Company", ""),
                            "Industry": st.session_state.demographics.get("Industry", ""),
                            "Industry L2": st.session_state.demographics.get("Industry L2", ""),
                            "Primary Region": st.session_state.demographics.get("Primary Region", ""),
                            "Currency": st.session_state.demographics.get("Currency", ""),
                            "FY / Period": st.session_state.demographics.get("FY / Period", ""),
                            "Region": spec.region,
                            "KPI": spec.metric_name,
                            "Output Label": out_label,
                            "Metric Value": "" if v is None else v,
                            "Metric Unit": unit,
                            "Component": comp,
                            "Component Value": inputs.get(comp, ""),
                            "Comment": st.session_state.kpi_comments.get(kpi_key, "")
                        })

    # ============================================================
    # EXPORT SECTION
    # ============================================================

    st.markdown("---")
    st.subheader("Export")

    export_df = pd.DataFrame(export_rows, columns=[
        "Function", "Lever", "Company", "Industry", "Industry L2",
        "Primary Region", "Currency", "FY / Period", "Region", "KPI", "Value", "Unit"
    ])

    component_export_df = pd.DataFrame(component_export_rows, columns=[
        "Function", "Lever", "Company", "Industry", "Industry L2",
        "Primary Region", "Currency", "FY / Period", "Region", "KPI", "Output Label",
        "Metric Value", "Metric Unit", "Component", "Component Value", "Comment"
    ])

    st.markdown("### KPI export preview")
    st.dataframe(export_df, use_container_width=True, hide_index=True)

    st.markdown("### Component-level export preview")
    st.dataframe(component_export_df, use_container_width=True, hide_index=True)

    # Main KPI export
    towrite_main = io.BytesIO()
    with pd.ExcelWriter(towrite_main, engine="xlsxwriter") as writer:
        export_df.to_excel(writer, index=False, sheet_name="AR KPIs")
    towrite_main.seek(0)

    st.download_button(
        "Download KPI Excel",
        data=towrite_main.getvalue(),
        file_name="ar_kpis_export.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
        key="download_ar_excel"
    )

    # Component-level export
    towrite_components = io.BytesIO()
    with pd.ExcelWriter(towrite_components, engine="xlsxwriter") as writer:
        component_export_df.to_excel(writer, index=False, sheet_name="AR KPI Components")
    towrite_components.seek(0)

    st.download_button(
        "Download Component Excel",
        data=towrite_components.getvalue(),
        file_name="ar_kpi_components_export.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
        key="download_ar_component_excel"
    )

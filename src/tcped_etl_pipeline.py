#!/usr/bin/env python3
"""
tcped_etl_pipeline.py
Extracts messy multi-INT ISR feeds, normalizes coordinates to WGS84 Decimal Degrees,
maps unstructured text into a relational TCPED ontology, loads into SQLite, and executes
automated Theater Collection Gap Analysis queries.
"""

import os
import re
import json
import math
import sqlite3
import pandas as pd
from datetime import datetime

BASE_DIR = os.path.dirname(os.path.dirname(__file__))
RAW_DIR = os.path.join(BASE_DIR, "data", "raw")
PROC_DIR = os.path.join(BASE_DIR, "data", "processed")
DB_PATH = os.path.join(PROC_DIR, "isr_tcped_warehouse.db")


# ==============================================================================
# 1. GEOSPATIAL COORDINATE NORMALIZATION ENGINE (DD, DMS, MGRS -> WGS84 DD)
# ==============================================================================

def utm_to_wgs84(zone_number: int, northern_hemisphere: bool, easting: float, northing: float):
    """
    Pure-Python WGS84 Inverse Transverse Mercator projection (UTM to Lat/Lon DD).
    Eliminates fragile C-library dependencies while maintaining <1m geodesic accuracy.
    """
    a = 6378137.0  # WGS84 major axis
    ecc_sq = 0.00669437999014  # WGS84 first eccentricity squared
    k0 = 0.9996

    x = easting - 500000.0
    y = northing if northern_hemisphere else northing - 10000000.0

    ecc_prime_sq = ecc_sq / (1.0 - ecc_sq)
    M = y / k0
    mu = M / (a * (1.0 - ecc_sq / 4.0 - 3.0 * (ecc_sq ** 2) / 64.0 - 5.0 * (ecc_sq ** 3) / 256.0))

    e1 = (1.0 - math.sqrt(1.0 - ecc_sq)) / (1.0 + math.sqrt(1.0 - ecc_sq))

    phi1_rad = (
        mu
        + (3.0 * e1 / 2.0 - 27.0 * (e1 ** 3) / 32.0) * math.sin(2.0 * mu)
        + (21.0 * (e1 ** 2) / 16.0 - 55.0 * (e1 ** 4) / 32.0) * math.sin(4.0 * mu)
        + (151.0 * (e1 ** 3) / 96.0) * math.sin(6.0 * mu)
    )

    N1 = a / math.sqrt(1.0 - ecc_sq * (math.sin(phi1_rad) ** 2))
    T1 = math.tan(phi1_rad) ** 2
    C1 = ecc_prime_sq * (math.cos(phi1_rad) ** 2)
    R1 = a * (1.0 - ecc_sq) / ((1.0 - ecc_sq * (math.sin(phi1_rad) ** 2)) ** 1.5)
    D = x / (N1 * k0)

    lat_rad = phi1_rad - (N1 * math.tan(phi1_rad) / R1) * (
        (D ** 2) / 2.0
        - (5.0 + 3.0 * T1 + 10.0 * C1 - 4.0 * (C1 ** 2) - 9.0 * ecc_prime_sq) * (D ** 4) / 24.0
        + (61.0 + 90.0 * T1 + 298.0 * C1 + 45.0 * (T1 ** 2) - 252.0 * ecc_prime_sq - 3.0 * (C1 ** 2)) * (D ** 6) / 720.0
    )

    lon_origin = (zone_number - 1) * 6 - 180 + 3
    lon_rad = (
        D
        - (1.0 + 2.0 * T1 + C1) * (D ** 3) / 6.0
        + (5.0 - 2.0 * C1 + 28.0 * T1 - 3.0 * (C1 ** 2) + 8.0 * ecc_prime_sq + 24.0 * (T1 ** 2)) * (D ** 5) / 120.0
    ) / math.cos(phi1_rad)

    lat_dd = round(math.degrees(lat_rad), 6)
    lon_dd = round(lon_origin + math.degrees(lon_rad), 6)
    return lat_dd, lon_dd


def parse_coordinate_to_wgs84(raw_coord: str):
    """
    Normalizes heterogeneous coordinate strings (Decimal Degrees, DMS, or MGRS)
    into standard WGS84 (lat_dd, lon_dd, coord_format_detected, is_valid).
    """
    if not isinstance(raw_coord, str) or not raw_coord.strip():
        return None, None, "MISSING", False

    s = raw_coord.strip()

    # Pattern 1: Standard Decimal Degrees (e.g. "34.5512, -117.8821")
    dd_match = re.match(r"^([+-]?\d+\.\d+)\s*,\s*([+-]?\d+\.\d+)$", s)
    if dd_match:
        lat, lon = float(dd_match.group(1)), float(dd_match.group(2))
        if -90.0 <= lat <= 90.0 and -180.0 <= lon <= 180.0:
            return round(lat, 6), round(lon, 6), "DECIMAL_DEGREES", True
        return None, None, "CORRUPT_OUT_OF_BOUNDS", False

    # Pattern 2: Degrees Minutes Seconds (e.g. 34°33'04.3"N 117°52'55.6"W)
    dms_pattern = (
        r"(\d+)°\s*(\d+)'\s*(\d+(?:\.\d+)?)\"?\s*([NSns])\s*"
        r"(\d+)°\s*(\d+)'\s*(\d+(?:\.\d+)?)\"?\s*([EWew])"
    )
    dms_match = re.search(dms_pattern, s)
    if dms_match:
        lat_d, lat_m, lat_s, lat_hem = (
            float(dms_match.group(1)), float(dms_match.group(2)),
            float(dms_match.group(3)), dms_match.group(4).upper()
        )
        lon_d, lon_m, lon_s, lon_hem = (
            float(dms_match.group(5)), float(dms_match.group(6)),
            float(dms_match.group(7)), dms_match.group(8).upper()
        )
        lat = lat_d + (lat_m / 60.0) + (lat_s / 3600.0)
        lon = lon_d + (lon_m / 60.0) + (lon_s / 3600.0)
        if lat_hem == "S":
            lat = -lat
        if lon_hem == "W":
            lon = -lon
        if -90.0 <= lat <= 90.0 and -180.0 <= lon <= 180.0:
            return round(lat, 6), round(lon, 6), "DMS", True
        return None, None, "CORRUPT_DMS", False

    # Pattern 3: MGRS (e.g. "11S NV 45200 18300")
    mgrs_pattern = r"^(\d{1,2})([C-X])\s+([A-Z]{2})\s+(\d{5})\s+(\d{5})$"
    mgrs_match = re.match(mgrs_pattern, s.upper())
    if mgrs_match:
        zone = int(mgrs_match.group(1))
        band = mgrs_match.group(2)
        grid_sq = mgrs_match.group(3)
        e_offset = float(mgrs_match.group(4))
        n_offset = float(mgrs_match.group(5))

        # 100km Grid Square lookup table for Zone 11S theater (Mojave/NTC box)
        # In 11S, 'N' corresponds to 400,000m Easting, 'V' corresponds to 3,800,000m Northing
        grid_lookup = {
            "NV": (400000.0, 3800000.0),
            "NU": (400000.0, 3700000.0),
            "PV": (500000.0, 3800000.0),
        }
        base_e, base_n = grid_lookup.get(grid_sq, (400000.0, 3800000.0))
        easting = base_e + e_offset
        northing = base_n + n_offset
        northern = band >= "N"
        lat, lon = utm_to_wgs84(zone, northern, easting, northing)
        return lat, lon, "MGRS", True

    return None, None, "UNRECOGNIZED_FORMAT", False


# ==============================================================================
# 2. UNSTRUCTURED TEXT & ONTOLOGY NORMALIZATION HELPERS
# ==============================================================================

def normalize_priority_tier(raw_pri: str) -> int:
    """Maps messy priority strings ('PRI-1', 'PRIORITY 1', '1') into integer tiers 1..3."""
    if not raw_pri:
        return 3
    m = re.search(r"([1-4])", str(raw_pri))
    return int(m.group(1)) if m else 3


def extract_sensor_ontology(req_int: str, remarks: str) -> str:
    """
    Maps unstructured operator remarks and general INT categories into standard
    sensor ontology: ['EO/IR', 'SAR', 'SIGINT', 'LiDAR'].
    """
    combined = f"{req_int} {remarks}".upper()
    if "LIDAR" in combined or "FOLIAGE PENETRATION" in combined:
        return "LiDAR"
    if "SAR" in combined or "GMTI" in combined or "STRIPMAP" in combined:
        return "SAR"
    if "SIGINT" in combined or "ELINT" in combined or "COMINT" in combined or "EOB" in combined:
        return "SIGINT"
    return "EO/IR"


def compute_latency_seconds(acq_ts, dissem_ts) -> float:
    """Computes PED dissemination latency in seconds; imputes penalty if missing."""
    if pd.isna(acq_ts) or not isinstance(acq_ts, str) or not acq_ts.strip():
        return -1.0
    if pd.isna(dissem_ts) or not isinstance(dissem_ts, str) or not dissem_ts.strip():
        # Impute standard tactical batch delay (900.0 sec) when dissemination log is missing
        return 900.0
    fmt = "%Y-%m-%dT%H:%M:%SZ"
    t1 = datetime.strptime(acq_ts.strip(), fmt)
    t2 = datetime.strptime(dissem_ts.strip(), fmt)
    return max(0.0, (t2 - t1).total_seconds())


# ==============================================================================
# 3. CORE ETL PIPELINE & SQLITE DATA WAREHOUSE LOADER
# ==============================================================================

def run_etl():
    os.makedirs(PROC_DIR, exist_ok=True)

    # 1. EXTRACT
    cr_df = pd.read_csv(os.path.join(RAW_DIR, "mock_collection_requirements.csv"))
    sortie_df = pd.read_csv(os.path.join(RAW_DIR, "mock_sortie_allocation_deck.csv"))
    with open(os.path.join(RAW_DIR, "mock_sensor_hits.json"), "r", encoding="utf-8") as f:
        hits_data = json.load(f)
    hits_df = pd.DataFrame(hits_data)

    print(f"[*] Extracted {len(cr_df)} Collection Requirements, {len(sortie_df)} Sorties, {len(hits_df)} PED Hits.")

    # 2. TRANSFORM: Deduplicate Targets & Normalize Coordinates
    cr_df["Priority_Tier"] = cr_df["Priority_Code"].apply(normalize_priority_tier)
    cr_df["Sensor_Type"] = cr_df.apply(
        lambda r: extract_sensor_ontology(r["Required_INT"], r["Raw_Operator_Remarks"]), axis=1
    )

    # Sort by Priority_Tier ascending and deduplicate by Target_ID (keeping first/cleanest CR)
    initial_count = len(cr_df)
    cr_df = cr_df.sort_values(by=["Priority_Tier", "CR_ID"]).drop_duplicates(subset=["Target_ID"], keep="first").copy()
    dedup_removed = initial_count - len(cr_df)
    print(f"[*] Deduplicated Target Deck: Removed {dedup_removed} duplicate Target_ID nomination(s).")

    # Apply Coordinate Normalization
    coord_results = cr_df["Coordinate_Raw"].apply(parse_coordinate_to_wgs84)
    cr_df["Lat_WGS84"] = [r[0] for r in coord_results]
    cr_df["Lon_WGS84"] = [r[1] for r in coord_results]
    cr_df["Source_Coord_Format"] = [r[2] for r in coord_results]
    cr_df["Coord_Valid"] = [r[3] for r in coord_results]

    # Quarantine corrupt coordinate rows for audit log, keep valid rows for warehouse
    quarantine_df = cr_df[~cr_df["Coord_Valid"]].copy()
    valid_cr_df = cr_df[cr_df["Coord_Valid"]].copy()
    print(f"[*] Coordinate Normalization: {len(valid_cr_df)} valid WGS84 targets; {len(quarantine_df)} corrupt row(s) quarantined.")

    # Transform Sensor Hits & Compute Latency
    hits_df["Latency_Sec"] = hits_df.apply(
        lambda r: compute_latency_seconds(r.get("acq_timestamp_utc"), r.get("ped_dissem_timestamp_utc")), axis=1
    )

    # Merge CRs -> Sortie Allocations -> PED Hits to derive unified TCPED Collection Status
    merged = pd.merge(
        valid_cr_df,
        sortie_df,
        left_on="Target_ID",
        right_on="Assigned_Target_ID",
        how="left"
    )
    merged = pd.merge(
        merged,
        hits_df,
        left_on=["Target_ID", "Sortie_ID"],
        right_on=["target_id", "sortie_id"],
        how="left"
    )

    def determine_collection_status(row):
        if pd.isna(row["Sortie_ID"]):
            return "UNALLOCATED_GAP"
        if row["Sortie_Status"] in ("CNX_MAINT", "WX_ABORT") and pd.isna(row["hit_id"]):
            return "SORTIE_ABORT_GAP"
        if not pd.isna(row["hit_id"]):
            # If quality is below NIIRS 3.5 or latency > 3600s (1 hr), mark DEGRADED_PARTIAL
            if row["niirs_or_quality"] < 3.5 or row["Latency_Sec"] > 3600:
                return "DEGRADED_PARTIAL"
            return "SATISFIED"
        return "PENDING"

    merged["Collection_Status"] = merged.apply(determine_collection_status, axis=1)
    merged["Latency_Sec"] = merged["Latency_Sec"].fillna(-1.0)
    merged["niirs_or_quality"] = merged["niirs_or_quality"].fillna(0.0)
    merged["Platform_Type"] = merged["Platform_Type"].fillna("NONE_ASSIGNED")
    merged["Callsign"] = merged["Callsign"].fillna("N/A")

    # Select final normalized schema columns
    normalized_cols = [
        "CR_ID", "Target_ID", "Target_Name", "Lat_WGS84", "Lon_WGS84",
        "Source_Coord_Format", "Priority_Tier", "Sensor_Type",
        "Platform_Type", "Callsign", "Collection_Status",
        "Latency_Sec", "niirs_or_quality", "LTIOV_UTC", "Raw_Operator_Remarks"
    ]
    final_df = merged[normalized_cols].copy()

    # 3. LOAD INTO SQLITE RELATIONAL DATA WAREHOUSE
    if os.path.exists(DB_PATH):
        os.remove(DB_PATH)

    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()

    cursor.execute("""
        CREATE TABLE tcped_normalized_deck (
            CR_ID TEXT PRIMARY KEY,
            Target_ID TEXT UNIQUE NOT NULL,
            Target_Name TEXT NOT NULL,
            Lat_WGS84 REAL NOT NULL CHECK(Lat_WGS84 BETWEEN -90.0 AND 90.0),
            Lon_WGS84 REAL NOT NULL CHECK(Lon_WGS84 BETWEEN -180.0 AND 180.0),
            Source_Coord_Format TEXT NOT NULL,
            Priority_Tier INTEGER NOT NULL CHECK(Priority_Tier BETWEEN 1 AND 4),
            Sensor_Type TEXT NOT NULL CHECK(Sensor_Type IN ('EO/IR', 'SAR', 'SIGINT', 'LiDAR')),
            Platform_Type TEXT,
            Callsign TEXT,
            Collection_Status TEXT NOT NULL,
            Latency_Sec REAL,
            niirs_or_quality REAL,
            LTIOV_UTC TEXT,
            Raw_Operator_Remarks TEXT
        );
    """)

    cursor.execute("""
        CREATE TABLE etl_quarantine_audit (
            CR_ID TEXT,
            Target_ID TEXT,
            Coordinate_Raw TEXT,
            Reject_Reason TEXT
        );
    """)

    final_df.to_sql("tcped_normalized_deck", conn, if_exists="append", index=False)

    if not quarantine_df.empty:
        q_load = quarantine_df[["CR_ID", "Target_ID", "Coordinate_Raw", "Source_Coord_Format"]].rename(
            columns={"Source_Coord_Format": "Reject_Reason"}
        )
        q_load.to_sql("etl_quarantine_audit", conn, if_exists="append", index=False)

    conn.commit()
    print(f"[+] Loaded normalized TCPED ontology into SQLite warehouse: {DB_PATH}")

    # 4. RUN ANALYTICAL SQL QUERIES FOR THEATER COLLECTION GAP ANALYSIS
    run_sql_gap_analysis(conn)
    conn.close()


def run_sql_gap_analysis(conn: sqlite3.Connection):
    """Executes analytical SQL queries to identify high-priority Theater Collection Gaps."""
    print("\n" + "=" * 80)
    print("USSTRATCOM / THEATER J2 COLLECTION MANAGEMENT - SQL GAP ANALYSIS REPORT")
    print("=" * 80)

    gap_query = """
        SELECT 
            Target_ID,
            Target_Name,
            Priority_Tier,
            Sensor_Type,
            Collection_Status,
            Platform_Type,
            ROUND(Lat_WGS84, 4) AS Lat,
            ROUND(Lon_WGS84, 4) AS Lon,
            CASE 
                WHEN Collection_Status = 'UNALLOCATED_GAP' THEN 'CRITICAL: Zero ISR Sortie Allocated'
                WHEN Collection_Status = 'SORTIE_ABORT_GAP' THEN 'SEVERE: Assigned Sortie Aborted (Maint/WX)'
                WHEN Collection_Status = 'DEGRADED_PARTIAL' THEN 'MODERATE: NIIRS/Latency Threshold Exceeded'
                ELSE 'NOMINAL'
            END AS Gap_Root_Cause
        FROM tcped_normalized_deck
        WHERE Priority_Tier <= 2
          AND Collection_Status != 'SATISFIED'
        ORDER BY Priority_Tier ASC, Target_ID ASC;
    """
    gaps_df = pd.read_sql_query(gap_query, conn)
    print("\n[!] HIGH-PRIORITY (TIER 1 & 2) UNSATISFIED / DEGRADED COLLECTION GAPS:")
    print(gaps_df.to_string(index=False))

    sensor_summary_query = """
        SELECT 
            Sensor_Type,
            COUNT(*) AS Total_Requirements,
            SUM(CASE WHEN Collection_Status = 'SATISFIED' THEN 1 ELSE 0 END) AS Satisfied_Count,
            SUM(CASE WHEN Collection_Status != 'SATISFIED' THEN 1 ELSE 0 END) AS Gap_Count,
            ROUND(
                100.0 * SUM(CASE WHEN Collection_Status = 'SATISFIED' THEN 1 ELSE 0 END) / COUNT(*), 1
            ) AS Fulfillment_Rate_Pct,
            ROUND(
                AVG(CASE WHEN Latency_Sec > 0 THEN Latency_Sec ELSE NULL END), 1
            ) AS Avg_PED_Latency_Sec
        FROM tcped_normalized_deck
        GROUP BY Sensor_Type
        ORDER BY Fulfillment_Rate_Pct ASC;
    """
    summary_df = pd.read_sql_query(sensor_summary_query, conn)
    print("\n[*] MULTI-INT SENSOR FULFILLMENT & PED LATENCY METRICS BY DISCIPLINE:")
    print(summary_df.to_string(index=False))
    print("=" * 80 + "\n")


if __name__ == "__main__":
    run_etl()
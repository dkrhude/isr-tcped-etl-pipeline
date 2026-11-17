#!/usr/bin/env python3
"""
generate_mock_isr_data.py
Simulates unclassified, disparate, dirty exports from theater ISR Collection Management
systems (modeled conceptually after GIMS/PRISM requirements, ROME reconnaissance decks,
and raw downlinked multi-INT sensor telemetry logs).
"""

import os
import json
import random
import csv
from datetime import datetime, timedelta

RAW_DIR = os.path.join(os.path.dirname(os.path.dirname(__file__)), "data", "raw")

def ensure_dirs():
    os.makedirs(RAW_DIR, exist_ok=True)

def generate_collection_requirements_csv():
    """
    Generates simulated theater intelligence collection requirements (CRs) with:
    - Mixed coordinate formats (Decimal Degrees, DMS, MGRS, and corrupt strings)
    - Duplicate Target_IDs (simulating multi-unit nominations for the same BE/facility)
    - Unstructured operator remarks requiring regex ontology extraction
    """
    filepath = os.path.join(RAW_DIR, "mock_collection_requirements.csv")
    headers = [
        "CR_ID", "Target_ID", "Target_Name", "Coordinate_Raw",
        "Priority_Code", "Required_INT", "LTIOV_UTC", "Raw_Operator_Remarks"
    ]

    rows = [
        # Valid DD, Priority 1, SAR requirement
        ["CR-2026-001", "TGT-ALPHA-01", "SAM Garrison North", "34.5512, -117.8821", "PRI-1", "GEOINT", "2026-10-30T06:00:00Z", "REQ: SAR stripmap 0.5m NIIRS 6 min; wx independent all-weather collect."],
        # Duplicate Target_ID with slightly different naming & DMS coordinates
        ["CR-2026-002", "TGT-ALPHA-01", "SAM Garrison N (Dup)", "34°33'04.3\"N 117°52'55.6\"W", "1", "GEOINT", "2026-10-30T06:00:00Z", "Duplicate entry from JTF sub-unit; REQ: SAR moving target indicator."],
        # High Priority SIGINT requirement (Will intentionally have ZERO collection hits -> GAP)
        ["CR-2026-003", "TGT-BRAVO-02", "Mobile C2 Relay Node", "34°45'18.0\"N 117°31'12.0\"W", "PRIORITY 1", "SIGINT", "2026-10-30T04:00:00Z", "CRITICAL: ELINT/COMINT emitter geolocation required prior to H-Hour. Sensor: SIGINT."],
        # MGRS coordinate format, Priority 2, EO/IR requirement
        ["CR-2026-004", "TGT-CHARLIE-03", "Logistics Depot Railhead", "11S NV 45200 18300", "PRI-2", "GEOINT", "2026-10-30T12:00:00Z", "Day/Night EO/IR FMV dwell 30 min; confirm BDA and railcar loading."],
        # High Priority LiDAR/3D requirement (Will have ZERO allocation -> GAP)
        ["CR-2026-005", "TGT-DELTA-04", "Mountain Tunnel Entrance", "11S NV 61800 39400", "PRI-1", "MASINT/GEOINT", "2026-10-30T08:00:00Z", "REQ: High-density LiDAR foliage penetration & elevation model."],
        # Decimal Degrees with reversed/dirty whitespace, Priority 3
        ["CR-2026-006", "TGT-ECHO-05", "Coastal Radar Site", "  34.2105 , -118.1044  ", "3", "SIGINT", "2026-10-30T18:00:00Z", "Routine electronic order of battle (EOB) maintenance collect; SIGINT."],
        # Partially satisfied / High-latency target
        ["CR-2026-007", "TGT-FOXTROT-06", "Forward FARP Pad", "34°18'50.4\"N 117°15'21.6\"W", "PRI-1", "GEOINT", "2026-10-30T05:00:00Z", "Urgent EO/IR pre-strike verification of rotary-wing dispersal."],
        # Corrupt Coordinate Row (Testing ETL Quarantine & Fault Tolerance)
        ["CR-2026-008", "TGT-CORRUPT-99", "Unknown Bunker", "999.9999, -295.0000", "PRI-2", "GEOINT", "2026-10-30T10:00:00Z", "Corrupt coordinate feed from legacy tactical terminal; EO/IR."],
        # Second High-Priority Unallocated Target (Collection Gap)
        ["CR-2026-009", "TGT-GOLF-07", "Mobile Ballistic TEL Hide", "11S NV 78100 51200", "PRI-1", "GEOINT", "2026-10-30T03:30:00Z", "Time-Sensitive Target (TST); requires SAR / GMTI wide-area search."],
        # Low priority target with full coverage (demonstrates allocation inefficiency / over-collection)
        ["CR-2026-010", "TGT-HOTEL-08", "Abandoned Airfield Strip", "34.8910, -117.4020", "PRI-3", "GEOINT", "2026-10-30T22:00:00Z", "Routine weekly change detection; EO/IR."]
    ]

    with open(filepath, mode="w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(headers)
        writer.writerows(rows)

    print(f"[+] Generated raw collection requirements: {filepath} ({len(rows)} records)")


def generate_platform_sortie_deck_csv():
    """
    Generates simulated ATO / Reconnaissance Sortie Allocation Deck.
    Notice that TGT-BRAVO-02, TGT-DELTA-04, and TGT-GOLF-07 have NO allocated sorties
    or suffer from severe weather/maintenance aborts.
    """
    filepath = os.path.join(RAW_DIR, "mock_sortie_allocation_deck.csv")
    headers = [
        "Sortie_ID", "Platform_Type", "Callsign", "Assigned_Target_ID",
        "Primary_Sensor", "Scheduled_TOT_UTC", "Sortie_Status"
    ]

    rows = [
        ["SRT-101", "MQ-9A Block 5", "REAPER21", "TGT-ALPHA-01", "Lynx SAR", "2026-10-29T23:00:00Z", "EXECUTED"],
        ["SRT-102", "RQ-4B Global Hawk", "FORTE10", "TGT-CHARLIE-03", "MS-177 EO/IR", "2026-10-29T23:30:00Z", "EXECUTED"],
        ["SRT-103", "RC-135V/W Rivet Joint", "OLIVE55", "TGT-BRAVO-02", "SIGINT Suite", "2026-10-30T01:00:00Z", "CNX_MAINT"],  # Canceled! Causes Gap
        ["SRT-104", "RC-135V/W Rivet Joint", "OLIVE55", "TGT-ECHO-05", "SIGINT Suite", "2026-10-29T22:15:00Z", "EXECUTED"],
        ["SRT-105", "MQ-1C Gray Eagle", "GHOST14", "TGT-FOXTROT-06", "MTS-A EO/IR", "2026-10-29T21:00:00Z", "WX_ABORT"],  # Aborted!
        ["SRT-106", "MQ-9A Block 5", "REAPER22", "TGT-HOTEL-08", "MTS-B EO/IR", "2026-10-29T20:00:00Z", "EXECUTED"]
        # Note: TGT-DELTA-04 and TGT-GOLF-07 have zero sorties even scheduled.
    ]

    with open(filepath, mode="w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(headers)
        writer.writerows(rows)

    print(f"[+] Generated raw sortie allocation deck: {filepath} ({len(rows)} records)")


def generate_sensor_telemetry_hits_json():
    """
    Generates simulated PED / sensor downlink logs in JSON format.
    Includes missing timestamps, latency variations, and unstructured remarks.
    """
    filepath = os.path.join(RAW_DIR, "mock_sensor_hits.json")

    hits = [
        {
            "hit_id": "PED-9001",
            "sortie_id": "SRT-101",
            "target_id": "TGT-ALPHA-01",
            "sensor_payload": "SAR_STRIPMAP",
            "acq_timestamp_utc": "2026-10-29T23:02:10Z",
            "ped_dissem_timestamp_utc": "2026-10-29T23:08:45Z",
            "niirs_or_quality": 6.2,
            "collection_remarks": "COLLECTED: Clear SAR return on TEL garages; 4 vehicles detected."
        },
        {
            "hit_id": "PED-9002",
            "sortie_id": "SRT-102",
            "target_id": "TGT-CHARLIE-03",
            "sensor_payload": "EO_IR_FMV",
            "acq_timestamp_utc": "2026-10-29T23:35:00Z",
            "ped_dissem_timestamp_utc": "2026-10-29T23:39:10Z",
            "niirs_or_quality": 5.8,
            "collection_remarks": "COLLECTED: IR tracks thermal signatures on locomotive."
        },
        {
            "hit_id": "PED-9003",
            "sortie_id": "SRT-104",
            "target_id": "TGT-ECHO-05",
            "sensor_payload": "HIGH_BAND_SIGINT",
            "acq_timestamp_utc": "2026-10-29T22:18:00Z",
            "ped_dissem_timestamp_utc": None,  # Missing dissemination timestamp (test imputer)
            "niirs_or_quality": 8.0,
            "collection_remarks": "COLLECTED: S-Band acquisition radar PRF logged; LOB +/- 1.2 deg."
        },
        {
            "hit_id": "PED-9004",
            "sortie_id": "SRT-105",
            "target_id": "TGT-FOXTROT-06",
            "sensor_payload": "EO_IR_FMV",
            "acq_timestamp_utc": "2026-10-29T21:05:00Z",
            "ped_dissem_timestamp_utc": "2026-10-29T22:55:00Z",  # 110 min latency -> Stale/Partial
            "niirs_or_quality": 2.1,
            "collection_remarks": "DEGRADED: 85% cloud deck obscuration; partial perimeter visible only."
        },
        {
            "hit_id": "PED-9005",
            "sortie_id": "SRT-106",
            "target_id": "TGT-HOTEL-08",
            "sensor_payload": "EO_IR_FMV",
            "acq_timestamp_utc": "2026-10-29T20:04:00Z",
            "ped_dissem_timestamp_utc": "2026-10-29T20:06:30Z",
            "niirs_or_quality": 7.1,
            "collection_remarks": "COLLECTED: Runway clear, no activity."
        }
    ]

    with open(filepath, mode="w", encoding="utf-8") as f:
        json.dump(hits, f, indent=4)

    print(f"[+] Generated raw sensor PED hits JSON: {filepath} ({len(hits)} records)")


def main():
    ensure_dirs()
    generate_collection_requirements_csv()
    generate_platform_sortie_deck_csv()
    generate_sensor_telemetry_hits_json()
    print("[*] Mock ISR Multi-INT dataset generation complete.\n")


if __name__ == "__main__":
    main()
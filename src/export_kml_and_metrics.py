#!/usr/bin/env python3
"""
export_kml_and_metrics.py
Queries the normalized SQLite ISR data warehouse to generate:
1. A clean flat CSV ready for ArcGIS Pro / Portal or Tableau dashboard ingestion.
2. An OGC-compliant Google Earth Pro (.kml) file plotting target placemarks and
   color-coded 3D extruded sensor coverage / gap polygons.
"""

import os
import math
import sqlite3
import pandas as pd

BASE_DIR = os.path.dirname(os.path.dirname(__file__))
PROC_DIR = os.path.join(BASE_DIR, "data", "processed")
DB_PATH = os.path.join(PROC_DIR, "isr_tcped_warehouse.db")
CSV_OUT = os.path.join(PROC_DIR, "arcgis_tableau_tcped_export.csv")
KML_OUT = os.path.join(PROC_DIR, "theater_isr_collection_gaps.kml")


def generate_octagon_polygon_coords(lat_dd: float, lon_dd: float, radius_km: float = 3.5, altitude_m: float = 1200.0) -> str:
    """
    Computes an 8-point closed polygon ring around a target coordinate to visualize
    sensor footprint coverage or collection gap bounding boxes in Google Earth Pro.
    """
    coords = []
    lat_rad = math.radians(lat_dd)
    deg_per_km_lat = 1.0 / 111.132
    deg_per_km_lon = 1.0 / (111.320 * math.cos(lat_rad))

    for i in range(9):  # 9 points to close the ring (0 == 8)
        angle_rad = math.radians(i * 45.0)
        d_lat = radius_km * math.sin(angle_rad) * deg_per_km_lat
        d_lon = radius_km * math.cos(angle_rad) * deg_per_km_lon
        pt_lat = lat_dd + d_lat
        pt_lon = lon_dd + d_lon
        coords.append(f"{pt_lon:.6f},{pt_lat:.6f},{altitude_m:.1f}")

    return " ".join(coords)


def export_csv_and_kml():
    if not os.path.exists(DB_PATH):
        raise FileNotFoundError(f"Database not found at {DB_PATH}. Run tcped_etl_pipeline.py first.")

    conn = sqlite3.connect(DB_PATH)
    df = pd.read_sql_query("SELECT * FROM tcped_normalized_deck ORDER BY Priority_Tier ASC, Target_ID ASC;", conn)
    conn.close()

    # 1. Export Clean ArcGIS / Tableau CSV
    df["Is_Collection_Gap"] = df["Collection_Status"].apply(lambda s: 1 if s != "SATISFIED" else 0)
    df["PED_Latency_Min"] = df["Latency_Sec"].apply(lambda s: round(s / 60.0, 2) if s > 0 else None)
    df.to_csv(CSV_OUT, index=False, encoding="utf-8")
    print(f"[+] Exported ArcGIS / Tableau ready CSV: {CSV_OUT} ({len(df)} records)")

    # 2. Build OGC KML 2.2 Document
    # Note: KML uses AABBGGRR hex color format:
    # Green (Satisfied): 7f00ff00 | Amber (Degraded): 7f00a5ff | Red (Gap): 7f0000ff
    kml_header = """<?xml version="1.0" encoding="UTF-8"?>
<kml xmlns="http://www.opengis.net/kml/2.2">
<Document>
    <name>Theater ISR TCPED Collection Deck &amp; Gap Analysis</name>
    <description>Unclassified Physics/ETL Generated Multi-INT Coverage &amp; Collection Gap Overlay</description>

    <!-- Style: SATISFIED Collection (Green) -->
    <Style id="style_SATISFIED">
        <IconStyle>
            <color>ff00ff00</color>
            <scale>1.1</scale>
            <Icon><href>http://maps.google.com/mapfiles/kml/shapes/target.png</href></Icon>
        </IconStyle>
        <LineStyle><color>ff00ff00</color><width>2</width></LineStyle>
        <PolyStyle><color>5500ff00</color><fill>1</fill><outline>1</outline></PolyStyle>
    </Style>

    <!-- Style: DEGRADED_PARTIAL Collection (Amber/Orange) -->
    <Style id="style_DEGRADED_PARTIAL">
        <IconStyle>
            <color>ff00a5ff</color>
            <scale>1.2</scale>
            <Icon><href>http://maps.google.com/mapfiles/kml/shapes/caution.png</href></Icon>
        </IconStyle>
        <LineStyle><color>ff00a5ff</color><width>2.5</width></LineStyle>
        <PolyStyle><color>6600a5ff</color><fill>1</fill><outline>1</outline></PolyStyle>
    </Style>

    <!-- Style: CRITICAL COLLECTION GAP (Red) -->
    <Style id="style_GAP">
        <IconStyle>
            <color>ff0000ff</color>
            <scale>1.3</scale>
            <Icon><href>http://maps.google.com/mapfiles/kml/shapes/forbidden.png</href></Icon>
        </IconStyle>
        <LineStyle><color>ff0000ff</color><width>3</width></LineStyle>
        <PolyStyle><color>770000ff</color><fill>1</fill><outline>1</outline></PolyStyle>
    </Style>
"""

    placemarks = []
    for _, row in df.iterrows():
        status = row["Collection_Status"]
        if status == "SATISFIED":
            style_id = "style_SATISFIED"
            extrude_alt = 800.0
        elif status == "DEGRADED_PARTIAL":
            style_id = "style_DEGRADED_PARTIAL"
            extrude_alt = 1400.0
        else:
            style_id = "style_GAP"
            extrude_alt = 2200.0  # Higher 3D extrusion visually emphasizes critical gaps

        poly_ring = generate_octagon_polygon_coords(
            row["Lat_WGS84"], row["Lon_WGS84"], radius_km=3.0, altitude_m=extrude_alt
        )

        latency_str = f"{row['Latency_Sec']:.0f} sec" if row["Latency_Sec"] >= 0 else "N/A (No Downlink)"
        niirs_str = f"{row['niirs_or_quality']:.1f}" if row["niirs_or_quality"] > 0 else "N/A"

        pm_xml = f"""
    <Placemark>
        <name>[P{row['Priority_Tier']}] {row['Target_ID']} - {row['Sensor_Type']} ({status})</name>
        <styleUrl>#{style_id}</styleUrl>
        <description><![CDATA[
            <h3>{row['Target_Name']} ({row['Target_ID']})</h3>
            <table border="1" cellpadding="4" cellspacing="0" style="border-collapse:collapse; font-family:sans-serif; font-size:12px;">
                <tr><td><b>Collection Requirement ID</b></td><td>{row['CR_ID']}</td></tr>
                <tr><td><b>Priority Tier</b></td><td>Tier {row['Priority_Tier']}</td></tr>
                <tr><td><b>Required Sensor Ontology</b></td><td>{row['Sensor_Type']}</td></tr>
                <tr><td><b>Collection Status</b></td><td><b>{status}</b></td></tr>
                <tr><td><b>Allocated Platform / Callsign</b></td><td>{row['Platform_Type']} ({row['Callsign']})</td></tr>
                <tr><td><b>PED Dissemination Latency</b></td><td>{latency_str}</td></tr>
                <tr><td><b>NIIRS / Signal Quality</b></td><td>{niirs_str}</td></tr>
                <tr><td><b>Normalized WGS84 Coords</b></td><td>{row['Lat_WGS84']:.5f}, {row['Lon_WGS84']:.5f} (from {row['Source_Coord_Format']})</td></tr>
                <tr><td><b>LTIOV (UTC)</b></td><td>{row['LTIOV_UTC']}</td></tr>
                <tr><td><b>Operator Remarks</b></td><td>{row['Raw_Operator_Remarks']}</td></tr>
            </table>
        ]]></description>
        <MultiGeometry>
            <Point>
                <altitudeMode>clampToGround</altitudeMode>
                <coordinates>{row['Lon_WGS84']:.6f},{row['Lat_WGS84']:.6f},0</coordinates>
            </Point>
            <Polygon>
                <extrude>1</extrude>
                <altitudeMode>relativeToGround</altitudeMode>
                <outerBoundaryIs>
                    <LinearRing>
                        <coordinates>{poly_ring}</coordinates>
                    </LinearRing>
                </outerBoundaryIs>
            </Polygon>
        </MultiGeometry>
    </Placemark>"""
        placemarks.append(pm_xml)

    kml_footer = "\n</Document>\n</kml>\n"

    with open(KML_OUT, mode="w", encoding="utf-8") as f:
        f.write(kml_header + "\n".join(placemarks) + kml_footer)

    print(f"[+] Exported 3D Google Earth Pro KML overlay: {KML_OUT}")


if __name__ == "__main__":
    export_csv_and_kml()
# Unclassified Multi-INT TCPED & Collection Management ETL Pipeline (`isr-tcped-etl-pipeline`)

![Python 3.10+](https://img.shields.io/badge/Python-3.10%2B-blue.svg)
![SQLite3](https://img.shields.io/badge/Database-SQLite3-lightgrey.svg)
![Geospatial](https://img.shields.io/badge/GIS-OGC%20KML%202.2%20%7C%20WGS84-green.svg)
![Excel VBA](https://img.shields.io/badge/Excel_VBA-J2_Briefing_Automation-217346.svg)
![Classification](https://img.shields.io/badge/Classification-100%25%20UNCLASSIFIED%20%2F%20SYNTHETIC-brightgreen.svg)

## Overview
`isr-tcped-etl-pipeline` is an end-to-end Python and SQL data engineering architecture designed to ingest, cleanse, normalize, and analyze disparate theater Intelligence, Surveillance, and Reconnaissance (ISR) Collection Management datasets. 

Modeled conceptually on the data structures and interoperability challenges encountered across theater collection management databases (such as GIMS, PRISM, and ROME exports) and airborne Processing, Exploitation, and Dissemination (PED) logs, this pipeline automates **Collection Deck Normalization**, **Multi-Format Geodesy Conversion**, and **Theater Collection Gap Analysis**.

> **OPSEC & Classification Notice:** All data, target nominations, platform allocations, coordinates, and sensor parameters in this repository are **100% synthetic and unclassified**, mathematically generated to demonstrate software architecture and ISR data modeling techniques.

---

## Key Capabilities & Engineering Highlights

1. **Multi-Format Coordinate Cleansing & Geodesy Engine (`DD`, `DMS`, `MGRS` $\rightarrow$ `WGS84`):**
   - Automatically detects and converts heterogeneous coordinate strings—including standard Decimal Degrees, Degrees-Minutes-Seconds (`34°33'04.3"N 117°52'55.6"W`), and Military Grid Reference System (`11S NV 45200 18300`)—into standardized `WGS84` Decimal Degrees using an embedded inverse Transverse Mercator (`UTM`) mathematical projection.
   - Identifies and quarantines out-of-bounds or corrupted tactical feeds into a dedicated relational audit table (`etl_quarantine_audit`) without halting pipeline execution.

2. **TCPED Ontology Mapping & Deduplication:**
   - Resolves duplicate target nominations (`Target_ID`) across subordinate units while preserving highest-priority theater requirements.
   - Parses unstructured free-text operator remarks via deterministic regular expressions into a strict relational ontology (`Sensor_Type` $\in \{\text{EO/IR}, \text{SAR}, \text{SIGINT}, \text{LiDAR}\}$, `Priority_Tier` $\in \{1..4\}$, `Collection_Status`, and `Latency_Sec`).
   - Computes real-time PED dissemination latency ($\Delta t = t_{\text{dissem}} - t_{\text{acq}}$) and flags degraded collections based on National Imagery Interpretability Rating Scale (NIIRS) thresholds and Latest Time Information is of Value (LTIOV) constraints.

3. **Automated SQL Collection Gap Analysis:**
   - Loads cleaned datasets into an embedded relational data warehouse (`sqlite3`) enforced by `CHECK` and uniqueness constraints.
   - Executes analytical SQL queries to isolate **Theater Collection Gaps**—specifically identifying Priority Tier 1 and Tier 2 intelligence requirements with zero platform allocation (`UNALLOCATED_GAP`), weather/maintenance aborts (`SORTIE_ABORT_GAP`), or substandard sensor quality (`DEGRADED_PARTIAL`).

4. **Automated GIS & Executive Dashboard Generation:**
   - Exports a normalized flat file (`arcgis_tableau_tcped_export.csv`) engineered for direct ingestion into **ArcGIS Pro / ArcGIS Portal** and **Tableau** executive dashboards.
   - Programmatically constructs an OGC-compliant **Google Earth Pro (`theater_isr_collection_gaps.kml`)** overlay featuring 3D-extruded, color-coded octagon sensor footprints (Green = Satisfied, Amber = Degraded/High Latency, Red = Unsatisfied Collection Gap) and interactive HTML metadata callouts.

---

## System Architecture

```text
+-------------------------------------------------------------------------+
|                   1. SYNTHETIC RAW FEEDS (data/raw/)                    |
|  - mock_collection_requirements.csv (Dirty DD/DMS/MGRS, Duplicate IDs)  |
|  - mock_sortie_allocation_deck.csv  (ATO/RATO Platform Allocations)     |
|  - mock_sensor_hits.json            (PED Timestamps, NIIRS, Remarks)    |
+------------------------------------+------------------------------------+
                                     |
                                     v
+-------------------------------------------------------------------------+
|                2. ETL & NORMALIZATION (tcped_etl_pipeline.py)           |
|  - Deduplicates Target_IDs by Priority Tier                             |
|  - Converts DMS & MGRS (Inverse UTM) -> WGS84 Decimal Degrees           |
|  - Quarantines corrupt coordinates to [etl_quarantine_audit]            |
|  - Maps unstructured remarks -> [EO/IR, SAR, SIGINT, LiDAR] ontology    |
+------------------------------------+------------------------------------+
                                     |
                                     v
+-------------------------------------------------------------------------+
|          3. SQLITE RELATIONAL WAREHOUSE (isr_tcped_warehouse.db)        |
|  - Table: tcped_normalized_deck (Strict WGS84 & Ontology Constraints)   |
|  - Table: etl_quarantine_audit  (Rejected Feed Traceability)            |
|  - Analytical SQL Queries: Tier 1/2 Collection Gap & Latency Reporting  |
+------------------------------------+------------------------------------+
                                     |
                                     v
+-------------------------------------------------------------------------+
|         4. GEOSPATIAL & BI EXPORTS (export_kml_and_metrics.py)          |
|  - arcgis_tableau_tcped_export.csv (ArcGIS Portal / Tableau Ready)      |
|  - theater_isr_collection_gaps.kml (3D Extruded KML 2.2 Polygons)       |
+-------------------------------------------------------------------------+

---

## Automated J2 Executive Excel Briefing (`vba/J2_TCPED_Briefing_Generator.bas`)
To support staff briefing workflows on locked-down workstations without ActiveX (`Scripting.Dictionary`) dependencies, this repository includes a pure-VBA module:
1. Open a blank **Microsoft Excel** workbook and press **Alt + F11** to launch the Visual Basic Editor.
2. Click **File -> Import File...** and select `vba/J2_TCPED_Briefing_Generator.bas`.
3. Return to Excel, press **Alt + F8**, select **`GenerateJ2BriefingDeck`**, and click **Run** (select `data/processed/arcgis_tableau_tcped_export.csv` if prompted).
4. The macro builds two formatted tabs—**`J2_Executive_Brief`** (KPI summary cards, Multi-INT fulfillment rollup, embedded clustered column chart, and Priority Tier 1/2 Action Deck) and **`Normalized_TCPED_Deck`** (`ListObject` table with status highlighting).
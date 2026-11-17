#!/usr/bin/env python3
"""
run_pipeline.py
Single-command orchestrator for the Unclassified Multi-INT TCPED & Collection Deck ETL Pipeline.
"""

from src.generate_mock_isr_data import main as generate_data
from src.tcped_etl_pipeline import run_etl
from src.export_kml_and_metrics import export_csv_and_kml

if __name__ == "__main__":
    print("================================================================================")
    print("INITIALIZING MULTI-INT TCPED & COLLECTION MANAGEMENT ETL PIPELINE")
    print("================================================================================\n")
    generate_data()
    run_etl()
    export_csv_and_kml()
    print("\n[SUCCESS] Pipeline execution complete. All artifacts written to data/processed/.")
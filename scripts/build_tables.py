#!/usr/bin/env python3
"""
Table Builders - Generate on-demand tables for different audiences.

Usage:
    python scripts/build_tables.py --help
    python scripts/build_tables.py executive --output-dir /path --with-values
    python scripts/build_tables.py qa --output-dir /path
"""

import argparse
import json
import pandas as pd
from pathlib import Path
from datetime import datetime
from typing import Optional, Dict, Any
import sys

# Direct import from same directory
from metrics_catalog import MetricsCatalog, MetricType


def load_latest_values(db_path: Optional[Path] = None, json_path: Optional[Path] = None) -> Dict[str, Any]:
    """Load latest metric values from JSON or SQLite."""
    if json_path and Path(json_path).exists():
        try:
            data = json.load(open(json_path))
            metrics = data.get('metrics', {})
            
            # Handle both dict and list formats
            if isinstance(metrics, dict):
                # Dict format: { metric_key: { value_num: ..., ... } }
                return {key: val.get('value_num') for key, val in metrics.items() if isinstance(val, dict) and 'value_num' in val}
            else:
                # List format: [{ metric_key: ..., value_num: ..., ... }]
                return {m['metric_key']: m['value_num'] for m in metrics if isinstance(m, dict) and 'metric_key' in m}
        except Exception as e:
            print(f"[warn] Could not load values from {json_path}: {e}")
    return {}


def build_executive_dashboard(catalog: MetricsCatalog, values: Optional[Dict] = None, output_path: Optional[Path] = None) -> pd.DataFrame:
    """
    Build executive steering committee dashboard.
    All summary metrics across all domains, with traffic light alerts.
    """
    table = catalog.build_table(metric_type=MetricType.SUMMARY, values=values)
    df = pd.DataFrame(table)
    
    # Add alert level based on thresholds
    def get_alert(row):
        val = row.get('value')
        amber = row.get('threshold_amber')
        red = row.get('threshold_red')
        inverted = row.get('threshold_inverted', False)
        
        if val is None or amber is None or red is None:
            return 'na'
        
        if inverted:  # Higher is better
            if val >= amber:
                return 'green'
            elif val >= red:
                return 'amber'
            else:
                return 'red'
        else:  # Lower is better
            if val <= amber:
                return 'green'
            elif val <= red:
                return 'amber'
            else:
                return 'red'
    
    df['alert_level'] = df.apply(get_alert, axis=1)
    
    # Reorder columns for readability
    cols = ['domain', 'metric_label', 'value', 'unit', 'alert_level', 'threshold_amber', 'threshold_red', 'audience']
    df = df[[c for c in cols if c in df.columns]]
    
    # Sort by domain, then alert level (red first)
    alert_order = {'red': 0, 'amber': 1, 'green': 2, 'na': 3}
    df['alert_sort'] = df['alert_level'].map(alert_order)
    df = df.sort_values(['domain', 'alert_sort']).drop('alert_sort', axis=1)
    
    if output_path:
        df.to_excel(output_path, sheet_name='Executive Dashboard', index=False)
        print(f"[ok] Executive dashboard: {output_path}")
    
    return df


def build_qa_report(catalog: MetricsCatalog, values: Optional[Dict] = None, output_path: Optional[Path] = None) -> pd.DataFrame:
    """
    Build QA manager weekly report.
    All testing metrics (summary + detail) with thresholds.
    """
    table = catalog.build_table(domain='testing', values=values)
    df = pd.DataFrame(table)
    
    # Sort by type (summary first) then by metric
    type_order = {'summary': 0, 'detail': 1}
    df['type_sort'] = df['type'].map(type_order)
    df = df.sort_values('type_sort').drop('type_sort', axis=1)
    
    if output_path:
        with pd.ExcelWriter(output_path) as writer:
            df.to_excel(writer, sheet_name='Testing Metrics', index=False)
            
            # Add a summary sheet
            summary = df[df['type'] == 'summary'][['metric_label', 'value', 'unit']]
            summary.to_excel(writer, sheet_name='Summary', index=False)
        
        print(f"[ok] QA report: {output_path}")
    
    return df


def build_pmo_standup(catalog: MetricsCatalog, values: Optional[Dict] = None, output_path: Optional[Path] = None) -> pd.DataFrame:
    """
    Build PMO daily standup dashboard.
    All detail metrics that need attention today.
    """
    table = catalog.build_table(metric_type=MetricType.DETAIL, values=values)
    df = pd.DataFrame(table)
    
    # Filter to metrics with alert thresholds (actionable items)
    df = df[df['threshold_amber'].notna()]
    
    # Sort by domain
    df = df.sort_values('domain')
    
    if output_path:
        df.to_excel(output_path, sheet_name='Daily Standup', index=False)
        print(f"[ok] PMO standup dashboard: {output_path}")
    
    return df


def build_by_domain_workbook(catalog: MetricsCatalog, values: Optional[Dict] = None, output_path: Optional[Path] = None) -> None:
    """
    Build multi-sheet workbook with one sheet per domain.
    """
    domains = set(m.domain for m in catalog.metrics.values())
    
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        # Add dashboard sheet
        exec_df = pd.DataFrame(catalog.build_table(metric_type=MetricType.SUMMARY, values=values))
        exec_df.to_excel(writer, sheet_name='Dashboard', index=False)
        
        # Add sheets for each domain
        for domain in sorted(domains):
            table = catalog.build_table(domain=domain, values=values)
            df = pd.DataFrame(table)
            sheet_name = domain.replace('_', ' ').title()[:31]  # Excel limit
            df.to_excel(writer, sheet_name=sheet_name, index=False)
    
    print(f"[ok] Multi-domain workbook: {output_path}")


def build_json_exports(catalog: MetricsCatalog, values: Optional[Dict] = None, output_dir: Path = Path('.')) -> None:
    """
    Export tables as JSON for API/programmatic access.
    """
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    
    # Catalog definition
    catalog_json = output_dir / 'metrics_catalog_definition.json'
    catalog.to_json(catalog_json)
    print(f"[ok] Metrics catalog definition: {catalog_json}")
    
    # Executive dashboard
    exec_table = catalog.build_table(metric_type=MetricType.SUMMARY, values=values)
    exec_json = output_dir / 'executive_dashboard.json'
    with open(exec_json, 'w') as f:
        json.dump({
            'as_of_date': datetime.now().isoformat(),
            'metrics': exec_table
        }, f, indent=2)
    print(f"[ok] Executive dashboard JSON: {exec_json}")
    
    # All detail metrics
    detail_table = catalog.build_table(metric_type=MetricType.DETAIL, values=values)
    detail_json = output_dir / 'detail_metrics.json'
    with open(detail_json, 'w') as f:
        json.dump({
            'as_of_date': datetime.now().isoformat(),
            'metrics': detail_table
        }, f, indent=2)
    print(f"[ok] Detail metrics JSON: {detail_json}")


def main():
    parser = argparse.ArgumentParser(description='Build on-demand metric tables for different audiences')
    parser.add_argument(
        'table_type',
        choices=['executive', 'qa', 'pmo', 'all-domains'],
        help='Type of table to build'
    )
    parser.add_argument('--output-dir', required=True, help='Output directory')
    parser.add_argument('--values-json', help='Path to JSON file with metric values (from pmo_metrics_latest.json)')
    parser.add_argument('--with-values', action='store_true', help='Load values from default location if available')
    
    args = parser.parse_args()
    
    catalog = MetricsCatalog()
    output_dir = Path(args.output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    
    # Load values if requested
    values = {}
    if args.with_values or args.values_json:
        json_path = args.values_json or output_dir / 'pmo_metrics_latest.json'
        values = load_latest_values(json_path=json_path)
    
    # Build requested table
    if args.table_type == 'executive':
        output_path = output_dir / f'Executive_Dashboard_{datetime.now().strftime("%Y%m%d")}.xlsx'
        build_executive_dashboard(catalog, values, output_path)
    
    elif args.table_type == 'qa':
        output_path = output_dir / f'QA_Report_{datetime.now().strftime("%Y%m%d")}.xlsx'
        build_qa_report(catalog, values, output_path)
    
    elif args.table_type == 'pmo':
        output_path = output_dir / f'PMO_Standup_{datetime.now().strftime("%Y%m%d")}.xlsx'
        build_pmo_standup(catalog, values, output_path)
    
    elif args.table_type == 'all-domains':
        output_path = output_dir / f'All_Metrics_{datetime.now().strftime("%Y%m%d")}.xlsx'
        build_by_domain_workbook(catalog, values, output_path)
    
    # Always generate JSON exports
    build_json_exports(catalog, values, output_dir)
    
    print()
    print(f"Generated tables in: {output_dir}")


if __name__ == '__main__':
    main()

#!/usr/bin/env python3
"""Quick test of metrics catalog export."""

from scripts.metrics_catalog import MetricsCatalog, MetricType

catalog = MetricsCatalog()

# Export catalog to JSON
output_path = r'C:\Users\jsw73\OneDrive - AMETEK Inc\Ametek SAP S4 Daily Report\metrics_catalog.json'
catalog.to_json(output_path)
print(f'[ok] Exported metrics catalog to: {output_path}')

# Show example: Build Executive Dashboard (Summary metrics only)
print()
print('EXECUTIVE DASHBOARD - Summary Metrics Only:')
print('-'*70)
summaries = catalog.query(metric_type=MetricType.SUMMARY)
table = catalog.build_table(metric_type=MetricType.SUMMARY)

for row in table[:5]:
    print(f"  {row['domain']:<15} {row['metric_key']:<30} {row['metric_label']}")

print(f'  ... and {len(table) - 5} more summary metrics\n')

# Show example: Build Domain-Specific Table
print('PROJECT PLAN DETAIL METRICS - Team Lead Report:')
print('-'*70)
detail_table = catalog.build_table(domain='project_plan', metric_type=MetricType.DETAIL)
for row in detail_table:
    thresh = f"(A:{row['threshold_amber']}, R:{row['threshold_red']})" if row['threshold_amber'] else '(none)'
    print(f"  {row['metric_key']:<35} {thresh}")

# Show example: Build Testing Summary
print()
print('TESTING SUMMARY METRICS:')
print('-'*70)
test_table = catalog.build_table(domain='testing', metric_type=MetricType.SUMMARY)
for row in test_table:
    print(f"  {row['metric_key']:<35} {row['metric_label']}")

print()
print(f'Total metrics in catalog: {len(catalog.metrics)}')
print(f'  Summary: {len(catalog.query(metric_type=MetricType.SUMMARY))}')
print(f'  Detail:  {len(catalog.query(metric_type=MetricType.DETAIL))}')

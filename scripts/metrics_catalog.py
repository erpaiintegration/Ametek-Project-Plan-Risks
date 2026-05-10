#!/usr/bin/env python3
"""
Metrics Catalog - Structured registry of all PMO metrics organized by domain and type.

Defines:
- All 50+ metrics with metadata (domain, category, type, thresholds, formula)
- Summary vs Detail classification
- Functions to query and build tables on-demand
- One source of truth for metric definitions callable by any program

Usage:
    from metrics_catalog import MetricsCatalog
    
    catalog = MetricsCatalog()
    
    # Query by domain + type
    schedule_summaries = catalog.query(domain='project_plan', metric_type='summary')
    testing_details = catalog.query(domain='testing', metric_type='detail')
    
    # Build table for specific use case
    exec_dashboard = catalog.build_table(domain='project_plan', metric_type='summary')
    team_standup = catalog.build_table(domain=None, metric_type='summary')  # all domains
"""

import json
from dataclasses import dataclass, asdict
from typing import Optional, List, Dict, Any
from enum import Enum
from pathlib import Path


class MetricType(Enum):
    """Metric classification: summary (KPI) or detail (drill-down)."""
    SUMMARY = "summary"
    DETAIL = "detail"


class AlertLevel(Enum):
    """Alert severity levels."""
    GREEN = "green"
    AMBER = "amber"
    RED = "red"
    NA = "na"


@dataclass
class Threshold:
    """Alert thresholds for a metric."""
    amber: Optional[float] = None
    red: Optional[float] = None
    inverted: bool = False  # Higher = better (e.g., pass rate, completion %)


@dataclass
class MetricDefinition:
    """Single metric definition with all metadata."""
    
    metric_key: str  # Unique identifier (e.g., "tasks_total")
    metric_label: str  # Human-readable label (e.g., "Total project tasks")
    domain: str  # Domain (project_plan, testing, risks, issues, actions, pmo_health)
    metric_type: MetricType  # SUMMARY or DETAIL
    unit: str  # Unit of measurement (count, %, text, etc.)
    threshold: Optional[Threshold] = None
    
    # Computation guidance
    description: str = ""  # Business definition
    formula: str = ""  # How to compute (e.g., "COUNT(tasks WHERE status != 'Complete')")
    audience: str = ""  # Who cares (PMO, Dev Lead, QA Manager, etc.)
    
    def to_dict(self) -> Dict[str, Any]:
        """Serialize to dict."""
        return {
            'metric_key': self.metric_key,
            'metric_label': self.metric_label,
            'domain': self.domain,
            'metric_type': self.metric_type.value,
            'unit': self.unit,
            'description': self.description,
            'formula': self.formula,
            'audience': self.audience,
            'threshold': {
                'amber': self.threshold.amber,
                'red': self.threshold.red,
                'inverted': self.threshold.inverted,
            } if self.threshold else None,
        }


class MetricsCatalog:
    """Registry of all PMO metrics organized by domain and type."""
    
    def __init__(self):
        self.metrics: Dict[str, MetricDefinition] = {}
        self._register_all()
    
    def _register_all(self):
        """Register all 50+ PMO metrics."""
        
        # ==================== PROJECT PLAN / SCHEDULE ====================
        
        # Summary metrics
        self._register(MetricDefinition(
            metric_key="tasks_total",
            metric_label="Total project tasks",
            domain="project_plan",
            metric_type=MetricType.SUMMARY,
            unit="count",
            description="All tasks in the project plan",
            audience="PMO, Executive",
        ))
        
        self._register(MetricDefinition(
            metric_key="tasks_complete",
            metric_label="Completed tasks",
            domain="project_plan",
            metric_type=MetricType.SUMMARY,
            unit="count",
            description="Tasks marked complete or 100% done",
            audience="PMO",
        ))
        
        self._register(MetricDefinition(
            metric_key="tasks_open",
            metric_label="Open tasks",
            domain="project_plan",
            metric_type=MetricType.SUMMARY,
            unit="count",
            description="Tasks not yet complete (future, not started, in progress, at risk)",
            audience="Team Leads, PMO",
        ))
        
        self._register(MetricDefinition(
            metric_key="schedule_pct_complete",
            metric_label="Overall schedule completion %",
            domain="project_plan",
            metric_type=MetricType.SUMMARY,
            unit="%",
            description="Total tasks_complete / tasks_total",
            formula="SUM(% Complete) / COUNT(tasks)",
            threshold=Threshold(amber=60, red=40, inverted=True),
            audience="Executive, PMO",
        ))
        
        self._register(MetricDefinition(
            metric_key="milestones_total",
            metric_label="Total milestones",
            domain="project_plan",
            metric_type=MetricType.SUMMARY,
            unit="count",
            audience="PMO, Sponsor",
        ))
        
        self._register(MetricDefinition(
            metric_key="milestones_complete",
            metric_label="Completed milestones",
            domain="project_plan",
            metric_type=MetricType.SUMMARY,
            unit="count",
            audience="PMO",
        ))
        
        self._register(MetricDefinition(
            metric_key="critical_tasks_total",
            metric_label="Critical path tasks",
            domain="project_plan",
            metric_type=MetricType.SUMMARY,
            unit="count",
            description="Tasks on the critical path",
            audience="PMO, Scheduler",
        ))
        
        # Detail metrics
        self._register(MetricDefinition(
            metric_key="tasks_late_vs_baseline",
            metric_label="Tasks late vs baseline finish",
            domain="project_plan",
            metric_type=MetricType.DETAIL,
            unit="count",
            description="Open tasks with finish_date > baseline_finish",
            formula="COUNT(tasks WHERE status != 'Complete' AND finish_date > baseline_finish)",
            threshold=Threshold(amber=10, red=25),
            audience="PMO, Scheduler",
        ))
        
        self._register(MetricDefinition(
            metric_key="critical_tasks_late",
            metric_label="Critical path tasks behind schedule",
            domain="project_plan",
            metric_type=MetricType.DETAIL,
            unit="count",
            threshold=Threshold(amber=1, red=3),
            audience="PMO, Scheduler",
        ))
        
        self._register(MetricDefinition(
            metric_key="tasks_due_7_days",
            metric_label="Tasks due within 7 days",
            domain="project_plan",
            metric_type=MetricType.DETAIL,
            unit="count",
            formula="COUNT(tasks WHERE finish_date BETWEEN today AND today+7 AND status != 'Complete')",
            audience="Team Leads, Daily Standup",
        ))
        
        self._register(MetricDefinition(
            metric_key="tasks_due_14_days",
            metric_label="Tasks due within 14 days",
            domain="project_plan",
            metric_type=MetricType.DETAIL,
            unit="count",
            audience="Team Leads",
        ))
        
        self._register(MetricDefinition(
            metric_key="tasks_due_7_days_no_progress",
            metric_label="Tasks due in 7 days with 0% progress",
            domain="project_plan",
            metric_type=MetricType.DETAIL,
            unit="count",
            description="High-risk: due soon but not started",
            formula="COUNT(tasks WHERE finish_date BETWEEN today AND today+7 AND % Complete = 0)",
            threshold=Threshold(amber=3, red=8),
            audience="PMO, Team Leads",
        ))
        
        self._register(MetricDefinition(
            metric_key="milestones_overdue",
            metric_label="Overdue milestones",
            domain="project_plan",
            metric_type=MetricType.DETAIL,
            unit="count",
            threshold=Threshold(amber=1, red=2),
            audience="PMO, Sponsor",
        ))
        
        # ==================== TESTING ====================
        
        # Summary metrics
        self._register(MetricDefinition(
            metric_key="tests_total",
            metric_label="Total test cases",
            domain="testing",
            metric_type=MetricType.SUMMARY,
            unit="count",
            audience="QA Manager",
        ))
        
        self._register(MetricDefinition(
            metric_key="tests_executable",
            metric_label="Executable test cases",
            domain="testing",
            metric_type=MetricType.SUMMARY,
            unit="count",
            description="Tests that can be run (excludes N/A)",
            audience="QA Manager",
        ))
        
        self._register(MetricDefinition(
            metric_key="tests_passed",
            metric_label="Passed tests",
            domain="testing",
            metric_type=MetricType.SUMMARY,
            unit="count",
            audience="QA Manager",
        ))
        
        self._register(MetricDefinition(
            metric_key="tests_pass_rate_pct",
            metric_label="Test pass rate %",
            domain="testing",
            metric_type=MetricType.SUMMARY,
            unit="%",
            description="passed / executable",
            formula="passed_count / executable_count * 100",
            threshold=Threshold(amber=85, red=70, inverted=True),
            audience="QA Manager, Executive",
        ))
        
        # Detail metrics
        self._register(MetricDefinition(
            metric_key="tests_failed",
            metric_label="Failed tests",
            domain="testing",
            metric_type=MetricType.DETAIL,
            unit="count",
            threshold=Threshold(amber=5, red=15),
            audience="QA Manager",
        ))
        
        self._register(MetricDefinition(
            metric_key="tests_blocked",
            metric_label="Blocked tests",
            domain="testing",
            metric_type=MetricType.DETAIL,
            unit="count",
            threshold=Threshold(amber=3, red=10),
            audience="QA Manager, Dev Lead",
        ))
        
        self._register(MetricDefinition(
            metric_key="tests_not_yet_run",
            metric_label="Not yet run",
            domain="testing",
            metric_type=MetricType.DETAIL,
            unit="count",
            audience="QA Manager",
        ))
        
        self._register(MetricDefinition(
            metric_key="tests_incomplete",
            metric_label="Incomplete tests",
            domain="testing",
            metric_type=MetricType.DETAIL,
            unit="count",
            audience="QA Manager",
        ))
        
        self._register(MetricDefinition(
            metric_key="tests_not_applicable",
            metric_label="Not applicable tests",
            domain="testing",
            metric_type=MetricType.DETAIL,
            unit="count",
            description="Skipped / excluded from scope",
            audience="QA Manager",
        ))
        
        self._register(MetricDefinition(
            metric_key="tests_fail_rate_pct",
            metric_label="Test fail rate %",
            domain="testing",
            metric_type=MetricType.DETAIL,
            unit="%",
            description="failed / executable",
            audience="QA Manager",
        ))
        
        self._register(MetricDefinition(
            metric_key="tests_with_open_defects",
            metric_label="Tests with open defects",
            domain="testing",
            metric_type=MetricType.DETAIL,
            unit="count",
            audience="QA Manager, Dev Lead",
        ))
        
        # ==================== RISKS ====================
        
        # Summary metrics
        self._register(MetricDefinition(
            metric_key="risks_total",
            metric_label="Total identified risks",
            domain="risks",
            metric_type=MetricType.SUMMARY,
            unit="count",
            audience="PMO, Risk Manager",
        ))
        
        self._register(MetricDefinition(
            metric_key="risks_open",
            metric_label="Open risks",
            domain="risks",
            metric_type=MetricType.SUMMARY,
            unit="count",
            description="Active risks not yet closed or mitigated",
            audience="PMO, Risk Manager",
        ))
        
        self._register(MetricDefinition(
            metric_key="risks_closed",
            metric_label="Closed risks",
            domain="risks",
            metric_type=MetricType.SUMMARY,
            unit="count",
            audience="PMO",
        ))
        
        # Detail metrics
        self._register(MetricDefinition(
            metric_key="risks_open_high_exposure",
            metric_label="Open high-exposure risks",
            domain="risks",
            metric_type=MetricType.DETAIL,
            unit="count",
            description="Open risks with exposure score >= 9",
            threshold=Threshold(amber=1, red=3),
            audience="PMO, Sponsor",
        ))
        
        self._register(MetricDefinition(
            metric_key="risks_open_medium_exposure",
            metric_label="Open medium-exposure risks",
            domain="risks",
            metric_type=MetricType.DETAIL,
            unit="count",
            description="Open risks with exposure 4-8",
            threshold=Threshold(amber=5, red=10),
            audience="PMO",
        ))
        
        self._register(MetricDefinition(
            metric_key="risks_open_low_exposure",
            metric_label="Open low-exposure risks",
            domain="risks",
            metric_type=MetricType.DETAIL,
            unit="count",
            description="Open risks with exposure < 4",
            audience="PMO",
        ))
        
        self._register(MetricDefinition(
            metric_key="risks_overdue",
            metric_label="Overdue risks (response due)",
            domain="risks",
            metric_type=MetricType.DETAIL,
            unit="count",
            threshold=Threshold(amber=2, red=5),
            audience="PMO, Risk Manager",
        ))
        
        self._register(MetricDefinition(
            metric_key="risks_avg_open_exposure",
            metric_label="Average exposure (open risks)",
            domain="risks",
            metric_type=MetricType.DETAIL,
            unit="count",
            description="Mean of exposure scores for all open risks",
            audience="PMO",
        ))
        
        # ==================== ISSUES ====================
        
        # Summary metrics
        self._register(MetricDefinition(
            metric_key="issues_total",
            metric_label="Total issues",
            domain="issues",
            metric_type=MetricType.SUMMARY,
            unit="count",
            audience="PMO, Issue Manager",
        ))
        
        self._register(MetricDefinition(
            metric_key="issues_open",
            metric_label="Open issues",
            domain="issues",
            metric_type=MetricType.SUMMARY,
            unit="count",
            threshold=Threshold(amber=10, red=20),
            audience="PMO, Team Leads",
        ))
        
        self._register(MetricDefinition(
            metric_key="issues_closed",
            metric_label="Closed issues",
            domain="issues",
            metric_type=MetricType.SUMMARY,
            unit="count",
            audience="PMO",
        ))
        
        # Detail metrics
        self._register(MetricDefinition(
            metric_key="issues_escalated",
            metric_label="Escalated issues",
            domain="issues",
            metric_type=MetricType.DETAIL,
            unit="count",
            threshold=Threshold(amber=1, red=3),
            audience="PMO, Leadership",
        ))
        
        self._register(MetricDefinition(
            metric_key="issues_critical_high",
            metric_label="Critical or High priority issues",
            domain="issues",
            metric_type=MetricType.DETAIL,
            unit="count",
            threshold=Threshold(amber=2, red=5),
            audience="PMO, Leadership",
        ))
        
        self._register(MetricDefinition(
            metric_key="issues_overdue",
            metric_label="Overdue issues",
            domain="issues",
            metric_type=MetricType.DETAIL,
            unit="count",
            threshold=Threshold(amber=3, red=8),
            audience="PMO, Issue Manager",
        ))
        
        # ==================== ACTIONS ====================
        
        # Summary metrics
        self._register(MetricDefinition(
            metric_key="actions_total",
            metric_label="Total action items",
            domain="actions",
            metric_type=MetricType.SUMMARY,
            unit="count",
            audience="PMO, Issue Manager",
        ))
        
        self._register(MetricDefinition(
            metric_key="actions_open",
            metric_label="Open action items",
            domain="actions",
            metric_type=MetricType.SUMMARY,
            unit="count",
            audience="PMO, Team Leads",
        ))
        
        self._register(MetricDefinition(
            metric_key="actions_closed",
            metric_label="Closed actions",
            domain="actions",
            metric_type=MetricType.SUMMARY,
            unit="count",
            audience="PMO",
        ))
        
        # Detail metrics
        self._register(MetricDefinition(
            metric_key="actions_in_progress",
            metric_label="Actions in progress",
            domain="actions",
            metric_type=MetricType.DETAIL,
            unit="count",
            audience="PMO",
        ))
        
        self._register(MetricDefinition(
            metric_key="actions_overdue",
            metric_label="Overdue actions",
            domain="actions",
            metric_type=MetricType.DETAIL,
            unit="count",
            threshold=Threshold(amber=5, red=15),
            audience="PMO, Team Leads",
        ))
        
        self._register(MetricDefinition(
            metric_key="actions_due_7_days",
            metric_label="Actions due within 7 days",
            domain="actions",
            metric_type=MetricType.DETAIL,
            unit="count",
            audience="Team Leads, Daily Standup",
        ))
        
        self._register(MetricDefinition(
            metric_key="actions_critical_high",
            metric_label="Critical or High priority open actions",
            domain="actions",
            metric_type=MetricType.DETAIL,
            unit="count",
            threshold=Threshold(amber=10, red=20),
            audience="PMO, Leadership",
        ))
        
        # ==================== PMO HEALTH / CROSS-DOMAIN ====================
        
        # Summary metrics
        self._register(MetricDefinition(
            metric_key="workstreams_in_plan",
            metric_label="Workstreams in plan",
            domain="pmo_health",
            metric_type=MetricType.SUMMARY,
            unit="count",
            audience="PMO, Sponsor",
        ))
        
        self._register(MetricDefinition(
            metric_key="workstreams_with_open_issues",
            metric_label="Workstreams with open issues",
            domain="pmo_health",
            metric_type=MetricType.SUMMARY,
            unit="count",
            audience="PMO",
        ))
        
        # Detail metrics
        self._register(MetricDefinition(
            metric_key="workstreams_with_late_tasks",
            metric_label="Workstreams with late tasks",
            domain="pmo_health",
            metric_type=MetricType.DETAIL,
            unit="count",
            description="Workstreams with >= 1 task behind baseline",
            threshold=Threshold(amber=3, red=7),
            audience="PMO, Leadership",
        ))
        
        self._register(MetricDefinition(
            metric_key="workstreams_with_overdue_actions",
            metric_label="Workstreams with overdue actions",
            domain="pmo_health",
            metric_type=MetricType.DETAIL,
            unit="count",
            threshold=Threshold(amber=2, red=5),
            audience="PMO, Leadership",
        ))
        
        self._register(MetricDefinition(
            metric_key="report_run_id",
            metric_label="Latest report run ID",
            domain="pmo_health",
            metric_type=MetricType.SUMMARY,
            unit="text",
            description="Timestamp of most recent metric computation (YYYYMMDDTHHmmss)",
            audience="System",
        ))
    
    def _register(self, metric: MetricDefinition):
        """Register a metric in the catalog."""
        self.metrics[metric.metric_key] = metric
    
    def query(
        self,
        domain: Optional[str] = None,
        metric_type: Optional[MetricType] = None,
    ) -> List[MetricDefinition]:
        """
        Query metrics by domain and/or type.
        
        Args:
            domain: Filter by domain (e.g., 'project_plan', 'testing', etc.) or None for all
            metric_type: Filter by SUMMARY, DETAIL, or None for all
        
        Returns:
            List of matching metrics
        """
        results = list(self.metrics.values())
        
        if domain:
            results = [m for m in results if m.domain == domain]
        
        if metric_type:
            results = [m for m in results if m.metric_type == metric_type]
        
        return sorted(results, key=lambda m: (m.domain, m.metric_type.value, m.metric_key))
    
    def get(self, metric_key: str) -> Optional[MetricDefinition]:
        """Get a single metric by key."""
        return self.metrics.get(metric_key)
    
    def build_table(
        self,
        domain: Optional[str] = None,
        metric_type: Optional[MetricType] = None,
        values: Optional[Dict[str, Any]] = None,
    ) -> List[Dict[str, Any]]:
        """
        Build a table (list of dicts) with metrics and optional values.
        
        Args:
            domain: Filter domain
            metric_type: Filter type
            values: Dict mapping metric_key -> value for display
        
        Returns:
            List of row dicts with metric metadata and values
        """
        metrics = self.query(domain=domain, metric_type=metric_type)
        rows = []
        
        for metric in metrics:
            row = {
                'metric_key': metric.metric_key,
                'metric_label': metric.metric_label,
                'domain': metric.domain,
                'type': metric.metric_type.value,
                'unit': metric.unit,
                'description': metric.description,
                'audience': metric.audience,
                'formula': metric.formula,
                'value': values.get(metric.metric_key) if values else None,
                'threshold_amber': metric.threshold.amber if metric.threshold else None,
                'threshold_red': metric.threshold.red if metric.threshold else None,
                'threshold_inverted': metric.threshold.inverted if metric.threshold else False,
            }
            rows.append(row)
        
        return rows
    
    def to_dict(self) -> Dict[str, Dict[str, Any]]:
        """Export entire catalog as dict."""
        return {
            key: metric.to_dict()
            for key, metric in self.metrics.items()
        }
    
    def to_json(self, path: Optional[Path] = None) -> str:
        """Export catalog as JSON."""
        import json
        result = json.dumps(self.to_dict(), indent=2)
        if path:
            Path(path).write_text(result, encoding='utf-8')
        return result
    
    def summary(self) -> Dict[str, int]:
        """Get summary counts by domain and type."""
        summary = {}
        for domain in set(m.domain for m in self.metrics.values()):
            for mtype in MetricType:
                key = f"{domain}_{mtype.value}"
                count = len(self.query(domain=domain, metric_type=mtype))
                if count > 0:
                    summary[key] = count
        return summary


if __name__ == '__main__':
    # Example usage
    catalog = MetricsCatalog()
    
    print("=" * 80)
    print("PMO METRICS CATALOG")
    print("=" * 80)
    print()
    
    # Show summary
    print("Metric Counts by Domain & Type:")
    print("-" * 40)
    for key, count in sorted(catalog.summary().items()):
        domain, mtype = key.rsplit('_', 1)
        print(f"  {domain:<20} {mtype:<10} : {count:>3}")
    print()
    
    # Show sample queries
    print("Project Plan Summary Metrics:")
    print("-" * 80)
    for metric in catalog.query(domain='project_plan', metric_type=MetricType.SUMMARY):
        print(f"  {metric.metric_key:<30} {metric.metric_label}")
    print()
    
    print("Testing Detail Metrics:")
    print("-" * 80)
    for metric in catalog.query(domain='testing', metric_type=MetricType.DETAIL):
        print(f"  {metric.metric_key:<30} {metric.metric_label}")
    print()
    
    # Show all summary metrics
    print("ALL SUMMARY METRICS (Executive Dashboard):")
    print("-" * 80)
    summaries = catalog.query(metric_type=MetricType.SUMMARY)
    for metric in summaries:
        print(f"  [{metric.domain:<15}] {metric.metric_key:<30} {metric.metric_label}")
    print(f"\nTotal summary metrics: {len(summaries)}")
    print()
    
    # Export catalog
    json_output = catalog.to_json()
    print(f"Total metrics in catalog: {len(catalog.metrics)}")
    print()
    print("Sample JSON export (first 500 chars):")
    print(json_output[:500] + "...")

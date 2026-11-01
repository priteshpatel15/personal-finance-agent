from __future__ import annotations

from datetime import date, datetime
from decimal import Decimal
from html import escape
from pathlib import Path

from personal_budget_plaid_excel.domain.budget import monthly_budget_summary
from personal_budget_plaid_excel.domain.models import MonthlyBudgetSummary
from personal_budget_plaid_excel.infrastructure.workbook import BudgetWorkbook


class HtmlReportService:
    def __init__(self, workbook_path: Path) -> None:
        self.workbook_path = workbook_path

    def build_summary(self, month: date) -> MonthlyBudgetSummary:
        workbook = BudgetWorkbook(self.workbook_path)
        last_synced_at, sync_status = workbook.last_sync_context()
        return monthly_budget_summary(
            workbook.read_budgets(),
            workbook.read_transactions(),
            month,
            last_synced_at=last_synced_at,
            sync_status=sync_status,
        )

    def write_report(self, output_dir: Path, month: date) -> Path:
        output_dir.mkdir(parents=True, exist_ok=True)
        summary = self.build_summary(month)
        stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
        output_path = output_dir / f"budget-report-{summary.month:%Y-%m}-{stamp}.html"
        output_path.write_text(render_html(summary), encoding="utf-8")
        return output_path


def render_html(summary: MonthlyBudgetSummary) -> str:
    rows = "\n".join(_category_row(position) for position in summary.positions)
    top_over = [p for p in summary.positions if p.remaining < 0]
    alerts = "".join(
        f"<li><strong>{escape(p.category)}</strong> is over by {_money(-p.remaining)}.</li>"
        for p in top_over[:8]
    )
    if summary.uncategorized_count:
        alerts += f"<li>{summary.uncategorized_count} transaction(s) need a Selected Category.</li>"
    if not alerts:
        alerts = "<li>No over-budget categories in posted spend.</li>"
    freshness = summary.last_synced_at.isoformat() if summary.last_synced_at else "No successful Plaid sync recorded"
    return f"""<!doctype html>
<html>
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Budget Report {summary.month:%B %Y}</title>
  <style>
    body {{ font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif; margin: 24px; color: #1f2937; }}
    h1 {{ margin-bottom: 4px; }}
    .meta {{ color: #4b5563; margin-top: 0; }}
    .cards {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(180px, 1fr)); gap: 12px; margin: 20px 0; }}
    .card {{ border: 1px solid #d1d5db; border-radius: 8px; padding: 12px; }}
    .label {{ color: #6b7280; font-size: 12px; text-transform: uppercase; letter-spacing: .04em; }}
    .value {{ font-size: 24px; font-weight: 700; margin-top: 4px; }}
    table {{ border-collapse: collapse; width: 100%; margin-top: 20px; }}
    th, td {{ border-bottom: 1px solid #e5e7eb; padding: 8px; text-align: right; }}
    th:first-child, td:first-child, th:nth-child(2), td:nth-child(2) {{ text-align: left; }}
    th {{ background: #1f4e78; color: #fff; }}
    tr.over td {{ background: #fef2f2; }}
    tr.watch td {{ background: #fffbeb; }}
    .alerts {{ border: 1px solid #d1d5db; border-radius: 8px; padding: 12px 16px; }}
  </style>
</head>
<body>
  <h1>Budget Report: {summary.month:%B %Y}</h1>
  <p class="meta">Sync status: {escape(summary.sync_status)} | Freshness: {escape(freshness)}</p>
  <section class="cards">
    <div class="card"><div class="label">Monthly Budget</div><div class="value">{_money(summary.total_budget)}</div></div>
    <div class="card"><div class="label">Posted Spend</div><div class="value">{_money(summary.total_posted_spend)}</div></div>
    <div class="card"><div class="label">Remaining</div><div class="value">{_money(summary.total_remaining)}</div></div>
    <div class="card"><div class="label">Pending Impact</div><div class="value">{_money(summary.total_pending_impact)}</div></div>
  </section>
  <section class="alerts">
    <h2>Attention</h2>
    <ul>{alerts}</ul>
  </section>
  <table>
    <thead>
      <tr>
        <th>Cost Type</th><th>Category</th><th>Budget</th><th>Posted</th><th>Pending</th><th>Remaining</th><th>% Used</th><th>Review</th><th>Status</th>
      </tr>
    </thead>
    <tbody>{rows}</tbody>
  </table>
</body>
</html>
"""


def _category_row(position) -> str:
    css = "over" if position.status == "Over" else "watch" if position.status == "Watch" else ""
    return (
        f'<tr class="{css}"><td>{escape(position.cost_type)}</td>'
        f"<td>{escape(position.category)}</td>"
        f"<td>{_money(position.monthly_budget)}</td>"
        f"<td>{_money(position.posted_spend)}</td>"
        f"<td>{_money(position.pending_impact)}</td>"
        f"<td>{_money(position.remaining)}</td>"
        f"<td>{_percent(position.percent_used)}</td>"
        f"<td>{position.review_needed_count}</td>"
        f"<td>{escape(position.status)}</td></tr>"
    )


def _money(value: Decimal) -> str:
    return f"${value:,.2f}"


def _percent(value: Decimal) -> str:
    return f"{value * Decimal('100'):.1f}%"


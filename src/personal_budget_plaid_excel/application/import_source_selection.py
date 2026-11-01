from __future__ import annotations

from dataclasses import dataclass

from personal_budget_plaid_excel.domain.source_mode import ImportSourceMode


@dataclass(frozen=True)
class MonthlyImportSourceDecision:
    """Public, side-effect-free decision made before monthly data collection."""

    status: str
    source: ImportSourceMode | None
    available_sources: tuple[ImportSourceMode, ...]

    @property
    def data_collection_allowed(self) -> bool:
        return self.source is not None

    def public_payload(self) -> dict[str, object]:
        if self.source is None:
            return {
                "status": "source_choice_required",
                "reason": "monthly_import_source_required",
                "prompt": "Choose an import source before collecting monthly transactions.",
                "choices": [
                    {"id": source.value, "label": source.display_name}
                    for source in self.available_sources
                ],
                "data_collection_started": False,
            }
        return {
            "status": "source_selected",
            "import_source": {
                "id": self.source.value,
                "label": self.source.display_name,
            },
            "data_collection_started": False,
        }


class SelectMonthlyImportSource:
    """Resolve an explicit source without touching files or external systems."""

    def execute(
        self,
        requested_source: str | ImportSourceMode | None,
    ) -> MonthlyImportSourceDecision:
        available = tuple(ImportSourceMode)
        if requested_source is None or (
            isinstance(requested_source, str) and not requested_source.strip()
        ):
            return MonthlyImportSourceDecision(
                status="source_choice_required",
                source=None,
                available_sources=available,
            )

        selected = (
            requested_source
            if isinstance(requested_source, ImportSourceMode)
            else ImportSourceMode.parse(requested_source)
        )
        return MonthlyImportSourceDecision(
            status="source_selected",
            source=selected,
            available_sources=available,
        )


def source_label(source: ImportSourceMode) -> str:
    return source.display_name

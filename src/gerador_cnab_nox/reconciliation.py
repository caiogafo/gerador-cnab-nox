"""Reconciliation of identified credits, using integer cents throughout."""

from __future__ import annotations

import os
from dataclasses import asdict, dataclass
from decimal import Decimal, InvalidOperation
from typing import Literal

from .normalize import digits, mask_document

ReconciliationStatus = Literal["PERFECT_MATCH", "TOLERATED_WARNING", "CRITICAL_ERROR"]


def tolerance_cents() -> int:
    raw = os.environ.get("MAX_TOLERANCE_DIFF_REAIS", "100.00")
    try:
        value = Decimal(raw)
        if not value.is_finite() or value < 0 or value * 100 != (value * 100).to_integral():
            raise ValueError
        return int(value * 100)
    except (ValueError, InvalidOperation) as exc:
        raise ValueError(
            "MAX_TOLERANCE_DIFF_REAIS deve ser decimal não negativo, com até 2 casas"
        ) from exc


def cents(value: Decimal) -> int:
    if not value.is_finite() or value < 0 or value * 100 != (value * 100).to_integral():
        raise ValueError("Valor deve ser não negativo e expresso em centavos inteiros")
    return int(value * 100)


@dataclass(frozen=True)
class CreditorReconciliationResult:
    creditorId: str
    documentNumber: str
    name: str
    pfmiValueCents: int
    analyticValueCents: int
    diffCents: int
    status: ReconciliationStatus
    warningMessage: str | None = None
    errorMessage: str | None = None


@dataclass(frozen=True)
class BatchValidationSummary:
    canGenerateTxt: bool
    totalPfmiCents: int
    totalAnalyticCents: int
    totalDiffCents: int
    hasWarnings: bool
    warningCount: int
    errors: list[CreditorReconciliationResult]
    warnings: list[CreditorReconciliationResult]
    clearedRecords: list[CreditorReconciliationResult]

    def to_dict(self) -> dict:
        return asdict(self)

    def to_audit_dict(self) -> dict:
        """Preserve the application's existing redaction of personal data in logs."""
        result = self.to_dict()
        for category in ("errors", "warnings", "clearedRecords"):
            for record in result[category]:
                record["documentNumber"] = mask_document(record["documentNumber"])
                record["name"] = "[omitido]"
        return result


def reconcile_credit(
    creditor_id: str, document: str, name: str, pfmi: Decimal, analytic: Decimal,
    *, limit_cents: int | None = None,
) -> CreditorReconciliationResult:
    limit = tolerance_cents() if limit_cents is None else limit_cents
    if isinstance(limit, bool) or not isinstance(limit, int) or limit < 0:
        raise ValueError("Tolerância deve ser um inteiro não negativo em centavos")
    pfmi_cents, analytic_cents = cents(pfmi), cents(analytic)
    difference = pfmi_cents - analytic_cents
    status: ReconciliationStatus = (
        "PERFECT_MATCH" if difference == 0
        else "TOLERATED_WARNING" if abs(difference) <= limit
        else "CRITICAL_ERROR"
    )
    message = (
        f"{creditor_id}: {status}; PFMI R$ {pfmi:.2f}; Analítico R$ {analytic:.2f}; "
        f"diferença R$ {Decimal(difference) / 100:+.2f}; "
        f"limite R$ {Decimal(limit) / 100:.2f}. "
    )
    return CreditorReconciliationResult(
        creditor_id, digits(document), name, pfmi_cents, analytic_cents, difference, status,
        message + "Será utilizado o valor da PFMI no TXT."
        if status == "TOLERATED_WARNING" else None,
        message + "DIVERGENCIA_VALOR acima da tolerância; geração bloqueada."
        if status == "CRITICAL_ERROR" else None,
    )


def summarize(
    records: list[CreditorReconciliationResult], *, other_errors: bool = False,
) -> BatchValidationSummary:
    errors = [r for r in records if r.status == "CRITICAL_ERROR"]
    warnings = [r for r in records if r.status == "TOLERATED_WARNING"]
    return BatchValidationSummary(
        not errors and not other_errors,
        sum(r.pfmiValueCents for r in records),
        sum(r.analyticValueCents for r in records),
        sum(r.diffCents for r in records),
        bool(warnings), len(warnings), errors, warnings,
        [r for r in records if r.status == "PERFECT_MATCH"],
    )

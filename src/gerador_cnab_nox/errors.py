from __future__ import annotations


class UserFacingError(Exception):
    """Erro esperado que pode ser apresentado diretamente à usuária."""


class InputFileError(UserFacingError):
    """Entrada ausente, ilegível, corrompida ou estruturalmente incompatível."""


class ValidationError(UserFacingError):
    """Pendências que impedem a geração do TXT."""

    def __init__(self, issues: list[str]):
        self.issues = issues
        preview = "\n".join(f"- {item}" for item in issues[:20])
        suffix = "\n- ..." if len(issues) > 20 else ""
        super().__init__(f"O TXT não foi gerado. Corrija as pendências:\n{preview}{suffix}")

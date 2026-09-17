"""Presentation-only guidance. Original service messages remain available verbatim."""

from __future__ import annotations

import re
from decimal import Decimal


def money(value: Decimal) -> str:
    return f"{value:,.2f}".translate(str.maketrans({",": ".", ".": ","}))


_PENDING_MESSAGES = (
    (
        "CREDITO_NAO_LOCALIZADO",
        "Não encontramos um crédito individual com o mesmo valor. Confira o instrumento.",
    ),
    (
        "COMPOSICAO_BLOQUEADA_SELECAO_MANUAL",
        "A composição não fechou automaticamente. Confira o instrumento e, se estiver "
        "correto, informe a aba/linha do Analítico em COMPOSICAO_SELECAO_MANUAL_ABA/"
        "COMPOSICAO_SELECAO_MANUAL_LINHA.",
    ),
    (
        "MULTIPLOS_CREDITOS_PRINCIPAIS",
        "Encontramos mais de um crédito. Selecione os créditos corretos.",
    ),
    (
        "CONFLITO_BASE_VENCIMENTOS",
        "Há mais de um vencimento possível para este sacado. Confira a Base de Vencimentos.",
    ),
    (
        "SACADO_NAO_LOCALIZADO",
        "Não encontramos o vencimento deste sacado na Base de Vencimentos.",
    ),
    (
        "DOCUMENTOS_CANDIDATOS_CONFLITANTES",
        "Os créditos encontrados têm documentos diferentes. Confira e escolha manualmente.",
    ),
    (
        "DIVERGENCIA_VALOR",
        "O valor do Analítico diverge da PFMI. Dentro da tolerância configurada, o TXT "
        "usa a PFMI com alerta; acima do limite, corrija as fontes e prepare novamente.",
    ),
    ("DIVERGENCIA_NOME", "O nome no sistema é diferente do nome na PFMI. Confira e corrija."),
    ("DOC_CEDENTE_INVALIDO", "O documento do cedente está inválido. Confira e corrija."),
    ("DOC_SACADO_INVALIDO", "O documento do sacado está inválido. Confira e corrija."),
    (
        "EMISSAO_NOVA_CESSAO_AUSENTE",
        "Falta informar a data de emissão da nova cessão para este arquivo PFMI.",
    ),
    (
        "EMISSAO_NOVA_CESSAO_INVALIDA",
        "A data de emissão da nova cessão informada não é válida. Confira o formato dd/mm/aaaa.",
    ),
    (
        "EMISSAO_NOVA_CESSAO_POSTERIOR_LIQUIDACAO",
        "A data de emissão da nova cessão é posterior à liquidação. Corrija antes de continuar.",
    ),
)


def explain_pending(codes: str) -> str:
    """Plain-language message for a pending row/group, from its PENDENCIAS codes.

    An openpyxl round trip reads a blank PENDENCIAS cell back as None (not
    ""), which a group with no other issue can now reach in practice since
    the composition nominal suggestion can clear the row's only pendency.
    """
    codes = codes or ""
    for code, message in _PENDING_MESSAGES:
        if code in codes:
            return message
    return "Este item precisa de conferência manual antes de continuar."


# Uma instrução objetiva por linha, priorizada (nunca a lista inteira de
# flags empilhada) - o oposto de explain_pending acima, que descreve o
# problema; esta diz que campo preencher/aprovar e onde. Deliberadamente uma
# função pura e independente, sem alterar explain_pending nem seus chamadores
# já testados no restante da GUI.
_ACTION_INSTRUCTIONS = (
    ("COMPOSICAO_BLOQUEADA_SELECAO_MANUAL", "Indicar Aba/Linha do Analítico (Cols CM/CN)."),
    ("CREDITO_NAO_LOCALIZADO", "Indicar Aba/Linha do Analítico (Cols CM/CN)."),
    ("DIVERGENCIA_VALOR", "Conferir tolerância; acima do limite, corrigir as fontes."),
    ("DIVERGENCIA_NOME", "Validar divergência de grafia (APROVADO=SIM)."),
)


def format_action_instruction(pendencias: set[str], composicao_estado: str = "") -> str:
    """One prioritized, objective instruction for the operator from a row's
    PENDENCIAS codes and its COMPOSICAO_ESTADO - never every flag stacked
    together. `composicao_estado` maps a found-but-unapproved composition
    (COMPOSICAO_ESTADO=="PROPOSTA") to its own instruction, since that state
    is not itself one of the PENDENCIAS codes."""
    codes = set(pendencias or ())
    for code, instruction in _ACTION_INSTRUCTIONS:
        if code in codes:
            return instruction
    if composicao_estado == "PROPOSTA":
        return "Validar instrumento e marcar COMPOSICAO_APROVADA=SIM."
    if codes:
        return "Este item precisa de conferência manual antes de continuar."
    return ""


def explain_issue(message: str) -> str:
    """Use only locations present in the issue, without reopening the workbook."""
    if "TOLERATED_WARNING" in message or "CRITICAL_ERROR" in message:
        return message
    row = re.search(r"CREDITOS linha (\d+)", message)
    group = re.search(r"Grupo ([^:]+):", message)
    location = f"CREDITOS, linha {row[1]}: " if row else ""
    if group:
        location = f"Grupo {group[1]}: "
    if "foi alterad" in message and "no Excel" in message:
        action = "Há uma alteração válida de campo final no Excel. Confira os detalhes."
    elif "COMISSAO_PARAMETRO_INVALIDO" in message:
        action = (
            "Volte à configuração, corrija a comissão da PFMI indicada e prepare um novo Excel."
        )
    elif "pendência de origem PFMI" in message or "corrija a PFMI" in message:
        action = "Confira a pendência na PFMI de origem, corrija o arquivo e prepare novamente."
    elif any(
        term in message
        for term in (
            "CREDITO_NAO_LOCALIZADO",
            "sem origem válida",
            "sem crédito original do Analítico",
        )
    ):
        action = (
            "Confira a indicação na PFMI e os registros do Analítico. Depois, prepare novamente."
        )
    elif "DT_EMISSAO_TITULO" in message:
        action = (
            "Confira a data de emissão na aba CREDITOS; ela não pode ser posterior "
            "à liquidação. Salve o Excel."
            if "posterior" in message
            else "Informe uma data de emissão válida na aba CREDITOS e salve o Excel."
        )
    elif "DATA_LIQUIDACAO" in message:
        action = "Corrija a data de liquidação na aba RESUMO e salve o Excel."
    elif "PRIMEIRA_SEQUENCIA" in message or "Faixa de sequência" in message:
        action = "Confira a primeira sequência na aba RESUMO e salve o Excel."
    elif "DIVERGENCIA_VALOR_JUSTIFICATIVA" in message:
        action = (
            "DIVERGENCIA_VALOR_JUSTIFICATIVA e DIVERGENCIA_VALOR_APROVADA=SIM andam juntas: "
            "preencha as duas, ou nenhuma das duas, na aba CREDITOS."
        )
    elif "DIVERGENCIA_VALOR" in message and "APROVADA" in message:
        action = (
            "A composição só pode ser aprovada quando o principal está identificado e, se "
            "houver divergência de valor, ela foi aprovada com justificativa em todas as "
            "linhas da composição."
        )
    elif "SELECAO_MANUAL_DOCUMENTO não fecha exatamente" in message:
        action = (
            "O documento identifica o crédito certo, mas o valor do Analítico diverge do "
            "total da PFMI. Confira o instrumento; se estiver correto, aprove o uso do valor "
            "da PFMI em DIVERGENCIA_VALOR_APROVADA, com justificativa."
        )
    elif "registro selecionado não fecha exatamente" in message:
        action = (
            "O registro escolhido para a composição diverge do total da PFMI. Confira o "
            "instrumento; se estiver correto, aprove o uso do valor da PFMI em "
            "DIVERGENCIA_VALOR_APROVADA, com justificativa, igual em todas as linhas."
        )
    elif "DIVERGENCIA_VALOR" in message:
        action = (
            "Confira a diferença em centavos entre o Analítico e a PFMI. Até o limite "
            "configurado, o TXT usa a PFMI com alerta; acima dele, corrija as fontes e "
            "prepare novamente. A aprovação manual não ultrapassa esse limite."
        )
    elif "DIVERGENCIA_NOME" in message:
        action = (
            "Confira e corrija o nome do cedente. APROVADO=SIM confirma somente "
            "divergência de nome, respeitando a modalidade."
        )
    elif "DOC_CEDENTE" in message or "DOC_SACADO" in message:
        person = "cedente" if "DOC_CEDENTE" in message else "sacado"
        action = f"Confira e corrija o documento do {person} na aba CREDITOS."
    elif "nenhum crédito selecionado" in message or "Nenhum crédito válido" in message:
        action = "Confira a seleção de créditos na aba CREDITOS e as demais pendências do grupo."
    elif "COMPOSICAO_APROVADA" in message or "PREENCHER_VL_NOMINAL_COMPOSICAO" in message:
        action = (
            "Preencha VL_NOMINAL de cada componente e marque COMPOSICAO_APROVADA=SIM "
            "em todas as linhas da composição após conferir o instrumento."
        )
    elif "INCLUIR_CNAB" in message:
        action = "Preencha INCLUIR_CNAB com SIM ou NAO na aba CREDITOS e salve o Excel."
    elif "selecionado mais de uma vez" in message:
        action = "Revise a seleção: um mesmo crédito não pode ser usado em mais de um grupo."
    elif "LOG_AUDITORIA_NAO_GRAVADO" in message:
        action = "O arquivo foi salvo, mas não foi possível gravar o registro de auditoria."
    elif "NOME_TRUNCADO_CNAB" in message:
        action = "Um nome será limitado a 40 caracteres no TXT. O nome completo permanece no Excel."
    elif any(word in message.lower() for word in ("manifesto", "proveniência", "esquema")):
        action = (
            "Os controles do Excel são incompatíveis ou foram alterados. Prepare um novo Excel."
        )
    else:
        action = "Confira este apontamento nos detalhes técnicos e as orientações do Excel."
    return location + action

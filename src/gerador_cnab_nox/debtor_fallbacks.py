"""
Este módulo é responsável por unificar a base histórica de sacados (Massa Falida)
com os novos cadastros feitos localmente pelo operador via Interface Gráfica (GUI).
"""

from __future__ import annotations

import json
import os
import re
import tempfile
from datetime import UTC, datetime
from pathlib import Path

from .errors import InputFileError
from .normalize import digits, normalize_failure, technical_name, validate_document

# Delimitadores usados para encontrar o bloco JSON dentro do arquivo Markdown histórico
BEGIN = "<!-- CNAB_NOX_FALLBACKS_BEGIN -->"
END = "<!-- CNAB_NOX_FALLBACKS_END -->"


def _unique_object(pairs):
    """
    Garante que não existam chaves (CNPJs/Falências) duplicadas ao ler um JSON.
    Se encontrar duplicidade, levanta um erro em vez de sobrescrever silenciosamente.
    """
    result = {}
    for key, value in pairs:
        if key in result:
            raise ValueError(f"chave duplicada: {key}")
        result[key] = value
    return result


def _load_markdown(path=None):
    """
    Carrega a lista de sacados do arquivo histórico Markdown (EQUIVALENCIAS_FALENCIAS.md).
    Procura por um bloco de código JSON específico entre as marcações BEGIN e END.
    """
    # Define o caminho do Markdown: via parâmetro, variável de ambiente ou caminho relativo padrão
    path = Path(path or os.environ.get("CNAB_NOX_EQUIVALENCIAS_PATH") or (
        Path(__file__).resolve().parents[2] / "docs" / "EQUIVALENCIAS_FALENCIAS.md"
    ))
    
    try:
        text = path.read_text(encoding="utf-8")
    except FileNotFoundError:
        return {}
    except OSError as exc:
        raise InputFileError(f"Não foi possível ler o cadastro de sacados: {path.name}") from exc

    result = {}
    
    # Se o arquivo contiver os delimitadores, extrai e valida o conteúdo JSON
    if BEGIN in text or END in text:
        try:
            if text.count(BEGIN) != 1 or text.count(END) != 1:
                raise ValueError("delimitadores duplicados ou ausentes")
            if text.index(BEGIN) > text.index(END):
                raise ValueError("delimitadores fora de ordem")
                
            # Extrai apenas a string que está entre os delimitadores
            block = text.split(BEGIN, 1)[1].split(END, 1)[0].strip()
            
            if not block.startswith("```json\n") or not block.endswith("```"):
                raise ValueError("bloco JSON inválido")
                
            # Lê o JSON garantindo que não há chaves duplicadas
            data = json.loads(block[8:-3], object_pairs_hook=_unique_object)
            
            if not isinstance(data, dict):
                raise ValueError("cadastro deve ser um objeto")

            # Varre os dados do Markdown, validando CNPJ e sanitizando a chave da falência
            for failure, entry in data.items():
                raw = entry["cnpj"]
                if not isinstance(raw, str) or not re.fullmatch(r"[0-9. /-]+", raw):
                    raise ValueError("CNPJ deve conter somente números e pontuação")
                    
                document = digits(raw)
                if len(document) != 14:
                    raise ValueError("CNPJ deve conter 14 dígitos")
                    
                name = entry["nome"]
                key = normalize_failure(failure)
                
                if not key or key in result or not isinstance(name, str) or not name.strip():
                    raise ValueError("falência ou nome inválido/duplicado")
                    
                # Adiciona ao resultado com a indicação da fonte de onde veio
                result[key] = {"document": document, "name": name.strip(), "source": str(path)}
                
        except (KeyError, TypeError, ValueError, AttributeError) as exc:
            raise InputFileError(f"Cadastro de sacados inválido: {exc}") from exc
            
    return result


def default_path():
    """
    Retorna o caminho padrão do arquivo JSON local (cache do operador).
    No Windows, utiliza a pasta %LOCALAPPDATA%.
    """
    base = Path(os.environ.get("LOCALAPPDATA", Path.home() / ".local" / "share"))
    return base / "GeradorCNABNOX" / "sacados_cadastrados.json"


def validate_record(failure, document, name):
    """
    Valida rigorosamente os dados antes de salvar no cache ou ler do cache.
    Esta função garante a integridade (Graceful Degradation e Validação Restrita).
    """
    if not isinstance(failure, str) or not isinstance(name, str):
        raise InputFileError("Informe a falência e o nome oficial em texto.")
        
    # Sanitiza a chave (ex: remove acentos, padroniza maiúsculas)
    key = normalize_failure(failure)
    
    if not key or not name.strip():
        raise InputFileError("Informe a falência e o nome oficial do sacado.")
        
    # Usa a validação matemática para garantir que é um CNPJ real
    doc, kind = validate_document(document)
    if not isinstance(document, str) or kind != 2:
        raise InputFileError("Informe um CNPJ de 14 dígitos com dígitos verificadores válidos.")
        
    try:
        # Garante que o nome atende aos requisitos técnicos (sem caracteres especiais não permitidos)
        technical_name(name)
    except ValueError as exc:
        raise InputFileError(str(exc)) from exc
        
    return key, {"document": doc, "name": name.strip()}


def load_local(path=None):
    """
    Carrega o arquivo JSON local do operador contendo as massas falidas cadastradas.
    """
    path = Path(path) if path is not None else default_path()
    try:
        data = json.loads(path.read_text(encoding="utf-8"), object_pairs_hook=_unique_object)
    except FileNotFoundError:
        return {}
    except (OSError, ValueError) as exc:
        raise InputFileError("Não foi possível ler o cadastro local de sacados.") from exc
        
    try:
        # Valida a versão do esquema do JSON
        if not isinstance(data, dict) or data.get("version") != 1:
            raise ValueError("versão desconhecida")
        if not isinstance(data.get("sacados"), dict):
            raise ValueError("sacados deve ser um objeto")
            
        records = {}
        for failure, entry in data["sacados"].items():
            # Passa todos os registros pela validação rigorosa
            key, record = validate_record(failure, entry["document"], entry["name"])
            stamp = entry["updated_at"]
            
            if not isinstance(stamp, str) or not stamp.endswith("Z"):
                raise ValueError("data de atualização inválida")
                
            # Verifica se o formato da data é um ISO válido
            datetime.fromisoformat(stamp)
            
            if key in records:
                raise ValueError("falência duplicada")
                
            records[key] = {**record, "updated_at": stamp}
            
        return records
    except (KeyError, TypeError, ValueError) as exc:
        raise InputFileError(f"Cadastro local de sacados inválido: {exc}") from exc


def save_debtor(failure, document, name, path=None):
    """
    Salva um novo sacado no JSON local de forma segura utilizando Substituição Atômica.
    Isso impede que o arquivo corrompa se houver queda de energia ou travamento do sistema.
    """
    # 1. Valida todos os dados informados via GUI
    key, record = validate_record(failure, document, name)
    path = Path(path) if path is not None else default_path()
    
    # 2. Carrega os registros já existentes
    records = load_local(path)
    
    # 3. Adiciona a data/hora atual em padrão UTC (ISO 8601)
    record["updated_at"] = datetime.now(UTC).isoformat(timespec="seconds").replace("+00:00", "Z")
    records[key] = record
    
    # Garante que a pasta existe (%LOCALAPPDATA%/GeradorCNABNOX)
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = None
    
    try:
        # 4. Gravação atômica: cria um arquivo temporário primeiro
        with tempfile.NamedTemporaryFile(mode="w", encoding="utf-8", dir=path.parent, delete=False) as handle:
            temporary = Path(handle.name)
            
            # Constrói o schema JSON
            json.dump({"version": 1, "sacados": records}, handle, ensure_ascii=False, indent=2)
            
            # Garante que os dados saíram do buffer da aplicação e foram pro HD
            handle.flush()
            os.fsync(handle.fileno())
            
        # 5. Substitui o arquivo original pelo temporário de uma só vez (atômico)
        os.replace(temporary, path)
    finally:
        # Em caso de falha durante a escrita, apaga o arquivo temporário
        if temporary is not None:
            temporary.unlink(missing_ok=True)
            
    return records


def load_fallbacks(path=None):
    """
    Função principal de consulta para o pipeline ETL (Motor do gerador).
    Carrega o Markdown e depois o JSON local.
    Em caso de conflito, os dados do JSON local (CACHE_LOCAL) sobrescrevem os do Markdown.
    """
    # 1. Traz a Source of Truth original
    records = _load_markdown(path)
    
    # 2. Traz o dicionário do operador e mescla as chaves
    for key, entry in load_local().items():
        # A flag 'source_kind' pode ser usada depois para gerar o alerta de "FALLBACK_SACADO_APLICADO_VIA_CACHE_LOCAL"
        records[key] = {**entry, "source": str(default_path()), "source_kind": "CACHE_LOCAL"}
        
    return records
import os
import json
import tempfile
from pathlib import Path
from typing import Dict, Any
from datetime import datetime, timezone

# Assumindo que a função existente de leitura do MD e a de normalização 
# estão importadas adequadamente no projeto original.
# from .normalize import normalize_failure
# from .md_parser import parse_equivalencias_md

def _get_local_cache_path() -> Path:
    """Resolve o caminho seguro para o cache local do JSON."""
    if os.name == 'nt':
        base_dir = Path(os.environ.get('LOCALAPPDATA', '~')).expanduser()
    else:
        base_dir = Path.home() / '.local' / 'share'
        
    app_dir = base_dir / 'GeradorCNABNOX'
    app_dir.mkdir(parents=True, exist_ok=True)
    return app_dir / 'sacados_cadastrados.json'

def load_fallbacks(local_json_path: Path = None) -> Dict[str, Any]:
    """
    1. Carrega docs/EQUIVALENCIAS_FALENCIAS.md (Source of Truth original).
    2. Carrega o cache JSON local.
    3. Retorna o dicionário com o merge (JSON tem prioridade).
    """
    # 1. Carrega os dados históricos do MD (substituir pela chamada real do projeto)
    # fallbacks_md = parse_equivalencias_md("docs/EQUIVALENCIAS_FALENCIAS.md")
    fallbacks_unificados = {} 
    
    # 2. Localiza e carrega o JSON
    cache_path = local_json_path or _get_local_cache_path()
    
    if cache_path.exists():
        try:
            with open(cache_path, 'r', encoding='utf-8') as f:
                data = json.load(f)
                
            # 3. Merge: JSON prevalece sobre o MD
            if isinstance(data, dict) and "sacados" in data:
                for key, value in data["sacados"].items():
                    fallbacks_unificados[key] = value
        except json.JSONDecodeError:
            # Graceful Degradation: Se o JSON estiver corrompido, ignora a sua leitura,
            # mas não quebra o pipeline. Dependerá apenas do MD.
            pass
            
    return fallbacks_unificados

def save_sacado_to_cache(chave_falencia: str, document: str, name: str) -> None:
    """
    Guarda um novo sacado no JSON local utilizando substituição atómica (os.replace).
    A chave_falencia já deve vir sanitizada do frontend.
    """
    cache_path = _get_local_cache_path()
    
    # Carrega estado atual
    data = {"version": 1, "sacados": {}}
    if cache_path.exists():
        try:
            with open(cache_path, 'r', encoding='utf-8') as f:
                data = json.load(f)
        except json.JSONDecodeError:
            # Ficheiro corrompido será sobrescrito de forma limpa
            pass
            
    # Adiciona/Atualiza o registo
    data["sacados"][chave_falencia] = {
        "document": document,
        "name": name,
        "updated_at": datetime.now(timezone.utc).isoformat()
    }
    
    # Escrita Atómica
    fd, temp_path = tempfile.mkstemp(dir=cache_path.parent, prefix="tmp_sacados_", suffix=".json")
    try:
        with os.fdopen(fd, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        
        # Substitui o ficheiro antigo pelo novo instantaneamente
        os.replace(temp_path, cache_path)
    except Exception as e:
        os.unlink(temp_path)
        raise RuntimeError(f"Falha ao guardar cache de sacados: {str(e)}")
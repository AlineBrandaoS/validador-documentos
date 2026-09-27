"""
CHECKLIST DE MOBILIZAÇÃO (admissão + segurança do trabalho)
===========================================================

Recebe a planilha de colaboradores (NOME e FUNÇÃO), descobre quais documentos
cada um precisa (admissão, NRs e RACs) e confere nas pastas o que foi entregue,
o que está ilegível e o que está FALTANDO.

De onde vêm as regras:
    - Documentos de todos: lista DOCUMENTOS_TODOS abaixo (baseada no checklist SGC).
    - NR11, NR12, NR35, treinamento de equipamento e RACs por função:
      arquivo regras_funcoes.xlsx (na pasta do programa). Você controla essa tabela.
      Quando aparece uma função nova, a IA sugere as regras e marca a linha como
      "IA - REVISAR". Depois de conferir, troque a ORIGEM para "CONFIRMADA".
    - Quem tem qualquer RAC precisa também: PRO de cada RAC, controle de frequência
      ART e frequência PST.
    - NR35 também é exigida quando a planilha de colaboradores marcar APTO NR35 = SIM.

Como o programa reconhece cada documento:
    1. Pelo nome do arquivo, no padrão do SGC (ex.: NR18_JOAO DA SILVA.pdf, RAC01_..., ASO_...).
    2. Se o nome não disser, pelo tipo que a IA identificou ao ler o documento.
"""

from __future__ import annotations
 
import difflib
import re
import unicodedata
from pathlib import Path
from typing import Any
 
import pandas as pd
from openpyxl.utils import get_column_letter
 
import validador_legibilidade as v
 
PASTA_PROGRAMA = Path(__file__).resolve().parent
ARQUIVO_REGRAS = PASTA_PROGRAMA / "regras_funcoes.xlsx"
 
# -----------------------------------------------------------------------------
# Documentos exigidos de TODOS os colaboradores (edite aqui se mudar o checklist)
# (chave, descrição, grupo)
# -----------------------------------------------------------------------------
DOCUMENTOS_TODOS = [
    ("IDENTIFICACAO", "Documento de identificação (RG, CPF ou CNH)", "Admissão"),
    ("CTPS", "CTPS digital, ficha de registro ou contrato (assinados)", "Admissão"),
    ("RESIDENCIA", "Comprovante de residência", "Admissão"),
    ("ESCOLARIDADE", "Comprovante de escolaridade", "Admissão"),
    ("FOTO", "Foto do colaborador", "Admissão"),
    ("ASO", "ASO (exames e riscos conforme PCMSO)", "Segurança"),
    ("FICHA_EPI", "Ficha de EPI", "Segurança"),
    ("ORDEM_SERVICO", "Ordem de serviço (NR01)", "Segurança"),
    ("NR06", "Certificado NR06 + conteúdo programático", "Segurança"),
    ("NR18", "Certificado NR18 + conteúdo programático", "Segurança"),
]
 
DOCS_CONDICIONAIS = {
    "NR11": "Certificado NR11 + conteúdo programático, carta de anuência e atestado de proficiência",
    "NR12": "Certificado NR12 (Eng. Mecânico) + conteúdo programático, carta de anuência e atestado de proficiência",
    "NR35": "Certificado NR35 + conteúdo programático, carta de anuência e atestado de proficiência",
    "TREINAMENTO_EQUIPAMENTO": "Certificado de treinamento do equipamento (operador)",
}
 
RACS = {
    1: "Trabalho em Altura",
    2: "Veículos Automotores Leves",
    3: "Equipamentos Móveis",
    4: "Bloqueio, Identificação e Zero Energia (LOTO)",
    5: "Içamento de Carga",
    6: "Espaço Confinado",
    7: "Proteção de Máquinas",
    8: "Atividades no Terreno / Escavações",
    9: "Explosivos e Detonação",
}
 
COLUNAS_REGRAS = (["FUNÇÃO", "NR11", "NR12", "NR35", "TREINAMENTO EQUIPAMENTO"]
                  + [f"RAC{n:02d}" for n in RACS] + ["ORIGEM", "JUSTIFICATIVA"])
 
OK, FALTANDO, NAO_SE_APLICA = "OK", "FALTANDO", "—"
NAO_EXIGIDA = "NÃO EXIGIDA (ASO)"
DESCR_DISPENSADOS = {
    "NR35": "Certificado NR35",
    "RAC01": "Certificado RAC 01 – Trabalho em Altura",
    "PRO_RAC01": "PRO RAC 01 – Trabalho em Altura",
}
ILEGIVEL_REENVIAR = "ILEGÍVEL"
CORES = {OK: "C6EFCE", FALTANDO: "FFC7CE", ILEGIVEL_REENVIAR: "F4B084", v.REVISAR: "FFEB9C",
         "APTO": "C6EFCE", "INAPTO": "FFC7CE", "NÃO CONSTA": "FFEB9C", "NÃO EXIGIDA (ASO)": "D9D9D9"}
 
 
def normalizar(texto: Any) -> str:
    texto = unicodedata.normalize("NFKD", str(texto or "")).encode("ascii", "ignore").decode()
    texto = re.sub(r"[^A-Z0-9]+", " ", texto.upper())
    return re.sub(r"\s+", " ", texto).strip()
 
 
def marcado(valor: Any) -> bool:
    return normalizar(valor) in {"X", "SIM", "S", "1", "TRUE", "VERDADEIRO", "OK"}
 
 
# =============================================================================
# 1. PLANILHA DE COLABORADORES
# =============================================================================
 
def ler_colaboradores(caminho: Path) -> list[dict]:
    if caminho.suffix.lower() == ".csv":
        tabela = pd.read_csv(caminho, sep=None, engine="python", dtype=str)
    else:
        tabela = pd.read_excel(caminho, dtype=str)
    colunas = {normalizar(c): c for c in tabela.columns}
 
    def achar(*opcoes: str) -> str | None:
        for opcao in opcoes:
            for norm, original in colunas.items():
                if norm == opcao or norm.startswith(opcao):
                    return original
        return None
 
    col_nome = achar("NOME", "COLABORADOR", "FUNCIONARIO")
    col_funcao = achar("FUNCAO", "CARGO", "FUNCOES")
    col_nr35 = achar("APTO NR35", "APTO NR 35", "NR35", "NR 35", "APTO ALTURA")
    if not col_nome or not col_funcao:
        raise RuntimeError("A planilha de colaboradores precisa ter as colunas NOME e FUNÇÃO.")
 
    pessoas = []
    for _, linha in tabela.iterrows():
        nome = str(linha[col_nome] or "").strip()
        if not nome or nome.lower() == "nan":
            continue
        funcao = str(linha[col_funcao] or "").strip()
        pessoas.append({
            "nome": nome,
            "funcao": "" if funcao.lower() == "nan" else funcao,
            "apto_nr35": marcado(linha[col_nr35]) if col_nr35 else False,
        })
    if not pessoas:
        raise RuntimeError("A planilha de colaboradores está vazia.")
    return pessoas
 
 
def criar_modelo_colaboradores(destino: Path) -> Path:
    pd.DataFrame({"NOME": ["JOAO DA SILVA", "MARIA DE SOUZA"],
                  "FUNÇÃO": ["OPERADOR DE ESCAVADEIRA", "AJUDANTE"],
                  "APTO NR35 (opcional)": ["NÃO", "SIM"]}).to_excel(destino, index=False)
    return destino
 
 
# =============================================================================
# 2. REGRAS POR FUNÇÃO (regras_funcoes.xlsx + sugestão da IA)
# =============================================================================
 
def ler_regras() -> dict[str, dict]:
    if not ARQUIVO_REGRAS.exists():
        return {}
    tabela = pd.read_excel(ARQUIVO_REGRAS, dtype=str).fillna("")
    regras = {}
    for _, linha in tabela.iterrows():
        funcao = str(linha.get("FUNÇÃO", "")).strip()
        if not funcao:
            continue
        regras[normalizar(funcao)] = {
            "funcao": funcao,
            "NR11": marcado(linha.get("NR11")),
            "NR12": marcado(linha.get("NR12")),
            "NR35": marcado(linha.get("NR35")),
            "TREINAMENTO_EQUIPAMENTO": marcado(linha.get("TREINAMENTO EQUIPAMENTO")),
            "racs": [n for n in RACS if marcado(linha.get(f"RAC{n:02d}"))],
            "origem": str(linha.get("ORIGEM", "")).strip() or "CONFIRMADA",
            "justificativa": str(linha.get("JUSTIFICATIVA", "")).strip(),
        }
    return regras
 
 
def salvar_regras(regras: dict[str, dict]) -> None:
    linhas = []
    for r in sorted(regras.values(), key=lambda r: r["funcao"]):
        linha = {"FUNÇÃO": r["funcao"],
                 "NR11": "X" if r["NR11"] else "", "NR12": "X" if r["NR12"] else "",
                 "NR35": "X" if r["NR35"] else "",
                 "TREINAMENTO EQUIPAMENTO": "X" if r["TREINAMENTO_EQUIPAMENTO"] else ""}
        for n in RACS:
            linha[f"RAC{n:02d}"] = "X" if n in r["racs"] else ""
        linha["ORIGEM"] = r["origem"]
        linha["JUSTIFICATIVA"] = r["justificativa"]
        linhas.append(linha)
    tabela = pd.DataFrame(linhas, columns=COLUNAS_REGRAS)
    with pd.ExcelWriter(ARQUIVO_REGRAS, engine="openpyxl") as writer:
        tabela.to_excel(writer, sheet_name="Regras", index=False)
        legenda = pd.DataFrame(
            [("Como usar", "Marque X nas NRs/RACs que a função exige. NR06 e NR18 valem para todos "
              "e não precisam ser marcadas.")]
            + [("ORIGEM", "IA - REVISAR = sugerida pela IA, confira e troque para CONFIRMADA.")]
            + [(f"RAC{n:02d}", nome) for n, nome in RACS.items()]
            + [("Quem tem RAC", "Precisa também: PRO de cada RAC, controle de frequência ART e "
                "frequência PST (o programa cobra automaticamente).")],
            columns=["Item", "Descrição"])
        legenda.to_excel(writer, sheet_name="Legenda", index=False)
        from openpyxl.styles import Font, PatternFill
        aba = writer.book["Regras"]
        for celula in aba[1]:
            celula.font = Font(bold=True, color="FFFFFF")
            celula.fill = PatternFill("solid", fgColor="1F4E78")
        aba.column_dimensions["A"].width = 38
        aba.column_dimensions[get_column_letter(len(COLUNAS_REGRAS))].width = 70
        aba.freeze_panes = "B2"
        for linha in range(2, aba.max_row + 1):
            if str(aba.cell(row=linha, column=len(COLUNAS_REGRAS) - 1).value).startswith("IA"):
                for col in range(1, len(COLUNAS_REGRAS) + 1):
                    aba.cell(row=linha, column=col).fill = PatternFill("solid", fgColor="FFEB9C")
        writer.book["Legenda"].column_dimensions["B"].width = 100
 
 
PROMPT_REGRAS = """Você é técnico de segurança do trabalho em obras de construção civil \
(contratada de mineradora que usa os RACs - Requisitos de Atividades Críticas).
 
Para cada função informada, indique quais treinamentos ela exige, seguindo estas regras:
- NR06 e NR18 valem para todos (não precisa informar).
- NR11: somente quem opera equipamentos de transporte/movimentação de materiais \
(empilhadeira, ponte rolante, munck, guindaste) ou faz amarração/sinalização de carga.
- NR12: somente funções que operam ou fazem manutenção em máquinas e equipamentos.
- NR35: somente funções que normalmente trabalham em altura (acima de 2 m).
- TREINAMENTO EQUIPAMENTO: somente operadores de máquinas/equipamentos.
- RACs: {racs}
Seja conservador: marque só o que a função realmente exerce na rotina. Se o nome da \
função for vago, marque só o que for certo e explique na justificativa o que precisa \
ser confirmado. Justificativa curta, em português."""
 
FERRAMENTA_REGRAS = {
    "name": "registrar_regras",
    "description": "Registra os treinamentos exigidos por função.",
    "input_schema": {
        "type": "object",
        "properties": {
            "funcoes": {
                "type": "array",
                "items": {
                    "type": "object",
                    "properties": {
                        "funcao": {"type": "string"},
                        "nr11": {"type": "boolean"},
                        "nr12": {"type": "boolean"},
                        "nr35": {"type": "boolean"},
                        "treinamento_equipamento": {"type": "boolean"},
                        "racs": {"type": "array", "items": {"type": "integer", "minimum": 1, "maximum": 9}},
                        "justificativa": {"type": "string"},
                    },
                    "required": ["funcao", "nr11", "nr12", "nr35", "treinamento_equipamento",
                                 "racs", "justificativa"],
                },
            }
        },
        "required": ["funcoes"],
    },
}
 
 
def sugerir_regras_com_ia(funcoes: list[str]) -> dict[str, dict]:
    """Pede à IA as regras das funções novas. Tudo volta marcado como 'IA - REVISAR'."""
    cliente = v.obter_cliente_ia()
    lista_racs = "; ".join(f"RAC {n:02d} - {nome}" for n, nome in RACS.items())
    sugestoes: dict[str, dict] = {}
    for inicio in range(0, len(funcoes), 25):
        lote = funcoes[inicio:inicio + 25]
        ferramentas: Any = [FERRAMENTA_REGRAS]
        escolha: Any = {"type": "tool", "name": "registrar_regras"}
        mensagens: Any = [{"role": "user", "content":
                           "Funções:\n" + "\n".join(f"- {f}" for f in lote)}]
        resposta = cliente.messages.create(
            model=v.MODELO_IA, max_tokens=4096, system=PROMPT_REGRAS.format(racs=lista_racs),
            tools=ferramentas, tool_choice=escolha, messages=mensagens)
        dados: Any = None
        for bloco in resposta.content:
            if getattr(bloco, "type", None) == "tool_use":
                dados = getattr(bloco, "input", None)
        itens = (dados or {}).get("funcoes") or []
        por_nome = {normalizar(i.get("funcao")): i for i in itens if isinstance(i, dict)}
        for funcao in lote:
            item = por_nome.get(normalizar(funcao))
            if item is None:  # a IA não devolveu esta função: fica só com o básico, para revisar
                item = {"justificativa": "IA não sugeriu regras para esta função: preencher manualmente."}
            sugestoes[normalizar(funcao)] = {
                "funcao": funcao,
                "NR11": bool(item.get("nr11")), "NR12": bool(item.get("nr12")),
                "NR35": bool(item.get("nr35")),
                "TREINAMENTO_EQUIPAMENTO": bool(item.get("treinamento_equipamento")),
                "racs": sorted({int(n) for n in item.get("racs") or [] if str(n).isdigit() and 1 <= int(n) <= 9}),
                "origem": "IA - REVISAR",
                "justificativa": str(item.get("justificativa") or ""),
            }
    return sugestoes
 
 
def achar_regra(funcao: str, regras: dict[str, dict]) -> dict | None:
    chave = normalizar(funcao)
    if chave in regras:
        return regras[chave]
    parecidas = difflib.get_close_matches(chave, list(regras), n=1, cutoff=0.92)
    return regras[parecidas[0]] if parecidas else None
 
 
# =============================================================================
# 3. O QUE CADA COLABORADOR PRECISA
# =============================================================================
 
def documentos_exigidos(pessoa: dict, regra: dict | None) -> list[tuple[str, str, str]]:
    exigidos = list(DOCUMENTOS_TODOS)
    regra = regra or {}
    nota = " (regra sugerida pela IA — confirmar)" if str(regra.get("origem", "")).startswith("IA") else ""
    for chave in ("NR11", "NR12", "TREINAMENTO_EQUIPAMENTO"):
        if regra.get(chave):
            exigidos.append((chave, DOCS_CONDICIONAIS[chave] + nota, "Segurança"))
    if regra.get("NR35") or pessoa.get("apto_nr35") or pessoa.get("aso_altura"):
        motivo = ("" if regra.get("NR35") else
                  " (ASO apto para trabalho em altura)" if pessoa.get("aso_altura")
                  else " (APTO NR35 na planilha)")
        exigidos.append(("NR35", DOCS_CONDICIONAIS["NR35"] + (motivo or nota), "Segurança"))
    racs = regra.get("racs") or []
    for n in racs:
        exigidos.append((f"RAC{n:02d}", f"Certificado RAC {n:02d} – {RACS[n]}{nota}", "RAC"))
        exigidos.append((f"PRO_RAC{n:02d}", f"PRO RAC {n:02d} – {RACS[n]}{nota}", "RAC"))
    if racs:
        exigidos.append(("ART", "Controle de frequência ART", "RAC"))
        exigidos.append(("PST", "Frequência PST", "RAC"))
    return exigidos
 
 
# =============================================================================
# 4. O QUE CADA ARQUIVO É
# =============================================================================
 
IGNORAR = "OUTRO"  # documento reconhecido, mas que não entra no checklist
 
PADROES_NOME = [
    # Documentos que existem na pasta mas não são itens do checklist
    (r"REGRAS?\s*DE\s*OURO|ALCOOL|DROGAS?\b|DIREITOS\s*HUMANOS|TBSSMA|INTEGRACAO|VACINA"
     r"|\bCONTA\b|\bBANCO\b|\bPIS\b|\bNIS\b|TITULO|RESERVISTA|CERTIDAO", IGNORAR),
    (r"\bPRO\s*RAC\s*0?([1-9])\b", "PRO_RAC"),
    (r"\bRAC\s*0?([1-9])\b", "RAC"),
    (r"\bNR\s*0?(\d{1,2})\b", "NR"),
    (r"\bASO\b", "ASO"),
    (r"\bFICHA\s*(DE\s*)?EPI\b|\bEPI\b", "FICHA_EPI"),
    (r"\bORDEM\s*(DE\s*)?SERVICO\b|\bOS\b", "ORDEM_SERVICO"),
    (r"\bART\b", "ART"),
    (r"\bPST\b", "PST"),
    (r"\bCTPS|CARTEIRA\s*DE\s*TRABALHO|FICHA\s*(DE\s*)?REGISTRO|\bCONTRATO\b", "CTPS"),
    (r"\bRG\b|IDENTIDADE|\bCNH\b|\bCPF\b|IDENTIFICACAO", "IDENTIFICACAO"),
    (r"RESIDENCIA|ENDERECO", "RESIDENCIA"),
    (r"ESCOLAR|DIPLOMA|HISTORICO\s*ESCOLAR|CONCLUSAO", "ESCOLARIDADE"),
    (r"\bFOTO\b|3X4", "FOTO"),
    (r"TREINAMENTO\s*(D[OE]\s*)?(EQUIPAMENTO|OPERADOR|OPERACAO)|\bOPERADOR\b|\bOPERACAO\b",
     "TREINAMENTO_EQUIPAMENTO"),
]
 
TIPO_IA_PARA_CHAVE = {
    "RG": "IDENTIFICACAO", "CPF": "IDENTIFICACAO", "CNH": "IDENTIFICACAO",
    "CTPS": "CTPS", "FICHA_REGISTRO_OU_CONTRATO": "CTPS",
    "COMPROVANTE_RESIDENCIA": "RESIDENCIA", "COMPROVANTE_ESCOLARIDADE": "ESCOLARIDADE",
    "FOTO_3X4": "FOTO", "ASO": "ASO", "FICHA_EPI": "FICHA_EPI",
    "ORDEM_SERVICO": "ORDEM_SERVICO", "CONTROLE_FREQUENCIA_ART": "ART",
    "FREQUENCIA_PST": "PST", "CERTIFICADO_TREINAMENTO_EQUIPAMENTO": "TREINAMENTO_EQUIPAMENTO",
}
 
 
def chaves_do_nome(nome_arquivo: str) -> set[str]:
    texto = normalizar(Path(nome_arquivo).stem)
    for padrao, tipo in PADROES_NOME:  # o primeiro padrão que bater define o documento
        achado = re.search(padrao, texto)
        if not achado:
            continue
        if tipo in ("NR", "RAC", "PRO_RAC"):
            numero = int(achado.group(1))
            if tipo == "NR":
                return {f"NR{numero:02d}"}
            return {f"{'PRO_RAC' if tipo == 'PRO_RAC' else 'RAC'}{numero:02d}"}
        return {tipo}
    return set()
 
 
def chaves_da_ia(resultado: dict) -> set[str]:
    chaves = set()
    for pagina in resultado.get("_paginas", []):
        tipo, numero = pagina.get("tipo"), pagina.get("numero") or ""
        if tipo in TIPO_IA_PARA_CHAVE:
            chaves.add(TIPO_IA_PARA_CHAVE[tipo])
        elif tipo == "CERTIFICADO_NR" and numero:
            chaves.add(f"NR{int(numero):02d}")
        elif tipo == "CERTIFICADO_RAC" and numero:
            chaves.add(f"RAC{int(numero):02d}")
        elif tipo == "PRO_RAC" and numero:
            chaves.add(f"PRO_RAC{int(numero):02d}")
    return chaves
 
 
def situacao_do_documento(arquivos: list[dict]) -> tuple[str, str]:
    """Junta todos os arquivos que atendem um item: basta um legível para ficar OK."""
    if not arquivos:
        return FALTANDO, ""
    nomes = ", ".join(a["Arquivo"] for a in arquivos)
    status = [a["Status"] for a in arquivos]
    if v.LEGIVEL in status:
        return OK, nomes
    if v.REVISAR in status:
        return v.REVISAR, nomes
    motivo = next((a["Motivo"] for a in arquivos if a["Status"] == v.ILEGIVEL), "")
    return ILEGIVEL_REENVIAR, f"{nomes}: {motivo}"
 
 
# =============================================================================
# 5. APTIDÃO NO ASO (apto / inapto, trabalho em altura, espaço confinado)
# =============================================================================
 
VERSAO_ASO = "1"
APTO, INAPTO, NAO_CONSTA = "APTO", "INAPTO", "NÃO CONSTA"
 
PROMPT_ASO = """Você é técnico de segurança do trabalho conferindo um ASO (Atestado de Saúde \
Ocupacional) para mobilização em obra. Leia as páginas e informe:
- apto_funcao: conclusão geral do ASO (APTO ou INAPTO para a função). Use \
NAO_IDENTIFICADO se não der para ler ou não houver conclusão.
- trabalho_altura: APTO se o ASO declara aptidão para trabalho em altura (NR35); \
INAPTO se declara inaptidão; NAO_CONSTA se o ASO não menciona trabalho em altura.
- espaco_confinado: mesma lógica para espaço confinado (NR33).
- data_exame no formato DD/MM/AAAA (vazio se não achar).
Não deduza: marque APTO só se estiver escrito ou assinalado no documento. \
confianca = sua certeza (0 a 1) sobre essas leituras. Não transcreva dados pessoais."""
 
FERRAMENTA_ASO = {
    "name": "registrar_aso",
    "description": "Registra as aptidões declaradas no ASO.",
    "input_schema": {
        "type": "object",
        "properties": {
            "apto_funcao": {"type": "string", "enum": ["APTO", "INAPTO", "NAO_IDENTIFICADO"]},
            "trabalho_altura": {"type": "string", "enum": ["APTO", "INAPTO", "NAO_CONSTA"]},
            "espaco_confinado": {"type": "string", "enum": ["APTO", "INAPTO", "NAO_CONSTA"]},
            "data_exame": {"type": "string"},
            "confianca": {"type": "number", "minimum": 0, "maximum": 1},
            "observacao": {"type": "string"},
        },
        "required": ["apto_funcao", "trabalho_altura", "espaco_confinado", "data_exame",
                     "confianca", "observacao"],
    },
}
 
 
def ler_aso(caminho: Path, cache: Any) -> dict:
    """Pede à IA as aptidões do ASO. Guarda no cache para não pagar duas vezes."""
    chave = f"ASO|{v.Cache.chave(caminho)}|a{VERSAO_ASO}"
    guardado = cache.obter(chave)
    if guardado:
        return guardado
    paginas = v.carregar_paginas(caminho)[:4]
    conteudo: Any = [{"type": "image", "source": {"type": "base64", "media_type": "image/jpeg",
                                                   "data": v.imagem_para_base64(p)}} for p in paginas]
    conteudo.append({"type": "text", "text": "Leia as aptidões deste ASO."})
    ferramentas: Any = [FERRAMENTA_ASO]
    escolha: Any = {"type": "tool", "name": "registrar_aso"}
    mensagens: Any = [{"role": "user", "content": conteudo}]
    resposta = v.obter_cliente_ia().messages.create(
        model=v.MODELO_IA, max_tokens=800, system=PROMPT_ASO,
        tools=ferramentas, tool_choice=escolha, messages=mensagens)
    dados: Any = {}
    for bloco in resposta.content:
        if getattr(bloco, "type", None) == "tool_use":
            dados = getattr(bloco, "input", None) or {}
    traduz = {"APTO": APTO, "INAPTO": INAPTO, "NAO_CONSTA": NAO_CONSTA, "NAO_IDENTIFICADO": NAO_CONSTA}
    try:
        confianca = float(dados.get("confianca", 0))
    except (TypeError, ValueError):
        confianca = 0.0
    resultado = {
        "apto": traduz.get(str(dados.get("apto_funcao")), NAO_CONSTA),
        "altura": traduz.get(str(dados.get("trabalho_altura")), NAO_CONSTA),
        "confinado": traduz.get(str(dados.get("espaco_confinado")), NAO_CONSTA),
        "data": str(dados.get("data_exame") or ""),
        "confianca": confianca,
        "obs": str(dados.get("observacao") or ""),
        "arquivo": caminho.name,
    }
    cache.salvar(chave, resultado)
    return resultado
 
 
def aptidao_do_colaborador(arquivos_aso: list[dict], raiz: Path, cache: Any, log) -> dict | None:
    """Lê o(s) ASO(s) legível(is) do colaborador e devolve a aptidão encontrada."""
    for arquivo in arquivos_aso:
        if arquivo["Status"] == v.ILEGIVEL:
            continue
        try:
            return ler_aso(raiz / arquivo["Caminho"], cache)
        except Exception as erro:
            log(f"AVISO: não consegui ler a aptidão do ASO {arquivo['Arquivo']} ({erro}).")
            return {"apto": NAO_CONSTA, "altura": NAO_CONSTA, "confinado": NAO_CONSTA, "data": "",
                    "confianca": 0.0, "obs": f"Falha na leitura: {erro}", "arquivo": arquivo["Arquivo"]}
    return None
 
 
def pendencias_do_aso(pessoa: dict, aso: dict | None, precisa_confinado: bool) -> list[dict]:
    """Confere a aptidão do ASO contra o que a função exige."""
    base = {"Colaborador": pessoa["nome"], "Função": pessoa["funcao"] or "-", "Grupo": "ASO"}
    if aso is None:
        return []  # sem ASO legível: a falta do ASO já aparece no checklist
    detalhe = f"{aso['arquivo']}" + (f" | exame {aso['data']}" if aso["data"] else "") + \
              (f" | {aso['obs']}" if aso["obs"] else "")
    if aso["confianca"] < v.CONFIANCA_MINIMA:
        return [{**base, "Documento": "Aptidão no ASO", "Situação": v.REVISAR,
                 "O que fazer": "Conferir manualmente a aptidão no ASO (IA em dúvida)", "Detalhe": detalhe}]
    lista = []
    if aso["apto"] == INAPTO:
        lista.append({**base, "Documento": "ASO: colaborador INAPTO para a função", "Situação": INAPTO,
                      "O que fazer": "Não mobilizar. Verificar com o SESMT/médico do PCMSO",
                      "Detalhe": detalhe})
    elif aso["apto"] == NAO_CONSTA:
        lista.append({**base, "Documento": "ASO sem conclusão de apto/inapto", "Situação": v.REVISAR,
                      "O que fazer": "Conferir o ASO manualmente", "Detalhe": detalhe})
    if precisa_confinado and aso["confinado"] != APTO:
        texto = ("ASO declara INAPTO para espaço confinado" if aso["confinado"] == INAPTO
                 else "ASO não traz aptidão para espaço confinado")
        lista.append({**base, "Documento": texto, "Situação": aso["confinado"],
                      "O que fazer": "Solicitar ASO com aptidão para espaço confinado", "Detalhe": detalhe})
    return lista
 
 
# =============================================================================
# 6. CHECKLIST
# =============================================================================
 
def gerar_checklist(resultados: list[dict], raiz: Path, planilha: Path, destino: Path,
                    usar_ia: bool, log=print) -> None:
    log("\nConferindo o checklist de documentos por colaborador…")
    pessoas = ler_colaboradores(planilha)
 
    # Regras das funções (e sugestão da IA para as novas)
    regras = ler_regras()
    novas = sorted({p["funcao"] for p in pessoas if p["funcao"] and achar_regra(p["funcao"], regras) is None})
    if novas and usar_ia:
        log(f"{len(novas)} função(ões) sem regra cadastrada: pedindo sugestão à IA…")
        try:
            regras.update(sugerir_regras_com_ia(novas))
            salvar_regras(regras)
            log(f"Sugestões salvas em {ARQUIVO_REGRAS.name} (marcadas 'IA - REVISAR').")
        except Exception as erro:
            log(f"AVISO: não foi possível consultar a IA para as regras ({erro}).")
    elif not ARQUIVO_REGRAS.exists():
        salvar_regras(regras)
 
    # Arquivos de cada pasta de colaborador
    por_pasta: dict[str, list[dict]] = {}
    for r in resultados:
        if r["Colaborador"] != "-":
            por_pasta.setdefault(normalizar(r["Colaborador"]), []).append(r)
    pastas_usadas: set[str] = set()
 
    cache_aso = v.Cache(raiz / "_cache_validador.json") if usar_ia else None
    if usar_ia:
        log("Lendo a aptidão (apto / trabalho em altura) nos ASOs…")
 
    matriz, pendencias = [], []
    for pessoa in pessoas:
        chave_nome = normalizar(pessoa["nome"])
        pasta = chave_nome if chave_nome in por_pasta else None
        if pasta is None:
            parecida = difflib.get_close_matches(chave_nome, list(por_pasta), n=1, cutoff=0.85)
            pasta = parecida[0] if parecida else None
        arquivos = por_pasta.get(pasta, []) if pasta else []
        if pasta:
            pastas_usadas.add(pasta)
 
        regra = achar_regra(pessoa["funcao"], regras) if pessoa["funcao"] else None
 
        mapa: dict[str, list[dict]] = {}
        for arquivo in arquivos:
            # Nome do arquivo + o que a IA viu em cada página (um PDF pode juntar
            # vários documentos, ex.: ordem de serviço + ficha de EPI). Arquivos que o
            # nome já identifica como fora do checklist (regra de ouro, álcool e drogas,
            # direitos humanos...) são ignorados para a IA não confundir.
            pelo_nome = chaves_do_nome(arquivo["Arquivo"])
            if IGNORAR in pelo_nome:
                continue
            for chave in pelo_nome | chaves_da_ia(arquivo):
                mapa.setdefault(chave, []).append(arquivo)
 
        # Aptidão do ASO manda na NR35 (e na RAC 01 - Trabalho em Altura):
        #   ASO apto para altura     -> NR35 exigida, mesmo que a função não peça
        #   ASO não apto / não consta -> NR35 e RAC 01 NÃO são exigidas (não vai trabalhar em altura)
        #   sem ASO ou IA em dúvida  -> vale a regra da função
        aso = aptidao_do_colaborador(mapa.get("ASO", []), raiz, cache_aso, log) if cache_aso else None
        dispensados: list[str] = []
        if aso and aso["confianca"] >= v.CONFIANCA_MINIMA:
            if aso["altura"] == APTO:
                pessoa = {**pessoa, "aso_altura": True}
            else:
                queria = bool((regra or {}).get("NR35")) or pessoa.get("apto_nr35")
                racs = list((regra or {}).get("racs") or [])
                if queria:
                    dispensados.append("NR35")
                if 1 in racs:
                    dispensados += ["RAC01", "PRO_RAC01"]
                pessoa = {**pessoa, "apto_nr35": False}
                if regra:
                    regra = {**regra, "NR35": False, "racs": [n for n in racs if n != 1]}
        exigidos = documentos_exigidos(pessoa, regra)
        codigos = {c for c, _, _ in exigidos}
 
        linha: dict[str, Any] = {"Colaborador": pessoa["nome"], "Função": pessoa["funcao"] or "-",
                                 "Pasta encontrada": (arquivos[0]["Colaborador"] if arquivos else "NÃO ENCONTRADA"),
                                 "ASO: aptidão": aso["apto"] if aso else "-",
                                 "ASO: altura": aso["altura"] if aso else "-"}
        if not arquivos:
            pendencias.append({"Colaborador": pessoa["nome"], "Função": pessoa["funcao"] or "-",
                               "Grupo": "Pasta", "Documento": "Pasta do colaborador não encontrada",
                               "Situação": FALTANDO,
                               "O que fazer": "Criar a pasta com o nome do colaborador (ou conferir a grafia)",
                               "Detalhe": "Por isso todos os documentos aparecem como FALTANDO"})
        faltas = 0
        for chave, descricao, grupo in exigidos:
            situacao, detalhe = situacao_do_documento(mapa.get(chave, []))
            linha[chave] = situacao
            if situacao != OK:
                faltas += 1
                acao = {FALTANDO: "Solicitar o documento",
                        ILEGIVEL_REENVIAR: "Solicitar novo envio legível",
                        v.REVISAR: "Conferir manualmente"}[situacao]
                pendencias.append({"Colaborador": pessoa["nome"], "Função": pessoa["funcao"] or "-",
                                   "Grupo": grupo, "Documento": descricao, "Situação": situacao,
                                   "O que fazer": acao, "Detalhe": detalhe})
        pend_aso = pendencias_do_aso(pessoa, aso, precisa_confinado="RAC06" in codigos)
        pendencias.extend(pend_aso)
        faltas += len(pend_aso)
        if not pessoa["funcao"]:
            pendencias.append({"Colaborador": pessoa["nome"], "Função": "-", "Grupo": "Planilha",
                               "Documento": "Função não informada", "Situação": v.REVISAR,
                               "O que fazer": "Preencher a função para conferir NRs e RACs", "Detalhe": ""})
        elif regra is None:
            pendencias.append({"Colaborador": pessoa["nome"], "Função": pessoa["funcao"], "Grupo": "Regras",
                               "Documento": "Função sem regra de NR/RAC", "Situação": v.REVISAR,
                               "O que fazer": f"Cadastrar a função em {ARQUIVO_REGRAS.name}",
                               "Detalhe": "Conferidos só os documentos obrigatórios para todos"})
        for chave in dispensados:
            linha[chave] = NAO_EXIGIDA
        linha["Pendências"] = faltas
        linha["Situação geral"] = "COMPLETO" if faltas == 0 else "PENDENTE"
        linha["_exigidos"] = exigidos + [(c, DESCR_DISPENSADOS[c], "Segurança") for c in dispensados]
        matriz.append(linha)
 
    for pasta, arquivos in por_pasta.items():
        if pasta not in pastas_usadas:
            pendencias.append({"Colaborador": arquivos[0]["Colaborador"], "Função": "-", "Grupo": "Planilha",
                               "Documento": "Pasta sem colaborador na planilha", "Situação": v.REVISAR,
                               "O que fazer": "Conferir o nome da pasta ou incluir na planilha", "Detalhe": ""})
 
    salvar_checklist(destino, matriz, pendencias, regras)
    completos = sum(1 for m in matriz if m["Situação geral"] == "COMPLETO")
    log(f"Checklist: {completos} de {len(matriz)} colaborador(es) completos | "
        f"{len(pendencias)} pendência(s).")
 
 
def salvar_checklist(destino: Path, matriz: list[dict], pendencias: list[dict],
                     regras: dict[str, dict]) -> None:
    from openpyxl import load_workbook
    from openpyxl.styles import Alignment, Font, PatternFill
 
    # Colunas do checklist na ordem: todos, condicionais, RACs
    descricoes: dict[str, str] = {}
    for linha in matriz:
        for chave, descricao, _ in linha["_exigidos"]:
            descricoes.setdefault(chave, descricao.split(" (regra")[0].split(" (APTO")[0])
    ordem = [c for c, _, _ in DOCUMENTOS_TODOS] + ["NR11", "NR12", "NR35", "TREINAMENTO_EQUIPAMENTO"]
    ordem += [f"{p}{n:02d}" for n in RACS for p in ("RAC", "PRO_RAC")] + ["ART", "PST"]
    colunas = [c for c in ordem if c in descricoes]
    titulos = {c: curto(c) for c in colunas}
 
    tabela = pd.DataFrame([
        {"Colaborador": m["Colaborador"], "Função": m["Função"], "Situação geral": m["Situação geral"],
         "Pendências": m["Pendências"], "Pasta encontrada": m["Pasta encontrada"],
         "ASO: aptidão": m["ASO: aptidão"], "ASO: altura": m["ASO: altura"],
         **{titulos[c]: m.get(c, NAO_SE_APLICA) for c in colunas}}
        for m in matriz])
    tabela_pend = pd.DataFrame(pendencias, columns=["Colaborador", "Função", "Grupo", "Documento",
                                                    "Situação", "O que fazer", "Detalhe"])
 
    livro = load_workbook(destino)
    for nome in ("Pendências", "Checklist"):
        if nome in livro.sheetnames:
            del livro[nome]
    aba_pend = livro.create_sheet("Pendências", 0)
    aba_check = livro.create_sheet("Checklist", 1)
    for aba, df in ((aba_pend, tabela_pend), (aba_check, tabela)):
        aba.append(list(df.columns))
        for registro in df.itertuples(index=False):
            aba.append(list(registro))
        for celula in aba[1]:
            celula.font = Font(bold=True, color="FFFFFF")
            celula.fill = PatternFill("solid", fgColor="1F4E78")
            celula.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
        aba.freeze_panes = "C2" if aba is aba_check else "A2"
        if aba.max_row > 1:
            aba.auto_filter.ref = aba.dimensions
        for linha in aba.iter_rows(min_row=2):
            for celula in linha:
                cor = CORES.get(str(celula.value))
                if str(celula.value) == "COMPLETO":
                    cor = CORES[OK]
                elif str(celula.value) == "PENDENTE":
                    cor = CORES[FALTANDO]
                if cor:
                    celula.fill = PatternFill("solid", fgColor=cor)
                    celula.font = Font(bold=True)
                celula.alignment = Alignment(vertical="top", wrap_text=True,
                                             horizontal="center" if cor else "left")
 
    larguras_pend = {"A": 30, "B": 26, "C": 12, "D": 55, "E": 12, "F": 32, "G": 60}
    for col, larg in larguras_pend.items():
        aba_pend.column_dimensions[col].width = larg
    aba_check.column_dimensions["A"].width = 30
    aba_check.column_dimensions["B"].width = 26
    aba_check.column_dimensions["C"].width = 12
    aba_check.column_dimensions["D"].width = 11
    aba_check.column_dimensions["E"].width = 24
    aba_check.column_dimensions["F"].width = 12
    aba_check.column_dimensions["G"].width = 12
    aba_check.row_dimensions[1].height = 45
    for idx in range(8, 8 + len(colunas)):
        aba_check.column_dimensions[get_column_letter(idx)].width = 12
 
    # Legenda das colunas
    linha_leg = aba_check.max_row + 3
    aba_check.cell(row=linha_leg, column=1, value="Legenda").font = Font(bold=True)
    for i, c in enumerate(colunas, start=1):
        aba_check.cell(row=linha_leg + i, column=1, value=titulos[c])
        aba_check.cell(row=linha_leg + i, column=2, value=descricoes[c])
    fim = linha_leg + len(colunas) + 2
    for i, (sit, txt) in enumerate([(OK, "entregue e legível"), (FALTANDO, "não encontrado na pasta"),
                                    (ILEGIVEL_REENVIAR, "entregue, mas ilegível: pedir de novo"),
                                    (v.REVISAR, "conferir manualmente"),
                                    (NAO_SE_APLICA, "não exigido para esta função"),
                                    (NAO_EXIGIDA, "dispensado: ASO não apto para trabalho em altura")]):
        c = aba_check.cell(row=fim + i, column=1, value=sit)
        if sit in CORES:
            c.fill = PatternFill("solid", fgColor=CORES[sit])
        aba_check.cell(row=fim + i, column=2, value=txt)
 
    # Regras usadas (para conferência)
    if "Regras usadas" in livro.sheetnames:
        del livro["Regras usadas"]
    aba_regras = livro.create_sheet("Regras usadas", 2)
    aba_regras.append(COLUNAS_REGRAS)
    for r in sorted(regras.values(), key=lambda r: r["funcao"]):
        aba_regras.append([r["funcao"], *("X" if r[k] else "" for k in
                                         ("NR11", "NR12", "NR35", "TREINAMENTO_EQUIPAMENTO")),
                           *("X" if n in r["racs"] else "" for n in RACS),
                           r["origem"], r["justificativa"]])
    for celula in aba_regras[1]:
        celula.font = Font(bold=True, color="FFFFFF")
        celula.fill = PatternFill("solid", fgColor="1F4E78")
    aba_regras.column_dimensions["A"].width = 38
    livro.active = 0
    livro.save(destino)
 
 
def curto(chave: str) -> str:
    nomes = {"IDENTIFICACAO": "Identificação", "CTPS": "CTPS/Registro", "RESIDENCIA": "Comp. residência",
             "ESCOLARIDADE": "Escolaridade", "FOTO": "Foto", "ASO": "ASO", "FICHA_EPI": "Ficha EPI",
             "ORDEM_SERVICO": "Ordem de serviço", "TREINAMENTO_EQUIPAMENTO": "Trein. equipamento",
             "ART": "Freq. ART", "PST": "Freq. PST"}
    if chave in nomes:
        return nomes[chave]
    if chave.startswith("PRO_RAC"):
        return f"PRO RAC {chave[-2:]}"
    if chave.startswith("RAC"):
        return f"RAC {chave[-2:]}"
    return chave  # NR06, NR18, NR11...
 
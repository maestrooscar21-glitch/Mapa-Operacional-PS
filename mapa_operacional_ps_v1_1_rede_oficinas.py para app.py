import io
import re
import math
import difflib
import unicodedata

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st


# ============================================================
# MAPA OPERACIONAL PS
# Versão 1.4 — fontes corrigidas: SAVE CSV + Cadastro XLSX
# ============================================================

st.set_page_config(
    page_title="Mapa Operacional PS",
    page_icon="🗺️",
    layout="wide",
)

st.title("🗺️ Mapa Operacional PS")
st.caption(
    "Ferramenta independente para visualizar demanda, serviços executados "
    "e cobertura territorial da rede de oficinas."
)


# ============================================================
# CONSTANTES
# ============================================================

REGIAO_POR_UF = {
    "AC": "Norte", "AP": "Norte", "AM": "Norte", "PA": "Norte",
    "RO": "Norte", "RR": "Norte", "TO": "Norte",
    "AL": "Nordeste", "BA": "Nordeste", "CE": "Nordeste", "MA": "Nordeste",
    "PB": "Nordeste", "PE": "Nordeste", "PI": "Nordeste", "RN": "Nordeste",
    "SE": "Nordeste",
    "DF": "Centro-Oeste", "GO": "Centro-Oeste", "MT": "Centro-Oeste",
    "MS": "Centro-Oeste",
    "ES": "Sudeste", "MG": "Sudeste", "RJ": "Sudeste", "SP": "Sudeste",
    "PR": "Sul", "RS": "Sul", "SC": "Sul",
}

CENTRO_UF = {
    "AC": (-8.77, -70.55), "AL": (-9.62, -36.82),
    "AP": (1.41, -51.77), "AM": (-3.47, -65.10),
    "BA": (-12.96, -41.70), "CE": (-5.20, -39.53),
    "DF": (-15.83, -47.86), "ES": (-19.19, -40.34),
    "GO": (-15.98, -49.86), "MA": (-5.42, -45.44),
    "MT": (-12.64, -55.42), "MS": (-20.51, -54.54),
    "MG": (-18.10, -44.38), "PA": (-3.79, -52.48),
    "PB": (-7.28, -36.72), "PR": (-24.89, -51.55),
    "PE": (-8.38, -37.86), "PI": (-6.60, -42.28),
    "RJ": (-22.25, -42.66), "RN": (-5.81, -36.59),
    "RS": (-30.17, -53.50), "RO": (-10.83, -63.34),
    "RR": (1.99, -61.33), "SC": (-27.45, -50.95),
    "SP": (-22.19, -48.79), "SE": (-10.57, -37.45),
    "TO": (-10.25, -48.25),
}

# Regras de negócio informadas pelo usuário.
OFICINAS_DESCARTADAS = {
    "MIGUEL ANGELO GONCALVES BITENCOURT",
    "VX TECH",
}

# Overrides manuais para oficinas vigentes que não estavam casando com o SAVE.
# O nome vigente é SEMPRE o do CSV do painel.
OVERRIDES_OFICINAS = {
    "VXTECH CONECTA LTDA": {
        "Rua": "10 R Caju",
        "Número": "1157",
        "Bairro": "Caju",
        "Cidade SAVE": "NOVA SANTA RITA",
        "UF SAVE": "RS",
        "CEP": "92480000",
        "Latitude": -29.83457765919663,
        "Longitude": -51.260103322002244,
        "Fonte localização": "Override manual OFS",
    },
    "JF INSTALACOES": {
        "Rua": "R. Vidoça Portela",
        "Número": "121",
        "Bairro": "Loteamento Nene Graeff",
        "Cidade SAVE": "Passo Fundo",
        "UF SAVE": "RS",
        "CEP": "",
        "Latitude": -28.259394848432738,
        "Longitude": -52.44547502883579,
        "Fonte localização": "Override manual OFS",
    },
}


# ============================================================
# UTILITÁRIOS
# ============================================================

def normalizar_texto(valor) -> str:
    texto = str(valor or "").strip().upper()
    return "".join(
        c
        for c in unicodedata.normalize("NFKD", texto)
        if not unicodedata.combining(c)
    )


def numero_float(valor):
    texto = str(valor or "").strip().replace(",", ".")
    try:
        return float(texto)
    except Exception:
        return None


def ler_csv_robusto(arquivo) -> pd.DataFrame:
    bruto = arquivo.getvalue()

    for encoding in ["utf-8-sig", "utf-8", "latin1"]:
        for sep in [",", ";", "\t"]:
            try:
                df = pd.read_csv(
                    io.BytesIO(bruto),
                    sep=sep,
                    encoding=encoding,
                    dtype=str,
                    keep_default_na=False,
                    low_memory=False,
                )
                if len(df.columns) >= 5:
                    return df
            except Exception:
                pass

    raise ValueError(
        "Não foi possível interpretar o CSV. "
        "Verifique delimitador e codificação."
    )


def ler_xlsx(arquivo) -> pd.DataFrame:
    return pd.read_excel(
        io.BytesIO(arquivo.getvalue()),
        dtype=str,
    ).fillna("")


def localizar_coluna(df, candidatos):
    mapa = {
        normalizar_texto(coluna): coluna
        for coluna in df.columns
    }
    for candidato in candidatos:
        chave = normalizar_texto(candidato)
        if chave in mapa:
            return mapa[chave]
    return None


def padronizar_atividade(df: pd.DataFrame, nome_arquivo: str) -> pd.DataFrame:
    base = df.copy()

    aliases = {
        "OS": ["OS", "Ordem de Serviço", "Ordem de Servico"],
        "Ticket Jira": ["Ticket Jira", "Ticket", "Jira"],
        "Placa": ["Placa"],
        "Data": ["Data", "Data da Atividade", "Data Atividade"],
        "Oficina": ["Oficina", "Nome da Oficina"],
        "Cidade": ["Cidade", "Município", "Municipio"],
        "Estado": ["Estado", "UF"],
        "CEP Cliente": ["CEP/Código Postal", "CEP", "Código Postal", "Codigo Postal"],
        "Endereço Cliente": ["Endereço", "Endereco"],
        "Tipo de Atividade": ["Tipo de Atividade", "Tipo Atividade"],
        "Status da Atividade": ["Status da Atividade", "Status"],
        "Cliente": ["Cliente", "Cliente Oficina Própria", "Cliente Oficina Propria"],
        "Recurso": ["Recurso", "Técnico", "Tecnico"],
    }

    out = pd.DataFrame(index=base.index)

    for destino, candidatos in aliases.items():
        origem = localizar_coluna(base, candidatos)
        out[destino] = (
            base[origem].astype(str)
            if origem
            else ""
        )

    out["Arquivo origem"] = nome_arquivo
    out["Data_dt"] = pd.to_datetime(
        out["Data"],
        dayfirst=True,
        errors="coerce",
    )

    out["UF"] = out["Estado"].astype(str).str.strip().str.upper()
    out["Região"] = out["UF"].map(REGIAO_POR_UF).fillna("Sem região")
    out["Chave Oficina"] = out["Oficina"].apply(normalizar_texto)

    out["__chave"] = out["OS"].astype(str).str.strip()
    faltando = out["__chave"].eq("")

    out.loc[faltando, "__chave"] = (
        out.loc[faltando, "Ticket Jira"].astype(str).str.strip()
        + "|"
        + out.loc[faltando, "Placa"].astype(str).str.strip()
        + "|"
        + out.loc[faltando, "Data"].astype(str).str.strip()
        + "|"
        + out.loc[faltando, "Oficina"].astype(str).str.strip()
    )

    return out


def consolidar_atividades(arquivos) -> pd.DataFrame:
    bases = []

    for arquivo in arquivos:
        bases.append(
            padronizar_atividade(
                ler_csv_robusto(arquivo),
                arquivo.name,
            )
        )

    if not bases:
        return pd.DataFrame()

    tudo = pd.concat(
        bases,
        ignore_index=True,
        sort=False,
    )

    tudo = tudo.sort_values(
        ["Data_dt", "Arquivo origem"],
        na_position="last",
    )

    return tudo.drop_duplicates(
        subset=["__chave"],
        keep="last",
    ).reset_index(drop=True)


def somente_concluidos(base: pd.DataFrame) -> pd.DataFrame:
    if base.empty:
        return base

    return base[
        base["Status da Atividade"]
        .map(normalizar_texto)
        .eq("CONCLUIDO")
    ].copy()


def preparar_cadastro_vigente(df: pd.DataFrame) -> pd.DataFrame:
    base = df.copy()

    obrigatorias = [
        "Oficina",
        "Cidade-base",
        "UF-base",
    ]

    faltantes = [
        coluna
        for coluna in obrigatorias
        if coluna not in base.columns
    ]
    if faltantes:
        raise ValueError(
            "Cadastro vigente sem colunas obrigatórias: "
            + ", ".join(faltantes)
        )

    if "Ativa" in base.columns:
        base = base[
            base["Ativa"].map(normalizar_texto).isin(
                {"SIM", "S", "TRUE", "1", "ATIVA", "ATIVO"}
            )
        ].copy()

    base["Oficina vigente"] = base["Oficina"].astype(str).str.strip()
    base["Chave Oficina"] = base["Oficina vigente"].apply(normalizar_texto)
    base["Cidade-base"] = base["Cidade-base"].astype(str).str.strip()
    base["UF-base"] = base["UF-base"].astype(str).str.strip().str.upper()

    # Regras de descarte manual.
    base = base[
        ~base["Chave Oficina"].isin(OFICINAS_DESCARTADAS)
    ].copy()

    return base


def preparar_save(df: pd.DataFrame) -> pd.DataFrame:
    base = df.copy()

    renomear = {
        "Nome Fantasia": "Nome SAVE",
        "Razão Social": "Razão Social SAVE",
        "Rua": "Rua",
        "Número": "Número",
        "Bairro": "Bairro",
        "UF": "UF SAVE",
        "Cidade": "Cidade SAVE",
        "CEP": "CEP",
        "Latitude": "Latitude",
        "Longitude": "Longitude",
    }

    existentes = {
        origem: destino
        for origem, destino in renomear.items()
        if origem in base.columns
    }
    base = base.rename(columns=existentes)

    for coluna in [
        "Nome SAVE",
        "Razão Social SAVE",
        "Rua",
        "Número",
        "Bairro",
        "UF SAVE",
        "Cidade SAVE",
        "CEP",
        "Latitude",
        "Longitude",
    ]:
        if coluna not in base.columns:
            base[coluna] = ""

    base["Chave SAVE"] = base["Nome SAVE"].apply(normalizar_texto)
    base["UF SAVE"] = base["UF SAVE"].astype(str).str.strip().str.upper()
    base["Cidade SAVE"] = base["Cidade SAVE"].astype(str).str.strip()

    base["Latitude"] = base["Latitude"].apply(numero_float)
    base["Longitude"] = base["Longitude"].apply(numero_float)

    return base


def melhor_match_save(linha_vigente, save: pd.DataFrame):
    chave = linha_vigente["Chave Oficina"]
    cidade = normalizar_texto(linha_vigente["Cidade-base"])
    uf = normalizar_texto(linha_vigente["UF-base"])

    # 1) exact match
    exato = save[
        save["Chave SAVE"] == chave
    ]
    if not exato.empty:
        return exato.iloc[0], "Nome exato", 1.0

    # 2) restringe por UF e cidade quando possível
    candidatos = save.copy()

    cand_uf = candidatos[
        candidatos["UF SAVE"].map(normalizar_texto) == uf
    ]
    if not cand_uf.empty:
        candidatos = cand_uf

    cand_cidade = candidatos[
        candidatos["Cidade SAVE"].map(normalizar_texto) == cidade
    ]
    if not cand_cidade.empty:
        candidatos = cand_cidade

    # 3) fuzzy conservador
    melhor = None
    melhor_score = 0.0

    for _, candidato in candidatos.iterrows():
        score = difflib.SequenceMatcher(
            None,
            chave,
            candidato["Chave SAVE"],
        ).ratio()

        if score > melhor_score:
            melhor_score = score
            melhor = candidato

    if melhor is not None and melhor_score >= 0.88:
        return melhor, "Fuzzy nome/cidade/UF", melhor_score

    return None, "Sem correspondência segura", melhor_score


def montar_rede_oficinas(cadastro_vigente, save) -> pd.DataFrame:
    vigente = preparar_cadastro_vigente(cadastro_vigente)
    save = preparar_save(save)

    linhas = []

    for _, linha in vigente.iterrows():
        chave = linha["Chave Oficina"]

        # Override manual.
        if chave in OVERRIDES_OFICINAS:
            o = OVERRIDES_OFICINAS[chave]

            registro = linha.to_dict()
            registro.update(o)
            registro["Nome SAVE"] = ""
            registro["Score match"] = 1.0
            registro["Tipo match"] = "Override manual"
            linhas.append(registro)
            continue

        candidato, tipo_match, score = melhor_match_save(
            linha,
            save,
        )

        registro = linha.to_dict()

        if candidato is not None:
            for coluna in [
                "Nome SAVE",
                "Razão Social SAVE",
                "Rua",
                "Número",
                "Bairro",
                "UF SAVE",
                "Cidade SAVE",
                "CEP",
                "Latitude",
                "Longitude",
            ]:
                registro[coluna] = candidato.get(coluna, "")

            registro["Fonte localização"] = "SAVE"
        else:
            registro.update(
                {
                    "Nome SAVE": "",
                    "Razão Social SAVE": "",
                    "Rua": "",
                    "Número": "",
                    "Bairro": "",
                    "UF SAVE": "",
                    "Cidade SAVE": "",
                    "CEP": "",
                    "Latitude": None,
                    "Longitude": None,
                    "Fonte localização": "Não localizada",
                }
            )

        registro["Score match"] = score
        registro["Tipo match"] = tipo_match
        linhas.append(registro)

    return pd.DataFrame(linhas)


def aplicar_filtros(
    base,
    periodo,
    regioes,
    ufs,
    cidades,
    oficinas,
    tipos,
):
    filtrada = base.copy()

    if isinstance(periodo, (list, tuple)) and len(periodo) == 2:
        inicio, fim = periodo
        filtrada = filtrada[
            filtrada["Data_dt"].notna()
            & (filtrada["Data_dt"].dt.date >= inicio)
            & (filtrada["Data_dt"].dt.date <= fim)
        ]

    if regioes:
        filtrada = filtrada[filtrada["Região"].isin(regioes)]

    if ufs:
        filtrada = filtrada[filtrada["UF"].isin(ufs)]

    if cidades:
        filtrada = filtrada[filtrada["Cidade"].isin(cidades)]

    if oficinas:
        filtrada = filtrada[filtrada["Oficina"].isin(oficinas)]

    if tipos:
        filtrada = filtrada[
            filtrada["Tipo de Atividade"].isin(tipos)
        ]

    return filtrada


# ============================================================
# INPUTS
# ============================================================

st.sidebar.header("📥 Bases")

st.sidebar.markdown("**1) Atividades**")
arquivos_atividades = st.sidebar.file_uploader(
    "Importe um ou vários CSVs de atividades",
    type=["csv"],
    accept_multiple_files=True,
    key="atividades",
    help=(
        "O app valida as colunas internas para confirmar que o arquivo "
        "parece ser um relatório de atividades."
    ),
)

st.sidebar.markdown("**2) Rede de oficinas**")
arquivos_rede = st.sidebar.file_uploader(
    "Importe o SAVE (CSV) + Cadastro da rede (XLS/XLSX)",
    type=["csv", "xls", "xlsx"],
    accept_multiple_files=True,
    key="cadastro_rede_auto",
    help=(
        "Selecione juntos o CSV do SAVE com os nomes vigentes e o XLS/XLSX do cadastro da rede com endereços/coordenadas. "
        "O app identifica automaticamente qual arquivo é qual pelas colunas internas."
    ),
)

st.sidebar.caption(
    "O CSV vigente define os nomes atuais das oficinas. "
    "O SAVE complementa endereço e coordenadas."
)

def validar_arquivo_atividades(df: pd.DataFrame) -> tuple[bool, str]:
    colunas = {normalizar_texto(c) for c in df.columns}
    sinais = {
        "OS",
        "STATUS DA ATIVIDADE",
        "TIPO DE ATIVIDADE",
        "OFICINA",
        "RECURSO",
    }
    presentes = len(colunas.intersection(sinais))
    if presentes >= 3:
        return True, ""
    return False, (
        "O arquivo não parece ser um relatório de atividades. "
        "Esperava encontrar colunas como OS, Status da Atividade, "
        "Tipo de Atividade, Oficina ou Recurso."
    )

def validar_cadastro_vigente(df: pd.DataFrame) -> tuple[bool, str]:
    colunas = {normalizar_texto(c) for c in df.columns}
    obrig = {"OFICINA", "CIDADE-BASE", "UF-BASE"}
    if obrig.issubset(colunas):
        return True, ""
    return False, (
        "Este CSV não parece ser o SAVE com os nomes vigentes das oficinas. "
        "Ele precisa conter Oficina, Cidade-base e UF-base."
    )

def validar_save(df: pd.DataFrame) -> tuple[bool, str]:
    colunas = {normalizar_texto(c) for c in df.columns}
    sinais_nome = {"NOME FANTASIA", "RAZAO SOCIAL"}
    sinais_geo = {"LATITUDE", "LONGITUDE", "CEP", "CIDADE", "UF"}
    if colunas.intersection(sinais_nome) and len(colunas.intersection(sinais_geo)) >= 3:
        return True, ""
    return False, (
        "Esta planilha não parece ser o Cadastro da rede de oficinas. "
        "Esperava encontrar nome/razão social e campos geográficos "
        "como Latitude, Longitude, CEP, Cidade e UF."
    )

if not arquivos_atividades:
    st.info("Importe os CSVs de atividades na barra lateral.")
    st.stop()

atividades_validas = []
for arquivo in arquivos_atividades:
    try:
        df_teste = ler_csv_robusto(arquivo)
        ok, msg = validar_arquivo_atividades(df_teste)
        if not ok:
            st.error(f"{arquivo.name}: {msg}")
            st.stop()
        atividades_validas.append(arquivo)
    except Exception as erro:
        st.error(f"Erro ao ler {arquivo.name}: {erro}")
        st.stop()

with st.spinner("Consolidando atividades..."):
    consolidado = consolidar_atividades(atividades_validas)
    concluidos = somente_concluidos(consolidado)

if concluidos.empty:
    st.warning("Não encontrei atividades concluídas.")
    st.stop()

# ============================================================
# REDE DE OFICINAS
# ============================================================

rede = pd.DataFrame()
save_vigente = None
cadastro_geo = None

if arquivos_rede:
    erros_rede = []

    for arquivo in arquivos_rede:
        nome = arquivo.name.lower()

        try:
            if nome.endswith(".csv"):
                df_rede = ler_csv_robusto(arquivo)

                ok_save, _ = validar_cadastro_vigente(df_rede)
                if ok_save and save_vigente is None:
                    save_vigente = df_rede
                    st.sidebar.success(
                        f"✅ SAVE identificado: {arquivo.name}"
                    )
                else:
                    erros_rede.append(
                        f"{arquivo.name}: o CSV não foi identificado como SAVE vigente."
                    )

            elif nome.endswith((".xls", ".xlsx")):
                df_rede = ler_xlsx(arquivo)

                ok_cadastro, _ = validar_save(df_rede)
                if ok_cadastro and cadastro_geo is None:
                    cadastro_geo = df_rede
                    st.sidebar.success(
                        f"✅ Cadastro da rede identificado: {arquivo.name}"
                    )
                else:
                    erros_rede.append(
                        f"{arquivo.name}: a planilha não foi identificada como Cadastro da rede."
                    )

        except Exception as erro:
            erros_rede.append(f"{arquivo.name}: {erro}")

    for erro in erros_rede:
        st.sidebar.warning(erro)

if save_vigente is not None and cadastro_geo is not None:
    try:
        # Regra de negócio:
        # - nomes vigentes vêm do SAVE (CSV)
        # - endereço/CEP/latitude/longitude vêm do Cadastro da rede (XLS/XLSX)
        rede = montar_rede_oficinas(
            save_vigente,
            cadastro_geo,
        )
    except Exception as erro:
        st.error(f"Erro ao montar rede de oficinas: {erro}")
elif arquivos_rede:
    faltando = []
    if save_vigente is None:
        faltando.append("SAVE (CSV)")
    if cadastro_geo is None:
        faltando.append("Cadastro da rede (XLS/XLSX)")
    st.sidebar.info(
        "Ainda falta identificar: " + " e ".join(faltando)
    )

# ============================================================
# FILTROS
# ============================================================

st.sidebar.divider()
st.sidebar.header("🔎 Filtros")

datas_validas = concluidos["Data_dt"].dropna()

if datas_validas.empty:
    periodo = None
else:
    data_min = datas_validas.min().date()
    data_max = datas_validas.max().date()
    periodo = st.sidebar.date_input(
        "Período",
        value=(data_min, data_max),
        min_value=data_min,
        max_value=data_max,
    )

regioes = st.sidebar.multiselect(
    "Região",
    sorted(
        v
        for v in concluidos["Região"].unique()
        if str(v).strip()
    ),
)

base_uf = concluidos
if regioes:
    base_uf = base_uf[base_uf["Região"].isin(regioes)]

ufs = st.sidebar.multiselect(
    "UF",
    sorted(
        v
        for v in base_uf["UF"].unique()
        if str(v).strip()
    ),
)

base_cidade = base_uf
if ufs:
    base_cidade = base_cidade[base_cidade["UF"].isin(ufs)]

cidades = st.sidebar.multiselect(
    "Cidade",
    sorted(
        v
        for v in base_cidade["Cidade"].unique()
        if str(v).strip()
    ),
)

oficinas = st.sidebar.multiselect(
    "Oficina",
    sorted(
        v
        for v in concluidos["Oficina"].unique()
        if str(v).strip()
    ),
)

tipos = st.sidebar.multiselect(
    "Tipo de atividade",
    sorted(
        v
        for v in concluidos["Tipo de Atividade"].unique()
        if str(v).strip()
    ),
)

filtrada = aplicar_filtros(
    concluidos,
    periodo,
    regioes,
    ufs,
    cidades,
    oficinas,
    tipos,
)

if filtrada.empty:
    st.warning("Nenhum serviço corresponde aos filtros.")
    st.stop()


# ============================================================
# KPIs
# ============================================================

k1, k2, k3, k4, k5 = st.columns(5)

k1.metric(
    "Serviços executados",
    f"{len(filtrada):,}".replace(",", "."),
)
k2.metric(
    "UFs atendidas",
    filtrada.loc[
        filtrada["UF"].str.strip() != "",
        "UF",
    ].nunique(),
)
k3.metric(
    "Cidades atendidas",
    filtrada.loc[
        filtrada["Cidade"].str.strip() != "",
        "Cidade",
    ].nunique(),
)
k4.metric(
    "Oficinas executoras",
    filtrada.loc[
        filtrada["Oficina"].str.strip() != "",
        "Oficina",
    ].nunique(),
)

if rede.empty:
    k5.metric("Oficinas geolocalizadas", "—")
else:
    localizadas = rede[
        rede["Latitude"].notna()
        & rede["Longitude"].notna()
    ]
    k5.metric(
        "Oficinas geolocalizadas",
        f"{len(localizadas)}/{len(rede)}",
    )


# ============================================================
# MAPA NACIONAL COM DUAS CAMADAS
# ============================================================

st.divider()
st.subheader("🌎 Demanda executada × rede de oficinas")

por_uf = (
    filtrada[
        filtrada["UF"].isin(CENTRO_UF)
    ]
    .groupby(["UF", "Região"])
    .size()
    .reset_index(name="Executados")
)

fig = go.Figure()

if not por_uf.empty:
    por_uf["Latitude"] = por_uf["UF"].map(
        lambda uf: CENTRO_UF[uf][0]
    )
    por_uf["Longitude"] = por_uf["UF"].map(
        lambda uf: CENTRO_UF[uf][1]
    )

    fig.add_trace(
        go.Scattergeo(
            lat=por_uf["Latitude"],
            lon=por_uf["Longitude"],
            text=[
                f"<b>{uf}</b><br>"
                f"Região: {regiao}<br>"
                f"Serviços executados: {qtd}"
                for uf, regiao, qtd in zip(
                    por_uf["UF"],
                    por_uf["Região"],
                    por_uf["Executados"],
                )
            ],
            hoverinfo="text",
            mode="markers+text",
            textposition="middle center",
            marker=dict(
                size=(
                    10
                    + por_uf["Executados"]
                    / max(por_uf["Executados"].max(), 1)
                    * 45
                ),
                opacity=0.55,
            ),
            name="Demanda executada por UF",
        )
    )

if not rede.empty:
    oficinas_mapa = rede[
        rede["Latitude"].notna()
        & rede["Longitude"].notna()
    ].copy()

    if regioes:
        oficinas_mapa = oficinas_mapa[
            oficinas_mapa["UF-base"].map(REGIAO_POR_UF).isin(regioes)
        ]
    if ufs:
        oficinas_mapa = oficinas_mapa[
            oficinas_mapa["UF-base"].isin(ufs)
        ]
    if oficinas:
        oficinas_mapa = oficinas_mapa[
            oficinas_mapa["Oficina vigente"].isin(oficinas)
        ]

    if not oficinas_mapa.empty:
        fig.add_trace(
            go.Scattergeo(
                lat=oficinas_mapa["Latitude"],
                lon=oficinas_mapa["Longitude"],
                text=[
                    f"<b>{nome}</b><br>"
                    f"{cidade}/{uf}<br>"
                    f"{rua}, {numero}<br>"
                    f"Fonte: {fonte}"
                    for nome, cidade, uf, rua, numero, fonte in zip(
                        oficinas_mapa["Oficina vigente"],
                        oficinas_mapa["Cidade-base"],
                        oficinas_mapa["UF-base"],
                        oficinas_mapa["Rua"],
                        oficinas_mapa["Número"],
                        oficinas_mapa["Fonte localização"],
                    )
                ],
                hoverinfo="text",
                mode="markers",
                marker=dict(
                    size=9,
                    symbol="diamond",
                    line=dict(width=1),
                ),
                name="Oficinas",
            )
        )

fig.update_geos(
    scope="south america",
    projection_type="natural earth",
    lataxis_range=[-35, 6],
    lonaxis_range=[-75, -32],
    showcountries=True,
    showcoastlines=True,
    showland=True,
)

fig.update_layout(
    height=680,
    margin=dict(l=0, r=0, t=10, b=0),
    legend=dict(
        orientation="h",
        yanchor="bottom",
        y=1.01,
        xanchor="left",
        x=0,
    ),
)

st.plotly_chart(
    fig,
    use_container_width=True,
)

st.caption(
    "Círculos = volume executado por UF. "
    "Losangos = localização real das oficinas com coordenadas disponíveis. "
    "A próxima camada geocodifica o cliente para calcular distância oficina → atendimento."
)


# ============================================================
# DIAGNÓSTICO DA REDE
# ============================================================

if not rede.empty:
    st.divider()
    st.subheader("🔧 Qualidade do cadastro geográfico da rede")

    localizadas = rede[
        rede["Latitude"].notna()
        & rede["Longitude"].notna()
    ].copy()

    pendentes = rede[
        rede["Latitude"].isna()
        | rede["Longitude"].isna()
    ].copy()

    r1, r2, r3 = st.columns(3)

    r1.metric(
        "Oficinas vigentes consideradas",
        len(rede),
    )
    r2.metric(
        "Com latitude + longitude",
        len(localizadas),
    )
    r3.metric(
        "Pendentes de coordenada",
        len(pendentes),
    )

    if not pendentes.empty:
        st.markdown("#### ⚠️ Oficinas ainda sem coordenada completa")
        st.dataframe(
            pendentes[
                [
                    "Oficina vigente",
                    "Cidade-base",
                    "UF-base",
                    "Tipo match",
                    "Fonte localização",
                ]
            ],
            use_container_width=True,
            hide_index=True,
        )


# ============================================================
# PARTICIPAÇÃO REGIONAL
# ============================================================

st.divider()
st.subheader("📊 Participação dos serviços por região")

por_regiao = (
    filtrada.groupby("Região")
    .size()
    .reset_index(name="Executados")
    .sort_values("Executados", ascending=False)
)

por_regiao["% do total"] = (
    por_regiao["Executados"]
    / por_regiao["Executados"].sum()
    * 100
)

grafico_regiao = px.bar(
    por_regiao,
    x="Região",
    y="Executados",
    text="Executados",
    hover_data={"% do total": ":.1f"},
)

grafico_regiao.update_layout(
    height=400,
    showlegend=False,
)

st.plotly_chart(
    grafico_regiao,
    use_container_width=True,
)


# ============================================================
# TOP CIDADES E OFICINAS
# ============================================================

c1, c2 = st.columns(2)

with c1:
    st.subheader("🏙️ Top cidades atendidas")
    top_cidades = (
        filtrada[
            filtrada["Cidade"].str.strip() != ""
        ]
        .groupby(["Cidade", "UF"])
        .size()
        .reset_index(name="Executados")
        .sort_values("Executados", ascending=False)
        .head(25)
    )
    st.dataframe(
        top_cidades,
        use_container_width=True,
        hide_index=True,
    )

with c2:
    st.subheader("🔧 Top oficinas executoras")
    top_oficinas = (
        filtrada[
            filtrada["Oficina"].str.strip() != ""
        ]
        .groupby("Oficina")
        .size()
        .reset_index(name="Executados")
        .sort_values("Executados", ascending=False)
        .head(25)
    )
    st.dataframe(
        top_oficinas,
        use_container_width=True,
        hide_index=True,
    )


# ============================================================
# COBERTURA TERRITORIAL ATUAL
# ============================================================

st.divider()
st.subheader("🧭 Cobertura territorial observada por oficina")

cobertura = (
    filtrada[
        filtrada["Oficina"].str.strip() != ""
    ]
    .groupby("Oficina")
    .agg(
        Serviços=("__chave", "count"),
        UFs=("UF", lambda s: s[s.str.strip() != ""].nunique()),
        Cidades=("Cidade", lambda s: s[s.str.strip() != ""].nunique()),
    )
    .reset_index()
    .sort_values(
        ["Serviços", "Cidades"],
        ascending=[False, False],
    )
)

if not rede.empty:
    apoio = rede[
        [
            "Oficina vigente",
            "Cidade-base",
            "UF-base",
            "Latitude",
            "Longitude",
            "Fonte localização",
        ]
    ].copy()

    cobertura = cobertura.merge(
        apoio,
        left_on="Oficina",
        right_on="Oficina vigente",
        how="left",
    )

st.dataframe(
    cobertura,
    use_container_width=True,
    hide_index=True,
)


# ============================================================
# MIX DE SERVIÇO
# ============================================================

st.divider()
st.subheader("🧰 Mix de serviços executados")

mix = (
    filtrada[
        filtrada["Tipo de Atividade"].str.strip() != ""
    ]
    .groupby("Tipo de Atividade")
    .size()
    .reset_index(name="Executados")
    .sort_values("Executados", ascending=False)
)

if not mix.empty:
    grafico_mix = px.bar(
        mix.head(20),
        x="Executados",
        y="Tipo de Atividade",
        orientation="h",
        text="Executados",
    )
    grafico_mix.update_layout(
        height=max(450, 35 * min(len(mix), 20)),
        yaxis={"categoryorder": "total ascending"},
    )
    st.plotly_chart(
        grafico_mix,
        use_container_width=True,
    )


# ============================================================
# DETALHE
# ============================================================

st.divider()
st.subheader("📋 Serviços que compõem a visão")

colunas_detalhe = [
    "Data",
    "OS",
    "Ticket Jira",
    "Placa",
    "Tipo de Atividade",
    "Oficina",
    "Recurso",
    "Cliente",
    "Cidade",
    "UF",
    "Região",
    "CEP Cliente",
    "Endereço Cliente",
]

st.dataframe(
    filtrada[
        [c for c in colunas_detalhe if c in filtrada.columns]
    ].sort_values(
        ["Data_dt", "UF", "Cidade", "Oficina"],
        ascending=[False, True, True, True],
    ),
    use_container_width=True,
    hide_index=True,
    height=520,
)

st.download_button(
    "⬇️ Baixar visão filtrada em CSV",
    data=filtrada.to_csv(
        index=False,
        encoding="utf-8-sig",
    ),
    file_name="mapa_operacional_visao_filtrada.csv",
    mime="text/csv",
)


# ============================================================
# PRÓXIMA CAMADA
# ============================================================

st.divider()
st.info(
    "Próxima evolução: geocodificar CEP/endereço do cliente, calcular "
    "distância oficina → atendimento, raio típico por oficina e cidades "
    "com demanda suficiente para estudar novos prestadores."
)

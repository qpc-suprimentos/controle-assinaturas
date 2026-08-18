# -*- coding: utf-8 -*-
"""
PAINEL DE ACOMPANHAMENTO DE CONTRATOS - CLICKSIGN VIA OUTLOOK
=============================================================

O QUE ESTE SCRIPT FAZ, EM PORTUGUES SIMPLES:

Voce coloca seu e-mail como OBSERVADOR em toda assinatura da Clicksign. Por causa
disso, a Clicksign te manda um e-mail a cada movimentacao do documento. Esses
e-mails, juntos, contam a historia completa de cada contrato: quem ja assinou,
quem falta, se foi finalizado, se foi cancelado.

Este script pega o "dump" desses e-mails (um arquivo JSON gerado pelo comando
/painel-contratos, que le o Outlook), aplica as regras de negocio que voce
definiu, e gera dois arquivos:

  1. 02-dados-tratados/contratos-clicksign.json  -> os dados limpos, auditaveis
  2. 03-painel/painel-contratos.html             -> o painel para abrir no navegador

Rodar este script duas vezes com o mesmo arquivo de entrada gera exatamente o
mesmo resultado. Nenhum numero e inventado: tudo que aparece no painel foi lido
de um e-mail, e cada linha do painel tem link para o e-mail de origem.

COMO RODAR:
    python 05-scripts/03-painel-contratos-clicksign.py
"""

import json
import os
import re
import shutil
import unicodedata
from datetime import datetime, timezone
from zoneinfo import ZoneInfo

# =============================================================================
# CONFIGURACAO  - e so aqui que voce mexe
# =============================================================================

# Pasta raiz do projeto (o script assume que esta em 05-scripts/)
PASTA_RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

# Arquivo de entrada: o dump bruto dos e-mails. O comando /painel-contratos
# grava um arquivo novo por data. Deixe None para pegar automaticamente o mais recente.
ARQUIVO_ENTRADA = None

# Onde salvar a saida
ARQUIVO_TRATADO = os.path.join(PASTA_RAIZ, "02-dados-tratados", "contratos-clicksign.json")
ARQUIVO_PAINEL = os.path.join(PASTA_RAIZ, "03-painel", "painel-contratos.html")
PASTA_HISTORICO = os.path.join(PASTA_RAIZ, "03-painel", "historico")

# A partir de quantos dias sem movimentacao um contrato EM ANDAMENTO e considerado "parado".
# 7 dias = uma semana util inteira sem ninguem assinar. Ajuste se seu ciclo for outro.
DIAS_PARA_CONSIDERAR_PARADO = 7

# Faltando quantos dias para a data limite o contrato entra em alerta de prazo.
DIAS_ALERTA_PRAZO = 10

# Incluir e-mails do ambiente de teste da Clicksign (SANDBOX)?
# Decisao do Valter em 17/08/2026: NAO. Sandbox nao tem valor legal.
INCLUIR_SANDBOX = False

# Fuso horario de quem le o painel. TUDO que e data ou hora no painel e
# convertido para ca antes de aparecer.
#
# Por que isto existe: o computador onde este script roda usa UTC, e o Brasil
# esta 3 horas atras. Sem esta linha o painel dizia "atualizado 13:11" quando no
# relogio do Valter eram 10:11. Pior que o horario feio: um e-mail recebido de
# madrugada (ex.: 02:00 UTC = 23:00 do dia anterior em Salvador) caia no dia
# seguinte e a contagem de "dias parado" saia um dia errada.
FUSO_LOCAL = ZoneInfo("America/Sao_Paulo")

# Data de referencia do painel. None = agora.
DATA_REFERENCIA = None


# =============================================================================
# REGRAS DE NEGOCIO - definidas pelo Valter, nao alterar sem pedido dele
# =============================================================================

# Um documento so entra no painel se o nome dele tiver o padrao de contrato de obra:
# uma sigla (AFE ou CT), a obra G200 e um numero. Isso exclui automaticamente
# documentos pessoais e de RH (PJ_QPC, F.PES.012) que chegam do mesmo remetente.
PADRAO_CONTRATO = re.compile(
    r"(?P<sigla>AFE|CT)[-\s.]*(?P<lsf>LSF)?[-\s.]*G200[-\s.]*(?P<numero>\d{2,3})",
    re.IGNORECASE,
)

# Padrao para achar o nome completo do arquivo do contrato dentro do texto do e-mail.
#
# Cuidado que custou um bug: o nome do contrato as vezes vem no meio de uma frase
# ("O processo de assinatura do documento 65.CT-LSF-G200-065-26 - MB Terraplenagem.pdf
# junto a organizacao QPC..."). Se o padrao comecar solto, ele engole "O processo de
# assinatura do documento" junto e o mesmo contrato vira dois no painel.
# Por isso o padrao comeca obrigatoriamente no numero de ordem ("65.") ou na sigla.
PADRAO_NOME_ARQUIVO = re.compile(
    r"((?:\d{1,3}\s*[.\-]\s*)?(?:AFE|CT)[-\s.]*(?:LSF)?[-\s.]*G200[^\n]*?\.pdf)",
    re.IGNORECASE,
)

# Linha de signatario. A Clicksign escreve "Fulano Assinou como testemunha" para
# quem ja assinou, e "fulano@ assinara como contratante" para quem ainda falta.
# E o VERBO que diz o status, nao a caixa alta.
PADRAO_SIGNATARIO = re.compile(
    r"^(?P<quem>.+?)\s+(?P<verbo>assinou|assinar[aá])\s+como\s+(?P<papel>.+?)\s*$",
    re.IGNORECASE | re.MULTILINE,
)

# Data limite de assinatura, em qualquer das tres redacoes que a Clicksign usa.
PADRAO_DATA_LIMITE = re.compile(
    r"Data limite(?:\s+(?:de|para)\s+assinatura)?:\s*(\d{2}/\d{2}/\d{4})", re.IGNORECASE
)

# Como cada assunto de e-mail se traduz em evento de contrato.
# A ordem importa: o primeiro que casar vence.
EVENTOS = [
    ("cancelado", re.compile(r"documento cancelado", re.IGNORECASE)),
    ("recusado", re.compile(r"documento recusado|recusou a assinatura", re.IGNORECASE)),
    ("finalizado", re.compile(r"foi finalizado|documento assinado:", re.IGNORECASE)),
    ("comprovante", re.compile(r"comprovante de assinatura", re.IGNORECASE)),
    ("prazo", re.compile(r"perto de atingir a data limite", re.IGNORECASE)),
    ("solicitacao", re.compile(r"assinar documento:|solicita", re.IGNORECASE)),
]

# Status final do contrato, na ordem de prioridade. Cancelado e recusado sao
# terminais: uma vez cancelado, nao volta a ser "em andamento".
ORDEM_STATUS = ["Cancelado", "Recusado", "Finalizado", "Em andamento"]


def erro(mensagem):
    """Para tudo e avisa em portugues. Regra do projeto: falhar alto, nao baixo."""
    raise SystemExit("\n[ERRO] " + mensagem + "\n")


def sem_acento(texto):
    """Tira acento para comparar textos sem depender de como foi digitado."""
    return "".join(
        c for c in unicodedata.normalize("NFD", texto) if unicodedata.category(c) != "Mn"
    )


def achar_arquivo_entrada():
    """Pega o dump de e-mails mais recente da pasta de dados brutos."""
    if ARQUIVO_ENTRADA:
        return ARQUIVO_ENTRADA
    pasta = os.path.join(PASTA_RAIZ, "01-dados-brutos")
    if not os.path.isdir(pasta):
        erro("A pasta 01-dados-brutos nao existe. Rode o comando /painel-contratos primeiro.")
    candidatos = sorted(
        f for f in os.listdir(pasta) if f.startswith("emails-clicksign-") and f.endswith(".json")
    )
    if not candidatos:
        erro(
            "Nenhum arquivo emails-clicksign-AAAA-MM-DD.json encontrado em 01-dados-brutos/.\n"
            "Esse arquivo e gerado pelo comando /painel-contratos, que le o seu Outlook."
        )
    return os.path.join(pasta, candidatos[-1])


def classificar_tipo(nome_documento):
    """
    Descobre se o contrato e AFE, LYON ou QPC olhando o nome do arquivo.

    Regra confirmada pelo Valter em 17/08/2026:
      - tem "AFE" no nome            -> AFE
      - nao tem AFE mas tem "LSF"    -> LYON
      - nao tem AFE nem LSF          -> QPC

    Exemplos reais:
      "71. AFE-LSF-G200-071.26 - VIAPOL"      -> AFE  (AFE vence, mesmo tendo LSF)
      "86.CT-LSF-G200-086.26 - BLINDADUS"     -> LYON
      "84.CT-LSFG200-084.26 - GEONORDESTE"    -> LYON (LSF colado no G200, sem hifen)
      "03. CT-G200-003.26 - CTMED (Exames)"   -> QPC
    """
    nome = sem_acento(nome_documento).upper()
    if "AFE" in nome:
        return "AFE"
    if "LSF" in nome:
        return "LYON"
    return "QPC"


def extrair_numero_contrato(nome_documento):
    """Devolve o numero do contrato (ex.: '086') lido do proprio nome do arquivo."""
    achado = PADRAO_CONTRATO.search(sem_acento(nome_documento))
    return achado.group("numero").zfill(3) if achado else None


def extrair_nome_documento(email):
    """
    Acha o nome do arquivo do contrato dentro do e-mail.

    Por que olhar assunto E corpo: nos e-mails de "Comprovante" e "Cancelado" o nome
    esta no assunto; nos de "foi finalizado" e "perto da data limite" o assunto e
    generico e o nome so aparece no corpo.
    """
    for texto in (email.get("assunto", ""), email.get("corpo", "")):
        achado = PADRAO_NOME_ARQUIVO.search(texto)
        if achado:
            nome = achado.group(1).strip().strip("-").strip()
            return re.sub(r"\s+", " ", nome)
    return None


def chave_documento(nome):
    """
    Identidade do documento para agrupar e-mails.

    Usamos o NOME DO ARQUIVO, nao o numero do contrato. Motivo pratico: o contrato
    065 foi cancelado como "MB Terrplenagem (Limpeza Vicinal)" e reemitido como
    "MB Terraplenagem (Limpeza)". Sao dois documentos na Clicksign e voce precisa
    ver os dois - o cancelamento e um fato do processo, nao um erro a ser escondido.
    """
    return re.sub(r"\s+", " ", sem_acento(nome).lower().replace(".pdf", "")).strip()


def identificar_evento(assunto, corpo):
    """Traduz o assunto do e-mail em um tipo de evento do contrato."""
    texto = sem_acento(assunto + " " + corpo)
    for nome_evento, padrao in EVENTOS:
        if padrao.search(texto):
            return nome_evento
    return "outro"


def parsear_signatarios(corpo):
    """
    Le a lista de signatarios de um e-mail de comprovante.

    A Clicksign lista, no mesmo e-mail, quem ja assinou e quem falta:
      "Fabio Matheus Assinou como testemunha"          -> ja assinou
      "nilt***@lyoncapital.com.br assinara como contratante" -> falta assinar
    """
    signatarios = []
    for achado in PADRAO_SIGNATARIO.finditer(corpo):
        quem = achado.group("quem").strip()
        verbo = sem_acento(achado.group("verbo")).lower()
        papel = achado.group("papel").strip().rstrip(".").strip()
        # Linhas de cabecalho as vezes casam por acidente; descarta o obvio.
        if not quem or len(quem) > 90:
            continue
        signatarios.append(
            {"quem": quem, "papel": papel, "assinou": verbo == "assinou"}
        )
    return signatarios


def ler_data(texto_iso):
    """
    Converte a data que veio do e-mail para o fuso de quem le o painel.

    O Outlook entrega tudo em UTC. Converter aqui, na entrada, garante que
    todo o resto do script ja trabalhe no horario de Salvador - inclusive a
    conta de dias, que compara datas de calendario.
    """
    return datetime.fromisoformat(texto_iso.replace("Z", "+00:00")).astimezone(FUSO_LOCAL)


def dias_entre(inicio, fim):
    return (fim.date() - inicio.date()).days


def processar():
    caminho_entrada = achar_arquivo_entrada()
    with open(caminho_entrada, encoding="utf-8") as arquivo:
        bruto = json.load(arquivo)

    if "emails" not in bruto:
        erro("O arquivo %s nao tem a chave 'emails'. Formato inesperado." % caminho_entrada)

    agora = (
        datetime.fromisoformat(DATA_REFERENCIA).astimezone(FUSO_LOCAL)
        if DATA_REFERENCIA
        else datetime.now(FUSO_LOCAL)
    )

    documentos = {}
    descartados = {"sandbox": 0, "nao_e_contrato_de_obra": 0}

    for email in bruto["emails"]:
        if email.get("sandbox") and not INCLUIR_SANDBOX:
            descartados["sandbox"] += 1
            continue

        nome = extrair_nome_documento(email)
        if not nome or not PADRAO_CONTRATO.search(sem_acento(nome)):
            descartados["nao_e_contrato_de_obra"] += 1
            continue

        chave = chave_documento(nome)
        evento = identificar_evento(email.get("assunto", ""), email.get("corpo", ""))
        recebido = ler_data(email["recebido_em"])

        doc = documentos.setdefault(
            chave,
            {
                "nome": nome,
                "numero": extrair_numero_contrato(nome),
                "tipo": classificar_tipo(nome),
                "eventos": [],
                "signatarios": [],
                "data_signatarios": None,
                "data_limite": None,
                "ultima_movimentacao": recebido,
                "link_email": email.get("web_link") or "",
                "link_clicksign": "",
                "finalizado_em": None,
                "cancelado_em": None,
            },
        )

        doc["eventos"].append({"tipo": evento, "em": recebido})

        # A movimentacao mais recente define a data de referencia do contrato.
        if recebido > doc["ultima_movimentacao"]:
            doc["ultima_movimentacao"] = recebido
            if email.get("web_link"):
                doc["link_email"] = email["web_link"]

        # Lista de signatarios: sempre a do comprovante MAIS RECENTE.
        # Comprovantes antigos mostram um retrato desatualizado das assinaturas.
        if evento == "comprovante":
            lista = parsear_signatarios(email.get("corpo", ""))
            if lista and (doc["data_signatarios"] is None or recebido > doc["data_signatarios"]):
                doc["signatarios"] = lista
                doc["data_signatarios"] = recebido

        if evento == "finalizado":
            if doc["finalizado_em"] is None or recebido < doc["finalizado_em"]:
                doc["finalizado_em"] = recebido
        if evento == "cancelado":
            doc["cancelado_em"] = recebido

        achado_prazo = PADRAO_DATA_LIMITE.search(email.get("corpo", ""))
        if achado_prazo:
            doc["data_limite"] = achado_prazo.group(1)

        achado_link = re.search(r"https://app\.clicksign\.com/\S+?(?=[\s.]|$)", email.get("corpo", ""))
        if achado_link and not doc["link_clicksign"]:
            doc["link_clicksign"] = achado_link.group(0)

    # ---- Guarda-corpo 1: dump vazio nao e painel vazio, e busca furada -------
    # Falha silenciosa e o pior tipo: o painel sai zerado com cara de correto e
    # voce vai procurar o problema na Clicksign, no lugar errado.
    if not bruto["emails"]:
        erro(
            "O dump nao tem NENHUM e-mail. Isso nao e um painel vazio - e uma busca que falhou.\n"
            "Causa mais provavel: a busca no Outlook usou o parametro 'order'.\n"
            "Sem 'folderName', o 'order' restringe a busca a Caixa de Entrada, e os e-mails\n"
            "da Clicksign nao estao nela. Refaca a busca SEM 'order'. NAO publique."
        )

    if not documentos:
        erro(
            "Foram lidos %d e-mails, mas nenhum e contrato de obra da G200.\n"
            "Confira se o dump em 01-dados-brutos/ e do remetente e do periodo certos.\n"
            "NAO vou gerar painel vazio." % len(bruto["emails"])
        )

    # ---- Guarda-corpo 2: queda brusca no numero de contratos ------------------
    # Contrato so sai do painel se for renomeado na Clicksign. Uma queda grande
    # de uma rodada para outra e sintoma de coleta furada, nao de realidade.
    if os.path.exists(ARQUIVO_TRATADO):
        try:
            with open(ARQUIVO_TRATADO, encoding="utf-8") as anterior:
                antes = len(json.load(anterior).get("contratos", []))
        except Exception:
            antes = 0
        if antes and len(documentos) < antes * 0.7:
            erro(
                "A rodada anterior tinha %d contratos e esta tem so %d - queda de %.0f%%.\n"
                "Contrato nao some sozinho do painel. Isso e sintoma de coleta incompleta.\n"
                "Confira a busca no Outlook antes de publicar. NAO vou sobrescrever o painel."
                % (antes, len(documentos), 100 * (1 - len(documentos) / antes))
            )

    # ---- Fecha o status de cada contrato e monta a saida ----------------------
    contratos = []
    for doc in documentos.values():
        if doc["cancelado_em"]:
            status = "Cancelado"
        elif any(e["tipo"] == "recusado" for e in doc["eventos"]):
            status = "Recusado"
        elif doc["finalizado_em"]:
            status = "Finalizado"
        else:
            status = "Em andamento"

        total_assinaturas = len(doc["signatarios"])
        ja_assinaram = sum(1 for s in doc["signatarios"] if s["assinou"])

        # Contrato finalizado tem, por definicao, todas as assinaturas colhidas -
        # mesmo que o ultimo comprovante que chegou por e-mail ainda mostre pendencia.
        if status == "Finalizado" and total_assinaturas:
            for s in doc["signatarios"]:
                s["assinou"] = True
            ja_assinaram = total_assinaturas

        dias_parado = dias_entre(doc["ultima_movimentacao"], agora)

        dias_para_limite = None
        if doc["data_limite"]:
            limite = datetime.strptime(doc["data_limite"], "%d/%m/%Y").replace(tzinfo=FUSO_LOCAL)
            dias_para_limite = dias_entre(agora, limite)

        contratos.append(
            {
                "nome": doc["nome"].replace(".pdf", ""),
                "numero": doc["numero"],
                "tipo": doc["tipo"],
                "status": status,
                "atualizado": doc["ultima_movimentacao"].strftime("%Y-%m-%dT%H:%M"),
                "atualizado_em_texto": doc["ultima_movimentacao"].strftime("%d/%m/%Y"),
                "dias_parado": dias_parado,
                "data_limite": doc["data_limite"],
                "dias_para_limite": dias_para_limite,
                "finalizado_em": doc["finalizado_em"].strftime("%d/%m/%Y") if doc["finalizado_em"] else None,
                "cancelado_em": doc["cancelado_em"].strftime("%d/%m/%Y") if doc["cancelado_em"] else None,
                "assinaturas_total": total_assinaturas,
                "assinaturas_ok": ja_assinaram,
                "signatarios": doc["signatarios"],
                "signatarios_em": doc["data_signatarios"].strftime("%d/%m/%Y") if doc["data_signatarios"] else None,
                "link_email": doc["link_email"],
                "link_clicksign": doc["link_clicksign"],
                "qtd_eventos": len(doc["eventos"]),
            }
        )

    contratos.sort(key=lambda c: c["atualizado"], reverse=True)  # mais recente primeiro, com hora para desempatar

    saida = {
        "gerado_em": agora.strftime("%d/%m/%Y %H:%M"),
        "fonte": os.path.basename(caminho_entrada),
        "regras": {
            "tipo": "tem AFE no nome \u2192 AFE; n\u00e3o tem AFE mas tem LSF \u2192 LYON; n\u00e3o tem nenhum dos dois \u2192 QPC",
            "sandbox": "excluido" if not INCLUIR_SANDBOX else "incluido",
            "parado": "contrato Em andamento sem movimenta\u00e7\u00e3o h\u00e1 mais de %d dias" % DIAS_PARA_CONSIDERAR_PARADO,
        },
        "descartados": descartados,
        "emails_lidos": len(bruto["emails"]),
        "contratos": contratos,
    }

    # ---- Conferencia obrigatoria: os totais tem que fechar --------------------
    soma_status = sum(
        1 for c in contratos if c["status"] in ORDEM_STATUS
    )
    if soma_status != len(contratos):
        erro(
            "A soma dos contratos por status (%d) nao bate com o total de contratos (%d). "
            "Nao vou gerar o painel com numero errado." % (soma_status, len(contratos))
        )

    os.makedirs(os.path.dirname(ARQUIVO_TRATADO), exist_ok=True)
    with open(ARQUIVO_TRATADO, "w", encoding="utf-8") as arquivo:
        json.dump(saida, arquivo, ensure_ascii=False, indent=2)

    # ---- Guarda o painel anterior antes de sobrescrever -----------------------
    if os.path.exists(ARQUIVO_PAINEL):
        os.makedirs(PASTA_HISTORICO, exist_ok=True)
        carimbo = datetime.fromtimestamp(os.path.getmtime(ARQUIVO_PAINEL)).strftime("%Y-%m-%d")
        shutil.copy2(ARQUIVO_PAINEL, os.path.join(PASTA_HISTORICO, "painel-%s.html" % carimbo))

    os.makedirs(os.path.dirname(ARQUIVO_PAINEL), exist_ok=True)
    modelo = os.path.join(os.path.dirname(os.path.abspath(__file__)), "modelo-painel.html")
    if not os.path.exists(modelo):
        erro("Modelo do painel nao encontrado em 05-scripts/modelo-painel.html")
    with open(modelo, encoding="utf-8") as arquivo:
        html = arquivo.read()

    html = html.replace(
        "/*DADOS_AQUI*/", json.dumps(saida, ensure_ascii=False)
    ).replace("/*PARADO_AQUI*/", str(DIAS_PARA_CONSIDERAR_PARADO)).replace(
        "/*ALERTA_PRAZO_AQUI*/", str(DIAS_ALERTA_PRAZO)
    ).replace("/*QTD_CONTRATOS*/", str(len(contratos)))

    with open(ARQUIVO_PAINEL, "w", encoding="utf-8") as arquivo:
        arquivo.write(html)

    # ---- Mostra a conta na tela, como manda a regra do projeto ---------------
    print("=" * 68)
    print("PAINEL DE CONTRATOS - CONFERENCIA")
    print("=" * 68)
    print("Fonte:                    %s" % os.path.basename(caminho_entrada))
    print("E-mails lidos:            %d" % len(bruto["emails"]))
    print("  descartados (sandbox):  %d" % descartados["sandbox"])
    print("  descartados (nao obra): %d" % descartados["nao_e_contrato_de_obra"])
    print("Contratos identificados:  %d" % len(contratos))
    print("-" * 68)
    for status in ORDEM_STATUS:
        quantidade = sum(1 for c in contratos if c["status"] == status)
        print("  %-14s %d" % (status + ":", quantidade))
    print("-" * 68)
    for tipo in ("AFE", "LYON", "QPC"):
        quantidade = sum(1 for c in contratos if c["tipo"] == tipo)
        print("  %-14s %d" % (tipo + ":", quantidade))
    print("-" * 68)
    pendentes = sum(
        c["assinaturas_total"] - c["assinaturas_ok"]
        for c in contratos
        if c["status"] == "Em andamento"
    )
    finalizados = sum(1 for c in contratos if c["status"] == "Finalizado")
    print("Assinaturas pendentes:    %d" % pendentes)
    print(
        "%% concluido:              %d / %d = %.1f%%"
        % (finalizados, len(contratos), finalizados / len(contratos) * 100)
    )
    print("=" * 68)
    print("Dados tratados: %s" % ARQUIVO_TRATADO)
    print("Painel gerado:  %s" % ARQUIVO_PAINEL)


if __name__ == "__main__":
    processar()

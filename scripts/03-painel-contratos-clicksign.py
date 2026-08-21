# -*- coding: utf-8 -*-
"""
PAINEL DE ACOMPANHAMENTO DE CONTRATOS - CLICKSIGN VIA OUTLOOK
=============================================================

O QUE ESTE SCRIPT FAZ, EM PORTUGUES SIMPLES:

Voce coloca seu e-mail como OBSERVADOR em toda assinatura da Clicksign. Por causa
disso, a Clicksign te manda um e-mail a cada movimentacao do documento. Esses
e-mails, juntos, contam a historia completa de cada contrato: quem ja assinou,
quem falta, se foi finalizado, se foi cancelado.

DE ONDE VEM CADA COISA - LEIA ISTO ANTES DE MEXER EM QUALQUER LINHA:

  FONTE VIVA (a unica que muda):  os e-mails da Clicksign no Outlook.
      Voce esta como observador em TODO contrato que ainda esta rodando, entao
      qualquer movimentacao a partir de agora chega por e-mail. Uma regra do
      Outlook joga esses e-mails na pasta "Clicksign". E daqui que sai 100% da
      informacao nova, todos os dias, para sempre.

  HISTORICO CONGELADO (nunca mais muda):  02-dados-tratados/historico-congelado.json.
      E o retrato do passado, tirado UMA VEZ SO da planilha F.SUP.G200.002 em
      17/08/2026, so para o painel nao comecar em branco. A planilha nunca mais
      sera carregada. Este arquivo e so ponto de partida: contrato que ja estava
      fechado antes de 17/08/2026 continua aparecendo por causa dele.

  Isto tudo e um paliativo ate a API da Clicksign existir. Quando ela chegar,
  a API substitui a FONTE VIVA e o historico congelado continua exatamente como
  esta, sem precisar ser refeito.

Este script pega o "dump" desses e-mails (um arquivo JSON gerado pelo comando
/painel-contratos, que le o Outlook), junta com o historico congelado, aplica as
regras de negocio que voce definiu, e gera dois arquivos:

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
# HISTORICO CONGELADO - retrato do passado, tirado da planilha em 17/08/2026.
# NAO E ATUALIZAVEL e nao deve ser regerado: a planilha nao volta mais. Serve so
# para o painel ja nascer com os contratos antigos dentro.
ARQUIVO_HISTORICO = os.path.join(PASTA_RAIZ, "02-dados-tratados", "historico-congelado.json")

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




# =============================================================================
# QUEM E QUEM - unificacao de identidade dos signatarios
# =============================================================================
#
# O PROBLEMA que isto resolve: a mesma pessoa chega com dois nomes diferentes.
# O historico escreve "Nilton LYON"; a Clicksign manda "nilt***********@lyoncapital.com.br".
# Sem unificar, o painel mostrava os dois e o ranking de quem trava saia errado:
# Nilton aparecia com 13 numa linha e 2 em outra, quando o real e 15.
#
# COMO CASAMOS: pelo pedaco do e-mail que a Clicksign NAO mascara (o comeco) mais
# o dominio. "nilt***********@lyoncapital.com.br" vira a etiqueta
# "nilt@lyoncapital.com.br". Isso e estavel mesmo se a Clicksign mudar a
# quantidade de asteriscos.
#
# Decisao do Valter em 19/08/2026: uma identidade so por pessoa. O painel mostra
# o nome (que se le) com o e-mail embaixo (que e o identificador de verdade).
PESSOAS = [
    {"nome": "Nilton LYON",   "email": "nilt@lyoncapital.com.br",      "apelidos": ["Nilton Bertuchi"]},
    {"nome": "Luiz LYON",     "email": "luiz@lyoncapital.com.br",      "apelidos": ["Luiz Guilherme"]},
    {"nome": "Fábio LYON",    "email": "fabi@lyoncapital.com.br",      "apelidos": ["Fabio Matheus"]},
    {"nome": "Flávia LYON",   "email": "flav@lyoncapital.com.br",      "apelidos": ["Flavia Lina Doi Utiyama"]},
    {"nome": "Lucas M. LYON", "email": "luca@lyoncapital.com.br",      "apelidos": ["Lucas Marrucci"]},
    {"nome": "Hassan QPC",    "email": "hlue@qpc.com.br",              "apelidos": ["Hassan Fair Luedy"]},
    {"nome": "Felippe QPC",   "email": "fpam@qpc.com.br",              "apelidos": ["Felippe Pamponet Esquivel"]},
    {"nome": "Emerson ABR",   "email": "emer@abrgerenciamento.com",    "apelidos": ["Emerson Leal"]},
    {"nome": "Lucas G. QPC",  "email": None,                           "apelidos": ["Lucas Maron Grimaldi"]},
    {"nome": "Sergio QPC",    "email": None,                           "apelidos": []},
]

# Quem saiu da obra: sai das PENDENCIAS, mas o que ja assinou continua no painel.
# Use isto quando a pessoa realmente participou do fluxo e depois deixou a obra.
# Hoje esta vazio - o caso do Sergio virou outra coisa, explicada logo abaixo.
SAIRAM_DA_OBRA = set()

# ---------------------------------------------------------------------------
# COLUNA FANTASMA DA PLANILHA - decisao do Valter em 20/08/2026
# ---------------------------------------------------------------------------
#
# O Sergio saiu da obra em maio/2026, mas a planilha das colegas continuou com
# uma coluna no nome dele, e essa coluna continuou sendo preenchida. Ele nao sai
# so das pendencias: ele sai INTEIRO do historico congelado, porque a evidencia
# diz que ele nao e signatario de coisa nenhuma na Clicksign.
#
# A prova, checada em 20/08/2026 - nos CINCO contratos em que da para cruzar as
# duas fontes, a Clicksign nao tem o Sergio no fluxo:
#
#     contrato 084  ->  Clicksign lista 10 signatarios; a planilha lista 11.
#                       O 11o e o Sergio, com data 03/08/2026.
#     mesma coisa em 079, 050, 081 e 065.
#
# Mais tres indicios na mesma direcao:
#
#   1. ZERO mencoes ao Sergio em 76 e-mails da Clicksign. Nenhuma.
#   2. 47 das 70 datas dele sao copia exata da data da coluna vizinha - isso e
#      arrasto de celula no Excel, nao assinatura.
#   3. As DUAS unicas datas impossiveis da planilha inteira (17/03/2027 e
#      02/06/2028) estao na coluna dele. Nenhuma outra coluna tem data furada.
#
# O que NAO consegui provar, e fica registrado por honestidade: 33 das 70 datas
# tem data propria, quase todas entre fevereiro e maio, quando ele estava na
# obra. Podem ser assinaturas reais de contratos para os quais nao tenho e-mail.
# O Valter foi consultado com essa ressalva e escolheu tirar completamente.
#
# Nada foi destruido: o 02-dados-tratados/historico-congelado.json continua com
# os registros dele. A remocao acontece aqui, na leitura, e desfazer e apagar
# uma linha.
NAO_ESTAO_NA_CLICKSIGN = {"Sergio QPC"}


def chave_pessoa(quem):
    """
    Reduz um signatario a uma etiqueta unica, venha ele do historico ou do e-mail.

    "nilt***********@lyoncapital.com.br" -> "nilt@lyoncapital.com.br"
    "Nilton Bertuchi"                    -> "nilton bertuchi"
    """
    texto = str(quem).strip()
    if "@" in texto:
        usuario, _, dominio = texto.partition("@")
        return (usuario.split("*")[0].strip().lower() + "@" + dominio.strip().lower())
    return re.sub(r"\s+", " ", sem_acento(texto).lower()).strip()


def montar_indice_pessoas():
    indice = {}
    for pessoa in PESSOAS:
        if pessoa["email"]:
            indice[pessoa["email"].lower()] = pessoa
        indice[chave_pessoa(pessoa["nome"])] = pessoa
        for apelido in pessoa["apelidos"]:
            indice[chave_pessoa(apelido)] = pessoa
    return indice


INDICE_PESSOAS = montar_indice_pessoas()


def identificar_pessoa(quem):
    """Devolve (nome_para_mostrar, email_ou_None). Quem nao esta na tabela passa direto."""
    pessoa = INDICE_PESSOAS.get(chave_pessoa(quem))
    if pessoa:
        return pessoa["nome"], pessoa["email"]
    # Nao esta na tabela: e um fornecedor ou alguem novo. Mostra como veio.
    return str(quem).strip(), (str(quem).strip() if "@" in str(quem) else None)


def normalizar_signatarios(lista):
    """
    Unifica identidade, descarta coluna fantasma e tira quem saiu das pendencias.

    Se a mesma pessoa aparecer duas vezes no mesmo contrato (um registro vindo do
    historico e outro do e-mail), vale ASSINADO - porque assinatura nao se desfaz.
    """
    por_pessoa = {}
    for assinatura in lista:
        nome, email = identificar_pessoa(assinatura["quem"])
        if nome in NAO_ESTAO_NA_CLICKSIGN:
            # Coluna fantasma: nao e signatario do documento. Sai inteiro, tenha
            # a planilha marcado ASSINADO ou nao. Ver o bloco de comentario la em
            # cima com a evidencia.
            continue
        if nome in SAIRAM_DA_OBRA and not assinatura.get("assinou"):
            continue
        atual = por_pessoa.get(nome)
        if atual is None:
            por_pessoa[nome] = {
                "quem": nome,
                "email": email,
                "papel": assinatura.get("papel") or "signatario",
                "assinou": bool(assinatura.get("assinou")),
                "data": assinatura.get("data"),
            }
        else:
            atual["assinou"] = atual["assinou"] or bool(assinatura.get("assinou"))
            atual["data"] = atual["data"] or assinatura.get("data")
            atual["email"] = atual["email"] or email
    return list(por_pessoa.values())


def chave_do_contrato(nome_documento):
    """
    Etiqueta que liga um documento da Clicksign a uma linha do historico.

    Mesma regra usada para montar o historico: o que importa e ser aditivo ou
    nao, e o numero do contrato. Assim "63.CT-LSF-G200-063-26 - INOSERVICE.pdf"
    (nome do e-mail) e "CT-LSF-G200-063.26" (historico) viram os dois "CT-063".
    """
    texto = sem_acento(str(nome_documento)).upper()
    aditivo = re.search(r"ADIT\.?\s*V?\.?\s*(\d{1,2})", texto)
    numero = re.search(r"G200[-\s.]*(\d{2,3})", texto)
    if not numero:
        return None
    if aditivo:
        return "ADIT%s-%s" % (aditivo.group(1).zfill(2), numero.group(1).zfill(3))
    return "CT-%s" % numero.group(1).zfill(3)


def ordem_pelo_numero(chave):
    """
    Numero para ordenar a lista do jeito que o Valter le: 001, 002, 003...

    O aditivo cai logo depois do contrato de origem, e nao no fim da lista:
      CT-046 -> 4600 ; ADIT01-046 -> 4601 ; ADIT02-046 -> 4602
    """
    sufixo = re.search(r"#doc(\d+)$", chave)
    chave = re.sub(r"#doc\d+$", "", chave)
    achado = re.search(r"(\d{2,3})$", chave)
    numero = int(achado.group(1)) if achado else 0
    aditivo = re.match(r"ADIT(\d{2})", chave)
    base = numero * 1000 + (int(aditivo.group(1)) if aditivo else 0) * 10
    return base + (int(sufixo.group(1)) if sufixo else 0)


def carregar_historico():
    """Le o retrato congelado do passado. Ele nao muda entre uma rodada e outra."""
    if not os.path.exists(ARQUIVO_HISTORICO):
        return None
    with open(ARQUIVO_HISTORICO, encoding="utf-8") as arquivo:
        return json.load(arquivo)


def processar():
    agora = (
        datetime.fromisoformat(DATA_REFERENCIA).astimezone(FUSO_LOCAL)
        if DATA_REFERENCIA
        else datetime.now(FUSO_LOCAL)
    )

    # =========================================================================
    # PONTO DE PARTIDA - o historico congelado (passado, nao muda mais)
    # =========================================================================
    base = carregar_historico()
    if not base:
        erro(
            "Nao achei o historico congelado em 02-dados-tratados/historico-congelado.json.\n"
            "Este arquivo nao se regera sozinho: ele e o retrato de 17/08/2026 e vem\n"
            "junto com os scripts no repositorio. Baixe-o de novo do GitHub\n"
            "(dados/historico-congelado.json) - nao tente reconstruir pela planilha."
        )

    data_historico = datetime.strptime(base["data_do_retrato"], "%Y-%m-%d").replace(tzinfo=FUSO_LOCAL)
    registros = {}
    for item in base["contratos"]:
        item["signatarios"] = normalizar_signatarios(item["signatarios"])
        assinados = [s for s in item["signatarios"] if s["assinou"]]
        registros[item["chave"]] = {
            "fonte": "histórico",
            "data_fonte": data_historico,
            "chave": item["chave"],
            "identificacao": item["identificacao"],
            "razao_social": item["razao_social"],
            "servico": item["servico"],
            "cnpj": item["cnpj"],
            "aditivo": item["aditivo"],
            "tipo": item["tipo"],
            "status": "Finalizado" if item["concluido"] else "Em andamento",
            "signatarios": item["signatarios"],
            "assinaturas_ok": len(assinados),
            "assinaturas_total": len(item["signatarios"]),
            "dias_parado": item["dias_sem_assinatura"],
            "data_cadastro": item["data_cadastro"],
            "data_limite": None,
            "finalizado_em": None,
            "cancelado_em": None,
            "link_email": "",
            "link_clicksign": "",
            "qtd_eventos": 0,
            "historico": [],
        }

    # =========================================================================
    # FONTE 2 - os e-mails da Clicksign (o que aconteceu depois do retrato)
    # =========================================================================
    caminho_entrada = achar_arquivo_entrada()
    with open(caminho_entrada, encoding="utf-8") as arquivo:
        bruto = json.load(arquivo)

    if "emails" not in bruto:
        erro("O arquivo %s nao tem a chave 'emails'. Formato inesperado." % caminho_entrada)

    if not bruto["emails"]:
        erro(
            "O dump nao tem NENHUM e-mail. Isso nao e um painel vazio - e uma busca que falhou.\n"
            "Causa mais provavel: a busca no Outlook usou o parametro 'order'.\n"
            "Sem 'folderName', o 'order' restringe a busca a Caixa de Entrada, e os e-mails\n"
            "da Clicksign nao estao nela. Refaca a busca SEM 'order'. NAO publique."
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

        chave_arquivo = chave_documento(nome)
        evento = identificar_evento(email.get("assunto", ""), email.get("corpo", ""))
        recebido = ler_data(email["recebido_em"])

        doc = documentos.setdefault(chave_arquivo, {
            "nome": nome, "eventos": [], "signatarios": [], "data_signatarios": None,
            "data_limite": None, "ultima_movimentacao": recebido,
            "link_email": email.get("web_link") or "", "link_clicksign": "",
            "finalizado_em": None, "cancelado_em": None,
        })

        doc["eventos"].append({"tipo": evento, "em": recebido})
        if recebido > doc["ultima_movimentacao"]:
            doc["ultima_movimentacao"] = recebido
            if email.get("web_link"):
                doc["link_email"] = email["web_link"]

        if evento == "comprovante":
            lista = parsear_signatarios(email.get("corpo", ""))
            if lista and (doc["data_signatarios"] is None or recebido > doc["data_signatarios"]):
                doc["signatarios"] = normalizar_signatarios([
                    {"quem": s["quem"], "papel": s["papel"], "assinou": s["assinou"], "data": None}
                    for s in lista
                ])
                doc["data_signatarios"] = recebido

        if evento == "finalizado" and (doc["finalizado_em"] is None or recebido < doc["finalizado_em"]):
            doc["finalizado_em"] = recebido
        if evento == "cancelado":
            doc["cancelado_em"] = recebido

        achado_prazo = PADRAO_DATA_LIMITE.search(email.get("corpo", ""))
        if achado_prazo:
            doc["data_limite"] = achado_prazo.group(1)

        achado_link = re.search(r"https://app\.clicksign\.com/\S+?(?=[\s.]|$)", email.get("corpo", ""))
        if achado_link and not doc["link_clicksign"]:
            doc["link_clicksign"] = achado_link.group(0)

    # Um contrato pode ter varios documentos na Clicksign: o 065 foi cancelado
    # como "Limpeza Vicinal" e reemitido como "Limpeza"; o 063 expirou e foi
    # reenviado. Agrupamos por contrato para depois decidir qual e o vigente.
    por_contrato = {}
    for doc in documentos.values():
        chave = chave_do_contrato(doc["nome"])
        if not chave:
            continue
        por_contrato.setdefault(chave, []).append(doc)

    # =========================================================================
    # FUSAO DAS DUAS FONTES - por PESSOA, nao por lista inteira
    # =========================================================================
    #
    # A versao anterior escolhia uma fonte e jogava a outra fora. Isso quebrou o
    # contrato 081: o historico tem uma coluna generica "Contratada" e o e-mail
    # tem as pessoas de verdade (luis***@mandrade.com.br e pinh*@mandrade.com.br).
    # Como o historico era mais novo, ele vencia e as pessoas reais sumiam.
    #
    # Agora juntamos as duas listas pessoa a pessoa. Duas regras cuidam do resto:
    #
    #   1. ASSINATURA NAO SE DESFAZ. Se qualquer uma das fontes diz que fulano
    #      assinou, ele assinou. Isso resolve sozinho o caso do 084, em que o
    #      e-mail antigo dizia "falta o Nilton" e o historico dizia que nao.
    #
    #   2. "Contratada" e um CARGO, nao uma pessoa. Quando o e-mail diz quem e o
    #      fornecedor, o cargo generico sai de cena. O contrato 050 mostra por que
    #      isso importa: ele tem DUAS pessoas como contratada (Eliana e Alan), e
    #      o historico so tem uma coluna.
    def fundir_signatarios(do_historico, do_email):
        tem_contratada_real = any(
            "contratada" in (s.get("papel") or "").lower() for s in do_email
        )
        juntos = []
        for s in do_historico:
            if tem_contratada_real and s["quem"].strip().lower() == "contratada":
                continue
            juntos.append(s)
        return normalizar_signatarios(juntos + list(do_email))

    def situacao_do_documento(assinaturas, encerrou_em, cancelado, recusado, historico_concluiu):
        """
        Decide o status DEPOIS que as duas listas ja foram fundidas.

        Fazer isso antes era o bug: o e-mail de 14/08 do contrato 084 dizia que
        faltava o Nilton, entao o documento era marcado "Expirado"; o historico
        completava a assinatura logo em seguida e o rotulo errado ficava.
        """
        pendentes = [s for s in assinaturas if not s["assinou"]]
        if cancelado:
            return "Cancelado"
        if recusado:
            return "Recusado"
        if assinaturas and not pendentes:
            return "Finalizado"
        if historico_concluiu and not pendentes:
            return "Finalizado"
        if encerrou_em and pendentes:
            # A Clicksign encerra por prazo vencido e manda o mesmo e-mail de
            # "finalizado". Com pendencia na lista, isso e prazo estourado.
            return "Expirado"
        if encerrou_em and not assinaturas:
            # Encerrou e nao temos a lista: nao da para dizer se todos assinaram.
            return "Finalizado"
        return "Em andamento"

    for chave, docs in por_contrato.items():
        docs.sort(key=lambda d: d["ultima_movimentacao"], reverse=True)
        do_historico = registros.get(chave)

        for posicao, doc in enumerate(docs):
            recusado = any(e["tipo"] == "recusado" for e in doc["eventos"])
            encerrou_em = doc["finalizado_em"] or doc["cancelado_em"]

            # ---- Este documento e o mesmo que o historico descreve? ------------
            # Se o documento do e-mail JA ENCERROU e o historico, DEPOIS dele,
            # mostra o contrato ainda em andamento, sao documentos DIFERENTES: o
            # antigo morreu e um novo foi disparado. Foi o 063, que expirou em
            # 12/07 e foi reenviado. Quando o encerramento e o registro mais
            # recente, e o mesmo documento que chegou ao fim - caso do 050.
            documento_diferente = (
                do_historico is not None
                and encerrou_em is not None
                and encerrou_em < do_historico["data_fonte"]
                and do_historico["status"] == "Em andamento"
            )
            e_o_vigente = posicao == 0 and not documento_diferente

            assinaturas = list(doc["signatarios"])
            registro = {
                "fonte": "e-mail",
                "data_fonte": doc["ultima_movimentacao"],
                "chave": chave,
                "identificacao": doc["nome"].replace(".pdf", ""),
                "aditivo": chave.startswith("ADIT"),
                "dias_parado": dias_entre(doc["ultima_movimentacao"], agora),
                "data_limite": doc["data_limite"],
                "finalizado_em": doc["finalizado_em"].strftime("%d/%m/%Y") if doc["finalizado_em"] else None,
                "cancelado_em": doc["cancelado_em"].strftime("%d/%m/%Y") if doc["cancelado_em"] else None,
                "link_email": doc["link_email"],
                "link_clicksign": doc["link_clicksign"],
                "qtd_eventos": len(doc["eventos"]),
                "signatarios_em": doc["data_signatarios"].strftime("%d/%m/%Y") if doc["data_signatarios"] else None,
                "tipo": classificar_tipo(doc["nome"]),
                "historico": [],
                "encerrado": False,
                # Quantas pessoas a CLICKSIGN declarou no fluxo deste documento.
                # So existe quando um "Comprovante de assinatura" foi lido - e o
                # comprovante lista TODO MUNDO do fluxo, tenha assinado ou nao.
                # E a nossa unica referencia externa para conferir o historico.
                "clicksign_declarou": (
                    len(doc["signatarios"]) if doc["data_signatarios"] else None
                ),
            }

            if e_o_vigente and do_historico is not None:
                assinaturas = fundir_signatarios(do_historico["signatarios"], assinaturas)
                for campo in ("razao_social", "servico", "cnpj", "tipo", "data_cadastro"):
                    registro[campo] = do_historico.get(campo)
                registro["identificacao"] = do_historico["identificacao"]
                # O historico esta parado em 17/08/2026 e o e-mail so anda para a
                # frente. Enquanto o ultimo e-mail deste contrato for anterior ao
                # retrato, quem sabe mais sobre ele ainda e o retrato.
                if do_historico["data_fonte"] > doc["ultima_movimentacao"]:
                    registro["dias_parado"] = do_historico["dias_parado"]
                    registro["data_fonte"] = do_historico["data_fonte"]
                    registro["fonte"] = "histórico + e-mail"
                else:
                    registro["fonte"] = "e-mail (sobre o histórico)"
                historico_concluiu = do_historico["status"] == "Finalizado"
                do_historico = None
            else:
                historico_concluiu = False
                if do_historico is not None:
                    for campo in ("razao_social", "servico", "cnpj"):
                        registro[campo] = do_historico.get(campo)

            registro["signatarios"] = assinaturas
            registro["status"] = situacao_do_documento(
                assinaturas, encerrou_em, doc["cancelado_em"], recusado, historico_concluiu
            )
            if registro["status"] == "Finalizado":
                for s in assinaturas:
                    s["assinou"] = True

            if e_o_vigente:
                registros[chave] = registro
            else:
                # Um documento encerrado que foi SUBSTITUIDO por outro do mesmo
                # contrato nao foi concluido - se tivesse sido, ninguem teria
                # disparado um novo. Sem a lista de assinaturas nao da para
                # provar isso pelo conteudo, mas a existencia do substituto
                # prova pelo processo. Foi exatamente o caso do 063, que o
                # painel mostrava como fechado e o Valter corrigiu.
                if registro["status"] == "Finalizado":
                    registro["status"] = "Expirado"
                # Documento que ja encerrou e nao e o vigente: linha propria.
                # O Valter pediu isso - cancelamento e expiracao sao fatos do
                # processo e nao podem sumir do painel.
                registro["encerrado"] = True
                registros["%s#doc%d" % (chave, posicao + 1)] = registro

    if not registros:
        erro("Nem o historico congelado nem os e-mails renderam contrato nenhum.")

    # =========================================================================
    # CONFERENCIA CONTRA A CLICKSIGN - trava criada em 21/08/2026
    # =========================================================================
    #
    # POR QUE ISTO EXISTE, EM PORTUGUES SIMPLES:
    #
    # A planilha que virou o historico congelado tinha uma coluna com o nome de
    # uma pessoa que NAO E SIGNATARIA de nada na Clicksign (o Sergio, que saiu da
    # obra em maio/2026 e cuja coluna continuou sendo preenchida). Resultado: o
    # painel mostrava 11 pessoas onde a Clicksign so tem 10, e ninguem percebeu
    # por dois dias - foi o Valter quem viu.
    #
    # A regra que pega esse tipo de erro e simples e barata: o "Comprovante de
    # assinatura" da Clicksign lista TODAS as pessoas do fluxo, tenham assinado
    # ou nao. Entao, num contrato em que lemos um comprovante, o painel NUNCA
    # pode ter mais gente do que a Clicksign declarou. Se tiver, o excesso veio
    # do historico e e contaminacao.
    #
    # Por que isto PARA a rodada em vez de so avisar: um painel que inventa
    # signatario e pior que painel nenhum, porque as pessoas cobram assinatura de
    # quem nao tem que assinar. E o conserto e uma linha - acrescentar o nome em
    # NAO_ESTAO_NA_CLICKSIGN, la em cima, depois de conferir.
    #
    # Contar a MENOS e normal e nao para nada: um contrato pode ter ganhado
    # signatario depois do comprovante que lemos.
    contaminados = []
    for chave, reg in registros.items():
        declarou = reg.get("clicksign_declarou")
        if not declarou:
            continue
        no_painel = len(reg.get("signatarios") or [])
        if no_painel > declarou:
            do_email = {s["quem"] for s in reg.get("signatarios") or [] if s.get("quem")}
            contaminados.append((reg.get("identificacao") or chave, declarou, no_painel, sorted(do_email)))

    if contaminados:
        linhas = [
            "O painel tem MAIS signatarios do que a Clicksign declarou. Isso e",
            "contaminacao vinda do historico congelado - alguem esta na planilha",
            "e nao esta no fluxo da Clicksign.",
            "",
        ]
        for identificacao, declarou, no_painel, pessoas in contaminados:
            linhas.append("  %s: Clicksign diz %d, painel diz %d" % (identificacao, declarou, no_painel))
            linhas.append("     no painel: %s" % ", ".join(pessoas))
        linhas += [
            "",
            "O QUE FAZER: abra o ultimo 'Comprovante de assinatura' desse contrato",
            "no Outlook, veja quem a Clicksign lista, e descubra quem sobra. Se for",
            "alguem que nao participa mais do fluxo, acrescente o nome em",
            "NAO_ESTAO_NA_CLICKSIGN, no topo deste script, com a evidencia ao lado.",
            "NAO publique um painel que inventa signatario.",
        ]
        erro("\n".join(linhas))

    # ---- Guarda-corpo: queda brusca vs a rodada anterior ---------------------
    if os.path.exists(ARQUIVO_TRATADO):
        try:
            with open(ARQUIVO_TRATADO, encoding="utf-8") as anterior:
                antes = len(json.load(anterior).get("contratos", []))
        except Exception:
            antes = 0
        if antes and len(registros) < antes * 0.7:
            erro(
                "A rodada anterior tinha %d contratos e esta tem so %d - queda de %.0f%%.\n"
                "Contrato nao some sozinho do painel. Isso e sintoma de coleta incompleta.\n"
                "NAO vou sobrescrever o painel." % (antes, len(registros), 100 * (1 - len(registros) / antes))
            )

    # =========================================================================
    # MONTAGEM FINAL
    # =========================================================================
    contratos = []
    for reg in registros.values():
        dias_para_limite = None
        if reg.get("data_limite"):
            limite = datetime.strptime(reg["data_limite"], "%d/%m/%Y").replace(tzinfo=FUSO_LOCAL)
            dias_para_limite = dias_entre(agora, limite)

        rotulo = reg["identificacao"]
        if reg.get("razao_social"):
            rotulo = "%s — %s" % (reg["identificacao"], reg["razao_social"])

        contratos.append({
            "nome": rotulo,
            "numero": reg["chave"],
            "ordem_numero": ordem_pelo_numero(reg["chave"]),
            "encerrado": reg.get("encerrado", False),
            "tipo": reg.get("tipo") or "LYON",
            "aditivo": reg.get("aditivo", False),
            "fonte": reg["fonte"],
            "servico": reg.get("servico"),
            "cnpj": reg.get("cnpj"),
            "status": reg["status"],
            "atualizado": reg["data_fonte"].strftime("%Y-%m-%dT%H:%M"),
            "atualizado_em_texto": reg["data_fonte"].strftime("%d/%m/%Y"),
            "dias_parado": reg["dias_parado"] if reg["dias_parado"] is not None else 0,
            "data_limite": reg.get("data_limite"),
            "dias_para_limite": dias_para_limite,
            "finalizado_em": reg.get("finalizado_em"),
            "cancelado_em": reg.get("cancelado_em"),
            "clicksign_declarou": reg.get("clicksign_declarou"),
            "assinaturas_total": len(reg["signatarios"]),
            "assinaturas_ok": sum(1 for s in reg["signatarios"] if s["assinou"]),
            "signatarios": reg["signatarios"],
            "signatarios_em": reg.get("signatarios_em"),
            "link_email": reg.get("link_email", ""),
            "link_clicksign": reg.get("link_clicksign", ""),
            "qtd_eventos": reg.get("qtd_eventos", 0),
            "historico": reg.get("historico", []),
        })

    contratos.sort(key=lambda c: (c["status"] == "Finalizado", c["atualizado"]), reverse=False)
    contratos.sort(key=lambda c: (c["status"] != "Em andamento", -(c["dias_parado"] or 0)))

    saida = {
        "gerado_em": agora.strftime("%d/%m/%Y %H:%M"),
        "fonte_viva": "e-mails da Clicksign no Outlook (pasta Clicksign) - %s"
                      % os.path.basename(caminho_entrada),
        "historico_congelado": "%s, retrato unico de %s, nao recebe mais atualizacao"
                               % (base["fonte_original"], base["data_do_retrato_texto"]),
        "regras": {
            "tipo": "o historico cobre apenas contratos LYON; AFE e QPC so aparecem pelo e-mail",
            "conflito": "o historico parou em %s; tudo que aconteceu depois vem do e-mail, "
                        "e assinatura registrada em qualquer uma das duas nao se desfaz"
                        % base["data_do_retrato_texto"],
            "nao_se_aplica": "signatario marcado N.A ou - sai da conta: nao aparece como pendente",
            "parado": "contrato Em andamento sem movimentação há mais de %d dias" % DIAS_PARA_CONSIDERAR_PARADO,
        },
        "descartados": descartados,
        "emails_lidos": len(bruto["emails"]),
        "so_do_historico": sum(1 for c in contratos if c["fonte"] == "histórico"),
        "com_email": sum(1 for c in contratos if "mail" in c["fonte"]),
        "retrato_historico": base["data_do_retrato_texto"],
        "contratos": contratos,
    }

    os.makedirs(os.path.dirname(ARQUIVO_TRATADO), exist_ok=True)
    with open(ARQUIVO_TRATADO, "w", encoding="utf-8") as arquivo:
        json.dump(saida, arquivo, ensure_ascii=False, indent=2)

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

    html = (
        html.replace("/*DADOS_AQUI*/", json.dumps(saida, ensure_ascii=False))
        .replace("/*PARADO_AQUI*/", str(DIAS_PARA_CONSIDERAR_PARADO))
        .replace("/*ALERTA_PRAZO_AQUI*/", str(DIAS_ALERTA_PRAZO))
        .replace("/*QTD_CONTRATOS*/", str(len(contratos)))
    )
    with open(ARQUIVO_PAINEL, "w", encoding="utf-8") as arquivo:
        arquivo.write(html)

    # ---- Mostra a conta -----------------------------------------------------
    print("=" * 70)
    print("PAINEL DE CONTRATOS - CONFERENCIA")
    print("=" * 70)
    print("Historico congelado:      retrato de %s (nao muda mais)" % base["data_do_retrato_texto"])
    print("E-mails (fonte viva):     %s" % os.path.basename(caminho_entrada))
    print("  lidos:                  %d" % len(bruto["emails"]))
    print("  descartados (sandbox):  %d" % descartados["sandbox"])
    print("  descartados (nao obra): %d" % descartados["nao_e_contrato_de_obra"])
    print("-" * 70)
    print("Contratos no painel:      %d" % len(contratos))
    print("  so no historico:        %d  (sem nenhum e-mail desde o retrato)" % saida["so_do_historico"])
    print("  com noticia por e-mail: %d" % saida["com_email"])
    print("  sendo aditivos:         %d" % sum(1 for c in contratos if c["aditivo"]))
    print("-" * 70)
    for status in ORDEM_STATUS + ["Expirado"]:
        quantidade = sum(1 for c in contratos if c["status"] == status)
        if quantidade:
            print("  %-16s %d" % (status + ":", quantidade))
    conferidos = [c for c in contratos if c.get("clicksign_declarou")]
    print("Conferidos contra a Clicksign: %d contratos, todos batendo" % len(conferidos))
    print("-" * 70)
    pendentes = sum(
        c["assinaturas_total"] - c["assinaturas_ok"]
        for c in contratos if c["status"] == "Em andamento"
    )
    finalizados = sum(1 for c in contratos if c["status"] == "Finalizado")
    print("Assinaturas pendentes:    %d" % pendentes)
    print("%% concluido:              %d / %d = %.1f%%"
          % (finalizados, len(contratos), finalizados / len(contratos) * 100))
    print("=" * 70)
    print("Dados tratados: %s" % ARQUIVO_TRATADO)
    print("Painel gerado:  %s" % ARQUIVO_PAINEL)


if __name__ == "__main__":
    processar()

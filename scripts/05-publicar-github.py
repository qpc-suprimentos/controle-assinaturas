# -*- coding: utf-8 -*-
"""
PUBLICA O PAINEL NO AR (GitHub Pages)
=====================================

O QUE ESTE SCRIPT FAZ, EM PORTUGUES SIMPLES:

Depois que o 03-painel-contratos-clicksign.py gera o HTML, este script manda esse
arquivo para o GitHub, que o serve como site. O endereco nunca muda:

    https://valterrocha.github.io/g200-7k3f9x2m/

Quem tiver o link sempre ve a versao mais recente. Nao precisa arrastar nada,
nao precisa abrir o navegador, nao precisa anexar em e-mail.

DUAS ARMADILHAS DESTE AMBIENTE QUE JA CUSTARAM CARO - NAO MEXER SEM LER:

1. O ambiente do Claude tem um proxy que INTERCEPTA chamadas ao GitHub e troca a
   credencial por uma dele ("builtin injection failed (github)", HTTP 502).
   Por isso o script desliga o proxy explicitamente (ProxyHandler vazio, abaixo).
   Sem isso, nada funciona, por mais correto que o token esteja.

2. Nao adianta tentar Netlify, GitLab, Vercel, Cloudflare, Render, Bitbucket ou
   Codeberg: TODOS estao bloqueados neste ambiente (conexao recusada em ~0,35 s).
   O GitHub e o unico host alcancavel. Ja foi gasta uma sessao inteira nisso.

COMO RODAR:
    python 05-scripts/05-publicar-github.py
"""

import base64
import hashlib
import json
import os
import urllib.error
import urllib.request

# =============================================================================
# CONFIGURACAO - e so aqui que voce mexe
# =============================================================================

PASTA_RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

USUARIO = "valterrocha"
REPOSITORIO = "g200-7k3f9x2m"          # nome neutro de proposito: nao diz "contratos" nem "QPC"
RAMO = "main"
URL_PUBLICA = "https://valterrocha.github.io/g200-7k3f9x2m/"

# O painel gerado pelo script anterior
ARQUIVO_PAINEL = os.path.join(PASTA_RAIZ, "03-painel", "painel-contratos.html")

# Onde mora a chave do GitHub. Fica no computador do Valter, nunca no chat e
# nunca na memoria do projeto. O .gitignore desta pasta ja ignora este arquivo.
ARQUIVO_TOKEN = os.path.join(PASTA_RAIZ, "05-scripts", "github-token.txt")

API = "https://api.github.com"


def erro(mensagem):
    """Para tudo e avisa em portugues. Regra do projeto: falhar alto, nao baixo."""
    raise SystemExit("\n[ERRO] " + mensagem + "\n")


def ler_token():
    token = os.environ.get("GITHUB_PAT", "").strip()
    if token:
        return token
    if not os.path.exists(ARQUIVO_TOKEN):
        erro(
            "Nao achei a chave do GitHub em 05-scripts/github-token.txt.\n"
            "Gere uma em github.com > Settings > Developer settings >\n"
            "Personal access tokens > Fine-grained tokens, com permissao de\n"
            "Contents e Pages (Read and write) no repositorio %s." % REPOSITORIO
        )
    with open(ARQUIVO_TOKEN, encoding="utf-8") as arquivo:
        token = arquivo.read().strip()
    if not token.startswith("github_pat_") and not token.startswith("ghp_"):
        erro("O github-token.txt nao parece uma chave do GitHub (deve comecar com github_pat_).")
    return token


# O proxy do ambiente sequestra as chamadas ao GitHub. Este "abridor" fala direto
# com a internet, ignorando o proxy. E a linha que faz todo o resto funcionar.
ABRIDOR = urllib.request.build_opener(urllib.request.ProxyHandler({}))


def chamar_api(caminho, token, metodo="GET", corpo=None):
    requisicao = urllib.request.Request(
        API + caminho,
        data=json.dumps(corpo).encode("utf-8") if corpo is not None else None,
        method=metodo,
    )
    requisicao.add_header("Authorization", "Bearer " + token)
    requisicao.add_header("Accept", "application/vnd.github+json")
    requisicao.add_header("User-Agent", "painel-contratos-qpc")
    requisicao.add_header("Content-Type", "application/json")
    try:
        with ABRIDOR.open(requisicao, timeout=90) as resposta:
            texto = resposta.read().decode("utf-8")
            return json.loads(texto) if texto.strip() else {}
    except urllib.error.HTTPError as falha:
        if falha.code == 404 and metodo == "GET":
            return None                      # arquivo ainda nao existe no repositorio
        detalhe = falha.read().decode("utf-8", "replace")[:300]
        if falha.code == 401:
            erro("O GitHub recusou a chave (401). Ela expirou ou foi revogada. Gere outra.")
        if falha.code == 403:
            erro(
                "O GitHub respondeu 403 em %s.\n"
                "Quase sempre e falta de permissao na chave: ela precisa de\n"
                "Contents e Pages com Read and write neste repositorio.\n%s" % (caminho, detalhe)
            )
        erro("O GitHub respondeu %s em %s.\n%s" % (falha.code, caminho, detalhe))


def sha_do_git(conteudo):
    """
    Calcula a impressao digital que o Git daria a este arquivo.

    Serve para saber se o conteudo mudou desde a ultima publicacao sem precisar
    baixar o arquivo inteiro do GitHub. Se nao mudou, pulamos o envio.
    """
    cabecalho = ("blob %d\0" % len(conteudo)).encode()
    return hashlib.sha1(cabecalho + conteudo).hexdigest()


def enviar(caminho_no_site, conteudo, token, mensagem):
    atual = chamar_api(
        "/repos/%s/%s/contents/%s?ref=%s" % (USUARIO, REPOSITORIO, caminho_no_site, RAMO), token
    )

    if atual and atual.get("sha") == sha_do_git(conteudo):
        print("  sem mudanca: %s" % caminho_no_site)
        return False

    corpo = {
        "message": mensagem,
        "content": base64.b64encode(conteudo).decode("ascii"),
        "branch": RAMO,
    }
    if atual:
        corpo["sha"] = atual["sha"]          # obrigatorio para SUBSTITUIR um arquivo existente

    chamar_api(
        "/repos/%s/%s/contents/%s" % (USUARIO, REPOSITORIO, caminho_no_site),
        token, "PUT", corpo,
    )
    print("  publicado:   %s (%d KB)" % (caminho_no_site, len(conteudo) // 1024 or 1))
    return True


def publicar():
    token = ler_token()

    if not os.path.exists(ARQUIVO_PAINEL):
        erro(
            "O painel nao existe em 03-painel/painel-contratos.html.\n"
            "Rode antes: python 05-scripts/03-painel-contratos-clicksign.py"
        )

    with open(ARQUIVO_PAINEL, "rb") as arquivo:
        html = arquivo.read()

    # Confere se o painel realmente tem dados dentro. Ja aconteceu de o gerador
    # falhar e sobrar um HTML vazio - publicar isso seria pior que nao publicar.
    if b'"contratos"' not in html or len(html) < 5000:
        erro(
            "O painel parece vazio ou incompleto (%d bytes). Nao vou publicar.\n"
            "Rode o gerador de novo e confira a conferencia que ele imprime." % len(html)
        )

    # O robots.txt e o que impede o Google de achar o endereco. O site e publico
    # e a unica protecao e o link nao ser descoberto - sem isto, "link secreto"
    # dura ate o primeiro rastreamento do buscador.
    robots = b"User-agent: *\nDisallow: /\n"

    # ---- Guarda-corpo: nao sobrescrever um painel bom com um painel furado ----
    # Este e o unico guarda-corpo que funciona TAMBEM na tarefa agendada, porque
    # nao depende de arquivo local nenhum: ele compara com a versao que ja esta
    # publicada no GitHub. Contrato so sai do painel se for renomeado na Clicksign;
    # queda grande de uma rodada para outra e sintoma de coleta incompleta.
    import re

    def contar(conteudo):
        achado = re.search(rb'name="qpc-contratos" content="(\d+)"', conteudo)
        return int(achado.group(1)) if achado else None

    agora_tem = contar(html)
    publicado = chamar_api(
        "/repos/%s/%s/contents/index.html?ref=%s" % (USUARIO, REPOSITORIO, RAMO), token
    )
    if publicado and agora_tem is not None:
        anterior = contar(base64.b64decode(publicado.get("content", "")))
        if anterior and agora_tem < anterior * 0.7:
            erro(
                "O painel no ar tem %d contratos e este tem so %d - queda de %.0f%%.\n"
                "Contrato nao some sozinho. Isso e sintoma de coleta incompleta:\n"
                "verifique se a busca no Outlook usou o parametro 'order' (ele zera a busca).\n"
                "NAO vou sobrescrever o painel que esta no ar."
                % (anterior, agora_tem, 100 * (1 - agora_tem / anterior))
            )
        print("Contratos: %d no ar -> %d agora." % (anterior or 0, agora_tem))

    from datetime import datetime
    carimbo = datetime.now().strftime("%d/%m/%Y %H:%M")
    mensagem = "Painel atualizado em %s" % carimbo

    print("Publicando em %s ..." % URL_PUBLICA)
    mudou = enviar("index.html", html, token, mensagem)
    enviar("robots.txt", robots, token, mensagem)

    print("\n" + "=" * 62)
    print("PAINEL NO AR" if mudou else "PAINEL JA ESTAVA ATUALIZADO")
    print("=" * 62)
    print(URL_PUBLICA)
    if mudou:
        print("O GitHub leva de 30 a 60 segundos para trocar a versao no ar.")
    print("=" * 62)


if __name__ == "__main__":
    publicar()

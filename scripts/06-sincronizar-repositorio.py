# -*- coding: utf-8 -*-
"""
SINCRONIZA OS SCRIPTS E O HISTORICO CONGELADO COM O GITHUB
==========================================================

O QUE ESTE SCRIPT FAZ, EM PORTUGUES SIMPLES:

As duas tarefas programadas (8:30 e 17:00) rodam na nuvem, sem o seu computador
ligado. Elas nao enxergam a pasta do projeto na sua maquina: elas BAIXAM do
GitHub tudo de que precisam. Entao o repositorio precisa estar sempre com a
versao boa dos scripts. Este script leva para la:

    scripts/03-painel-contratos-clicksign.py   -> o gerador
    scripts/05-publicar-github.py              -> o publicador
    scripts/modelo-painel.html                 -> o visual do painel
    dados/historico-congelado.json             -> o retrato de 17/08/2026

E APAGA de la o que aposentamos, para que ninguem baixe coisa velha por engano.

ORDEM QUE IMPORTA: rode ESTE script ANTES de mexer nas tarefas programadas.
Se a tarefa for alterada para procurar um arquivo que ainda nao subiu, ela
quebra na proxima rodada das 8:30.

COMO RODAR:
    python 05-scripts/06-sincronizar-repositorio.py
"""

import base64
import hashlib
import json
import os
import urllib.error
import urllib.request

PASTA_RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
USUARIO = "qpc-suprimentos"
REPOSITORIO = "controle-assinaturas"
RAMO = "main"
API = "https://api.github.com"

ARQUIVO_TOKEN = os.path.join(PASTA_RAIZ, "05-scripts", "github-token.txt")

# (caminho no repositorio, caminho aqui)
A_ENVIAR = [
    ("scripts/03-painel-contratos-clicksign.py", "05-scripts/03-painel-contratos-clicksign.py"),
    ("scripts/05-publicar-github.py",            "05-scripts/05-publicar-github.py"),
    ("scripts/06-sincronizar-repositorio.py",    "05-scripts/06-sincronizar-repositorio.py"),
    ("scripts/modelo-painel.html",               "05-scripts/modelo-painel.html"),
    ("dados/historico-congelado.json",           "02-dados-tratados/historico-congelado.json"),
]

# Ficaram para tras quando a planilha saiu de cena. Se sobrarem no repositorio,
# uma tarefa antiga pode baixar a base velha e o painel volta a mentir.
A_APAGAR = [
    "scripts/06-ler-planilha-controle.py",
    "dados/base-planilha.json",
]

# O proxy deste ambiente sequestra chamadas ao GitHub. Ver 05-publicar-github.py.
ABRIDOR = urllib.request.build_opener(urllib.request.ProxyHandler({}))


def erro(mensagem):
    raise SystemExit("\n[ERRO] " + mensagem + "\n")


def ler_token():
    token = os.environ.get("GITHUB_PAT", "").strip()
    if token:
        return token
    if not os.path.exists(ARQUIVO_TOKEN):
        erro("Nao achei a chave do GitHub em 05-scripts/github-token.txt.")
    with open(ARQUIVO_TOKEN, encoding="utf-8") as arquivo:
        return arquivo.read().strip()


def chamar(caminho, token, metodo="GET", corpo=None):
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
            return None
        erro("GitHub respondeu %s em %s\n%s"
             % (falha.code, caminho, falha.read().decode("utf-8", "replace")[:300]))


def sha_do_git(conteudo):
    return hashlib.sha1(("blob %d\0" % len(conteudo)).encode() + conteudo).hexdigest()


def sincronizar():
    token = ler_token()
    print("Sincronizando %s/%s ..." % (USUARIO, REPOSITORIO))

    for no_repo, aqui in A_ENVIAR:
        caminho = os.path.join(PASTA_RAIZ, aqui)
        if not os.path.exists(caminho):
            erro("Nao achei %s. Nao vou subir um repositorio pela metade." % aqui)
        with open(caminho, "rb") as arquivo:
            conteudo = arquivo.read()

        atual = chamar("/repos/%s/%s/contents/%s?ref=%s" % (USUARIO, REPOSITORIO, no_repo, RAMO), token)
        if atual and atual.get("sha") == sha_do_git(conteudo):
            print("  sem mudanca: %s" % no_repo)
            continue

        corpo = {
            "message": "Sincroniza %s" % no_repo,
            "content": base64.b64encode(conteudo).decode("ascii"),
            "branch": RAMO,
        }
        if atual:
            corpo["sha"] = atual["sha"]
        chamar("/repos/%s/%s/contents/%s" % (USUARIO, REPOSITORIO, no_repo), token, "PUT", corpo)
        print("  enviado:     %s (%d KB)" % (no_repo, len(conteudo) // 1024 or 1))

    for no_repo in A_APAGAR:
        atual = chamar("/repos/%s/%s/contents/%s?ref=%s" % (USUARIO, REPOSITORIO, no_repo, RAMO), token)
        if not atual:
            print("  ja nao existe: %s" % no_repo)
            continue
        chamar(
            "/repos/%s/%s/contents/%s" % (USUARIO, REPOSITORIO, no_repo), token, "DELETE",
            {"message": "Aposenta %s (a planilha nao alimenta mais o painel)" % no_repo,
             "sha": atual["sha"], "branch": RAMO},
        )
        print("  apagado:     %s" % no_repo)

    print("\nRepositorio sincronizado. So agora e seguro alterar as tarefas programadas.")


if __name__ == "__main__":
    sincronizar()

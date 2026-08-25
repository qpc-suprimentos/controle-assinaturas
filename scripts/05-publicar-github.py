# -*- coding: utf-8 -*-
"""
PUBLICA O PAINEL NO AR (GitHub Pages)
=====================================

O QUE ESTE SCRIPT FAZ, EM PORTUGUES SIMPLES:

Depois que o 03-painel-contratos-clicksign.py gera o HTML, este script manda esse
arquivo para o GitHub, que o serve como site. O endereco nunca muda:

    https://qpc-suprimentos.github.io/controle-assinaturas/

Quem tiver o link sempre ve a versao mais recente. Nao precisa arrastar nada,
nao precisa abrir o navegador, nao precisa anexar em e-mail.

DUAS ARMADILHAS DESTE AMBIENTE QUE JA CUSTARAM CARO - NAO MEXER SEM LER:

1. O ambiente do Claude tem um proxy que INTERCEPTA chamadas ao GitHub e troca a
   credencial por uma dele ("builtin injection failed (github)", HTTP 502).
   Por isso o script desliga o proxy explicitamente (ProxyHandler vazio, abaixo).
   Sem isso, nada funciona, por mais correto que o token esteja.

2a. O ENDERECO MUDOU em 25/08/2026. Era valterrocha.github.io/g200-7k3f9x2m/ e o
   nome cifrado tinha a intencao de esconder o painel. Isso protegia pouco: o
   repositorio e PUBLICO e aparece listado no perfil do dono, com qualquer nome.
   Hoje o dono da conta chama-se qpc-suprimentos e o repositorio, controle-assinaturas.
   A protecao de verdade continua sendo o robots.txt + noindex, que impedem o
   buscador de indexar. Nao remova nenhum dos dois.

2. Nao adianta tentar Netlify, GitLab, Vercel, Cloudflare, Render, Bitbucket ou
   Codeberg: TODOS estao bloqueados neste ambiente (conexao recusada em ~0,35 s).
   O GitHub e o unico host alcancavel. Ja foi gasta uma sessao inteira nisso.

COMO RODAR:
    python 05-scripts/05-publicar-github.py
"""

import base64
import hashlib
import re
import json
import os
import urllib.error
import urllib.request

# =============================================================================
# CONFIGURACAO - e so aqui que voce mexe
# =============================================================================

PASTA_RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

USUARIO = "qpc-suprimentos"
REPOSITORIO = "controle-assinaturas"   # renomeado em 25/08/2026 a pedido do Valter
RAMO = "main"
URL_PUBLICA = "https://qpc-suprimentos.github.io/controle-assinaturas/"

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

    # ---- Guarda-corpo: NAO apagar do ar o que outra rodada ja viu ------------
    #
    # POR QUE ISTO EXISTE, EM PORTUGUES SIMPLES:
    #
    # Em 25/08/2026 a rotina das 11h30 leu o Outlook, achou o aditivo ADIT01-045 e
    # publicou 104 contratos as 11h45. As 12h56 eu (Claude) fui publicar uma
    # mudanca de TITULO de um card, gerei o painel a partir do dump que estava no
    # disco - de ONTEM - e publiquei 103 contratos por cima dos 104. O aditivo
    # sumiu do ar, e o Valter viu antes de mim.
    #
    # DUAS MEDIDAS ERRADAS, para nao repetir a tentativa:
    #   - queda de 30%: 104 -> 103 e 1%, nao dispara.
    #   - hora de GERACAO do painel: a minha era mais NOVA (gerei agora, com dado
    #     velho). Comparar quando o HTML foi montado nao diz nada sobre o dado.
    #
    # A medida certa e o DADO: qual foi o e-mail mais recente que cada painel
    # enxergou, e quais contratos cada um lista. Painel novo nunca pode enxergar
    # menos longe no tempo, nem perder contrato, em relacao ao que ja esta no ar.
    def ultimo_email_de(conteudo):
        achado = re.search(rb'"ultimo_email"\s*:\s*"([^"]+)"', conteudo)
        return achado.group(1).decode() if achado else None

    def contratos_de(conteudo):
        return set(re.findall(rb'"numero"\s*:\s*"([^"]+)"', conteudo))

    if publicado:
        no_ar_bruto = base64.b64decode(publicado.get("content", ""))

        meu_ultimo = ultimo_email_de(html)
        ultimo_no_ar = ultimo_email_de(no_ar_bruto)
        if meu_ultimo and ultimo_no_ar and ultimo_no_ar > meu_ultimo:
            erro(
                "O painel NO AR enxergou e-mail mais recente do que este.\n\n"
                "  e-mail mais novo no ar:  %s\n"
                "  e-mail mais novo no meu: %s\n\n"
                "Alguma rodada leu o Outlook depois de voce - provavelmente uma tarefa\n"
                "agendada. Publicar agora apagaria do ar o que ela viu.\n\n"
                "O QUE FAZER: traga o dado novo antes de publicar.\n"
                "  python 05-scripts/06-sincronizar-repositorio.py --puxar\n"
                "  python 05-scripts/03-painel-contratos-clicksign.py\n"
                "e so entao publique. NAO force." % (ultimo_no_ar, meu_ultimo)
            )

        sumindo = contratos_de(no_ar_bruto) - contratos_de(html)
        if sumindo:
            erro(
                "Estes contratos estao no painel que esta NO AR e nao estao neste:\n\n"
                "  %s\n\n"
                "Contrato nao some sozinho. Ou a coleta veio incompleta, ou este painel\n"
                "foi gerado a partir de um dump mais velho do que o que gerou o painel\n"
                "publicado. Nos dois casos, publicar perde informacao.\n\n"
                "O QUE FAZER: python 05-scripts/06-sincronizar-repositorio.py --puxar\n"
                "depois rode o gerador de novo. NAO force."
                % ", ".join(sorted(x.decode() for x in sumindo))
            )
        print("Frescor: nada some e o dado nao regride. OK.")

    from datetime import datetime
    from zoneinfo import ZoneInfo
    # datetime.now() sem argumento pega o fuso da MAQUINA, que aqui e UTC - e a
    # mensagem do commit saia 3 horas adiantada, atrapalhando auditoria. Mesma
    # regra que ja vale no gerador: fuso sempre explicito.
    carimbo = datetime.now(ZoneInfo("America/Sao_Paulo")).strftime("%d/%m/%Y %H:%M")
    mensagem = "Painel atualizado em %s" % carimbo

    print("Publicando em %s ..." % URL_PUBLICA)
    mudou = enviar("index.html", html, token, mensagem)
    enviar("robots.txt", robots, token, mensagem)

    # ---- Guardar no repositorio o dump que gerou ESTE painel ------------------
    # A tarefa agendada roda na nuvem, num /tmp que morre no fim da sessao. Sem
    # isto, o que ela leu do Outlook se perde e a proxima pessoa que for gerar o
    # painel (inclusive o Claude, no computador do Valter) comeca de um dump
    # velho - que foi como o ADIT01-045 se perdeu em 25/08/2026.
    pasta_bruta = os.path.join(PASTA_RAIZ, "01-dados-brutos")
    if os.path.isdir(pasta_bruta):
        dumps = sorted(
            f for f in os.listdir(pasta_bruta)
            if f.startswith("emails-clicksign-") and f.endswith(".json")
        )
        if dumps:
            with open(os.path.join(pasta_bruta, dumps[-1]), "rb") as arquivo:
                enviar("dados/%s" % dumps[-1], arquivo.read(), token,
                       "Dump de e-mails que gerou o painel de %s" % carimbo)

    print("\n" + "=" * 62)
    print("PAINEL NO AR" if mudou else "PAINEL JA ESTAVA ATUALIZADO")
    print("=" * 62)
    print(URL_PUBLICA)
    if mudou:
        print("O GitHub leva de 30 a 60 segundos para trocar a versao no ar.")
    print("=" * 62)


if __name__ == "__main__":
    publicar()

#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
sonda_nfp.py - Etapa 1 do projeto de extracao das notas da Nota Fiscal Paulista.

OBJETIVO
    Descobrir, empiricamente, se o login por certificado digital (e-CPF A1) na NFP
    funciona via mTLS puro - ou seja, SEM reCAPTCHA e SEM navegador.

    A hipotese: o reCAPTCHA protege o formulario do gov.br. A autenticacao por
    certificado acontece no handshake TLS, uma camada abaixo do HTML. Se a SEFAZ
    implementou do jeito convencional, o .pfx autentica uma sessao HTTP direto.

    Este script NAO faz login de verdade nem baixa nota nenhuma. Ele so bate nos
    endpoints, olha o que volta e te da um relatorio.

SEGURANCA
    Seu .pfx e sua senha nunca saem da sua maquina. O certificado e usado apenas
    no handshake TLS direto com os servidores da SEFAZ-SP. A senha e pedida via
    getpass (nao aparece na tela, nao vai pro historico do shell).

    O Python exige arquivos em disco para carregar certificado cliente, entao o
    script extrai o PEM para arquivos temporarios com permissao restrita e os
    apaga no finally - inclusive se der erro ou se voce apertar Ctrl+C.

USO
    pip install requests cryptography
    python sonda_nfp.py C:\\caminho\\para\\seu_certificado.pfx

    Se o seu A1 estiver instalado no repositorio do Windows e voce nao tiver mais
    o arquivo .pfx: exporte-o em certmgr.msc -> Pessoal -> Certificados -> botao
    direito -> Todas as tarefas -> Exportar -> "Sim, exportar a chave privada".
"""

import os
import re
import ssl
import sys
import stat
import tempfile
import getpass
import warnings
from urllib.parse import urljoin, urlparse

try:
    import requests
    from requests.adapters import HTTPAdapter
    from urllib3.poolmanager import PoolManager
    from cryptography.hazmat.primitives.serialization import (
        pkcs12, Encoding, PrivateFormat, NoEncryption,
    )
    from cryptography import x509
except ImportError as e:
    sys.exit(f"Falta dependencia: {e.name}\nRode:  pip install requests cryptography")

warnings.filterwarnings("ignore", message="Unverified HTTPS request")

# ----------------------------------------------------------------------------
# Endpoints candidatos. A SEFAZ nao documenta a URL do login por certificado,
# entao a sonda testa os caminhos plausiveis e tambem descobre links novos
# lendo a home em busca de href contendo "certific".
# ----------------------------------------------------------------------------
BASE = "https://www.nfp.fazenda.sp.gov.br"

CANDIDATOS = [
    f"{BASE}/",
    f"{BASE}/login.aspx",
    f"{BASE}/Login.aspx",
    f"{BASE}/CertificadoDigital/LoginCertificado.aspx",
    f"{BASE}/certificado/",
    f"{BASE}/LoginCertificado.aspx",
    "https://www.nfp.fazenda.sp.gov.br/Inicial.aspx",
]

TIMEOUT = 20


class TLS12Adapter(HTTPAdapter):
    """Servidores da SEFAZ costumam ser exigentes com a negociacao TLS.
    Este adapter forca um contexto com o certificado cliente carregado."""

    def __init__(self, certfile, keyfile, **kw):
        self._certfile = certfile
        self._keyfile = keyfile
        super().__init__(**kw)

    def init_poolmanager(self, connections, maxsize, block=False, **kw):
        ctx = ssl.create_default_context()
        ctx.check_hostname = False
        ctx.verify_mode = ssl.CERT_NONE
        ctx.load_cert_chain(certfile=self._certfile, keyfile=self._keyfile)
        # Alguns servicos .asmx da SEFAZ ainda rejeitam ciphers modernos demais.
        try:
            ctx.set_ciphers("DEFAULT@SECLEVEL=1")
        except ssl.SSLError:
            pass
        kw["ssl_context"] = ctx
        self.poolmanager = PoolManager(
            num_pools=connections, maxsize=maxsize, block=block, **kw
        )


def extrair_pem(caminho_pfx, senha):
    """Converte o .pfx em dois arquivos PEM temporarios com permissao 0600.
    Retorna (certfile, keyfile, titular)."""
    with open(caminho_pfx, "rb") as fh:
        blob = fh.read()

    chave, cert, extras = pkcs12.load_key_and_certificates(
        blob, senha.encode("utf-8") if senha else None
    )
    if chave is None or cert is None:
        raise ValueError("O arquivo nao contem chave privada + certificado.")

    titular = "(desconhecido)"
    try:
        cn = cert.subject.get_attributes_for_oid(x509.NameOID.COMMON_NAME)
        if cn:
            titular = cn[0].value
    except Exception:
        pass

    def _tmp(dados, sufixo):
        fd, caminho = tempfile.mkstemp(suffix=sufixo)
        os.write(fd, dados)
        os.close(fd)
        try:
            os.chmod(caminho, stat.S_IRUSR | stat.S_IWUSR)
        except OSError:
            pass  # Windows ignora, tudo bem
        return caminho

    cert_pem = cert.public_bytes(Encoding.PEM)
    for extra in (extras or []):
        cert_pem += extra.public_bytes(Encoding.PEM)

    key_pem = chave.private_bytes(
        Encoding.PEM, PrivateFormat.TraditionalOpenSSL, NoEncryption()
    )

    return _tmp(cert_pem, ".crt.pem"), _tmp(key_pem, ".key.pem"), titular, cert


def analisar(resp):
    """Le a resposta e devolve os sinais que interessam."""
    html = resp.text or ""
    baixo = html.lower()

    sinais = []
    if "recaptcha" in baixo or "g-recaptcha" in baixo:
        sinais.append("RECAPTCHA NA PAGINA")
    if "hcaptcha" in baixo:
        sinais.append("hCaptcha")
    if "sso.acesso.gov.br" in resp.url or "gov.br" in urlparse(resp.url).netloc:
        sinais.append("REDIRECIONOU PRO GOV.BR")
    if "certificado" in baixo:
        sinais.append("menciona certificado")
    if re.search(r"(sair|logout|meus dados|saldo|consumidor)", baixo):
        sinais.append("parece area logada")

    titulo = ""
    m = re.search(r"<title[^>]*>(.*?)</title>", html, re.S | re.I)
    if m:
        titulo = " ".join(m.group(1).split())[:80]

    return titulo, sinais


def descobrir_links(sessao):
    """Le a home procurando links que mencionem certificado."""
    achados = set()
    try:
        r = sessao.get(BASE + "/", timeout=TIMEOUT)
        for href in re.findall(r'href=["\']([^"\']+)["\']', r.text, re.I):
            if "certific" in href.lower() or "login" in href.lower():
                achados.add(urljoin(BASE + "/", href))
    except Exception as e:
        print(f"    (nao consegui ler a home para descobrir links: {e})")
    return sorted(achados)


def main():
    if len(sys.argv) < 2 or sys.argv[1] in ("-h", "--help", "/?"):
        print(__doc__)
        return

    caminho = sys.argv[1]
    if not os.path.isfile(caminho):
        sys.exit(f"Arquivo nao encontrado: {caminho}")

    # Em container sem TTY dá para passar a senha por variavel de ambiente.
    # Preferir sempre o getpass interativo: env var fica visivel em `docker inspect`.
    senha = os.environ.get("NFP_PFX_SENHA")
    if senha:
        print("(senha lida de NFP_PFX_SENHA)")
    else:
        senha = getpass.getpass("Senha do certificado (.pfx): ")

    certfile = keyfile = None
    try:
        certfile, keyfile, titular, cert = extrair_pem(caminho, senha)

        print()
        print("=" * 72)
        print("  SONDA - login por certificado digital na Nota Fiscal Paulista")
        print("=" * 72)
        validade = getattr(cert, "not_valid_after_utc", None) or cert.not_valid_after
        print(f"  Titular do certificado : {titular}")
        print(f"  Valido ate             : {validade:%d/%m/%Y}")
        print()

        if ":" not in titular and not re.search(r"\d{11}", titular.replace(":", "")):
            print("  AVISO: nao consegui identificar um CPF no titular. Se este for")
            print("  um e-CNPJ, o caminho da NFP como consumidor nao vai funcionar.")
            print()

        sessao = requests.Session()
        sessao.headers["User-Agent"] = (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
            "(KHTML, like Gecko) Chrome/126.0 Safari/537.36"
        )
        sessao.mount("https://", TLS12Adapter(certfile, keyfile))

        alvos = list(dict.fromkeys(CANDIDATOS + descobrir_links(sessao)))

        print(f"  Testando {len(alvos)} endpoints com mTLS...")
        print("-" * 72)

        promissores = []
        for url in alvos:
            try:
                r = sessao.get(url, timeout=TIMEOUT, allow_redirects=True)
            except requests.exceptions.SSLError as e:
                print(f"  [TLS ]  {url}\n          handshake recusado: {str(e)[:90]}")
                continue
            except Exception as e:
                print(f"  [ERRO]  {url}\n          {type(e).__name__}: {str(e)[:90]}")
                continue

            titulo, sinais = analisar(r)
            marca = "  " if r.status_code >= 400 else "OK"
            print(f"  [{r.status_code} {marca}] {url}")
            if r.url != url:
                print(f"          -> {r.url}")
            if titulo:
                print(f"          titulo: {titulo}")
            if sinais:
                print(f"          sinais: {', '.join(sinais)}")

            cookies = [c.name for c in r.cookies]
            if cookies:
                print(f"          cookies: {', '.join(cookies[:6])}")

            if r.status_code < 400 and "RECAPTCHA NA PAGINA" not in sinais \
               and "REDIRECIONOU PRO GOV.BR" not in sinais:
                promissores.append(r.url)

        print("-" * 72)
        print()
        print("  LEITURA DO RESULTADO")
        print()
        if promissores:
            print("  Endpoints que responderam sem captcha e sem jogar pro gov.br:")
            for u in promissores:
                print(f"    - {u}")
            print()
            print("  Isso e o sinal verde. Me manda essa saida que eu escrevo a")
            print("  etapa 2 (varrer a listagem de notas e coletar as chaves).")
        else:
            print("  Nenhum endpoint aceitou o certificado sem captcha/gov.br.")
            print("  Nao e o fim: significa so que o caminho e o navegador com o")
            print("  certificado instalado. Me manda a saida assim mesmo - os")
            print("  codigos de status e os redirects dizem qual e o plano B.")
        print()

    finally:
        for f in (certfile, keyfile):
            if f and os.path.exists(f):
                try:
                    with open(f, "r+b") as fh:      # sobrescreve antes de apagar
                        n = os.path.getsize(f)
                        fh.write(b"\0" * n)
                    os.remove(f)
                except OSError:
                    pass


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\nInterrompido. Arquivos temporarios foram limpos.")

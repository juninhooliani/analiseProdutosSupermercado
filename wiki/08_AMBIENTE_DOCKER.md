# 08. Ambiente de Execução (Docker)

## 1. Por que container

A coleta precisa de dependências específicas (`cryptography` para ler o `.pfx`,
e mais adiante Playwright, se a rota exigir navegador). Instalar isso direto no
Windows polui o ambiente e torna o resultado irreprodutível.

Além disso, o `.venv` do projeto e o Python do sistema divergem com facilidade —
o container elimina a variável "funcionou na minha máquina".

## 2. Estrutura

```text
/
├── docker/
│   ├── Dockerfile                  # imagem de coleta (python:3.12-slim)
│   ├── requirements.txt            # dependencias pinadas
│   └── exportar_certificado.ps1    # extrai o e-CPF do Windows para .pfx
├── docker-compose.yml              # servico `sonda`
├── .dockerignore                   # impede o .pfx de entrar no build
└── certs/                          # .pfx fica aqui (ignorado pelo git)
```

## 3. Modelo de segurança do certificado

Três camadas, e vale entender o porquê de cada uma:

| Camada | Mecanismo | Protege de |
|---|---|---|
| `.gitignore` | `certs/`, `*.pfx`, `*.p12`, `*.pem`, `*.key` | commit acidental do certificado |
| `.dockerignore` | mesmas entradas | o `.pfx` ser copiado **para dentro da imagem** — imagens são camadas imutáveis e vazam se publicadas |
| volume `:ro` | `./certs:/certs:ro` | escrita/alteração do certificado pelo container |

A senha é pedida via `getpass` (não ecoa, não entra no histórico). Existe um
fallback por variável de ambiente `NFP_PFX_SENHA` para execução sem TTY, mas
**prefira o modo interativo**: variável de ambiente aparece em `docker inspect`.

O container roda como usuário não-root (`coletor`, UID 1000).

## 4. Uso

### 4.1. Obter o `.pfx`

O container Linux não enxerga o repositório de certificados do Windows. Se o
seu A1 está instalado no Windows e você não tem mais o arquivo:

```powershell
powershell -ExecutionPolicy Bypass -File docker\exportar_certificado.ps1
```

O script lista os certificados pessoais válidos, exporta o escolhido para
`certs\certificado.pfx` e pede uma senha nova para o arquivo.

> Só funciona se a chave privada tiver sido marcada como exportável na
> instalação. Se não foi, reinstale o A1 a partir do arquivo original marcando
> *"Marcar esta chave como exportável"*.

Alternativa manual: `certmgr.msc` -> Pessoal -> Certificados -> botão direito ->
Todas as tarefas -> Exportar -> *"Sim, exportar a chave privada"*.

### 4.2. Construir e rodar a sonda

```bash
docker compose build
docker compose run --rm sonda /certs/certificado.pfx
```

O `run` (e não `up`) é intencional: a sonda é um processo de vida curta que
precisa de terminal interativo para o `getpass`. Por isso o compose declara
`stdin_open` e `tty`.

### 4.3. Rede

O container herda a rede do host. É justamente o que o ambiente de
desenvolvimento remoto **não** tem — daí a necessidade de rodar a coleta aqui.

Para conferir que o container alcança a SEFAZ:

```bash
docker compose run --rm --entrypoint python sonda -c \
  "import socket;print(socket.gethostbyname('www.nfp.fazenda.sp.gov.br'))"
```

## 5. Evolução prevista

Se a sonda indicar que a rota exige navegador (redirect para o SSO do gov.br),
a imagem passa a ser `mcr.microsoft.com/playwright/python`, e o certificado
precisará ser importado no perfil do Chromium dentro do container via política
`AutoSelectCertificateForUrls`. Isso está fora do escopo da etapa 1.

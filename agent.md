# Estado Atual - Agente

> Este arquivo é o ponto de entrada de contexto para agentes de IA.
> A base de conhecimento completa está em `wiki/`. Leia `wiki/06_AI_AGENT_RULES.md`
> antes de escrever código.

## 1. Objetivos do Usuário
1. **Migração de stack**: transformar a análise de cupons fiscais em um
   aplicativo web/mobile moderno.
2. **Coleta automatizada de NFC-e**: obter o histórico de notas fiscais do
   usuário para calcular padrão de compra, gasto mensal e qual estabelecimento
   tem o menor preço por produto.

## 2. Decisão de arquitetura da coleta

A rota de coleta é a **sessão autenticada por e-CPF (certificado A1)** no
portal da Nota Fiscal Paulista. Detalhes em `wiki/07_COLETA_ECPF.md`.

O que foi verificado e descartado:

- `NFeDistribuicaoDFe` não cobre NFC-e (modelo 65), só modelo 55.
- `SAE-NFC-e` exige e-CNPJ e só devolve notas do próprio emitente.
- O portal da NFP não tem exportação de XML nem CSV.

> **Quebra de captcha por OCR está fora de escopo.** O robô anterior
> (`extrator_sefaz.py`, com `ddddocr`) foi **removido**. A justificativa está em
> `wiki/06_AI_AGENT_RULES.md`, seção 3 — não reabra esse caminho.

## 3. Estado do código

| Item | Estado |
|---|---|
| `sonda_nfp.py` | Pronto. Testa se o login mTLS na NFP dispensa captcha. **Ainda não executado.** |
| `docker/` + `docker-compose.yml` | Prontos. Ambiente de execução da coleta. Ver `wiki/08_AMBIENTE_DOCKER.md`. |
| `Supermercado.py`, `LErCUpomTexto.py`, `categorias.py` | Legado funcional: processa cupons TXT e gera `compras.xlsx`. |
| `web_app/` | Next.js + React + TailwindCSS inicializado. Sem dashboard ainda. |
| `banco_compras.db` | Não criado. |
| `XMLs/` | Vazia. Destino dos XMLs coletados. |

## 4. Como executar

**Coleta (etapa 1 - sonda):**
```bash
docker compose build
docker compose run --rm sonda /certs/certificado.pfx
```
Antes, obtenha o `.pfx`: `powershell -ExecutionPolicy Bypass -File docker\exportar_certificado.ps1`

**Legado (processamento de TXT):**
```bash
python Supermercado.py
```

**Interface:**
```bash
cd web_app && npm run dev   # http://localhost:3000
```

## 5. Próximos Passos

> **Há tarefas pendentes de execução.** Ver `HANDOFF_AGENTE_LOCAL.md` — contém
> os dois comandos que precisam rodar em terminal com rede (publicar o histórico
> limpo e executar a sonda de coleta), com as verificações e os limites de cada um.

1. **Sonda NFP Executada (`sonda_nfp.py`)**: Concluída. Resultado: a NFP redireciona requisições mTLS puras para a tela de login (`login.aspx` / gov.br). O plano B suportado é a automação via **Playwright utilizando o certificado e-CPF instalado no navegador (com a política `AutoSelectCertificateForUrls`)**.
2. Criar `banco_compras.db` (SQLite) conforme `wiki/02_DATA_MODEL.md`, usando **GTIN como chave de normalização de produto**.
3. Construir o parser de XMLs e a ingestão no banco de dados.
4. Construir o dashboard do `web_app` sobre esse banco.

## 6. Higiene de dados no repositório

O repositório é **público**. O histórico local foi reescrito para remover
dados sensíveis e está pronto para ser publicado com `--force-with-lease`.

O que foi removido do histórico versionado:

| Item | Situação anterior | Agora |
|---|---|---|
| `.env` (`OPENAI_API_KEY`) | commitado localmente, **nunca publicado** | fora do histórico e no `.gitignore` |
| `cupons_txt/*.txt` (contêm CPF) | **publicados** no GitHub desde "Add files via upload" | purgados de todo o histórico do `main` |
| `compras.xlsx` (histórico de compras) | commitado localmente | ignorado |
| CPF hardcoded em `LErCUpomTexto.py` | commitado localmente | mascarado como `000.000.000-00` |

Verificação: varredura de todos os blobs alcançáveis pelo `main` não encontra
nenhum CPF real nem chave de API.

### Regras permanentes

- **Nunca** versionar `cupons_txt/`, `XMLs/*.xml`, `compras.xlsx`, `certs/` ou `.env`.
  As pastas ficam no repositório vazias, com `.gitkeep`.
- Cupons de exemplo dentro de código devem usar CPF fictício (`000.000.000-00`)
  e razão social genérica.

### Refs locais que NÃO devem ser publicados

`refs/original/refs/heads/main` e a tag `backup-antes-limpeza` ainda contêm os
objetos antigos, de propósito, como rede de segurança. Um `git push` normal não
os leva — mas **nunca** use `git push --tags` nem `git push --mirror` aqui.
Para descartá-los depois de confirmar que está tudo certo:

```bash
git update-ref -d refs/original/refs/heads/main
git tag -d backup-antes-limpeza
git reflog expire --expire=now --all && git gc --prune=now --aggressive
```

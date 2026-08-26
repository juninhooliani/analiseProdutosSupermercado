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
1. **Executar `sonda_nfp.py`** e interpretar a saída conforme a tabela da seção 4
   de `wiki/07_COLETA_ECPF.md`. Tudo depende desse resultado.
2. Implementar a listagem de chaves de acesso por período.
3. Implementar o download do XML por chave.
4. Criar `banco_compras.db` (SQLite) conforme `wiki/02_DATA_MODEL.md`,
   usando **GTIN como chave de normalização de produto**.
5. Construir o dashboard do `web_app` sobre esse banco.

## 6. Pendência de segurança

O `.env` (com `OPENAI_API_KEY`) está no commit `87dd921`, que ainda **não foi
enviado** ao GitHub — e o repositório é público. Já foram criados o `.gitignore`
e feito `git rm --cached .env`, mas **o blob continua no histórico local**.
Não faça `git push` antes de limpar o histórico ou rotacionar a chave.

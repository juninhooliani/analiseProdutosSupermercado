# 07. Coleta via e-CPF (certificado digital A1)

> Complementa o `03_INGESTION_STRATEGY.md`. Documenta uma rota de ingestão
> que não estava mapeada e que substitui a tentativa de quebra de captcha.

## 1. Correção de premissa

O `03_INGESTION_STRATEGY.md` afirma que a raspagem é *"inviável sem serviços
pagos de quebra de captcha"*. Isso é verdade para o acesso **anônimo**, mas não
para o acesso **autenticado**: o sistema da NFP aceita login por certificado
digital `e-CPF ou e-CNPJ, tipo A1 ou A3` no perfil Consumidor.

Um handshake TLS com certificado de cliente acontece **abaixo da camada HTTP**.
Não há captcha em um handshake — o reCAPTCHA protege o formulário do gov.br,
não a autenticação mútua. Se a SEFAZ implementou o login por certificado da
forma convencional, o `.pfx` autentica um `requests.Session()` direto, sem
navegador e sem OCR.

Isso ainda **não está confirmado**. É exatamente o que a sonda testa.

## 2. O que foi verificado sobre as rotas de coleta

| Rota | Situação | Fonte |
|---|---|---|
| `NFeDistribuicaoDFe` (lote com certificado) | **Não cobre NFC-e modelo 65.** Só modelo 55. | Documentação SEFAZ |
| `SAE-NFC-e` (lote, até 2.000 chaves, janela de 100 dias) | **Fechado para pessoa física.** Nota Técnica: *"o contribuinte a ser pesquisado será o que possui o CNPJ que consta no certificado digital (e-CNPJ)"*. Consultas por CPF não são permitidas. | Nota Técnica SAE-NFC-e v1.0.0 |
| Exportação no portal da NFP | **Não existe.** Nem XML, nem CSV. Só tela HTML. | Guia de Consulta de Documentos Fiscais |
| Login por e-CPF na NFP | **Aceito.** A confirmar se dispensa captcha. | Guia de Certificado Digital NFP |
| XML por chave como consumidor identificado | **Permitido**, desde que o CPF tenha sido informado na compra. | SEFAZ |

Conclusão: **não existe download em massa oficial de XML para pessoa física.**
A melhor rota disponível é a sessão autenticada por e-CPF no portal da NFP.

## 3. Por que essa rota é superior à quebra de captcha

Além de não violar os termos do portal, ela é tecnicamente melhor:

- **Traz o GTIN.** O XML tem o código de barras de cada item; a tela HTML não.
  Isso resolve de graça o passo 3 do "Pipeline de Normalização de Produtos" —
  `LEITE UHT INT ELEGANCE 1L` e `LEITE INTEGRAL ELEGANCE 1000ML` viram o mesmo
  GTIN, sem depender de regex nem de classificação por IA.
- **Não quebra.** OCR de captcha falha a cada mudança de fonte ou de ruído da
  imagem. Um certificado vale por 1 ano.
- **É a sua própria conta.** A consulta pública por chave é anônima e o captcha
  ali existe justamente como controle anti-automação; contorná-lo com OCR abre
  a porta para varrer notas de terceiros, já que a chave de acesso é
  parcialmente previsível (CNPJ + série + número sequencial).

## 4. Estado atual: etapa 1 (sonda)

O arquivo `sonda_nfp.py` na raiz **não faz login nem baixa nota**. Ele testa os
endpoints da NFP com mTLS e reporta o que volta.

### Uso (recomendado: container)

Ver `08_AMBIENTE_DOCKER.md`. O ambiente de desenvolvimento remoto **não tem
rede** para os domínios da SEFAZ; a coleta roda no host, via Docker.

```bash
# 1. exportar o e-CPF do Windows para certs\certificado.pfx
powershell -ExecutionPolicy Bypass -File docker\exportar_certificado.ps1

# 2. construir e rodar
docker compose build
docker compose run --rm sonda /certs/certificado.pfx
```

### Uso (alternativo: Python direto no host)

```
pip install requests cryptography
python sonda_nfp.py "D:\caminho\para\certificado.pfx"
```

A senha é pedida via `getpass` (não ecoa, não entra no histórico do shell). O
certificado é usado apenas no handshake com a SEFAZ. O Python exige o
certificado em disco, então o script extrai o PEM para arquivos temporários com
permissão restrita, **sobrescreve o conteúdo e apaga** no `finally` — inclusive
em caso de erro ou Ctrl+C.

Se o A1 estiver instalado no repositório do Windows e o `.pfx` não existir mais:
`certmgr.msc` -> Pessoal -> Certificados -> botão direito -> Todas as tarefas ->
Exportar -> *"Sim, exportar a chave privada"*.

### Como ler o resultado

| Saída | Significado | Próximo passo |
|---|---|---|
| `200` sem reCAPTCHA e sem redirect pro gov.br | mTLS funciona | Etapa 2 vira cliente HTTP puro, sem navegador |
| `403` ou handshake TLS recusado | Endpoint existe, quer o certificado por outro caminho | Ajustar o adapter conforme o erro |
| Redirect para `sso.acesso.gov.br` | Login por certificado foi absorvido pelo SSO federal | Playwright com o certificado instalado no Windows |

## 5. Roteiro

1. **Sonda mTLS** — `sonda_nfp.py`. *(feito, aguardando execução)*
2. **Listagem de chaves** — varrer a consulta de documentos fiscais por período
   e coletar as chaves de 44 dígitos.
3. **Download do XML por chave** — com o mesmo e-CPF.
4. **Parser + carga** — alimentar `TransactionItem` do `02_DATA_MODEL.md`,
   usando o GTIN como chave de normalização.

## 6. Observação sobre a janela de 24 horas

Se a rota acabar exigindo navegador, vale lembrar: o código de verificação que
a NFP manda por e-mail libera os detalhes por **24 horas** e vale por 1 hora.
É **um código por janela, não por nota** — ou seja, uma autenticação humana por
dia é suficiente para varrer meses de histórico.

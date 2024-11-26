# analiseProdutosSupermercado
Processa todos cupons fiscais em formato txt em uma pasta e gera uma planilha em excel com as colunas: Número do Cupom	Data	Hora	Dia da Semana	Supermercado	Número	Código de Barras	Descrição	Quantidade	Unidade	Preço Unitário	Total	Categoria
. A categoria pode ser de forma automática através da API do chatGPT (paga) ou de um arquivo categorias.py que está inserido no projeto.

Crie uma pasta cupons_txt, e jogue os cupons fiscais em formato txt nesse diretório

Em breve será adicionado o processamento de XML.

Para conexão com a API do Cht, é necessário criar um arquivo .env com OPENAI_API_KEY=sua_chave_aqui onde é chamado na linha:

openai.api_key = os.getenv("OPENAI_API_KEY")



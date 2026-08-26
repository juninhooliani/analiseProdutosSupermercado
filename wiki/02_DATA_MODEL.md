# 02. Modelo de Dados Unificado (Data Model)

## 1. Schema da Transação Financeira (`Transaction`)
Representa qualquer fluxo monetário no sistema (entradas, saídas, transferências, faturas).

```json
{
  "id": "uuid-v4",
  "date": "YYYY-MM-DDTHH:MM:SS",
  "type": "EXPENSE | INCOME | TRANSFER",
  "amount": 250.75,
  "category_id": "cat_alimentacao",
  "subcategory_id": "sub_supermercado",
  "account_id": "acc_nubank_credito",
  "description": "Compras da Semana - Carrefour",
  "status": "COMPLETED | PENDING",
  "source": "NF_SEFAZ | MANUAL | BANK_OFX | CSV",
  "has_items": true,
  "metadata": {
    "store_name": "Carrefour Comércio e Indústria LTDA",
    "cnpj": "45.543.915/0001-81",
    "invoice_key": "3526..."
  }
}
```

## 2. Schema de Itens Detalhados (`TransactionItem`)
Utilizado quando a transação possui detalhamento de produtos (ex: Notas Fiscais de mercado ou farmácia).

```json
{
  "id": "uuid-v4",
  "transaction_id": "uuid-v4",
  "raw_name": "ARROZ BRANCO T1 5KG TIO JOAO",
  "normalized_name": "Arroz Branco 5kg",
  "category_id": "cat_mercearia",
  "quantity": 1.0,
  "unit_price": 28.90,
  "total_price": 28.90,
  "unit_of_measure": "UN",
  "discount": 0.00
}
```

## 3. Schemas de Apoio (`Account`, `Category`, `Budget`)

### Conta / Meio de Pagamento (`Account`)
```json
{
  "id": "acc_nubank_credito",
  "name": "Cartão Nubank",
  "type": "CREDIT_CARD | CHECKING_ACCOUNT | CASH | INVESTMENT",
  "currency": "BRL"
}
```

### Categorias (`Category`)
Estrutura hierárquica (Categoria Pai -> Subcategoria):
- **Alimentação** (Supermercado, Feira, Restaurante, Delivery)
- **Moradia** (Aluguel, Condomínio, Energia, Água, Internet)
- **Transporte** (Combustível, Uber, Manutenção, IPVA)
- **Saúde** (Farmácia, Consultas, Plano de Saúde)
- **Lazer** (Assinaturas, Viagens, Hobbies)
- **Receitas** (Salário, Freelance, Rendimentos)

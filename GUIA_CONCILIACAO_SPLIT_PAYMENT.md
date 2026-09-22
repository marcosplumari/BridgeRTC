# Guia de Conciliação Financeira: Split Payment e Caixa do Lojista

## 1. O Desafio Contábil e Financeiro do Split Payment (LC 214/2025)

Com a Reforma Tributária e a entrada em vigor do **Split Payment**, a sistemática de fluxo de caixa das empresas muda radicalmente:

### Antes da Reforma:
- Cliente pagava **R$ 100,00** no caixa (Pix ou Cartão).
- O lojista recebia **R$ 100,00** na conta (ou ~R$ 99,00 após tarifas de cartão).
- Os impostos (ICMS, PIS, COFINS, ISS) eram recolhidos apenas no mês seguinte via guia de arrecadação (DAS / DARE / DARF).

### Com o Split Payment:
- O consumidor paga **R$ 100,00**.
- A instituição financeira / PSP (banco ou credenciadora), de forma automática e instantânea no momento da liquidação:
  1. Segrega os tributos devidos ao Comitê Gestor do IBS e à Receita Federal (CBS). Exemplo alíquota padrão estimada: ~**26,5%** (R$ 26,50).
  2. Desconta a tarifa operacional do banco/adquirente (ex: R$ 0,99).
  3. Credita na conta corrente do lojista **apenas o valor líquido** (ex: **R$ 72,51**).

Se o ERP der baixa no Contas a Receber esperando encontrar R$ 100,00 no extrato bancário, haverá **furo no caixa e na conciliação**.

---

## 2. Como a Informação Retorna da API do Banco

Na especificação do Bacen e da SEFAZ para liquidação com Split Payment, o endpoint `GET /v2/cob/{txid}` e os webhooks de confirmação retornam o objeto `componentesValor`:

```json
{
  "txid": "PDV202609221200001",
  "status": "CONCLUIDA",
  "pago": true,
  "valorBruto": 100.00,
  "valorTributosRetidos": 26.50,
  "valorTarifaBancaria": 0.99,
  "valorLiquidoRecebido": 72.51,
  "detalheSplit": {
    "cbsRetido": 8.80,
    "ibsRetido": 17.70,
    "totalImpostosGoverno": 26.50,
    "tarifaPSP": 0.99,
    "liquidoDisponivelCaixa": 72.51
  },
  "endToEndId": "E1234567820260922120000000001"
}
```

---

## 3. O que o Software em VB6 Deve Fazer no Contas a Receber

O formulário `frmPixSupermercado` na DLL `BridgeRTC` disponibiliza as seguintes propriedades públicas após a confirmação:
- `ValorBruto` (ex: 100.00)
- `ValorTributosRetidos` (ex: 26.50)
- `ValorTarifa` (ex: 0.99)
- `ValorLiquido` (ex: 72.51)
- `ValorCbs` (ex: 8.80)
- `ValorIbs` (ex: 17.70)
- `EndToEndId` (Identificador único Bacen)

### Modelo de Lançamento Contábil / Financeiro no ERP:
Ao salvar o recebimento da venda no banco de dados do ERP:
1. **Baixa da Venda (Contas a Receber)**: Liquida o título original pelo valor cheio de **R$ 100,00** (para não ficar pendência do cliente).
2. **Entrada na Conta Corrente / Caixa**: Lança débito bancário real de **R$ 72,51** (bate 100% com o extrato bancário do dia).
3. **Lançamento de Tributos Antecipados (Split Payment)**: Lança despesa/dedução de receita de **R$ 26,50** sob a conta "Impostos Retidos na Fonte - CBS/IBS Split".
4. **Lançamento de Tarifa Transacional**: Lança despesa bancária de **R$ 0,99** sob a conta "Tarifas de Cobrança Pix".

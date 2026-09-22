# BridgeRTC

Biblioteca em C# (.NET Framework 4.8) com interoperabilidade COM para **Visual Basic 6.0 (VB6)**, voltada ao atendimento da **Reforma Tributária (Lei Complementar nº 214/2025 - Split Payment)** e integração nativa com a **API Pix do Banco Central do Brasil (Bacen v2)** diretamente com instituições financeiras.

---

## Novos Recursos Implementados

1. **Estorno e Devolução de Pix no Caixa (`DevolverPix` / `ConsultarDevolucaoPix`)**:
   - Atende à especificação oficial do Bacen (`PUT /v2/pix/{e2eid}/devolucao/{idDevolucao}`).
   - Permite que o operador de caixa estorne a venda imediatamente caso haja erro de impressão ou cancelamento de compra, devolvendo o dinheiro à conta do cliente em segundos.

2. **Geração de XML de Pagamento com `<tpIntegra>1</tpIntegra>`**:
   - Atende às exigências das SEFAZ estaduais para pagamentos integrados no PDV.
   - Gera as tags `<pag>`, `<card>`, `<tpIntegra>1</tpIntegra>` e `<infTransacPag>` com o `idTransacao` (`endToEndId`).

3. **Frente de Caixa (PDV Supermercado) Resiliente em VB6**:
   - Temporizador de timeout regressivo (3 minutos de espera máxima configurável).
   - Botão de **Reconsulta Forçada ("Reconsultar Banco Agora")** para contornar instabilidades temporárias de internet no caixa sem perder vendas.
   - Painel dinâmico de exibição da segregação do Split Payment (Bruto, CBS/IBS retidos e Líquido em conta).

---

## Documentação Técnica

Consulte os arquivos na raiz deste repositório:
- **`DOCUMENTACAO_PIX_BRIDGERTC.md`**: Manual de integração, credenciamento em bancos e procedimentos de homologação/produção.
- **`GUIA_CONCILIACAO_SPLIT_PAYMENT.md`**: Manual detalhado de conciliação financeira para contas a receber e fluxo de caixa.

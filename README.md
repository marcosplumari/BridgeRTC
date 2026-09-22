# BridgeRTC

Biblioteca em C# (.NET Framework 4.8) com interoperabilidade COM para **Visual Basic 6.0 (VB6)**, voltada ao atendimento da **Reforma Tributária (Lei Complementar nº 214/2025 - Split Payment)** e integração nativa com a **API Pix do Banco Central do Brasil (Bacen v2)** diretamente com instituições financeiras, sem necessidade de gateways intermediários.

---

## Recursos Implementados

1. **API Pix Bacen v2 Direta (Multi-Bancos)**:
   - Suporte nativo ao padrão do Banco Central (Bacen).
   - Compatível com Banco do Brasil, Itaú, Santander, Banco Inter, Sicoob, Bradesco, Efí (Gerencianet) e conexões customizadas.
   - Autenticação OAuth 2.0 com certificado digital cliente mTLS (.pfx / ICP-Brasil).
   - Criação de cobrança imediata (`PUT /v2/cob/{txid}`) com geração de payload `pixCopiaECola`.
   - Consulta de liquidação (`GET /v2/cob/{txid}`) com captura automática do `endToEndId`.
   - Modo Simulador integrado para agilizar o desenvolvimento e testes do time de software.

2. **Interface Frente de Caixa (PDV Supermercado) em VB6**:
   - Formulário modal `frmPixSupermercado.frm`.
   - Exibição de QR Code em tela e texto copia-e-cola.
   - Polling automático via Timer a cada 2,5 segundos.
   - Auto-fechamento imediato após a confirmação do pagamento no banco do consumidor.
   - Liberação dos dados para emissão e impressão instantânea da NFC-e.

3. **Split Payment e Reforma Tributária (RTC)**:
   - Geração das tags XML `<pag>` e `<infTransacPag>` conforme a NT 2026.006.
   - Vinculação posterior de pagamentos à SEFAZ via evento 110300.
   - Retornos estruturados em formato JSON padronizado (`STATUS`, `DESCRICAO`, `RETORNO`).

---

## Como Compilar e Registrar

```cmd
:: 1. Abra o Prompt de Comando do Desenvolvedor do Visual Studio como Administrador
:: 2. Compile a DLL:
msbuild BridgeRTC.csproj /p:Configuration=Release

:: 3. Registre no COM do Windows:
registrar_dll.bat
```

---

## Documentação Técnica Completa

Consulte o arquivo **`DOCUMENTACAO_PIX_BRIDGERTC.md`** na raiz deste repositório para o manual detalhado com diagramas de fluxo, credenciamento em bancos, pré-requisitos e testes em homologação e produção.

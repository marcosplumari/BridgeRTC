# Guia Técnico de Integração Direta Pix (Bacen v2) & Split Payment

Este documento fornece à equipe de desenvolvimento e arquitetura todas as instruções, fluxos de funcionamento, chamadas em VB6 via DLL C# `BridgeRTC`, credenciamento com instituições financeiras e procedimentos para homologação e produção.

---

## 1. Visão Geral da Arquitetura

O ecossistema opera em 3 camadas:
1. **Frente de Caixa (PDV) em Visual Basic 6.0**: Interface gráfica de venda e interação com o operador/cliente.
2. **DLL BridgeRTC (.NET Framework 4.8 / Interop COM)**: Camada de comunicação de alto nível com criptografia TLS 1.2, certificados mTLS e padronização JSON.
3. **API Pix Padronizada pelo Banco Central (Bacen)** e **WebServices SEFAZ/RTC**: Endpoints das instituições financeiras e da administração tributária.

```
+----------------+      COM Interop      +----------------------+      mTLS / REST      +-----------------------+
|  PDV em VB6    | --------------------> |   BridgeRTC (C# DLL) | --------------------> | API Pix Bacen (Banco) |
| (Supermercado) | <-------------------- | (.NET Framework 4.8) | <-------------------- | (BB, Itaú, Inter,...) |
+----------------+                       +----------------------+                       +-----------------------+
        |                                           |
        | Fecha janela ao confirmar                 | Emissão NFC-e com tPag 17 e idTransacao (endToEndId)
        v                                           v
+---------------------------------------------------------------------------------------------------------------+
|                                      SEFAZ / Reforma Tributária (RTC)                                         |
+---------------------------------------------------------------------------------------------------------------+
```

---

## 2. Requisitos Oficiais do Bacen (Sem Intermediários / Gateways)

### 2.1. A Informação de API Pública Direta está Correta?
**Sim.** O Banco Central do Brasil, ao criar o Regulamento do Pix (Resolução BCB nº 1/2020), definiu uma **especificação aberta e obrigatória** chamada **API Pix** para todas as instituições financeiras e de pagamento que ofertam conta para pessoas jurídicas.

- **Repositório Oficial do Bacen:** [github.com/bacen/pix-api](https://github.com/bacen/pix-api)
- **Documentação Swagger/OpenAPI:** [bacen.github.io/pix-api](https://bacen.github.io/pix-api/)
- **Padrão de Endpoints:**
  - `POST /oauth/token`: Obtenção do token OAuth 2.0 (mTLS obrigatório).
  - `PUT /v2/cob/{txid}`: Criação da cobrança imediata com o valor e tempo de expiração. Retorna o payload `pixCopiaECola` e `location`.
  - `GET /v2/cob/{txid}`: Consulta do status da cobrança (`ATIVA`, `CONCLUIDA`, etc.) e identificador `endToEndId`.

### 2.2. A Software House Precisa de Homologação ou Certificação no Bacen?
**Não.** A empresa desenvolvedora do ERP **não** necessita de credenciamento ou certificação junto ao Banco Central. 
Quem contrata e autoriza a API é o **lojista (correntista PJ)**:
1. O lojista acessa o portal corporativo do seu banco (ex: *Itaú Empresas*, *BB Developers*, *Santander Developers*, *Inter Empresas*, *Sicoob*, *Bradesco Net Empresa*).
2. Cria uma "Aplicação" para a API Pix do Bacen.
3. Faz o upload ou emissão do **Certificado Digital mTLS** (geralmente o e-CNPJ da própria empresa em `.pfx`).
4. Obtém o `Client_Id` e `Client_Secret`.
5. O ERP apenas recebe essas chaves e o caminho do arquivo `.pfx` nas configurações do sistema.

---

## 3. Fluxo de Caixa de Supermercado (Passo a Passo)

```
[Operador]             [PDV VB6]             [BridgeRTC]             [Banco Bacen]            [Cliente Celular]
    |                      |                      |                        |                          |
    |-- Finaliza Venda --->|                      |                        |                          |
    |   (Forma: PIX)       |-- IniciarCobranca -->|                        |                          |
    |                      |                      |-- PUT /v2/cob/{txid} ->|                          |
    |                      |                      |<-- Retorna CopiaCola --|                          |
    |                      |<- Exibe QR Code -----|                        |                          |
    |                      |   (Abre Modal)       |                        |                          |
    |                      |                      |                        |<-- Escaneia e Paga ------|
    |                      |-- Polling 2.5s ----->|                        |                          |
    |                      |                      |-- GET /v2/cob/{txid} ->|                          |
    |                      |                      |<-- Status: CONCLUIDA --|                          |
    |                      |<- Fecha Modal Pix ---|    (endToEndId)        |                          |
    |                      |                                                                          |
    |                      |-- Monta XML NFC-e com tPag=17 e idTransacao (endToEndId)                |
    |                      |-- Envia NFC-e à SEFAZ e imprime cupom fiscal                             |
    |                      |-- Limpa a tela para o próximo cliente                                    |
```

---

## 4. Métodos Disponíveis na DLL BridgeRTC

| Método | Finalidade | Parâmetros Principais | Retorno |
| :--- | :--- | :--- | :--- |
| `ConfigurarAmbiente` | Define Produção (1) ou Homologação (2) | `tpAmb (1 ou 2)`, `uf ("SP", "RJ", ...)` | `void` |
| `ConfigurarPix` | Configura credenciais e banco | `instituicao`, `clientId`, `clientSecret`, `chavePix`, `caminhoPfx`, `senha` | `void` |
| `CriarCobrancaPix` | Gera a cobrança e o QR Code no Bacen | `txid`, `valor`, `expiracaoSegundos`, `solicitacaoPagador` | JSON com `STATUS`, `DESCRICAO`, `RETORNO` (`pixCopiaECola`, `location`) |
| `ConsultarCobrancaPix` | Consulta se foi pago | `txid` | JSON com `status ("CONCLUIDA")`, `pago (true/false)`, `endToEndId` |
| `SimularPagamentoPix` | Confirma o pagamento no mock local | `txid` | JSON com status concluído imediato |
| `GerarGrupoPagamentoXml` | Gera o XML de pagamento da NFC-e | `codigoMeioPagto`, `valor`, `idTransacao`, `cnpjInstituicao`, `cAut` | XML pronto para a tag `<pag>` |
| `VincularPagamentoDFe` | Envia evento SEFAZ (Reforma Tributária) | `chaveDFe`, `idTransacao`, `codigoMeioPagto`, `valor`, `cnpjInst`, `cAut` | Retorno da SEFAZ |

---

## 5. Como o Time de Desenvolvimento Deve Testar

### 5.1. Teste Rápido (Sem Dependência de Bancos - Modo Simulador)
Para testar a interface do VB6 imediatamente sem gastar centavos ou depender de aprovação de contas:
1. No VB6, abra o projeto `ExemploVB6\ProjetoExemplo.vbp`.
2. Clique no botão **"Abrir Tela PIX Supermercado"**.
3. O formulário `frmPixSupermercado` será carregado com um QR Code gerado instantaneamente e iniciará o polling.
4. Clique no botão **"Simular Pagamento (Dev/Teste)"**.
5. O timer detectará a transação como `CONCLUIDA`, capturará o `endToEndId` gerado, tocará o sinal sonoro, fechará o formulário automaticamente e exibirá a confirmação para emissão da NFC-e.

### 5.2. Teste em Homologação com Bancos Reais
Para testar com um banco parceiro (ex: Banco do Brasil, Itaú, Santander, Inter, Sicoob ou Efí):
1. No formulário VB6, configure as credenciais obtidas no portal de desenvolvedores do banco:
   ```vb
   mBridge.ConfigurarAmbiente 2, "SP" ' 2 = Homologacao
   mBridge.ConfigurarPix "BANCO_DO_BRASIL", "SEU_CLIENT_ID_SANDBOX", "SEU_CLIENT_SECRET", "sua-chave-pix@empresa.com", "C:\certificados\pix_homolog.pfx", "senha123"
   ```
2. Ao chamar `mBridge.CriarCobrancaPix`, a DLL fará a requisição mTLS real aos servidores de sandbox do banco.
3. No sandbox de cada banco, utilize as ferramentas de simulação fornecidas no portal (ex: webhook simulator ou botão de pagamento fake) para aprovar a transação.

### 5.3. Mudança para PRODUÇÃO
Quando o cliente for entrar em operação real:
1. No Internet Banking PJ do cliente, gere as credenciais de **Produção**.
2. Altere o parâmetro de ambiente para `1`:
   ```vb
   mBridge.ConfigurarAmbiente 1, "SP" ' 1 = Producao
   mBridge.ConfigurarPix "BANCO_DO_BRASIL", txtClientId.Text, txtClientSecret.Text, txtChavePix.Text, txtCaminhoPfx.Text, txtSenhaCert.Text
   ```
3. O cliente realiza um Pix real de R$ 1,00 lendo com o app de qualquer banco para validar o encerramento do caixa e a impressão da NFC-e.

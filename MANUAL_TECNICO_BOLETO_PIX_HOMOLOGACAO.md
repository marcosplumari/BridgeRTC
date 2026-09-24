# Manual Técnico: Pix Cobrança, Boleto Híbrido e Transição Homologação/Produção

**Projeto:** BridgeRTC (ERP VB6 + DLL C# .NET Framework 4.8 via COM Interop)  
**Autor:** Equipe de Arquitetura & Integrações  
**Destinatários:** Desenvolvedores VB6, Engenheiros C#, Equipe de Suporte e Implantação  

---

## 1. Visão Geral da Arquitetura

O ERP legado em Visual Basic 6 (VB6) não possui suporte nativo a protocolos modernos de segurança (TLS 1.2/1.3, mTLS com certificados A1, JWT, OAuth 2.0 e JSON REST). A DLL **BridgeRTC** atua como uma ponte transparente (*COM Interop*), encapsulando:
1. Autenticação OAuth 2.0 com mTLS e renovação automática de tokens.
2. Chamadas aos endpoints padronizados do Banco Central (Pix Cobrança `/v2/cob` e Boleto Híbrido com Vencimento `/v2/cobv`).
3. Cálculos financeiros de juros, multa, descontos e tributos do Split Payment (LC 214/2025).
4. Retorno estruturado em JSON padronizado com campos `STATUS` (`OK` ou `ERRO`), `DESCRICAO` e `RETORNO`.

---

## 2. Perguntas Frequentes & Diretrizes de Negócio

### 2.1. Client ID, Client Secret e Certificado Digital A1
* **Certificado Digital A1 (mTLS):** É o mesmo para toda a comunicação com a instituição bancária. Trata-se do certificado da empresa/lojista (.pfx) emitido por Autoridade Certificadora ICP-Brasil, responsável por autenticar o túnel seguro com o banco.
* **Client ID e Client Secret:** Em bancos tradicionais (Itaú, Banco do Brasil, Santander, Bradesco), os portais para desenvolvedores costumam separar as credenciais por **Aplicação / Produto**. Há uma aplicação para **Pix** e outra para **Cobrança Bancária / Boletos**. Em fintechs (Inter, Efí/Gerencianet), as credenciais podem ser compartilhadas na mesma chave, configurando-se os escopos permitidos no painel.

### 2.2. Modelo de Cobrança e Tarifas Bancárias
* **Cobrança por Liquidação:** Vale a regra de ouro da FEBRABAN: **só há cobrança de tarifa quando o título é efetivamente pago**. Não existe cobrança cumulativa (se o cliente pagar via Pix, cobra-se apenas a tarifa acordada para o boleto).
* **Valor Fixo:** A cobrança no Brasil para boletos e boletos híbridos é em valor fixo por título liquidado (ex: R$ 1,50 a R$ 3,50, conforme negociação com a agência), e não percentual.
* **Consultas via API Gratuitas:** O banco **não cobra** por requisições de consulta (`GET /cob` ou `GET /cobv`). Seu ERP pode consultar o status do título em tempo real sem custo adicional.

### 2.3. Leiaute do Boleto no ActiveReports
* **Estrutura Tradicional Preservada:** O código de barras 2 de 5 intercalado (44 dígitos) e a Linha Digitável devem permanecer inalterados na parte inferior da folha (Ficha de Compensação), respeitando as margens e dimensões homologadas.
* **Posicionamento do QR Code Pix:** Inserir o QR Code preferencialmente no **Recibo do Pagador** (topo da página A4) ou à esquerda dos dados do beneficiário. Dimensões recomendadas: entre **2,5 cm e 3,0 cm** para garantir leitura rápida pelo celular.
* **Identificação Visual:** Incluir o logotipo oficial do Pix e o aviso explicativo: *"Pague com Pix usando o QR Code acima ou utilize o código de barras abaixo"*.

### 2.4. Conciliação: Retorno CNAB 400/240 vs Consulta Online
* **Compatibilidade com o CNAB:** O campo `Nosso Número` registrado via API é o mesmo que constará no arquivo de retorno CNAB gerado pelo banco ao final do dia (ocorrência 06 - Liquidação). O leitor de CNAB existente no seu ERP continuará processando a baixa automaticamente.
* **Baixa Imediata por Tela de Consulta:** O operador pode abrir uma tela avulsa de pesquisa de títulos no ERP e consultar o status diretamente via API. Se o retorno indicar status `LIQUIDADO` ou `CONCLUIDA`, a baixa no Contas a Receber pode ser efetuada no mesmo instante, sem aguardar o arquivo de retorno noturno.

---

## 3. As Três Camadas de Teste

| Camada | Nome | O que Testa | Comunicação de Rede | Risco |
|---|---|---|---|---|
| **Camada 1** | **Simulador Interno** | Lógica do VB6, timers, interface gráfica, ActiveReports, gravação no banco de dados local. | Nenhuma (mock em memória). | Zero. |
| **Camada 2** | **Sandbox / Homologação** | Autenticação mTLS real com o banco, formatação de payloads, validação de schemas bancários. | Conecta aos endpoints de teste dos bancos. | Zero (moeda fictícia). |
| **Camada 3** | **Produção Controlada** | Fluxo ponta a ponta com dinheiro real, conciliação de caixa e baixa de títulos. | Conecta aos endpoints de produção dos bancos. | Valor real (usar R$ 0,01 a R$ 1,00). |

---

## 4. Configuração no VB6 e na DLL C#

### 4.1. Como alternar entre os ambientes no VB6

Na chamada de inicialização da DLL, configure os parâmetros conforme o ambiente desejado:

```vb
Dim mBridge As Object
Set mBridge = CreateObject("BridgeRTC.BridgeRTCService")

' =============================================================================
' CAMADA 1: MODO SIMULADOR (Padrao para desenvolvimento interno)
' =============================================================================
mBridge.ConfigurarAmbiente 2, "SP"
mBridge.ConfigurarPix "SIMULADOR", "", "", "suporte@meuerp.com.br", "", ""

' =============================================================================
' CAMADA 2: HOMOLOGACAO / SANDBOX BANCARIO
' =============================================================================
' mBridge.ConfigurarAmbiente 2, "SP"  ' 2 = Homologacao
' mBridge.ConfigurarPix "ITAU", "CLIENT_ID_TESTE", "CLIENT_SECRET_TESTE", "chave_pix_hm", "C:\Certificados\banco_hm.pfx", "SenhaCert123"

' Se o banco exigir URLs customizadas de sandbox:
' mBridge.ConfigurarPixCustomizado "https://api-sandbox.banco.com.br/pix/v2", _
'                                 "https://oauth-sandbox.banco.com.br/token", _
'                                 "CLIENT_ID_TESTE", "CLIENT_SECRET_TESTE", _
'                                 "chave_pix_hm", "C:\Certificados\banco_hm.pfx", "SenhaCert123"

' =============================================================================
' CAMADA 3: PRODUCAO (Ambiente real de vendas)
' =============================================================================
' mBridge.ConfigurarAmbiente 1, "SP"  ' 1 = Producao
' mBridge.ConfigurarPix "ITAU", "CLIENT_ID_PROD", "CLIENT_SECRET_PROD", "chave_pix_real", "C:\Certificados\empresa_prod.pfx", "SenhaCert123"
```

---

## 5. Roteiro Passo a Passo para Configuração no Banco e Virada de Chave (Go-Live)

### Etapa 1: Cadastro no Portal de Desenvolvedores
1. O titular da conta corrente (lojista ou representante legal) acessa o portal para desenvolvedores do banco:
   * **Itaú:** Itaú Developers / Portal de APIs Itaú Empresas
   * **Banco do Brasil:** Portal Developers BB
   * **Santander:** Santander Devs
   * **Inter:** Conta Digital PJ -> Aba de Integrações / APIs
   * **Bradesco:** Portal do Desenvolvedor Bradesco
2. Cria-se uma nova **Aplicação / App**.
3. Selecionam-se os escopos necessários:
   * Pix Cobrança Imediata: `cob.write`, `cob.read`, `pix.read`, `pix.write` (para devolução).
   * Boleto Híbrido / CobV: `cobv.write`, `cobv.read`.

### Etapa 2: Instalação do Certificado Digital A1
1. Exportar o certificado A1 da empresa com chave privada no formato `.pfx` ou `.p12`.
2. Em alguns bancos (ex: Itaú/Santander), deve-se fazer o upload da chave pública do certificado no portal do banco para autorização do mTLS.
3. Salvar o arquivo `.pfx` em uma pasta com permissão de leitura para o usuário do Windows que executa o ERP (ex: `C:\ERP\Certificados\cert_banco.pfx`).

### Etapa 3: Homologação contra o Sandbox Bancário
1. No portal do desenvolvedor, obter as chaves de teste: `Client_Id` e `Client_Secret` de sandbox.
2. Configurar o ERP com `ConfigurarAmbiente 2, "SP"`.
3. Executar o fluxo no VB6:
   * Criar Cobrança Pix (`CriarCobrancaPix`).
   * Criar Boleto Híbrido (`CriarBoletoHibridoPix`).
   * No painel do desenvolvedor do banco, utilizar a ferramenta de simulação ("Simular Pagamento") fornecida pelo sandbox para liquidar o `txid`.
   * Verificar se o ERP recebe o status `CONCLUIDA` e fecha o modal corretamente.

### Etapa 4: Solicitação de Virada de Chave para Produção
1. Com os testes de sandbox validados, solicitar no portal a "Publicação em Produção" ou "Virada de Chave".
2. Vincular a aplicação à **Conta Corrente PJ** ativa da empresa que receberá os créditos.
3. O banco gerará as credenciais definitivas de Produção (`Client_Id` e `Client_Secret`).

### Etapa 5: Teste Real em Produção (Smoketest)
1. No cadastro de parâmetros do ERP, apontar para `ConfigurarAmbiente 1, "SP"` e inserir as credenciais de produção.
2. Emitir uma venda no caixa com Pix no valor de **R$ 0,01** ou **R$ 1,00**.
3. O operador lê o QR Code com o aplicativo de qualquer banco no celular e paga.
4. O ERP detecta o pagamento, fecha a tela e emite a NFC-e.
5. Em seguida, testar o botão de estorno/devolução (`DevolverPix`) para comprovar o ciclo completo.

---

## 6. Checklist de Validação para a Equipe de TI

- [ ] DLL `BridgeRTC.dll` compilada em `Release | x86` ou `Any CPU` (.NET Framework 4.8).
- [ ] Registro COM efetuado via `regasm.exe BridgeRTC.dll /codebase /tlb` com permissões de administrador.
- [ ] Certificado A1 (.pfx) instalado no disco do PDV/Servidor com caminho e senha preenchidos.
- [ ] Teste de emissão com Simulação Interna concluído com sucesso.
- [ ] Teste de conexão mTLS com as URLs de homologação do banco concluído.
- [ ] Impressão do Boleto Híbrido no ActiveReports validada com o QR Code nítido e código de barras 2 de 5 legível por leitor ótico.
- [ ] Smoketest de R$ 1,00 em produção executado e liquidado na conta corrente da empresa.

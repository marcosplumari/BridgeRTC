# BridgeRTC

Biblioteca em C# (.NET Framework 4.8) com interoperabilidade COM para **Visual Basic 6.0 (VB6)**, desenvolvida para conectar sistemas legados e ERPs à **Reforma Tributária (Lei Complementar nº 214/2025 - Split Payment)** e à **API Pix do Banco Central do Brasil (Bacen v2)** diretamente com os bancos, sem necessidade de gateways intermediários.

---

## Índice
1. [Visão Geral da Arquitetura](#1-visão-geral-da-arquitetura)
2. [Requisitos de Certificado Digital (A1 Obrigatório)](#2-requisitos-de-certificado-digital-a1-obrigatório)
3. [Fluxo de Configuração Passo a Passo](#3-fluxo-de-configuração-passo-a-passo)
4. [Tabela de Endpoints Oficiais (Bacen v2 e SEFAZ)](#4-tabela-de-endpoints-oficiais-bacen-v2-e-sefaz)
5. [Padrão de Resposta JSON e Códigos de Erro](#5-padrão-de-resposta-json-e-códigos-de-erro)
6. [Exemplo Completo de Uso no VB6](#6-exemplo-completo-de-uso-no-vb6)
7. [Exemplo Completo de Uso no C#](#7-exemplo-completo-de-uso-no-c)
8. [Compilação e Registro da DLL no Windows](#8-compilação-e-registro-da-dll-no-windows)

---

## 1. Visão Geral da Arquitetura

O `BridgeRTC` atua como uma ponte de comunicação de alta performance:
- **No VB6**: É instanciado via COM (`CreateObject("BridgeRTC.BridgeRTCService")`), operando com métodos simples que recebem e devolvem strings JSON.
- **Internamente em C#**: Gerencia conexões seguras sob TLS 1.2, autenticação mTLS (Mutual TLS), chamadas REST HTTP, cálculos de segregação de Split Payment e montagem de tags fiscais para a NFC-e / NF-e.
- **Nos Bancos**: Comunica-se com os endpoints oficiais padronizados pelo Bacen (Itaú, Banco do Brasil, Santander, Inter, Sicoob, Bradesco, Efí ou simulador de testes).

```
+-------------------+           COM Interop          +--------------------------+
|  Sistema em VB6   | <----------------------------> |   BridgeRTC.dll (C#)     |
| (PDV / Caixa ERP) |   (JSON STATUS/DESCRICAO)      |  (.NET Framework 4.8)    |
+-------------------+                                +--------------------------+
                                                      /                        \
                                 mTLS / REST (Bacen) /                          \  SOAP / WebServices
                                                    v                            v
                                    +-----------------------+        +-----------------------+
                                    |   API Pix dos Bancos  |        |      SEFAZ / RTC      |
                                    |  (BB, Itaú, Inter...) |        | (Evento 110300/NFC-e) |
                                    +-----------------------+        +-----------------------+
```

---

## 2. Requisitos de Certificado Digital (A1 Obrigatório)

Para a comunicação com os bancos e SEFAZ via API, o uso de **Certificado Digital A1 (arquivo `.pfx`) é estritamente obrigatório**.

### Por que o A1 e nunca o A3 (Token USB / Cartão SmartCard)?
1. **Sem bloqueio de tela:** O certificado A3 exige digitação manual de senha PIN em janelas modais do Windows. Em um caixa de supermercado realizando dezenas de transações por hora, o caixa travaria a cada consulta.
2. **Compatibilidade mTLS:** O handshake mTLS dos bancos exige que a aplicação anexe a chave privada na requisição HTTP. O arquivo A1 (`.pfx`) permite exportação de chaves em memória (`X509KeyStorageFlags.MachineKeySet | Exportable`).

### Como obter e configurar:
- O lojista utiliza o mesmo e-CNPJ A1 que já usa para emitir notas fiscais, ou um certificado gerado no próprio portal de desenvolvedores do banco.
- Salve o arquivo em local acessível (ex: `C:\MeuERP\Certificados\empresa.pfx`).
- A DLL carrega o arquivo diretamente em memória passando o caminho e a senha.

---

## 3. Fluxo de Configuração Passo a Passo

### 3.1. Credenciamento no Banco (Realizado pelo Lojista)
1. O lojista acessa o portal corporativo do seu banco (ex: *Itaú Empresas*, *BB Developers*, *Santander Developers*, *Inter Empresas*).
2. Cria uma aplicação para **API Pix**.
3. Faz o upload da chave pública do seu certificado A1.
4. Obtém o `Client_Id` e `Client_Secret`.
5. Cadastra uma **Chave Pix** (CNPJ, e-mail, telefone ou chave aleatória) associada à conta da empresa.

### 3.2. Configuração no Software
No momento da inicialização do caixa:
```vb
' 1. Configura Ambiente: 1 = Produção | 2 = Homologação / Sandbox
bridge.ConfigurarAmbiente 2, "SP"

' 2. Configura a Instituição Financeira
' Opções: "BANCO_DO_BRASIL", "ITAU", "SANTANDER", "INTER", "SICOOB", "BRADESCO", "EFIPAY", ou "SIMULADOR"
bridge.ConfigurarPix "BANCO_DO_BRASIL", "SEU_CLIENT_ID", "SEU_CLIENT_SECRET", "chave-pix@empresa.com.br", "C:\cert\empresa.pfx", "senha123"
```

---

## 4. Tabela de Endpoints Oficiais (Bacen v2 e SEFAZ)

A biblioteca implementa rigorosamente a especificação aberta do Banco Central do Brasil ([Regulamento Bacen v2](https://github.com/bacen/pix-api)):

| Operação | Método HTTP | Endpoint Bacen / SEFAZ | Função na BridgeRTC |
| :--- | :---: | :--- | :--- |
| **Token OAuth** | `POST` | `/oauth/token` (mTLS) | Automático antes de cada requisição |
| **Criar Cobrança Pix** | `PUT` | `/v2/cob/{txid}` | `CriarCobrancaPix` |
| **Consultar Pagamento** | `GET` | `/v2/cob/{txid}` | `ConsultarCobrancaPix` |
| **Estornar/Devolver Pix** | `PUT` | `/v2/pix/{e2eid}/devolucao/{id}` | `DevolverPix` |
| **Consultar Devolução** | `GET` | `/v2/pix/{e2eid}/devolucao/{id}` | `ConsultarDevolucaoPix` |
| **Conciliação Split** | `GET` | `/v2/cob/{txid}` (componentesValor) | `ObterConciliacaoFinanceiraPix` |
| **XML NFC-e (tpIntegra)** | N/A | Tag fiscal `<pag>` com `<tpIntegra>1` | `GerarGrupoPagamentoXml` |
| **Vincular DFe SEFAZ** | `POST` | Evento SEFAZ `110300` | `VincularPagamentoDFe` |

---

## 5. Padrão de Resposta JSON e Códigos de Erro

Todas as funções retornam uma string JSON estruturada:
```json
{
  "STATUS": "OK" | "ERRO",
  "DESCRICAO": "Mensagem amigável com a descrição da resposta",
  "RETORNO": { ... }
}
```

### 5.1. Exemplo de Retorno com Sucesso (`ConsultarCobrancaPix` - Pagamento Confirmado)
```json
{
  "STATUS": "OK",
  "DESCRICAO": "Pagamento Pix confirmado no PSP",
  "RETORNO": {
    "txid": "PDV202609221200001",
    "status": "CONCLUIDA",
    "pago": true,
    "endToEndId": "E000000002026092212000000001",
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
    }
  }
}
```

### 5.2. Tabela de Tratamento de Erros

| Situação | Causa Raiz | Mensagem Retornada | Como Resolver |
| :--- | :--- | :--- | :--- |
| **Certificado Inexistente** | Caminho do arquivo `.pfx` incorreto | `Falha ao carregar certificado digital: Arquivo nao encontrado` | Verificar caminho do arquivo no disco |
| **Senha Incorreta** | Senha do certificado digital inválida | `Falha ao carregar certificado digital: The specified network password is not correct` | Corrigir a senha configurada no ERP |
| **Chave DFe Inválida** | Chave com tamanho diferente de 44 dígitos | `Chave do DF-e invalida ou incompleta (deve conter 44 digitos)` | Passar a chave completa da NFC-e/NFe |
| **HTTP 401 Unauthorized** | Client ID ou Client Secret incorretos no banco | `Erro HTTP (Unauthorized): Client authentication failed` | Revisar credenciais geradas no portal do banco |
| **HTTP 429 Too Many Requests** | Excesso de requisições por segundo (Rate Limit) | `Erro HTTP (TooManyRequests)` | Ajustar o intervalo do polling no PDV (2,5s a 3s) |
| **HTTP 504 / Timeout** | Instabilidade momentânea na internet da loja | `Erro HTTP: The operation has timed out` | Utilizar o botão "Reconsultar Banco Agora" |
| **Pagamento Pendente** | Cliente ainda não confirmou no aplicativo | `STATUS: OK`, com `status: ATIVA` e `pago: false` | Manter a tela aberta aguardando ou aguardar timeout |

---

## 6. Exemplo Completo de Uso no VB6

```vb
Dim bridge As Object
Set bridge = CreateObject("BridgeRTC.BridgeRTCService")

' 1. Configura Ambiente e Banco
bridge.ConfigurarAmbiente 2, "SP" ' 1 = Produção, 2 = Homologação
bridge.ConfigurarPix "BANCO_DO_BRASIL", "CLIENT_ID", "CLIENT_SECRET", "chave@empresa.com", "C:\cert\pix.pfx", "senha123"

' 2. Cria Cobrança Imediata para o Caixa (3 minutos de validade)
Dim sRespCob As String, txid As String
txid = "PDV" & Format(Now, "yyyymmddhhnnss")
sRespCob = bridge.CriarCobrancaPix(txid, 100.0, 180, "Venda no Caixa 01")

' 3. Consulta periódica (dentro de um Timer a cada 2.5s)
Dim sRespConsulta As String
sRespConsulta = bridge.ConsultarCobrancaPix(txid)

If InStr(sRespConsulta, """status"":""CONCLUIDA""") > 0 Then
    ' Pagamento confirmado pelo banco!
    ' 4. Gera o XML de pagamento com tpIntegra=1 para a NFC-e
    Dim sXmlPagto As String
    sXmlPagto = bridge.GerarGrupoPagamentoXml("17", 100.0, "E000000002026092200001", "00000000000191", "AUTH123")
    
    ' 5. Conciliação no Contas a Receber
    ' - Baixa R$ 100,00 no cliente
    ' - Registra entrada real no banco de R$ 72,51
    ' - Registra retenção tributária de R$ 26,50 (Split Payment LC 214/2025)
End If

' 6. Se o cliente desistir ou der erro de impressão, efetua o estorno imediato:
Dim sRespEstorno As String
sRespEstorno = bridge.DevolverPix("E000000002026092200001", "DEV001", 100.0, "Desistencia da compra no PDV")
```

---

## 7. Exemplo Completo de Uso no C#

```csharp
using System;
using BridgeRTC;

class Program
{
    static void Main()
    {
        var bridge = new BridgeRTCService();
        
        // 1. Configurações
        bridge.ConfigurarAmbiente(tpAmb: 2, uf: "SP");
        bridge.ConfigurarPix(
            instituicao: "ITAU",
            clientId: "SEU_CLIENT_ID",
            clientSecret: "SEU_CLIENT_SECRET",
            chavePix: "sua-chave-pix",
            caminhoCertificadoPfx: @"C:\certificados\empresa.pfx",
            senhaCertificado: "senha123"
        );

        // 2. Criar cobrança
        string txid = "PDV" + DateTime.Now.ToString("yyyyMMddHHmmss");
        string jsonCob = bridge.CriarCobrancaPix(txid, 100.00, 180, "Venda Loja Matriz");
        Console.WriteLine("Cobranca Gerada: " + jsonCob);

        // 3. Consultar liquidação
        string jsonConsulta = bridge.ConsultarCobrancaPix(txid);
        Console.WriteLine("Status Pagamento: " + jsonConsulta);

        // 4. Em caso de cancelamento da venda, estorna
        string jsonDevolucao = bridge.DevolverPix("E000000002026092200001", "DEV" + txid, 100.00, "Cancelamento PDV");
        Console.WriteLine("Devolucao: " + jsonDevolucao);
    }
}
```

---

## 8. Compilação e Registro da DLL no Windows

Para compilar e registrar a biblioteca em computadores com Windows e VB6:

```cmd
:: 1. Abra o Prompt de Comando do Desenvolvedor do Visual Studio como Administrador
:: 2. Compile a DLL em modo Release:
msbuild BridgeRTC.csproj /p:Configuration=Release

:: 3. Registre no mecanismo COM do Windows:
:: Para sistemas 64-bit que rodam VB6 (32-bit):
%SystemRoot%\Microsoft.NET\Framework\v4.0.30319\regasm.exe bin\Release\BridgeRTC.dll /codebase /tlb

:: Ou execute diretamente o arquivo em lote incluído na raiz:
registrar_dll.bat
```


### Suporte a Qualquer Banco ou Instituição (Configuração Customizada)

Para bancos ou cooperativas que não possuem URLs fixas embutidas no switch da DLL (como Sicredi, Caixa, Safra, cooperativas regionais, etc.), utilize o método `ConfigurarPixCustomizado`:

```vb
' Exemplo VB6 para qualquer instituição bancária:
Dim respConfig As String
Dim urlBase As String
Dim urlOAuth As String

urlBase = "https://api-pix.sicredi.com.br/pix/v2"
urlOAuth = "https://api-pix.sicredi.com.br/oauth/v2/token"

respConfig = bridge.ConfigurarPixCustomizado( _
    urlBase, _
    urlOAuth, _
    "SEU_CLIENT_ID", _
    "SEU_CLIENT_SECRET", _
    "suachave@pix.com.br", _
    "C:\Certificados\empresa.pfx", _
    "123456" _
)

' respConfig retorna JSON com STATUS = "OK" ou "ERRO"
' Validações automáticas:
' - URL Base e URL OAuth devem ser HTTPS absoluto
' - Validação e teste de abertura do certificado A1 (.pfx)
' - Limpeza e preparação segura do cache de autenticação
```

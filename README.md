# BridgeRTC

Biblioteca de classes em **C# (.NET Framework 4.8)** com suporte a **COM Interop**, desenvolvida para atuar como ponte entre softwares legado em **Visual Basic 6 (VB6)** e as exigências da **Reforma Tributária do Consumo (RTC)** e **Split Payment** (Lei Complementar e Nota Técnica 2026.006).

---

## 🎯 Objetivo

Fornecer ao seu ERP/PDV uma interface simples para:
1. **Emissão Imediata:** Gerar o bloco XML de pagamento (`<pag><detPag><infTransacPag>`) com o grupo YC (Split Payment) para inclusão direta na NF-e (mod 55) e NFC-e (mod 65).
2. **Vinculação Posterior:** Transmitir para a SEFAZ o **Evento de Vinculação da Transação de Pagamento (Código 110300)** amarrando a chave de 44 dígitos da nota ao identificador da transação bancária/Pix/TEF.

---

## 📦 Estrutura do Repositório

```text
├── BridgeRTC.sln                   # Solução para Visual Studio 2019 / 2022
├── BridgeRTC.csproj                # Projeto C# configurado com RegisterForComInterop
├── BridgeRTCService.cs             # Código-fonte da classe COM e interface
├── Properties/
│   └── AssemblyInfo.cs             # Configurações de visibilidade COM e GUIDs
├── ExemploVB6/
│   ├── ProjetoExemplo.vbp          # Arquivo de projeto do Visual Basic 6
│   └── Form1.frm                   # Tela de exemplo no Visual Basic 6
├── registrar_dll.bat               # Script para registrar a DLL no Windows via RegAsm
├── .gitignore                      # Filtro padrão do Git para binários e temporários
└── README.md                       # Documentação técnica
```

---

## 🚀 Como Compilar e Registrar

### 1. Compilação no Visual Studio
1. Abra o arquivo `BridgeRTC.sln`.
2. Selecione a configuração **Release | Any CPU**.
3. Compile a solução (**Build Solution** ou `Ctrl + Shift + B`).
4. A DLL será gerada na pasta `bin\Release\BridgeRTC.dll`.

### 2. Registro no Computador Cliente (Windows)
Execute o arquivo `registrar_dll.bat` como Administrador:

```cmd
%SystemRoot%\Microsoft.NET\Framework\v4.0.30319\regasm.exe bin\Release\BridgeRTC.dll /codebase /tlb
```

---

## 💻 Exemplo de Uso no Visual Basic 6

```vb
Dim oBridge As Object
Dim sJson As String

' 1. Cria a instancia COM
Set oBridge = CreateObject("BridgeRTC.BridgeRTCService")

' 2. Configura ambiente (2 = Homologacao, 1 = Producao)
oBridge.ConfigurarAmbiente 2, "SP"

' 3. Executa a vinculacao posterior (Evento 110300)
sJson = oBridge.VincularPagamentoDFe( _
    "35260900000000000000550010000000011000000010", _
    "PIX-E2E-ID1234567890", _
    "17", _
    150.75, _
    "00000000000191", _
    "AUT987654" _
)

MsgBox sJson
Set oBridge = Nothing
```

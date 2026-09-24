VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Teste BridgeRTC - Split Payment e Pix Bacen"
   ClientHeight    =   9200
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7800
   LinkTopic       =   "Form1"
   ScaleHeight     =   9200
   ScaleWidth      =   7800
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdStatus 
      Caption         =   "1. Consultar Status Servico"
      Height          =   450
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   3400
   End
   Begin VB.CommandButton cmdGerarXml 
      Caption         =   "2. Gerar Grupo Pagto XML (tpIntegra=1)"
      Height          =   450
      Left            =   3960
      TabIndex        =   1
      Top             =   240
      Width           =   3400
   End
   Begin VB.CommandButton cmdVincular 
      Caption         =   "3. Testar Vinculacao SEFAZ (Evento 110300)"
      Height          =   450
      Left            =   360
      TabIndex        =   2
      Top             =   780
      Width           =   7095
   End
   Begin VB.CommandButton cmdConciliacao 
      Caption         =   "4. Conciliacao Financeira Split Payment"
      Height          =   450
      Left            =   360
      TabIndex        =   3
      Top             =   1320
      Width           =   7095
   End
   Begin VB.CommandButton cmdEstornoPix 
      Caption         =   "5. Testar Devolucao / Estorno de Pix no Caixa"
      Height          =   450
      Left            =   360
      TabIndex        =   4
      Top             =   1860
      Width           =   7095
   End
   Begin VB.CommandButton cmdBancoCustomizado 
      Caption         =   "6. Testar Banco Customizado (ConfigurarPixCustomizado)"
      Height          =   450
      Left            =   360
      TabIndex        =   5
      Top             =   2400
      Width           =   7095
   End
   Begin VB.CommandButton cmdAbrirBoletoPix 
      Caption         =   "7. Testar Boleto Hibrido com Pix (ActiveReports / CobV)"
      Height          =   450
      Left            =   360
      TabIndex        =   6
      Top             =   2940
      Width           =   7095
   End
   Begin VB.CommandButton cmdAbrirPixSupermercado 
      Caption         =   "Abrir Tela PIX Supermercado (Modal / Auto Fechamento)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   360
      TabIndex        =   6
      Top             =   3480
      Width           =   7095
   End
   Begin VB.Label lblStatus 
      Caption         =   "Retorno JSON da DLL:"
      Height          =   255
      Left            =   360
      TabIndex        =   7
      Top             =   4100
      Width           =   2000
   End
   Begin VB.TextBox txtResultado 
      Height          =   4600
      Left            =   360
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   8
      Top             =   4380
      Width           =   7095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAbrirPixSupermercado_Click()
    Load frmPixSupermercado
    frmPixSupermercado.IniciarCobranca 100#, "Venda Caixa 01 - Supermercado"
    frmPixSupermercado.Show vbModal
    
    If frmPixSupermercado.PagamentoConfirmado Then
        Dim sLog As String
        sLog = "=== PAGAMENTO PIX CONFIRMADO NO CAIXA ===" & vbCrLf & _
               "EndToEndId: " & frmPixSupermercado.EndToEndId & vbCrLf & vbCrLf & _
               "----- CONCILIACAO FINANCEIRA (LANCAMENTO CONTAS A RECEBER) -----" & vbCrLf & _
               "Valor Bruto da Venda............: R$ " & FormatNumber(frmPixSupermercado.ValorBruto, 2) & vbCrLf & _
               "(-) CBS retida (Governo Federal): R$ " & FormatNumber(frmPixSupermercado.ValorCbs, 2) & vbCrLf & _
               "(-) IBS retido (Estados/Munic.)..: R$ " & FormatNumber(frmPixSupermercado.ValorIbs, 2) & vbCrLf & _
               "(-) Total Impostos Retidos......: R$ " & FormatNumber(frmPixSupermercado.ValorTributosRetidos, 2) & vbCrLf & _
               "(-) Tarifa Transacional PSP.....: R$ " & FormatNumber(frmPixSupermercado.ValorTarifa, 2) & vbCrLf & _
               "(=) LIQUIDO A ENTRAR NO CAIXA...: R$ " & FormatNumber(frmPixSupermercado.ValorLiquido, 2) & vbCrLf & vbCrLf & _
               "Emitindo NFC-e com tpIntegra=1 e tPag=17 (Pix)..." & vbCrLf & vbCrLf & _
               "Retorno completo JSON da DLL:" & vbCrLf & _
               frmPixSupermercado.JsonUltimoRetorno
        txtResultado.Text = sLog
        MsgBox "Pagamento concluído! Líquido de R$ " & FormatNumber(frmPixSupermercado.ValorLiquido, 2) & " disponibilizado.", vbInformation, "Venda Concluída"
    Else
        txtResultado.Text = "Operação Pix cancelada ou expirada pelo operador."
    End If
    
    Unload frmPixSupermercado
End Sub

Private Sub cmdBancoCustomizado_Click()
    On Error GoTo TrataErro
    Dim bridge As Object
    Set bridge = CreateObject("BridgeRTC.BridgeRTCService")
    
    ' Exemplo configurando uma instituicao que nao esta no switch fixo (ex: Sicredi)
    Dim urlBase As String
    Dim urlOAuth As String
    Dim retornoJson As String
    
    urlBase = "https://api-pix.sicredi.com.br/pix/v2"
    urlOAuth = "https://api-pix.sicredi.com.br/oauth/v2/token"
    
    ' Executa a configuracao customizada
    retornoJson = bridge.ConfigurarPixCustomizado( _
        urlBase, _
        urlOAuth, _
        "CLIENT_ID_EXEMPLO_12345", _
        "CLIENT_SECRET_EXEMPLO_67890", _
        "contato@lojistaexemplo.com.br", _
        "", _
        "" _
    )
    
    txtResultado.Text = "=== RETORNO DO ConfigurarPixCustomizado ===" & vbCrLf & _
                        retornoJson
    
    ' Identifica o STATUS no JSON retornado
    If InStr(retornoJson, ""STATUS":"OK"") > 0 Then
        MsgBox "Configuração validada com sucesso na DLL!" & vbCrLf & _
               "Status: OK", vbInformation, "Banco Customizado"
    Else
        MsgBox "Falha na validação da configuração:" & vbCrLf & _
               retornoJson, vbExclamation, "Banco Customizado - ERRO"
    End If
    
    Set bridge = Nothing
    Exit Sub
TrataErro:
    txtResultado.Text = "Erro: " & Err.Description
End Sub

Private Sub cmdEstornoPix_Click()
    On Error GoTo TrataErro
    Dim bridge As Object
    Set bridge = CreateObject("BridgeRTC.BridgeRTCService")
    bridge.ConfigurarAmbiente 2, "SP"
    bridge.ConfigurarPix "SIMULADOR", "", "", "", "", ""
    
    ' Simula o estorno de um Pix efetuado (ex: desistência ou erro de impressão)
    Dim e2eId As String, idDev As String
    e2eId = "E1234567820260922120000000001"
    idDev = "DEV" & Format(Now, "yyyymmddhhnnss")
    
    txtResultado.Text = bridge.DevolverPix(e2eId, idDev, 100#, "Cancelamento de compra pelo cliente no PDV")
    Set bridge = Nothing
    Exit Sub
TrataErro:
    txtResultado.Text = "Erro: " & Err.Description
End Sub

Private Sub cmdConciliacao_Click()
    On Error GoTo TrataErro
    Dim bridge As Object
    Set bridge = CreateObject("BridgeRTC.BridgeRTCService")
    bridge.ConfigurarAmbiente 2, "SP"
    bridge.ConfigurarPix "SIMULADOR", "", "", "", "", ""
    
    Dim txid As String
    txid = "CONC" & Format(Now, "yyyymmddhhnnss")
    bridge.CriarCobrancaPix txid, 100#, 3600, "Venda Demonstração Split"
    bridge.SimularPagamentoPix txid
    
    txtResultado.Text = bridge.ObterConciliacaoFinanceiraPix(txid)
    Set bridge = Nothing
    Exit Sub
TrataErro:
    txtResultado.Text = "Erro: " & Err.Description
End Sub

Private Sub cmdStatus_Click()
    On Error GoTo TrataErro
    Dim bridge As Object
    Set bridge = CreateObject("BridgeRTC.BridgeRTCService")
    bridge.ConfigurarAmbiente 2, "SP"
    txtResultado.Text = bridge.ConsultarStatusServico()
    Set bridge = Nothing
    Exit Sub
TrataErro:
    txtResultado.Text = "Erro: " & Err.Description
End Sub

Private Sub cmdGerarXml_Click()
    On Error GoTo TrataErro
    Dim bridge As Object
    Set bridge = CreateObject("BridgeRTC.BridgeRTCService")
    txtResultado.Text = bridge.GerarGrupoPagamentoXml("17", 100#, "E12345678202609220000000001", "00000000000191", "AUTH987654")
    Set bridge = Nothing
    Exit Sub
TrataErro:
    txtResultado.Text = "Erro: " & Err.Description
End Sub

Private Sub cmdVincular_Click()
    On Error GoTo TrataErro
    Dim bridge As Object
    Set bridge = CreateObject("BridgeRTC.BridgeRTCService")
    bridge.ConfigurarAmbiente 2, "SP"
    txtResultado.Text = bridge.VincularPagamentoDFe("35260100000000000191550010000000011000000018", "E12345678202609220000000001", "17", 100#, "00000000000191", "AUTH987654")
    Set bridge = Nothing
    Exit Sub
TrataErro:
    txtResultado.Text = "Erro: " & Err.Description
End Sub

Private Sub cmdAbrirBoletoPix_Click()
    Load frmBoletoPix
    frmBoletoPix.Show vbModal
    Unload frmBoletoPix
End Sub

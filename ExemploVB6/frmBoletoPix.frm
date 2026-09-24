VERSION 5.00
Begin VB.Form frmBoletoPix 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Boleto Hibrido com Pix (ActiveReports / CobV)"
   ClientHeight    =   8200
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8600
   LinkTopic       =   "frmBoletoPix"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8200
   ScaleWidth      =   8600
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraSimulador 
      Caption         =   "2. Simular Leitura e Pagamento pelo Aplicativo Bancario"
      Height          =   1600
      Left            =   240
      TabIndex        =   16
      Top             =   3480
      Width           =   8175
      Begin VB.CommandButton cmdPagarAtrasado 
         Caption         =   "Pagar c/ 5 Dias de Atraso (+ Multa e Juros)"
         Height          =   420
         Left            =   4920
         TabIndex        =   20
         Top             =   960
         Width           =   3015
      End
      Begin VB.CommandButton cmdPagarNoVencimento 
         Caption         =   "Pagar no Vencimento (Valor Normal)"
         Height          =   420
         Left            =   240
         TabIndex        =   18
         Top             =   960
         Width           =   2295
      End
      Begin VB.CommandButton cmdPagarAntecipado 
         Caption         =   "Pagar Antecipado (Aplica Desconto)"
         Height          =   420
         Left            =   2640
         TabIndex        =   19
         Top             =   960
         Width           =   2175
      End
      Begin VB.Label lblInfoSimulacao 
         Caption         =   "O banco do cliente consulta o QR Code em tempo real e decide sozinho se aplica desconto ou calcula multa e juros diarios:"
         Height          =   495
         Left            =   240
         TabIndex        =   17
         Top             =   360
         Width           =   7695
      End
   End
   Begin VB.Frame fraDadosBoleto 
      Caption         =   "1. Dados da Emissao do Boleto (ERP)"
      Height          =   3135
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   8175
      Begin VB.CommandButton cmdGerarBoletoPix 
         Caption         =   "Gerar Pix Copia e Cola para ActiveReports"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   240
         TabIndex        =   15
         Top             =   2520
         Width           =   7695
      End
      Begin VB.TextBox txtJurosMes 
         Height          =   315
         Left            =   6720
         TabIndex        =   14
         Text            =   "1,00"
         Top             =   1920
         Width           =   1215
      End
      Begin VB.TextBox txtMultaPerc 
         Height          =   315
         Left            =   4080
         TabIndex        =   12
         Text            =   "2,00"
         Top             =   1920
         Width           =   1215
      End
      Begin VB.TextBox txtDataDesconto 
         Height          =   315
         Left            =   1560
         TabIndex        =   10
         Text            =   "10/10/2026"
         Top             =   1920
         Width           =   1215
      End
      Begin VB.TextBox txtValorDesconto 
         Height          =   315
         Left            =   6720
         TabIndex        =   8
         Text            =   "10,00"
         Top             =   1320
         Width           =   1215
      End
      Begin VB.TextBox txtVencimento 
         Height          =   315
         Left            =   4080
         TabIndex        =   6
         Text            =   "15/10/2026"
         Top             =   1320
         Width           =   1215
      End
      Begin VB.TextBox txtValorOriginal 
         Height          =   315
         Left            =   1560
         TabIndex        =   4
         Text            =   "150,00"
         Top             =   1320
         Width           =   1215
      End
      Begin VB.TextBox txtNomeCliente 
         Height          =   315
         Left            =   4080
         TabIndex        =   3
         Text            =   "MERCADO EXEMPLO LTDA"
         Top             =   720
         Width           =   3855
      End
      Begin VB.TextBox txtNossoNumero 
         Height          =   315
         Left            =   1560
         TabIndex        =   2
         Text            =   "BOL20260924001"
         Top             =   720
         Width           =   2295
      End
      Begin VB.Label lblJuros 
         Caption         =   "Juros Mes (%):"
         Height          =   255
         Left            =   5520
         TabIndex        =   13
         Top             =   1950
         Width           =   1095
      End
      Begin VB.Label lblMulta 
         Caption         =   "Multa (%):"
         Height          =   255
         Left            =   3000
         TabIndex        =   11
         Top             =   1950
         Width           =   975
      End
      Begin VB.Label lblDataDesc 
         Caption         =   "Desc. ate:"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   1950
         Width           =   1215
      End
      Begin VB.Label lblValorDesc 
         Caption         =   "Desconto R$:"
         Height          =   255
         Left            =   5520
         TabIndex        =   7
         Top             =   1350
         Width           =   1095
      End
      Begin VB.Label lblVenc 
         Caption         =   "Vencimento:"
         Height          =   255
         Left            =   3000
         TabIndex        =   5
         Top             =   1350
         Width           =   975
      End
      Begin VB.Label lblValor 
         Caption         =   "Valor Original:"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   1350
         Width           =   1215
      End
   End
   Begin VB.TextBox txtResultado 
      Height          =   2655
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   21
      Top             =   5400
      Width           =   8175
   End
   Begin VB.Label lblRetorno 
      Caption         =   "3. Retorno da DLL / String Pix para alimentar o ActiveReports:"
      Height          =   255
      Left            =   240
      TabIndex        =   22
      Top             =   5160
      Width           =   5000
   End
End
Attribute VB_Name = "frmBoletoPix"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mBridge As Object
Private mUltimoCopiaECola As String

Private Sub Form_Load()
    On Error Resume Next
    Set mBridge = CreateObject("BridgeRTC.BridgeRTCService")

    ' =========================================================================
    ' SELECAO DO AMBIENTE DE TRABALHO (Escolha UMA das 3 opcoes abaixo):
    ' =========================================================================

    ' --- OPCAO 1: MODO SIMULADOR (Para desenvolvimento interno sem comunicacao bancaria) ---
    mBridge.ConfigurarAmbiente 2, "SP"
    mBridge.ConfigurarPix "SIMULADOR", "", "", "", "", ""

    ' --- OPCAO 2: MODO HOMOLOGACAO / SANDBOX (Testes reais contra a API de testes do Banco) ---
    ' Descomente as linhas abaixo e comente a OPCAO 1 quando for homologar:
    ' mBridge.ConfigurarAmbiente 2, "SP"  ' 2 = Homologacao
    ' mBridge.ConfigurarPix "ITAU", "SEU_CLIENT_ID_HM", "SEU_CLIENT_SECRET_HM", "sua_chave_pix_hm", "C:\Certificados\empresa_hm.pfx", "SenhaCertificado123"
    ' OU para banco customizado:
    ' mBridge.ConfigurarPixCustomizado "https://api-sandbox.banco.com.br/pix/v2", "https://oauth-sandbox.banco.com.br/token", "SEU_CLIENT_ID", "SEU_SECRET", "chave_pix", "C:\Cert\cert.pfx", "senha"

    ' --- OPCAO 3: MODO PRODUCAO (Boletos e Pix reais valendo dinheiro) ---
    ' Descomente apenas quando a aplicacao no banco for aprovada e estiver no ar:
    ' mBridge.ConfigurarAmbiente 1, "SP"  ' 1 = Producao
    ' mBridge.ConfigurarPix "ITAU", "CLIENT_ID_PRODUCAO", "CLIENT_SECRET_PRODUCAO", "sua_chave_pix_real", "C:\Certificados\empresa_prod.pfx", "SenhaCertificado123"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mBridge = Nothing
End Sub

Private Sub cmdGerarBoletoPix_Click()
    On Error GoTo TrataErro
    
    Dim dValor As Double
    Dim dDesconto As Double
    Dim dMulta As Double
    Dim dJuros As Double
    
    dValor = CDbl(Replace(txtValorOriginal.Text, ",", "."))
    dDesconto = CDbl(Replace(txtValorDesconto.Text, ",", "."))
    dMulta = CDbl(Replace(txtMultaPerc.Text, ",", "."))
    dJuros = CDbl(Replace(txtJurosMes.Text, ",", "."))
    
    Dim sResp As String
    sResp = mBridge.CriarBoletoHibridoPix( _
        txtNossoNumero.Text, _
        dValor, _
        txtVencimento.Text, _
        30, _
        dDesconto, _
        txtDataDesconto.Text, _
        dMulta, _
        dJuros, _
        "12345678000195", _
        txtNomeCliente.Text, _
        "Boleto ref. Venda " & txtNossoNumero.Text _
    )
    
    mUltimoCopiaECola = ExtrairValorJson(sResp, "pixCopiaECola")
    
    Dim sRelatorio As String
    sRelatorio = "=== BOLETO HIBRIDO COM PIX GERADO COM SUCESSO ===" & vbCrLf & _
                 "Nosso Numero: " & txtNossoNumero.Text & vbCrLf & _
                 "Valor Nominal: R$ " & FormatNumber(dValor, 2) & vbCrLf & _
                 "Vencimento...: " & txtVencimento.Text & vbCrLf & _
                 "Regra Desconto: R$ " & FormatNumber(dDesconto, 2) & " ate " & txtDataDesconto.Text & vbCrLf & _
                 "Regra Atraso..: Multa de " & txtMultaPerc.Text & "% + Juros de " & txtJurosMes.Text & "% ao mes" & vbCrLf & vbCrLf & _
                 "----- STRING PARA ALIMENTAR O ACTIVE REPORTS -----" & vbCrLf & _
                 "Exemplo ActiveReports:" & vbCrLf & _
                 "Me.ImgQrCode.Picture = SuaFuncaoDesenharQrCode(sPixCopiaECola)" & vbCrLf & vbCrLf & _
                 "Codigo Copia e Cola (Payload Oficial):" & vbCrLf & _
                 mUltimoCopiaECola & vbCrLf & vbCrLf & _
                 "JSON Completo da DLL:" & vbCrLf & sResp
                 
    txtResultado.Text = sRelatorio
    MsgBox "Pix Copia e Cola gerado com sucesso! Passe esta string para o ActiveReports.", vbInformation, "Boleto Hibrido"
    Exit Sub
TrataErro:
    txtResultado.Text = "Erro ao gerar Boleto Pix: " & Err.Description
End Sub

Private Sub cmdPagarAntecipado_Click()
    On Error GoTo TrataErro
    Dim sDataPagto As String
    sDataPagto = txtDataDesconto.Text
    
    Dim sResp As String
    sResp = mBridge.SimularPagamentoBoletoPix(txtNossoNumero.Text, sDataPagto)
    
    txtResultado.Text = "=== SIMULACAO 1: CLIENTE PAGOU ANTECIPADO COM DESCONTO ===" & vbCrLf & sResp
    MsgBox "O banco aplicou o desconto automaticamente! Veja o JSON retornado.", vbInformation, "Pagamento com Desconto"
    Exit Sub
TrataErro:
    txtResultado.Text = "Erro: " & Err.Description
End Sub

Private Sub cmdPagarNoVencimento_Click()
    On Error GoTo TrataErro
    Dim sResp As String
    sResp = mBridge.SimularPagamentoBoletoPix(txtNossoNumero.Text, txtVencimento.Text)
    
    txtResultado.Text = "=== SIMULACAO 2: CLIENTE PAGOU NO VENCIMENTO ===" & vbCrLf & sResp
    MsgBox "Pagamento no vencimento: cobrado o valor original exato.", vbInformation, "Pagamento Normal"
    Exit Sub
TrataErro:
    txtResultado.Text = "Erro: " & Err.Description
End Sub

Private Sub cmdPagarAtrasado_Click()
    On Error GoTo TrataErro
    Dim dtVenc As Date
    dtVenc = CDate(txtVencimento.Text)
    Dim dtAtraso As Date
    dtAtraso = DateAdd("d", 5, dtVenc)
    
    Dim sResp As String
    sResp = mBridge.SimularPagamentoBoletoPix(txtNossoNumero.Text, Format(dtAtraso, "dd/mm/yyyy"))
    
    txtResultado.Text = "=== SIMULACAO 3: CLIENTE PAGOU COM 5 DIAS DE ATRASO ===" & vbCrLf & sResp
    MsgBox "O banco calculou a multa e juros pro-rata automaticamente! Veja os valores no JSON.", vbInformation, "Pagamento em Atraso"
    Exit Sub
TrataErro:
    txtResultado.Text = "Erro: " & Err.Description
End Sub

Private Function ExtrairValorJson(sJson As String, sCampo As String) As String
    Dim posIni As Long, posFim As Long
    Dim busca As String
    busca = Chr(34) & sCampo & Chr(34) & ":" & Chr(34)
    posIni = InStr(sJson, busca)
    If posIni > 0 Then
        posIni = posIni + Len(busca)
        posFim = InStr(posIni, sJson, Chr(34))
        If posFim > posIni Then
            ExtrairValorJson = Mid(sJson, posIni, posFim - posIni)
            Exit Function
        End If
    End If
    ExtrairValorJson = ""
End Function

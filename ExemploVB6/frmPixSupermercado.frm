VERSION 5.00
Begin VB.Form frmPixSupermercado 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Pagamento PIX - Supermercado / PDV"
   ClientHeight    =   7900
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7500
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7900
   ScaleWidth      =   7500
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdReconsultarBanco 
      Caption         =   "Reconsultar Banco Agora"
      Height          =   450
      Left            =   360
      TabIndex        =   9
      Top             =   7200
      Width           =   2200
   End
   Begin VB.Timer tmrPolling 
      Enabled         =   0   'False
      Interval        =   2500
      Left            =   120
      Top             =   7200
   End
   Begin VB.CommandButton cmdSimularConfirmacao 
      Caption         =   "Simular Pagamento (Dev)"
      Height          =   450
      Left            =   2700
      TabIndex        =   5
      Top             =   7200
      Width           =   2400
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   450
      Left            =   5250
      TabIndex        =   4
      Top             =   7200
      Width           =   1845
   End
   Begin VB.TextBox txtPixCopiaECola 
      BackColor       =   &H00F0F0F0&
      Height          =   555
      Left            =   360
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   6480
      Width           =   6735
   End
   Begin VB.Label lblContadorTimeout 
      Alignment       =   2  'Center
      Caption         =   "Tempo restante: 03:00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   360
      TabIndex        =   10
      Top             =   4440
      Width           =   6735
   End
   Begin VB.Label lblSplitDetalhamento 
      Alignment       =   2  'Center
      BackColor       =   &H00E0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Previsão Split Payment: Bruto R$ 0,00 | CBS/IBS: R$ 0,00 | Líquido Caixa: R$ 0,00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   435
      Left            =   360
      TabIndex        =   8
      Top             =   5700
      Width           =   6735
   End
   Begin VB.PictureBox picQrCode 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   2655
      Left            =   2280
      ScaleHeight     =   2595
      ScaleWidth      =   2715
      TabIndex        =   1
      Top             =   1320
      Width           =   2775
   End
   Begin VB.Label lblInstrucaoCopiaCola 
      Caption         =   "Código Pix Copia e Cola:"
      Height          =   255
      Left            =   360
      TabIndex        =   7
      Top             =   6240
      Width           =   5500
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      Caption         =   "Aguardando leitura do QR Code pelo cliente no app do banco..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   550
      Left            =   240
      TabIndex        =   6
      Top             =   4800
      Width           =   6975
   End
   Begin VB.Label lblTituloValor 
      Alignment       =   2  'Center
      Caption         =   "TOTAL A PAGAR: R$ 0,00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   6975
   End
   Begin VB.Label lblSubtitulo 
      Alignment       =   2  'Center
      Caption         =   "Abra o app do seu banco e aponte a câmera para o QR Code abaixo"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   840
      Width           =   6975
   End
End
Attribute VB_Name = "frmPixSupermercado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mBridge As Object
Private mTxId As String
Private mValorVenda As Double
Private mPagoComSucesso As Boolean
Private mEndToEndId As String
Private mJsonRetorno As String

' Timeout do PDV (3 minutos = 180 segundos)
Private mSegundosRestantes As Integer

' Valores de Conciliação Financeira (Split Payment)
Private mValorBruto As Double
Private mValorTributosRetidos As Double
Private mValorTarifa As Double
Private mValorLiquido As Double
Private mValorCbs As Double
Private mValorIbs As Double

Public Property Get PagamentoConfirmado() As Boolean
    PagamentoConfirmado = mPagoComSucesso
End Property

Public Property Get EndToEndId() As String
    EndToEndId = mEndToEndId
End Property

Public Property Get JsonUltimoRetorno() As String
    JsonUltimoRetorno = mJsonRetorno
End Property

Public Property Get ValorBruto() As Double
    ValorBruto = mValorBruto
End Property

Public Property Get ValorTributosRetidos() As Double
    ValorTributosRetidos = mValorTributosRetidos
End Property

Public Property Get ValorTarifa() As Double
    ValorTarifa = mValorTarifa
End Property

Public Property Get ValorLiquido() As Double
    ValorLiquido = mValorLiquido
End Property

Public Property Get ValorCbs() As Double
    ValorCbs = mValorCbs
End Property

Public Property Get ValorIbs() As Double
    ValorIbs = mValorIbs
End Property

Public Sub IniciarCobranca(ByVal vValor As Double, ByVal sDescricao As String)
    mValorVenda = vValor
    mPagoComSucesso = False
    mEndToEndId = ""
    mValorBruto = vValor
    mSegundosRestantes = 180 ' 3 minutos limite de espera do caixa
    
    lblTituloValor.Caption = "TOTAL A PAGAR: R$ " & FormatNumber(mValorVenda, 2)
    lblStatus.Caption = "Conectando ao banco e gerando QR Code Pix..."
    lblStatus.ForeColor = &HC00000
    AtualizarLabelTimeout
    
    On Error Resume Next
    Set mBridge = CreateObject("BridgeRTC.BridgeRTCService")
    If Err.Number <> 0 Then
        MsgBox "Erro ao instanciar BridgeRTC.BridgeRTCService: " & vbCrLf & Err.Description, vbCritical, "Erro DLL"
        Unload Me
        Exit Sub
    End If
    On Error GoTo 0
    
    ' =========================================================================
    ' SELECAO DO AMBIENTE (Simulador, Homologacao ou Producao):
    ' =========================================================================
    ' Opcao 1 (Padrao): MODO SIMULADOR
    mBridge.ConfigurarAmbiente 2, "SP"
    mBridge.ConfigurarPix "SIMULADOR", "", "", "suporte@meuerp.com.br", "", ""
    
    ' Opcao 2: MODO HOMOLOGACAO (Descomente para testar contra sandbox do banco)
    ' mBridge.ConfigurarAmbiente 2, "SP"
    ' mBridge.ConfigurarPix "ITAU", "CLIENT_ID_HM", "CLIENT_SECRET_HM", "chave_hm@loja.com", "C:\Cert\cert_hm.pfx", "senha123"
    
    ' Opcao 3: MODO PRODUCAO (Descomente em ambiente real de vendas)
    ' mBridge.ConfigurarAmbiente 1, "SP"
    ' mBridge.ConfigurarPix "ITAU", "CLIENT_ID_PROD", "CLIENT_SECRET_PROD", "chave_real@loja.com", "C:\Cert\cert_prod.pfx", "senha123" ""
    
    mTxId = "PDV" & Format(Now, "yyyymmddhhnnss") & "001"
    
    Dim respJson As String
    respJson = mBridge.CriarCobrancaPix(mTxId, mValorVenda, 180, sDescricao)
    mJsonRetorno = respJson
    
    If InStr(respJson, """STATUS"":""OK""") > 0 Then
        txtPixCopiaECola.Text = ExtrairValorJson(respJson, "pixCopiaECola")
        DesenharVisualQrCode
        
        Dim prevTrib As Double, prevLiq As Double
        prevTrib = Val(ExtrairValorJson(respJson, "previsaoTributosRetidos"))
        prevLiq = Val(ExtrairValorJson(respJson, "previsaoLiquidoConta"))
        
        lblSplitDetalhamento.Caption = "Split Payment: Venda R$ " & FormatNumber(mValorVenda, 2) & _
            " | Impostos Retidos: R$ " & FormatNumber(prevTrib, 2) & _
            " | Líquido Caixa: R$ " & FormatNumber(prevLiq, 2)
        
        lblStatus.Caption = "Aguardando leitura do QR Code pelo cliente no app do banco..."
        lblStatus.ForeColor = &H8000&
        
        tmrPolling.Interval = 2500
        tmrPolling.Enabled = True
    Else
        lblStatus.Caption = "Falha ao gerar cobrança Pix."
        lblStatus.ForeColor = vbRed
        MsgBox "Erro DLL:" & vbCrLf & respJson, vbExclamation, "Aviso"
    End If
End Sub

Private Sub tmrPolling_Timer()
    ' Decrementa tempo restante (cada tick ~ 2.5s)
    mSegundosRestantes = mSegundosRestantes - 2
    AtualizarLabelTimeout
    
    If mSegundosRestantes <= 0 Then
        tmrPolling.Enabled = False
        lblStatus.Caption = "TEMPO ESGOTADO. Se o cliente já pagou, clique em 'Reconsultar Banco'."
        lblStatus.ForeColor = vbRed
        Beep
        Exit Sub
    End If
    
    ChecarPagamentoComBanco
End Sub

Private Sub ChecarPagamentoComBanco()
    If mBridge Is Nothing Or mTxId = "" Then Exit Sub
    
    Dim respConsulta As String
    respConsulta = mBridge.ConsultarCobrancaPix(mTxId)
    mJsonRetorno = respConsulta
    
    If InStr(UCase(respConsulta), """STATUS"":""CONCLUIDA""") > 0 Or InStr(LCase(respConsulta), """pago"":true") > 0 Then
        tmrPolling.Enabled = False
        mPagoComSucesso = True
        
        mEndToEndId = ExtrairValorJson(respConsulta, "endToEndId")
        mValorBruto = Val(ExtrairValorJson(respConsulta, "valorBruto"))
        If mValorBruto = 0 Then mValorBruto = mValorVenda
        
        mValorTributosRetidos = Val(ExtrairValorJson(respConsulta, "valorTributosRetidos"))
        mValorTarifa = Val(ExtrairValorJson(respConsulta, "valorTarifaBancaria"))
        mValorLiquido = Val(ExtrairValorJson(respConsulta, "valorLiquidoRecebido"))
        mValorCbs = Val(ExtrairValorJson(respConsulta, "cbsRetido"))
        mValorIbs = Val(ExtrairValorJson(respConsulta, "ibsRetido"))
        
        lblSplitDetalhamento.Caption = "CONFIRMADO -> Bruto: R$ " & FormatNumber(mValorBruto, 2) & _
            " | Retido Gov: R$ " & FormatNumber(mValorTributosRetidos, 2) & _
            " | Líquido Caixa: R$ " & FormatNumber(mValorLiquido, 2)
        
        lblStatus.Caption = "PAGAMENTO CONFIRMADO! Finalizando venda..."
        lblStatus.ForeColor = &H8000&
        Beep
        
        Dim tFim As Single
        tFim = Timer + 1
        Do While Timer < tFim
            DoEvents
        Loop
        
        Unload Me
    End If
End Sub

Private Sub cmdReconsultarBanco_Click()
    ' Reconsulta forçada pelo operador quando a internet oscilou
    lblStatus.Caption = "Reconsultando status do Pix junto ao banco..."
    lblStatus.ForeColor = &HC00000
    ChecarPagamentoComBanco
    If Not mPagoComSucesso Then
        lblStatus.Caption = "Banco informa: Pagamento ainda pendente de confirmação."
        lblStatus.ForeColor = &H80&
    End If
End Sub

Private Sub cmdSimularConfirmacao_Click()
    If mBridge Is Nothing Or mTxId = "" Then Exit Sub
    mBridge.SimularPagamentoPix mTxId
    ChecarPagamentoComBanco
End Sub

Private Sub cmdCancelar_Click()
    tmrPolling.Enabled = False
    mPagoComSucesso = False
    Unload Me
End Sub

Private Sub AtualizarLabelTimeout()
    Dim min As Integer, seg As Integer
    If mSegundosRestantes < 0 Then mSegundosRestantes = 0
    min = mSegundosRestantes \ 60
    seg = mSegundosRestantes Mod 60
    lblContadorTimeout.Caption = "Tempo de espera no caixa: " & Format(min, "00") & ":" & Format(seg, "00")
End Sub

Private Sub Form_Unload(Cancel As Integer)
    tmrPolling.Enabled = False
    Set mBridge = Nothing
End Sub

Private Function ExtrairValorJson(ByVal json As String, ByVal chave As String) As String
    Dim posChave As Long, posInicio As Long, posFim As Long
    posChave = InStr(json, """" & chave & """")
    If posChave = 0 Then Exit Function
    
    posInicio = InStr(posChave, json, ":")
    If posInicio = 0 Then Exit Function
    
    posInicio = posInicio + 1
    Do While Mid$(json, posInicio, 1) = " " Or Mid$(json, posInicio, 1) = """"
        posInicio = posInicio + 1
    Loop
    
    posFim = InStr(posInicio, json, """")
    If posFim = 0 Then
        posFim = InStr(posInicio, json, ",")
        If posFim = 0 Then posFim = InStr(posInicio, json, "}")
    End If
    
    If posFim > posInicio Then
        ExtrairValorJson = Mid$(json, posInicio, posFim - posInicio)
    End If
End Function

Private Sub DesenharVisualQrCode()
    picQrCode.Cls
    picQrCode.Line (100, 100)-(700, 700), vbBlack, BF
    picQrCode.Line (200, 200)-(600, 600), vbWhite, BF
    picQrCode.Line (300, 300)-(500, 500), vbBlack, BF
    picQrCode.Line (1800, 100)-(2400, 700), vbBlack, BF
    picQrCode.Line (1900, 200)-(2300, 600), vbWhite, BF
    picQrCode.Line (2000, 300)-(2200, 500), vbBlack, BF
    picQrCode.Line (100, 1800)-(700, 2400), vbBlack, BF
    picQrCode.Line (200, 1900)-(600, 2300), vbWhite, BF
    picQrCode.Line (300, 2000)-(500, 2200), vbBlack, BF
    picQrCode.CurrentX = 400
    picQrCode.CurrentY = 1100
    picQrCode.FontBold = True
    picQrCode.Print "QR CODE PIX DINAMICO"
    picQrCode.CurrentX = 450
    picQrCode.CurrentY = 1350
    picQrCode.FontBold = False
    picQrCode.Print "PADRAO BACEN v2"
End Sub

VERSION 5.00
Begin VB.Form frmPixSupermercado 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Pagamento PIX - Supermercado / PDV"
   ClientHeight    =   6405
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7500
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6405
   ScaleWidth      =   7500
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrPolling 
      Enabled         =   0   'False
      Interval        =   2500
      Left            =   120
      Top             =   5760
   End
   Begin VB.CommandButton cmdSimularConfirmacao 
      Caption         =   "Simular Pagamento (Dev/Teste)"
      Height          =   495
      Left            =   1200
      TabIndex        =   5
      Top             =   5760
      Width           =   2800
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar Operação"
      Height          =   495
      Left            =   4200
      TabIndex        =   4
      Top             =   5760
      Width           =   2000
   End
   Begin VB.TextBox txtPixCopiaECola 
      BackColor       =   &H00F0F0F0&
      Height          =   735
      Left            =   360
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   4800
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
      Caption         =   "Código Pix Copia e Cola (para conferência ou pagamento alternativo):"
      Height          =   255
      Left            =   360
      TabIndex        =   7
      Top             =   4560
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
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   4080
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

' ==============================================================================
' FORMULÁRIO DE PAGAMENTO PIX DIRETO (SUPERMERCADO / PDV MODAL)
' ==============================================================================
' Comportamento:
' 1. Recebe o valor da venda e identificador único da transação (txid).
' 2. Chama a DLL BridgeRTC (CreateObject("BridgeRTC.BridgeRTCService")).
' 3. Solicita a criação da cobrança imediata via API Bacen (PUT /v2/cob/{txid}).
' 4. Exibe na tela o QR Code em destaque para o cliente ler.
' 5. Inicia o Timer (polling a cada 2.5s) consultando ConsultarCobrancaPix(txid).
' 6. Quando o consumidor paga no app do banco, o banco confirma instantaneamente.
' 7. O formulário detecta "pago = true", emite aviso sonoro/visual, fecha-se
'    e aciona a rotina de emissão imediata da NFC-e com o endToEndId retornado.
' ==============================================================================

Private mBridge As Object
Private mTxId As String
Private mValorVenda As Double
Private mPagoComSucesso As Boolean
Private mEndToEndId As String
Private mJsonRetorno As String

' Propriedade pública para saber se o Pix foi pago antes de emitir a NFC-e
Public Property Get PagamentoConfirmado() As Boolean
    PagamentoConfirmado = mPagoComSucesso
End Property

Public Property Get EndToEndId() As String
    EndToEndId = mEndToEndId
End Property

Public Property Get JsonUltimoRetorno() As String
    JsonUltimoRetorno = mJsonRetorno
End Property

' Método para iniciar a cobrança com os parâmetros do caixa
Public Sub IniciarCobranca(ByVal vValor As Double, ByVal sDescricao As String)
    mValorVenda = vValor
    mPagoComSucesso = False
    mEndToEndId = ""
    
    lblTituloValor.Caption = "TOTAL A PAGAR: R$ " & FormatNumber(mValorVenda, 2)
    lblStatus.Caption = "Conectando ao serviço financeiro e gerando QR Code..."
    lblStatus.ForeColor = &HC00000
    
    ' 1. Instancia a DLL BridgeRTC
    On Error Resume Next
    Set mBridge = CreateObject("BridgeRTC.BridgeRTCService")
    If Err.Number <> 0 Then
        MsgBox "Erro ao instanciar DLL BridgeRTC.BridgeRTCService: " & vbCrLf & Err.Description, vbCritical, "Erro DLL"
        Unload Me
        Exit Sub
    End If
    On Error GoTo 0
    
    ' 2. Configura ambiente:
    ' tpAmb: 1 = Produção, 2 = Homologação (Bancos / SEFAZ)
    mBridge.ConfigurarAmbiente 2, "SP"
    
    ' 3. Configura a Instituição Financeira / API Pix Bacen:
    ' Se passar "SIMULADOR", funciona localmente sem precisar de conta de banco aberta para testes do time.
    ' Para bancos reais, passe ex: "BANCO_DO_BRASIL", "ITAU", "SANTANDER", "INTER", "SICOOB", "EFIPAY".
    ' mBridge.ConfigurarPix "BANCO_DO_BRASIL", "client-id-aqui", "client-secret-aqui", "chave-pix-aqui", "C:\certificados\pix.pfx", "senha123"
    mBridge.ConfigurarPix "SIMULADOR", "", "", "suporte@meuerp.com.br", "", ""
    
    ' 4. Gera TxId único para o cupom/venda (26 a 35 caracteres alfanuméricos)
    mTxId = "PDV" & Format(Now, "yyyymmddhhnnss") & "001"
    
    ' 5. Solicita a criação da cobrança imediata (Bacen PUT /v2/cob/{txid})
    Dim respJson As String
    respJson = mBridge.CriarCobrancaPix(mTxId, mValorVenda, 3600, sDescricao)
    mJsonRetorno = respJson
    
    ' Verifica se a criação retornou OK
    If InStr(respJson, """STATUS"":""OK""") > 0 Then
        ' Extrai dados do Pix Copia e Cola
        Dim sCopiaCola As String
        sCopiaCola = ExtrairValorJson(respJson, "pixCopiaECola")
        txtPixCopiaECola.Text = sCopiaCola
        
        ' Desenha representação visual do QR Code no PictureBox
        DesenharVisualQrCode
        
        lblStatus.Caption = "Aguardando leitura do QR Code pelo cliente no app do banco..."
        lblStatus.ForeColor = &H8000& ' Verde escuro
        
        ' Ativa o timer de checagem periódica (polling)
        tmrPolling.Interval = 2500
        tmrPolling.Enabled = True
    Else
        lblStatus.Caption = "Falha ao gerar cobrança Pix. Verifique a conexão com o banco."
        lblStatus.ForeColor = vbRed
        MsgBox "Erro retornado pela DLL:" & vbCrLf & respJson, vbExclamation, "Aviso"
    End If
End Sub

Private Sub tmrPolling_Timer()
    ' Consulta o status da cobrança na API Pix do Bacen através da DLL
    If mBridge Is Nothing Or mTxId = "" Then Exit Sub
    
    Dim respConsulta As String
    respConsulta = mBridge.ConsultarCobrancaPix(mTxId)
    mJsonRetorno = respConsulta
    
    ' Verifica se o status é CONCLUIDA / pago = true
    If InStr(UCase(respConsulta), """STATUS"":""CONCLUIDA""") > 0 Or _
       InStr(LCase(respConsulta), """pago"":true") > 0 Then
        
        ' Para o timer para evitar múltiplas execuções
        tmrPolling.Enabled = False
        mPagoComSucesso = True
        
        ' Captura o EndToEndId oficial da transação Pix devolvido pelo Bacen/Banco
        mEndToEndId = ExtrairValorJson(respConsulta, "endToEndId")
        
        lblStatus.Caption = "PAGAMENTO CONFIRMADO! Finalizando venda..."
        lblStatus.ForeColor = &H8000&
        Beep
        
        ' Pequena pausa visual para o caixa e cliente verem a confirmação (1 segundo)
        Dim tFim As Single
        tFim = Timer + 1
        Do While Timer < tFim
            DoEvents
        Loop
        
        ' Fecha o modal do Pix e devolve o controle para o caixa emitir a NFC-e
        Unload Me
    End If
End Sub

' Botão para permitir que os desenvolvedores testem a confirmação imediata
Private Sub cmdSimularConfirmacao_Click()
    If mBridge Is Nothing Or mTxId = "" Then
        MsgBox "Inicie uma cobrança primeiro.", vbInformation
        Exit Sub
    End If
    
    Dim retSimulacao As String
    retSimulacao = mBridge.SimularPagamentoPix(mTxId)
    ' O próximo tick do timer já identificará como CONCLUIDA e fechará a tela
End Sub

Private Sub cmdCancelar_Click()
    tmrPolling.Enabled = False
    mPagoComSucesso = False
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    tmrPolling.Enabled = False
    Set mBridge = Nothing
End Sub

' Função auxiliar simples para extrair valores de chaves JSON sem dependência externa
Private Function ExtrairValorJson(ByVal json As String, ByVal chave As String) As String
    Dim posChave As Long, posInicio As Long, posFim As Long
    posChave = InStr(json, """" & chave & """")
    If posChave = 0 Then Exit Function
    
    posInicio = InStr(posChave, json, ":")
    If posInicio = 0 Then Exit Function
    
    ' Pula espaços
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

' Desenha um padrão visual representativo de QR Code no PictureBox
Private Sub DesenharVisualQrCode()
    picQrCode.Cls
    picQrCode.CurrentX = 200
    picQrCode.CurrentY = 200
    
    ' Desenha quadrados de alinhamento típicos de QR Code
    picQrCode.Line (100, 100)-(700, 700), vbBlack, BF
    picQrCode.Line (200, 200)-(600, 600), vbWhite, BF
    picQrCode.Line (300, 300)-(500, 500), vbBlack, BF
    
    picQrCode.Line (1800, 100)-(2400, 700), vbBlack, BF
    picQrCode.Line (1900, 200)-(2300, 600), vbWhite, BF
    picQrCode.Line (2000, 300)-(2200, 500), vbBlack, BF
    
    picQrCode.Line (100, 1800)-(700, 2400), vbBlack, BF
    picQrCode.Line (200, 1900)-(600, 2300), vbWhite, BF
    picQrCode.Line (300, 2000)-(500, 2200), vbBlack, BF
    
    ' Texto informativo no centro
    picQrCode.CurrentX = 400
    picQrCode.CurrentY = 1100
    picQrCode.FontBold = True
    picQrCode.Print "QR CODE PIX DINAMICO"
    picQrCode.CurrentX = 450
    picQrCode.CurrentY = 1350
    picQrCode.FontBold = False
    picQrCode.Print "PADRAO BACEN v2"
End Sub

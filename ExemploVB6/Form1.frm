VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Teste Split Payment Bridge (VB6)"
   ClientHeight    =   4200
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6800
   LinkTopic       =   "Form1"
   ScaleHeight     =   4200
   ScaleWidth      =   6800
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdVincular 
      Caption         =   "1. Testar Vinculacao (Evento 110300)"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   3000
   End
   Begin VB.CommandButton cmdGerarXml 
      Caption         =   "2. Gerar Grupo XML Pagamento"
      Height          =   495
      Left            =   3480
      TabIndex        =   1
      Top             =   240
      Width           =   3000
   End
   Begin VB.TextBox txtResultado 
      Height          =   3135
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   840
      Width           =   6255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdVincular_Click()
    On Error GoTo TrataErro
    
    Dim oBridge As Object
    Dim sJson As String
    
    txtResultado.Text = "Instanciando SplitPaymentBridge..."
    Set oBridge = CreateObject("SplitPaymentBridge.SplitPaymentService")
    
    ' Configura ambiente (2 = Homologacao, 1 = Producao)
    oBridge.ConfigurarAmbiente 2, "SP"
    
    ' Realiza vinculacao da nota ao pagamento
    sJson = oBridge.VincularPagamentoDFe( _
        "35260900000000000000550010000000011000000010", _
        "PIX-E2E-ID1234567890", _
        "17", _
        150.75, _
        "00000000000191", _
        "AUT987654" _
    )
    
    txtResultado.Text = sJson
    Set oBridge = Nothing
    Exit Sub

TrataErro:
    txtResultado.Text = "Erro: " & Err.Description
    If Not oBridge Is Nothing Then Set oBridge = Nothing
End Sub

Private Sub cmdGerarXml_Click()
    On Error GoTo TrataErro
    
    Dim oBridge As Object
    Dim sJson As String
    
    Set oBridge = CreateObject("SplitPaymentBridge.SplitPaymentService")
    
    ' Gera o bloco XML <pag> para inclusao imediata na emissao
    sJson = oBridge.GerarGrupoPagamentoXml("24", 99.5, "TEF-NSU-001234", "00000000000191", "AUTH5544")
    
    txtResultado.Text = sJson
    Set oBridge = Nothing
    Exit Sub

TrataErro:
    txtResultado.Text = "Erro: " & Err.Description
    If Not oBridge Is Nothing Then Set oBridge = Nothing
End Sub

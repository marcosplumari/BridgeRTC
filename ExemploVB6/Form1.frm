VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Teste BridgeRTC - Split Payment e Pix Bacen"
   ClientHeight    =   6200
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7800
   LinkTopic       =   "Form1"
   ScaleHeight     =   6200
   ScaleWidth      =   7800
   StartUpPosition =   2  'CenterScreen
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
      Height          =   550
      Left            =   360
      TabIndex        =   5
      Top             =   1680
      Width           =   7095
   End
   Begin VB.CommandButton cmdStatus 
      Caption         =   "1. Consultar Status Servico"
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   3400
   End
   Begin VB.CommandButton cmdGerarXml 
      Caption         =   "2. Gerar Grupo Pagamento XML"
      Height          =   495
      Left            =   3960
      TabIndex        =   1
      Top             =   240
      Width           =   3400
   End
   Begin VB.CommandButton cmdVincular 
      Caption         =   "3. Testar Vinculacao SEFAZ (Evento 110300)"
      Height          =   495
      Left            =   360
      TabIndex        =   2
      Top             =   960
      Width           =   7095
   End
   Begin VB.TextBox txtResultado 
      Height          =   3300
      Left            =   360
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   3
      Top             =   2400
      Width           =   7095
   End
   Begin VB.Label lblStatus 
      Caption         =   "Retorno JSON da DLL:"
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   2160
      Width           =   2000
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAbrirPixSupermercado_Click()
    ' Demonstração do fluxo de supermercado:
    ' 1. Abre modal com o QR Code
    ' 2. O cliente paga no celular
    ' 3. A tela fecha sozinha e emite a NFC-e vinculada
    Load frmPixSupermercado
    frmPixSupermercado.IniciarCobranca 48.5, "Compra Caixa 01 - Supermercado"
    frmPixSupermercado.Show vbModal
    
    If frmPixSupermercado.PagamentoConfirmado Then
        txtResultado.Text = "=== PAGAMENTO PIX CONFIRMADO NO CAIXA ===" & vbCrLf & _
                            "EndToEndId: " & frmPixSupermercado.EndToEndId & vbCrLf & _
                            "Emitindo NFC-e com meio de pagamento 17 (Pix)..." & vbCrLf & vbCrLf & _
                            "Retorno completo da API Bacen:" & vbCrLf & _
                            frmPixSupermercado.JsonUltimoRetorno
        MsgBox "Venda concluída com sucesso! NFC-e emitida.", vbInformation, "Caixa Liberado"
    Else
        txtResultado.Text = "Operação Pix cancelada ou expirada pelo operador."
    End If
    
    Unload frmPixSupermercado
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
    txtResultado.Text = bridge.GerarGrupoPagamentoXml("17", 48.5, "E12345678202609220000000001", "00000000000191", "AUTH987654")
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
    txtResultado.Text = bridge.VincularPagamentoDFe("35260100000000000191550010000000011000000018", "E12345678202609220000000001", "17", 48.5, "00000000000191", "AUTH987654")
    Set bridge = Nothing
    Exit Sub
TrataErro:
    txtResultado.Text = "Erro: " & Err.Description
End Sub

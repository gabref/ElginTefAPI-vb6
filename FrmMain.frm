VERSION 5.00
Begin VB.Form FrmMain 
   Caption         =   "Menu"
   ClientHeight    =   4290
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   6345
   LinkTopic       =   "Form1"
   ScaleHeight     =   4290
   ScaleWidth      =   6345
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnCarregar 
      Caption         =   "Carregar Funções"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   840
      TabIndex        =   1
      Top             =   2760
      Width           =   4455
   End
   Begin VB.ListBox lstMenu 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1410
      ItemData        =   "FrmMain.frx":0000
      Left            =   840
      List            =   "FrmMain.frx":0002
      TabIndex        =   0
      Top             =   840
      Width           =   4455
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim buttonClicked As Boolean

Private Sub btnCarregar_Click()
    Dim selected As Integer
    selected = lstMenu.ListIndex
    If selected = Defines.PAGINA_PAGAMENTOS Then
        FrmPagamento.Show
    ElseIf selected = Defines.PAGINA_COLETA_PINPAD Then
        FrmColetaPinPad.Show
    Else
        MsgBox "Selecione uma opção válida"
    End If
End Sub

Private Function Iniciar() As String
    Dim resultado As String
    Dim payload As JsonBag
    Set payload = New JsonBag
    
    ' add examples
    
    resultado = IniciarOperacaoTEF(Stringify(payload))
    
    ' logs
    Set payload = Nothing
    
    Iniciar = resultado
End Function

Private Function Vender(ByVal cartao As Integer, ByVal sequencial As String, ByVal operacao As Integer) As String
    Dim resultado As String
    Dim payload As JsonBag
    Set payload = New JsonBag
    
    ' logs
    
    payload.Item("sequencial") = sequencial
    
    ' verificar valorTotal
    If operacao = 1 Then
        resultado = RealizarPagamentoTEF(CLng(cartao), Stringify(payload), True)
    Else
        resultado = RealizarPixTEF(Stringify(payload), True)
    End If
    
    ' logs
    
    Set payload = Nothing
    
    Vender = resultado
End Function

Private Sub Form_Load()
    lstMenu.AddItem ("Operações TEF e Adm")
    lstMenu.AddItem ("Operações de Coleta PinPad")
    lstMenu.ListIndex = 0
End Sub

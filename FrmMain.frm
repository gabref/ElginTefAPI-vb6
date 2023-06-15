VERSION 5.00
Begin VB.Form FrmMain 
   Caption         =   "Menu"
   ClientHeight    =   6750
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   11055
   LinkTopic       =   "Form1"
   ScaleHeight     =   6750
   ScaleWidth      =   11055
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnCarregar 
      Caption         =   "Carregar Funções"
      Height          =   615
      Left            =   1080
      TabIndex        =   4
      Top             =   5400
      Width           =   3975
   End
   Begin VB.ListBox lstMenu 
      Height          =   1035
      ItemData        =   "FrmMain.frx":0000
      Left            =   840
      List            =   "FrmMain.frx":0002
      TabIndex        =   3
      Top             =   3960
      Width           =   4215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   1095
      Left            =   840
      TabIndex        =   2
      Top             =   1920
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   1095
      Left            =   840
      TabIndex        =   1
      Top             =   480
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   2295
      Left            =   4440
      TabIndex        =   0
      Top             =   960
      Width           =   3975
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
    ElseIf selected = Defines.PAGINA_ADM Then
    Else
    End If
End Sub

Private Sub Command1_Click()
    Dim test As String
    Dim testjb As JsonBag
    
    Dim start As String
    Dim retorno As String
    Dim sequencial As String
    Dim resp As String
    
    start = StrPtrToString(Iniciar)
    MsgBox start
    retorno = GetRetorno(start)
    If retorno = "1" Then
        sequencial = incrementarSequencial(GetSequencial(start))
        
        MsgBox "vamos esperar um pouco o click do 2"
        
        Do While Not buttonClicked
            DoEvents
        Loop
        buttonClicked = False
        
        resp = StrPtrToString(Vender(0, sequencial, 1))
        
        MsgBox "Resp vender: " & resp
        
    End If
    
End Sub

Private Sub Command2_Click()
    buttonClicked = True
End Sub

Private Sub Command3_Click()
    Debug.Print lstMenu.Visible
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
    lstMenu.AddItem ("Operações TEF")
    lstMenu.AddItem ("Operações Administrativas")
    lstMenu.AddItem ("Operações de Coleta PinPad")
    lstMenu.ListIndex = 0
    
    buttonClicked = False
End Sub

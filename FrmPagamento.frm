VERSION 5.00
Begin VB.Form FrmPagamento 
   Caption         =   "Pagamentos"
   ClientHeight    =   8460
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   13650
   LinkTopic       =   "Form1"
   ScaleHeight     =   8460
   ScaleWidth      =   13650
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frameLogs 
      Caption         =   "Logs"
      Height          =   7935
      Left            =   5760
      TabIndex        =   10
      Top             =   240
      Width           =   7695
      Begin VB.TextBox txtLogs 
         Height          =   7335
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   11
         Top             =   360
         Width           =   7215
      End
   End
   Begin VB.CommandButton btnCancelar 
      Caption         =   "Cancelar"
      Height          =   495
      Left            =   1920
      TabIndex        =   9
      Top             =   7440
      Width           =   1575
   End
   Begin VB.Frame frameOperador 
      Caption         =   "Processamento Operador"
      Height          =   3375
      Left            =   240
      TabIndex        =   4
      Top             =   4800
      Width           =   5295
      Begin VB.CommandButton btnOk 
         Caption         =   "OK"
         Height          =   495
         Left            =   3480
         TabIndex        =   8
         Top             =   2640
         Width           =   1575
      End
      Begin VB.ListBox lstOperador 
         Height          =   645
         Left            =   240
         TabIndex        =   7
         Top             =   1680
         Visible         =   0   'False
         Width           =   4815
      End
      Begin VB.TextBox txtOperador 
         Height          =   285
         Left            =   120
         TabIndex        =   6
         Top             =   1080
         Width           =   4935
      End
      Begin VB.Label lblOperador 
         Caption         =   "Label Operador"
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   480
         Width           =   4935
      End
   End
   Begin VB.CommandButton btnIniciarPIX 
      Caption         =   "Iniciar PIX"
      Height          =   975
      Left            =   3240
      TabIndex        =   3
      Top             =   3240
      Width           =   1695
   End
   Begin VB.Frame frameValor 
      Caption         =   "Valor"
      Height          =   4455
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   5295
      Begin VB.CommandButton btnIniciarTEF 
         Caption         =   "Iniciar TEF"
         Height          =   975
         Left            =   480
         TabIndex        =   2
         Top             =   3120
         Width           =   1695
      End
      Begin VB.TextBox lblValor 
         Height          =   615
         Left            =   480
         TabIndex        =   1
         Text            =   "1.27"
         Top             =   480
         Width           =   3975
      End
   End
End
Attribute VB_Name = "FrmPagamento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim retornoUI As String
Dim valorTotal As String
Dim operacaoAtual As Integer
Dim cancelarColeta As String
Dim continuaColeta As Boolean

' =============================
' ====== MÉTODOS DE UI ========
' =============================

Private Sub OkEvent()
    Dim retList As String
    Dim retTxt As String
    
    retornoUI = ""
    
    If lstOperador.Visible Then
        If lstOperador.ListIndex = -1 Then
            MsgBox "Escolha uma opção"
            Exit Sub
        End If
    End If
    
    If txtOperador.Visible Then
        If txtOperador.Text = "" Then
            MsgBox "Escreva o valor pedido"
            Exit Sub
        End If
    End If
    
    retList = CStr(lstOperador.ListIndex)
    retTxt = txtOperador.Text
    
    ' reseta UI
    txtOperador.Text = ""
    lblOperador.Visible = False
    txtOperador.Visible = False
    btnOk.Visible = False
    btnCancelar.Visible = False
    
    ' define variavel global como retorno do usuário
    If lstOperador.Visible Then
        retornoUI = retList
    Else
        retornoUI = retTxt
    End If
    lstOperador.Visible = False
    
    ' retoma a execução do fluxo de coleta
    continuaColeta = True
End Sub

Private Sub btnOk_Click()
    OkEvent
End Sub

Private Sub Form_Load()
    continuarColeta = False
    lblOperador.Visible = False
    txtOperador.Visible = False
    lstOperador.Visible = False
    btnCancelar.Visible = False
    btnOk.Visible = False
End Sub

Private Sub txtOperador_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then       'User pressed the Enter key
        OkEvent
    End If
End Sub

Private Sub btnCancelar_Click()
    ' define a variável global retornoUI = 0
    retornoUI = "0"
    cancelarColeta = "9"
    continuarColeta = True
End Sub

Private Sub btnIniciarPIX_Click()
    operacaoAtual = Defines.OPERACAO_PIX
    
    lblOperador.Visible = True
    lblOperador.Caption = "AGUARDE..."
    
    lblValor.Text = ""
    
    ElginTEF
End Sub

Private Sub btnIniciarTEF_Click()
    operacaoAtual = Defines.OPERACAO_TEF
    
    lblOperador.Visible = True
    lblOperador.Caption = "AGUARDE..."
    
    lblValor.Text = ""
    
    ElginTEF
End Sub

Private Sub printTela(ByVal msg As String)
    ' reseta UI
    lstOperador.Visible = False
    txtOperador.Visible = False
    lblOperador.Visible = False
    btnOk.Visible = False
    btnCancelar.Visible = False
    ' imgQrCode.visible = false
    
    ' qrcode pix
    If InStr(msg, "QRCODE;") Then
    Else
        lblOperador.Caption = msg
        lblOperador.Visible = True
        
        If Utils.MostrarBotoes(msg) Then
            txtOperador.Visible = True
            txtOperador.SetFocus
            btnOk.Visible = True
            btnCancelar.Visible = True
        End If
    End If
End Sub

Private Sub printTelaArray(elements() As String)
    Dim i As Long
    
    lstOperador.Clear
    
    lstOperador.Visible = False
    txtOperador.Visible = False
    lblOperador.Visible = False
    btnOk.Visible = False
    btnCancelar.Visible = False
    ' imgqrcode visible  = false
    
    lblOperador.Visible = True
    btnCancelar.Visible = True
    btnOk.Visible = True
    
    For i = LBound(elements) To UBound(elements)
        lstOperador.AddItem (elements(i))
    Next i
    
    lstOperador.Visible = True
End Sub

Private Sub writeLogs(ByVal logs As String)
    txtLogs.Text = txtLogs & Defines.DIV_LOGS & logs
End Sub

' ================================================
' =============== LÓGICA DO TEF ==================
' ================================================

Private Sub ElginTEF()
    Dim start As String
    Dim retorno As String
    Dim sequencial As String
    Dim resp As String
    Dim comprovanteLoja As String
    Dim comprovanteCliente As String
    Dim cnf As String
    Dim endFinalizar As String
    
    ' (1) INICIAR CONEXÃO COM CLIENT
    start = Iniciar
    
    ' fas o parse do retorno da função iniciar
    retorno = GetRetorno(start)
    ' dependendo do resultado da função iniciar definido na variável "retorno" o
    ' fluxo poderá terminar ou continuar
    If retorno <> "1" Then
        Finalizar
        Exit Sub
    End If
    
    ' (2) REALIZAR OPERAÇÃO
    sequencial = incrementarSequencial(GetSequencial(start))
    
    If operacaoAtual = Defines.OPERACAO_TEF Then
        resp = Vender(0, sequencial, Defines.OPERACAO_TEF)
    Else
        resp = Vender(0, sequencial, Defines.OPERACAO_PIX)
    End If
    
    retorno = GetRetorno(resp)
    
    If retorno = "" Then
        resp = Coletar(Defines.OPERACAO_TEF, Jsonify(resp))
        retorno = GetRetorno(resp)
    End If
    
    ' (3) VERIFICAR RESULTADO / CONFIRMAR
    If retorno = "" Then
        writeLogs ("ERRO AO COLETAR DADOS")
        printTela ("ERROR AO COLETAR DADOS")
    ElseIf retorno = "0" Then
        comprovanteLoja = GetComprovante(resp, "loja")
        comprovanteCliente = GetComprovante(resp, "cliente")
        writeLogs (comprovanteLoja)
        writeLogs (comprovanteCliente)
        writeLogs ("TRANSAÇÃO OK< INICIANDO CONFIRMAÇÃO...")
        printTela ("TRANSAÇÃO OK< INICIANDO CONFIRMAÇÃO...")
        
        sequencial = GetSequencial(resp)
        
        cnf = Confirmar(sequencial)
        
        retorno = GetRetorno(cnf)
        
        If retorno <> "1" Then
            Finalizar
        End If
    ElseIf retorno = "1" Then
        writeLogs ("TRANSAÇÃO OK")
        printTela ("TRANSAÇÃO OK")
    Else
        writeLogs ("ERRO NA TRANSAÇÃO")
        printTela ("ERRO NA TRANSAÇÃO")
    End If
    
    ' (4) FINALIZA CONEXÃO
    endFinalizar = Finalizar
    retorno = GetRetorno(endFinalizar)
    If retorno <> "1" Then
        Finalizar
        Exit Sub
    End If
End Sub

' ================================================
' ====== MÉTODOS PARA CONTROLE DA TRANSAÇÃO ======
' ================================================


Private Function Iniciar() As String
    Dim resultado As String
    Dim payload As JsonBag
    Set payload = New JsonBag
    
    ' add examples
    
    resultado = StrPtrToString(IniciarOperacaoTEF(Stringify(payload)))
    
    ' logs
    Set payload = Nothing
    
    Iniciar = resultado
End Function

Private Function Vender(ByVal cartao As Integer, ByVal sequencial As String, ByVal operacao As Integer) As String
    Dim resultado As String
    Dim payload As JsonBag
    Set payload = New JsonBag
    
    ' logs
    writeLogs ("VENDER: " & "SEQUENCIAL USADO NA VENDA" & sequencial)
    
    payload.Item("sequencial") = sequencial
    
    ' verificar valorTotal
    If valorTotal <> "" Then
        payload.Item("valorTotal") = valorTotal
    End If
    
    If operacao = Defines.OPERACAO_TEF Then
        resultado = StrPtrToString(RealizarPagamentoTEF(CLng(cartao), Stringify(payload), True))
    Else
        resultado = StrPtrToString(RealizarPixTEF(Stringify(payload), True))
    End If
    
    ' logs
    writeLogs ("VENDER: " & Jsonify(resultado).JSON)
    
    Set payload = Nothing
    
    Vender = resultado
End Function

Private Function Coletar(ByVal operacao As Integer, ByVal root As JsonBag) As String
    ' chaves utilizadas na coleta
    Dim coletaRetorno As String ' In/Out; out: 0 = continuar coleta, 9 = cancelar coleta
    Dim coletaSequencial As String ' In/Out
    Dim coletaMensagem As String ' In/[Out]
    Dim coletaTipo As String ' In
    Dim coletaOpcao As String ' In
    Dim coletaMascara As String
    Dim coletaInformacao As String ' Out
    Dim payload As JsonBag
    Dim resp As String
    Dim retorno As String
    Dim opcoes() As String
    Dim elements() As String
    Dim i As Integer
    
    ' extrair dados da resposta / coleta
    coletaRetorno = GetStringValue(root, "tef", "automacao_coleta_retorno")
    coletaSequencial = GetStringValue(root, "tef", "automacao_coleta_sequencial")
    coletaMensagem = GetStringValue(root, "tef", "mensagemResultado")
    coletaTipo = GetStringValue(root, "tef", "automacao_coleta_tipo")
    coletaOpcao = GetStringValue(root, "tef", "automacao_coleta_opcao")
    coletaMascara = GetStringValue(root, "tef", "automacao_coleta_mascara")
    writeLogs ("COLETAR: " & UCase(coletaMensagem))
    printTela (UCase(coletaMensagem))
    
    ' em caso de erro, encerra coleta
    If coletaRetorno <> "0" Then
        Coletar = Stringify(root)
        Exit Function
    End If
    
    ' em caso de sucesso, monta o novo payload e continua a coleta
    Set payload = New JsonBag
    payload.Item("automacao_coleta_retorno") = coletaRetorno
    payload.Item("automacao_coleta_sequencial") = coletaSequencial
    
    ' COLETA DADOS DO USUÁRIO
    If coletaTipo <> "" Then
        If coletaOpcao = "" Then
            writeLogs ("INFORME O VALOR SOLICITADO: ")
            coletaInformacao = ReadInput
            payload.Item("automacao_coleta_informacao") = coletaInformacao
        ElseIf coletaOpcao <> "" Then
            opcoes = Split(coletaOpcao, ";")
            ReDim elements(UBound(opcoes))
            
            For i = 0 To UBound(opcoes)
                elements(i) = "[" & i & "] " & UCase(opcoes(i)) & vbCrLf
                writeLogs ("[" & i & "] " & UCase(opcoes(i)) & vbCrLf)
            Next i
            
            printTelaArray elements
            writeLogs (vbCrLf & "SELECIONE A OPÇÃO DESEJADA: ")
            
            Dim read As String
            read = ReadInput
            coletaInformacao = opcoes(CInt(read))
            payload.Item("automacao_coleta_informacao") = coletaInformacao
        End If
        
        ' verifica variável global "cancelarColeta"
        If cancelarColeta <> "" Then
            payload.Item("automacao_coleta_retorno") = cancelarColeta
            cancelarColeta = ""
        End If
    End If
    
    ' informa dados coletados
    If operacao = Defines.OPERACAO_ADM Then
        resp = StrPtrToString(RealizarAdmTEF(0, Stringify(payload), False))
    Else
        If operacaoAtual = Defines.OPERACAO_PIX Then
            resp = StrPtrToString(RealizarPixTEF(Stringify(payload), False))
        Else
            resp = StrPtrToString(RealizarPagamentoTEF(0, Stringify(payload), False))
        End If
    End If
    
    ' libera memória ocupada pela instancia jsonbag
    Set payload = Nothing
    
    writeLogs (Jsonify(resp).JSON)
    
    ' verificar fim da coleta
    retorno = GetRetorno(resp)
    If retorno <> "" Then
        Coletar = resp
        Exit Function
    End If
    
    ' passa para próxima fase da coleta chamando novamente a função até
    ' que a coleta seja finalizada
    Coletar = Coletar(operacao, Jsonify(resp))
End Function


Private Function Confirmar(ByVal sequencial As String) As String
    Dim resultado As String
    
    writeLogs ("CONFIRMAR: " & "SEQUENCIAL DA OPERAÇÂO A SER CONFIRMADA: ")
    printTela ("AGUARDE, CONFIRMANDO OPERAÇÃO...")
    
    resultado = StrPtrToString(ConfirmarOperacaoTEF(CLng(sequencial), 1))
    writeLogs ("CONFIRMAR: " & Jsonify(resultado).JSON)
    Confirmar = resultado
End Function

Private Function Finalizar() As String
    Dim resultado As String
    
    resultado = StrPtrToString(FinalizarOperacaoTEF(1))
    writeLogs ("FINALIZAR: " & Jsonify(resultado).JSON)
    valorTotal = ""
    printTela ("OPERAÇÃO FINALIZADA")
    Finalizar = resultado
End Function

Private Function ReadInput() As String
    Do While Not continuaColeta
        DoEvents
    Loop
    continuaColeta = False
    ReadInput = retornoUI
End Function

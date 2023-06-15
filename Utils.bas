Attribute VB_Name = "Utils"
Public Function incrementarSequencial(ByVal sequencial As String) As String
    Dim seq As String
    If IsNumeric(sequencial) Then
        seq = CInt(sequencial)
    Else
      seq = ""
    End If
    incrementarSequencial = seq
End Function

Public Function Jsonify(ByVal jsonString As String) As JsonBag
    Dim JB As JsonBag
    Set JB = New JsonBag
    JB.JSON = jsonString
    Set Jsonify = JB
End Function

Public Function Stringify(ByVal JSON As JsonBag) As String
    Stringify = JSON.JSON
End Function

Public Function GetStringValue(ByVal JSON As JsonBag, ByVal outerKey As String, Optional ByVal innerKey As String = "") As String
    Dim ret As String
    Dim innerjson As JsonBag
    
    If JSON.Exists(outerKey) Then
        If innerKey = "" Then
            ret = JSON.Item(outerKey)
        ElseIf TypeOf JSON.Item(outerKey) Is JsonBag Then
            Set innerjson = JSON.Item(outerKey)
            If innerjson.Exists(innerKey) Then
                ret = JSON.Item(outerKey).Item(innerKey)
            Else
                ret = ""
            End If
        Else
            ret = ""
        End If
    Else
        ret = ""
    End If
    GetStringValue = ret
End Function

Public Function GetRetorno(ByVal resp As String) As String
    GetRetorno = GetStringValue(Jsonify(resp), "tef", "retorno")
End Function

Public Function GetSequencial(ByVal resp As String) As String
    GetSequencial = GetStringValue(Jsonify(resp), "tef", "sequencial")
End Function

Public Function GetComprovante(ByVal resp As String, ByVal via As String) As String
    Dim ret As String
    If via = "loja" Then
        ret = GetStringValue(Jsonify(resp), "tef", "comprovanteDiferenciadoLoja")
    ElseIf via = "cliente" Then
        ret = GetStringValue(Jsonify(resp), "tef", "comprovanteDiferenciadoPortador")
    Else
        ret = ""
    End If
    GetComprovante = ret
End Function

Public Function MostrarBotoes(ByVal mensagem As String) As Boolean
    Dim msgArray As Variant
    msgArray = Array("aguarde", "finalizada", "passagem", "cancelada", "iniciando confirmação")
    
    Dim I As Integer
    Dim msgToLower As String
    msgToLower = LCase(mensagem)
    
    Dim showButtons As Boolean
    showButtons = True
    
    For I = LBound(msgArray) To UBound(msgArray)
        If InStr(msgToLower, LCase(msgArray(I))) <> 0 Then
            showButtons = False
            Exit For
        End If
    Next I
    
    MostrarBotoes = showButtons
End Function


'Função para conversão de ponteiro para String
Public Function StrPtrToString(ByVal ponteiro As Long) As String
    Dim Saida As String
    Saida = SysAllocStringByteLen(ponteiro, lstrlenA(ponteiro))
    StrPtrToString = Saida
End Function

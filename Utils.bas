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
    JB.json = jsonString
    Set Jsonify = JB
End Function

Public Function Stringify(ByVal json As JsonBag) As String
    Stringify = json.json
End Function

Public Function GetStringValue(ByVal json As JsonBag, ByVal outerKey As String, Optional ByVal innerKey As String = "") As String
    Dim ret As String
    Dim innerjson As JsonBag
    
    If json.Exists(outerKey) Then
        If innerKey = "" Then
            ret = json.Item(outerKey)
        ElseIf TypeOf json.Item(outerKey) Is JsonBag Then
            Set innerjson = json.Item(outerKey)
            If innerjson.Exists(innerKey) Then
                ret = json.Item(outerKey).Item(innerKey)
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


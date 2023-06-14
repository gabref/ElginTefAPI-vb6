VERSION 5.00
Begin VB.Form FrmMain 
   Caption         =   "Form1"
   ClientHeight    =   6750
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   11055
   LinkTopic       =   "Form1"
   ScaleHeight     =   6750
   ScaleWidth      =   11055
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   2295
      Left            =   3480
      TabIndex        =   0
      Top             =   1920
      Width           =   3975
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim test As String
    Dim testjb As JsonBag
    Dim F As Integer
    Dim JsonData As String
    
    F = FreeFile(0)
    Open "JsonSample.txt" For Input As #F
    JsonData = Input$(LOF(F), #F)
    Close #F
    
    Set testjb = Jsonify(JsonData)
    test = Stringify(testjb)
    
End Sub

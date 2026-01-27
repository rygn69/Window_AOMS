VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3030
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   735
      Left            =   2280
      TabIndex        =   2
      Top             =   840
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   360
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   1560
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   600
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Text2.Text = EncryptString(Text1.Text)
End Sub

Public Function EncryptString(strNative As String) As String
Dim strEncrypt As String
Dim strI As Integer, temp1 As Integer
Dim intLen As Integer
Dim intAve As Integer

  If Len(strNative) = 0 Then
    Encrypt_PWord = ""
    Exit Function
  End If
  strEncrypt = ""
  intLen = Len(strNative)
  intAve = 0
  For strI = 1 To intLen
    intAve = intAve + Asc(Mid(strNative, strI, 1))
  Next strI
  intAve = (intAve / intLen) + intLen
  For strI = 1 To intLen
    temp1 = Asc(Mid(strNative, strI, 1)) + strI '+ intAve
    strEncrypt = Chr(temp1) & strEncrypt
  Next strI
  strEncrypt = strEncrypt 'Chr(intAve) & strEncrypt
  EncryptString = strEncrypt
End Function

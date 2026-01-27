VERSION 5.00
Begin {BD4B4E61-F7B8-11D0-964D-00A0C9273C2A} rptJEVNew 
   ClientHeight    =   13830
   ClientLeft      =   0
   ClientTop       =   390
   ClientWidth     =   17820
   OleObjectBlob   =   "rptJEVNew.dsx":0000
End
Attribute VB_Name = "rptJEVNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Trantype As Integer

Private Sub Section8_Format(ByVal pFormattingInfo As Object)
    If Trantype = 1 Then txtCollection.SetText Chr(254) Else txtCollection.SetText Chr(168)
    If Trantype = 2 Then txtCheck.SetText Chr(254) Else txtCheck.SetText Chr(168)
    If Trantype = 3 Then txtCash.SetText Chr(254) Else txtCash.SetText Chr(168)
    If Trantype = 4 Then txtOthers.SetText Chr(254) Else txtOthers.SetText Chr(168)
End Sub


VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frm_UnfilteredcashFlow 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Unfiltered Cash Flow"
   ClientHeight    =   7590
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4110
   Icon            =   "frm_UnfilteredcashFlow.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7590
   ScaleWidth      =   4110
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   3120
      Top             =   120
   End
   Begin lvButton.lvButtons_H lvButtons_H1 
      Height          =   615
      Left            =   3120
      TabIndex        =   1
      Top             =   600
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   1085
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      ImgAlign        =   4
      Image           =   "frm_UnfilteredcashFlow.frx":076A
      ImgSize         =   24
      cBack           =   -2147483633
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6030
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   3855
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   1200
      TabIndex        =   3
      Top             =   720
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "MMMM yyyy"
      Format          =   58916867
      CurrentDate     =   41047
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Month Year:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   -120
      TabIndex        =   4
      Top             =   840
      Width           =   1290
   End
   Begin VB.Label Label1 
      Caption         =   "List of Ulfiltered Cash Flow"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   3855
   End
End
Attribute VB_Name = "frm_UnfilteredcashFlow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub LoadUCashflow()
Dim rec As New ADODB.Recordset
Dim x As Long
End Sub

Public Sub Loaddata()
Call lvButtons_H1_Click
End Sub
Private Sub lvButtons_H1_Click()
Dim rec As New ADODB.Recordset
Dim x As Long
List1.Clear
rec.Open "Exec [MPproc_CheckIfFilterIntoCF_gf] @date_ = '" & DTPicker1.Value & "'", opndbaseFMIS, adOpenStatic, adLockOptimistic
If rec.RecordCount > 0 Then
    For x = 1 To rec.RecordCount
            If CheckIfHaveCFentry(rec.Fields!jevno) = False Then
                opndbaseFMIS.Execute "Update dbo.tblAMIS_FinalJEV set filterInCashflow = 1 where jevno = '" & rec.Fields!jevno & "' and actioncode = 1"
            Else
                List1.AddItem rec.Fields!jevno
            End If
        rec.MoveNext
    Next x
End If
rec.Close
End Sub
Private Sub Timer1_Timer()
Me.ZOrder (0)
DoEvents
End Sub

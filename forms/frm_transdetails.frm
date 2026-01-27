VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frm_transdetails 
   Caption         =   "Find Transaction Details "
   ClientHeight    =   7470
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9120
   Icon            =   "frm_transdetails.frx":0000
   MDIChild        =   -1  'True
   Picture         =   "frm_transdetails.frx":076A
   ScaleHeight     =   7470
   ScaleWidth      =   9120
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Caption         =   "Search By:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   8
      Top             =   1200
      Width           =   6975
      Begin VB.OptionButton Option1 
         Caption         =   "Responsibility Center"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   2280
         TabIndex        =   10
         Top             =   360
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Obligation number"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Value           =   -1  'True
         Width           =   2055
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   5055
      Left            =   120
      TabIndex        =   3
      Top             =   3435
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   8916
      _Version        =   393216
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      AllowUserResizing=   1
      Appearance      =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.ListBox lstobrno 
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
      Height          =   3630
      ItemData        =   "frm_transdetails.frx":AE19
      Left            =   120
      List            =   "frm_transdetails.frx":AE20
      TabIndex        =   2
      Top             =   4200
      Visible         =   0   'False
      Width           =   2175
   End
   Begin MSComCtl2.Animation Animation1 
      Height          =   495
      Left            =   8280
      TabIndex        =   4
      Top             =   2880
      Visible         =   0   'False
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   873
      _Version        =   393216
      AutoPlay        =   -1  'True
      FullWidth       =   25
      FullHeight      =   33
   End
   Begin lvButton.lvButtons_H lvButtons_H6 
      Height          =   495
      Left            =   2760
      TabIndex        =   5
      Top             =   2760
      Visible         =   0   'False
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   873
      Caption         =   "&Find"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFore           =   0
      cFHover         =   33023
      cBhover         =   8438015
      LockHover       =   3
      cGradient       =   33023
      Gradient        =   3
      CapStyle        =   1
      Mode            =   0
      Value           =   0   'False
      Image           =   "frm_transdetails.frx":AE39
      cBack           =   16777215
   End
   Begin lvButton.lvButtons_H lvButtons_H1 
      Height          =   495
      Left            =   3840
      TabIndex        =   6
      Top             =   2760
      Visible         =   0   'False
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   873
      Caption         =   "&Close"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFore           =   0
      cFHover         =   33023
      cBhover         =   8438015
      LockHover       =   3
      cGradient       =   33023
      Gradient        =   3
      CapStyle        =   1
      Mode            =   0
      Value           =   0   'False
      Image           =   "frm_transdetails.frx":20E63
      cBack           =   16777215
   End
   Begin VB.Frame frme_RC 
      Caption         =   "Search By:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   120
      TabIndex        =   11
      Top             =   2040
      Width           =   6975
      Begin VB.CheckBox Check1 
         Caption         =   "Direct to PTO"
         Height          =   255
         Left            =   2040
         TabIndex        =   20
         Top             =   360
         Width           =   1335
      End
      Begin VB.TextBox txtSearch 
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
         Height          =   360
         Left            =   120
         TabIndex        =   12
         Top             =   720
         Width           =   2175
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Obligation number:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   360
         Width           =   2175
      End
   End
   Begin VB.Frame frme_OBlig 
      Caption         =   "Search By:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   120
      TabIndex        =   14
      Top             =   2040
      Visible         =   0   'False
      Width           =   6975
      Begin VB.OptionButton Option1 
         Caption         =   "Paid Transaction"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   4560
         TabIndex        =   19
         Top             =   960
         Width           =   2055
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Accounts Payable"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   4560
         TabIndex        =   18
         Top             =   600
         Width           =   2055
      End
      Begin VB.OptionButton Option1 
         Caption         =   "All"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   4560
         TabIndex        =   17
         Top             =   240
         Width           =   2055
      End
      Begin VB.ComboBox cmbrc 
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
         Height          =   360
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   16
         Text            =   "cmbrc"
         Top             =   720
         Width           =   4095
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Responsibility Center"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   360
         Visible         =   0   'False
         Width           =   2175
      End
   End
   Begin VB.Label lblrecordcount 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   8520
      TabIndex        =   7
      Top             =   1560
      Width           =   45
   End
   Begin VB.Label lblException 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Transaction Details"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1860
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Note: After typing Press ENTER"
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
      Left            =   120
      TabIndex        =   0
      Top             =   390
      Width           =   2670
   End
End
Attribute VB_Name = "frm_transdetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbrc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    
Else
KeyAscii = AutoFind(cmbRC, KeyAscii, True)
End If
End Sub

Private Sub Form_Load()
 LoadOffice
End Sub

Private Sub Form_Resize()
On Error Resume Next
MSHFlexGrid1.Width = Me.ScaleWidth - 350
  MSHFlexGrid1.Height = Me.ScaleHeight - MSHFlexGrid1.Top - 500
 
End Sub

Private Sub lvButtons_H1_Click()
If MsgBox("Are you sure do you want to close this form?", vbCritical + vbYesNo, "System Message") = vbYes Then
Unload Me
End If
End Sub

Private Sub lvButtons_H6_Click()
Call LoadDetails
End Sub

Private Sub Option1_Click(Index As Integer)
Select Case Index
Case 0
LoadCondition
Case 1
LoadCondition
Case 2

Case 3
Case 4
End Select
End Sub
Private Sub LoadCondition()
If Option1(0).Value = True Then
frme_OBlig.Visible = True
frme_RC.Visible = False
Else
frme_OBlig.Visible = False
frme_RC.Visible = True
End If
End Sub

Private Sub txtSearch_KeyPress(KeyAscii As Integer)
    Dim rec As New ADODB.Recordset
    Dim ZZ As Long
If KeyAscii = 13 Then
    Call LoadDetails
End If
End Sub
Public Sub LoadDetails()
Dim rec As New ADODB.Recordset
    Animation1.Visible = True
    Animation1.Open App.path & AViLocation & "\refresh.avi"
    Animation1.Play
    DoEvents
    If Check1.Value = 1 Then
    rec.Open "EXECUTE MPproc_GetTransthroughOBR @Obrno = '" & txtSearch.Text & "',@PTO = 1", opndbaseFMIS, adOpenStatic
    Else
    rec.Open "EXECUTE MPproc_GetTransthroughOBR @Obrno = '" & txtSearch.Text & "',@PTO = 0", opndbaseFMIS, adOpenStatic
    End If
    MSHFlexGrid1.Rows = 2
    MSHFlexGrid1.Cols = 10
    
    If rec.RecordCount > 0 Then
        Set MSHFlexGrid1.DataSource = rec
    End If
    Call SetGrid
    Animation1.Visible = False
    rec.Close
    Set rec = Nothing
End Sub
Private Sub SetGrid()
Dim cc As Integer

On Error Resume Next
'    MSHFlexGrid1.Rows = 2
    ' IIf(LCase(Trim(lblMode)) = "edit", 6, 5)
    
'    MSHFlexGrid1.TextMatrix(0, 0) = "Dvno"
'    MSHFlexGrid1.TextMatrix(0, 1) = "Obr No"
'    MSHFlexGrid1.TextMatrix(0, 2) = "RCI no"
'    MSHFlexGrid1.TextMatrix(0, 3) = "Check No."
'    MSHFlexGrid1.TextMatrix(0, 4) = "Check Date"
'    MSHFlexGrid1.TextMatrix(0, 5) = "PTV No."
'    MSHFlexGrid1.TextMatrix(0, 6) = "Particular"
'    MSHFlexGrid1.TextMatrix(0, 7) = "Officename"
'    MSHFlexGrid1.TextMatrix(0, 8) = "Officename"
'    MSHFlexGrid1.TextMatrix(0, 9) = "Officename"
    
    MSHFlexGrid1.ColWidth(0) = 1500
    MSHFlexGrid1.ColWidth(1) = 2000
    MSHFlexGrid1.ColWidth(2) = 7000
    MSHFlexGrid1.ColWidth(3) = 6000
    MSHFlexGrid1.ColWidth(4) = 3000
    MSHFlexGrid1.ColWidth(5) = 1500
'    MSHFlexGrid1.ColWidth(6) = 1500
 '   MSHFlexGrid1.ColWidth(7) = 1500
  '  MSHFlexGrid1.ColWidth(8) = 1500
   ' MSHFlexGrid1.ColWidth(9) = 1500
    'MSHFlexGrid1.ColWidth(10) = 1000
    
    For cc = 0 To MSHFlexGrid1.Cols - 1
        MSHFlexGrid1.Row = 0
        MSHFlexGrid1.col = cc
        MSHFlexGrid1.CellAlignment = 4
    Next cc
 
    'Else
    '    MSHFlexGrid1.ColWidth(5) = 0
    'End If
End Sub
Public Sub LoadOffice()
Dim OREc As New ADODB.Recordset
Dim x As Integer

cmbRC.Clear

OREc.Open ("Select distinct * FRom tblREF_AIS_Offices Order By [OfficeMedium]"), opndbaseFMIS, adOpenStatic, adLockOptimistic
If OREc.RecordCount > 0 Then
    For x = 1 To OREc.RecordCount
        cmbRC.AddItem OREc![OfficeMedium]
        cmbRC.ItemData(cmbRC.NewIndex) = OREc!fmisofficeid
        OREc.MoveNext
    Next x
End If
OREc.Close
Set OREc = Nothing

End Sub


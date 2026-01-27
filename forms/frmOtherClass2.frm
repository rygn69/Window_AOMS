VERSION 5.00
Object = "{1693405E-2DC9-4248-B52F-4AC9145DA2AF}#1.0#0"; "WinXPC Engine.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmOtherClass2 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7470
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10785
   Icon            =   "frmOtherClass2.frx":0000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmOtherClass2.frx":076A
   ScaleHeight     =   7470
   ScaleWidth      =   10785
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   6285
      Left            =   120
      ScaleHeight     =   6255
      ScaleWidth      =   10545
      TabIndex        =   1
      Top             =   1080
      Width           =   10575
      Begin VB.TextBox txt_entry 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   7560
         TabIndex        =   12
         Top             =   840
         Visible         =   0   'False
         Width           =   435
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSFlexGrid1 
         Height          =   6255
         Left            =   0
         TabIndex        =   10
         Top             =   0
         Width           =   10575
         _ExtentX        =   18653
         _ExtentY        =   11033
         _Version        =   393216
         ForeColorSel    =   65535
         AllowUserResizing=   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin lvButton.lvButtons_H lvlbrowse 
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   720
         Visible         =   0   'False
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   661
         Caption         =   "..."
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   13.5
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
         ImgAlign        =   5
         Image           =   "frmOtherClass2.frx":AE19
         cBack           =   16777215
      End
   End
   Begin lvButton.lvButtons_H lvButtons_H1 
      Height          =   615
      Left            =   9840
      TabIndex        =   2
      Top             =   360
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   1085
      Caption         =   "&Ok"
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
      Image           =   "frmOtherClass2.frx":AF73
      cBack           =   16777215
   End
   Begin WinXPC_Engine.WindowsXPC WindowsXPC1 
      Left            =   6240
      Top             =   8880
      _ExtentX        =   6588
      _ExtentY        =   1085
      ColorScheme     =   2
      Common_Dialog   =   0   'False
   End
   Begin VB.PictureBox freeSizer1 
      Height          =   480
      Left            =   5040
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   3
      Top             =   9240
      Width           =   1200
   End
   Begin lvButton.lvButtons_H lvButtons_H4 
      Height          =   615
      Left            =   13200
      TabIndex        =   0
      Top             =   8280
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1085
      Caption         =   "&Back"
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
      cGradient       =   0
      CapStyle        =   1
      Mode            =   0
      Value           =   0   'False
      Image           =   "frmOtherClass2.frx":B2C5
      ImgSize         =   24
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H lvButtons_H3 
      Height          =   495
      Left            =   4920
      TabIndex        =   9
      Top             =   2400
      Visible         =   0   'False
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   873
      Caption         =   "&Load"
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
      Image           =   "frmOtherClass2.frx":EDCF
      cBack           =   16777215
   End
   Begin lvButton.lvButtons_H lvButtons_H5 
      Height          =   615
      Left            =   6120
      TabIndex        =   11
      Top             =   1560
      Visible         =   0   'False
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1085
      Caption         =   "&Query Settings"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
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
      Image           =   "frmOtherClass2.frx":128D9
      cBack           =   16777215
   End
   Begin VB.Frame Frame1 
      Caption         =   "View"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   5040
      TabIndex        =   4
      Top             =   3480
      Visible         =   0   'False
      Width           =   6015
      Begin VB.OptionButton Option3 
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
         Left            =   4680
         TabIndex        =   7
         Top             =   360
         Width           =   975
      End
      Begin VB.OptionButton Option2 
         Caption         =   "W/o Accountcode"
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
         Left            =   2400
         TabIndex        =   6
         Top             =   360
         Width           =   1935
      End
      Begin VB.OptionButton Option1 
         Caption         =   "With Accountcode"
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
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   1935
      End
   End
   Begin VB.ComboBox Combo1 
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
      Left            =   720
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   2880
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Note: To change the Accountcode, Double Click it."
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
      TabIndex        =   16
      Top             =   720
      Width           =   4275
   End
   Begin VB.Label lblException 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Conversion Table for Chart of Accounts"
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
      TabIndex        =   15
      Top             =   120
      Width           =   3825
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Please identify the list below"
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
      TabIndex        =   14
      Top             =   390
      Width           =   2430
   End
End
Attribute VB_Name = "frmOtherClass2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Reff, datetimeentered, UserID, Accountname, CName As String
Public Damount, Camount, Gamount As Currency
Public isEdit, inRec, ifCMB, IfEdit, IfNew, insert, delete, IsOK As Boolean
Public Transtype As Integer

Private Sub SetGrid()
Dim cc As Integer

    MSFlexGrid1.Clear
    MSFlexGrid1.Rows = 2
    MSFlexGrid1.Cols = 6 ' IIf(LCase(Trim(lblMode)) = "edit", 6, 5)
    
    'Name
    MSFlexGrid1.TextMatrix(0, 0) = "ID"
    MSFlexGrid1.TextMatrix(0, 1) = "AccountCode"
    MSFlexGrid1.TextMatrix(0, 2) = "Accounts and Explanation"
    MSFlexGrid1.TextMatrix(0, 3) = "Debit"
    MSFlexGrid1.TextMatrix(0, 4) = "Credit"
    MSFlexGrid1.TextMatrix(0, 5) = "ActionCode"
    
    MSFlexGrid1.ColWidth(0) = 0
    MSFlexGrid1.ColWidth(1) = 700
    MSFlexGrid1.ColWidth(2) = 8000
    MSFlexGrid1.ColWidth(3) = 1500
    MSFlexGrid1.ColWidth(4) = 1500
    MSFlexGrid1.ColWidth(5) = 0
    MSFlexGrid1.ColAlignment(1) = 1
End Sub

Private Sub Form_Load()
Call loadDetails(Reff, Transtype)
End Sub
Private Sub lvButtons_H1_Click()
Unload Me
End Sub

Private Sub loadDetails(ByVal Reff As String, types As Integer)
On Error GoTo bad
Dim sql As String
Dim rec As New ADODB.Recordset
sql = "Exec [fmis].[dbo].[MPproc_ChckIfHaveAccnt]  reff = '" & Trim(Reff) & "',@type = " & types & ""
Set rec = opndbaseFMIS.Execute("Exec [fmis].[dbo].[MPproc_NullAccntcode]  @reff = '" & Trim(Reff) & "',@type = " & types & "")
If rec.RecordCount > 0 Then
Call SetGrid
Set MSFlexGrid1.DataSource = rec
End If
Exit Sub
bad:
MsgBox err.Description
End Sub

Private Sub lvButtons_H5_Click()
frm_relatedtableForCOA.Show 1
End Sub


Private Sub lvlbrowse_Click()
Dim rec As New ADODB.Recordset
IsOK = False
With frmforCOA
.Nme = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2)
.accntcode = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 3)
.Trnno = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 1)
.Show 1
End With
If IsOK = True Then
      rec.Open "select tables,columns,conditions from tblAMIS_RelatedTableforCOA where trnno = " & Combo1.ItemData(Combo1.ListIndex) & "", opndbaseFMIS, adOpenStatic
      If rec.RecordCount > 0 Then
        Call UpdateExtractor(IIf(IsNull(rec!Tables), "", rec!Tables), IIf(IsNull(rec!columns), "", rec!columns), Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 3)), IIf(IsNull(rec!Conditions), "", rec!Conditions), Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 1)))
      End If
      rec.Close
End If
End Sub
Private Sub MSFlexGrid1_DblClick()
On Error GoTo bad
    Select Case MSFlexGrid1.col
    Case 3 'Debit/Credit
        
            Dim rec As New ADODB.Recordset
            IsOK = False
            With frmforCOA
            .Nme = Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2))
            .accntcode = Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 3))
            .Trnno = Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 1))
            .Show 1
            End With
            If IsOK = True Then
                  rec.Open "select tables,columns,conditions from tblAMIS_RelatedTableforCOA where trnno = " & Combo1.ItemData(Combo1.ListIndex) & "", opndbaseFMIS, adOpenStatic
                  If rec.RecordCount > 0 Then
                    Call UpdateExtractor(IIf(IsNull(rec!Tables), "", rec!Tables), IIf(IsNull(rec!columns), "", rec!columns), Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 3)), IIf(IsNull(rec!Conditions), "", rec!Conditions), Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 1)))
                  End If
                  rec.Close
            End If
        
    End Select
Exit Sub
bad:
MsgBox err.Description
End Sub

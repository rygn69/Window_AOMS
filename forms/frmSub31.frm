VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{1693405E-2DC9-4248-B52F-4AC9145DA2AF}#1.0#0"; "WinXPC Engine.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmSub31 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Account Details"
   ClientHeight    =   9090
   ClientLeft      =   4005
   ClientTop       =   1665
   ClientWidth     =   13275
   Icon            =   "frmSub31.frx":0000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9090
   ScaleWidth      =   13275
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtformula 
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
      Left            =   4440
      TabIndex        =   19
      Top             =   360
      Width           =   6615
   End
   Begin lvButton.lvButtons_H lvButtons_H1 
      Height          =   615
      Left            =   11160
      TabIndex        =   6
      Top             =   120
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1085
      Caption         =   "&OK"
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
      Image           =   "frmSub31.frx":076A
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
      TabIndex        =   24
      Top             =   9120
      Width           =   1200
   End
   Begin lvButton.lvButtons_H lvButtons_H4 
      Height          =   615
      Left            =   12000
      TabIndex        =   0
      Top             =   8400
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
      Image           =   "frmSub31.frx":0ABC
      ImgSize         =   24
      cBack           =   -2147483633
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   7455
      Left            =   0
      ScaleHeight     =   7425
      ScaleWidth      =   13185
      TabIndex        =   1
      Top             =   840
      Width           =   13215
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   6615
         Left            =   1200
         ScaleHeight     =   6585
         ScaleWidth      =   8610
         TabIndex        =   8
         Top             =   300
         Visible         =   0   'False
         Width           =   8635
         Begin VB.CheckBox Check1 
            BackColor       =   &H80000005&
            Caption         =   "Many"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   21
            Top             =   6720
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.TextBox txtdetails 
            Appearance      =   0  'Flat
            Height          =   495
            Left            =   840
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   13
            Top             =   80
            Width           =   7215
         End
         Begin VB.TextBox txtfind 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3360
            TabIndex        =   11
            Top             =   6120
            Width           =   3375
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
            Height          =   5415
            Left            =   120
            TabIndex        =   9
            Top             =   600
            Width           =   8385
            _ExtentX        =   14790
            _ExtentY        =   9551
            _Version        =   393216
            BackColor       =   16777215
            BackColorSel    =   8454143
            ForeColorSel    =   0
            GridLinesUnpopulated=   1
            SelectionMode   =   1
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
         Begin lvButton.lvButtons_H lvButtons_H3 
            Height          =   375
            Left            =   8160
            TabIndex        =   22
            Top             =   120
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   661
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
            Image           =   "frmSub31.frx":45C6
            cBack           =   16777215
         End
         Begin lvButton.lvButtons_H lvButtons_H5 
            Height          =   375
            Left            =   120
            TabIndex        =   23
            Top             =   6120
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   661
            Caption         =   "Import"
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
            cFore           =   0
            cFHover         =   33023
            cBhover         =   8438015
            LockHover       =   3
            cGradient       =   33023
            Gradient        =   3
            CapStyle        =   1
            Mode            =   0
            Value           =   0   'False
            Image           =   "frmSub31.frx":4720
            cBack           =   16777215
         End
         Begin VB.Label Label4 
            BackColor       =   &H80000005&
            Caption         =   "Details:"
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
            TabIndex        =   14
            Top             =   120
            Width           =   855
         End
         Begin VB.Label Label3 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Press ENTER "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   495
            Left            =   6840
            TabIndex        =   12
            Top             =   6075
            Width           =   1215
         End
         Begin VB.Label Label2 
            BackColor       =   &H80000005&
            Caption         =   "Search Name:"
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
            Left            =   2040
            TabIndex        =   10
            Top             =   6165
            Width           =   1335
         End
      End
      Begin VB.TextBox txt_entry 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
         BorderStyle     =   0  'None
         Height          =   525
         Left            =   8760
         TabIndex        =   3
         Top             =   3480
         Visible         =   0   'False
         Width           =   1665
      End
      Begin VB.ComboBox cmbEntry 
         BackColor       =   &H0080FFFF&
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
         Left            =   10320
         TabIndex        =   2
         Text            =   "cmbEntry"
         Top             =   3120
         Visible         =   0   'False
         Width           =   1695
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   7440
         Left            =   0
         TabIndex        =   4
         Top             =   0
         Width           =   13200
         _ExtentX        =   23283
         _ExtentY        =   13123
         _Version        =   393216
         FixedCols       =   0
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
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   6495
         Left            =   2400
         TabIndex        =   5
         Top             =   300
         Visible         =   0   'False
         Width           =   9375
         _ExtentX        =   16536
         _ExtentY        =   11456
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
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
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Code"
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Explaination"
            Object.Width           =   14111
         EndProperty
      End
   End
   Begin lvButton.lvButtons_H lvButtons_H2 
      Height          =   615
      Left            =   120
      TabIndex        =   7
      Top             =   8400
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   1085
      Caption         =   "&Auto Generate"
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
      Image           =   "frmSub31.frx":59A2
      cBack           =   16777215
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "fx"
      BeginProperty Font 
         Name            =   "Bell MT"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4080
      TabIndex        =   20
      Top             =   360
      Width           =   360
   End
   Begin VB.Label lblname 
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
      Left            =   1680
      TabIndex        =   18
      Top             =   120
      Width           =   6255
   End
   Begin VB.Label Label5 
      Caption         =   "Claimant Name:"
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
      TabIndex        =   17
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label lblamount 
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
      Left            =   1680
      TabIndex        =   16
      Top             =   480
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "Gross Amount:"
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
      Top             =   480
      Width           =   1575
   End
   Begin VB.Menu popup 
      Caption         =   "popup"
      Visible         =   0   'False
      Begin VB.Menu Payroll 
         Caption         =   "Payroll"
      End
      Begin VB.Menu Property 
         Caption         =   "Property"
      End
   End
End
Attribute VB_Name = "frmSub31"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public reff, datetimeentered, UserID, Accountname, CName As String
Public Damount, Camount, Gamount As Currency
Public isEdit, inRec, ifCMB, IfEdit, IfNew, insert, delete As Boolean
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
    MSFlexGrid1.ColWidth(1) = 1700
    MSFlexGrid1.ColWidth(2) = 8000
    MSFlexGrid1.ColWidth(3) = 1500
    MSFlexGrid1.ColWidth(4) = 1500
    MSFlexGrid1.ColWidth(5) = 0
    MSFlexGrid1.ColAlignment(1) = 1
    
    
End Sub


Private Sub cmbEntry_Change()

    If Len(Trim(cmbEntry.Text)) >= 3 Then
        LoadAccountsbySub (cmbEntry.Text)
        txtdetails.Text = LoadAccountsByName(cmbEntry.Text, "Fullname")
    Else
    Call GetAccountNamebyorder("Accountcode")
    End If
    txtfind.Text = ""
Picture2.Move MSFlexGrid1.CellLeft + cmbEntry.Width
End Sub
Private Sub GotonextCell()
'txt_entry.Move MSFlexGrid1.CellLeft(MSHFlexGrid1.Row, 3), MSFlexGrid1.CellTop, MSFlexGrid1.CellWidth, MSFlexGrid1.CellHeight
End Sub

Public Function GetAccountNamebyorder(ByVal Condition As String)
Dim rec As New ADODB.Recordset
Dim x
Dim z As Integer
rec.Open "Select Accountcode,Accountname from tblREF_AIS_ChartOfAccountsMother where accountcode like '" & cmbEntry.Text & "%' and accountname like '" & Trim(txtfind.Text) & "%' order by Accountname", opndbaseFMIS, adOpenStatic, adLockOptimistic
    'lst.ListItems.Clear
        MSHFlexGrid1.Clear
        MSHFlexGrid1.Rows = 2
    If rec.RecordCount > 0 Then
    
'        For z = 1 To rec.RecordCount
'                    Set x = lst.ListItems.Add(, , rec.Fields!Accountcode)
'                    x.SubItems(1) = Trim(rec.Fields!Accountname)
'            rec.MoveNext
'        Next z
    
    Set MSHFlexGrid1.DataSource = rec
        MSHFlexGrid1.Cols = 4
        MSHFlexGrid1.TextMatrix(0, 1) = "Code"
        MSHFlexGrid1.TextMatrix(0, 2) = "Explanation"
        MSHFlexGrid1.TextMatrix(0, 3) = "Formula"
        
        MSHFlexGrid1.ColWidth(0) = 0
        MSHFlexGrid1.ColWidth(1) = 700
        MSHFlexGrid1.ColWidth(2) = 8000
        MSHFlexGrid1.ColWidth(3) = 0
        
        
    End If
'rec.Close
Set rec = Nothing
End Function
Private Sub cmbEntry_Click()
   If Len(Trim(cmbEntry.Text)) > 3 Then
        LoadAccountsbySub (cmbEntry.Text)
    Else
    Call GetAccountNamebyorder("Accountname")
    End If
End Sub

Private Sub cmbEntry_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       ' If cmbEntry.ListIndex <> -1 Then
            inRec = False
            Accountname = LoadAccountsByName(cmbEntry.Text, "Summary")
            If Trim(cmbEntry.Text) <> "" Then
                If inRec = False Then
                    If cmbEntry.Text <> MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 1) Then
                    MsgBox "Invalid Accountcode Please Select Another Account..!", vbCritical, "System Information"
                    Exit Sub
                    End If
                End If
                ifCMB = True
                If Chckentry = False Then
                Exit Sub
                End If
                ifCMB = False
            End If
            
            
            MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 1) = cmbEntry.Text
            
            If MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2) = "TOTAL" Then
                    MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 3) = ""
                    MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 4) = ""
                    
            Else
                If MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2) = "" Then
                MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1
                End If
                MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2) = Accountname
            End If
            
            
        If cmbEntry.Text = "" Then
            If MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2) <> "TOTAL" Then
               
                    MSFlexGrid1.RemoveItem (MSFlexGrid1.Row)
                    If MSFlexGrid1.TextMatrix(MSFlexGrid1.Rows - 1, 2) <> "TOTAL" Then
                        MSFlexGrid1.Rows = MSFlexGrid1.Rows - 1
                    
                End If
            End If
        Else
            If MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2) <> "TOTAL" Then
                MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2) = Accountname
                    If MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2) = "" Then
                    MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1
                    End If
            Else
            MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1
            MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2) = Accountname
            End If
        End If
        cmbEntry.Visible = False
        ListView2.Visible = False
        Picture2.Visible = False
        Call GetSum
        MSFlexGrid1.SetFocus
        isEdit = True
    Else
       ' KeyAscii = AutoFind(cmbEntry, KeyAscii, True)
        ListView2.Visible = True
        ListView2.Move MSFlexGrid1.CellLeft + cmbEntry.Width
        Picture2.Visible = True
       Picture2.Move MSFlexGrid1.CellLeft + cmbEntry.Width
    End If
End Sub
Private Sub Form_Load()
Call SetGrid
lblname.Caption = CName
lblamount.Caption = Format(Gamount, "#,##0.00")
If isEdit = True Then
    Loaddetails (reff)
End If
'Call LoadAccountsByFund(Accountcode)
End Sub
Private Function Loaddetails(ByVal reff As String)
Dim DRec As New ADODB.Recordset
DRec.Open ("Select trnno ,ChildAccountcode, Debit,credit,actioncode,datetimeentered,userid From tblAMIS_AccoutingEntries Where [reffno]='" & reff & "' And (ActionCode=1)"), opndbaseFMIS, adOpenStatic, adLockOptimistic
    Call SetGrid
    If DRec.RecordCount > 0 Then
    
        datetimeentered = DRec!datetimeentered
        UserID = DRec!UserID
        For x = 1 To DRec.RecordCount
            MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1
            MSFlexGrid1.TextMatrix(x, 0) = DRec![Trnno]
            MSFlexGrid1.TextMatrix(x, 1) = DRec!childaccountcode
            MSFlexGrid1.TextMatrix(x, 2) = LoadAccountsByName(DRec!childaccountcode, "Summary")
            MSFlexGrid1.TextMatrix(x, 4) = IIf((Format(DRec!Credit, "#,##0.00") = "0.00"), "", Format(DRec!Credit, "#,##0.00"))
            MSFlexGrid1.TextMatrix(x, 3) = IIf((Format(DRec!Debit, "#,##0.00") = "0.00"), "", Format(DRec!Debit, "#,##0.00"))
            'MSFlexGrid1.TextMatrix(x, 1) = DRec!ActionCode
            DRec.MoveNext
        Next x
        Call GetSum
    Else
    MSFlexGrid1.TextMatrix(1, 2) = "TOTAL"
    End If
    DRec.Close
    Set DRec = Nothing

End Function
Private Function sumAmount(ByVal amnt As String) As String
On Error GoTo sum
Dim x As Integer
Dim Y As String
Dim str() As String
    If Left(amnt, 1) = "+" Then
    Else
    amnt = "+" & amnt
    End If
 
 str = Split(Trim(amnt), "+", -1, vbTextCompare)
 Y = 0

 For x = 1 To 1000
Y = Val(Y) + Val(str(x))
 Next x
 Exit Function
sum:
 If err.Number = 9 Then
 sumAmount = Y
Else
MsgBox "Incorrect Format", vbInformation, "System Message"
End If
End Function
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If isEdit = True Then
    
         q = MsgBox("Save Change..?", vbInformation + vbYesNoCancel, "System Message")
        If q = vbYes Then
            If ChkEntry = True Then
                SaveEntry
            Else
            MsgBox "Total Debit and Total credit Amount are not Equal..Please Check you entry..", vbInformation, "System Message"
            Cancel = 1
            End If
        ElseIf q = vbCancel Then
        Cancel = 1
        End If
    
End If
End Sub

Private Sub lvButtons_H1_Click()
Unload Me
End Sub
Private Function SaveEntry()
Dim x As Integer
opndbaseFMIS.Execute "Update tblAMIS_AccoutingEntries set actioncode = 2,datetimeentered = '" & Trim(datetimeentered) & "," & Format(Now, "yyyy/mm/dd hh:mm:ss AMPM") & "',userid = '" & Trim(UserID) & "," & Trim(ActiveUserID) & "' where reffno = '" & reff & "'"
     For x = 1 To MSFlexGrid1.Rows - 1
        If MSFlexGrid1.TextMatrix(x, 2) <> "TOTAL" Then
            If MSFlexGrid1.TextMatrix(x, 1) <> "" Then
                If MSFlexGrid1.TextMatrix(x, 3) <> "" Or MSFlexGrid1.TextMatrix(x, 4) <> "" Then
                    opndbaseFMIS.Execute "Insert Into [tblAMIS_AccoutingEntries] (reffNo,debit,credit,ChildAccountcode,actioncode,datetimeentered,userid,transtype) values ('" & reff & "'," & CDbl(IIf(IsNumeric(MSFlexGrid1.TextMatrix(x, 3)) = True, MSFlexGrid1.TextMatrix(x, 3), 0)) & "," & CDbl(IIf(IsNumeric(MSFlexGrid1.TextMatrix(x, 4)) = True, MSFlexGrid1.TextMatrix(x, 4), 0)) & "," & _
                    "'" & MSFlexGrid1.TextMatrix(x, 1) & "',1,'" & Format(Now, "yyyy/mm/dd hh:mm:ss AMPM") & "','" & ActiveUserID & "'," & Transtype & ")"
                End If
            End If
        Else
            Exit For
        End If
    Next x
End Function
Private Function ChkEntry() As Boolean
    ChkEntry = False
        If Damount = Camount And Camount > 0 Then
                ChkEntry = True
        End If
End Function

Private Sub lvButtons_H3_Click()
Picture2.Visible = False
cmbEntry.Visible = False
End Sub

Private Sub lvButtons_H5_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
PopupMenu popup
End Sub

Private Sub MSFlexGrid1_Click()
On Error GoTo bad
    Select Case MSFlexGrid1.col
    Case 1 'AccntCode
        txt_entry.Visible = False
        Picture2.Visible = True
        cmbEntry.Move MSFlexGrid1.CellLeft, MSFlexGrid1.CellTop, MSFlexGrid1.CellWidth
        cmbEntry.Visible = True
        If Len(Trim(MSFlexGrid1.Text)) <> 0 Then
            cmbEntry.Text = MSFlexGrid1.Text
            cmbEntry.SetFocus
        Else
            cmbEntry.Text = ""
            cmbEntry.SetFocus
        End If
    Case 3 To 5 'Debit/Credit
        cmbEntry.Visible = False
        Picture2.Visible = False
        txt_entry.Move MSFlexGrid1.CellLeft, MSFlexGrid1.CellTop, MSFlexGrid1.CellWidth, MSFlexGrid1.CellHeight
        txt_entry.Visible = True
        If Len(Trim(MSFlexGrid1.Text)) <> 0 Then
            txt_entry.Text = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 5)
            
            txt_entry.SelStart = 0
            txt_entry.SelLength = Len(txt_entry.Text)
        Else
            txt_entry.Text = ""
        End If
        txt_entry.SetFocus
    
    Case Else
        txt_entry.Visible = False
        cmbEntry.Visible = False
        Picture2.Visible = False
    End Select
Exit Sub
bad:
MsgBox err.Description
End Sub
Private Function LoadAccountsByFund(ByVal accountcode As String)
Dim ARec As New ADODB.Recordset
Dim x As Integer

    cmbEntry.Clear
    cmbEntry.Visible = False
    FundName = GetFundName(fundmedium)
    ARec.Open ("exec Proc_CodeExplaination @accountcode = '" & accountcode & "'"), opndbaseFMIS, adOpenStatic, adLockOptimistic
    If ARec.RecordCount > 0 Then
        For x = 1 To ARec.RecordCount
            cmbEntry.AddItem ARec![childaccountcode]
            cmbEntry.ItemData(cmbEntry.NewIndex) = ARec!Subcode1
            ARec.MoveNext
        Next x
    End If
    ARec.Close
    Set ARec = Nothing
End Function
Public Function LoadAccountsbySub(ByVal accountcode As String)
Dim ARec As New ADODB.Recordset
Dim x As Integer
Dim xx As Variant
Dim str() As String
Dim lvl As Integer
Dim Code As Long
Dim childcode As String
Dim z
    xx = Split(accountcode, "-")
    str() = Split(accountcode, "-", -1, vbTextCompare)
    lvl = UBound(xx) + 1
    If lvl = 1 Then
        lvl = 0
    End If
    
    Select Case (lvl)
        Case 0
        childcode = str(0)
        Case 2
        childcode = str(0)
        Case 3
        childcode = str(0) & "-" & str(1)
        Case 4
        childcode = str(0) & "-" & str(1) & "-" & str(2)
        Case 5
        childcode = str(0) & "-" & str(1) & "-" & str(2) & "-" & str(3)
        Case 6
        childcode = str(0) & "-" & str(1) & "-" & str(2) & "-" & str(3) & "-" & str(4)
        Case 7
        childcode = str(0) & "-" & str(1) & "-" & str(2) & "-" & str(3) & "-" & str(4) & "-" & str(5)
    End Select
    
    If Right(Trim(accountcode), 1) <> "-" Then
        accountcode = accountcode & "-"
        If lvl <> 0 Then
            lvl = lvl + 1
        Else
            lvl = lvl + 2
        End If
    End If
    
    ListView2.ListItems.Clear
    ARec.Open ("Exec Proc_GetSubCode @find = '" & Trim(txtfind.Text) & "' , @lvl = " & lvl & ",@childcode = '" & accountcode & "'"), opndbaseFMIS, adOpenStatic, adLockOptimistic
        MSHFlexGrid1.Clear
        MSHFlexGrid1.Cols = 3
        MSHFlexGrid1.Rows = 2
        If ARec.RecordCount > 0 Then
            Set MSHFlexGrid1.DataSource = ARec
        End If
        MSHFlexGrid1.TextMatrix(0, 1) = "Code"
        MSHFlexGrid1.TextMatrix(0, 2) = "Explanation"
        MSHFlexGrid1.ColWidth(0) = 0
        MSHFlexGrid1.ColWidth(1) = 700
        MSHFlexGrid1.ColWidth(2) = 6000
    ARec.Close
    Set ARec = Nothing
End Function
Private Function LoadAccountsByName(ByVal accountcode As String, ByVal Condition As String)
Dim ARec As New ADODB.Recordset
Dim x As Integer
    ARec.Open "exec Proc_getNamebychildCode @childaccountcode = '" & accountcode & "', @Condition = '" & Condition & "'", opndbaseFMIS, adOpenStatic
        If ARec.RecordCount > 0 Then
            LoadAccountsByName = ARec!Accountfullname
        inRec = True
        End If
    ARec.Close
    Set ARec = Nothing
End Function

Private Sub MSFlexGrid1_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo bad
If KeyCode = vbKeyDelete Then
    Dim i, x As Integer
   With MSFlexGrid1
    If .Row <> .RowSel Then
        If .RowSel < .Row Then
            x = 0
            For i = .RowSel To .Row
                x = x + 1
            Next i
            
            For i = 1 To x
                If .TextMatrix(.RowSel, 2) <> "TOTAL" Then
                .RemoveItem (.RowSel)
                End If
            Next i
        Else
            x = 0
            For i = .Row To .RowSel
                x = x + 1
            Next i
            
            For i = 1 To x
                If .TextMatrix(.Row, 2) <> "TOTAL" Then
                .RemoveItem (.Row)
                End If
            Next i
        End If
    Else
    If .TextMatrix(.Row, 2) <> "TOTAL" Then
    .RemoveItem (.Row)
    End If
    End If
    Call GetSum
End With
Exit Sub
bad:
If err.Number = 30015 Then

Else
MsgBox err.Description
End If
End If
End Sub

Private Sub MSHFlexGrid1_DblClick()
If MSHFlexGrid1.Rows > 1 Then
    If Trim(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 1)) <> "" Then
        If Right(Trim(cmbEntry.Text), 1) = "-" Then
        cmbEntry.Text = cmbEntry.Text & Trim(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 1))
        Else
            If Len(Trim(cmbEntry.Text)) < 3 Then
            cmbEntry.Text = Trim(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 1))
            Else
            cmbEntry.Text = Trim(cmbEntry.Text) & "-" & Trim(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 1))
            End If
        End If
    End If
End If
txtfind.Text = ""
cmbEntry.SetFocus
End Sub

Private Sub MSHFlexGrid1_KeyPress(KeyAscii As Integer)
Dim accountcode As String
If KeyAscii = 13 Then
    ifCMB = False
    If Chckentry() = True Then
         If Trim(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 1)) <> "" Then
            If Right(Trim(cmbEntry.Text), 1) = "-" Then
                MSFlexGrid1.TextMatrix((MSFlexGrid1.Rows - 1), 1) = cmbEntry.Text & Trim(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 1))
                MSFlexGrid1.TextMatrix((MSFlexGrid1.Rows - 1), 2) = LoadAccountsByName(cmbEntry.Text & Trim(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 1)), "Summary")
            Else
               If Len(Trim(cmbEntry.Text)) < 3 Then
                MSFlexGrid1.TextMatrix((MSFlexGrid1.Rows - 1), 1) = Trim(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 1))
                MSFlexGrid1.TextMatrix((MSFlexGrid1.Rows - 1), 2) = Trim(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 2))
                Else
                MSFlexGrid1.TextMatrix((MSFlexGrid1.Rows - 1), 1) = cmbEntry.Text & "-" & Trim(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 1))
                MSFlexGrid1.TextMatrix((MSFlexGrid1.Rows - 1), 2) = LoadAccountsByName(cmbEntry.Text & "-" & Trim(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 1)), "Summary")
                End If
            End If
            MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1
            cmbEntry.Move MSFlexGrid1.CellLeft, ((MSFlexGrid1.Rows - 1) * MSFlexGrid1.CellHeight), MSFlexGrid1.CellWidth
            Call GetSum
        End If
    End If
End If
End Sub
Public Function Chckentry() As Boolean
Dim x As Integer
Dim accountcode As String
Chckentry = True

    If cmbEntry.Text <> MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 1) Then
            If ifCMB = True Then
                accountcode = Trim(cmbEntry.Text)
            Else
                If Right(Trim(cmbEntry.Text), 1) = "-" Then
                    accountcode = cmbEntry.Text & Trim(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 1))
                 
                Else
                   If Len(Trim(cmbEntry.Text)) < 3 Then
                        accountcode = Trim(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 1))
                
                   Else
                       accountcode = cmbEntry.Text & "-" & Trim(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 1))
                   End If
                End If
            End If
        For x = 1 To MSFlexGrid1.Rows - 1
            If Trim(MSFlexGrid1.TextMatrix(x, 1)) = accountcode Then
               ' MsgBox "Acocuntcode and Explaination Already on the List..!", vbInformation, "System Message"
                'Chckentry = False
                'Exit For
            End If
        Next x
    End If
End Function

Private Sub Payroll_Click()
Dim name As String

If Right(Trim(cmbEntry.Text), 1) = "-" Then
   name = LoadAccountsByName(Left(Trim(cmbEntry.Text), Len(Trim(cmbEntry.Text)) - 1), "Fullname")
Else
   name = LoadAccountsByName(Trim(cmbEntry.Text), "Fullname")
End If

            
                If Trim(name) = "" Then
                    MsgBox "Invalid Accountcode Please Select Another Account..!", vbCritical, "System Information"
                    Exit Sub
                End If
    With frmSub2
    .accountcode = cmbEntry.Text
    .Accountname = LoadAccountsByName(cmbEntry.Text, "Fullname")
    .Show 1
    End With
End Sub

Private Sub Property_Click()
frmProperty.Show 1
End Sub

Private Sub txt_entry_Change()
txtformula.Text = txt_entry.Text
End Sub

Private Sub txt_entry_KeyPress(KeyAscii As Integer)
Dim tamount As Currency
 On Error GoTo bad
    If KeyAscii = 13 Then
            
            If IsNumeric(txt_entry.Text) = False And txt_entry.Text <> "" Then
                If InStr(1, txt_entry.Text, "+") = 0 Then
                    MsgBox "None Numeric Entry, Please Check Your Entry", vbCritical, "System Message"
                    Exit Sub
                End If
            End If
                tamount = sumAmount(txt_entry.Text)
                MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 5) = txt_entry.Text
                MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, MSFlexGrid1.col) = IIf((tamount = 0), "", Format((tamount), "#,##0.00"))
                txt_entry.Visible = False
            
            If MSFlexGrid1.col = 3 Then
                If MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 3) <> "" Then
                    MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 4) = ""
                End If
            ElseIf MSFlexGrid1.col = 4 Then
                If MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 4) <> "" Then
                    MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 3) = ""
                End If
            End If
        Call GetSum
        isEdit = True
    End If
Exit Sub
bad:
    Call LoadErr(err.Number, err.Source & ", " & Me.name & ", " & Me.Caption, err.Description)
End Sub
Public Sub GetSum()
Dim x As Integer
    Damount = 0
    Camount = 0
    For x = 1 To MSFlexGrid1.Rows - 1
        If MSFlexGrid1.TextMatrix(x, 1) <> "" Then
            Damount = Damount + CCur(IIf(MSFlexGrid1.TextMatrix(x, 3) = "", 0, MSFlexGrid1.TextMatrix(x, 3)))
            Camount = Camount + CCur(IIf(MSFlexGrid1.TextMatrix(x, 4) = "", 0, MSFlexGrid1.TextMatrix(x, 4)))
        Else

           ' MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1
            MSFlexGrid1.TextMatrix(x, 2) = "TOTAL"
            MSFlexGrid1.TextMatrix(x, 3) = IIf((Damount = 0), "", Format(Damount, "#,##0.00"))
            MSFlexGrid1.TextMatrix(x, 4) = IIf((Camount = 0), "", Format(Camount, "#,##0.00"))
            Exit For
        End If
    Next x
End Sub
Private Function whatAction()

End Function

Private Sub txtfind_Change()
    If Len(Trim(cmbEntry.Text)) >= 3 Then
        LoadAccountsbySub (cmbEntry.Text)
        txtdetails.Text = LoadAccountsByName(cmbEntry.Text, "Fullname")
    Else
    Call GetAccountNamebyorder("Accountcode")
    End If
End Sub

Private Sub txtfind_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    ifCMB = False
    If Chckentry() = True Then
         If Trim(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 1)) <> "" Then
            If Right(Trim(cmbEntry.Text), 1) = "-" Then
                MSFlexGrid1.TextMatrix((MSFlexGrid1.Rows - 1), 1) = cmbEntry.Text & Trim(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 1))
                MSFlexGrid1.TextMatrix((MSFlexGrid1.Rows - 1), 2) = LoadAccountsByName(cmbEntry.Text & Trim(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 1)), "Summary")
            Else
               If Len(Trim(cmbEntry.Text)) < 3 Then
                MSFlexGrid1.TextMatrix((MSFlexGrid1.Rows - 1), 1) = Trim(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 1))
                MSFlexGrid1.TextMatrix((MSFlexGrid1.Rows - 1), 2) = Trim(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 2))
                Else
                MSFlexGrid1.TextMatrix((MSFlexGrid1.Rows - 1), 1) = cmbEntry.Text & "-" & Trim(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 1))
                MSFlexGrid1.TextMatrix((MSFlexGrid1.Rows - 1), 2) = LoadAccountsByName(cmbEntry.Text & "-" & Trim(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 1)), "Summary")
                End If
            End If
            MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1
            cmbEntry.Move MSFlexGrid1.CellLeft, ((MSFlexGrid1.Rows - 1) * MSFlexGrid1.CellHeight), MSFlexGrid1.CellWidth
            Call GetSum
        End If
    End If
End If
End Sub

Private Sub txtformula_Change()
txt_entry.Text = txtformula.Text
End Sub

Private Sub txtformula_KeyPress(KeyAscii As Integer)
Dim tamount As Currency
 On Error GoTo bad
    If KeyAscii = 13 Then
            If IsNumeric(txt_entry.Text) = False And txt_entry.Text <> "" Then
                 If InStr(1, txt_entry.Text, "+") = 0 Then
                    MsgBox "None Numeric Entry, Please Check Your Entry", vbCritical, "System Message"
                    Exit Sub
                End If
            End If
                tamount = sumAmount(txt_entry.Text)
                MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 5) = txt_entry.Text
                MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, MSFlexGrid1.col) = Format((tamount), "#,##0.00")
                txt_entry.Visible = False
            
            If MSFlexGrid1.col = 3 Then
                If MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 3) <> "" Then
                    MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 4) = ""
                End If
            ElseIf MSFlexGrid1.col = 4 Then
                If MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 4) <> "" Then
                    MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 3) = ""
                End If
            End If
        Call GetSum
        isEdit = True
    End If
Exit Sub
bad:
    Call LoadErr(err.Number, err.Source & ", " & Me.name & ", " & Me.Caption, err.Description)
End Sub

VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{1693405E-2DC9-4248-B52F-4AC9145DA2AF}#1.0#0"; "WinXPC Engine.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frm_PPE 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Import from Property"
   ClientHeight    =   8910
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10725
   Icon            =   "frm_PPE.frx":0000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8910
   ScaleWidth      =   10725
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6780
      Left            =   120
      TabIndex        =   15
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Frame Frame2 
      Height          =   975
      Left            =   1800
      TabIndex        =   9
      Top             =   120
      Width           =   3135
      Begin VB.TextBox txtARE 
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
         Left            =   1440
         TabIndex        =   11
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label8 
         Caption         =   "Enter ARE no."
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
         TabIndex        =   10
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.TextBox txtgamount 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   5760
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   8280
      Width           =   3375
   End
   Begin lvButton.lvButtons_H lvButtons_H1 
      Height          =   615
      Left            =   7320
      TabIndex        =   5
      Top             =   360
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1085
      Caption         =   "&Add to Journal"
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
      Image           =   "frm_PPE.frx":076A
      cBack           =   16777215
   End
   Begin WinXPC_Engine.WindowsXPC WindowsXPC1 
      Left            =   120
      Top             =   9000
      _ExtentX        =   6588
      _ExtentY        =   1085
      ColorScheme     =   2
      Common_Dialog   =   0   'False
   End
   Begin VB.Frame Frame1 
      Caption         =   "Normal Balance"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   8520
      TabIndex        =   2
      Top             =   120
      Width           =   2055
      Begin VB.OptionButton OptCredit 
         Caption         =   "Credit"
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
         Left            =   1080
         TabIndex        =   4
         Top             =   480
         Width           =   855
      End
      Begin VB.OptionButton OptDebit 
         Caption         =   "Debit"
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
         TabIndex        =   3
         Top             =   480
         Width           =   855
      End
   End
   Begin lvButton.lvButtons_H lvButtons_H4 
      Height          =   495
      Left            =   9360
      TabIndex        =   0
      Top             =   8280
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
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
      Image           =   "frm_PPE.frx":0ABC
      ImgSize         =   24
      cBack           =   -2147483633
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   6735
      Left            =   1800
      ScaleHeight     =   6705
      ScaleWidth      =   8760
      TabIndex        =   1
      Top             =   1200
      Width           =   8790
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   6710
         Left            =   0
         TabIndex        =   12
         Top             =   0
         Width           =   8745
         _ExtentX        =   15425
         _ExtentY        =   11827
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
   End
   Begin VB.PictureBox freeSizer1 
      Height          =   480
      Left            =   7320
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   14
      Top             =   9840
      Width           =   1200
   End
   Begin lvButton.lvButtons_H lvButtons_H2 
      Height          =   615
      Left            =   5040
      TabIndex        =   8
      Top             =   360
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1085
      Caption         =   "&Add"
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
      Image           =   "frm_PPE.frx":45C6
      cBack           =   16777215
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   8400
      Visible         =   0   'False
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin lvButton.lvButtons_H lvButtons_H3 
      Height          =   615
      Left            =   6180
      TabIndex        =   17
      Top             =   360
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1085
      Caption         =   "&Delete"
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
      Image           =   "frm_PPE.frx":80D0
      cBack           =   16777215
   End
   Begin VB.Label Label1 
      Caption         =   "ARE List"
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
      TabIndex        =   16
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "Total Amount"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4080
      TabIndex        =   7
      Top             =   8400
      Width           =   1575
   End
End
Attribute VB_Name = "frm_PPE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public accountcode, Accountname, reff As String
Public RC As Long

Private Sub Form_Load()
Loaddata
loadARE
End Sub

Private Sub lvButtons_H1_Click()
Dim x As Integer
Dim xx As Variant
Dim str() As String
Dim lvl As Integer
Dim Code As Long
Dim childcode As String
Dim z
    
   
    If MsgBox("Are you sure do you want to add into Journal?", vbInformation + vbYesNo, "System Message") = vbYes Then
     ProgressBar1.Visible = True
     ProgressBar1.Max = MSHFlexGrid1.Rows - 1
        Me.Visible = False
        For x = 1 To MSHFlexGrid1.Rows - 1
'        childcode = Trim(txtaccountcode.Text) & "-" & Trim(MSHFlexGrid1.TextMatrix(x, 1))
        xx = Split(Trim(childcode), "-")
        str() = Split(Trim(childcode), "-", -1, vbTextCompare)
        lvl = UBound(xx) + 1
    
        
            If Trim(MSHFlexGrid1.TextMatrix(x, 1)) <> "" Or Trim(MSHFlexGrid1.TextMatrix(x, 2)) <> "" Then
                With frmSub3
                    .Picture2.Visible = False
                    .cmbEntry.Visible = False
                    If OptCredit.Value = True Then
                    Credit = Trim(MSHFlexGrid1.TextMatrix(x, 3))
                    Debit = 0
                    Else
                    Credit = 0
                    Debit = Trim(MSHFlexGrid1.TextMatrix(x, 3))
                    End If
                    opndbaseFMIS.Execute "Insert into tblAMIs_tmpjournal (Dvno,Accountcode,Debit,Credit) values ('" & Trim(reff) & "','" & Trim(MSHFlexGrid1.TextMatrix(x, 1)) & "','" & Debit & "','" & Credit & "')"
                End With
                DoEvents
                ProgressBar1.Value = x
            End If
        Next x
    End If
    Call frmSub3.GetSum
    ProgressBar1.Visible = False
    Unload Me
End Sub
Private Sub SetGrid()
Dim cc As Integer
    MSHFlexGrid1.TextMatrix(0, 0) = "ID"
    MSHFlexGrid1.TextMatrix(0, 1) = "Acocuntcode"
    MSHFlexGrid1.TextMatrix(0, 2) = "Name"
    MSHFlexGrid1.TextMatrix(0, 3) = "Amount"
    
    
    MSHFlexGrid1.ColWidth(0) = 0
    MSHFlexGrid1.ColWidth(1) = 1700
    MSHFlexGrid1.ColWidth(2) = 5500
    MSHFlexGrid1.ColWidth(3) = 1500
    MSHFlexGrid1.ColAlignment(1) = 1
    MSHFlexGrid1.ColWidth(4) = 0
    MSHFlexGrid1.ColWidth(5) = 0
End Sub
Private Function LoadAccountsByName(ByVal accountcode As String, ByVal Condition As String)
Dim ARec As New ADODB.Recordset
Dim x As Integer
    Set ARec = opndbaseFMIS.Execute("exec Proc_getNamebychildCode @childaccountcode = '" & accountcode & "', @Condition = '" & Condition & "'")
        If ARec.RecordCount > 0 Then
            LoadAccountsByName = ARec!Accountfullname
        inRec = True
        End If
    ARec.Close
    Set ARec = Nothing
End Function
Private Sub lvButtons_H2_Click()
Dim x As Integer
For x = 0 To List1.ListCount
    If UCase(txtARE.Text) = List1.List(x) Then
        MsgBox "Oops! ARE number Already exist...Please Check it..", vbCritical, "System Message"
        Exit Sub
    End If
Next x
opndbaseFMIS.Execute "Insert into tblAMIS_DvnoAndPPE (reffno,AREno) values ('" & reff & "','" & txtARE.Text & "')"
Loaddata
loadARE
End Sub
Private Sub Loaddata()
Dim rec As New ADODB.Recordset
Dim x, y As Integer
'On Error GoTo bad

    Set rec = opndbaseFMIS.Execute("EXEc [fmis].[dbo].[GetJournalfromPropertyByreff] @REFFNO = '" & reff & "'")
        MSHFlexGrid1.Clear
        'MSHFlexGrid1.FixedRows = 1
        MSHFlexGrid1.Rows = 1
        MSHFlexGrid1.Cols = 4
        If rec.RecordCount > 0 Then
       '
            With rec
                For x = 1 To rec.RecordCount
                    MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
                    opndbaseFMIS.Execute "Execute Proc_ChkAlrdyInCOAfromPPE @Articlename = '" & !Articlename & "',@ClassName = '" & !Classification & "',@chldcode = '" & !accntcode & "'"
                    MSHFlexGrid1.TextMatrix(x, 1) = ExecFunction("SELECT [fmis].[dbo].[GetAccntcodebyName] ('" & !Articlename & "','" & !accntcode & "',3)")
                    MSHFlexGrid1.TextMatrix(x, 2) = !Articlename
                    MSHFlexGrid1.TextMatrix(x, 3) = !Untvalue
                    rec.MoveNext
                Next x
            End With
            
        
       ' Set MSHFlexGrid1.DataSource = rec
        End If
'    MSHFlexGrid1.FixedRows = 1
    Call SetGrid
    Call GettotalAMount
rec.Close
Exit Sub
bad:
If err.Number = 3704 Then
    MsgBox "No Record Found", vbInformation, "System Message"
End If
End Sub

Private Function GettotalAMount()
Dim x As Integer
Dim Gamount As Currency
Gamount = 0
For x = 1 To MSHFlexGrid1.Rows - 1
    Gamount = Gamount + CCur(IIf((MSHFlexGrid1.TextMatrix(x, 3) = ""), "", MSHFlexGrid1.TextMatrix(x, 3)))
    MSHFlexGrid1.TextMatrix(x, 3) = Format(CCur(MSHFlexGrid1.TextMatrix(x, 3)), "#,##0.00")
Next x
txtgamount.Text = Format(Gamount, "#,##0.00")
End Function

Private Sub lvButtons_H3_Click()
If MsgBox("Are you sure do you want to Remove the ARE number?", vbInformation + vbYesNo, "System Confirmation") = vbYes Then
    opndbaseFMIS.Execute "Delete from tblAMIS_DvnoAndPPE where reffno = '" & reff & "' and areno = '" & List1.Text & "'"
    Loaddata
    loadARE
End If
End Sub
Private Sub loadARE()
Dim rec As New ADODB.Recordset
Dim x As Integer
Set rec = opndbaseFMIS.Execute("Select AREno from tblAMIS_DvnoAndPPE where reffno = '" & reff & "'")
List1.Clear
If rec.RecordCount > 0 Then
    For x = 1 To rec.RecordCount
        List1.AddItem Trim(UCase(rec!AREno))
        rec.MoveNext
    Next x
End If
End Sub


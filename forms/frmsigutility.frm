VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frm_Signatory 
   Caption         =   "Utility"
   ClientHeight    =   7380
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8985
   Icon            =   "frmsigutility.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7380
   ScaleWidth      =   8985
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ImageList itb32x32 
      Left            =   8280
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   17
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsigutility.frx":076A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsigutility.frx":20FC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsigutility.frx":3A8E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsigutility.frx":5420
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsigutility.frx":6DB2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsigutility.frx":8744
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsigutility.frx":A0D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsigutility.frx":BA68
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsigutility.frx":D3FA
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsigutility.frx":ED8E
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsigutility.frx":FA6A
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsigutility.frx":1034A
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsigutility.frx":11026
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsigutility.frx":11D02
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsigutility.frx":129DE
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsigutility.frx":136BA
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsigutility.frx":14396
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6495
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   11456
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      OLEDropMode     =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Name"
      TabPicture(0)   =   "frmsigutility.frx":14C72
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label4"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label5"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label6"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label7"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label8"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "ListView1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "txtname"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Command1"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txtid"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "txtposition"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "txtext"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "cmb_type"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).ControlCount=   12
      TabCaption(1)   =   "Type of Signatory"
      TabPicture(1)   =   "frmsigutility.frx":15AC4
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txtdesc"
      Tab(1).Control(1)=   "txttype"
      Tab(1).Control(2)=   "lsttype"
      Tab(1).Control(3)=   "Label3"
      Tab(1).Control(4)=   "Label2"
      Tab(1).Control(5)=   "Label1"
      Tab(1).ControlCount=   6
      Begin VB.ComboBox cmb_type 
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
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   2040
         Width           =   3255
      End
      Begin VB.TextBox txtext 
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
         Left            =   1200
         TabIndex        =   16
         Top             =   960
         Width           =   1935
      End
      Begin VB.TextBox txtposition 
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
         Left            =   1200
         TabIndex        =   14
         Top             =   1440
         Width           =   3255
      End
      Begin VB.TextBox txtid 
         Height          =   375
         Left            =   4920
         TabIndex        =   12
         Top             =   480
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.CommandButton Command1 
         Caption         =   "..."
         Height          =   375
         Left            =   4560
         TabIndex        =   11
         Top             =   480
         Width           =   375
      End
      Begin VB.TextBox txtname 
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
         Left            =   1200
         TabIndex        =   10
         Top             =   480
         Width           =   3255
      End
      Begin VB.TextBox txtdesc 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   -73680
         TabIndex        =   7
         Top             =   960
         Width           =   3975
      End
      Begin VB.TextBox txttype 
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
         Left            =   -73680
         TabIndex        =   5
         Top             =   530
         Width           =   3975
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   3495
         Left            =   120
         TabIndex        =   2
         Top             =   2880
         Width           =   8535
         _ExtentX        =   15055
         _ExtentY        =   6165
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
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "trnno"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "ID"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Name"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Position"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Type of Sinatory"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView lsttype 
         Height          =   3015
         Left            =   -74880
         TabIndex        =   3
         Top             =   3360
         Width           =   8535
         _ExtentX        =   15055
         _ExtentY        =   5318
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
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "ID"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Type Name"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Description"
            Object.Width           =   5292
         EndProperty
      End
      Begin VB.Label Label8 
         Caption         =   "Type of Signatory"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   18
         Top             =   1920
         Width           =   975
      End
      Begin VB.Label Label7 
         Caption         =   "Record(s)"
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
         Top             =   2520
         Width           =   1815
      End
      Begin VB.Label Label6 
         Caption         =   "Extension:"
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
         TabIndex        =   15
         Top             =   1035
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "Position:"
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
         TabIndex        =   13
         Top             =   1515
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "Name:"
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
         TabIndex        =   9
         Top             =   555
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Record(s)"
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
         Left            =   -74880
         TabIndex        =   8
         Top             =   3000
         Width           =   1815
      End
      Begin VB.Label Label2 
         Caption         =   "Type Description:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   -74880
         TabIndex        =   6
         Top             =   1035
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Type Name:"
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
         Left            =   -74880
         TabIndex        =   4
         Top             =   600
         Width           =   1215
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Height          =   810
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   10560
      _ExtentX        =   18627
      _ExtentY        =   1429
      ButtonWidth     =   1058
      ButtonHeight    =   1429
      Appearance      =   1
      Style           =   1
      ImageList       =   "itb32x32"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "New"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Add"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Delete"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Close"
            ImageIndex      =   7
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frm_Signatory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Timer1_Timer()

End Sub

Private Sub Command1_Click()
ActiveFormCaller = Me.name
frmCDClaimantRegistry.Frame1.Enabled = False
frmCDClaimantRegistry.Show 1
frmCDClaimantRegistry.Frame1.Enabled = True
txtposition.Text = GetEmpPosition(txtid.Text)
End Sub

Private Sub Form_Load()
LoadSigRecords
GEtTypeOfSig
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

Select Case Button

Case "New"
clr
Case "Add"
        Select Case SSTab1.Tab
        Case 0
            If txtname.Text = "" Or txtid.Text = "" Or txtposition.Text = "" Or cmb_type.Text = "" Then
            MsgBox "Please Complete the Parameter to Save the Entry..", vbInformation, "System Message"
            Exit Sub
            End If
            opndbaseFMIS.Execute "Insert into tblReff_Signatory([id],[fullname],position,typsig) values ('" & txtid.Text & "','" & Trim(txtname.Text) & "','" & txtposition.Text & "','" & cmb_type.Text & "') "
            MsgBox "Successfully Save..", vbInformation, "System Message"
            clr
        Case 1
            If txttype.Text = "" Or txtdesc.Text = "" Then
            MsgBox "Please Complete the Parameter to Save the Entry..", vbInformation, "System Message"
            Exit Sub
            End If
            opndbaseFMIS.Execute "Insert into tblReff_TypeOfSig([typName],[Description]) values ('" & txttype.Text & "','" & txtdesc.Text & "') "
            MsgBox "Successfully Save..", vbInformation, "System Message"
            clr
        End Select
Case "Delete"
        Select Case SSTab1.Tab
        Case 0
            If MsgBox("Are you sure do want to delete?", vbCritical + vbYesNo, "System Message") = vbYes Then
                opndbaseFMIS.Execute "Delete from tblReff_Signatory where trnno = '" & ListView1.SelectedItem.Text & "'"
            End If
        Case 1
            If MsgBox("Are you sure do want to delete?", vbCritical + vbYesNo, "System Message") = vbYes Then
                    opndbaseFMIS.Execute "Delete from tblReff_TypeOfSig where trnno = '" & lsttype.SelectedItem.Text & "'"
            End If
        End Select
End Select
LoadSigRecords
GEtTypeOfSig
End Sub
Private Function GetEmpPosition(ByVal id As String)
Dim rec As New ADODB.Recordset
rec.Open "Select position from pmis.dbo.employee where swipemployeeid = '" & id & "'", opndbaseFMIS, adOpenStatic, adLockOptimistic
    If rec.RecordCount > 0 Then
        GetEmpPosition = Trim(rec!Position)
    End If
rec.Close
End Function
Private Function GEtTypeOfSig()
Dim rec As New ADODB.Recordset
Dim x As Integer
rec.Open "Select * from tblReff_TypeOfSig", opndbaseFMIS, adOpenStatic, adLockOptimistic
cmb_type.Clear
If rec.RecordCount > 0 Then
    For x = 1 To rec.RecordCount
        cmb_type.AddItem (rec!TypName)
        cmb_type.ItemData(cmb_type.NewIndex) = rec!Trnno
        rec.MoveNext
    Next x
End If
rec.Close
End Function
Private Function LoadSigRecords()
Dim rec As New ADODB.Recordset
Dim x As Integer
Dim z
Select Case SSTab1.Tab
Case 0
    ListView1.ListItems.Clear
    rec.Open "Select * from tblReff_Signatory", opndbaseFMIS, adOpenStatic, adLockOptimistic
        If rec.RecordCount > 0 Then
        For x = 1 To rec.RecordCount
            Set z = ListView1.ListItems.Add(, , rec!Trnno)
            z.SubItems(1) = rec!id
            z.SubItems(2) = rec!FullName
            z.SubItems(3) = rec!Position
            z.SubItems(4) = rec!typSig
            rec.MoveNext
        Next x
        End If
    rec.Close
Case 1
    lsttype.ListItems.Clear
    rec.Open "Select * from tblReff_TypeOfSig", opndbaseFMIS, adOpenStatic, adLockOptimistic
        If rec.RecordCount > 0 Then
        For x = 1 To rec.RecordCount
            Set z = lsttype.ListItems.Add(, , rec!Trnno)
            z.SubItems(1) = rec!TypName
            z.SubItems(2) = rec!Description
            rec.MoveNext
        Next x
        End If
    rec.Close
End Select

End Function
Private Function clr()
Select Case SSTab1.Tab
Case 0
txtid.Text = ""
txtname.Text = ""
txtposition.Text = ""
txtext.Text = ""
Case 1
txtdesc.Text = ""
txttype.Text = ""
End Select
End Function

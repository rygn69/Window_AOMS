VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frm_UtilityConnection 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Manage Connection"
   ClientHeight    =   6705
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9345
   Icon            =   "frm_UtilityConnection.frx":0000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6705
   ScaleWidth      =   9345
   Begin VB.Frame Frame2 
      Caption         =   "Entries"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   120
      TabIndex        =   11
      Top             =   4080
      Width           =   9135
      Begin MSComctlLib.ListView ListView1 
         Height          =   1575
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   2778
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
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "ID"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Name"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Description"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Connection String"
            Object.Width           =   5292
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Connection Details"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3855
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   9135
      Begin VB.TextBox txtconnString 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1620
         IMEMode         =   3  'DISABLE
         Left            =   2400
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   9
         Top             =   2040
         Width           =   6615
      End
      Begin VB.TextBox txtconndesc 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1020
         Left            =   2400
         TabIndex        =   7
         Top             =   840
         Width           =   3975
      End
      Begin VB.TextBox txtConName 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   2400
         TabIndex        =   5
         Top             =   240
         Width           =   3975
      End
      Begin VB.Label Label3 
         Caption         =   "Connection String:"
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
         Left            =   120
         TabIndex        =   10
         Top             =   2040
         Width           =   2055
      End
      Begin VB.Label Label2 
         Caption         =   "Connection Desc:"
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
         Left            =   120
         TabIndex        =   8
         Top             =   840
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Connection Name:"
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
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   2055
      End
   End
   Begin lvButton.lvButtons_H lvButtons_H3 
      Height          =   495
      Left            =   5040
      TabIndex        =   0
      Top             =   6120
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      Caption         =   "&Save"
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
      Image           =   "frm_UtilityConnection.frx":076A
      ImgSize         =   24
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdupdate 
      Height          =   495
      Left            =   3600
      TabIndex        =   1
      Top             =   6120
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      Caption         =   "&New"
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
      Image           =   "frm_UtilityConnection.frx":0ABC
      ImgSize         =   24
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H lvButtons_H1 
      Height          =   495
      Left            =   6480
      TabIndex        =   2
      Top             =   6120
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
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
      cGradient       =   0
      CapStyle        =   1
      Mode            =   0
      Value           =   0   'False
      Image           =   "frm_UtilityConnection.frx":170E
      ImgSize         =   24
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H lvButtons_H2 
      Height          =   495
      Left            =   7920
      TabIndex        =   3
      Top             =   6120
      Width           =   1335
      _ExtentX        =   2355
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
      cGradient       =   0
      CapStyle        =   1
      Mode            =   0
      Value           =   0   'False
      Image           =   "frm_UtilityConnection.frx":5218
      ImgSize         =   24
      cBack           =   -2147483633
   End
End
Attribute VB_Name = "frm_UtilityConnection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Trnno As Integer

Private Sub cmdupdate_Click()
txtConName.Text = ""
txtconndesc.Text = ""
txtconnString.Text = ""
Trnno = 0
LoadData
End Sub

Private Sub Form_Load()
LoadData
End Sub

Private Sub ListView1_Click()
Trnno = 0
txtConName.Text = ""
txtconndesc.Text = ""
txtconnString.Text = ""
If ListView1.ListItems.Count <> 0 Then
    With ListView1
        Trnno = .SelectedItem.Text
        txtConName.Text = Trim(.SelectedItem.ListSubItems(1).Text)
        txtconndesc.Text = Trim(.SelectedItem.ListSubItems(2).Text)
        txtconnString.Text = MPDll.DecryptString(Trim(.SelectedItem.ListSubItems(3).Text))
    End With
End If
End Sub

Private Sub lvButtons_H1_Click()
If MsgBox("Are you sure do you want to Delete this entry?", vbCritical + vbYesNo, "System Confirmation") = vbYes Then
    opndbaseFMIS.Execute "delete from tblreff_ManageConnection where trnno = " & Trnno & ""
    MsgBox "Delete Successfully...!", vbInformation, "System Message"
End If
End Sub

Private Sub lvButtons_H2_Click()
Unload Me
End Sub

Private Sub lvButtons_H3_Click()
Dim rec As New ADODB.Recordset
If Trim(txtConName.Text) = "" Or txtconndesc.Text = "" Or txtconnString.Text = "" Then
    MsgBox "Complete the Fields to Proceed the transaction", vbInformation, "System Message"
    Exit Sub
End If

rec.Open "Select * from tblreff_ManageConnection where trnno = " & Trnno & "", opndbaseFMIS, adOpenStatic, adLockOptimistic
If rec.RecordCount > 0 Then
    If MsgBox("Are you sure do you want to update the Connection String?", vbInformation + vbYesNo, "System Confirmation") = vbYes Then
        opndbaseFMIS.Execute "Update tblreff_ManageConnection set Name = '" & Replace(Trim(txtConName.Text), "'", "''") & "',description = '" & Replace(Trim(txtconndesc.Text), "'", "''") & "',String = '" & Replace(Trim(txtconnString.Text), "'", "''") & "'"
        MsgBox "Update Successfully...!", vbInformation, "System Message"
    End If
Else
    If MsgBox("Are you Sure do you want to Save?", vbInformation + vbYesNo, "System Confirmation") = vbYes Then
       opndbaseFMIS.Execute "insert into  tblreff_ManageConnection (Name,description,string) values ('" & Replace(Trim(txtConName.Text), "'", "''") & "','" & Replace(Trim(txtconndesc.Text), "'", "''") & "','" & Replace(Trim(txtconnString.Text), "'", "''") & "')"
       MsgBox "Save Successfully...!", vbInformation, "System Message"
    End If
End If
LoadData
txtConName.Text = ""
txtconndesc.Text = ""
txtconnString.Text = ""
rec.Close
Set rec = Nothing
End Sub
Private Function LoadData()
Dim rec As New ADODB.Recordset
Dim x As Integer
Dim z
rec.Open "Select * from tblreff_ManageConnection order by name", opndbaseFMIS, adOpenStatic, adLockOptimistic
    If rec.RecordCount > 0 Then
    ListView1.ListItems.Clear
    With ListView1
        For x = 1 To rec.RecordCount
        
            Set z = .ListItems.Add(, , rec!Trnno)
                z.SubItems(1) = Trim(rec!name)
                z.SubItems(2) = Trim(rec!Description)
                z.SubItems(3) = MPDll.EncryptString(Trim(rec!String))
            rec.MoveNext
        Next x
    End With
    End If
rec.Close
Set rec = Nothing
End Function

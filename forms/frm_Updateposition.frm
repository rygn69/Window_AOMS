VERSION 5.00
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frm_Updateposition 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Update Position"
   ClientHeight    =   5415
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9030
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5415
   ScaleWidth      =   9030
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "Encrypt/Decrypt String"
      Height          =   1815
      Left            =   240
      TabIndex        =   7
      Top             =   3360
      Width           =   6015
      Begin VB.TextBox txtencrypt 
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
         TabIndex        =   10
         Top             =   480
         Width           =   4215
      End
      Begin VB.TextBox txtdecrypt 
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
         TabIndex        =   9
         Top             =   1200
         Width           =   4215
      End
      Begin lvButton.lvButtons_H lvButtons_H1 
         Height          =   375
         Left            =   4440
         TabIndex        =   8
         Top             =   480
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         Caption         =   "Encrypt"
         CapAlign        =   2
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
         cBack           =   -2147483633
      End
      Begin lvButton.lvButtons_H lvButtons_H2 
         Height          =   375
         Left            =   4440
         TabIndex        =   11
         Top             =   1200
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         Caption         =   "Decrypt"
         CapAlign        =   2
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
         cBack           =   -2147483633
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Change"
      Height          =   495
      Left            =   7080
      TabIndex        =   5
      Top             =   2400
      Width           =   1695
   End
   Begin VB.ComboBox Combo2 
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
      Left            =   240
      TabIndex        =   3
      Text            =   "Combo1"
      Top             =   1800
      Width           =   8535
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
      Left            =   240
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   720
      Width           =   8535
   End
   Begin VB.Label Label3 
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
      Left            =   6600
      TabIndex        =   6
      Top             =   360
      Width           =   2175
   End
   Begin VB.Label Label2 
      Caption         =   "Position Name"
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
      Left            =   240
      TabIndex        =   4
      Top             =   1440
      Width           =   2175
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Change to"
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
      Left            =   240
      TabIndex        =   2
      Top             =   1200
      Width           =   8535
   End
   Begin VB.Label sd 
      Caption         =   "Position Name"
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
      Left            =   240
      TabIndex        =   1
      Top             =   360
      Width           =   2175
   End
End
Attribute VB_Name = "frm_Updateposition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If MsgBox("Are You sure Do you want to Update the Position?", vbInformation + vbYesNo, "System Confirmation") = vbYes Then
opndbaseFMIS.Execute ("update [pmis].[dbo].[Employee]  set Position = '" & Combo2.Text & "' where Position = '" & Combo1.Text & "'")
Call LoadpositionFromEmployee
End If
End Sub
Private Sub LoadpositionFromEmployee()
Dim rec As New ADODB.Recordset
Dim x
 Combo1.Clear
Set rec = opndbaseFMIS.Execute("SELECT     Position FROM         pmis.dbo.Employee WHERE     (Cause IN(SELECT     Cause FROM          pmis.dbo.tblEmploymentConnection WHERE      (Active = 1))) AND (LEN(SwipEmployeeID) > 0) and position  not in (SELECT [Pos_Name] FROM [pmis].[dbo].[RefsPositions]) group by Position")
If rec.RecordCount > 0 Then
For x = 1 To rec.RecordCount
    Combo1.AddItem rec!Position
    rec.MoveNext
Next x
Label3.Caption = rec.RecordCount
Combo1.ListIndex = 0
rec.Close
Set rec = Nothing
End If
End Sub
Private Sub LoadpositionFromPosition()
Dim rec As New ADODB.Recordset
Dim x
 Combo2.Clear
Set rec = opndbaseFMIS.Execute("SELECT [Pos_Name]  FROM [pmis].[dbo].[RefsPositions]")
If rec.RecordCount > 0 Then
For x = 1 To rec.RecordCount
    Combo2.AddItem rec!Pos_Name
    rec.MoveNext
Next x
End If
End Sub
Private Sub Form_Load()
'On Error Resume Next
'Call LoadpositionFromEmployee
'Call LoadpositionFromPosition
'If Trim(ActiveUserID) = "8500" Then
'    Frame1.Enabled = True
'    Me.Height = 5850
'End If
End Sub

Private Sub lvButtons_H1_Click()
txtdecrypt.Text = EncryptString(txtencrypt.Text)
End Sub

Private Sub lvButtons_H2_Click()
txtdecrypt.Text = DecryptString(txtencrypt.Text)
End Sub


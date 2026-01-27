VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmUpdate 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Accounting Operation Management System Updater"
   ClientHeight    =   4905
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5715
   Icon            =   "frmUpdate.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4905
   ScaleWidth      =   5715
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "Crrent System Use"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   120
      TabIndex        =   11
      Top             =   1800
      Width           =   5535
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
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
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   360
         Width           =   975
      End
      Begin VB.Label lblCname 
         BackStyle       =   0  'Transparent
         Caption         =   "Accounting Operation Management System"
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
         Left            =   960
         TabIndex        =   16
         Top             =   360
         Width           =   4095
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Version:"
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
         Top             =   720
         Width           =   975
      End
      Begin VB.Label lblCversion 
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
         Height          =   255
         Left            =   960
         TabIndex        =   14
         Top             =   720
         Width           =   4095
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Size:"
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
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label lblCsize 
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
         Height          =   255
         Left            =   960
         TabIndex        =   12
         Top             =   1080
         Width           =   4095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Available System Update"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   5535
      Begin VB.Label lblASize 
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
         Height          =   255
         Left            =   960
         TabIndex        =   10
         Top             =   1080
         Width           =   4095
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Size:"
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
         TabIndex        =   9
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label lblAversion 
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
         Height          =   255
         Left            =   960
         TabIndex        =   8
         Top             =   720
         Width           =   4095
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Version:"
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
         TabIndex        =   7
         Top             =   720
         Width           =   975
      End
      Begin VB.Label lblAname 
         BackStyle       =   0  'Transparent
         Caption         =   "Accounting Operation Management System"
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
         Left            =   960
         TabIndex        =   6
         Top             =   360
         Width           =   4095
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
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
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4560
      TabIndex        =   2
      Top             =   4440
      Width           =   1095
   End
   Begin MSComctlLib.ProgressBar proStat 
      Height          =   420
      Left            =   120
      TabIndex        =   1
      Top             =   3840
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   741
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Update"
      Height          =   375
      Left            =   3360
      TabIndex        =   0
      Top             =   4440
      Width           =   1095
   End
   Begin VB.Label lblStatus 
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
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   3480
      Width           =   5295
   End
End
Attribute VB_Name = "frmUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim FSO As New FileSystemObject
Dim AppPath As String
Dim DEXE, CEXE, REXE As Double



Private Sub Command1_Click()
Dim x As Long
'On Error GoTo bad
Dim nAMEFOLD As String
If lblAversion.Caption <> lblCversion.Caption Then
    If FSO.FolderExists(App.Path & "\BackUp Old Exe") = False Then
        FSO.CreateFolder App.Path & "\BackUp Old Exe"
    End If
    nAMEFOLD = "'" & App.Path & "\BackUp Old Exe'"
    nAMEFOLD = Replace(nAMEFOLD, "'", """")
    Shell ("attrib +h +r +s " & nAMEFOLD)

    If FSO.FileExists(App.Path & "\Accounting Operation Management System.exe") = True Then
        If FSO.FileExists(App.Path & "\BackUp Old Exe\" & App.Path & "\BackUp Old Exe\" & "AOMS " & FSO.GetFileVersion(Trim(App.Path & "\Accounting Operation Management System.exe")) & ".exe") = True Then
        FSO.MoveFile App.Path & "\Accounting Operation Management System.exe", App.Path & "\BackUp Old Exe\" & "AOMS " & FSO.GetFileVersion(Trim(App.Path & "\Accounting Operation Management System.exe")) & ".exe"
        Else
        FSO.DeleteFile App.Path & "\Accounting Operation Management System.exe", True
        End If
    End If
    FSO.CopyFile AppPath, App.Path & "\Accounting Operation Management System.exe", True
    proStat.Max = CDbl(lblASize.Caption)
    For x = 1 To CDbl(lblASize.Caption)
        x = FSO.GetFile(Trim(App.Path & "\Accounting Operation Management System.exe")).Size
    Next x
    Call Form_Load
Else
    MsgBox "Your System is Updated, not necessary to Update.", vbInformation, "System Message"
End If
Exit Sub
bad:
MsgBox Err.Description
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
On Error GoTo bad
Dim rec As New ADODB.Recordset
rec.Open "Select * from tblAMIS_SystemUpdate", opndbaseFMIS, adOpenStatic
    If rec.RecordCount > 0 Then
        AppPath = rec!Path
        lblAname.Caption = FSO.GetFileName(Trim(AppPath))
        lblAversion.Caption = FSO.GetFileVersion(Trim(AppPath))
        lblASize.Caption = Format((FSO.GetFile(Trim(AppPath)).Size), "#,###") & " bytes"
        
        If FSO.FileExists(App.Path & "\Accounting Operation Management System.exe") = True Then
        lblCname.Caption = FSO.GetFileName(Trim(App.Path & "\Accounting Operation Management System.exe"))
        lblCversion.Caption = FSO.GetFileVersion(Trim(App.Path & "\Accounting Operation Management System.exe"))
        lblCsize.Caption = Format(FSO.GetFile(Trim(App.Path & "\Accounting Operation Management System.exe")).Size, "#,###") & " bytes"
        Else
        lblCname.Caption = "Not Available"
        lblCversion.Caption = "Not Available"
        lblCsize.Caption = "Not Available"
        End If
    End If
rec.Close
Exit Sub
bad:
MsgBox Err.Description
End Sub

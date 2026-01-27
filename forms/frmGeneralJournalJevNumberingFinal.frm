VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
Object = "{1693405E-2DC9-4248-B52F-4AC9145DA2AF}#1.0#0"; "WinXPC Engine.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form frmGeneralJournalJevNumbering 
   Caption         =   "JEV Numbering for Liquidation Of Cash Advance Genearl Journal"
   ClientHeight    =   9420
   ClientLeft      =   735
   ClientTop       =   1215
   ClientWidth     =   14415
   Icon            =   "frmGeneralJournalJevNumberingFinal.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   9420
   ScaleWidth      =   14415
   Begin VB.CommandButton Command1 
      Caption         =   "Load Reports"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   1200
      TabIndex        =   15
      Top             =   3015
      Width           =   1125
   End
   Begin VB.Frame Frame5 
      Height          =   1530
      Left            =   2535
      TabIndex        =   13
      Top             =   885
      Width           =   11835
      Begin VB.CommandButton cmd_post 
         Caption         =   "Post (JEV No.)"
         Height          =   1005
         Left            =   1800
         Picture         =   "frmGeneralJournalJevNumberingFinal.frx":0E42
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   240
         Width           =   1755
      End
      Begin VB.CommandButton cmd_Mass 
         Caption         =   "Mass JEV Nos."
         Height          =   1005
         Left            =   120
         Picture         =   "frmGeneralJournalJevNumberingFinal.frx":493C
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   240
         Width           =   1635
      End
      Begin VB.Frame Frame3 
         Caption         =   "CN Details"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1245
         Left            =   3840
         TabIndex        =   18
         Top             =   120
         Width           =   8010
         Begin VB.ComboBox Combo1 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   1140
            Style           =   1  'Simple Combo
            TabIndex        =   27
            Top             =   240
            Width           =   2565
         End
         Begin VB.ComboBox cmb_AccountName 
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
            Left            =   4950
            TabIndex        =   20
            Top             =   1755
            Width           =   2085
         End
         Begin VB.ComboBox Cmb_CnNo 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   1140
            Style           =   1  'Simple Combo
            TabIndex        =   19
            Top             =   750
            Width           =   2565
         End
         Begin MSComCtl2.DTPicker DTPCNdate 
            CausesValidation=   0   'False
            Height          =   390
            Left            =   5340
            TabIndex        =   21
            Top             =   240
            Width           =   2565
            _ExtentX        =   4524
            _ExtentY        =   688
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "mm/dd/yyyy"
            Format          =   87097345
            UpDown          =   -1  'True
            CurrentDate     =   38240
         End
         Begin MSComCtl2.DTPicker DTPRdate 
            CausesValidation=   0   'False
            Height          =   390
            Left            =   5340
            TabIndex        =   22
            Top             =   720
            Width           =   2565
            _ExtentX        =   4524
            _ExtentY        =   688
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "MM/dd/yyyy"
            Format          =   87097345
            UpDown          =   -1  'True
            CurrentDate     =   40583
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "CN No:"
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
            Left            =   360
            TabIndex        =   28
            Top             =   825
            Width           =   1005
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fundtype:"
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
            Left            =   120
            TabIndex        =   26
            Top             =   360
            Width           =   885
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "CN Date:"
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
            Left            =   4440
            TabIndex        =   25
            Top             =   315
            Width           =   1170
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Account Name"
            Height          =   195
            Left            =   3600
            TabIndex        =   24
            Top             =   1875
            Width           =   1065
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Date Received"
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
            Left            =   3840
            TabIndex        =   23
            Top             =   795
            Width           =   1725
         End
      End
      Begin VB.Shape Shape3 
         BackColor       =   &H00000000&
         BackStyle       =   1  'Opaque
         Height          =   1155
         Left            =   60
         Top             =   165
         Width           =   3555
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00000000&
         BackStyle       =   1  'Opaque
         Height          =   945
         Left            =   -15
         Top             =   -1020
         Width           =   11880
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00000000&
      Caption         =   "Special Account"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   780
      Left            =   75
      TabIndex        =   11
      Top             =   1035
      Width           =   2250
      Begin VB.ComboBox cmb_FundType 
         Height          =   315
         Left            =   75
         TabIndex        =   12
         Top             =   300
         Width           =   2100
      End
   End
   Begin VB.TextBox txt_RecordID 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   60
      TabIndex        =   9
      ToolTipText     =   "Type only CN No. then press Enter"
      Top             =   3855
      Width           =   2280
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   2400
      Top             =   4275
   End
   Begin VB.ListBox List2 
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
      Height          =   4350
      Left            =   75
      TabIndex        =   7
      Top             =   4605
      Width           =   2265
   End
   Begin VB.Frame Frame6 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "CN No."
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   60
      TabIndex        =   6
      Top             =   4290
      Width           =   2280
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "For the Period"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   60
      TabIndex        =   3
      Top             =   1920
      Width           =   2265
      Begin MSComCtl2.DTPicker DTPicker1 
         CausesValidation=   0   'False
         Height          =   390
         Left            =   90
         TabIndex        =   4
         Top             =   435
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   688
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "MMMM yyyy"
         Format          =   87097347
         UpDown          =   -1  'True
         CurrentDate     =   38240
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         Height          =   195
         Left            =   1530
         TabIndex        =   5
         Top             =   165
         Visible         =   0   'False
         Width           =   480
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   6600
      Left            =   2535
      ScaleHeight     =   6570
      ScaleWidth      =   11820
      TabIndex        =   1
      Top             =   2475
      Width           =   11850
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grd_details 
         Height          =   6570
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   11820
         _ExtentX        =   20849
         _ExtentY        =   11589
         _Version        =   393216
         FixedCols       =   0
         ForeColorFixed  =   128
         ScrollTrack     =   -1  'True
         FocusRect       =   0
         SelectionMode   =   1
         AllowUserResizing=   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
   Begin WinXPC_Engine.WindowsXPC WindowsXPC1 
      Left            =   9360
      Top             =   9000
      _ExtentX        =   6588
      _ExtentY        =   1085
      ColorScheme     =   2
      Common_Dialog   =   0   'False
      PictureControl  =   0   'False
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   840
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   14415
      _ExtentX        =   25426
      _ExtentY        =   1482
      ButtonWidth     =   2117
      ButtonHeight    =   1429
      Appearance      =   1
      Style           =   1
      ImageList       =   "itb32x32"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Print Report"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "slash"
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Close"
            ImageIndex      =   7
         EndProperty
      EndProperty
      Begin MSComctlLib.ImageList itb32x32 
         Left            =   5520
         Top             =   0
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
               Picture         =   "frmGeneralJournalJevNumberingFinal.frx":8436
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmGeneralJournalJevNumberingFinal.frx":9DC8
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmGeneralJournalJevNumberingFinal.frx":B75A
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmGeneralJournalJevNumberingFinal.frx":D0EC
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmGeneralJournalJevNumberingFinal.frx":EA7E
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmGeneralJournalJevNumberingFinal.frx":10410
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmGeneralJournalJevNumberingFinal.frx":11DA2
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmGeneralJournalJevNumberingFinal.frx":13734
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmGeneralJournalJevNumberingFinal.frx":150C6
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmGeneralJournalJevNumberingFinal.frx":16A5A
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmGeneralJournalJevNumberingFinal.frx":17736
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmGeneralJournalJevNumberingFinal.frx":18016
               Key             =   ""
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmGeneralJournalJevNumberingFinal.frx":18CF2
               Key             =   ""
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmGeneralJournalJevNumberingFinal.frx":199CE
               Key             =   ""
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmGeneralJournalJevNumberingFinal.frx":1A6AA
               Key             =   ""
            EndProperty
            BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmGeneralJournalJevNumberingFinal.frx":1B386
               Key             =   ""
            EndProperty
            BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmGeneralJournalJevNumberingFinal.frx":1C062
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComCtl2.Animation Animation1 
         Height          =   450
         Left            =   11760
         TabIndex        =   31
         Top             =   120
         Width           =   480
         _ExtentX        =   847
         _ExtentY        =   794
         _Version        =   393216
         FullWidth       =   32
         FullHeight      =   30
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1365
      Top             =   9615
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGeneralJournalJevNumberingFinal.frx":1C93E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGeneralJournalJevNumberingFinal.frx":1D9C0
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGeneralJournalJevNumberingFinal.frx":1FAFA
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGeneralJournalJevNumberingFinal.frx":20D7C
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGeneralJournalJevNumberingFinal.frx":23C06
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label14"
      Height          =   195
      Left            =   13755
      TabIndex        =   17
      Top             =   9135
      Width           =   570
   End
   Begin VB.Label Label13 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Label13"
      ForeColor       =   &H0000FF00&
      Height          =   420
      Left            =   90
      TabIndex        =   16
      Top             =   3015
      Width           =   1020
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Search CN No.:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   60
      TabIndex        =   14
      Top             =   3585
      Width           =   1125
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label8"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   105
      TabIndex        =   10
      Top             =   9060
      Width           =   480
   End
   Begin VB.Label Label7 
      Caption         =   "Label7"
      Height          =   210
      Left            =   3540
      TabIndex        =   8
      Top             =   4500
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   9540
      Left            =   -30
      Top             =   825
      Width           =   2445
   End
End
Attribute VB_Name = "frmGeneralJournalJevNumbering"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim tmpAccName As String
Dim FMISNo As String


Private Sub cmb_FundType_Change()
Call SetGrid
List2.Clear
Combo1.Text = ""
cmb_AccountName.Text = ""
Cmb_CnNo.Text = ""
DTPCNdate.Value = Now
DTPRdate.Value = Now
End Sub

Private Sub cmb_FundType_Click()
Call cmb_FundType_Change
End Sub

Private Sub cmd_Mass_Click()
JevOk = False
frmPOstdate.Show 1
If JevOk = True Then
Label13.Caption = "JEV Numbering..."
Label13.Refresh
Animation1.Visible = True
Animation1.Open AViLocation & "\horizontaloading.avi"
Animation1.Play
Call JEVMassNumbering(cmb_FundType.Text)
Animation1.Stop
Animation1.Close
Animation1.Visible = False
Label13.Caption = ""
Else
MsgBox "Cannot Generate the System JEV Number,If you cancel to Set the Date", vbInformation, "System Message"
End If
End Sub
Private Sub JEVMassNumbering(ByVal FundType As String)
Dim opnJEV As New ADODB.Recordset
Dim rec As New ADODB.Recordset
Dim sql As String
Dim cc As Integer
Dim DVNo As String
Dim LastJEVSNno As Long
Dim last As String
    

    rec.Open ("EXEC [dbo].[Proc_GetMaxJevSeries] @transtype = 4,@jevyeardate = '" & Year(DatePost) & "' ,@fundtype = '" & cmb_FundType.Text & "'"), opndbaseFMIS, adOpenStatic, adLockOptimistic
    LastJEVSNno = rec.Fields!MAXJEVSERIES
    rec.Close

For cc = 1 To grd_details.Rows - 1

    DVNo = (grd_details.TextMatrix(cc, 0))
    

'    If grd_details.TextMatrix(cc, 4) = 2 Then
    
        sql = "SELECT tblAMIS_IncomingDVTrns.FundType as FundType, tblAMIS_JournalEntry.TransType as TransType, tblAMIS_JournalEntry.DVNo as DVNo, " & _
                "          tblAMIS_JournalEntry.TransDate as TransDate, tblAMIS_JournalEntry.JEVSeriesNo as JEVSeriesNo,(Select FundCode from tblRefBMS_Funds where FundMedium=tblAMIS_IncomingDVTrns.FundType) as FundCode " & _
                " FROM tblAMIS_IncomingDVTrns INNER JOIN " & _
                "          tblAMIS_JournalEntry ON tblAMIS_IncomingDVTrns.DVNo = tblAMIS_JournalEntry.DVNo " & _
                " Where (tblAMIS_JournalEntry.ActionCode = 1) And (tblAMIS_IncomingDVTrns.ActionCode = 1) " & _
                " GROUP BY tblAMIS_IncomingDVTrns.FundType, tblAMIS_JournalEntry.TransType, tblAMIS_JournalEntry.DVNo, " & _
                "          tblAMIS_JournalEntry.TransDate , tblAMIS_JournalEntry.JEVSeriesNo " & _
                " HAVING   tblAMIS_JournalEntry.DVNo ='" & DVNo & "'"
        opnJEV.Open sql, opndbaseFMIS, adOpenStatic, adLockOptimistic
        If opnJEV.RecordCount <> 0 Then
            'grd_details.TextMatrix(cc, 14) = opnJEV!FundCode & "-" & Right(Year(Date), 2) & "-" & Format(Month(Date), "00") & "-" & Format(opnJEV!TransType, "00") & "-" & LastJEVSNno
            grd_details.TextMatrix(cc, 4) = cmb_FundType.ItemData(cmb_FundType.ListIndex) & "-" & Right(DTPicker1.Year, 2) & "-" & Format((DTPicker1.Month), "00") & "-" & Format(opnJEV!Transtype, "00") & "-" & Format(LastJEVSNno, "00000")
            LastJEVSNno = LastJEVSNno + 1
        Else 'No REcord Found yet in the AMIS
            grd_details.TextMatrix(cc, 4) = "000-00-00-00-xxxxx"
        End If
        opnJEV.Close
        Set opnJEV = Nothing
Next cc
End Sub

Private Sub cmd_post_Click()
Dim cc, tmp As Integer

If MsgBox("Save JEV Nos.?", vbQuestion + vbYesNo, "System Confirmation") = vbYes Then
        For cc = 1 To grd_details.Rows - 1
            
                If Len(Trim(grd_details.TextMatrix(cc, 4))) > 0 Then
                    If grd_details.TextMatrix(cc, 4) = "000-00-00-00-xxxxx" Then
                    Else
                        If IsFormatCorrect(grd_details.TextMatrix(cc, 4)) = True Then
                        
                        Call GEtCompleteJEVDetails(grd_details.TextMatrix(cc, 0), "DVNO", DTPCNdate.Value, "", "", _
                        grd_details.TextMatrix(cc, 2), grd_details.TextMatrix(cc, 4), "", "", "0", "0", "0", "4", "", grd_details.TextMatrix(cc, 0), "", cmb_FundType.Text, "", "", "", "", ExtractJEVSNo(grd_details.TextMatrix(cc, 4)), DatePost, "")
                            'Updating table from CN
                            opndbaseFMIS.Execute "Update tblAMIS_CreditNotice set AlreadySaved2JEV=1 where dvno='" & Trim(grd_details.TextMatrix(cc, 0)) & "' and actioncode = 1 "
                            
                            'Updating Accounting REcord...
                            tmp = ExtractJEVSNo(grd_details.TextMatrix(cc, 4))
                            
                           ' opndbaseFMIS.Execute "update tblAMIS_JournalEntry set JEVNo='" & grd_details.TextMatrix(cc, 14) & "', " & _
                           ' " JEVSeriesNo=" & tmp & ",JEVBy='" & ActiveUserID & "', " & _
                           ' " JEVDate='" & Date & "' where DVNo='" & grd_details.TextMatrix(cc, 5) & "'"
                        
                        opndbaseFMIS.Execute "update tblAMIS_JournalEntry set JEVNo='" & grd_details.TextMatrix(cc, 4) & "', " & _
                            " JEVSeriesNo=" & tmp & ",JEVBy='" & ActiveUserID & "', " & _
                            " JEVDate='" & Date & "',transtype = 4 where DVNo='" & Trim(grd_details.TextMatrix(cc, 0)) & "'"
                        
                        
                        
                        'GetDVNobyChkNo(grd_details.TextMatrix(grd_details.Row, 3))
                        
                        
                        
                        
                        End If
                    End If
                End If
           
        Next cc
MsgBox "Posting to JEV, Successful!", vbInformation, "System Information"
Call Command1_Click 'Loading Back Active RCI Numbers...
'List1.ListIndex = GetIndex4ListBox(List1, tmpRCINo)
End If
End Sub

Private Sub Command1_Click()
Label13.Caption = "Loading, Please wait.."
Label13.Refresh
Call LoadSavedReport(txt_RecordID.Text, DTPicker1.Year, DTPicker1.Month, cmb_FundType.Text)
Label13.Caption = ""
End Sub

Private Sub DTPicker1_Change()
DTPicker1.Value = DTPicker1.Month & "/1/" & DTPicker1.Year
Label6.Caption = MonthName(DTPicker1.Month) & " " & DTPicker1.Year
Call cmb_FundType_Click
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then
        Unload Me
End If
End Sub



Private Sub Form_Load()
WindowsXPC1.InitSubClassing
Me.Top = (Screen.Height - Me.Height) / 2
Me.Left = (Screen.Width - Me.Width) / 2
DTPicker1.Value = Month(Date) & "/1/" & Year(Date)
Label6.Caption = MonthName(DTPicker1.Month) & " " & DTPicker1.Year
Label8.Caption = ""
Label13.Caption = ""
Label14.Caption = ""
Timer1.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
WindowsXPC1.EndWinXPCSubClassing
Set frmCDCashDisbursedReport = Nothing
End Sub

Private Sub SetGrid()
Dim cc As Integer

grd_details.Clear

grd_details.Cols = 5

grd_details.Rows = 2

grd_details.TextMatrix(0, 0) = "Dvno"
grd_details.TextMatrix(0, 1) = "Claimant"
grd_details.TextMatrix(0, 2) = "Particular"
grd_details.TextMatrix(0, 3) = "Amount"
'grd_details.TextMatrix(0, 4) = "Code"
'grd_details.TextMatrix(0, 5) = "Amount"
'grd_details.TextMatrix(0, 6) = "Balance Amt."
'grd_details.TextMatrix(0, 7) = "ReconcilingSeqNo"
'grd_details.TextMatrix(0, 8) = "DVNo"
'grd_details.TextMatrix(0, 9) = "AlreadySaved"
grd_details.TextMatrix(0, 4) = "JEVNo"

grd_details.ColWidth(0) = 1200
grd_details.ColWidth(1) = 3000
grd_details.ColWidth(2) = 4000
grd_details.ColWidth(3) = 1500
grd_details.ColWidth(4) = 2000
'grd_details.ColWidth(5) = 1300
'grd_details.ColWidth(6) = 1300
'grd_details.ColWidth(7) = 1800
'grd_details.ColWidth(8) = 1800
'grd_details.ColWidth(9) = 0
'grd_details.ColWidth(10) = 1400

For cc = 0 To grd_details.Cols - 1
    grd_details.Row = 0
    grd_details.col = cc
    grd_details.CellAlignment = 4
Next cc
End Sub
Private Sub LoadAccountName(ByVal BankAccntNo As String, ByVal FundType As String, ByVal BankID As String)
Dim opnAcctname As New ADODB.Recordset
Dim xx As Long

opnAcctname.Open "Select * from vw_DepositoryBank where BankID='" & BankID & "' and FundType='" & FundType & "' and BankAccountNo='" & BankAccntNo & "' and active=1", opndbaseFMIS, adOpenStatic, adLockOptimistic
If opnAcctname.RecordCount <> 0 Then
    cmb_AccountName.Clear
    Do Until opnAcctname.EOF
        cmb_AccountName.AddItem (opnAcctname!Accountname)
        cmb_AccountName.ItemData(xx) = opnAcctname!FmisAccountcode
        xx = xx + 1
        opnAcctname.MoveNext
    Loop
Else
    cmb_AccountName.Clear
End If
opnAcctname.Close
Set opnAcctname = Nothing
End Sub
Private Sub LoadBackBreakdown(ByVal cnno As String)
Dim opnvoucher As New ADODB.Recordset
Dim sql As String

sql = "SELECT min(b.trnno),a.dvno,a.obrno,a.particular,a.claimantcode,a.gamount,a.fundtype,b.cnno,b.cndate,receiveddate " & _
        "FROM tblAMIS_IncomingDVTrns as a inner join tblAMIS_CreditNotice as b ON a.dvno = b.dvno  " & _
        "WHERE b.cnno = '" & cnno & "' and b.ACTIONCODE = 1 and a.actioncode = 1 and b.AlreadySaved2JEV = 0 " & _
        "group by a.obrno,a.particular,a.claimantcode,a.gamount,b.cnno,b.cndate,receiveddate,a.dvno,a.fundtype order by a.dvno"

'Debug.Print sql

opnvoucher.Open sql, opndbaseFMIS, adOpenStatic, adLockOptimistic

If opnvoucher.RecordCount > 0 Then
    Combo1.Text = Trim(opnvoucher!FundType)
    Cmb_CnNo.Text = Trim(opnvoucher.Fields!cnno)
    DTPCNdate.Value = opnvoucher.Fields!cndate
    DTPRdate.Value = opnvoucher.Fields!receiveddate
'    Do Until opnvoucher.EOF
'            With opnvoucher
'            Set X = ListView2.ListItems.Add(, , .Fields!DVNo)
'            X.SubItems(1) = .Fields!ObrNo
'            X.SubItems(2) = .Fields!particulargrd_details
'            X.SubItems(3) = GetClaimantDetails(IIf(IsNull(!ClaimantCode), "N/A", !ClaimantCode), "Name")
'            X.SubItems(4) = Format(.Fields!GAmount, "#,###.00")
'            X.SubItems(5) = .Fields!DVNo
'            End With
'            opnvoucher.MoveNext
'    Loop
    Call SetGrid
    grd_details.Rows = opnvoucher.RecordCount + 1
    Do Until opnvoucher.EOF
        grd_details.TextMatrix(opnvoucher.AbsolutePosition, 0) = IIf(IsNull(opnvoucher!DVNo), 0, opnvoucher!DVNo)
        grd_details.TextMatrix(opnvoucher.AbsolutePosition, 1) = GetClaimantDetails(IIf(IsNull(opnvoucher!ClaimantCode), "N/A", opnvoucher!ClaimantCode), "Name")
        grd_details.TextMatrix(opnvoucher.AbsolutePosition, 2) = opnvoucher.Fields!Particular
        grd_details.TextMatrix(opnvoucher.AbsolutePosition, 3) = Format(opnvoucher.Fields!Gamount, "#,###.00")
        grd_details.TextMatrix(opnvoucher.AbsolutePosition, 4) = ""
'        grd_details.TextMatrix(opnvoucher.AbsolutePosition, 5) = Format(opnvoucher!Amount, "###,##0.00") 'Cash Advanced Amount
'        grd_details.TextMatrix(opnvoucher.AbsolutePosition, 6) = IIf(IsNull(opnvoucher!balanceamt), "", Format(opnvoucher!balanceamt, "###,##0.00")) 'Liquiditing Amount
'        grd_details.TextMatrix(opnvoucher.AbsolutePosition, 7) = IIf(IsNull(opnvoucher!ReconcilingSeqNo), "", opnvoucher!ReconcilingSeqNo) 'Normal Balance
'        grd_details.TextMatrix(opnvoucher.AbsolutePosition, 8) = opnvoucher!DVNo
'        'grd_details.TextMatrix(opnvoucher.AbsolutePosition, 9) = IIf(IsNull(opnvoucher!RefundORNo), 0, 1)
        
        
        opnvoucher.MoveNext
    Loop
Else
   ' Call ClearCmb
    Call SetGrid
'MsgBox "No Record Found On the Database", vbInformation, "System Message"
End If
 
    
opnvoucher.Close
Set opnvoucher = Nothing
End Sub
Private Sub grd_details_DblClick()
If Len(grd_details.TextMatrix(grd_details.Row, 4)) > 0 Then
    ActiveFormCaller = Me.name
    ForTheGridRowNo = grd_details.Row

    If Len(grd_details.TextMatrix(grd_details.Row, 4)) <> 0 Then 'Kung Naa nay JEV No
        frmJEVNumberingAssignment_New.txt_Jevno.Text = grd_details.TextMatrix(grd_details.Row, 4)
        frmJEVNumberingAssignment_New.txt_DVNo.Text = grd_details.TextMatrix(grd_details.Row, 0)
        frmJEVNumberingAssignment_New.Show vbModal
    Else
        frmJEVNumberingAssignment_New.txt_DVNo.Text = grd_details.TextMatrix(grd_details.Row, 0)
        frmJEVNumberingAssignment_New.Show vbModal
    End If
Else
    MsgBox "There is no Voucher Attachment for this Check!" & Chr(13) & Chr(13) & "Please Select a New..", vbInformation, "System Information"
End If
End Sub

Private Sub Timer1_Timer()
Call SetGrid
Call LoadFundType(cmb_FundType)
Call LoadSavedReport(ActiveUserID, DTPicker1.Year, DTPicker1.Month, cmb_FundType.Text)
Timer1.Enabled = False
End Sub

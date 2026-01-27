VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmJEVPreparationforMemoEntry 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Memo Entry"
   ClientHeight    =   9375
   ClientLeft      =   -165
   ClientTop       =   2850
   ClientWidth     =   11565
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   -1  'True
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmJEVPreparationforMEMOEntry.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9375
   ScaleWidth      =   11565
   Visible         =   0   'False
   Begin VB.CheckBox chkSC 
      Caption         =   "Single Click"
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
      TabIndex        =   50
      Top             =   4560
      Value           =   1  'Checked
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10920
      TabIndex        =   46
      ToolTipText     =   "Click me to Generate Jev number"
      Top             =   3960
      Width           =   495
   End
   Begin VB.TextBox txtJEVNo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   5640
      MaxLength       =   18
      TabIndex        =   45
      ToolTipText     =   "format(000-00-00-00-00000)"
      Top             =   3960
      Width           =   5205
   End
   Begin VB.TextBox txtformula 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   2640
      TabIndex        =   43
      Top             =   4560
      Width           =   8775
   End
   Begin VB.CommandButton btnSearch 
      Caption         =   "Search"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   12840
      TabIndex        =   41
      Top             =   840
      Visible         =   0   'False
      Width           =   1665
   End
   Begin VB.TextBox txtClaimantCode 
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
      Left            =   -120
      Locked          =   -1  'True
      TabIndex        =   39
      Top             =   3360
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.CheckBox chkSTP 
      Caption         =   "Shoot-To-Print"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   12000
      TabIndex        =   37
      Top             =   8880
      Width           =   1815
   End
   Begin VB.ComboBox cmb_month 
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
      Left            =   12990
      TabIndex        =   30
      Top             =   2430
      Width           =   1470
   End
   Begin VB.TextBox txtDate 
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
      Left            =   5910
      Locked          =   -1  'True
      TabIndex        =   27
      Top             =   1340
      Width           =   2265
   End
   Begin VB.ComboBox cmb_trnYear 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   12990
      TabIndex        =   26
      Top             =   2010
      Width           =   1470
   End
   Begin VB.Frame Frame2 
      Caption         =   "Transaction Details"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1965
      Left            =   180
      TabIndex        =   14
      Top             =   1920
      Width           =   11235
      Begin VB.ComboBox txtFund 
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
         Style           =   2  'Dropdown List
         TabIndex        =   48
         Top             =   1440
         Width           =   4095
      End
      Begin VB.ComboBox cmbrc 
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
         Left            =   9840
         Locked          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   42
         Top             =   2160
         Visible         =   0   'False
         Width           =   4095
      End
      Begin VB.CommandButton btnParticular 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   9480
         TabIndex        =   40
         ToolTipText     =   "Click here to edit particulars..."
         Top             =   240
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CommandButton btnClaimant 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4680
         TabIndex        =   38
         ToolTipText     =   "Click here to select claimant..."
         Top             =   2400
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.TextBox txtAmount 
         Alignment       =   2  'Center
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
         Left            =   7680
         TabIndex        =   23
         Top             =   1440
         Width           =   1785
      End
      Begin VB.TextBox txtParticular 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1080
         Left            =   1440
         MultiLine       =   -1  'True
         TabIndex        =   21
         Top             =   240
         Width           =   8010
      End
      Begin VB.TextBox txtAlobs 
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
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   2640
         Width           =   4260
      End
      Begin VB.TextBox txtClaimant 
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
         Left            =   315
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   2385
         Visible         =   0   'False
         Width           =   4260
      End
      Begin VB.TextBox txtRC 
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
         Left            =   9840
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   2460
         Visible         =   0   'False
         Width           =   4050
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fund Type:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   240
         TabIndex        =   25
         Top             =   1440
         Width           =   1185
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Amount (Gross):"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   5880
         TabIndex        =   24
         Top             =   1440
         Width           =   1755
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Particular:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   360
         TabIndex        =   22
         Top             =   360
         Width           =   1050
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Report No:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   180
         TabIndex        =   20
         Top             =   2670
         Width           =   1155
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Claimant"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   180
         TabIndex        =   18
         Top             =   2130
         Visible         =   0   'False
         Width           =   600
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Responsibility Center"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   9780
         TabIndex        =   17
         Top             =   2070
         Visible         =   0   'False
         Width           =   2235
      End
   End
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
      Height          =   5580
      Left            =   12000
      TabIndex        =   12
      Top             =   3240
      Width           =   2505
   End
   Begin VB.CommandButton btnPrtJEV 
      Caption         =   "Print JEV"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   11985
      TabIndex        =   11
      Top             =   9240
      Width           =   2535
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   4215
      Left            =   165
      ScaleHeight     =   4185
      ScaleWidth      =   11160
      TabIndex        =   3
      Top             =   5040
      Width           =   11190
      Begin VB.ComboBox cmbEntry 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   0
         TabIndex        =   36
         Text            =   "cmbEntry"
         Top             =   600
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txt_entry 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   4920
         TabIndex        =   29
         Top             =   1800
         Visible         =   0   'False
         Width           =   1665
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   4200
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Width           =   11160
         _ExtentX        =   19685
         _ExtentY        =   7408
         _Version        =   393216
         FixedCols       =   0
         ScrollTrack     =   -1  'True
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
   End
   Begin VB.TextBox txtDVNo 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   405
      TabIndex        =   1
      Top             =   1215
      Width           =   4845
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   840
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11565
      _ExtentX        =   20399
      _ExtentY        =   1482
      ButtonWidth     =   1085
      ButtonHeight    =   1429
      Appearance      =   1
      Style           =   1
      ImageList       =   "itb32x32"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   6
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "New"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&New"
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Save"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "Delete"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Close"
            ImageIndex      =   7
         EndProperty
      EndProperty
      Begin MSComctlLib.ImageList itb32x32 
         Left            =   7560
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
               Picture         =   "frmJEVPreparationforMEMOEntry.frx":076A
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmJEVPreparationforMEMOEntry.frx":20FC
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmJEVPreparationforMEMOEntry.frx":3A8E
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmJEVPreparationforMEMOEntry.frx":5420
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmJEVPreparationforMEMOEntry.frx":6DB2
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmJEVPreparationforMEMOEntry.frx":8744
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmJEVPreparationforMEMOEntry.frx":A0D6
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmJEVPreparationforMEMOEntry.frx":BA68
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmJEVPreparationforMEMOEntry.frx":D3FA
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmJEVPreparationforMEMOEntry.frx":ED8E
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmJEVPreparationforMEMOEntry.frx":FA6A
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmJEVPreparationforMEMOEntry.frx":1034A
               Key             =   ""
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmJEVPreparationforMEMOEntry.frx":11026
               Key             =   ""
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmJEVPreparationforMEMOEntry.frx":11D02
               Key             =   ""
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmJEVPreparationforMEMOEntry.frx":129DE
               Key             =   ""
            EndProperty
            BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmJEVPreparationforMEMOEntry.frx":136BA
               Key             =   ""
            EndProperty
            BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmJEVPreparationforMEMOEntry.frx":14396
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "JEV Transaction Type"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Left            =   600
      TabIndex        =   6
      Top             =   5280
      Width           =   7830
      Begin VB.OptionButton optOther 
         Caption         =   "Other"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   6405
         TabIndex        =   10
         Tag             =   "04"
         Top             =   300
         Width           =   1230
      End
      Begin VB.OptionButton optCash 
         Caption         =   "Cash Disbursement"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   4245
         TabIndex        =   9
         Tag             =   "03"
         Top             =   300
         Width           =   2100
      End
      Begin VB.OptionButton optCheck 
         Caption         =   "Check Disbursement"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1965
         TabIndex        =   8
         Tag             =   "02"
         Top             =   300
         Width           =   2100
      End
      Begin VB.OptionButton optCollection 
         Caption         =   "Collection"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   270
         TabIndex        =   7
         Tag             =   "01"
         Top             =   285
         Value           =   -1  'True
         Width           =   1260
      End
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   8160
      TabIndex        =   49
      Top             =   1320
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   661
      _Version        =   393216
      Format          =   271319041
      CurrentDate     =   40631
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Refference:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   360
      TabIndex        =   47
      Top             =   960
      Width           =   1215
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
      Left            =   2280
      TabIndex        =   44
      Top             =   4560
      Width           =   360
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mode:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   8880
      TabIndex        =   35
      Top             =   1290
      Width           =   750
   End
   Begin VB.Label lblMode 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NEW"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   9795
      TabIndex        =   34
      Top             =   1320
      Width           =   615
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "Trn Year :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   12180
      TabIndex        =   32
      Top             =   2085
      Width           =   945
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "Month of:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   12180
      TabIndex        =   31
      Top             =   2505
      Width           =   915
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date Prepared"
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
      Height          =   240
      Left            =   5910
      TabIndex        =   28
      Top             =   1100
      Width           =   1335
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   885
      Left            =   12045
      Top             =   1950
      Width           =   2475
   End
   Begin VB.Label Label3 
      Caption         =   "Vouchers Prepared with JEV"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   12030
      TabIndex        =   13
      Top             =   2955
      Width           =   2430
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Accounting Entries"
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
      Left            =   195
      TabIndex        =   4
      Top             =   4800
      Width           =   1650
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Memo  Number:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   3720
      TabIndex        =   2
      Top             =   4080
      Width           =   1890
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   975
      Left            =   8640
      Top             =   840
      Width           =   2235
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   960
      Left            =   -15
      Top             =   855
      Width           =   8625
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Assigned Disbursement Voucher No :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   5400
      TabIndex        =   33
      Top             =   960
      Visible         =   0   'False
      Width           =   2640
   End
End
Attribute VB_Name = "frmJEVPreparationforMemoEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim Edited As Boolean
Dim xDebit As Currency
Dim xCredit As Currency
Dim xObR As String
Dim xNAcode As String
Dim CUFlag As Boolean           'Claimant Update Flag
Dim XFlag As Boolean
Public isfrom_jevNumbering As Boolean
Public EditCount, IsSaveAccntng As Boolean
Dim not_coloraly_total_debit, not_coloraly_total_credit, coloraly_total_debit, coloraly_total_credit As Double


Private Sub btnClaimant_Click()
    CUFlag = True
    ActiveFormCaller = "frmJEVPreparation"
    frmCDClaimantRegistry.Show 1
End Sub

Private Sub btnParticular_Click()
    CUFlag = True
    
    If txtParticular.Locked = True Then
    txtParticular.Locked = False
    Else
    txtParticular.Locked = True
    End If
    
End Sub

Private Sub btnPrtJEV_Click()
Dim sql As String

If Edited = True Then
'    SQL = "SELECT dbo.tblAMIS_IncomingDVTrns.RCenterCode, dbo.tblAMIS_JournalEntry.TransDate, dbo.tblAMIS_JournalEntry.TransType," & _
'            "dbo.tblAMIS_JournalEntry.FmisAccntCode, dbo.tblREF_AIS_ChartofAccounts.AccountNameFull, dbo.tblREF_AIS_ChartofAccounts.ChildAccountCode," & _
'            "dbo.tblAMIS_JournalEntry.Amount, dbo.tblAMIS_JournalEntry.DebitCredit, dbo.tblAMIS_JournalEntry.Actioncode," & _
'            "dbo.tblAMIS_IncomingDVTrns.Particular , dbo.tblAMIS_IncomingDVTrns.ClaimantCode FROM dbo.tblAMIS_JournalEntry INNER JOIN " & _
'            "dbo.tblAMIS_IncomingDVTrns ON dbo.tblAMIS_JournalEntry.DVNo = dbo.tblAMIS_IncomingDVTrns.DVNo AND " & _
'            "dbo.tblAMIS_JournalEntry.Actioncode = dbo.tblAMIS_IncomingDVTrns.Actioncode INNER JOIN " & _
'            "dbo.tblREF_AIS_ChartofAccounts ON dbo.tblAMIS_JournalEntry.FmisAccntCode = dbo.tblREF_AIS_ChartofAccounts.FMISAccountCode AND " & _
'            "dbo.tblAMIS_JournalEntry.ActionCode = dbo.tblREF_AIS_ChartofAccounts.Active " & _
'            "WHERE (dbo.tblREF_AIS_ChartofAccounts.FundType ='" & GetFundName(txtFund.Text) & "') AND (dbo.tblAMIS_JournalEntry.DVNo ='" & List1.Text & "')"
    
'sql = "SELECT dbo.tblAMIS_IncomingDVTrns.RCenterCode, dbo.tblAMIS_JournalEntry.TransDate, dbo.tblAMIS_JournalEntry.TransType," & _
'            "dbo.tblAMIS_JournalEntry.FmisAccntCode, dbo.tblREF_AIS_ChartofAccounts.AccountNameFull, dbo.tblREF_AIS_ChartofAccounts.ChildAccountCode," & _
'            "dbo.tblAMIS_JournalEntry.Amount, dbo.tblAMIS_JournalEntry.DebitCredit, dbo.tblAMIS_JournalEntry.Actioncode," & _
'            "dbo.tblAMIS_IncomingDVTrns.Particular , dbo.tblAMIS_IncomingDVTrns.ClaimantCode FROM dbo.tblAMIS_JournalEntry INNER JOIN " & _
'            "dbo.tblAMIS_IncomingDVTrns ON dbo.tblAMIS_JournalEntry.DVNo = dbo.tblAMIS_IncomingDVTrns.DVNo AND " & _
'            "dbo.tblAMIS_JournalEntry.Actioncode = dbo.tblAMIS_IncomingDVTrns.Actioncode INNER JOIN " & _
'            "dbo.tblREF_AIS_ChartofAccounts ON dbo.tblAMIS_JournalEntry.FmisAccntCode = dbo.tblREF_AIS_ChartofAccounts.FMISAccountCode AND " & _
'            "(dbo.tblAMIS_JournalEntry.ActionCode = dbo.tblREF_AIS_ChartofAccounts.Active or dbo.tblAMIS_JournalEntry.ActionCode=5 )" & _
'            "WHERE (dbo.tblREF_AIS_ChartofAccounts.FundType ='" & GetFundName(txtFund.Text) & "') AND (dbo.tblAMIS_JournalEntry.DVNo ='" & List1.Text & "')"
'
    'Debug.Print sql
    sql = "select * from vw_MP_JevPreperCollection where fundtype ='" & GetFundName(txtfund.Text) & "' AND (ptvno ='" & Me.txtDVNo.Text & "') and (actioncode = 1 or actioncode = 5)"
    ReportName = "JEV"
    rptJEVCollection.txtClaimDesc.SetText txtParticular.Text & ", " & txtClaimant.Text & ", " & txtAlobs.Text
    rptJEVCollection.txtRC.SetText txtRC.Text
    rptJEVCollection.txtClerk.SetText getUserName(ActiveUserID, "FullName")
    
    If chkSTP.Value = 1 Then
        rptJEVCollection.Line1.Suppress = True
        rptJEVCollection.Line2.Suppress = True
        rptJEVCollection.Line3.Suppress = True
        rptJEVCollection.Line4.Suppress = True
        rptJEVCollection.Line5.Suppress = True
        rptJEVCollection.Line6.Suppress = True
        rptJEVCollection.Line8.Suppress = True
        rptJEVCollection.Line9.Suppress = True
        rptJEVCollection.Line10.Suppress = True
        rptJEVCollection.Line11.Suppress = True
        rptJEVCollection.Line12.Suppress = True
        rptJEVCollection.Line13.Suppress = True
        rptJEVCollection.Line14.Suppress = True
        rptJEVCollection.Line15.Suppress = True
        rptJEVCollection.Line16.Suppress = True
        rptJEVCollection.Line17.Suppress = True
        rptJEVCollection.Line18.Suppress = True
        rptJEVCollection.Line19.Suppress = True
        
        rptJEVCollection.Text1.Suppress = True
        rptJEVCollection.Text2.Suppress = True
        rptJEVCollection.Text3.Suppress = True
        rptJEVCollection.Text4.Suppress = True
        rptJEVCollection.Text8.Suppress = True
        rptJEVCollection.Text9.Suppress = True
        rptJEVCollection.Text12.Suppress = True
        rptJEVCollection.Text13.Suppress = True
        rptJEVCollection.Text15.Suppress = True
        rptJEVCollection.Text16.Suppress = True
        rptJEVCollection.Text17.Suppress = True
        rptJEVCollection.Text18.Suppress = True
        rptJEVCollection.Text19.Suppress = True
        rptJEVCollection.Text20.Suppress = True
        rptJEVCollection.Text21.Suppress = True
        rptJEVCollection.Text22.Suppress = True
        rptJEVCollection.Text25.Suppress = True
        
    End If
    
    rptJEVCollection.Database.SetDataSource opndbaseFMIS.Execute(sql)
    rptJEVCollection.Database.Verify
   frmViewercollect.Show 1
End If

End Sub


Private Sub btnSearch_Click()
    frmDVSearch.Show 1
End Sub

Private Sub cmb_month_Click()
    Call LoadPrevTrans
End Sub

Private Sub LoadPrevTrans()
Dim PRec As New ADODB.Recordset
Dim x As Integer
    List1.Clear
    List1.Enabled = False
    PRec.Open ("Select top 500 refno, min(trnno) as trno  From tblAMIS_FinalJEV Where Actioncode = 1 and IsAdjustment =1 and transtype = 4  Group By jevno order by trnno desc"), opndbaseFMIS, adOpenStatic, adLockOptimistic
    If PRec.RecordCount > 0 Then
        For x = 1 To PRec.RecordCount
            List1.AddItem PRec!ptvNo
            PRec.MoveNext
        Next x
        List1.Enabled = True
    End If
    PRec.Close
    Set PRec = Nothing
End Sub

Private Sub cmb_trnYear_Click()
    Call LoadPrevTrans
End Sub

Private Sub cmbEntry_KeyPress(KeyAscii As Integer)
    
    
    If KeyAscii = 13 Then
        If cmbEntry.ListIndex <> -1 Then
            MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 1) = cmbEntry.Text
            MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 0) = cmbEntry.ItemData(cmbEntry.ListIndex)
            MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 5) = "1"
            If MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2) = "TOTAL" Then
                MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 3) = ""
                MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 4) = ""
            End If
            MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2) = GetAccountNameByFMISAccountCode(cmbEntry.ItemData(cmbEntry.ListIndex))
            If MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 1) = "101" Then
                MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 5) = "5"
            End If
        ElseIf cmbEntry.Text = "" Then
            MSFlexGrid1.RemoveItem (MSFlexGrid1.Row)

        End If
        cmbEntry.Visible = False
        Call GetSum
    Else
        KeyAscii = AutoFind(cmbEntry, KeyAscii, True)
    End If

End Sub





Private Sub cmdOK_Click()

End Sub

Private Sub cmbRC_Click()
    If Trim(cmbRC.Text) <> "" Then
        txtRC = Trim(cmbRC.Text)
        txtRC.Visible = True
        cmbRC.Visible = False
    End If
End Sub

Private Sub cmbRC_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        cmbRC_Click
    End If
End Sub

Private Sub Command1_Click()
Dim rec As New ADODB.Recordset
Dim Lastno As Double
JevOk = False
frmPOstdate.Show 1
If JevOk = True Then
    rec.Open ("EXEC [dbo].[Proc_GetMaxJevSeries_New] @transtype = 5,@jevyeardate = '" & DatePost & "' ,@fundtype = '" & txtfund.Text & "'"), opndbaseFMIS, adOpenStatic, adLockOptimistic
        Lastno = rec.Fields!MAXJEVSERIES
    rec.Close
    
    txtJEVNo.Text = txtfund.ItemData(txtfund.ListIndex) & "-" & Right(Year(DatePost), 2) & "-" & Format(Month(DatePost), "00") & "-" & "05" & "-" & Format(Lastno, "0000")
End If
End Sub

Private Sub DTPicker1_CallbackKeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer, ByVal CallbackField As String, CallbackDate As Date)
txtdate.Text = Format(DTPicker1.Value, "MM/dd/yyyy")
End Sub

Private Sub DTPicker1_Change()
txtdate.Text = Format(DTPicker1.Value, "MM/dd/yyyy")
End Sub

Private Sub DTPicker1_Click()
txtdate.Text = Format(DTPicker1.Value, "MM/dd/yyyy")
End Sub

Private Sub Form_Load()
    Call SetGrid
    ActiveUserID = Trim(ActiveUserID)
    DTPicker1.Value = Now
    txtdate.Text = Format(DTPicker1.Value, "MM/dd/yyyy")
    Call LoadFundType(txtfund)
End Sub
Private Sub SetGrid()
Dim cc As Integer

    MSFlexGrid1.Clear
    MSFlexGrid1.Rows = 50
    MSFlexGrid1.Cols = 7
    
    MSFlexGrid1.TextMatrix(0, 1) = "Account Code"
    MSFlexGrid1.TextMatrix(0, 2) = "Accounts and Explanation"
    MSFlexGrid1.TextMatrix(0, 3) = "Debit"
    MSFlexGrid1.TextMatrix(0, 4) = "Credit"
    MSFlexGrid1.TextMatrix(0, 5) = "ActionCode"
    MSFlexGrid1.TextMatrix(0, 6) = "Formula"
    
    MSFlexGrid1.ColWidth(0) = 0
    MSFlexGrid1.ColWidth(1) = 2500
    MSFlexGrid1.ColWidth(2) = 5000
    MSFlexGrid1.ColWidth(3) = 1500
    MSFlexGrid1.ColWidth(4) = 1500
    MSFlexGrid1.ColWidth(5) = 0
    MSFlexGrid1.ColWidth(6) = 0
    For cc = 0 To MSFlexGrid1.Cols - 1
        MSFlexGrid1.Row = 0
        MSFlexGrid1.col = cc
        MSFlexGrid1.CellAlignment = 4
    Next cc
End Sub

Private Sub List1_Click()
    Call LoadJEVDetails(List1.Text)
    cmbRC.Visible = False
   ' txtRC.Visible = True
End Sub

Private Sub LoadJEVDetails(ByVal dvno As String)
Dim DRec As New ADODB.Recordset
Dim x As Integer
    
    CUFlag = False
    txtParticular.Locked = True
    xNAcode = ""
    Edited = True
    lblMode.Caption = "EDIT"
    DRec.Open ("Select * From [tblAMIS_COllectionDepositt] Where [ptvno]='" & dvno & "' And (ActionCode=1 or ActionCode=5)"), opndbaseFMIS, adOpenStatic, adLockOptimistic
    If DRec.RecordCount > 0 Then
        
        txtDVNo.Text = DRec![ptvNo]
        'txtJEVNo.Text = DRec!JEVNo
        txtdate.Text = DRec![TransDate]
        'txtParticular.Text = DRec!Particular
        'txtFund.Text = GetFundMedium(DRec!FundCode)
        'txtAmount.Text = DRec!Amount
        
'        If CInt(optCollection.Tag) = DRec![TransType] Then optCollection.Value = True
'        If CInt(optCheck.Tag) = DRec![TransType] Then optCheck.Value = True
'        If CInt(optCash.Tag) = DRec![TransType] Then optCash.Value = True
'        If CInt(optOther.Tag) = DRec![TransType] Then optOther.Value = True
        
'        If DRec!continuing = 1 Then
'            XFlag = True
'        Else
'            XFlag = False
'        End If
    
    End If
    DRec.Close
    Set DRec = Nothing
        
    DRec.Open ("Select * FRom tblCMS_CDCheckBook where DVNo='" & txtDVNo.Text & "' and ActionCode=1"), opndbaseFMIS, adOpenStatic, adLockOptimistic
    If DRec.RecordCount > 0 Then
    
       txtAlobs.Text = DRec!chknumber
        ' txtClaimant.Text = GetClaimant(DVRec!ClaimantCode)
         'txtClaimantCode.Text = DVRec!ClaimantCode
        ' txtRC.Text = GetOfficeName(DVRec!RCenter, "OfficeMedium")
         txtParticular.Text = DRec!Particular
         txtfund.Text = GetFundMedium(DRec!fundcode)
         txtAmount.Text = DRec!amount
    End If
    DRec.Close
    Set DRec = Nothing
'
    Call SetGrid
    'DRec.Close
    DRec.Open ("Select * From [tblAMIS_COllectionDepositt] Where [ptvno]='" & dvno & "' And (ActionCode=1 or actioncode=5)"), opndbaseFMIS, adOpenStatic, adLockOptimistic
    If DRec.RecordCount > 0 Then
        For x = 1 To DRec.RecordCount
            MSFlexGrid1.TextMatrix(x, 0) = DRec![FmisAccntCode]
            MSFlexGrid1.TextMatrix(x, 1) = GetAccountCodeByFMISAccountCode(DRec![FmisAccntCode])
            MSFlexGrid1.TextMatrix(x, 2) = GetAccountNameByFMISAccountCode(DRec![FmisAccntCode])
            If DRec![debitcredit] = 0 Then
                MSFlexGrid1.TextMatrix(x, 4) = DRec!amount
            Else
                MSFlexGrid1.TextMatrix(x, 3) = DRec!amount
            End If
            If LCase(Trim(lblMode)) = "edit" Then MSFlexGrid1.TextMatrix(x, 5) = DRec!ActionCode  ' for coloraly purpose
            DRec.MoveNext
        Next x
        Call GetSum
    End If
    DRec.Close
    Set DRec = Nothing
    
    Call LoadAccountsByFund(Trim(txtfund.Text))

End Sub

Private Sub MSFlexGrid1_Click()
If chkSC.Value = 1 Then
Call MSFlexGrid1_DblClick
End If
End Sub

Private Sub MSFlexGrid1_DblClick()
If Trim(txtJEVNo.Text) = "" Then
    MsgBox "JEV Number is Empty, Please Specify..!", vbInformation, "System Information"
    Exit Sub
End If
If Trim(txtParticular.Text) <> "" Then
    With frmSub3
        .reff = txtJEVNo.Text
        .Gamount = txtAmount.Text
        .CName = "N/A"
        .isEdit = True
       Set .frm = Me
        Call LoadAcctngEntries(txtJEVNo.Text)
        .Show 1
        Call GetAccntngEntries
    End With
End If
End Sub
Public Sub GetAccntngEntries()
Dim DRec As New ADODB.Recordset
Dim x As Integer
Call SetGrid
    'DRec.Close
    If IsSaveAccntng = False Then
        Set DRec = opndbaseFMIS.Execute("Select left(ChildAccountcode,3) as childcode,sum(Debit) as sumdebit,sum(credit) as sumcredit From tblAMIS_AccoutingEntries Where [reffno]='" & txtJEVNo.Text & "' And (ActionCode=1) group by reffno,actioncode,left(ChildAccountcode,3)")
        If DRec.RecordCount > 0 Then
            For x = 1 To DRec.RecordCount
    '            MSFlexGrid1.TextMatrix(x, 0) = DRec![Trnno]
                MSFlexGrid1.TextMatrix(x, 1) = DRec!childcode
                MSFlexGrid1.TextMatrix(x, 2) = GetAccountNameByAccountcode(DRec!childcode)
                MSFlexGrid1.TextMatrix(x, 4) = IIf((Format(DRec!sumCredit, "#,##0.00") = "0.00"), "", Format(DRec!sumCredit, "#,##0.00"))
                MSFlexGrid1.TextMatrix(x, 3) = IIf((Format(DRec!sumDebit, "#,##0.00") = "0.00"), "", Format(DRec!sumDebit, "#,##0.00"))
              MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1
               ' If LCase(Trim(lblMode.Caption)) = "edit" Then MSFlexGrid1.TextMatrix(x, 5) = DRec!ActionCode  ' for coloraly purpose
                DRec.MoveNext
            Next x
        Else
        IsSaveAccntng = True
        Call GetAccntngEntries
        End If
    Else
        Set DRec = opndbaseFMIS.Execute("Select left(Accountcode,3) as childcode,sum(Debit) as sumdebit,sum(credit) as sumcredit From tblAMIS_tmpjournal Where [dvno]='" & txtJEVNo.Text & "' group by Dvno,left(Accountcode,3)")
    If DRec.RecordCount > 0 Then
        For x = 1 To DRec.RecordCount
            'MSFlexGrid1.TextMatrix(x, 0) = DRec![Trnno]
            
            MSFlexGrid1.TextMatrix(x, 1) = DRec!childcode
            MSFlexGrid1.TextMatrix(x, 2) = GetAccountNameByAccountcode(DRec!childcode)
            MSFlexGrid1.TextMatrix(x, 4) = IIf((Format(DRec!sumCredit, "#,##0.00") = "0.00"), "", Format(DRec!sumCredit, "#,##0.00"))
            MSFlexGrid1.TextMatrix(x, 3) = IIf((Format(DRec!sumDebit, "#,##0.00") = "0.00"), "", Format(DRec!sumDebit, "#,##0.00"))
            MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1
            'If LCase(Trim(lblMode.Caption)) = "edit" Then MSFlexGrid1.TextMatrix(x, 5) = DRec!ActionCode  ' for coloraly purpose
            DRec.MoveNext
        Next x
    End If
    End If
    Call GetSum
    DRec.Close
    Set DRec = Nothing
End Sub

Public Function LoadAcctngEntries(ByVal dvno As String)
Dim DRec As New ADODB.Recordset
Dim rec As New ADODB.Recordset
Dim x As Integer
    DRec.Open ("Select ChildAccountcode,Debit ,Credit From tblAMIS_AccoutingEntries Where [reffno]='" & dvno & "' And (ActionCode=1) "), opndbaseFMIS, adOpenStatic, adLockOptimistic
    If DRec.RecordCount > 0 Then
        If EditCount = False Then
        EditCount = True
            rec.Open "Select dvno from tblAMIs_tmpjournal where dvno = '" & dvno & "'", opndbaseFMIS, adOpenStatic
            If rec.RecordCount > 0 Then
                    If MsgBox("This transaction Have a temporary Accounting Entries, do you want to Delete?", vbCritical + vbYesNo, "System Information") = vbYes Then
                        opndbaseFMIS.Execute "Delete from tblAMIs_tmpjournal where Dvno = '" & dvno & "'"
                        For x = 1 To DRec.RecordCount
                        DoEvents
                            opndbaseFMIS.Execute "Insert into tblAMIs_tmpjournal (Dvno,Accountcode,Debit,Credit) values ('" & Trim(dvno) & "','" & Trim(DRec!childaccountcode) & "'," & DRec!Debit & "," & DRec!Credit & ")"
                            DRec.MoveNext
                        Next x
                    End If
            Else
            For x = 1 To DRec.RecordCount
                        DoEvents
                            opndbaseFMIS.Execute "Insert into tblAMIs_tmpjournal (Dvno,Accountcode,Debit,Credit) values ('" & Trim(dvno) & "','" & Trim(DRec!childaccountcode) & "'," & DRec!Debit & "," & DRec!Credit & ")"
                            DRec.MoveNext
                        Next x
            End If
            rec.Close
        End If
    Else
         If EditCount = False Then
            EditCount = True
            rec.Open "Select dvno from tblAMIs_tmpjournal where dvno = '" & dvno & "'", opndbaseFMIS, adOpenStatic
            If rec.RecordCount > 0 Then
                    If MsgBox("This transaction Have a temporary Accounting Entries, do you want to Delete?", vbCritical + vbYesNo, "System Information") = vbYes Then
                        opndbaseFMIS.Execute "Delete from tblAMIs_tmpjournal where Dvno = '" & dvno & "'"
                        For x = 1 To DRec.RecordCount
                        DoEvents
                            opndbaseFMIS.Execute "Insert into tblAMIs_tmpjournal (Dvno,Accountcode,Debit,Credit) values ('" & Trim(dvno) & "','" & Trim(DRec!childaccountcode) & "'," & DRec!Debit & "," & DRec!Credit & ")"
                            DRec.MoveNext
                        Next x
                    End If
            End If
            rec.Close
        End If
    End If
    DRec.Close
    Set DRec = Nothing
End Function


Private Sub optCash_Click()
    'txtJEVNo.Text = GetNewJEV(optCash.Tag)
End Sub

Private Sub optCheck_Click()
    'txtJEVNo.Text = GetNewJEV(optCheck.Tag)
End Sub

Private Sub optCollection_Click()
    'txtJEVNo.Text = GetNewJEV(optCollection.Tag)
End Sub

Private Sub optOther_Click()
    'txtJEVNo.Text = GetNewJEV(optOther.Tag)
End Sub

Private Function GetNewJEV(ByVal JournalCode As String) As String
Dim Jrec As New ADODB.Recordset
Dim xCode As String

    GetNewJEV = ""
    xCode = GetFundCODE(txtfund.Text) & "-" & Format(Now, "yy-mm") & "-" & JournalCode
    Jrec.Open ("Select * from tblAMIS_JournalEntry where JEVNo like '" & xCode & "%' Order by trnno desc"), opndbaseFMIS, adOpenStatic, adLockOptimistic
    If Jrec.RecordCount > 0 Then
        GetNewJEV = xCode & "-" & Format(CInt(Right(Jrec!jevno, 3)) + 1, "000")
    Else
        GetNewJEV = xCode & "-001"
    End If
    Jrec.Close
    Set Jrec = Nothing
    
End Function

'-----RICHARD--------
Private Function getdetails(signal As Integer) As String
Dim rs As New ADODB.Recordset
Set rs = opndbaseFMIS.Execute("select top 1 rcenter,rcentercode,claimantcode,transactiondate,nonalobs,ooe from [tblAMIS_IncomingDVTrns] Where DVNo='" & Trim(txtDVNo.Text) & "'")
If Not rs.EOF Then
    If signal = 1 Then
        getdetails = Trim(rs(0))
    ElseIf signal = 2 Then
        getdetails = Trim(rs(1))
    ElseIf signal = 3 Then
        getdetails = Trim(rs(2))
    ElseIf signal = 4 Then
        getdetails = Trim(rs(3))
    ElseIf signal = 5 Then
        getdetails = Trim(rs(4))
    ElseIf signal = 6 Then
        getdetails = Trim(rs(5))

    End If
End If
End Function



Public Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim x As Integer
Dim xType As Integer, coloraly_signal As Integer

    Select Case Button:
    Case "New":
                XFlag = False
                CUFlag = False
                Edited = False
                xNAcode = ""
                lblMode.Caption = "NEW"
                txtDVNo.Text = ""
                txtAlobs.Text = ""
                txtClaimant.Text = ""
                txtClaimantCode.Text = ""
                txtRC.Text = ""
                txtParticular.Text = ""
                'txtFund.Text = ""
                txtAmount.Text = ""
                txtJEVNo.Text = ""
                txtdate.Text = Format(Now, "MMMM dd, yyyy")
                optCollection.Value = True
                chkSTP.Value = 0
                Call SetGrid
                
    Case "Save":
                If ChkEntry = True Then
                
                If CheckIfExistInFinalJEV(txtJEVNo.Text) = True Then
                    If MsgBox("JEV number already exist in the Database, Do you want System Generated JEV Number?", vbInformation + vbYesNo, "System Message") = vbYes Then
                        Call Command1_Click
                    End If
                End If
                
                    If MsgBox("Are you sure you want to save this transaction?", vbQuestion + vbYesNo) = vbYes Then
                        Call SaveAcctngEntries(txtJEVNo.Text)
                        Call GEtCompleteJEVDetails(txtJEVNo.Text, "Reffno", txtdate.Text, "", "" _
                            , Replace(txtParticular.Text, "'", "''"), txtJEVNo.Text, "", "", txtAmount.Text, "0", "0", 5, "", "", "", txtfund.Text, "", "", txtAlobs.Text, txtDVNo.Text, ExtractJEVSNo(txtJEVNo.Text), DatePost, "")
                        
                        MsgBox "Successfully Save", vbInformation, "System Message"
                        XFlag = False
                        CUFlag = False
                        Edited = False
                        xNAcode = ""
                        lblMode.Caption = "NEW"
                        txtDVNo.Text = ""
                        txtAlobs.Text = ""
                        txtClaimant.Text = ""
                        txtClaimantCode.Text = ""
                        txtRC.Text = ""
                        txtParticular.Text = ""
                        'txtFund.Text = ""
                        txtAmount.Text = ""
                        txtJEVNo.Text = ""
                        txtdate.Text = Format(Now, "MMMM dd, yyyy")
                        optCollection.Value = True
                        chkSTP.Value = 0
                        Call SetGrid
                    End If
                Else
                
debit_credit_error:
                    MsgBox "Save operation cancelled!" & vbCrLf & vbCrLf & "Please check your entry.", vbExclamation + vbOKOnly
                
                End If
    Case "Delete":
               
    Case "Close":
            Unload Me
    End Select
    
End Sub
Public Function SaveAcctngEntries(ByVal dvno As String)
Dim DRec As New ADODB.Recordset
Dim rec As New ADODB.Recordset
Dim x As Integer
Dim xType As Integer
If optCollection.Value = True Then xType = CInt(optCollection.Tag)
If optOther.Value = True Then xType = CInt(optOther.Tag)
    DRec.Open ("Select Accountcode,sum(Debit) as Debit ,sum(Credit) as Credit From tblAMIs_tmpjournal Where [dvno]='" & dvno & "' group by accountcode"), opndbaseFMIS, adOpenStatic, adLockOptimistic
    If DRec.RecordCount > 0 Then
        opndbaseFMIS.Execute "update tblAMIS_AccoutingEntries set actioncode =2 where reffno = '" & txtJEVNo.Text & "' and actioncode =1" ', datetimeentered = rtrim(ltrim(DateTimeEntered)) +'," & Format(Now, "yyyy/mm/dd hh:mm:ss AMPM") & "',userid = UserID + '," & Trim(ActiveUserID) & "'
        For x = 1 To DRec.RecordCount
            opndbaseFMIS.Execute "Insert into tblAMIS_AccoutingEntries (reffNo,ChildAccountcode,debit,credit,actioncode,datetimeentered,transtype,userid) values " & _
            "('" & Trim(txtJEVNo.Text) & "','" & Trim(DRec!accountcode) & "'," & DRec!Debit & "," & DRec!Credit & ",1,'" & Format(Now, "yyyy/mm/dd hh:mm:ss AMPM") & "'," & xType & ",'" & Trim(ActiveUserID) & "')"
            DRec.MoveNext
            DoEvents
        Next x
        opndbaseFMIS.Execute "delete from tblAMIs_tmpjournal where dvno = '" & txtJEVNo.Text & "'"
    End If
    DRec.Close
    Set DRec = Nothing
End Function
Private Function coloraly() As Boolean
Dim x As Integer
    For x = 1 To MSFlexGrid1.Rows - 1
        If MSFlexGrid1.TextMatrix(x, 2) <> "TOTAL" Then
            If MSFlexGrid1.TextMatrix(x, 5) <> "" Then
                If MSFlexGrid1.TextMatrix(x, 5) = "5" Then
                    coloraly = True
                    Exit Function
                End If
            End If
        Else
            Exit For
        End If
    Next x
End Function


Private Function ChkEntry() As Boolean

    ChkEntry = False
    If Trim(txtDVNo.Text) <> "" And txtParticular.Text <> "" And txtfund.Text <> "" And txtAmount.Text <> "" And txtJEVNo.Text <> "" Then
        If xDebit = xCredit And xDebit > 0 Then
        If coloraly = True Then GoTo coloraly_jmp 'coloraly consideration - set chkentry to true even if not balance
            If Format(xDebit, "###,##0.00") = Format(txtAmount.Text, "###,##0.00") Then
coloraly_jmp:
                ChkEntry = True
            End If
        End If
    End If
    
End Function

Private Sub LoadExcessDetails(ByVal ObR As String)
Dim OREc As New ADODB.Recordset
Dim x As Integer
Dim y As Integer

    Call SetGrid
    OREc.Open ("Select * from [tblBMS_ExcessControl] where AlobsNo='" & ObR & "' and ActionCode=1"), opndbaseFMIS, adOpenStatic, adLockOptimistic
    If OREc.RecordCount > 0 Then
        For x = 1 To OREc.RecordCount
            For y = 0 To cmbEntry.ListCount - 1
                If cmbEntry.List(y) = "401" Then
                    cmbEntry.ListIndex = y
                    Exit For
                Else
                    If y = cmbEntry.ListCount - 1 Then
                        cmbEntry.ListIndex = -1
                    End If
                End If
            Next y
            MSFlexGrid1.TextMatrix(x, 0) = cmbEntry.ItemData(cmbEntry.ListIndex)
            MSFlexGrid1.TextMatrix(x, 1) = "401"
            MSFlexGrid1.TextMatrix(x, 2) = GetAccountNameByFMISAccountCode(cmbEntry.ItemData(cmbEntry.ListIndex))
            MSFlexGrid1.TextMatrix(x, 4) = OREc!amount
            OREc.MoveNext
        Next x
        Call GetSum
    End If
    OREc.Close
    Set OREc = Nothing
    
End Sub


Private Sub LoadObRDetails(ByVal ObR As String)
Dim OREc As New ADODB.Recordset
Dim x As Integer
    
    Call SetGrid
    OREc.Open ("Select * from tblBMS_SubsidiaryLedger where AlobsNo='" & ObR & "' and ActionCode=1"), opndbaseFMIS, adOpenStatic, adLockOptimistic
    If OREc.RecordCount > 0 Then
        For x = 1 To OREc.RecordCount
            MSFlexGrid1.TextMatrix(x, 0) = OREc!FmisAccountcode
            MSFlexGrid1.TextMatrix(x, 1) = GetAccountCodeByFMISAccountCode(OREc!FmisAccountcode)
            MSFlexGrid1.TextMatrix(x, 2) = GetAccountNameByFMISAccountCode(OREc!FmisAccountcode)
            MSFlexGrid1.TextMatrix(x, 4) = OREc!amount
            OREc.MoveNext
        Next x
        Call GetSum
    End If
    OREc.Close
    Set OREc = Nothing
    
End Sub

Public Sub LoadAccountsByFund(ByVal fundmedium As String)
Dim ARec As New ADODB.Recordset
Dim x As Integer
Dim FundName As String

    cmbEntry.Clear
    cmbEntry.Visible = False
    If Left(fundmedium, 3) = "Eco" Then
    FundName = "Economic Enterprises"
    Else
    FundName = fundmedium
    End If
    ARec.Open ("Select distinct * from [tblREF_AIS_ChartofAccounts] Where [Active]=1 and [FundType]='" & FundName & "' Order by [ChildAccountCode]"), opndbaseFMIS, adOpenStatic, adLockOptimistic
    If ARec.RecordCount > 0 Then
        For x = 1 To ARec.RecordCount
            cmbEntry.AddItem ARec![childaccountcode]
            cmbEntry.ItemData(cmbEntry.NewIndex) = ARec![FmisAccountcode]
            ARec.MoveNext
        Next x
    End If
    ARec.Close
    Set ARec = Nothing
    
End Sub

Private Sub txt_entry_Change()
txtformula.Text = txt_entry.Text
End Sub

Private Sub txt_entry_KeyPress(KeyAscii As Integer)
Dim tamount As String
    If KeyAscii = 13 Then
        tamount = sumAmount(txt_entry.Text)
        MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, MSFlexGrid1.col) = tamount
        MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 6) = txt_entry.Text
        If MSFlexGrid1.col = 3 Then
            If Trim(txt_entry.Text) <> "" Then
                MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 4) = ""
            Else
                MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 3) = ""
            End If
        
        ElseIf MSFlexGrid1.col <> 5 Then
            
            If Trim(txt_entry.Text) <> "" Then
                MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 3) = ""
            Else
                MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 4) = ""
            End If
        End If
        txt_entry.Visible = False
        txtformula.Text = ""
        Call GetSum
    End If
End Sub
Private Function sumAmount(ByVal amnt As String) As String
On Error GoTo sum
Dim x As Integer
Dim y As String
Dim str() As String
    If Left(amnt, 1) = "+" Then
    Else
    amnt = "+" & amnt
    End If
 
 str = Split(Trim(amnt), "+", -1, vbTextCompare)
 y = 0

 For x = 1 To 1000
y = Val(y) + Val(str(x))
 Next x
 Exit Function
sum:
 If err.Number = 9 Then
 sumAmount = y
Else
MsgBox "Incorrect Format", vbInformation, "System Message"
End If
End Function

Private Sub GetSum()
On Error GoTo bad
Dim x As Integer
    not_coloraly_total_debit = 0
    not_coloraly_total_credit = 0
     coloraly_total_credit = 0
     coloraly_total_debit = 0
      
    xDebit = 0
    xCredit = 0
    For x = 1 To MSFlexGrid1.Rows - 1
        If MSFlexGrid1.TextMatrix(x, 1) <> "" Then
            xDebit = xDebit + CCur(IIf(MSFlexGrid1.TextMatrix(x, 3) = "", 0, MSFlexGrid1.TextMatrix(x, 3)))
            xCredit = xCredit + CCur(IIf(MSFlexGrid1.TextMatrix(x, 4) = "", 0, MSFlexGrid1.TextMatrix(x, 4)))
        Else
            MSFlexGrid1.TextMatrix(x, 2) = "TOTAL"
            MSFlexGrid1.TextMatrix(x, 3) = Format(xDebit, "#,##0.00")
            MSFlexGrid1.TextMatrix(x, 4) = Format(xCredit, "#,##0.00")
            Exit For
        End If
    Next x
Exit Sub
bad:
MsgBox err.description
End Sub

Private Function ChkIfAlreadyJEV(ByVal dvno As String) As String
Dim Jrec As New ADODB.Recordset

    ChkIfAlreadyJEV = ""
    Jrec.Open ("Select * from tblAMIS_COllectionDepositt where PTVNO='" & dvno & "' and (Actioncode=1 or Actioncode=5) "), opndbaseFMIS, adOpenStatic, adLockOptimistic
    If Jrec.RecordCount > 0 Then
'        If Not IsNull(JREc!ApprovedByID) Then
'            ChkIfAlreadyJEV = "Approved" & "-" & JREc!JEVNo
'        Else
            ChkIfAlreadyJEV = dvno
       ' End If
    End If
    Jrec.Close
    Set Jrec = Nothing
    
End Function
Public Sub loaddt()
Dim DVRec As New ADODB.Recordset
Dim xAlreadyJEV As String
        CUFlag = False
        txtParticular.Locked = True
        
        xNAcode = ""
        txtDVNo.Text = Trim(txtDVNo.Text)
        If ChkPTVExist(txtDVNo.Text) = True Then
            xAlreadyJEV = ChkIfAlreadyJEV(txtDVNo.Text)
            If xAlreadyJEV = "" Then
                DVRec.Open ("Select * FRom tblCMS_CDCheckBook where DVNo='" & txtDVNo.Text & "' and (ActionCode=1)"), opndbaseFMIS, adOpenStatic, adLockOptimistic
                If txtDVNo.Text <> "" Then
                If DVRec.RecordCount > 0 Then
                   ' If DVRec.Fields!closed = True Then
                        
                            txtAlobs.Text = DVRec!chknumber
                           ' txtClaimant.Text = GetClaimant(DVRec!ClaimantCode)
                            'txtClaimantCode.Text = DVRec!ClaimantCode
                           ' txtRC.Text = GetOfficeName(DVRec!RCenter, "OfficeMedium")
                            txtParticular.Text = DVRec!Particular
                            txtfund.Text = GetFundMedium(DVRec!fundcode)
                            txtAmount.Text = DVRec!amount
                            optCollection.Value = True
                            
                            Call optCollection_Click
                            Call LoadAccountsByFund(Trim(txtfund.Text))
                            
                            'XFlag = False
                            
                    
'                    Else
'                        MsgBox "Please Close PTV No. " & txtDVNo.Text & " On PTO first!", vbExclamation + vbOKOnly
'                        Call Toolbar1_ButtonClick(Toolbar1.Buttons.Item(1))
'                    End If
                End If
                End If
                DVRec.Close
                Set DVRec = Nothing
            Else
                List1.Text = xAlreadyJEV
                Call LoadJEVDetails(xAlreadyJEV)
            End If
        Else
            MsgBox "Invalid DV Number!", vbExclamation
            Call Toolbar1_ButtonClick(Toolbar1.Buttons.Item(1))
        End If

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

Private Sub txtformula_Change()
txt_entry.Text = txtformula.Text

End Sub

Private Sub txtformula_KeyPress(KeyAscii As Integer)
Dim tamount As String
If KeyAscii = 13 Then
        tamount = sumAmount(txt_entry.Text)
        MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, MSFlexGrid1.col) = tamount
        MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 6) = txt_entry.Text
        If MSFlexGrid1.col = 3 Then
            If Trim(txt_entry.Text) <> "" Then
                MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 4) = ""
            Else
                MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 3) = ""
            End If
        
        ElseIf MSFlexGrid1.col <> 5 Then
            
            If Trim(txt_entry.Text) <> "" Then
                MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 3) = ""
            Else
                MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 4) = ""
            End If
        End If
        txt_entry.Visible = False
        txtformula.Text = ""
        Call GetSum
    End If
End Sub

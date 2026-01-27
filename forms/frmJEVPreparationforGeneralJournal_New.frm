VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmJEVPreparationforGeneralJournal_New 
   Caption         =   "JEV Preparation For General Journal Through PTV Number"
   ClientHeight    =   9795
   ClientLeft      =   -150
   ClientTop       =   2865
   ClientWidth     =   14640
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   -1  'True
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmJEVPreparationforGeneralJournal_New.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9795
   ScaleWidth      =   14640
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
      Left            =   2880
      TabIndex        =   46
      Top             =   3960
      Visible         =   0   'False
      Width           =   8775
   End
   Begin VB.CommandButton btnReturn 
      Caption         =   "Return To PA"
      Enabled         =   0   'False
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
      Left            =   11040
      TabIndex        =   44
      Top             =   840
      Visible         =   0   'False
      Width           =   1665
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
      TabIndex        =   43
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
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   41
      Top             =   3240
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
      TabIndex        =   39
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
      TabIndex        =   31
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
      Left            =   5670
      Locked          =   -1  'True
      TabIndex        =   28
      Top             =   1350
      Width           =   2565
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
      TabIndex        =   27
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
      Left            =   420
      TabIndex        =   14
      Top             =   1920
      Width           =   11235
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
         TabIndex        =   45
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
         TabIndex        =   42
         ToolTipText     =   "Click here to edit particulars..."
         Top             =   240
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
         TabIndex        =   40
         ToolTipText     =   "Click here to select claimant..."
         Top             =   2400
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.TextBox txtFund 
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
         TabIndex        =   25
         Top             =   1440
         Width           =   4260
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
         Left            =   7560
         Locked          =   -1  'True
         TabIndex        =   23
         Top             =   1440
         Width           =   1860
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
         Locked          =   -1  'True
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
         TabIndex        =   26
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
         Left            =   5760
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
      Top             =   3225
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
      Height          =   5295
      Left            =   405
      ScaleHeight     =   5265
      ScaleWidth      =   11160
      TabIndex        =   3
      Top             =   4440
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
         Left            =   840
         TabIndex        =   38
         Text            =   "cmbEntry"
         Top             =   1200
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
         Left            =   5520
         TabIndex        =   30
         Top             =   840
         Visible         =   0   'False
         Width           =   1665
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   5280
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Width           =   11160
         _ExtentX        =   19685
         _ExtentY        =   9313
         _Version        =   393216
         FixedCols       =   0
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
      Width           =   14640
      _ExtentX        =   25823
      _ExtentY        =   1482
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
            Caption         =   "Save"
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
               Picture         =   "frmJEVPreparationforGeneralJournal_New.frx":076A
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmJEVPreparationforGeneralJournal_New.frx":20FC
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmJEVPreparationforGeneralJournal_New.frx":3A8E
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmJEVPreparationforGeneralJournal_New.frx":5420
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmJEVPreparationforGeneralJournal_New.frx":6DB2
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmJEVPreparationforGeneralJournal_New.frx":8744
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmJEVPreparationforGeneralJournal_New.frx":A0D6
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmJEVPreparationforGeneralJournal_New.frx":BA68
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmJEVPreparationforGeneralJournal_New.frx":D3FA
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmJEVPreparationforGeneralJournal_New.frx":ED8E
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmJEVPreparationforGeneralJournal_New.frx":FA6A
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmJEVPreparationforGeneralJournal_New.frx":1034A
               Key             =   ""
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmJEVPreparationforGeneralJournal_New.frx":11026
               Key             =   ""
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmJEVPreparationforGeneralJournal_New.frx":11D02
               Key             =   ""
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmJEVPreparationforGeneralJournal_New.frx":129DE
               Key             =   ""
            EndProperty
            BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmJEVPreparationforGeneralJournal_New.frx":136BA
               Key             =   ""
            EndProperty
            BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmJEVPreparationforGeneralJournal_New.frx":14396
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.TextBox txtJEVNo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
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
      Left            =   10440
      Locked          =   -1  'True
      TabIndex        =   35
      Top             =   240
      Visible         =   0   'False
      Width           =   3885
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
      Left            =   840
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
      Left            =   2520
      TabIndex        =   47
      Top             =   3960
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mode :"
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
      TabIndex        =   37
      Top             =   1290
      Width           =   825
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
      TabIndex        =   36
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
      TabIndex        =   33
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
      TabIndex        =   32
      Top             =   2505
      Width           =   915
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date Prepared"
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
      Left            =   5670
      TabIndex        =   29
      Top             =   1125
      Width           =   1035
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
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   435
      TabIndex        =   4
      Top             =   4200
      Width           =   1335
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Enter PTV  Number:"
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
      Left            =   390
      TabIndex        =   2
      Top             =   960
      Width           =   1425
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
      TabIndex        =   34
      Top             =   960
      Visible         =   0   'False
      Width           =   2640
   End
End
Attribute VB_Name = "frmJEVPreparationforGeneralJournal_New"
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
    sql = "Exec Proc_JevPrinting @dvno = '" & List1.Text & "'"
    ReportName = "JEVNEW"
    rptJEVNew.txtClaimDesc.SetText txtParticular.Text & ", " & txtClaimant.Text & ", " & txtAlobs.Text
    rptJEVNew.txtRC.SetText txtRC.Text
    rptJEVNew.txtClerk.SetText getUserName(ActiveUserID, "FullName")
    rptJEVNew.Trantype = 4
    If chkSTP.Value = 1 Then
        rptJEVNew.Line1.Suppress = True
        rptJEVNew.Line2.Suppress = True
        rptJEVNew.Line3.Suppress = True
        rptJEVNew.Line4.Suppress = True
        rptJEVNew.Line5.Suppress = True
        rptJEVNew.Line6.Suppress = True
        rptJEVNew.Line8.Suppress = True
        rptJEVNew.Line9.Suppress = True
        rptJEVNew.Line10.Suppress = True
        rptJEVNew.Line11.Suppress = True
        rptJEVNew.Line12.Suppress = True
        rptJEVNew.Line13.Suppress = True
        rptJEVNew.Line14.Suppress = True
        rptJEVNew.Line15.Suppress = True
        rptJEVNew.Line16.Suppress = True
        rptJEVNew.Line17.Suppress = True
        
        rptJEVNew.Line19.Suppress = True
        
        rptJEVNew.Text1.Suppress = True
        rptJEVNew.Text2.Suppress = True
        rptJEVNew.Text3.Suppress = True
        rptJEVNew.Text4.Suppress = True
        rptJEVNew.Text8.Suppress = True
        rptJEVNew.Text9.Suppress = True
        rptJEVNew.Text12.Suppress = True
        rptJEVNew.Text13.Suppress = True
        rptJEVNew.Text15.Suppress = True
        rptJEVNew.Text16.Suppress = True
        rptJEVNew.Text17.Suppress = True
        rptJEVNew.Text18.Suppress = True
        rptJEVNew.Text19.Suppress = True
        rptJEVNew.Text20.Suppress = True
        rptJEVNew.Text21.Suppress = True
        rptJEVNew.Text22.Suppress = True
        rptJEVNew.Text25.Suppress = True
        
    End If
    rptJEVNew.DiscardSavedData
    rptJEVNew.Database.SetDataSource opndbaseFMIS.Execute(sql)
    rptJEVNew.Database.Verify
   frmViewer.Show 1
End If
End Sub
Private Sub btnReturn_Click()
    If MsgBox("Are you sure you want to return DV No.: " & txtDVNo.Text & " to Pre-Audit?", vbQuestion + vbYesNo, "System Security") = vbYes Then
        If ChkIfAlreadyJEV(txtDVNo.Text) = "" Then
            opndbaseFMIS.Execute "Update [tblAMIS_IncomingDVTrns] set [PAout]=0, [PAoutDate]=null, [PADesc]=null, [OutBy]=null where [DVNo]='" & txtDVNo.Text & "' and actioncode=1"
        End If
        Call Toolbar1_ButtonClick(Toolbar1.Buttons.Item(1))
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
    PRec.Open ("Select top 500 ptvno, min(trnno) as trno  From tblAMIS_COllectionDepositt Where Actioncode in(1,5) and transtype = 4 and jevseriesno < 1 Group By ptvno order by trno desc"), opndbaseFMIS, adOpenStatic, adLockOptimistic
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
    If Trim(cmbrc.Text) <> "" Then
        txtRC = Trim(cmbrc.Text)
        txtRC.Visible = True
        cmbrc.Visible = False
    End If
End Sub

Private Sub cmbRC_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        cmbRC_Click
    End If
End Sub

Private Sub Form_Load()
    
    Edited = False
    
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2
    
    Call Toolbar1_ButtonClick(Toolbar1.Buttons.Item(1))
    
    ActiveUserID = Trim(ActiveUserID)
    Call loaddt
End Sub

Private Sub SetGrid()
Dim cc As Integer

    MSFlexGrid1.Clear
    MSFlexGrid1.Rows = 50
    MSFlexGrid1.Cols = 7 ' IIf(LCase(Trim(lblMode)) = "edit", 6, 5)
    
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
    
    'If LCase(Trim(lblMode)) = "Edit" Then
    MSFlexGrid1.ColWidth(5) = 1000
    MSFlexGrid1.ColWidth(6) = 0
    
    'Else
    '    MSFlexGrid1.ColWidth(5) = 0
    'End If
    
    
    For cc = 0 To MSFlexGrid1.Cols - 1
        MSFlexGrid1.Row = 0
        MSFlexGrid1.col = cc
        MSFlexGrid1.CellAlignment = 4
    Next cc
End Sub

Private Sub List1_Click()
    Call LoadJEVDetails(List1.Text)
    cmbrc.Visible = False
   ' txtRC.Visible = True
End Sub

Private Sub LoadJEVDetails(ByVal DVNo As String)
Dim DRec As New ADODB.Recordset
Dim x As Integer
    
    CUFlag = False
    txtParticular.Locked = True
    xNAcode = ""
    Edited = True
    lblMode.Caption = "EDIT"
    DRec.Open ("Select * From [tblAMIS_COllectionDepositt] Where [ptvno]='" & DVNo & "' And (ActionCode=1 or ActionCode=5)"), opndbaseFMIS, adOpenStatic, adLockOptimistic
    If DRec.RecordCount > 0 Then
        txtDVNo.Text = DRec![ptvNo]
        txtdate.Text = DRec![TransDate]
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
         txtAmount.Text = DRec!AMOUNT
    End If
    DRec.Close
    Set DRec = Nothing
    Call LoadAcctngEntries(txtDVNo)
End Sub
Public Function LoadAcctngEntries(ByVal DVNo As String)
Dim DRec As New ADODB.Recordset
Dim x As Integer
Call SetGrid
    DRec.Open ("Select left(ChildAccountcode,3) as childcode,sum(Debit) as sumdebit,sum(credit) as sumcredit From tblAMIS_AccoutingEntries Where [reffno]='" & DVNo & "' And (ActionCode=1) group by reffno,actioncode,left(ChildAccountcode,3)"), opndbaseFMIS, adOpenStatic, adLockOptimistic
    If DRec.RecordCount > 0 Then
        For x = 1 To DRec.RecordCount
'            MSFlexGrid1.TextMatrix(x, 0) = DRec![Trnno]
            MSFlexGrid1.TextMatrix(x, 1) = DRec!childcode
            MSFlexGrid1.TextMatrix(x, 2) = GetAccountNameByAccountcode(DRec!childcode)
            MSFlexGrid1.TextMatrix(x, 4) = IIf((Format(DRec!sumCredit, "#,##0.00") = "0.00"), "", Format(DRec!sumCredit, "#,##0.00"))
            MSFlexGrid1.TextMatrix(x, 3) = IIf((Format(DRec!sumDebit, "#,##0.00") = "0.00"), "", Format(DRec!sumDebit, "#,##0.00"))
          
           ' If LCase(Trim(lblMode.Caption)) = "edit" Then MSFlexGrid1.TextMatrix(x, 5) = DRec!ActionCode  ' for coloraly purpose
            DRec.MoveNext
        Next x
        Call GetSum
    End If
    DRec.Close
    Set DRec = Nothing
End Function
Private Sub MSFlexGrid1_Click()
If Trim(txtAlobs.Text) <> "" Then
    With frmSub3
        .reff = txtDVNo.Text
        .Show 1
        Call LoadAcctngEntries(Trim(txtDVNo.Text))
    End With
End If
End Sub

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
        GetNewJEV = xCode & "-" & Format(CInt(Right(Jrec!JEVNO, 3)) + 1, "000")
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
                txtfund.Text = ""
                txtAmount.Text = ""
                txtJEVno.Text = ""
                txtdate.Text = Format(Now, "MMMM dd, yyyy")
                optCollection.Value = True
                chkSTP.Value = 0
                btnReturn.Enabled = False
                
                Call LoadTrnYear(cmb_trnYear)
                Call LoadTrnMonth(cmb_month)
                Call SetGrid
                
    Case "Save":
                If ChkEntry = True Then
                    If MsgBox("Are you sure you want to save this transaction?", vbQuestion + vbYesNo) = vbYes Then

                        If Edited = True Then
                            opndbaseFMIS.Execute "Update tblAMIS_COllectionDepositt set ActionCode=2, UserID=UserID + '," & ActiveUserID & "', DateTimeEntered=DateTimeEntered + '," & Format(Now, "yyyy/mm/dd hh:mm:ss AMPM") & "' Where ptvno='" & List1.Text & "' And ActionCode=1"
                        End If
                        
                        If CUFlag = True Then
                           opndbaseFMIS.Execute "Update [tblCMS_CDCheckBook] set [Particular]='" & Trim(Replace(txtParticular.Text, "'", "''")) & "' Where DVNo='" & Trim(txtDVNo.Text) & "' And ActionCode=1"
                        End If
                        If xNAcode <> "" Then
                            xObR = xNAcode
                        End If
                        opndbaseFMIS.Execute "Insert Into tblAMIS_COllectionDepositt (TransType,ptvno,reportno,FmisAccntCode,amount,debitcredit,TransDate,UserID,Actioncode,DateTimeEntered,IsNew) values (4,'" & Trim(Replace(txtDVNo.Text, "'", "''")) & "','" & xObR & "',0,0,0,'" & Format(Now, "yyyy/mm/dd hh:mm:ss AMPM") & "','" & ActiveUserID & "',1,'" & Format(Now, "yyyy/mm/dd hh:mm:ss AMPM") & "',1)"
                        Call Toolbar1_ButtonClick(Toolbar1.Buttons.Item(1))
                    End If
                Else
                End If
    Case "Delete":
                If Edited = True Then
                    If InStr(ChkIfAlreadyJEV(txtDVNo.Text), "Approved") <> 1 Then
                        If MsgBox("Are you sure you want to delete this transaction?", vbQuestion + vbYesNo) = vbYes Then
                            opndbaseFMIS.Execute "Update tblAMIS_COllectionDepositt set UserID=UserID + '," & ActiveUserID & "',Actioncode=3,DateTimeEntered=DateTimeEntered +'," & Format(Now, "mm/dd/yyyy hh:mm:ss AMPM") & "'  Where PTVNO='" & txtDVNo.Text & "' and Actioncode=1"
                            opndbaseFMIS.Execute "Update  tblAMIS_AccoutingEntries set Actioncode=3 Where reffno='" & txtDVNo.Text & "' and Actioncode=1"""
                            Call Toolbar1_ButtonClick(Toolbar1.Buttons(1))
                        End If
                    Else
                        MsgBox "This transaction is already approved!" & vbCrLf & vbCrLf & "Delete operation cancelled!", vbExclamation + vbOKOnly
                    End If
                End If
    Case "Close":
                If MsgBox("Are you sure you want to close this form?", vbQuestion + vbYesNo, "System Security") = vbYes Then
                    Unload Me
                End If
    End Select
    
    
End Sub

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
    If Trim(txtDVNo.Text) <> "" And txtAlobs.Text <> "" And txtParticular.Text <> "" And txtfund.Text <> "" And txtAmount.Text <> "" Then
        If xDebit = xCredit And xDebit > 0 Then
            If Format(xDebit, "###,##0.00") = Format((2 * CCur(txtAmount.Text)), "###,##0.00") Then
                ChkEntry = True
            Else
            MsgBox "Incorrect Entry,Please Check...!", vbInformation, "System Message"
            End If
        Else
        MsgBox "Your Total Debit and Total Credit Amount is not Equal or Accounting Entry is Empty, Please Check Your Entry", vbInformation, "System Message"
        End If
    Else
    MsgBox "Some Fields are Empty, Please check it..!", vbInformation, "System Message"
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
            MSFlexGrid1.TextMatrix(x, 4) = OREc!AMOUNT
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
            MSFlexGrid1.TextMatrix(x, 4) = OREc!AMOUNT
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
    FundName = GetFundName(fundmedium)
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
        If MSFlexGrid1.TextMatrix(x, 0) <> "" Then
            xDebit = xDebit + CCur(IIf(MSFlexGrid1.TextMatrix(x, 3) = "", 0, MSFlexGrid1.TextMatrix(x, 3)))
            xCredit = xCredit + CCur(IIf(MSFlexGrid1.TextMatrix(x, 4) = "", 0, MSFlexGrid1.TextMatrix(x, 4)))
                If Trim(MSFlexGrid1.TextMatrix(x, 5)) <> 5 Then
                    not_coloraly_total_debit = not_coloraly_total_debit + CCur(IIf(MSFlexGrid1.TextMatrix(x, 3) = "", 0, MSFlexGrid1.TextMatrix(x, 3)))
                    not_coloraly_total_credit = not_coloraly_total_credit + CCur(IIf(MSFlexGrid1.TextMatrix(x, 4) = "", 0, MSFlexGrid1.TextMatrix(x, 4)))
                Else
                    coloraly_total_debit = coloraly_total_debit + CCur(IIf(MSFlexGrid1.TextMatrix(x, 3) = "", 0, MSFlexGrid1.TextMatrix(x, 3)))
                    coloraly_total_credit = coloraly_total_credit + CCur(IIf(MSFlexGrid1.TextMatrix(x, 4) = "", 0, MSFlexGrid1.TextMatrix(x, 4)))
                End If
        Else
            MSFlexGrid1.TextMatrix(x, 2) = "TOTAL"
            MSFlexGrid1.TextMatrix(x, 3) = xDebit
            MSFlexGrid1.TextMatrix(x, 4) = xCredit
            Exit For
        End If
    Next x
Exit Sub
bad:
MsgBox err.Description
End Sub

Private Function ChkIfAlreadyJEV(ByVal DVNo As String) As String
Dim Jrec As New ADODB.Recordset

    ChkIfAlreadyJEV = ""
    Jrec.Open ("Select * from tblAMIS_COllectionDepositt where PTVNO='" & DVNo & "' and (Actioncode=1 or Actioncode=5) "), opndbaseFMIS, adOpenStatic, adLockOptimistic
    If Jrec.RecordCount > 0 Then
'        If Not IsNull(JREc!ApprovedByID) Then
'            ChkIfAlreadyJEV = "Approved" & "-" & JREc!JEVNo
'        Else
            ChkIfAlreadyJEV = DVNo
       ' End If
    End If
    Jrec.Close
    Set Jrec = Nothing
    
End Function

Private Sub txtDVNo_KeyPress(KeyAscii As Integer)
Dim DVRec As New ADODB.Recordset
Dim xAlreadyJEV As String

    If KeyAscii = 13 Then
        btnReturn.Enabled = False
        CUFlag = False
        txtParticular.Locked = True
        
        xNAcode = ""
        txtDVNo.Text = Trim(txtDVNo.Text)
        If ChkPTVExist(txtDVNo.Text) = True Then
            xAlreadyJEV = ChkIfAlreadyJEV(txtDVNo.Text)
            If xAlreadyJEV = "" Then
                DVRec.Open ("Select * FRom tblCMS_CDCheckBook where DVNo='" & txtDVNo.Text & "' and (ActionCode=1)"), opndbaseFMIS, adOpenStatic, adLockOptimistic
                If DVRec.RecordCount > 0 Then
                   ' If DVRec.Fields!closed = True Then
                        
                            txtAlobs.Text = DVRec!chknumber
                           ' txtClaimant.Text = GetClaimant(DVRec!ClaimantCode)
                            'txtClaimantCode.Text = DVRec!ClaimantCode
                           ' txtRC.Text = GetOfficeName(DVRec!RCenter, "OfficeMedium")
                            txtParticular.Text = DVRec!Particular
                            txtfund.Text = GetFundMedium(DVRec!fundcode)
                            txtAmount.Text = DVRec!AMOUNT
                            optCollection.Value = True
                            
                            Call optCollection_Click
                            Call LoadAccountsByFund(Trim(txtfund.Text))
                            
                            'XFlag = False
                            
                    
'                    Else
'                        MsgBox "Please Close PTV No. " & txtDVNo.Text & " On PTO first!", vbExclamation + vbOKOnly
'                        Call Toolbar1_ButtonClick(Toolbar1.Buttons.Item(1))
'                    End If
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
    End If
End Sub
Public Sub loaddt()
Dim DVRec As New ADODB.Recordset
Dim xAlreadyJEV As String

   
        btnReturn.Enabled = False
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
                            txtAmount.Text = DVRec!AMOUNT
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

cmbrc.Clear

OREc.Open ("Select distinct * FRom tblREF_AIS_Offices Order By [OfficeMedium]"), opndbaseFMIS, adOpenStatic, adLockOptimistic
If OREc.RecordCount > 0 Then
    For x = 1 To OREc.RecordCount
        cmbrc.AddItem OREc![OfficeMedium]
        cmbrc.ItemData(cmbrc.NewIndex) = OREc!fmisofficeid
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

Private Sub txtRC_Click()
'   cmbRC.Visible = True
    'txtRC.Visible = False
End Sub

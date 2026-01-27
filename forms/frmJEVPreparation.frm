VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form frm_AccountsPayableEntry 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Accounts Payable Entry"
   ClientHeight    =   9780
   ClientLeft      =   -165
   ClientTop       =   2850
   ClientWidth     =   14685
   Icon            =   "frmJEVPreparation.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "frmJEVPreparation.frx":076A
   ScaleHeight     =   9780
   ScaleWidth      =   14685
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4470
      Left            =   285
      ScaleHeight     =   4440
      ScaleWidth      =   11640
      TabIndex        =   45
      Top             =   5130
      Width           =   11670
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSFlexGrid1 
         Height          =   4455
         Left            =   0
         TabIndex        =   48
         Top             =   0
         Width           =   11655
         _ExtentX        =   20558
         _ExtentY        =   7858
         _Version        =   393216
         ScrollTrack     =   -1  'True
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.TextBox txt_entry 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   525
         Left            =   6240
         TabIndex        =   47
         Top             =   1680
         Visible         =   0   'False
         Width           =   1665
      End
      Begin VB.ComboBox cmbEntry 
         Height          =   315
         Left            =   6480
         TabIndex        =   46
         Text            =   "cmbEntry"
         Top             =   960
         Visible         =   0   'False
         Width           =   1215
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Corresponding  Code"
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
      Left            =   240
      TabIndex        =   42
      Top             =   3960
      Width           =   4575
      Begin VB.TextBox txtaccountcode 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,#0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
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
         Left            =   2160
         TabIndex        =   43
         Top             =   360
         Width           =   1620
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Accountcode"
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
         Left            =   840
         TabIndex        =   44
         Top             =   480
         Width           =   1185
      End
   End
   Begin MSComctlLib.ImageList itb32x32 
      Left            =   7800
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
            Picture         =   "frmJEVPreparation.frx":0CF4
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJEVPreparation.frx":2686
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJEVPreparation.frx":4018
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJEVPreparation.frx":59AA
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJEVPreparation.frx":733C
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJEVPreparation.frx":8CCE
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJEVPreparation.frx":A660
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJEVPreparation.frx":BFF2
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJEVPreparation.frx":D984
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJEVPreparation.frx":F318
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJEVPreparation.frx":FFF4
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJEVPreparation.frx":108D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJEVPreparation.frx":115B0
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJEVPreparation.frx":1228C
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJEVPreparation.frx":12F68
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJEVPreparation.frx":13C44
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJEVPreparation.frx":14920
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command2 
      Caption         =   "ADD Account Title"
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
      Left            =   11760
      TabIndex        =   34
      Top             =   840
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.CommandButton btnReturn 
      Caption         =   "Return To PA"
      Enabled         =   0   'False
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
      Left            =   10440
      TabIndex        =   33
      Top             =   840
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.CommandButton btnSearch 
      Caption         =   "Search"
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
      Left            =   13200
      TabIndex        =   32
      Top             =   840
      Width           =   1305
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
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   30
      Top             =   3120
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.CheckBox chkSTP 
      Caption         =   "Shoot-To-Print"
      Height          =   255
      Left            =   12240
      TabIndex        =   28
      Top             =   8880
      Width           =   1575
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
      Left            =   13230
      TabIndex        =   22
      Top             =   4590
      Width           =   1230
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
      TabIndex        =   20
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
      Left            =   13230
      TabIndex        =   19
      Top             =   4170
      Width           =   1230
   End
   Begin VB.Frame Frame2 
      Caption         =   "Transaction Details"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1965
      Left            =   240
      TabIndex        =   7
      Top             =   1920
      Width           =   14355
      Begin VB.TextBox txtrcenter 
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
         Height          =   360
         Left            =   9840
         TabIndex        =   49
         Top             =   600
         Width           =   3780
      End
      Begin VB.CommandButton Command7 
         Caption         =   "..."
         Height          =   255
         Left            =   13680
         TabIndex        =   35
         ToolTipText     =   "Click here to edit particulars..."
         Top             =   1320
         Width           =   255
      End
      Begin VB.CommandButton btnParticular 
         Caption         =   "..."
         Height          =   255
         Left            =   9120
         TabIndex        =   31
         ToolTipText     =   "Click here to edit particulars..."
         Top             =   240
         Width           =   255
      End
      Begin VB.CommandButton btnClaimant 
         Caption         =   "..."
         Height          =   255
         Left            =   4680
         TabIndex        =   29
         ToolTipText     =   "Click here to select claimant..."
         Top             =   1320
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.TextBox txtFund 
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
         Height          =   360
         Left            =   9870
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   1305
         Width           =   1860
      End
      Begin VB.TextBox txtAmount 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,#0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
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
         Left            =   12030
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   1275
         Width           =   1620
      End
      Begin VB.TextBox txtParticular 
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
         Height          =   1200
         Left            =   5160
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   13
         Top             =   540
         Width           =   4290
      End
      Begin VB.TextBox txtAlobs 
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
         Height          =   360
         Left            =   315
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   540
         Width           =   4260
      End
      Begin VB.TextBox txtClaimant 
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
         Height          =   360
         Left            =   315
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   1305
         Width           =   4260
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fund Type"
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
         Left            =   9840
         TabIndex        =   18
         Top             =   990
         Width           =   975
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Amount (Gross)"
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
         Left            =   12060
         TabIndex        =   16
         Top             =   1050
         Width           =   1380
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Particular"
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
         Left            =   5100
         TabIndex        =   14
         Top             =   280
         Width           =   840
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Alobs/OBR No:"
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
         Left            =   180
         TabIndex        =   12
         Top             =   270
         Width           =   1380
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Claimant"
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
         Left            =   180
         TabIndex        =   10
         Top             =   1050
         Width           =   780
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Responsibility Center"
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
         Left            =   9780
         TabIndex        =   9
         Top             =   239
         Width           =   1905
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
      Height          =   3420
      Left            =   12240
      TabIndex        =   5
      Top             =   5400
      Width           =   2265
   End
   Begin VB.CommandButton btnPrtJEV 
      Caption         =   "Print JEV"
      Height          =   360
      Left            =   12225
      TabIndex        =   4
      Top             =   9240
      Width           =   2295
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
      Left            =   10320
      Locked          =   -1  'True
      TabIndex        =   26
      Top             =   240
      Visible         =   0   'False
      Width           =   4125
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   840
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   14685
      _ExtentX        =   25903
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
   End
   Begin VB.Frame Frame1 
      Caption         =   "JEV Transaction Type"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   120
      Left            =   360
      TabIndex        =   37
      Top             =   5640
      Visible         =   0   'False
      Width           =   11670
      Begin VB.OptionButton optCollection 
         Caption         =   "Collection"
         Height          =   195
         Left            =   270
         TabIndex        =   41
         Tag             =   "01"
         Top             =   1125
         Value           =   -1  'True
         Width           =   1260
      End
      Begin VB.OptionButton optCheck 
         Caption         =   "Check Disbursement"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   165
         TabIndex        =   40
         Tag             =   "02"
         Top             =   420
         Width           =   2580
      End
      Begin VB.OptionButton optCash 
         Caption         =   "Cash Disbursement"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2925
         TabIndex        =   39
         Tag             =   "03"
         Top             =   420
         Width           =   2460
      End
      Begin VB.OptionButton optOther 
         Caption         =   "General Journal"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   5565
         TabIndex        =   38
         Tag             =   "04"
         Top             =   420
         Width           =   2190
      End
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
      Left            =   8640
      TabIndex        =   36
      Top             =   1320
      Width           =   825
   End
   Begin VB.Label lblMode 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Edit"
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
      Left            =   9480
      TabIndex        =   27
      Top             =   1320
      Width           =   615
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "Trn Year :"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   12420
      TabIndex        =   24
      Top             =   4245
      Width           =   705
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "Month of:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   12420
      TabIndex        =   23
      Top             =   4665
      Width           =   675
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
      Left            =   5670
      TabIndex        =   21
      Top             =   1125
      Width           =   1335
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   885
      Left            =   12240
      Top             =   4110
      Width           =   2235
   End
   Begin VB.Label Label3 
      Caption         =   "Vouchers Prepared with JEV"
      Height          =   225
      Left            =   12270
      TabIndex        =   6
      Top             =   5115
      Width           =   2190
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Accounting Entries"
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
      Left            =   315
      TabIndex        =   3
      Top             =   3540
      Width           =   2295
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Enter DV Number:"
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
      Left            =   390
      TabIndex        =   2
      Top             =   960
      Width           =   1605
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   975
      Left            =   8640
      Top             =   840
      Width           =   1755
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
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   5400
      TabIndex        =   25
      Top             =   960
      Visible         =   0   'False
      Width           =   2640
   End
End
Attribute VB_Name = "frm_AccountsPayableEntry"
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
Dim rcedit As Boolean
Dim CAnew As Boolean
Dim CAedit As Boolean
Dim ifsaveamount As Boolean
Dim ifColoraly As Boolean
Public isfrom_jevNumbering As Boolean
Dim not_coloraly_total_debit, not_coloraly_total_credit, coloraly_total_debit, coloraly_total_credit As Double

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
    PRec.Open ("Select DVNo From tblAMIS_AccountspayableEntry"), opndbaseFMIS, adOpenStatic, adLockOptimistic
    If PRec.RecordCount > 0 Then
        For x = 1 To PRec.RecordCount
            List1.AddItem PRec!dvno
            PRec.MoveNext
            DoEvents
        Next x
        List1.Enabled = True
    End If
    PRec.Close
    Set PRec = Nothing
    
End Sub

Private Sub Form_Load()
    Edited = False
    
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2
    
    Call Toolbar1_ButtonClick(Toolbar1.Buttons.Item(1))
    
    ActiveUserID = Trim(ActiveUserID)
 
    LoadPrevTransInGrid
End Sub
Private Sub LoadPrevTransInGrid()
Dim PRec As New ADODB.Recordset
Dim x As Integer

    SetGrid
    PRec.Open ("Exec MPproc_LoadAccntspayableEntry"), opndbaseFMIS, adOpenStatic, adLockOptimistic
    If PRec.RecordCount > 0 Then
        Set MSFlexGrid1.DataSource = PRec
    End If
    PRec.Close
    Set PRec = Nothing
End Sub

Private Sub SetGrid()
Dim cc As Integer

    MSFlexGrid1.Clear
    MSFlexGrid1.Cols = 6 ' IIf(LCase(Trim(lblMode)) = "edit", 6, 5)

    MSFlexGrid1.TextMatrix(0, 1) = "Account Code"
    MSFlexGrid1.TextMatrix(0, 2) = "Accounts and Explanation"
    MSFlexGrid1.TextMatrix(0, 3) = "Debit"
    MSFlexGrid1.TextMatrix(0, 4) = "Credit"

    MSFlexGrid1.ColWidth(0) = 0
    MSFlexGrid1.ColWidth(1) = 3000
    MSFlexGrid1.ColWidth(2) = 5250
    MSFlexGrid1.ColWidth(3) = 1500
    MSFlexGrid1.ColWidth(4) = 1500
    'If LCase(Trim(lblMode)) = "Edit" Then
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
On Error Resume Next
Dim rec As New ADODB.Recordset
    Set rec = opndbaseFMIS.Execute("EXECUTE  [fmis].[dbo].[MPproc_LoadAccntspayableDetails] @dvno = '" & List1.Text & "'")
        With rec
            If .RecordCount > 0 Then
                txtAlobs.Text = Trim(IIf(IsNull(.Fields!obrno), "", .Fields!obrno))
                txtParticular.Text = Trim(IIf(IsNull(.Fields!Particular), "", .Fields!Particular))
                txtrcenter.Text = Trim(GetOfficeName(.Fields!RCenter, "OfficeMedium"))
                txtFund.Text = Trim(IIf(IsNull(.Fields!FundType), "", .Fields!FundType))
                txtaccountcode.Text = Trim(IIf(IsNull(.Fields!accountcode), "", .Fields!accountcode))
                txtAmount.Text = Trim(IIf(IsNull(.Fields!amount), "", .Fields!amount))
                txtClaimant.Text = Trim(getClaimant(.Fields!Claimant))
                txtDVNo.Text = List1.Text
                Else
                MsgBox "Invalid DV Number!", vbExclamation
                Call Toolbar1_ButtonClick(Toolbar1.Buttons.Item(1))
            End If
        End With
End Sub
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
On Error GoTo bad

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
                txtParticular.Text = ""
                txtFund.Text = ""
                txtAmount.Text = ""
                txtJEVNo.Text = ""
                txtdate.Text = Format(Now, "MMMM dd, yyyy")
                optCollection.Value = True
                chkSTP.Value = 0
                btnReturn.Enabled = False
                txtaccountcode.Text = ""
                Call LoadTrnYear(cmb_trnYear)
                LoadPrevTransInGrid
                LoadPrevTrans
                Call LoadTrnMonth(cmb_month)
    Case "Save":
                Dim rec As New ADODB.Recordset
                Dim crec As New ADODB.Recordset
                Set crec = opndbaseFMIS.Execute("Select * from tblREF_AIS_ChartOfAccountsMother where accountcode = '" & txtaccountcode.Text & "'")
                If crec.RecordCount > 0 Then
                    If MsgBox("Are you sure you want to Save this transaction?", vbQuestion + vbYesNo) = vbYes Then
                        Set rec = opndbaseFMIS.Execute("Select * from tblAMIS_AccountspayableEntry where dvno = '" & txtDVNo.Text & "'")
                        If rec.RecordCount > 0 Then
                            opndbaseFMIS.Execute ("Delete from tblAMIS_AccountspayableEntry where dvno = '" & txtDVNo.Text & "'")
                        End If
                            opndbaseFMIS.Execute "Insert into tblAMIS_AccountspayableEntry (dvno,amount,accountcode) values ('" & txtDVNo.Text & "','" & txtAmount.Text & "','" & txtaccountcode.Text & "')"
                        rec.Close
                    End If
                    Set rec = Nothing
                    Call Toolbar1_ButtonClick(Toolbar1.Buttons(1))
                Else
                MsgBox "Invalid Accountcode", vbCritical, "System Message"
                End If
    Case "Delete":
                If MsgBox("Are you sure you want to delete this transaction?", vbQuestion + vbYesNo) = vbYes Then
                   opndbaseFMIS.Execute ("Delete from tblAMIS_AccountspayableEntry where dvno = '" & txtDVNo.Text & "'")
                   Call Toolbar1_ButtonClick(Toolbar1.Buttons(1))
                End If
    Case "Close":
                If MsgBox("Are you sure you want to close this form?", vbQuestion + vbYesNo, "System Security") = vbYes Then
                    Unload Me
                End If
    End Select
   
Exit Sub
bad:
    Call LoadErr(err.Number, err.Source & ", " & Me.name & ", " & Me.Caption, err.description)
End Sub
Private Function ChkIfAlreadyJEV(ByVal dvno As String) As String
Dim Jrec As New ADODB.Recordset

    ChkIfAlreadyJEV = ""
    Jrec.Open ("Select * from tblAMIS_JournalEntry where DVNo='" & dvno & "' and Actioncode=1"), opndbaseFMIS, adOpenStatic, adLockOptimistic
    If Jrec.RecordCount > 0 Then
        If Not IsNull(Jrec!ApprovedByID) Then
            ChkIfAlreadyJEV = "Approved" & "-" & Jrec!JEVNO
        Else
            ChkIfAlreadyJEV = dvno
        End If
    End If
    Jrec.Close
    Set Jrec = Nothing
    
End Function

Private Sub txtAmount_LostFocus()
txtAmount.Locked = True
txtAmount.Text = Format(txtAmount.Text, "#,##0.00")
End Sub
    
Private Sub txtDVNo_KeyPress(KeyAscii As Integer)
Dim DVRec As New ADODB.Recordset
Dim xAlreadyJEV As String
On Error Resume Next
    If KeyAscii = 13 Then
        Set DVRec = opndbaseFMIS.Execute("EXECUTE  [fmis].[dbo].[MPproc_LoadAccntspayableDetails] @dvno = '" & txtDVNo.Text & "'")
        With DVRec
            If .RecordCount > 0 Then
                txtAlobs.Text = Trim(IIf(IsNull(.Fields!obrno), "", .Fields!obrno))
                txtParticular.Text = Trim(IIf(IsNull(.Fields!Particular), "", .Fields!Particular))
                txtrcenter.Text = Trim(GetOfficeName(.Fields!RCenter, "OfficeMedium"))
                txtFund.Text = Trim(IIf(IsNull(.Fields!FundType), "", .Fields!FundType))
                txtaccountcode.Text = Trim(IIf(IsNull(.Fields!accountcode), "", .Fields!accountcode))
                txtAmount.Text = Trim(IIf(IsNull(.Fields!amount), "", .Fields!amount))
                txtClaimant.Text = Trim(getClaimant(.Fields!Claimant))
                txtaccountcode.SetFocus
            Else
                MsgBox "Invalid DV Number!", vbExclamation
                Call Toolbar1_ButtonClick(Toolbar1.Buttons.Item(1))
            End If
        End With
    End If
End Sub





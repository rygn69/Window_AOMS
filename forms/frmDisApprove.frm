VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "Mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.ocx"
Begin VB.Form frmJEVDisapprove 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " "
   ClientHeight    =   4455
   ClientLeft      =   -165
   ClientTop       =   2850
   ClientWidth     =   14715
   Icon            =   "frmDisApprove.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmDisApprove.frx":076A
   ScaleHeight     =   4455
   ScaleWidth      =   14715
   Begin VB.CheckBox chkcancel 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Cancel Transaction"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   10440
      TabIndex        =   77
      Top             =   1080
      Width           =   1815
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Return to Pre-Audit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8760
      TabIndex        =   76
      Top             =   1080
      Width           =   1575
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
            Picture         =   "frmDisApprove.frx":0CF4
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDisApprove.frx":2686
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDisApprove.frx":4018
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDisApprove.frx":59AA
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDisApprove.frx":733C
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDisApprove.frx":8CCE
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDisApprove.frx":A660
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDisApprove.frx":BFF2
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDisApprove.frx":D984
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDisApprove.frx":F318
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDisApprove.frx":FFF4
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDisApprove.frx":108D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDisApprove.frx":115B0
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDisApprove.frx":1228C
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDisApprove.frx":12F68
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDisApprove.frx":13C44
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDisApprove.frx":14920
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
      Left            =   17760
      TabIndex        =   68
      Top             =   840
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.PictureBox lblMode 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   15600
      ScaleHeight     =   315
      ScaleWidth      =   675
      TabIndex        =   44
      Top             =   1245
      Visible         =   0   'False
      Width           =   735
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
      Left            =   14760
      TabIndex        =   42
      Top             =   480
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
      Left            =   16080
      TabIndex        =   41
      Top             =   960
      Visible         =   0   'False
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
      TabIndex        =   39
      Top             =   3120
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.CheckBox chkSTP 
      Caption         =   "Shoot-To-Print"
      Height          =   255
      Left            =   12360
      TabIndex        =   37
      Top             =   9720
      Visible         =   0   'False
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
      TabIndex        =   30
      Top             =   11430
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
      TabIndex        =   27
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
      TabIndex        =   26
      Top             =   11010
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
      Height          =   2205
      Left            =   120
      TabIndex        =   14
      Top             =   1920
      Width           =   14475
      Begin VB.CommandButton Command7 
         Caption         =   "..."
         Height          =   255
         Left            =   13680
         TabIndex        =   74
         ToolTipText     =   "Click here to edit particulars..."
         Top             =   1320
         Visible         =   0   'False
         Width           =   255
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
         TabIndex        =   43
         Top             =   600
         Width           =   4095
      End
      Begin VB.CommandButton btnParticular 
         Caption         =   "..."
         Height          =   255
         Left            =   9120
         TabIndex        =   40
         ToolTipText     =   "Click here to edit particulars..."
         Top             =   240
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.CommandButton btnClaimant 
         Caption         =   "..."
         Height          =   255
         Left            =   4680
         TabIndex        =   38
         ToolTipText     =   "Click here to select claimant..."
         Top             =   1320
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
         Left            =   9870
         Locked          =   -1  'True
         TabIndex        =   24
         Top             =   1305
         Width           =   1860
      End
      Begin VB.TextBox txtAmount 
         Alignment       =   2  'Center
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
         TabIndex        =   22
         Top             =   1275
         Width           =   1620
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
         Height          =   1200
         Left            =   5160
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   20
         Top             =   540
         Width           =   4290
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
         Left            =   315
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   600
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
         TabIndex        =   15
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
         TabIndex        =   25
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
         TabIndex        =   23
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
         TabIndex        =   21
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
         TabIndex        =   19
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
         TabIndex        =   17
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
         TabIndex        =   16
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
      Height          =   4140
      Left            =   12240
      TabIndex        =   12
      Top             =   12240
      Width           =   2265
   End
   Begin VB.CommandButton btnPrtJEV 
      Caption         =   "Print JEV"
      Height          =   360
      Left            =   12225
      TabIndex        =   11
      Top             =   15240
      Visible         =   0   'False
      Width           =   2295
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
      Height          =   840
      Left            =   435
      TabIndex        =   6
      Top             =   10665
      Width           =   11670
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
         TabIndex        =   10
         Tag             =   "04"
         Top             =   420
         Width           =   2190
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
         TabIndex        =   9
         Tag             =   "03"
         Top             =   420
         Width           =   2460
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
         TabIndex        =   8
         Tag             =   "02"
         Top             =   420
         Width           =   2580
      End
      Begin VB.OptionButton optCollection 
         Caption         =   "Collection"
         Height          =   195
         Left            =   270
         TabIndex        =   7
         Tag             =   "01"
         Top             =   1125
         Value           =   -1  'True
         Width           =   1260
      End
      Begin VB.PictureBox ctlblink1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         ScaleHeight     =   195
         ScaleWidth      =   2235
         TabIndex        =   45
         Top             =   480
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.PictureBox ctlblink2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3120
         ScaleHeight     =   195
         ScaleWidth      =   2115
         TabIndex        =   46
         Top             =   480
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.PictureBox ctlblink3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5760
         ScaleHeight     =   195
         ScaleWidth      =   1755
         TabIndex        =   47
         Top             =   480
         Visible         =   0   'False
         Width           =   1815
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4590
      Left            =   405
      ScaleHeight     =   4560
      ScaleWidth      =   11640
      TabIndex        =   3
      Top             =   11970
      Width           =   11670
      Begin VB.ComboBox cmbEntry 
         Height          =   315
         Left            =   4200
         TabIndex        =   36
         Text            =   "cmbEntry"
         Top             =   2520
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txt_entry 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   525
         Left            =   5640
         TabIndex        =   29
         Top             =   1800
         Visible         =   0   'False
         Width           =   1665
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   4560
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Width           =   11640
         _ExtentX        =   20532
         _ExtentY        =   8043
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
      Begin VB.Frame fmeCA 
         Caption         =   "Cash Advance Details"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4575
         Left            =   0
         TabIndex        =   48
         Top             =   0
         Visible         =   0   'False
         Width           =   11655
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   375
            Left            =   7920
            TabIndex        =   75
            Top             =   840
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   661
            _Version        =   393216
            Format          =   57081857
            CurrentDate     =   40631
         End
         Begin VB.CommandButton Command4 
            Caption         =   "&Hide"
            Height          =   735
            Left            =   10800
            Picture         =   "frmDisApprove.frx":151FC
            Style           =   1  'Graphical
            TabIndex        =   51
            Top             =   240
            Width           =   735
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Edit"
            Height          =   735
            Left            =   9960
            Picture         =   "frmDisApprove.frx":15786
            Style           =   1  'Graphical
            TabIndex        =   67
            Top             =   240
            Width           =   855
         End
         Begin VB.CommandButton Command5 
            Caption         =   "New"
            Height          =   735
            Left            =   9120
            Picture         =   "frmDisApprove.frx":15BC8
            Style           =   1  'Graphical
            TabIndex        =   71
            Top             =   240
            Width           =   855
         End
         Begin VB.TextBox txtCclaimantcode 
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
            Left            =   8400
            Locked          =   -1  'True
            TabIndex        =   73
            Top             =   360
            Visible         =   0   'False
            Width           =   165
         End
         Begin VB.CommandButton Command6 
            Caption         =   "..."
            Height          =   375
            Left            =   8040
            TabIndex        =   72
            Top             =   360
            Width           =   375
         End
         Begin VB.TextBox txtCObrno 
            Alignment       =   2  'Center
            BackColor       =   &H80000000&
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
            Left            =   1080
            Locked          =   -1  'True
            TabIndex        =   69
            Top             =   1320
            Width           =   2535
         End
         Begin VB.TextBox txtCDvno 
            Alignment       =   2  'Center
            BackColor       =   &H80000000&
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
            Left            =   1080
            Locked          =   -1  'True
            TabIndex        =   58
            ToolTipText     =   "Type Dvno and press ENTER"
            Top             =   360
            Width           =   2535
         End
         Begin VB.CommandButton Command3 
            Caption         =   "&Add"
            Height          =   375
            Left            =   3720
            MouseIcon       =   "frmDisApprove.frx":1600A
            Picture         =   "frmDisApprove.frx":19B04
            TabIndex        =   57
            Top             =   360
            Width           =   495
         End
         Begin VB.TextBox txtCCheckno 
            Alignment       =   2  'Center
            BackColor       =   &H80000000&
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
            Left            =   1080
            Locked          =   -1  'True
            TabIndex        =   56
            ToolTipText     =   "Type Checkno and press ENTER"
            Top             =   840
            Width           =   2535
         End
         Begin VB.TextBox txtCChecdate 
            Alignment       =   2  'Center
            BackColor       =   &H80000000&
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "M/d/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   3
            EndProperty
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
            Left            =   5280
            Locked          =   -1  'True
            TabIndex        =   55
            Top             =   840
            Width           =   2655
         End
         Begin VB.TextBox txtCParticular 
            BackColor       =   &H80000000&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   1080
            Locked          =   -1  'True
            MaxLength       =   50
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   54
            Top             =   1800
            Width           =   6855
         End
         Begin VB.TextBox txtCamount 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000000&
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0.00;(#,##0.00)"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   5280
            Locked          =   -1  'True
            TabIndex        =   53
            Top             =   1320
            Width           =   2655
         End
         Begin VB.TextBox txtCClaimant 
            BackColor       =   &H80000000&
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
            Left            =   5280
            Locked          =   -1  'True
            TabIndex        =   52
            Top             =   360
            Width           =   2655
         End
         Begin VB.TextBox txtctotalAmnt 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000000&
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   8040
            Locked          =   -1  'True
            TabIndex        =   49
            Top             =   1920
            Width           =   2415
         End
         Begin MSComctlLib.ListView ListView1 
            Height          =   1935
            Left            =   120
            TabIndex        =   50
            Top             =   2520
            Width           =   11415
            _ExtentX        =   20135
            _ExtentY        =   3413
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
            NumItems        =   9
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "trnno"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   1
               Text            =   "Voucher No."
               Object.Width           =   3528
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   2
               Text            =   "checkno"
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   3
               Text            =   "checkdate"
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "particular"
               Object.Width           =   5292
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "claimant"
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   6
               Text            =   "amount"
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   7
               Text            =   "Obrno"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   8
               Text            =   "claimantcode"
               Object.Width           =   0
            EndProperty
         End
         Begin VB.Label Label23 
            Caption         =   "Obr No.:"
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
            TabIndex        =   70
            Top             =   1320
            Width           =   1215
         End
         Begin VB.Label Label15 
            Caption         =   "Dvno:"
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
            TabIndex        =   65
            Top             =   360
            Width           =   615
         End
         Begin VB.Label Label16 
            Caption         =   "Checkno:"
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
            TabIndex        =   64
            Top             =   840
            Width           =   975
         End
         Begin VB.Label Label17 
            Caption         =   "Checkdate:"
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
            Left            =   4200
            TabIndex        =   63
            Top             =   840
            Width           =   1215
         End
         Begin VB.Label Label18 
            Caption         =   "Particular:"
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
            TabIndex        =   62
            Top             =   1920
            Width           =   1095
         End
         Begin VB.Label Label19 
            Caption         =   "Amount:"
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
            Left            =   4320
            TabIndex        =   61
            Top             =   1320
            Width           =   975
         End
         Begin VB.Label Label20 
            Caption         =   "Claimant:"
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
            Left            =   4320
            TabIndex        =   60
            Top             =   480
            Width           =   975
         End
         Begin VB.Label Label21 
            Caption         =   "Total Amount:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   8400
            TabIndex        =   59
            Top             =   1560
            Width           =   1935
         End
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
      TabIndex        =   34
      Top             =   0
      Visible         =   0   'False
      Width           =   3885
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   855
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   14715
      _ExtentX        =   25956
      _ExtentY        =   1508
      ButtonWidth     =   2011
      ButtonHeight    =   1455
      Appearance      =   1
      Style           =   1
      ImageList       =   "itb32x32"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   8
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "New"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Update"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "Save"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Dis-Approve"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Close"
            ImageIndex      =   7
         EndProperty
      EndProperty
   End
   Begin VB.Label Label22 
      Caption         =   "<<--Cash Advance Details"
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
      Left            =   9240
      TabIndex        =   66
      Top             =   11640
      Visible         =   0   'False
      Width           =   3015
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
      Left            =   14760
      TabIndex        =   35
      Top             =   1290
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "Trn Year :"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   12420
      TabIndex        =   32
      Top             =   11085
      Width           =   705
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "Month of:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   12420
      TabIndex        =   31
      Top             =   11505
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
      TabIndex        =   28
      Top             =   1125
      Width           =   1335
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   885
      Left            =   12240
      Top             =   10950
      Width           =   2235
   End
   Begin VB.Label Label3 
      Caption         =   "Vouchers Prepared with JEV"
      Height          =   225
      Left            =   12270
      TabIndex        =   13
      Top             =   11955
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
      Left            =   435
      TabIndex        =   4
      Top             =   11580
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
      Visible         =   0   'False
      Width           =   3795
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
      TabIndex        =   33
      Top             =   960
      Visible         =   0   'False
      Width           =   2640
   End
End
Attribute VB_Name = "frmJEVDisapprove"
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


Private Sub btnClaimant_Click()
    CUFlag = True
    ActiveFormCaller = "frmJEVPreparation"
    frmCDClaimantRegistry.Show 1
End Sub

Private Sub btnParticular_Click()
    CUFlag = True
    txtparticular.Locked = False
    cmbRC.Locked = False
    rcedit = True
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
    
sql = "SELECT dbo.tblAMIS_IncomingDVTrns.RCenterCode, dbo.tblAMIS_JournalEntry.TransDate, dbo.tblAMIS_JournalEntry.TransType," & _
            "dbo.tblAMIS_JournalEntry.FmisAccntCode, dbo.tblREF_AIS_ChartofAccounts.AccountNameFull, dbo.tblREF_AIS_ChartofAccounts.ChildAccountCode," & _
            "dbo.tblAMIS_JournalEntry.Amount, dbo.tblAMIS_JournalEntry.DebitCredit, dbo.tblAMIS_JournalEntry.Actioncode," & _
            "dbo.tblAMIS_IncomingDVTrns.Particular , dbo.tblAMIS_IncomingDVTrns.ClaimantCode FROM dbo.tblAMIS_JournalEntry INNER JOIN " & _
            "dbo.tblAMIS_IncomingDVTrns ON dbo.tblAMIS_JournalEntry.DVNo = dbo.tblAMIS_IncomingDVTrns.DVNo AND " & _
            "dbo.tblAMIS_JournalEntry.Actioncode = dbo.tblAMIS_IncomingDVTrns.Actioncode INNER JOIN " & _
            "dbo.tblREF_AIS_ChartofAccounts ON dbo.tblAMIS_JournalEntry.FmisAccntCode = dbo.tblREF_AIS_ChartofAccounts.FMISAccountCode AND " & _
            "(dbo.tblAMIS_JournalEntry.ActionCode = dbo.tblREF_AIS_ChartofAccounts.Active or dbo.tblAMIS_JournalEntry.ActionCode=5 )" & _
            "WHERE (dbo.tblREF_AIS_ChartofAccounts.FundType ='" & GetFundName(txtFund.Text) & "') AND (dbo.tblAMIS_JournalEntry.DVNo ='" & List1.Text & "')"
    
    'Debug.Print sql
    ReportName = "JEV"
    rptJEV.txtClaimDesc.SetText txtparticular.Text & ", " & txtClaimant.Text & ", " & txtAlobs.Text
    rptJEV.txtRC.SetText cmbRC.Text
    rptJEV.txtClerk.SetText getUserName(ActiveUserID, "FullName")
    
    If chkSTP.Value = 1 Then
        rptJEV.Line1.Suppress = True
        rptJEV.Line2.Suppress = True
        rptJEV.Line3.Suppress = True
        rptJEV.Line4.Suppress = True
        rptJEV.Line5.Suppress = True
        rptJEV.Line6.Suppress = True
        rptJEV.Line8.Suppress = True
        rptJEV.Line9.Suppress = True
        rptJEV.Line10.Suppress = True
        rptJEV.Line11.Suppress = True
        rptJEV.Line12.Suppress = True
        rptJEV.Line13.Suppress = True
        rptJEV.Line14.Suppress = True
        rptJEV.Line15.Suppress = True
        rptJEV.Line16.Suppress = True
        rptJEV.Line17.Suppress = True
        rptJEV.Line18.Suppress = True
        rptJEV.Line19.Suppress = True
        
        rptJEV.Text1.Suppress = True
        rptJEV.Text2.Suppress = True
        rptJEV.Text3.Suppress = True
        rptJEV.Text4.Suppress = True
        rptJEV.Text8.Suppress = True
        rptJEV.Text9.Suppress = True
        rptJEV.Text12.Suppress = True
        rptJEV.Text13.Suppress = True
        rptJEV.Text15.Suppress = True
        rptJEV.Text16.Suppress = True
        rptJEV.Text17.Suppress = True
        rptJEV.Text18.Suppress = True
        rptJEV.Text19.Suppress = True
        rptJEV.Text20.Suppress = True
        rptJEV.Text21.Suppress = True
        rptJEV.Text22.Suppress = True
        rptJEV.Text25.Suppress = True
        
    End If
    
    rptJEV.Database.SetDataSource opndbaseFMIS.Execute(sql)
    rptJEV.Database.Verify
    If Trim(ActiveUserID) = "0313" Then
    rptJEV.Text23.SetText "Senior bookkeeper"
    End If
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

Private Sub Check2_Click()

End Sub

Private Sub cmb_month_Click()
    Call LoadPrevTrans
End Sub

Private Sub LoadPrevTrans()
Dim PRec As New ADODB.Recordset
Dim x As Integer

    List1.Clear
    List1.Enabled = False
    PRec.Open ("Select DVNo, min(trnno) as trnno From tblAMIS_JournalEntry Where (len(ApprovedByID)=4 or ApprovedByID is not null) and Actioncode=1  and jevseriesno < 1 Group By DVNo order by trnno desc"), opndbaseFMIS, adOpenStatic, adLockOptimistic
    If PRec.RecordCount > 0 Then
        For x = 1 To PRec.RecordCount
            List1.AddItem PRec!dvno
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
           ' SET MSFlexGrid1.SetFocus(MSFlexGrid1.Row, 3)
        ElseIf cmbEntry.Text = "" Then
            MSFlexGrid1.RemoveItem (MSFlexGrid1.Row)

        End If
        cmbEntry.Visible = False
        Call GetSum
    MSFlexGrid1.SetFocus
    End If
    
End Sub





Private Sub cmdOK_Click()

End Sub







Private Sub Command1_Click()
CAedit = True
CAnew = True
Iflck
IfCAedit
End Sub


Private Sub Command3_Click()
Dim x

'    If Trim(txtCDvno.Text) = "" Then
'    if msgbox("Are you sure this transaction  ")
'    End If

    If txtCamount.Text <> "" Then
        If IsNumeric(txtCamount.Text) = False Then
        MsgBox "None Numeric Entry in the Amount,Cannot Proceed the Transaction", vbInformation, "System Message"
        Exit Sub
        End If
        
        If CDbl(txtCamount.Text) <= 0 Then
        MsgBox "Please Specify the amount", vbInformation, "System Message"
        Exit Sub
        End If
        
    End If
    If IfexistDv(txtCDvno.Text) = False Then
        If txtCDvno.Text <> "" And txtCCheckno.Text <> "" And txtCamount.Text <> "" And txtCclaimantcode.Text <> "" Then
            Set x = ListView1.ListItems.Add(, , "")
                x.SubItems(1) = txtCDvno.Text
                x.SubItems(2) = txtCCheckno.Text
                x.SubItems(3) = txtCChecdate.Text
                x.SubItems(4) = txtCParticular.Text
                x.SubItems(5) = txtCClaimant.Text
                x.SubItems(6) = txtCamount.Text
                x.SubItems(7) = txtCObrno.Text
                x.SubItems(8) = txtCclaimantcode.Text
                txtctotalAmnt.Text = Format(GetCATotalamount(ListView1), "#,##0.00")
                CAedit = True
                Lclear
                CAnew = False
                Iflck
        Else
        MsgBox "Please check your entry", vbInformation, "System Message"
        End If
    Else
        MsgBox "Dvno Already on the List", vbInformation, "System Message"
    End If
End Sub
Private Function Lclear()
txtCamount.Text = ""
txtCChecdate.Text = ""
txtCCheckno.Text = ""
txtCClaimant.Text = ""
txtCclaimantcode.Text = ""
txtCDvno.Text = ""
txtCObrno.Text = ""
txtCParticular.Text = ""
End Function
Public Function IfexistDv(ByVal dvno As String) As Boolean
Dim y As Integer
IfexistDv = False
If ListView1.ListItems.Count <> 0 Then
    For y = 1 To ListView1.ListItems.Count
        If dvno = ListView1.ListItems(y).SubItems(1) Then
            IfexistDv = True
        End If
    Next y
End If
If dvno = "None" Then
IfexistDv = True
End If
End Function
Private Sub Command4_Click()
fmeCA.Visible = False
Lclear
End Sub

Private Sub Command5_Click()
CAnew = True
CAClear
Iflck
txtCCheckno.Text = ""
txtCDvno.Text = ""
End Sub
Private Function Iflck()
If CAnew = True Then
    Call unlcked(txtCCheckno)
    Call unlcked(txtCDvno)
    Call unlcked(txtCObrno)
    Call unlcked(txtCParticular)
    Call unlcked(txtCamount)
    'CAClear
Else
    Call lcked(txtCCheckno)
    Call lcked(txtCDvno)
    Call lcked(txtCObrno)
    Call lcked(txtCParticular)
    Call lcked(txtCClaimant)
    Call lcked(txtCChecdate)
    Call lcked(txtCamount)
End If
End Function

Private Sub Command6_Click()
ActiveFormCaller = "CAclaimant"
frmCDClaimantRegistry.Show 1
End Sub

Private Sub Command7_Click()
Iflock = False
frmLock.Show 1
If Iflock = True Then
ifsaveamount = True
txtAmount.Locked = False
End If
End Sub

Private Sub DTPicker1_CallbackKeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer, ByVal CallbackField As String, CallbackDate As Date)
txtCChecdate.Text = Format(DTPicker1.Value, "MM/dd/yyyy")

End Sub

Private Sub DTPicker1_Change()
txtCChecdate.Text = Format(DTPicker1.Value, "MM/dd/yyyy")
End Sub



Private Sub fmeCA_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Label22.FontBold = False
Label22.FontUnderline = False
End Sub
Private Sub Form_Load()
    Edited = False
    
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2
    
    Call Toolbar1_ButtonClick(Toolbar1.Buttons.Item(1))
    
    ActiveUserID = Trim(ActiveUserID)
    LoadOffice
    
End Sub
Public Function LoadCAdetails(ByVal field As String, ByVal Fieldtype As String)
Dim rec As New ADODB.Recordset
Dim rec1 As New ADODB.Recordset
    rec.Open "SELECT top 1    a.DVNo,a.obrno,A.CLAIMANTCODE, b.CheckNo, b.CheckDate, a.Particular, b.ClaimantName,a.GAmount " & _
            "FROM dbo.tblAMIS_IncomingDVTrns AS a inner join tblCMS_CDNewFMISVoucher as c on a.dvno = left(c.newcontrolno,14) inner join  " & _
            "dbo.tblCMS_CDPreparedCheck AS b ON c.fmisvoucherno = b.MixCode   " & _
            "Where (a.ActionCode = 1) And (b.ActionCode = 1) And " & field & " = '" & Replace(Fieldtype, "'", "''") & "'", opndbaseFMIS, adOpenStatic, adLockOptimistic
    If rec.RecordCount <> 0 Then
    
        txtCCheckno.Text = rec.Fields!checkno
        txtCChecdate.Text = rec.Fields!CheckDate
        txtCClaimant.Text = rec.Fields!claimantname
        txtCamount.Text = Format(rec.Fields!Gamount, "#,##0.00")
        txtCParticular.Text = rec.Fields!Particular
        txtCObrno.Text = rec.Fields!obrno
        txtCclaimantcode.Text = rec.Fields!ClaimantCode
        txtCDvno.Text = rec.Fields!dvno
    Else
        rec.Close
        Set rec = Nothing
        rec1.Open "SELECT top 1 percent    a.DVNo,a.obrno,a.claimantcode, b.CheckNo, b.CheckDate, a.Particular, b.ClaimantName,a.GAmount " & _
            "FROM dbo.tblAMIS_IncomingDVTrns AS a inner join " & _
            "dbo.tblCMS_CDPreparedCheck AS b ON a.dvno = b.MixCode   " & _
            "Where (a.ActionCode = 1) And (b.ActionCode = 1) And " & field & " = '" & Replace(Fieldtype, "'", "''") & "'", opndbaseFMIS, adOpenStatic, adLockOptimistic
        If rec1.RecordCount > 0 Then
        txtCDvno.Text = rec1.Fields!dvno
        txtCCheckno.Text = rec1.Fields!checkno
        txtCChecdate.Text = rec1.Fields!CheckDate
        txtCClaimant.Text = rec1.Fields!claimantname
        txtCamount.Text = Format(rec1.Fields!Gamount, "#,##0.00")
        txtCParticular.Text = rec1.Fields!Particular
        txtCObrno.Text = rec1.Fields!obrno
        txtCclaimantcode.Text = rec1.Fields!ClaimantCode
        Else
        MsgBox "Record Not Found", vbInformation, "System Message"
        End If
          rec1.Close
        Set rec1 = Nothing
    End If
End Function
Private Function IfCAedit()
If CAedit = True Then
    txtCDvno.BackColor = &HFFFFFF
    txtCDvno.Locked = False
    'txtctotalAmnt.Text = ""
    txtCCheckno.BackColor = &HFFFFFF
    txtCCheckno.Locked = False
    'CAClear
Else
    txtCCheckno.BackColor = &H80000004
    txtCCheckno.Locked = True
    txtCDvno.BackColor = &H80000004
    txtCDvno.Locked = True
End If
End Function
Public Function Ifliquidition()
If Trim(txtAlobs.Text) = "Liquidation of Cash Advance" Or Trim(txtAlobs.Text) = "Recoupment" Then
    Label22.Visible = True
    ListView1.Visible = True
    Call AllLoadCAdetails(ListView1, txtDVNo.Text, txtctotalAmnt)
Else
    Label22.Visible = False
    fmeCA.Visible = False
    CAClear
    ListView1.ListItems.Clear
    txtCCheckno.Text = ""
    txtctotalAmnt.Text = ""
    txtCDvno.Text = ""
End If
End Function
Private Sub SetGrid()
Dim cc As Integer

    MSFlexGrid1.Clear
    MSFlexGrid1.Rows = 50
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
    MSFlexGrid1.TextMatrix(0, 5) = "ActionCode"
    'If LCase(Trim(lblMode)) = "Edit" Then
        MSFlexGrid1.ColWidth(5) = 1500
    'Else
    '    MSFlexGrid1.ColWidth(5) = 0
    'End If
    
    
    For cc = 0 To MSFlexGrid1.Cols - 1
        MSFlexGrid1.Row = 0
        MSFlexGrid1.col = cc
        MSFlexGrid1.CellAlignment = 4
    Next cc
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Label22.FontBold = False
Label22.FontUnderline = False
End Sub

Private Sub Label22_Click()
fmeCA.Visible = True
End Sub

Private Sub Label22_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Label22.FontBold = True
Label22.FontUnderline = True
End Sub

Private Sub List1_Click()
    Call LoadJEVDetails(List1.Text)
    Ifliquidition
    rcedit = False
End Sub

Private Sub LoadJEVDetails(ByVal dvno As String)
Dim Drec As New ADODB.Recordset
Dim x As Integer
    
    CUFlag = False
    txtparticular.Locked = True
    xNAcode = ""
    Edited = True
   ' lblMode.Caption = "EDIT"
    Drec.Open ("Select  [DVNo],[TransDate],[TransType],Continuing From [tblAMIS_JournalEntry] Where [DVNo]='" & dvno & "' And ActionCode=1"), opndbaseFMIS, adOpenStatic, adLockOptimistic
    If Drec.RecordCount > 0 Then
        
        txtDVNo.Text = Drec![dvno]
        'txtJEVNo.Text = DRec!JEVNo
        txtDate.Text = Drec![TransDate]
        If CInt(optCollection.Tag) = Drec![Transtype] Then optCollection.Value = True
        If CInt(optCheck.Tag) = Drec![Transtype] Then optCheck.Value = True
        If CInt(optCash.Tag) = Drec![Transtype] Then optCash.Value = True
        If CInt(optOther.Tag) = Drec![Transtype] Then optOther.Value = True
        
        If Drec!Continuing = 1 Then
            XFlag = True
        Else
            XFlag = False
        End If
    
    End If
    Drec.Close
    Set Drec = Nothing
        
    Drec.Open ("Select top 1 * FRom tblAMIS_IncomingDVTrns where DVNo='" & txtDVNo.Text & "' and ActionCode=1 order by trnno desc"), opndbaseFMIS, adOpenStatic, adLockOptimistic
    If Drec.RecordCount > 0 Then
        txtClaimant.Text = getClaimant(IIf(IsNull(Drec!ClaimantCode), 0, Drec!ClaimantCode))
        txtClaimantCode = IIf(IsNull(Drec!ClaimantCode), 0, Drec!ClaimantCode)
        cmbRC.Text = GetOfficeName(Drec!RCenter, "OfficeMedium")
        txtparticular.Text = Drec!Particular
        txtFund.Text = Drec!FundType
        txtAmount.Text = Format(Drec![Gamount], "#,##0.00")
        If Drec!NonAlobs = 1 Then
            xObR = GetNonAlobsName(Drec!obrno)
            xNAcode = Drec!obrno
        Else
            xObR = Drec!obrno
        End If
        txtAlobs.Text = xObR
    End If
    Drec.Close
    Set Drec = Nothing
        
    Call SetGrid
    'DRec.Close
    Drec.Open ("Select [FmisAccntCode],[DebitCredit],ActionCode,Amount From [tblAMIS_JournalEntry] Where [DVNo]='" & dvno & "' And (ActionCode=1 or actioncode=5)"), opndbaseFMIS, adOpenStatic, adLockOptimistic
    If Drec.RecordCount > 0 Then
        For x = 1 To Drec.RecordCount
            MSFlexGrid1.TextMatrix(x, 0) = Drec![FmisAccntCode]
            MSFlexGrid1.TextMatrix(x, 1) = GetAccountCodeByFMISAccountCode(Drec![FmisAccntCode])
            MSFlexGrid1.TextMatrix(x, 2) = GetAccountNameByFMISAccountCode(Drec![FmisAccntCode])
            If Drec![debitcredit] = 0 Then
                MSFlexGrid1.TextMatrix(x, 4) = Format(Drec!amount, "#,##0.00")
            Else
                MSFlexGrid1.TextMatrix(x, 3) = Format(Drec!amount, "#,##0.00")
            End If
           ' If LCase(Trim(lblMode.Caption)) = "edit" Then MSFlexGrid1.TextMatrix(x, 5) = DRec!ActionCode  ' for coloraly purpose
            Drec.MoveNext
        Next x
        Call GetSum
    End If
    Drec.Close
    Set Drec = Nothing
End Sub
Private Sub ListView1_Click()
    If ListView1.ListItems.Count <> 0 Then
        txtCDvno.Text = Trim(ListView1.SelectedItem.SubItems(1))
        txtCCheckno.Text = Trim(ListView1.SelectedItem.SubItems(2))
        txtCChecdate.Text = Trim(ListView1.SelectedItem.SubItems(3))
        txtCParticular.Text = Trim(ListView1.SelectedItem.SubItems(4))
        txtCClaimant.Text = Trim(ListView1.SelectedItem.SubItems(5))
        txtCamount.Text = Trim(ListView1.SelectedItem.SubItems(6))
        txtCObrno.Text = Trim(ListView1.SelectedItem.SubItems(7))
        txtCclaimantcode.Text = Trim(ListView1.SelectedItem.SubItems(8))
    End If
End Sub
Private Function CAClear()
txtCamount.Text = ""
txtCChecdate.Text = ""
'txtCCheckno.Text = ""
txtCClaimant.Text = ""
txtCParticular.Text = ""
txtCObrno.Text = ""
End Function

Private Sub ListView1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDelete Then
ListView1.ListItems.Remove (ListView1.SelectedItem.Index)
txtctotalAmnt.Text = Format(GetCATotalamount(ListView1), "#,##0.00")
End If
End Sub

Private Sub MSFlexGrid1_Click()
On Error GoTo bad
    Select Case MSFlexGrid1.col
    Case 1 'AccntCode
        txt_entry.Visible = False
        cmbEntry.Move MSFlexGrid1.CellLeft, MSFlexGrid1.CellTop, MSFlexGrid1.CellWidth
        cmbEntry.Visible = True
        If Len(Trim(MSFlexGrid1.Text)) <> 0 Then
            cmbEntry.Text = MSFlexGrid1.Text
            cmbEntry.SetFocus
        Else
            cmbEntry.ListIndex = -1
            cmbEntry.SetFocus
        End If
    Case 3 To 5 'Debit/Credit
        cmbEntry.Visible = False
        txt_entry.Move MSFlexGrid1.CellLeft, MSFlexGrid1.CellTop, MSFlexGrid1.CellWidth, MSFlexGrid1.CellHeight
        txt_entry.Visible = True
        If Len(Trim(MSFlexGrid1.Text)) <> 0 Then
            txt_entry.Text = MSFlexGrid1.Text
            txt_entry.SelStart = 0
            txt_entry.SelLength = Len(txt_entry.Text)
        Else
            txt_entry.Text = ""
        End If
        txt_entry.SetFocus
    
    Case Else
        txt_entry.Visible = False
        cmbEntry.Visible = False
    End Select
Exit Sub
bad:
MsgBox err.description
End Sub

Private Sub MSFlexGrid1_KeyPress(KeyAscii As Integer)
Call MSFlexGrid1_Click
End Sub

Private Sub MSFlexGrid1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Label22.FontBold = False
Label22.FontUnderline = False
End Sub

Private Sub optCash_Click()
    'txtJEVNo.Text = GetNewJEV(optCash.Tag)
    loadblink
End Sub

Private Sub optCheck_Click()
    'txtJEVNo.Text = GetNewJEV(optCheck.Tag)
    loadblink
End Sub
Public Function loadblink()
If optCheck.Value = True Then
ctlblink1.Visible = True
ctlblink2.Visible = False
ctlblink3.Visible = False
ElseIf optCash.Value = True Then
ctlblink1.Visible = False
ctlblink2.Visible = True
ctlblink3.Visible = False
ElseIf optOther.Value = True Then
ctlblink1.Visible = False
ctlblink2.Visible = False
ctlblink3.Visible = True
End If
End Function

Private Sub optCollection_Click()
    'txtJEVNo.Text = GetNewJEV(optCollection.Tag)
End Sub

Private Sub optOther_Click()
    'txtJEVNo.Text = GetNewJEV(optOther.Tag)
    loadblink
End Sub

Private Function GetNewJEV(ByVal JournalCode As String) As String
Dim Jrec As New ADODB.Recordset
Dim xCode As String

    GetNewJEV = ""
    xCode = GetFundCODE(txtFund.Text) & "-" & Format(Now, "yy-mm") & "-" & JournalCode
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
On Error GoTo bad

    Select Case Button:
    Case "New":
                XFlag = False
                CUFlag = False
                Edited = False
                xNAcode = ""
              '  lblMode.Caption = "NEW"
                txtDVNo.Text = ""
                txtAlobs.Text = ""
                txtClaimant.Text = ""
                txtClaimantCode.Text = ""
                txtparticular.Text = ""
                txtFund.Text = ""
                txtAmount.Text = ""
                txtJEVNo.Text = ""
                txtDate.Text = Format(Now, "MMMM dd, yyyy")
                optCollection.Value = True
                chkSTP.Value = 0
                btnReturn.Enabled = False
                ListView1.ListItems.Clear
                CAClear
                txtctotalAmnt.Text = ""
                txtCDvno.Text = ""
'                Call LoadTrnYear(cmb_trnYear)
'                Call LoadTrnMonth(cmb_month)
            Call SetGrid
    Case "Dis-Approve":
        If txtAlobs.Text <> "" Then
                
                    If MsgBox("Are you sure do you want to Dis-Approve this transaction?", vbQuestion + vbYesNo) = vbYes Then
                        opndbaseFMIS.Execute "Update tblAMIS_JournalEntry set UserID=UserID + '," & ActiveUserID & "',Actioncode=7,DateTimeEntered=DateTimeEntered +'," & Format(Now, "mm/dd/yyyy hh:mm:ss AMPM") & "'  Where DVNo='" & txtDVNo.Text & "' and Actioncode=1"
                        If Check1.Value = 1 Then
                            Call btnReturn_Click
                        End If
                        
                        If chkcancel.Value = 1 Then
                            If ChkIfexistsPTO = True Then
                                MsgBox "The transaction is Active in PTO, if do want to force to cancel this transaction. Please cancel first in PTO", vbInformation, "System Message"
                            Else
                                If MsgBox("Are you sure do you want to Cancel the transaction?", vbInformation + vbYesNo, "System Message") = vbYes Then
                                   opndbaseFMIS.Execute "Update [tblAMIS_IncomingDVTrns] set Actioncode = 3, UserID=rtrim(ltrim(UserID)) + '," & Trim(ActiveUserID) & "',DateTimeEntered=rtrim(ltrim(DateTimeEntered)) +'," & Format(Now, "mm/dd/yyyy hh:mm:ss AMPM") & "' where [DVNo]='" & txtDVNo.Text & "' and actioncode=1"
                                End If
                            End If
                        End If
                        Call Toolbar1_ButtonClick(Toolbar1.Buttons(1))
                    End If
                
        Else
            MsgBox ("Operation Cancel..." & vbNewLine & "Please Check your Entry..!")
        End If
    Case "Close":
                If MsgBox("Are you sure do you want to close this form?", vbQuestion + vbYesNo, "System Security") = vbYes Then
                    Unload Me
                End If
    End Select
Exit Sub
bad:
    Call LoadErr(err.Number, err.Source & ", " & Me.name & ", " & Me.Caption, err.description)
End Sub
Private Function ChkIfexistsPTO() As Boolean
Dim rec As New ADODB.Recordset
ChkIfexistsPTO = False
rec.Open "Select  * from [tblCMS_EXCashVerification] where [VoucherNo] = '" & txtDVNo.Text & "' and actioncode = 1", opndbaseFMIS, adOpenStatic
    If rec.RecordCount > 0 Then
        ChkIfexistsPTO = True
    End If
rec.Close
End Function
Private Function UpdateCA()
Dim z As Integer
    If Trim(txtAlobs.Text) = "Liquidation of Cash Advance" Or Trim(txtAlobs.Text) = "Recoupment" Then
        If ListView1.ListItems.Count <> 0 Then
                opndbaseFMIS.Execute "Update tblAMIS_LiquiditionOfCA set Actioncode=2  Where liquidvno='" & txtDVNo.Text & "' and Actioncode=1"
                For z = 1 To ListView1.ListItems.Count
                    opndbaseFMIS.Execute "Insert into tblAMIS_LiquiditionOfCA ([liquiDvno],[CADvno],[checkno],[checkdate],[status],[actioncode],[amount],caclaimantcode,caobrno,caparticular) " & _
                                        " values ('" & txtDVNo.Text & "' , '" & ListView1.ListItems(z).SubItems(1) & "','" & ListView1.ListItems(z).SubItems(2) & "','" & ListView1.ListItems(z).SubItems(3) & "',0,1, " & CDbl(ListView1.ListItems(z).SubItems(6)) & ",'" & ListView1.ListItems(z).SubItems(8) & "','" & ListView1.ListItems(z).SubItems(7) & "','" & Replace(ListView1.ListItems(z).SubItems(4), "'", "''") & "') "
                Next z
         End If
    End If
End Function
Private Function coloraly() As Boolean

Dim Damount As Double
Dim Camount As Double
coloraly = False
Damount = 0
Camount = 0
ifColoraly = False
Dim x As Integer
    For x = 1 To MSFlexGrid1.Rows - 1
        If MSFlexGrid1.TextMatrix(x, 2) <> "TOTAL" Then
            If MSFlexGrid1.TextMatrix(x, 5) <> "" Then
                If MSFlexGrid1.TextMatrix(x, 5) = "5" Then
                    Damount = val(Damount) + val(CDbl(IIf(MSFlexGrid1.TextMatrix(x, 3) = "", 0, MSFlexGrid1.TextMatrix(x, 3))))
                    Camount = val(Camount) + val(CDbl(IIf(MSFlexGrid1.TextMatrix(x, 4) = "", 0, MSFlexGrid1.TextMatrix(x, 4))))
                    ifColoraly = True
                End If
            End If
            Else
            Exit For
        End If
    Next x
If ifColoraly = True Then
    If (CDbl(Damount) = CDbl(txtAmount.Text) And CDbl(Camount) = CDbl(txtAmount.Text)) Then
        coloraly = True
    End If
Else
coloraly = True
End If
End Function
Private Function ChkEntry() As Boolean
Dim Gamount As Double
    ChkEntry = False
    If Trim(txtDVNo.Text) <> "" And txtAlobs.Text <> "" And txtClaimant.Text <> "" And cmbRC.Text <> "" And txtparticular.Text <> "" And txtFund.Text <> "" And txtAmount.Text <> "" Then
        If xDebit = xCredit And xDebit > 0 Then
'                    If coloraly = True Then
'                        If ifColoraly = True Then
'                            If xDebit = (IIf(txtAmount.Text = "", "0", CDbl(txtAmount.Text)) * 2) Then
'                            ChkEntry = True
'                            End If
'                        Else
'                        GoTo jump
'                        End If
'                    Else
'jump:
                            If (optOther.Value = True Or optCash.Value = True) And xDebit = xCredit And xDebit > 0 And (Trim(txtAlobs.Text) = "Liquidation of Cash Advance" Or Trim(txtAlobs.Text) = "Recoupment") Then
                                 ChkEntry = True
                            ElseIf Format(xDebit, "###,##0.00") = Format(txtAmount.Text, "###,##0.00") Then

                            ChkEntry = True
                            End If
'                    End If
        End If
    End If
End Function
Private Sub LoadExcessDetails(ByVal ObR As String)
On Error GoTo bad
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
Exit Sub
bad:
MsgBox err.description
End Sub


Private Sub LoadObRDetails(ByVal ObR As String)
Dim OREc As New ADODB.Recordset
Dim x As Integer
    
    Call SetGrid
    OREc.Open ("Select * from tblBMS_SubsidiaryLedger where AlobsNo='" & ObR & "' and ActionCode=1"), opndbaseFMIS, adOpenStatic, adLockOptimistic
    If OREc.RecordCount > 0 Then
        For x = 1 To OREc.RecordCount
'            MSFlexGrid1.TextMatrix(x, 0) = OREc!FmisAccountcode
'            MSFlexGrid1.TextMatrix(x, 1) = GetAccountCodeByFMISAccountCode(OREc!FmisAccountcode)
'            MSFlexGrid1.TextMatrix(x, 2) = GetAccountNameByFMISAccountCode(OREc!FmisAccountcode)
'            MSFlexGrid1.TextMatrix(x, 4) = OREc!amount
'            OREc.MoveNext
        Next x
        Call GetSum
    End If
    OREc.Close
    Set OREc = Nothing
    
End Sub

Private Sub txt_entry_KeyPress(KeyAscii As Integer)
 On Error GoTo bad
    If KeyAscii = 13 Then
            If IsNumeric(txt_entry.Text) = False And txt_entry.Text <> "" Then
                MsgBox "None Numeric Entry, Please Check Your Entry", vbCritical, "System Message"
                Exit Sub
            End If
            MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, MSFlexGrid1.col) = Format((txt_entry.Text), "#,##0.00")
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
                If MSFlexGrid1.col = 5 Then
                    If txt_entry.Text = "1" Or txt_entry.Text = "5" Then
                    MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, MSFlexGrid1.col) = txt_entry.Text
                    Else
                    MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, MSFlexGrid1.col) = "1"
                    End If
                End If
        Call GetSum
    End If
Exit Sub
bad:
    Call LoadErr(err.Number, err.Source & ", " & Me.name & ", " & Me.Caption, err.description)
End Sub

Private Sub GetSum()

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
            MSFlexGrid1.TextMatrix(x, 3) = Format(xDebit, "#,##0.00")
            MSFlexGrid1.TextMatrix(x, 4) = Format(xCredit, "#,##0.00")
            Exit For
        End If
    Next x

End Sub

Private Function ChkIfAlreadyJEV(ByVal dvno As String) As String
Dim Jrec As New ADODB.Recordset

    ChkIfAlreadyJEV = ""
    Jrec.Open ("Select * from tblAMIS_JournalEntry where DVNo='" & dvno & "' and Actioncode=1"), opndbaseFMIS, adOpenStatic, adLockOptimistic
    If Jrec.RecordCount > 0 Then
        If Not IsNull(Jrec!ApprovedByID) Then
            ChkIfAlreadyJEV = "Approved" & "-" & Jrec!dvno
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

Private Sub txtCChecdate_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDelete Then
txtCChecdate.Text = ""
End If
End Sub

Private Sub txtCDvno_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If txtCDvno.Text <> "" Then
        If IFLiquidit(txtCDvno.Text, "cadvno") = True Then
        MsgBox "This Dvno Already Liquidit, Cannot Proccess the Trasaction", vbCritical, "System Message"
        CAClear
        ElseIf IFLiquidit(Trim(txtCCheckno.Text), "checkno") = True Then
        MsgBox "This Checkno Already Liquidit, Cannot Proccess the Trasaction", vbCritical, "System Message"
        CAClear
        Else
       Call LoadCAdetails("a.DVNo", txtCDvno.Text)
        End If
    Else
    MsgBox "No Record Foud...", vbInformation, "System Message"
    End If
Else
    If CAnew = True Then
    Else
    CAClear
    End If
End If
End Sub

Private Sub txtCCheckno_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If txtCCheckno.Text <> "" Then
        If IFLiquidit(Trim(txtCCheckno.Text), "checkno") = True Then
            MsgBox "This Checkno Already Liquidit, Cannot Proccess the Trasaction", vbCritical, "System Message"
            CAClear
        Else
        Call LoadCAdetails("b.checkno", txtCCheckno.Text)
        End If
    End If
Else
CAClear
End If
End Sub

Public Function IFLiquidit(ByVal field As String, ByVal Fieldtype As String) As Boolean
Dim rec As New ADODB.Recordset
IFLiquidit = False
rec.Open "Select " & Fieldtype & ",sum(amount) as Sumamount from tblAMIS_LiquiditionOfCA where  " & Fieldtype & " = '" & Replace(field, "'", "''") & "' and actioncode = 1 group by " & Fieldtype & "", opndbaseFMIS, adOpenStatic, adLockOptimistic
    If rec.RecordCount > 0 Then
        IFLiquidit = True
    End If
rec.Close
End Function
Private Sub txtCDvno_LostFocus()
If Trim(txtCDvno.Text) = "" Then
txtCDvno.Text = "N/A"
End If
End Sub
Private Sub txtDVNo_KeyPress(KeyAscii As Integer)
Dim DVRec As New ADODB.Recordset
Dim xAlreadyJEV As String

    If KeyAscii = 13 Then
        btnReturn.Enabled = False
        CUFlag = False
        txtparticular.Locked = True
        
        xNAcode = ""
        txtDVNo.Text = Trim(txtDVNo.Text)
        If ChkDVExist(txtDVNo.Text) = True Then
            xAlreadyJEV = ChkIfAlreadyJEV(txtDVNo.Text)
            If xAlreadyJEV = "" Then
                MsgBox "The DV No. " & txtDVNo.Text & " is not approved, Cannot Disapprove the transaction", vbExclamation + vbOKOnly
            Else
                DVRec.Open ("Select * FRom tblAMIS_IncomingDVTrns where DVNo='" & txtDVNo.Text & "' and (ActionCode=1 or ActionCode=5)"), opndbaseFMIS, adOpenStatic, adLockOptimistic
                If DVRec.RecordCount > 0 Then
                    If DVRec!PAout = 1 Then
                        If DVRec!returnflag = 0 Then
                            btnReturn.Enabled = True
                            If DVRec!NonAlobs = 1 Then
                                xObR = GetNonAlobsName(DVRec!obrno)
                                xNAcode = DVRec!obrno
                            Else
                                xObR = DVRec!obrno
                            End If
                            
                            txtAlobs.Text = xObR
                            txtClaimant.Text = getClaimant(DVRec!ClaimantCode)
                            txtClaimantCode.Text = DVRec!ClaimantCode
                           cmbRC.Text = GetOfficeName(DVRec!RCenter, "OfficeMedium")
                            txtparticular.Text = DVRec!Particular
                            txtFund.Text = DVRec!FundType
                            txtAmount.Text = Format(DVRec!Gamount, "#,##0.00")
                            optCollection.Value = True
                            Ifliquidition
                            Call optCollection_Click
                            
                            
                            XFlag = False
                            If DVRec!Continuing = 1 Then
                                    XFlag = True
                                Call LoadExcessDetails(DVRec!obrno)
                            Else
                                Call LoadObRDetails(DVRec!obrno)
                            End If
                            
                        Else
                            MsgBox "This transaction must pass to pre-audit first!", vbExclamation + vbOKOnly
                            Call Toolbar1_ButtonClick(Toolbar1.Buttons.Item(1))
                        End If
                    Else
                        MsgBox "Please log out DV No. " & txtDVNo.Text & " on pre-audit first!", vbExclamation + vbOKOnly
                        Call Toolbar1_ButtonClick(Toolbar1.Buttons.Item(1))
                    End If
                End If
                DVRec.Close
                Set DVRec = Nothing
                Call LoadJEVDetails(txtDVNo.Text)
'            Else
'                MsgBox "The DV No. " & txtDVNo.Text & " is not approved..", vbExclamation + vbOKOnly
'                Call Toolbar1_ButtonClick(Toolbar1.Buttons.Item(1))
                
            End If
        Else
            MsgBox "Invalid DV Number!", vbExclamation
            
        End If
    Else
                XFlag = False
                CUFlag = False
                Edited = False
                xNAcode = ""
              '  lblMode.Caption = "NEW"
              '  txtDVNo.Text = ""
                txtAlobs.Text = ""
                txtClaimant.Text = ""
                txtClaimantCode.Text = ""
                txtparticular.Text = ""
                txtFund.Text = ""
                txtAmount.Text = ""
                txtJEVNo.Text = ""
                txtDate.Text = Format(Now, "MMMM dd, yyyy")
                optCollection.Value = True
                chkSTP.Value = 0
                btnReturn.Enabled = False
                ListView1.ListItems.Clear
                CAClear
                txtctotalAmnt.Text = ""
                txtCDvno.Text = ""
            Call SetGrid
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



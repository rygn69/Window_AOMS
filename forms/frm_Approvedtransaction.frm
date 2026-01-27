VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.Form frm_Approvedtransaction 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "JEV Preparation"
   ClientHeight    =   9780
   ClientLeft      =   855
   ClientTop       =   795
   ClientWidth     =   14685
   Icon            =   "frm_Approvedtransaction.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "frm_Approvedtransaction.frx":076A
   ScaleHeight     =   9780
   ScaleWidth      =   14685
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
      Left            =   3000
      TabIndex        =   74
      Top             =   4800
      Value           =   1  'Checked
      Width           =   2055
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
            Picture         =   "frm_Approvedtransaction.frx":0CF4
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Approvedtransaction.frx":2686
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Approvedtransaction.frx":4018
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Approvedtransaction.frx":59AA
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Approvedtransaction.frx":733C
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Approvedtransaction.frx":8CCE
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Approvedtransaction.frx":A660
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Approvedtransaction.frx":BFF2
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Approvedtransaction.frx":D984
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Approvedtransaction.frx":F318
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Approvedtransaction.frx":FFF4
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Approvedtransaction.frx":108D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Approvedtransaction.frx":115B0
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Approvedtransaction.frx":1228C
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Approvedtransaction.frx":12F68
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Approvedtransaction.frx":13C44
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Approvedtransaction.frx":14920
            Key             =   ""
         EndProperty
      EndProperty
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
      Left            =   12000
      TabIndex        =   42
      Top             =   840
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
      TabIndex        =   41
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
      TabIndex        =   39
      Top             =   3120
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.CheckBox chkSTP 
      Caption         =   "Shoot-To-Print"
      Height          =   255
      Left            =   12240
      TabIndex        =   37
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
      TabIndex        =   30
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
      Height          =   1845
      Left            =   240
      TabIndex        =   14
      Top             =   1920
      Width           =   14355
      Begin VB.CommandButton Command7 
         Caption         =   "..."
         Height          =   255
         Left            =   13680
         TabIndex        =   72
         ToolTipText     =   "Click here to edit particulars..."
         Top             =   1320
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
         Width           =   255
      End
      Begin VB.CommandButton btnClaimant 
         Caption         =   "..."
         Height          =   255
         Left            =   4680
         TabIndex        =   38
         ToolTipText     =   "Click here to select claimant..."
         Top             =   1320
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
         Top             =   540
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
      Height          =   3420
      Left            =   12240
      TabIndex        =   12
      Top             =   5400
      Width           =   2265
   End
   Begin VB.CommandButton btnPrtJEV 
      Caption         =   "Print JEV"
      Height          =   360
      Left            =   12225
      TabIndex        =   11
      Top             =   9240
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
      Top             =   3825
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
         TabIndex        =   44
         Top             =   960
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
         TabIndex        =   45
         Top             =   1200
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
         TabIndex        =   46
         Top             =   960
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
      Top             =   5130
      Width           =   11670
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
         TabIndex        =   47
         Top             =   0
         Visible         =   0   'False
         Width           =   11655
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   375
            Left            =   7920
            TabIndex        =   73
            Top             =   840
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   661
            _Version        =   393216
            Format          =   51249153
            CurrentDate     =   40631
         End
         Begin VB.CommandButton Command4 
            Caption         =   "&Hide"
            Height          =   735
            Left            =   10800
            Picture         =   "frm_Approvedtransaction.frx":151FC
            Style           =   1  'Graphical
            TabIndex        =   50
            Top             =   240
            Width           =   735
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Edit"
            Height          =   735
            Left            =   9960
            Picture         =   "frm_Approvedtransaction.frx":15786
            Style           =   1  'Graphical
            TabIndex        =   66
            Top             =   240
            Width           =   855
         End
         Begin VB.CommandButton Command5 
            Caption         =   "New"
            Height          =   735
            Left            =   9120
            Picture         =   "frm_Approvedtransaction.frx":15BC8
            Style           =   1  'Graphical
            TabIndex        =   69
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
            TabIndex        =   71
            Top             =   360
            Visible         =   0   'False
            Width           =   165
         End
         Begin VB.CommandButton Command6 
            Caption         =   "..."
            Height          =   375
            Left            =   8040
            TabIndex        =   70
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
            TabIndex        =   67
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
            TabIndex        =   57
            ToolTipText     =   "Type Dvno and press ENTER"
            Top             =   360
            Width           =   2535
         End
         Begin VB.CommandButton Command3 
            Caption         =   "&Add"
            Height          =   375
            Left            =   3720
            MouseIcon       =   "frm_Approvedtransaction.frx":1600A
            Picture         =   "frm_Approvedtransaction.frx":19B04
            TabIndex        =   56
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
            TabIndex        =   55
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
            TabIndex        =   54
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
            TabIndex        =   53
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
            TabIndex        =   52
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
            TabIndex        =   51
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
            TabIndex        =   48
            Top             =   1920
            Width           =   2415
         End
         Begin MSComctlLib.ListView ListView1 
            Height          =   1935
            Left            =   120
            TabIndex        =   49
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
            TabIndex        =   68
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
            TabIndex        =   64
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
            TabIndex        =   63
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
            TabIndex        =   62
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
            TabIndex        =   61
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
            TabIndex        =   60
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
            TabIndex        =   59
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
            TabIndex        =   58
            Top             =   1560
            Width           =   1935
         End
      End
      Begin VB.ComboBox cmbEntry 
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
         TabIndex        =   36
         Text            =   "cmbEntry"
         Top             =   480
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.TextBox txt_entry 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   525
         Left            =   5160
         TabIndex        =   29
         Top             =   720
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
      Left            =   10680
      Locked          =   -1  'True
      TabIndex        =   34
      Top             =   120
      Visible         =   0   'False
      Width           =   3885
   End
   Begin VB.PictureBox freeSizer1 
      Height          =   480
      Left            =   8760
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   75
      Top             =   2160
      Width           =   1200
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
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Label24 
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
      Left            =   8760
      TabIndex        =   77
      Top             =   1270
      Width           =   750
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "New"
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
      Left            =   0
      TabIndex        =   76
      Top             =   0
      Width           =   525
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
      TabIndex        =   65
      Top             =   4800
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Label lblMode 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "New"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   300
      Left            =   9600
      TabIndex        =   35
      Top             =   1290
      Width           =   525
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "Trn Year :"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   12420
      TabIndex        =   32
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
      TabIndex        =   31
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
      TabIndex        =   28
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
      TabIndex        =   13
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
      Left            =   435
      TabIndex        =   4
      Top             =   4740
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
      TabIndex        =   33
      Top             =   960
      Visible         =   0   'False
      Width           =   2640
   End
End
Attribute VB_Name = "frm_Approvedtransaction"
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
Public isfrom_jevNumbering, EditCount, IsSaveAccntng As Boolean
Dim not_coloraly_total_debit, not_coloraly_total_credit, coloraly_total_debit, coloraly_total_credit As Double
Private Sub btnClaimant_Click()
    CUFlag = True
    ActiveFormCaller = "frmJEVPreparation"
    frmCDClaimantRegistry.Show 1
End Sub

Private Sub btnParticular_Click()
    CUFlag = True
    txtParticular.Locked = False
    cmbrc.Locked = False
    rcedit = True
End Sub

Private Sub btnPrtJEV_Click()
Dim sql As String

If Edited = True Then

sql = "Exec Proc_JevPrinting @dvno = '" & txtDVNo.Text & "'"
    
    'Debug.Print sql
    ReportName = "JEVNEW"
    rptJEVNew.txtClaimDesc.SetText txtParticular.Text & ", " & txtClaimant.Text & ", " & txtAlobs.Text
    rptJEVNew.txtRC.SetText cmbrc.Text
    rptJEVNew.txtClerk.SetText getUserName(ActiveUserID, "FullName")
    rptJEVNew.Text23.SetText GetEmpPosition(ActiveUserID)
    If optCash.Value = True Then: rptJEVNew.Trantype = 3
    If optCheck.Value = True Then: rptJEVNew.Trantype = 2
    If optCollection.Value = True Then: rptJEVNew.Trantype = 1
    If optOther.Value = True Then: rptJEVNew.Trantype = 4
    
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
    If Trim(ActiveUserID) = "0313" Then
    rptJEVNew.Text23.SetText "Senior bookkeeper"
    End If
    Call TransactionLogging("Print Preview", "Journal Entry Voucher", Me.Caption, Winsock1.LocalIP)
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
    frmDVSearch.Show
End Sub

Private Sub cmb_month_Click()
    Call LoadPrevTrans
End Sub

Private Sub LoadPrevTrans()
Dim PRec As New ADODB.Recordset
Dim x As Integer

    List1.Clear
    List1.Enabled = False
    Set PRec = opndbaseFMIS.Execute("Select DVNo, min(trnno) as trnno From tblAMIS_JournalEntry Where  ApprovedByID is null and Actioncode=1 and obrno <> 'NA-21' Group By DVNo order by trnno desc")
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

Private Sub cmb_trnYear_Click()
    Call LoadPrevTrans
End Sub


Private Sub cmbEntry_KeyPress(KeyAscii As Integer)
Dim name As String
    If KeyAscii = 13 Then
        If cmbEntry.ListIndex <> -1 Then
            MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 1) = cmbEntry.Text
            MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 0) = cmbEntry.ItemData(cmbEntry.ListIndex)
            MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 5) = "1"
            If MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2) = "TOTAL" Then
                MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 3) = ""
                MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 4) = ""
            End If
            name = GetAccountNameByAccountcode(cmbEntry.Text)
            MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2) = name
        ElseIf cmbEntry.Text = "" Then
            MSFlexGrid1.RemoveItem (MSFlexGrid1.Row)

        End If
        cmbEntry.Visible = False
        Call GetSum
    MSFlexGrid1.SetFocus
    Else
        KeyAscii = AutoFind(cmbEntry, KeyAscii, True)
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
Dim aa
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
    MSFlexGrid1.Rows = 2
    MSFlexGrid1.Cols = 6 ' IIf(LCase(Trim(lblMode)) = "edit", 6, 5)
    
    MSFlexGrid1.TextMatrix(0, 1) = "Account Code"
    MSFlexGrid1.TextMatrix(0, 2) = "Accounts and Explanation"
    MSFlexGrid1.TextMatrix(0, 3) = "Debit"
    MSFlexGrid1.TextMatrix(0, 4) = "Credit"
    
    MSFlexGrid1.ColWidth(0) = 0
    MSFlexGrid1.ColWidth(1) = 1500
    MSFlexGrid1.ColWidth(2) = 6550
    MSFlexGrid1.ColWidth(3) = 1500
    MSFlexGrid1.ColWidth(4) = 1500
    MSFlexGrid1.TextMatrix(0, 5) = "ActionCode"
    'If LCase(Trim(lblMode)) = "Edit" Then
       ' MSFlexGrid1.ColWidth(5) = 1500
    'Else
       MSFlexGrid1.ColWidth(5) = 0
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
    IsSaveAccntng = False
    Call LoadJEVDetails(List1.Text)
    Ifliquidition
    rcedit = False
    
End Sub

Private Sub LoadJEVDetails(ByVal dvno As String)
Dim DRec As New ADODB.Recordset
Dim x As Integer
    
    CUFlag = False
    txtParticular.Locked = True
    xNAcode = ""
    Edited = True
    lblMode.Caption = "EDIT"
    Set DRec = opndbaseFMIS.Execute("Select  [DVNo],[TransDate],[TransType],Continuing From [tblAMIS_JournalEntry] Where [DVNo]='" & dvno & "' And ActionCode=1")
    If DRec.RecordCount > 0 Then
        
        txtDVNo.Text = DRec![dvno]
        'txtJEVNo.Text = DRec!JEVNo
        txtDate.Text = DRec![TransDate]
        If CInt(optCollection.Tag) = DRec![Transtype] Then optCollection.Value = True
        If CInt(optCheck.Tag) = DRec![Transtype] Then optCheck.Value = True
        If CInt(optCash.Tag) = DRec![Transtype] Then optCash.Value = True
        If CInt(optOther.Tag) = DRec![Transtype] Then optOther.Value = True
        
        If DRec!Continuing = 1 Then
            XFlag = True
        Else
            XFlag = False
        End If
    
    End If
    DRec.Close
    Set DRec = Nothing
        
    Set DRec = opndbaseFMIS.Execute("Select top 1 * FRom tblAMIS_IncomingDVTrns where DVNo='" & txtDVNo.Text & "' and ActionCode=1 order by trnno desc")
    If DRec.RecordCount > 0 Then
        txtClaimant.Text = getClaimant(IIf(IsNull(DRec!ClaimantCode), 0, DRec!ClaimantCode))
        txtClaimantCode = IIf(IsNull(DRec!ClaimantCode), 0, DRec!ClaimantCode)
        cmbrc.Text = GetOfficeName(DRec!RCenter, "OfficeMedium")
        txtParticular.Text = DRec!Particular
        txtFund.Text = DRec!FundType
        txtAmount.Text = Format(DRec![Gamount], "#,##0.00")
        EditCount = False
        If DRec!NonAlobs = 1 Then
            xObR = GetNonAlobsName(DRec!obrno)
            xNAcode = DRec!obrno
        Else
            If DRec!moreobr = 1 Then
            xObR = Trim(DRec!obrno) & "," & Trim(DRec!obr2)
            Else
            xObR = DRec!obrno
            End If
        End If
        txtAlobs.Text = xObR
    End If
    DRec.Close
    Set DRec = Nothing
    Call GetAccntngEntries
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
                        opndbaseFMIS.Execute "Delete from tblAMIs_tmpjournal where Dvno = '" & txtDVNo.Text & "'"
                        For x = 1 To DRec.RecordCount
                        DoEvents
                            opndbaseFMIS.Execute "Insert into tblAMIs_tmpjournal (Dvno,Accountcode,Debit,Credit) values ('" & Trim(txtDVNo.Text) & "','" & Trim(DRec!childaccountcode) & "'," & DRec!Debit & "," & DRec!Credit & ")"
                            DRec.MoveNext
                        Next x
                    End If
            Else
            For x = 1 To DRec.RecordCount
                        DoEvents
                            opndbaseFMIS.Execute "Insert into tblAMIs_tmpjournal (Dvno,Accountcode,Debit,Credit) values ('" & Trim(txtDVNo.Text) & "','" & Trim(DRec!childaccountcode) & "'," & DRec!Debit & "," & DRec!Credit & ")"
                            DRec.MoveNext
                        Next x
            End If
            rec.Close
        End If
    End If
    DRec.Close
    Set DRec = Nothing
End Function
Public Function SaveAcctngEntries(ByVal dvno As String)
Dim DRec As New ADODB.Recordset
Dim rec As New ADODB.Recordset
Dim x As Integer
Dim xType As Integer
If optCollection.Value = True Then xType = CInt(optCollection.Tag)
If optCash.Value = True Then xType = CInt(optCash.Tag)
If optCheck.Value = True Then xType = CInt(optCheck.Tag)
If optOther.Value = True Then xType = CInt(optOther.Tag)
    DRec.Open ("Select Accountcode,sum(Debit) as Debit ,sum(Credit) AS Credit From tblAMIs_tmpjournal Where [dvno]='" & dvno & "'group by accountcode"), opndbaseFMIS, adOpenStatic, adLockOptimistic
    If DRec.RecordCount > 0 Then
        opndbaseFMIS.Execute "update tblAMIS_AccoutingEntries set actioncode =2 where reffno = '" & txtDVNo.Text & "' and actioncode =1" ', datetimeentered = rtrim(ltrim(DateTimeEntered)) +'," & Format(Now, "yyyy/mm/dd hh:mm:ss AMPM") & "',userid = UserID + '," & Trim(ActiveUserID) & "'
        For x = 1 To DRec.RecordCount
            DoEvents
            opndbaseFMIS.Execute "Insert into tblAMIS_AccoutingEntries (reffNo,ChildAccountcode,debit,credit,actioncode,datetimeentered,transtype,userid) values " & _
            "('" & Trim(txtDVNo.Text) & "','" & Trim(DRec!accountcode) & "'," & DRec!Debit & "," & DRec!Credit & ",1,'" & Format(Now, "yyyy/mm/dd hh:mm:ss AMPM") & "'," & xType & ",'" & Trim(ActiveUserID) & "')"
            DRec.MoveNext
        Next x
        opndbaseFMIS.Execute "delete from tblAMIs_tmpjournal where dvno = '" & txtDVNo.Text & "'"
    End If
    DRec.Close
    Set DRec = Nothing
End Function
Public Sub GetAccntngEntries()
Dim DRec As New ADODB.Recordset
Dim x As Integer
Call SetGrid
    'DRec.Close
    If IsSaveAccntng = False Then
        Set DRec = opndbaseFMIS.Execute("Select left(ChildAccountcode,3) as childcode,sum(Debit) as sumdebit,sum(credit) as sumcredit From tblAMIS_AccoutingEntries Where [reffno]='" & txtDVNo.Text & "' And (ActionCode=1) group by reffno,actioncode,left(ChildAccountcode,3)")
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
            
        End If
    Else
    If opndbaseFMIS.State = 1 Then
    
    End If
        Set DRec = opndbaseFMIS.Execute("Select left(Accountcode,3) as childcode,sum(Debit) as sumdebit,sum(credit) as sumcredit From tblAMIS_tmpjournal Where [dvno]='" & txtDVNo.Text & "' group by Dvno,left(Accountcode,3)")
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

Private Sub List1_KeyUp(KeyCode As Integer, Shift As Integer)
DoEvents
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
If chkSC.Value = 1 Then
Call MSFlexGrid1_DblClick
End If
End Sub
Private Sub MSFlexGrid1_DblClick()
If Trim(txtAlobs.Text) <> "" Then
    With frmSub3
        .reff = txtDVNo.Text
        .Gamount = txtAmount.Text
        .CName = UCase(txtClaimant.Text)
        .isEdit = True
       Set .frm = Me
        Call LoadAcctngEntries(txtDVNo.Text)
        .Show 1
        Call GetAccntngEntries
    End With
End If
End Sub
Private Sub MSFlexGrid1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Label22.FontBold = False
Label22.FontUnderline = False
End Sub
Private Sub optCash_Click()
    loadblink
End Sub

Private Sub optCheck_Click()
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
'On Error GoTo bad

    Select Case Button:
    Case "New":
                XFlag = False
                CUFlag = False
                Edited = False
                xNAcode = ""
                lblMode.Caption = "NEW"
                'txtDVNo.Text = ""
                txtAlobs.Text = ""
                txtClaimant.Text = ""
                txtClaimantCode.Text = ""
                txtParticular.Text = ""
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
                Call LoadTrnYear(cmb_trnYear)
                Call LoadTrnMonth(cmb_month)
                Call SetGrid
                
                
    Case "Save":
            
            If optCollection.Value = True Then
            MsgBox "Select Transaction Type", vbInformation, "System Message"
            Else
                        If ChkEntry = True Then
                           If Trim(txtAlobs.Text) = "Liquidation of Cash Advance" Or Trim(txtAlobs.Text) = "Recoupment" Then
                                If ListView1.ListItems.Count <> 0 Then
                                    If txtAmount.Text <> txtctotalAmnt.Text Then
                                        MsgBox "Gross Amount not Equal to your Total Cash Advance amount..!" & vbNewLine & "Please Check Yout Entry in Cash Advance Details..", vbCritical, "System Message"
                                        Exit Sub 'stop transaction
                                    End If
                                 Else
                                        MsgBox "Cash Advance Details Empty..!" & vbNewLine & "Please Entry the Cash Advance Details, to Proceed the transaction..", vbCritical, "System Message"
                                        Exit Sub ' stop transaction
                                 End If
                            End If
                        
                            If MsgBox("Are you sure you want to save this transaction?", vbQuestion + vbYesNo) = vbYes Then
                                
                                If optCollection.Value = True Then xType = CInt(optCollection.Tag)
                                If optCash.Value = True Then xType = CInt(optCash.Tag)
                                If optCheck.Value = True Then xType = CInt(optCheck.Tag)
                                If optOther.Value = True Then xType = CInt(optOther.Tag)
                                If xNAcode <> "" Then
                                    xObR = xNAcode
                                End If
                              'MsgBox xObR
                                    opndbaseFMIS.Execute "Update tblAMIS_JournalEntry set ActionCode=2, UserID=UserID + '," & ActiveUserID & "', DateTimeEntered=DateTimeEntered + '," & Format(Now, "yyyy/mm/dd hh:mm:ss AMPM") & "' Where DVNo='" & txtDVNo.Text & "' And ActionCode=1"
                                    opndbaseFMIS.Execute "Insert Into tblAMIS_JournalEntry (TransType,DVNo,ObrNo,TransDate,UserID,Actioncode,DateTimeEntered,Continuing,debitcredit,isnew,FmisAccntCode) values (" & xType & ",'" & Trim(Replace(txtDVNo.Text, "'", "''")) & "','" & Left(xObR, 19) & "','" & Format(Now, "yyyy/mm/dd hh:mm:ss AMPM") & "','" & ActiveUserID & "',1,'" & Format(Now, "yyyy/mm/dd hh:mm:ss AMPM") & "'," & IIf(XFlag, 1, 0) & ",0,1,0)"
                                    If IsSaveAccntng = True Then
                                        Call SaveAcctngEntries(txtDVNo.Text)
                                    End If
                                    EditCount = False
                                If CUFlag = True Then
                                    opndbaseFMIS.Execute "Update [tblAMIS_IncomingDVTrns] set [Particular]='" & Trim(Replace(txtParticular.Text, "'", "''")) & "', [ClaimantCode]='" & txtClaimantCode.Text & "',RCENTER = " & cmbrc.ItemData(cmbrc.ListIndex) & " Where DVNo='" & Trim(txtDVNo.Text) & "' And ActionCode=1"
                                    txtParticular.Locked = True
                                    cmbrc.Locked = True
                                    rcedit = False
                                End If
                                
                               
                                If ifsaveamount = True Then
                                    opndbaseFMIS.Execute "Update [tblAMIS_IncomingDVTrns] set gamount = '" & txtAmount.Text & "' Where DVNo='" & Trim(txtDVNo.Text) & "' And ActionCode=1"
                                End If
                                If CAedit = True Then
                                    UpdateCA
                                End If
                                Call Toolbar1_ButtonClick(Toolbar1.Buttons.Item(1))
                            End If
                        Else
                        
                        End If
                End If
    Case "Delete":
                If Edited = True Then
                    If InStr(ChkIfAlreadyJEV(txtDVNo.Text), "Approved") <> 1 Then
                        If MsgBox("Are you sure you want to delete this transaction?", vbQuestion + vbYesNo) = vbYes Then
                            opndbaseFMIS.Execute "Update tblAMIS_JournalEntry set UserID=UserID + '," & ActiveUserID & "',Actioncode=3,DateTimeEntered=DateTimeEntered +'," & Format(Now, "mm/dd/yyyy hh:mm:ss AMPM") & "'  Where DVNo='" & txtDVNo.Text & "' and Actioncode=1"
                            opndbaseFMIS.Execute "Update tblAMIS_AccoutingEntries set actioncode = 3,datetimeentered = DateTimeEntered +'," & Format(Now, "yyyy/mm/dd hh:mm:ss AMPM") & "',userid = '" & Trim(ActiveUserID) & "' where reffno = '" & txtDVNo.Text & "' and actioncode = 1"
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
Exit Sub
bad:
    Call LoadErr(err.Number, err.Source & ", " & Me.name & ", " & Me.Caption, err.description)
End Sub
Private Function ChkEntry() As Boolean
Dim Gamount As Double
    ChkEntry = False
    If Trim(txtDVNo.Text) <> "" And txtAlobs.Text <> "" And txtClaimant.Text <> "" And cmbrc.Text <> "" And txtParticular.Text <> "" And txtFund.Text <> "" And txtAmount.Text <> "" Then
        If xDebit = xCredit And xDebit > 0 Then
                If (optOther.Value = True Or optCash.Value = True) And (Trim(txtAlobs.Text) = "Liquidation of Cash Advance" Or Trim(txtAlobs.Text) = "Recoupment") Then
                    ChkEntry = True
                Else
                    If CCur(xDebit) > CCur(txtAmount.Text) Then
                        If MsgBox("Your total Credit and Debit Amount is Greater than to your Gross Amount" & vbNewLine & "Are you Sure this transaction is have a Corolary Entry?", vbCritical + vbYesNo, "System Information") = vbYes Then
                            ChkEntry = True
                        End If
                    ElseIf CCur(xDebit) = CCur(txtAmount.Text) Then
                    ChkEntry = True
                    Else
                        If MsgBox("Your total Credit and Debit Amount is Less than to your Gross Amount" & vbNewLine & "Are you Sure this transaction is have a Corolary Entry?", vbCritical + vbYesNo, "System Information") = vbYes Then
                        ChkEntry = True
                        End If
                    End If
                End If
        Else
        MsgBox "Your Total Debit and Total Credit Amount is not Equal or 0 value, Please Check Your Entry", vbInformation, "System Message"
        End If
    Else
    MsgBox "Some Fields are Empty, Please check it..!", vbInformation, "System Message"
    End If
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
                    Damount = Val(Damount) + Val(CDbl(IIf(MSFlexGrid1.TextMatrix(x, 3) = "", 0, MSFlexGrid1.TextMatrix(x, 3))))
                    Camount = Val(Camount) + Val(CDbl(IIf(MSFlexGrid1.TextMatrix(x, 4) = "", 0, MSFlexGrid1.TextMatrix(x, 4))))
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
'Dim OREc As New ADODB.Recordset
'Dim x As Integer
'
'    Call SetGrid
'    OREc.Open ("Select * from tblBMS_SubsidiaryLedger where AlobsNo='" & ObR & "' and ActionCode=1"), opndbaseFMIS, adOpenStatic, adLockOptimistic
'    If OREc.RecordCount > 0 Then
'        For x = 1 To OREc.RecordCount
'            MSFlexGrid1.TextMatrix(x, 0) = OREc!FmisAccountcode
'            MSFlexGrid1.TextMatrix(x, 1) = GetAccountCodeByFMISAccountCode(OREc!FmisAccountcode)
'            MSFlexGrid1.TextMatrix(x, 2) = GetAccountNameByFMISAccountCode(OREc!FmisAccountcode)
'            MSFlexGrid1.TextMatrix(x, 4) = OREc!Amount
'            OREc.MoveNext
'        Next x
'        Call GetSum
'    End If
'    OREc.Close
'    Set OREc = Nothing
'
End Sub

Public Sub LoadAccountsByFund(ByVal fundmedium As String)
Dim ARec As New ADODB.Recordset
Dim x As Integer
Dim FundName As String

    cmbEntry.Clear
    cmbEntry.Visible = False
    FundName = GetFundName(fundmedium)
    ARec.Open ("Select  [AccountCode],[FMISAccountCode] from [tblREF_AIS_ChartOfAccountsMother] Where [Actioncode]=0 order by [AccountCode]"), opndbaseFMIS, adOpenStatic, adLockOptimistic
    If ARec.RecordCount > 0 Then
        For x = 1 To ARec.RecordCount
            cmbEntry.AddItem Trim(ARec![accountcode])
            cmbEntry.ItemData(cmbEntry.NewIndex) = ARec![FmisAccountcode]
            ARec.MoveNext
        Next x
    End If
    ARec.Close
    Set ARec = Nothing
    
End Sub
Public Sub LoadAccountsByFund_OLD(ByVal fundmedium As String)
Dim ARec As New ADODB.Recordset
Dim x As Integer
Dim FundName As String

    cmbEntry.Clear
    cmbEntry.Visible = False
    FundName = GetFundName(fundmedium)
    ARec.Open ("Select  [ChildAccountCode],[FMISAccountCode] from [tblREF_AIS_ChartofAccounts] Where [Active]=1 and [FundType]='" & FundName & "' Order by [ChildAccountCode]"), opndbaseFMIS, adOpenStatic, adLockOptimistic
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
        If MSFlexGrid1.TextMatrix(x, 1) <> "" Then
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

Private Sub TxtDvno_Change()
'If Edited = False Then
'                XFlag = False
'                CUFlag = False
'                Edited = False
'                xNAcode = ""
'                lblMode.Caption = "NEW"
'                txtDVNo.Text = ""
'                txtAlobs.Text = ""
'                txtClaimant.Text = ""
'                txtClaimantCode.Text = ""
'                txtParticular.Text = ""
'                txtFund.Text = ""
'                txtAmount.Text = ""
'                txtJEVNo.Text = ""
'                txtDate.Text = Format(Now, "MMMM dd, yyyy")
'                optCollection.Value = True
'                chkSTP.Value = 0
'                btnReturn.Enabled = False
'                ListView1.ListItems.Clear
'                CAClear
'                txtctotalAmnt.Text = ""
'                txtCDvno.Text = ""
'                Call SetGrid
'End If
End Sub

Private Sub txtDVNo_KeyPress(KeyAscii As Integer)
Dim DVRec As New ADODB.Recordset
Dim xAlreadyJEV As String

    If KeyAscii = 13 Then
        btnReturn.Enabled = False
        CUFlag = False
        txtParticular.Locked = True
        
        xNAcode = ""
        txtDVNo.Text = Trim(txtDVNo.Text)
        If ChkDVExist(txtDVNo.Text) = True Then
            xAlreadyJEV = ChkIfAlreadyJEV(txtDVNo.Text)
            If xAlreadyJEV = "" Then
                MsgBox "The DV No. " & txtDVNo.Text & " is not approved, Cannot Disapprove the transaction", vbExclamation + vbOKOnly, "System Information"
            ElseIf InStr(1, xAlreadyJEV, "Approved") > 0 Then
                Call LoadJEVDetails(txtDVNo.Text)
            Else
                Call LoadJEVDetails(txtDVNo.Text)
            End If
        Else
            MsgBox "Invalid DV Number!", vbExclamation
            Call Toolbar1_ButtonClick(Toolbar1.Buttons.Item(1))
        End If
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



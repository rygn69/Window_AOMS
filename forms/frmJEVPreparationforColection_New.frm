VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmJEVPreparationforColection_New 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "JEV Preparation For Collection  and Deposit and General Journal through PTV Number"
   ClientHeight    =   9405
   ClientLeft      =   -165
   ClientTop       =   2850
   ClientWidth     =   14700
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   -1  'True
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmJEVPreparationforColection_New.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9405
   ScaleWidth      =   14700
   Visible         =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "JEV Transaction Type"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   840
      Left            =   360
      TabIndex        =   6
      Top             =   4080
      Width           =   11295
      Begin VB.OptionButton optOther 
         Caption         =   "General Journal"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   2240
         Style           =   1  'Graphical
         TabIndex        =   8
         Tag             =   "04"
         Top             =   285
         Width           =   2100
      End
      Begin VB.OptionButton optCollection 
         Caption         =   "Cash Receipts"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   150
         Style           =   1  'Graphical
         TabIndex        =   7
         Tag             =   "01"
         Top             =   285
         Width           =   2100
      End
   End
   Begin VB.CheckBox chkSC 
      Caption         =   "Single Click"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2640
      TabIndex        =   49
      Top             =   5040
      Value           =   1  'Checked
      Width           =   2055
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
      Left            =   5040
      TabIndex        =   44
      Top             =   4920
      Visible         =   0   'False
      Width           =   6495
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
      TabIndex        =   42
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
      Left            =   360
      Locked          =   -1  'True
      TabIndex        =   39
      Top             =   3480
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.CheckBox chkSTP 
      Caption         =   "Shoot-To-Print"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   12000
      TabIndex        =   37
      Top             =   8400
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
      ItemData        =   "frmJEVPreparationforColection_New.frx":076A
      Left            =   12990
      List            =   "frmJEVPreparationforColection_New.frx":076C
      TabIndex        =   29
      Top             =   4590
      Width           =   1470
   End
   Begin VB.TextBox txtDate 
      BeginProperty Font 
         Name            =   "Tahoma"
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
      TabIndex        =   26
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
      TabIndex        =   25
      Top             =   4170
      Width           =   1470
   End
   Begin VB.Frame Frame2 
      Caption         =   "Transaction Details"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1965
      Left            =   240
      TabIndex        =   12
      Top             =   2040
      Width           =   14115
      Begin VB.CommandButton Command1 
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
         Left            =   4560
         TabIndex        =   48
         ToolTipText     =   "Click here to Generate JEV number"
         Top             =   360
         Width           =   375
      End
      Begin VB.TextBox txt_JEVNO 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1800
         TabIndex        =   46
         Top             =   360
         Width           =   2700
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
         Top             =   2040
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
         Left            =   9240
         TabIndex        =   40
         ToolTipText     =   "Click here to edit particulars..."
         Top             =   840
         Width           =   375
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
      Begin VB.TextBox txtFund 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   11400
         Locked          =   -1  'True
         TabIndex        =   23
         Top             =   360
         Width           =   2580
      End
      Begin VB.TextBox txtAmount 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   11400
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   840
         Width           =   2580
      End
      Begin VB.TextBox txtParticular 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   960
         Left            =   1800
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   19
         Top             =   840
         Width           =   7410
      End
      Begin VB.TextBox txtAlobs 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   6120
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   360
         Width           =   3060
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
         TabIndex        =   14
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
         Left            =   11880
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   2460
         Visible         =   0   'False
         Width           =   4050
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "JEV Number:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   480
         TabIndex        =   47
         Top             =   360
         Width           =   1110
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fund Type:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   10290
         TabIndex        =   24
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Amount (Gross):"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   9810
         TabIndex        =   22
         Top             =   840
         Width           =   1425
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Particular:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   720
         TabIndex        =   20
         Top             =   840
         Width           =   885
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Report No:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   5040
         TabIndex        =   18
         Top             =   360
         Width           =   930
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
         TabIndex        =   16
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
         Left            =   11820
         TabIndex        =   15
         Top             =   2160
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
      Height          =   2940
      Left            =   12000
      TabIndex        =   10
      Top             =   5385
      Width           =   2505
   End
   Begin VB.CommandButton btnPrtJEV 
      Caption         =   "Print JEV"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   11985
      TabIndex        =   9
      Top             =   8760
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
      Height          =   3855
      Left            =   405
      ScaleHeight     =   3825
      ScaleWidth      =   11160
      TabIndex        =   3
      Top             =   5400
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
         Left            =   8160
         TabIndex        =   28
         Top             =   840
         Visible         =   0   'False
         Width           =   1665
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   3840
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Width           =   11160
         _ExtentX        =   19685
         _ExtentY        =   6773
         _Version        =   393216
         FixedCols       =   0
         ScrollTrack     =   -1  'True
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
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
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   405
      TabIndex        =   1
      Top             =   1335
      Width           =   4845
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   840
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   14700
      _ExtentX        =   25929
      _ExtentY        =   1482
      ButtonWidth     =   1032
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
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
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
               Picture         =   "frmJEVPreparationforColection_New.frx":076E
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmJEVPreparationforColection_New.frx":2100
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmJEVPreparationforColection_New.frx":3A92
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmJEVPreparationforColection_New.frx":5424
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmJEVPreparationforColection_New.frx":6DB6
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmJEVPreparationforColection_New.frx":8748
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmJEVPreparationforColection_New.frx":A0DA
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmJEVPreparationforColection_New.frx":BA6C
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmJEVPreparationforColection_New.frx":D3FE
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmJEVPreparationforColection_New.frx":ED92
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmJEVPreparationforColection_New.frx":FA6E
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmJEVPreparationforColection_New.frx":1034E
               Key             =   ""
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmJEVPreparationforColection_New.frx":1102A
               Key             =   ""
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmJEVPreparationforColection_New.frx":11D06
               Key             =   ""
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmJEVPreparationforColection_New.frx":129E2
               Key             =   ""
            EndProperty
            BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmJEVPreparationforColection_New.frx":136BE
               Key             =   ""
            EndProperty
            BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmJEVPreparationforColection_New.frx":1439A
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
      TabIndex        =   33
      Top             =   240
      Visible         =   0   'False
      Width           =   3885
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
      Left            =   4680
      TabIndex        =   45
      Top             =   4920
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
      TabIndex        =   35
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
      TabIndex        =   31
      Top             =   4245
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
      TabIndex        =   30
      Top             =   4665
      Width           =   915
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date Prepared"
      BeginProperty Font 
         Name            =   "Tahoma"
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
      TabIndex        =   27
      Top             =   1005
      Width           =   1230
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   885
      Left            =   12045
      Top             =   4110
      Width           =   2475
   End
   Begin VB.Label Label3 
      Caption         =   "Vouchers Prepared with JEV"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   12030
      TabIndex        =   11
      Top             =   5115
      Width           =   2430
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Accounting Entries"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   360
      TabIndex        =   4
      Top             =   5040
      Width           =   1575
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Enter PTV  Number:"
      BeginProperty Font 
         Name            =   "Tahoma"
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
      Width           =   1725
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   1095
      Left            =   8640
      Top             =   840
      Width           =   2235
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   1080
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
      TabIndex        =   32
      Top             =   1080
      Visible         =   0   'False
      Width           =   2640
   End
End
Attribute VB_Name = "frmJEVPreparationforColection_New"
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
Dim Ttype As Integer
Public isfrom_jevNumbering, EditCount, IsSaveAccntng As Boolean
Dim not_coloraly_total_debit, not_coloraly_total_credit, coloraly_total_debit, coloraly_total_credit As Double
Dim fundcodes As Integer


Private Sub btnClaimant_Click()
    CUFlag = True
    ActiveFormCaller = "frmJEVPreparation"
    frmCDClaimantRegistry.Show 1
End Sub

Private Sub btnParticular_Click()
    CUFlag = True
    If txtParticular.Locked = False Then
        txtParticular.Locked = True
    Else
        txtParticular.Locked = False
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
    sql = "Exec Proc_JevPrinting @dvno = '" & Trim(txtDVNo.Text) & "'"
    ReportName = "JEVNEW"
    rptJEVNew.txtClaimDesc.SetText txtParticular.Text & ", " & txtClaimant.Text & ", " & txtAlobs.Text
    rptJEVNew.txtRC.SetText txtRC.Text
    rptJEVNew.txtClerk.SetText getUserName(ActiveUserID, "FullName")
    rptJEVNew.Text23.SetText GetEmpPosition(ActiveUserID)
    rptJEVNew.txtJEVNo.SetText txt_JEVNO.Text
    rptJEVNew.txtDate.SetText Format(txtDate.Text, "MM/dd/yyyy")
    rptJEVNew.Trantype = 1
    
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
'Dim PRec As New ADODB.Recordset
'Dim x As Integer
'
'    List1.Clear
'    List1.Enabled = False
'    PRec.Open ("Select ptvno, trnno as trno  From tblAMIS_FinalJEV Where Actioncode=1 and transtype = 1 and isnew =1 and date_ = '" & Format(Now, "yyyy-MM-dd") & "' order by trno desc"), opndbaseFMIS, adOpenStatic, adLockOptimistic
'    If PRec.RecordCount > 0 Then
'        For x = 1 To PRec.RecordCount
'            List1.AddItem PRec!ptvNo
'            PRec.MoveNext
'        Next x
'        List1.Enabled = True
'    End If
'    PRec.Close
'    Set PRec = Nothing
'
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
        MSFlexGrid1.SetFocus
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

Private Sub Command1_Click()
Dim rec As New ADODB.Recordset
JevOk = False
frmPOstdate.Show 1
If JevOk = True Then
rec.Open ("EXEC [dbo].[Proc_GetMaxJevSeries_New] @transtype = " & getTranstype & ",@jevyeardate = '" & DatePost & "' ,@fundtype = '" & txtFund.Text & "'"), opndbaseFMIS, adOpenStatic, adLockOptimistic
txt_JEVNO.Text = Left(txtDVNo.Text, 3) & "-" & Right(Year(DatePost), 2) & "-" & Format(Month(DatePost), "00") & "-" & Format(getTranstype, "00") & "-" & Format(Trim(rec.Fields!MAXJEVSERIES), "0000")
rec.Close
Else
MsgBox "Cannot Generate the System JEV Number,If you cancel to Set the Date", vbInformation, "System Message"
txt_JEVNO.Text = ""
End If

End Sub
Private Sub getTranstype1()
If optOther.Value = True Then
    Ttype = 4
Else
    Ttype = 1
End If
End Sub
Private Sub Form_Load()
    
    Edited = False
    
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2
    
    Call Toolbar1_ButtonClick(Toolbar1.Buttons.Item(1))
    ActiveUserID = Trim(ActiveUserID)
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
    cmbrc.Visible = False
    IsSaveAccntng = False
    Call LoadJEVDetails(List1.Text)
    cmbrc.Visible = False
    IsSaveAccntng = False
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
    DRec.Open ("Select ptvno,Date_,Rdono,Particular,fundtype,gamount,jevno,Transtype From [tblAMIS_FinalJEV] Where [ptvno]='" & dvno & "' And ActionCode=1"), opndbaseFMIS, adOpenStatic, adLockOptimistic
    If DRec.RecordCount > 0 Then
        If DRec!Transtype = 4 Then
        optOther.Value = True
        Else
        optCollection.Value = True
        End If
        txtDVNo.Text = Trim(DRec![ptvNo])
        txtDate.Text = Trim(DRec![Date_])
        DatePost = DRec![Date_]
        txtAlobs.Text = Trim(DRec!RDOno)
        txtParticular.Text = Trim(DRec!Particular)
        txtFund.Text = Trim(DRec!FundType)
        txtAmount.Text = Format(DRec!Gamount, "#,##0.00")
        txt_JEVNO.Text = Trim(DRec!jevno)
        EditCount = False
       
    End If
    DRec.Close
    Set DRec = Nothing
Call GetAccntngEntries
End Sub
Private Sub MSFlexGrid1_Click()
If chkSC.Value = 1 Then
Call MSFlexGrid1_DblClick
End If
End Sub
Private Sub MSFlexGrid1_DblClick()
If Trim(txtAlobs.Text) <> "" Then
    With frmSub3
        .REFF = txtDVNo.Text
        .Gamount = txtAmount.Text
        .CName = "N/A"
        .isEdit = True
        Set .frm = Me
        Call LoadAcctngEntries(txtDVNo.Text)
        .Show 1
        Call GetAccntngEntries
    End With
End If
End Sub

Private Sub MSFlexGrid1_KeyPress(KeyAscii As Integer)

Call MSFlexGrid1_Click
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
        Else
        IsSaveAccntng = True
        Call GetAccntngEntries
        End If
    Else
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

Private Sub optCollection_Click()
txt_JEVNO.Text = ""
End Sub

Private Sub optOther_Click()
    txt_JEVNO.Text = ""
End Sub
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
                txtDVNo.Text = ""
                txtAlobs.Text = ""
                txtClaimant.Text = ""
                txtClaimantCode.Text = ""
                txtRC.Text = ""
                txtParticular.Text = ""
                txtFund.Text = ""
                txtAmount.Text = ""
                txtJEVNo.Text = ""
                txt_JEVNO.Text = ""
                txtDate.Text = Format(Now, "MMMM dd, yyyy")
                optCollection.Value = False
                optOther.Value = False
                chkSTP.Value = 0
                btnReturn.Enabled = False
               
                Call LoadTrnYear(cmb_trnYear)
                Call LoadTrnMonth(cmb_month)
                 Call GetPrevTrans(cmb_trnYear.Text, cmb_month.ItemData(cmb_month.ListIndex))
                Call SetGrid
                
    Case "Save":
                If Edited = True Then
                    MsgBox "Sorry the Transaction is not allow to Edit." & vbNewLine & "Cancel first posted transaction to Edit the transaction.", vbInformation, "System Message"
                    Exit Sub
                End If
                
                If ChkEntry = True Then
                    If MsgBox("Are you sure you want to save this transaction?", vbQuestion + vbYesNo) = vbYes Then

'                        If lblMode.Caption = "EDIT" Then
'                           opndbaseFMIS.Execute "Update [tblAMIS_FinalJEV] set datetimeentered= rtrim(ltrim(datetimeentered)) + '" & Now & "',JEVby = rtrim(ltrim(JEVby)) + '" & Trim(ActiveUserID) & "',actioncode =3 Where ptvno='" & Trim(txtDVNo.Text) & "' And ActionCode=1"
'                        End If
                            Call SaveAcctngEntries(txtDVNo.Text)
                            Call GEtCompleteJEVDetails(txtDVNo.Text, "PTV", "", "", "" _
                            , Replace(txtParticular.Text, "'", "''"), txt_JEVNO.Text, "", "", txtAmount.Text, "0", "0", getTranstype, "", "", "", txtFund.Text, "", "", txtAlobs.Text, "", ExtractJEVSNo(txt_JEVNO.Text), DatePost, txtDVNo.Text)
                            'Updating table from PTO....
                            
                            EditCount = False
                            opndbaseFMIS.Execute "Update tblCMS_CDcheckbook set AlreadySaved2JEV=1 WHERE DVNo='" & Trim(txtDVNo.Text) & "' And ActionCode=1"
                            MsgBox "JEV Post Successfully...!", vbInformation, "Syste Message"
                            Call Toolbar1_ButtonClick(Toolbar1.Buttons.Item(1))
                    End If
                End If
    Case "Delete":
                If Edited = True Then
                    If InStr(ChkIfAlreadyJEV(txtDVNo.Text), "Approved") <> 1 Then
                        If MsgBox("Are you sure you want to delete this transaction?", vbQuestion + vbYesNo) = vbYes Then
                            opndbaseFMIS.Execute "Update tblAMIS_FinalJEV set JEVby=case when JEVby is null then '' else RTRIM(LTRIM(JEVBY)) + ',' end + '" & Trim(ActiveUserID) & "',Actioncode=3,DateTimeEntered=rtrim(ltrim(DateTimeEntered)) +'," & Format(Now, "mm/dd/yyyy hh:mm:ss AMPM") & "'  Where PTVNO='" & txtDVNo.Text & "' and Actioncode=1"
                            opndbaseFMIS.Execute "Update  tblAMIS_AccoutingEntries set Actioncode=3 Where reffno='" & txtDVNo.Text & "' and Actioncode=1"
                            opndbaseFMIS.Execute "Update tblCMS_CDcheckbook set AlreadySaved2JEV=0 WHERE DVNo='" & Trim(txtDVNo.Text) & "' And ActionCode=1"
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
Call LoadErr(err.Number, err.Source & ", " & Me.name & ", " & Me.Caption & ".Toolbar1_ButtonClick", err.description)
End Sub
Private Function getTranstype() As Integer
If optCollection.Value = True Then
getTranstype = 1
Else
getTranstype = 4
End If
End Function
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
Public Function SaveAcctngEntries(ByVal dvno As String)
Dim DRec As New ADODB.Recordset
Dim rec As New ADODB.Recordset
Dim x As Integer
Dim xType As Integer
If optCollection.Value = True Then xType = CInt(optCollection.Tag)
If optOther.Value = True Then xType = CInt(optOther.Tag)
    DRec.Open ("Select Accountcode,sum(Debit) as Debit ,sum(Credit) as Credit From tblAMIs_tmpjournal Where [dvno]='" & dvno & "' group by Accountcode"), opndbaseFMIS, adOpenStatic, adLockOptimistic
    If DRec.RecordCount > 0 Then
        opndbaseFMIS.Execute "update tblAMIS_AccoutingEntries set actioncode =2 where reffno = '" & txtDVNo.Text & "' and actioncode =1" ', datetimeentered = rtrim(ltrim(DateTimeEntered)) +'," & Format(Now, "yyyy/mm/dd hh:mm:ss AMPM") & "',userid = UserID + '," & Trim(ActiveUserID) & "'
        For x = 1 To DRec.RecordCount
            opndbaseFMIS.Execute "Insert into tblAMIS_AccoutingEntries (reffNo,ChildAccountcode,debit,credit,actioncode,datetimeentered,transtype,userid) values " & _
            "('" & Trim(txtDVNo.Text) & "','" & Trim(DRec!accountcode) & "'," & DRec!Debit & "," & DRec!Credit & ",1,'" & Format(Now, "yyyy/mm/dd hh:mm:ss AMPM") & "'," & xType & ",'" & Trim(ActiveUserID) & "')"
            DRec.MoveNext
            DoEvents
        Next x
        opndbaseFMIS.Execute "delete from tblAMIs_tmpjournal where dvno = '" & txtDVNo.Text & "'"
    End If
    DRec.Close
    Set DRec = Nothing
End Function
Private Function GetPrevTrans(ByVal YEAR_ As Long, ByVal month_ As Integer)
Dim rec As New ADODB.Recordset
Dim x As Integer
List1.Clear
rec.Open "Exec GetPreTrans @Year = " & YEAR_ & ",@month = '" & month_ & "',@Whatfield = 'PTVNO'", opndbaseFMIS, adOpenStatic, adLockOptimistic
    If rec.RecordCount > 0 Then
        For x = 1 To rec.RecordCount
            List1.AddItem rec.Fields!ptvNo
            rec.MoveNext
        Next x
    End If
rec.Close
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

Private Function CheckWhatENtries() As Integer
Dim x As Long
CheckWhatENtries = 4
For x = 1 To MSFlexGrid1.Rows - 1
    If MSFlexGrid1.TextMatrix(x, 1) = "101" Then
        CheckWhatENtries = 1
        Exit For
    End If
Next x
End Function
Private Function ChkEntry() As Boolean
Dim transname As String
Dim opname As String
    ChkEntry = False
    If getTranstype <> CheckWhatENtries Then
        If getTranstype = 1 Then
        transname = "Cash Receipts"
        opname = "General Journal"
        Else
        transname = "General Journal"
        opname = "Cash Receipts"
        End If
    End If
    If CheckIfExistPTVinFinalJEV(txtDVNo.Text) = True Then 'check if EXIST PTV number  in Final JEV
        MsgBox "PTV number is Already Posted...Please Check your Entry", vbCritical, "System Message"
        Exit Function
    End If
    If Trim(txtDVNo.Text) <> "" And txtAlobs.Text <> "" And txt_JEVNO.Text <> "" And txtParticular.Text <> "" And txtFund.Text <> "" And txtAmount.Text <> "" Then
        If getTranstype <> CheckWhatENtries Then
            If MsgBox("Opps! " & opname & " Entry" & vbNewLine & "Are you sure this Entry is for " & transname & "?", vbCritical + vbYesNo, "System Message") = vbNo Then
            Exit Function
            End If
        End If
        If optCollection.Value = False And optOther.Value = False Then
        MsgBox "Please Specify the Transaction type..", vbCritical, "System Message"
        Exit Function
        End If
        If IsFormatCorrect(txt_JEVNO.Text) = True Then '
        
            If Edited = False Then
                If CheckIfExistInFinalJEV(txt_JEVNO.Text) = False Then 'check if EXIST JEV no  in Final JEV
                    If xDebit = xCredit And xDebit > 0 Then
                        If getTranstype = 4 Then
                            ChkEntry = True
                        Else
                            If Format(xDebit, "###,##0.00") = Format((2 * CCur(txtAmount.Text)), "###,##0.00") Then
                                ChkEntry = True
                            Else
                                If MsgBox("Your TOTAL ENTRY Amount is not Equal to the Expected amount, Are you sure do you want to proceed?", vbInformation + vbYesNo, "System Message") = vbYes Then
                                    ChkEntry = True
                                End If
                            End If
                        End If
                    Else
                    MsgBox "Your Total Debit and Total Credit Amount is not Equal or Accounting Entry is Empty, Please Check Your Entry", vbInformation, "System Message"
                    End If
                Else
                    MsgBox "JEV number already exist in the Database..", vbInformation, "System Message"
                        If MsgBox("Do you want to Generate JEV Number?", vbInformation + vbYesNo, "System Confirmation") = vbYes Then
                          Call Command1_Click
                        End If
                End If
            Else
                If xDebit = xCredit And xDebit > 0 Then
                        If getTranstype = 4 Then
                    
                            ChkEntry = True
                        Else
                            If Format(xDebit, "###,##0.00") = Format((2 * CCur(txtAmount.Text)), "###,##0.00") Then
                                ChkEntry = True
                            Else
                            If MsgBox("Your Entry Amount is not related to the transaction type, Are you sure do you want to proceed?", vbInformation + vbYesNo, "System Message") = vbYes Then
                                 ChkEntry = True
                            Else
                            Exit Function
                            End If
                            End If
                        End If
                Else
                MsgBox "Your Total Debit and Total Credit Amount is not Equal or Accounting Entry is Empty, Please Check Your Entry", vbInformation, "System Message"
                End If
            End If
        Else
            MsgBox "JEV Number is Incorrect Format..Please Check it", vbInformation, "System Information"""
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
        MSFlexGrid1.SetFocus
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
    Jrec.Open ("Select Ptvno from tblAMIS_FinalJEV where PTVNO='" & dvno & "' and Actioncode=1 "), opndbaseFMIS, adOpenStatic, adLockOptimistic
    If Jrec.RecordCount > 0 Then
            ChkIfAlreadyJEV = dvno
    End If
    Jrec.Close
    Set Jrec = Nothing
End Function

Private Sub txtDVNo_KeyPress(KeyAscii As Integer)
Dim DVRec As New ADODB.Recordset
Dim Jrec As New ADODB.Recordset
Dim rec As New ADODB.Recordset
Dim xAlreadyJEV As String
Dim x As Integer
If KeyAscii = 13 Then
        btnReturn.Enabled = False
        CUFlag = False
        txtParticular.Locked = True
        
        xNAcode = ""
        txtDVNo.Text = Trim(txtDVNo.Text)
        If ChkPTVExist(txtDVNo.Text) = True Then
            xAlreadyJEV = ChkIfAlreadyJEV(txtDVNo.Text)
            If xAlreadyJEV = "" Then
                Set DVRec = opndbaseFMIS.Execute("Select * FRom tblCMS_CDCheckBook where DVNo='" & txtDVNo.Text & "' and (ActionCode=1)")
                If DVRec.RecordCount > 0 Then
                    If DVRec!TransactionType = "DA" Or DVRec!TransactionType = "CA" Then
                    optOther.Value = True
                    ElseIf DVRec!TransactionType = "DS" Then
                    optCollection.Value = True
                    Else
                    optCollection.Value = False
                    optOther.Value = False
                    End If
                    
                    txtAlobs.Text = Trim(DVRec!chknumber)
                    txtParticular.Text = Trim(DVRec!Particular)
                    txtFund.Text = GetSFNameByCode(Left(txtDVNo.Text, 3))
                    txtAmount.Text = Format(DVRec!amount, "#,##0.00")
                    Edited = False
                    IsSaveAccntng = False
                    If MsgBox("Do you want to Generate Accounting Entries for this transaction?", vbInformation + vbYesNo, "System Message") = vbYes Then
'                        If CheckIfHavenullAccnt(txtDVNo.Text, 1) = True Then
''                            Set jrec = opndbaseFMIS.Execute("Exec MPproc_ChckIfHaveAccnt @reff = '" & txtDVNo.Text & "',@type = 1")
''                            If jrec.RecordCount > 0 Then
'                                With frmOtherClass
'                                     'Call Load(frmOtherClass)
'                                     .cmblistindex = 7
'                                     .Loadformload
'                                    'Call Load(frmOtherClass)
'                                    .Reff = txtDVNo.Text
'                                   ' Set .Combo1.ItemData(.Combo1.ListIndex) = 7
'
'                                    .Show 1
'                                End With
''                            End If
'                        End If
                        rec.Open "Select dvno from tblAMIs_tmpjournal where dvno = '" & txtDVNo.Text & "'", opndbaseFMIS, adOpenStatic
                            If rec.RecordCount > 0 Then
                                If MsgBox("This transaction Have a temporary Accounting Entries, do you want to Delete?", vbCritical + vbYesNo, "System Information") = vbYes Then
                                    opndbaseFMIS.Execute "Delete from tblAMIs_tmpjournal where Dvno = '" & txtDVNo.Text & "'"
                                        Set Jrec = opndbaseFMIS.Execute("Exec MPproc_GenerateAccntngEntries @reff = '" & txtDVNo.Text & "',@type = 1")
                                        If Jrec.RecordCount > 0 Then
                                            opndbaseFMIS.Execute "Insert into tblAMIs_tmpjournal (Dvno,Accountcode,Debit,Credit) values ('" & Trim(txtDVNo.Text) & "','101','" & txtAmount.Text & "','" & txtAmount.Text & "')"
                                           For x = 1 To Jrec.RecordCount
                                               opndbaseFMIS.Execute "Insert into tblAMIs_tmpjournal (Dvno,Accountcode,Debit,Credit) values ('" & Trim(txtDVNo.Text) & "','" & Trim(Jrec!accntcode) & "'," & IIf(IsNull(Jrec!Debit), 0, Jrec!Debit) & "," & IIf(IsNull(Jrec!Credit), 0, Jrec!Credit) & ")"
                                               Jrec.MoveNext
                                               DoEvents
                                           Next x
                                        End If
                                        Call GetAccntngEntries
                                        Jrec.Close
                                End If
                            Else
                                        Set Jrec = opndbaseFMIS.Execute("Exec MPproc_GenerateAccntngEntries @reff = '" & txtDVNo.Text & "',@type = 1")
                                        If Jrec.RecordCount > 0 Then
                                        opndbaseFMIS.Execute "Insert into tblAMIs_tmpjournal (Dvno,Accountcode,Debit,Credit) values ('" & Trim(txtDVNo.Text) & "','101'," & CCur(txtAmount.Text) & "," & CCur(txtAmount.Text) & ")"
                                           For x = 1 To Jrec.RecordCount
                                               opndbaseFMIS.Execute "Insert into tblAMIs_tmpjournal (Dvno,Accountcode,Debit,Credit) values ('" & Trim(txtDVNo.Text) & "','" & Trim(Jrec!accntcode) & "'," & IIf(IsNull(Jrec!Debit), 0, Jrec!Debit) & "," & IIf(IsNull(Jrec!Credit), 0, Jrec!Credit) & ")"
                                               Jrec.MoveNext
                                               DoEvents
                                           Next x
                                        End If
                                        Call GetAccntngEntries
                                        Jrec.Close
                            End If
                    End If
                End If
                DVRec.Close
                Set DVRec = Nothing
            Else
                IsSaveAccntng = False
                Call LoadJEVDetails(xAlreadyJEV)
            End If
        Else
            MsgBox "Invalid PTV Number!", vbExclamation
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
                            If Trim(DVRec!fundcode) = "126" Or Trim(DVRec!fundcode) = "124" Or Trim(DVRec!fundcode) = "114" Or Trim(DVRec!fundcode) = "104" Then
                            fundcodes = 119
                            Else
                            fundcodes = Trim(DVRec!fundcode)
                            End If
                            txtFund.Text = GetSFNameByCode(Left(txtDVNo.Text, 3))
                            txtAmount.Text = Format(DVRec!amount, "#,##0.00")
                            
                            
                            
                            'Call LoadAccountsByFund(Trim(txtFund.Text))
                            
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

VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "Mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{1693405E-2DC9-4248-B52F-4AC9145DA2AF}#1.0#0"; "WinXPC Engine.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form frmCDCashDisbursedReport 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "JEV Numbering for Cash Disbursement Report"
   ClientHeight    =   10230
   ClientLeft      =   2280
   ClientTop       =   1170
   ClientWidth     =   14490
   Icon            =   "frmCDCashDisbursedReport.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   10230
   ScaleWidth      =   14490
   Begin VB.CommandButton Command1 
      Caption         =   "Load Reports"
      Height          =   435
      Left            =   1200
      TabIndex        =   37
      Top             =   3015
      Width           =   1125
   End
   Begin VB.Frame Frame5 
      Height          =   3825
      Left            =   2535
      TabIndex        =   13
      Top             =   885
      Width           =   11835
      Begin VB.CommandButton cmd_post 
         Caption         =   "Post (JEV No.)"
         Height          =   1005
         Left            =   1800
         Picture         =   "frmCDCashDisbursedReport.frx":0E42
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   2640
         Visible         =   0   'False
         Width           =   1755
      End
      Begin VB.CommandButton cmd_Mass 
         Caption         =   "Mass JEV Nos."
         Height          =   1005
         Left            =   120
         Picture         =   "frmCDCashDisbursedReport.frx":493C
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   2640
         Visible         =   0   'False
         Width           =   1635
      End
      Begin VB.Frame Frame3 
         Caption         =   "Source of Fund"
         Height          =   1245
         Left            =   60
         TabIndex        =   16
         Top             =   1275
         Width           =   11730
         Begin VB.ComboBox cmb_accnumber 
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
            Left            =   7410
            TabIndex        =   20
            Top             =   270
            Width           =   3525
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
            Left            =   7410
            TabIndex        =   19
            Top             =   675
            Width           =   3525
         End
         Begin VB.ComboBox cmb_bank 
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
            Left            =   1530
            TabIndex        =   18
            Top             =   675
            Width           =   3525
         End
         Begin VB.ComboBox cmb_Fund 
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
            Left            =   1530
            Locked          =   -1  'True
            TabIndex        =   17
            Top             =   270
            Width           =   3525
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Account Number"
            Height          =   195
            Left            =   6060
            TabIndex        =   24
            Top             =   375
            Width           =   1200
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Account Name"
            Height          =   195
            Left            =   6060
            TabIndex        =   23
            Top             =   795
            Width           =   1065
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Drawee Bank"
            Height          =   195
            Left            =   180
            TabIndex        =   22
            Top             =   795
            Width           =   975
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fund Type"
            Height          =   195
            Left            =   180
            TabIndex        =   21
            Top             =   375
            Width           =   765
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Report No."
         Height          =   885
         Left            =   75
         TabIndex        =   14
         Top             =   210
         Width           =   4125
         Begin VB.TextBox txt_RDNo 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Left            =   150
            Locked          =   -1  'True
            TabIndex        =   15
            Top             =   255
            Width           =   3765
         End
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   1065
         Left            =   4080
         TabIndex        =   25
         Top             =   2580
         Width           =   7560
         _ExtentX        =   13335
         _ExtentY        =   1879
         View            =   3
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Check Number"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   1
            Text            =   "Check Amount"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Check Date"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "RCI No."
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Shape Shape3 
         BackColor       =   &H00000000&
         BackStyle       =   1  'Opaque
         Height          =   1155
         Left            =   0
         Top             =   2565
         Visible         =   0   'False
         Width           =   3690
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Check Amount:"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   4380
         TabIndex        =   35
         Top             =   300
         Width           =   1500
      End
      Begin VB.Label lbl_CheckAmount 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "000.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   5925
         TabIndex        =   34
         Top             =   300
         Width           =   1890
      End
      Begin VB.Label lbl_total 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "000.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   5925
         TabIndex        =   33
         Top             =   750
         Width           =   1890
      End
      Begin VB.Label lbl_TotalLiqAmount 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "000.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   9795
         TabIndex        =   32
         Top             =   300
         Width           =   1890
      End
      Begin VB.Label lbl_LackingAmount 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "000.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   9795
         TabIndex        =   31
         Top             =   750
         Width           =   1890
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Cash Advance :"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   4380
         TabIndex        =   30
         Top             =   750
         Width           =   1545
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Liquidated Amount :"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   7965
         TabIndex        =   29
         Top             =   300
         Width           =   1815
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Lacking Amount :"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   7965
         TabIndex        =   28
         Top             =   750
         Width           =   1650
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00000000&
         BackStyle       =   1  'Opaque
         Height          =   2265
         Left            =   -15
         Top             =   -1020
         Width           =   11880
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Special Account"
      Height          =   780
      Left            =   75
      TabIndex        =   11
      Top             =   1035
      Width           =   2250
      Begin VB.ComboBox cmb_FundType 
         Height          =   315
         Left            =   75
         Style           =   2  'Dropdown List
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
      ToolTipText     =   "Type only number then Enter (""FMISNo-00"") will apear"
      Top             =   3975
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
      Top             =   4845
      Width           =   2265
   End
   Begin VB.Frame Frame6 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "FMIS Report Nos."
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   60
      TabIndex        =   6
      Top             =   4530
      Width           =   2280
   End
   Begin VB.Frame Frame1 
      Caption         =   "For the Period"
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
         Format          =   149880835
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
      Height          =   5160
      Left            =   2535
      ScaleHeight     =   5130
      ScaleWidth      =   11820
      TabIndex        =   1
      Top             =   4755
      Width           =   11850
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   5130
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   11820
         _ExtentX        =   20849
         _ExtentY        =   9049
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
      Left            =   4770
      Top             =   9345
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
      Width           =   14490
      _ExtentX        =   25559
      _ExtentY        =   1482
      ButtonWidth     =   2117
      ButtonHeight    =   1429
      Appearance      =   1
      Style           =   1
      ImageList       =   "itb32x32"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   6
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Print Report"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "slash"
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Manual Entry"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Close"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
      EndProperty
      Begin MSComctlLib.ImageList itb32x32 
         Left            =   4320
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
               Picture         =   "frmCDCashDisbursedReport.frx":8436
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCDCashDisbursedReport.frx":9DC8
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCDCashDisbursedReport.frx":B75A
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCDCashDisbursedReport.frx":D0EC
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCDCashDisbursedReport.frx":EA7E
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCDCashDisbursedReport.frx":10410
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCDCashDisbursedReport.frx":11DA2
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCDCashDisbursedReport.frx":13734
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCDCashDisbursedReport.frx":150C6
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCDCashDisbursedReport.frx":16A5A
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCDCashDisbursedReport.frx":17736
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCDCashDisbursedReport.frx":18016
               Key             =   ""
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCDCashDisbursedReport.frx":18CF2
               Key             =   ""
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCDCashDisbursedReport.frx":199CE
               Key             =   ""
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCDCashDisbursedReport.frx":1A6AA
               Key             =   ""
            EndProperty
            BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCDCashDisbursedReport.frx":1B386
               Key             =   ""
            EndProperty
            BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCDCashDisbursedReport.frx":1C062
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComCtl2.Animation Animation1 
         Height          =   450
         Left            =   11400
         TabIndex        =   40
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
      Top             =   9255
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
            Picture         =   "frmCDCashDisbursedReport.frx":1C93E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCDCashDisbursedReport.frx":1D9C0
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCDCashDisbursedReport.frx":1FAFA
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCDCashDisbursedReport.frx":20D7C
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCDCashDisbursedReport.frx":23C06
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
      TabIndex        =   39
      Top             =   9975
      Width           =   570
   End
   Begin VB.Label Label13 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Label13"
      ForeColor       =   &H0000FF00&
      Height          =   300
      Left            =   90
      TabIndex        =   38
      Top             =   3015
      Width           =   1020
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Search FMIS No.:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   60
      TabIndex        =   27
      Top             =   3705
      Width           =   1275
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label8"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   105
      TabIndex        =   10
      Top             =   9300
      Width           =   480
   End
   Begin VB.Label Label7 
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
Attribute VB_Name = "frmCDCashDisbursedReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim tmpAccName As String
Dim FMISNo As String


Private Sub cmb_accnumber_Click()
Call LoadAccountName(cmb_accnumber.Text, cmb_Fund.Text, cmb_bank.Text)
End Sub


Private Sub cmb_Bank_Click()
Call LoadBankAccntNo(cmb_bank.Text, cmb_Fund.Text)
cmb_AccountName.Clear
End Sub

Private Sub cmb_Fund_click()
Call LoadDraweeBank
cmb_accnumber.Clear
cmb_AccountName.Clear

End Sub

Private Sub cmb_Fund_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    MsgBox cmb_Fund.ItemData(cmb_Fund.ListIndex)
End If
End Sub



Private Function GetTotalCheckAmount(ByVal RecordID As String) As Currency
Dim totalCheck As New ADODB.Recordset

totalCheck.Open " select sum(NetAmount) as Amount from vw_CDCashAdvancedChecks where mixcode='" & RecordID & "'", opndbaseFMIS, adOpenStatic, adLockOptimistic
    GetTotalCheckAmount = IIf(IsNull(totalCheck!amount), 0, totalCheck!amount)
totalCheck.Close
Set totalCheck = Nothing
End Function



Private Sub cmb_FundType_Change()
clr
End Sub
Private Function clr()
Call SetGrid
ListView1.ListItems.Clear
List2.Clear
txt_RDNo.Text = ""
cmb_accnumber.Text = ""
cmb_AccountName.Text = ""
cmb_Fund.Text = ""
cmb_bank.Text = ""
End Function

Private Sub cmd_Mass_Click()
JevOk = False
frmPOstdate.Show 1
If JevOk = True Then
Label13.Caption = "JEV Numbering..."
Label13.Refresh
Animation1.Visible = True
Animation1.Open App.path & AViLocation & "\horizontaloading.avi"
Animation1.Play
Call JEVMassNumbering(cmb_Fund.Text)
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
Dim dvno As String
Dim LastJEVSNno As Long
    
rec.Open ("EXEC [dbo].[Proc_GetMaxJevSeries_New] @transtype = 3,@jevyeardate = '" & DatePost & "' ,@fundtype = '" & cmb_FundType.Text & "'"), opndbaseFMIS, adOpenStatic, adLockOptimistic
 LastJEVSNno = rec.Fields!MAXJEVSERIES
 rec.Close
For cc = 1 To MSHFlexGrid1.Rows - 1


    If Len(MSHFlexGrid1.TextMatrix(cc, 8)) > 0 Then
    
        sql = "SELECT tblAMIS_IncomingDVTrns.FundType as FundType, tblAMIS_JournalEntry.TransType as TransType, tblAMIS_JournalEntry.DVNo as DVNo, " & _
                "          tblAMIS_JournalEntry.TransDate as TransDate, tblAMIS_JournalEntry.JEVSeriesNo as JEVSeriesNo,(Select FundCode from tblRefBMS_Funds where FundMedium=tblAMIS_IncomingDVTrns.FundType) as FundCode " & _
                " FROM tblAMIS_IncomingDVTrns INNER JOIN " & _
                "          tblAMIS_JournalEntry ON tblAMIS_IncomingDVTrns.DVNo = tblAMIS_JournalEntry.DVNo " & _
                " Where (tblAMIS_JournalEntry.ActionCode = 1) And (tblAMIS_IncomingDVTrns.ActionCode = 1) " & _
                " GROUP BY tblAMIS_IncomingDVTrns.FundType, tblAMIS_JournalEntry.TransType, tblAMIS_JournalEntry.DVNo, " & _
                "          tblAMIS_JournalEntry.TransDate , tblAMIS_JournalEntry.JEVSeriesNo " & _
                " HAVING   tblAMIS_JournalEntry.DVNo ='" & MSHFlexGrid1.TextMatrix(cc, 8) & "'"
    
        opnJEV.Open sql, opndbaseFMIS, adOpenStatic, adLockOptimistic
        If opnJEV.RecordCount <> 0 Then
            MSHFlexGrid1.TextMatrix(cc, 10) = cmb_FundType.ItemData(cmb_FundType.ListIndex) & "-" & Right(Year(DatePost), 2) & "-" & Format(Month(DatePost), "00") & "-03-" & Format(LastJEVSNno, "0000")
            LastJEVSNno = LastJEVSNno + 1
        Else 'No REcord Found yet in the AMIS
            MSHFlexGrid1.TextMatrix(cc, 10) = "000-00-00-00-xxxxx"
        End If
        opnJEV.Close
        Set opnJEV = Nothing

        
    End If
Next cc
End Sub

Private Sub cmd_post_Click()
Dim cc As Integer

If MsgBox("Save JEV Nos.?", vbQuestion + vbYesNo, "System Confirmation") = vbYes Then
 Animation1.Visible = True
Animation1.Open App.path & AViLocation & "\horizontaloading.avi"
Animation1.Play
    For cc = 1 To MSHFlexGrid1.Rows - 1
        
            If Len(Trim(MSHFlexGrid1.TextMatrix(cc, 10))) > 0 Then
                If IsFormatCorrect(MSHFlexGrid1.TextMatrix(cc, 10)) = True Then
                    Call GEtCompleteJEVDetails(MSHFlexGrid1.TextMatrix(cc, 8), "DVNO", ListView1.ListItems(1).ListSubItems(2).Text, ListView1.ListItems(1).ListSubItems(3).Text, ListView1.ListItems(1).Text _
                        , "", MSHFlexGrid1.TextMatrix(cc, 10), "", "", "0", "0", "0", "3", List2.List(List2.ListIndex), MSHFlexGrid1.TextMatrix(cc, 10), "", cmb_FundType.Text, "", "", "", "", ExtractJEVSNo(MSHFlexGrid1.TextMatrix(cc, 10)), DatePost, "")
'
'MSHFlexGrid1.TextMatrix(0, 0) = "Trnno"
'MSHFlexGrid1.TextMatrix(0, 1) = "Payee"
'MSHFlexGrid1.TextMatrix(0, 2) = "Payment Period"
'MSHFlexGrid1.TextMatrix(0, 3) = "Ref."
'MSHFlexGrid1.TextMatrix(0, 4) = "Code"
'MSHFlexGrid1.TextMatrix(0, 5) = "Amount"
'MSHFlexGrid1.TextMatrix(0, 6) = "Liq. Amt."
'MSHFlexGrid1.TextMatrix(0, 7) = "Norm Bal."
'MSHFlexGrid1.TextMatrix(0, 8) = "DVNo"
'MSHFlexGrid1.TextMatrix(0, 9) = "AlreadySaved"
'MSHFlexGrid1.TextMatrix(0, 10) = "JEVNo"
                    
                    'Updating table from PTO....
                    opndbaseFMIS.Execute "Update tblCMS_CDCashBook set AlreadySaved2JEV=1,DatePostedtoJEV='" & Date & "',PostedtoJEVUserid='" & ActiveUserID & "' where trnno=" & MSHFlexGrid1.TextMatrix(cc, 0) & ""
                    
                    'Updating Accounting REcord...
                    opndbaseFMIS.Execute "update tblAMIS_JournalEntry set JEVNo='" & MSHFlexGrid1.TextMatrix(cc, 10) & "',JEVSeriesNo=" & ExtractJEVSNo(MSHFlexGrid1.TextMatrix(cc, 10)) & ",JEVBy='" & ActiveUserID & "',JEVDate='" & DatePost & "',transtype = 3 where DVNo='" & MSHFlexGrid1.TextMatrix(cc, 8) & "'"
                    
                End If
            End If
       
    Next cc
Animation1.Stop
Animation1.Close
Animation1.Visible = False
MsgBox "Posting to JEV, Successful!", vbInformation, "System Information"
Command1_Click 'Loading Back Active Cash Disbursement Numbers...
List2.ListIndex = GetIndex4ListBox(List2, FMISNo)
End If
End Sub
Private Sub Timer1_Timer()
Call SetGrid
Call LoadFundType(cmb_FundType)
Timer1.Enabled = False
End Sub
Private Sub Command1_Click()
Label13.Caption = "Loading, Please wait.."
Label13.Refresh
Animation1.Visible = True
Animation1.Open App.path & AViLocation & "\horizontaloading.avi"
Animation1.Play
Call LoadSavedReport(ActiveUserID, DTPicker1.Year, DTPicker1.Month, cmb_FundType.Text)
Animation1.Stop
Animation1.Close
Animation1.Visible = False
Label13.Caption = ""
End Sub

Private Sub DTPicker1_Change()
DTPicker1.Value = DTPicker1.Month & "/1/" & DTPicker1.Year
Label6.Caption = MonthName(DTPicker1.Month) & " " & DTPicker1.Year
clr
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then
        Unload Me
End If
End Sub

Private Sub ClearCmb()
cmb_Fund.Clear
cmb_bank.Clear
cmb_accnumber.Clear
cmb_AccountName.Clear

End Sub
Public Sub LoadFund(ByVal cmb As ComboBox)
Dim opnfund As New ADODB.Recordset
Dim cc As Integer
opnfund.Open "Select * from tblRefBMS_Funds", opndbaseFMIS, adOpenStatic, adLockOptimistic
If opnfund.RecordCount <> 0 Then
    cmb.Clear
    Do Until opnfund.EOF
        cmb.AddItem (opnfund!FundName)
        cmb.ItemData(cc) = opnfund!fundcode
        cc = cc + 1
        opnfund.MoveNext
    Loop
Else
    cmb.Clear
End If
opnfund.Close
Set opnfund = Nothing
End Sub
Private Sub Clear()
lbl_CheckAmount.Caption = ""
lbl_total.Caption = ""
lbl_TotalLiqAmount.Caption = ""
lbl_LackingAmount.Caption = ""
Label13.Caption = ""
Label14.Caption = ""
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
Call SetGrid
Call LoadFundType(cmb_FundType)
'LoadDraweeBank
'LoadBankAccntNo
'Call LoadSavedReport(ActiveUserID, DTPicker1.Year, DTPicker1.Month, cmb_FundType.Text)
End Sub

Private Sub Form_Unload(Cancel As Integer)
WindowsXPC1.EndWinXPCSubClassing
Set frmCDCashDisbursedReport = Nothing
End Sub

Private Sub SetGrid()
Dim cc As Integer
MSHFlexGrid1.Clear
MSHFlexGrid1.Cols = 11
MSHFlexGrid1.Rows = 2
MSHFlexGrid1.TextMatrix(0, 0) = "Trnno"
MSHFlexGrid1.TextMatrix(0, 1) = "Payee"
MSHFlexGrid1.TextMatrix(0, 2) = "Payment Period"
MSHFlexGrid1.TextMatrix(0, 3) = "Ref."
MSHFlexGrid1.TextMatrix(0, 4) = "Code"
MSHFlexGrid1.TextMatrix(0, 5) = "Amount"
MSHFlexGrid1.TextMatrix(0, 6) = "Liq. Amt."
MSHFlexGrid1.TextMatrix(0, 7) = "Norm Bal."
MSHFlexGrid1.TextMatrix(0, 8) = "DVNo"
MSHFlexGrid1.TextMatrix(0, 9) = "AlreadySaved"
MSHFlexGrid1.TextMatrix(0, 10) = "JEVNo"

MSHFlexGrid1.ColWidth(0) = 800
MSHFlexGrid1.ColWidth(1) = 2300
MSHFlexGrid1.ColWidth(2) = 1400
MSHFlexGrid1.ColWidth(3) = 600
MSHFlexGrid1.ColWidth(4) = 1000
MSHFlexGrid1.ColWidth(5) = 1000
MSHFlexGrid1.ColWidth(6) = 1000
MSHFlexGrid1.ColWidth(7) = 800
MSHFlexGrid1.ColWidth(8) = 1200
MSHFlexGrid1.ColWidth(9) = 0
MSHFlexGrid1.ColWidth(10) = 1400

For cc = 0 To MSHFlexGrid1.Cols - 1
    MSHFlexGrid1.Row = 0
    MSHFlexGrid1.col = cc
    MSHFlexGrid1.CellAlignment = 4
Next cc
End Sub

Private Function GEtTotal() As Currency
Dim cc As Integer

For cc = 1 To MSHFlexGrid1.Rows - 1
    If GEtTotal <> 0 Then
        GEtTotal = GEtTotal + CCur(MSHFlexGrid1.TextMatrix(cc, 5))
    Else
        GEtTotal = CCur(MSHFlexGrid1.TextMatrix(cc, 5))
    End If
Next cc
End Function
Private Sub LoadCACheck(ByVal RecordID As String)
Dim opnCACheck As New ADODB.Recordset
Dim sitem As ListItem
Dim i As Integer


opnCACheck.Open "Select * from vw_CDCashAdvancedChecks where mixcode='" & RecordID & "'", opndbaseFMIS, adOpenStatic, adLockOptimistic
If opnCACheck.RecordCount <> 0 Then
    ListView1.ListItems.Clear
    Do Until opnCACheck.EOF
            Set sitem = ListView1.ListItems.Add()
            sitem.Text = opnCACheck!checkno
            sitem.SubItems(1) = Format(opnCACheck!NetAmount, "###,##0.00")
            sitem.SubItems(2) = opnCACheck!CheckDate
            sitem.SubItems(3) = GetRCINoPerCheck(opnCACheck!checkno)
        opnCACheck.MoveNext
    Loop
Else
    ListView1.ListItems.Clear
End If
opnCACheck.Close
Set opnCACheck = Nothing
End Sub

Private Sub LoadBreakdown(ByVal FMISVoucher As String, ByVal UserID As String)
Dim opnvoucher As New ADODB.Recordset

opnvoucher.Open "Select * from vw_CDCashAdvancedBreakDown where RecordID='" & FMISVoucher & "' and userid='" & UserID & "' and debitcredit=0 order by cbtrnno", opndbaseFMIS, adOpenStatic, adLockOptimistic

If opnvoucher.RecordCount <> 0 Then
    Call SetGrid
    MSHFlexGrid1.Rows = opnvoucher.RecordCount + 1
    Do Until opnvoucher.EOF
        MSHFlexGrid1.TextMatrix(opnvoucher.AbsolutePosition, 0) = opnvoucher!CBTrnno
        MSHFlexGrid1.TextMatrix(opnvoucher.AbsolutePosition, 1) = IIf(IsNull(opnvoucher!Claimant), "", opnvoucher!Claimant)
        MSHFlexGrid1.TextMatrix(opnvoucher.AbsolutePosition, 2) = IIf(IsNull(opnvoucher!PaymentPeriod), "", opnvoucher!PaymentPeriod)
        MSHFlexGrid1.TextMatrix(opnvoucher.AbsolutePosition, 3) = IIf(IsNull(opnvoucher!RefNo), "", opnvoucher!RefNo)
        MSHFlexGrid1.TextMatrix(opnvoucher.AbsolutePosition, 4) = opnvoucher!fundcode & "-" & opnvoucher!MotherFund
        MSHFlexGrid1.TextMatrix(opnvoucher.AbsolutePosition, 5) = Format(opnvoucher!amount, "###,##0.00") 'Cash Advanced Amount
        MSHFlexGrid1.TextMatrix(opnvoucher.AbsolutePosition, 6) = Format(opnvoucher!amount, "###,##0.00") 'Liquiditing Amount
        MSHFlexGrid1.TextMatrix(opnvoucher.AbsolutePosition, 7) = Format(0, "###,##0.00") 'Normal Balance
        MSHFlexGrid1.TextMatrix(opnvoucher.AbsolutePosition, 8) = IIf(IsNull(opnvoucher!controlno), "", opnvoucher!controlno)
        MSHFlexGrid1.TextMatrix(opnvoucher.AbsolutePosition, 9) = 0
        opnvoucher.MoveNext
    Loop
    lbl_total.Caption = Format(GEtTotal, "###,##0.00")
    lbl_TotalLiqAmount.Caption = Format(GetTotalSelColAmount(6), "###,##0.00")
    lbl_LackingAmount.Caption = Format(GetTotalSelColAmount(7), "###,##0.00")
    
Else
    Call Clear
    Call SetGrid
End If
opnvoucher.Close
Set opnvoucher = Nothing
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

Private Sub LoadBankAccntNo(ByVal BankID As String, ByVal FundType As String)
Dim opnaccnt As New ADODB.Recordset

opnaccnt.Open "Select BankAccountNo from vw_DepositoryBank where BankID='" & BankID & "' and FundType='" & FundType & "' and active=1 group by BankAccountNo", opndbaseFMIS, adOpenStatic, adLockOptimistic
If opnaccnt.RecordCount <> 0 Then
    cmb_accnumber.Clear
    Do Until opnaccnt.EOF
        cmb_accnumber.AddItem (opnaccnt!BankAccountNo)
        opnaccnt.MoveNext
    Loop
Else
    cmb_accnumber.Clear
End If
opnaccnt.Close
Set opnaccnt = Nothing
End Sub
Private Sub LoadDraweeBank()
Dim Banks As Variant
Dim x As Integer

Banks = readTXTDATA("CollectionFactors", "DraweeBank", App.path & "\data\SystemDefault.ini")
Banks = Split(Banks, ",")

cmb_bank.Clear
For x = 0 To UBound(Banks)
    cmb_bank.AddItem (Banks(x))
Next x

End Sub

Private Sub LoadSavedReport(ByVal UserID As String, ByVal TrnYear As Integer, ByVal trnMonth As Integer, ByVal fund As String)
Dim opnvoucher As New ADODB.Recordset
Dim cc As Integer
Dim sql, SFCOde As String


'If fund = "Eco-PTC" Then: SFCOde = "2473,36052,36055"
'If fund = "Eco-PNB" Then: SFCOde = "2476"
'If fund = "Eco-WATERWORKS" Then: SFCOde = "2475,36056"
'If fund = "Eco-ASERBAC" Then: SFCOde = "2477"
'
'If fund = "Eco-PTC" Or fund = "Eco-PNB" Or fund = "Eco-WATERWORKS" Or fund = "Eco-ASERBAC" Then
'fund = "Economic Enterprises"
'End If

If fund = "Provincial Learning Center" Then: SFCOde = "2473,36052,36055"
If fund = "Agricultural Resource Center" Then: SFCOde = "2476"
If fund = "Patin-ay Waterworks System" Then: SFCOde = "2475,36056"
If fund = "Eco-ASERBAC" Then: SFCOde = "2477"

If fund = "Provincial Learning Center" Or fund = "Agricultural Resource Center" Or fund = "Patin-ay Waterworks System" Or fund = "Eco-ASERBAC" Then
fund = "Economic Enterprises"
End If

Select Case (fund):
Case "Economic Enterprises"
    sql = " SELECT  tblCMS_CDCashBook.RecordID, tblCMS_CDCashBook.CompositionCode " & _
                          "  FROM  tblCMS_CDCashBook INNER JOIN " & _
                          " vw_DepositoryBank ON tblCMS_CDCashBook.CompositionCode = vw_DepositoryBank.FMISAccountCode LEFT OUTER JOIN " & _
                          " tblCMS_CDPreparedCheck ON tblCMS_CDCashBook.RecordID = tblCMS_CDPreparedCheck.MixCode " & _
            " WHERE         (tblCMS_CDCashBook.Actioncode = 1) AND (tblCMS_CDCashBook.DebitCredit = 1) AND (tblCMS_CDPreparedCheck.actioncode = 1) AND " & _
                          " (tblCMS_CDCashBook.RDONo IS NOT NULL or LEN(tblCMS_CDCashBook.RDONo) <> 0) AND (tblCMS_CDCashBook.CompositionCode in (" & SFCOde & ")) AND " & _
                          " (YEAR(tblCMS_CDPreparedCheck.CheckDate) = " & TrnYear & ") AND (MONTH(tblCMS_CDPreparedCheck.CheckDate) = " & trnMonth & " ) " & _
            "and tblCMS_CDCashBook.AlreadySaved2JEV = 0 ORDER BY tblCMS_CDCashBook.RecordID "
Case Else
    sql = " SELECT  tblCMS_CDCashBook.RecordID, tblCMS_CDCashBook.CompositionCode FROM  tblCMS_CDCashBook INNER JOIN " & _
                      " vw_DepositoryBank ON tblCMS_CDCashBook.CompositionCode = vw_DepositoryBank.FMISAccountCode LEFT OUTER JOIN " & _
                      " tblCMS_CDPreparedCheck ON tblCMS_CDCashBook.RecordID = tblCMS_CDPreparedCheck.MixCode " & _
        " WHERE         (tblCMS_CDCashBook.Actioncode = 1) AND (tblCMS_CDCashBook.DebitCredit = 1) AND (tblCMS_CDPreparedCheck.actioncode = 1) AND " & _
                      " (tblCMS_CDCashBook.RDONo IS NOT NULL or LEN(tblCMS_CDCashBook.RDONo) <> 0) AND (vw_DepositoryBank.FundType = '" & fund & "') AND " & _
                      " (YEAR(tblCMS_CDPreparedCheck.CheckDate) = " & TrnYear & ") AND (MONTH(tblCMS_CDPreparedCheck.CheckDate) = " & trnMonth & ")  " & _
        "and tblCMS_CDCashBook.AlreadySaved2JEV = 0 ORDER BY tblCMS_CDCashBook.RecordID "
End Select

'opnvoucher.Open "Select RecordID,compositioncode from vw_CDCreateRDONo where RDONo is not null and year(checkdate)=" & TrnYear & " and month(checkdate)=" & trnMonth & " and userid='" & UserID & "' or len(RDONo)<>0 and year(checkdate)=" & TrnYear & " and month(checkdate)=" & trnMonth & " order by RecordID", opndbaseFMIS, adOpenStatic, adLockOptimistic

'Debug.Print sql

opnvoucher.Open sql, opndbaseFMIS, adOpenStatic, adLockOptimistic



If opnvoucher.RecordCount <> 0 Then
    List2.Clear
    Do Until opnvoucher.EOF
        List2.AddItem (opnvoucher!RecordID)
        List2.ItemData(cc) = opnvoucher!compositioncode
        cc = cc + 1
        opnvoucher.MoveNext
    Loop
Else
    List2.Clear
End If
opnvoucher.Close
Set opnvoucher = Nothing

Label8.Caption = List2.ListCount & "Record/s Found"
End Sub



Private Sub List2_Click()
Animation1.Visible = True
Animation1.Open App.path & AViLocation & "\horizontaloading.avi"
Animation1.Play
Label13.Caption = "Loading Details..."
Label13.Refresh

FMISNo = List2.Text

Call LoadBackBreakdown(FMISNo)

Call LoadCACheck(FMISNo)
txt_RecordID.Text = FMISNo
Label13.Caption = ""
Label14.Caption = (MSHFlexGrid1.Rows - 1) & " Voucher/s Found..."
Animation1.Stop
Animation1.Close
Animation1.Visible = False
End Sub
Private Sub LoadBackBreakdown(ByVal ReportNo As String)
Dim opnvoucher As New ADODB.Recordset
'On Error GoTo bad
Dim sql As String

sql = "Select * from vw_CDCashAdvancedBreakDown where RecordID='" & ReportNo & "' and CompositionCode=" & List2.ItemData(List2.ListIndex) & " and debitcredit=0 and AlreadySaved2JEV =0 order by cbtrnno"
opnvoucher.Open sql, opndbaseFMIS, adOpenStatic, adLockOptimistic

If opnvoucher.RecordCount <> 0 Then
    Call ClearCmb
    txt_RDNo.Text = ""

    'Loading Other Details------------------------
    Call LoadFundType(cmb_Fund)
   ' cmb_Fund.AddItem "Economics Enterprise"
'    cmb_Fund.ListIndex = 9
    
    txt_RDNo.Text = opnvoucher!RDOno
    lbl_CheckAmount.Caption = Format(GetTotalCheckAmount(ReportNo), "###,##0.00")
    
    cmb_Fund.ListIndex = GetIndex(cmb_Fund, opnvoucher!FundType)
    cmb_bank.ListIndex = GetIndex(cmb_bank, opnvoucher!BankID)
    'cmb_accnumber.ListIndex = GetIndex(cmb_accnumber, opnvoucher!BankAccountNo)
   ' cmb_AccountName.ListIndex = GetIndex(cmb_AccountName, opnvoucher!Accountname)
    '---------------------------------------------
    
    Call SetGrid
    MSHFlexGrid1.Rows = opnvoucher.RecordCount + 1
    Do Until opnvoucher.EOF
        MSHFlexGrid1.TextMatrix(opnvoucher.AbsolutePosition, 0) = IIf(IsNull(opnvoucher!CBTrnno), 0, opnvoucher!CBTrnno)
        MSHFlexGrid1.TextMatrix(opnvoucher.AbsolutePosition, 1) = IIf(IsNull(opnvoucher!Claimant), "", opnvoucher!Claimant)
        MSHFlexGrid1.TextMatrix(opnvoucher.AbsolutePosition, 2) = IIf(IsNull(opnvoucher!PaymentPeriod), "", opnvoucher!PaymentPeriod)
        MSHFlexGrid1.TextMatrix(opnvoucher.AbsolutePosition, 3) = IIf(IsNull(opnvoucher!RefNo), "", opnvoucher!RefNo)
        MSHFlexGrid1.TextMatrix(opnvoucher.AbsolutePosition, 4) = opnvoucher!fundcode & "-" & opnvoucher!MotherFund
        MSHFlexGrid1.TextMatrix(opnvoucher.AbsolutePosition, 5) = Format(opnvoucher!amount, "###,##0.00") 'Cash Advanced Amount
        MSHFlexGrid1.TextMatrix(opnvoucher.AbsolutePosition, 6) = IIf(IsNull(opnvoucher!RefundORNo), Format(opnvoucher!amount, "###,##0.00"), Format(opnvoucher!amount - opnvoucher!RefundORAmount, "###,##0.00")) 'Liquiditing Amount
        MSHFlexGrid1.TextMatrix(opnvoucher.AbsolutePosition, 7) = IIf(IsNull(opnvoucher!RefundORNo), Format(0, "###,##0.00"), Format(opnvoucher!RefundORAmount, "###,##0.00")) 'Normal Balance
        MSHFlexGrid1.TextMatrix(opnvoucher.AbsolutePosition, 8) = opnvoucher!controlno
        MSHFlexGrid1.TextMatrix(opnvoucher.AbsolutePosition, 9) = IIf(IsNull(opnvoucher!RefundORNo), 0, 1)
        
        
        opnvoucher.MoveNext
    Loop
    lbl_total.Caption = Format(GEtTotal, "###,##0.00")
    lbl_TotalLiqAmount.Caption = Format(GetTotalSelColAmount(6), "###,##0.00")
    lbl_LackingAmount.Caption = Format(CCur(lbl_CheckAmount.Caption) - CCur(lbl_TotalLiqAmount.Caption), "###,##0.00")
    
Else
    Call ClearCmb
    txt_RDNo.Text = ""
    Call SetGrid
End If
opnvoucher.Close
Set opnvoucher = Nothing
Exit Sub
bad:
MsgBox "Invalid Entry.", vbInformation, "System Message"
End Sub
Private Function ExistBoth(ByVal RecordID As String) As String
Dim opnTable1 As New ADODB.Recordset
Dim existNtable1, existNtable2 As Boolean

opnTable1.Open "Select * from  tblCMS_CDLiquidationRefundOR where recid='" & RecordID & "'", opndbaseFMIS, adOpenStatic, adLockOptimistic
If opnTable1.RecordCount <> 0 Then
    existNtable1 = True
Else
    existNtable1 = False
End If
opnTable1.Close
Set opnTable1 = Nothing


opnTable1.Open "Select * from   tblCMS_CDLiquiditionRefundForOverCA where recid='" & RecordID & "'", opndbaseFMIS, adOpenStatic, adLockOptimistic
If opnTable1.RecordCount <> 0 Then
    existNtable2 = True
Else
    existNtable2 = False
End If
opnTable1.Close
Set opnTable1 = Nothing


If existNtable1 = True And existNtable2 = True Then
    ExistBoth = "BothExisting"
ElseIf existNtable1 = True And existNtable2 = False Then
    ExistBoth = "Table1"
ElseIf existNtable1 = False And existNtable2 = True Then
    ExistBoth = "Table2"
ElseIf existNtable1 = False And existNtable2 = False Then
    ExistBoth = "NoneExisting"
End If

End Function

Private Function GetTotalAmtOfReplacement(ByVal FMISNo As String, ByVal REpNo As String) As Currency 'This is for the Amount of OR having replaced for the Over Amount of Check Against the Actual Total Cash Advance (the edited cash advance)
Dim opnRepAmt As New ADODB.Recordset

opnRepAmt.Open "Select sum(ORAmount) as RepAmt from tblCMS_CDLiquiditionRefundForOverCA where RecID='" & FMISNo & "' and RDONo='" & REpNo & "'", opndbaseFMIS, adOpenStatic, adLockOptimistic
GetTotalAmtOfReplacement = IIf(IsNull(opnRepAmt!RepAmt), 0, opnRepAmt!RepAmt)
opnRepAmt.Close
Set opnRepAmt = Nothing

End Function




Private Function GetBackPrevAmtLacking(ByVal RecordID As String, ByVal ReportNo As String, ByVal Scenario As Integer) As Currency
Dim opntable As New ADODB.Recordset


If Scenario = 2 Or Scenario = 3 Then
    opntable.Open "Select sum(ORAmount) as TotalLacking from tblCMS_CDLiquidationRefundOR where RecID='" & RecordID & "' and RDONo='" & ReportNo & "'", opndbaseFMIS, adOpenStatic, adLockOptimistic
ElseIf Scenario = 1 Then
    opntable.Open "Select sum(ORAmount) as TotalLacking from tblCMS_CDLiquiditionRefundForOverCA where RecID='" & RecordID & "' and RDONo='" & ReportNo & "'", opndbaseFMIS, adOpenStatic, adLockOptimistic
End If

If opntable.RecordCount <> 0 Then
    GetBackPrevAmtLacking = opntable!TotalLacking
End If


End Function
Private Function CheckAmtLacking(ByVal RecordID As String, ByVal ReportNo As String, ByVal Scenario As Integer, ByVal LackingAmt As Currency) As Boolean
Dim opntable As New ADODB.Recordset


If Scenario = 2 Or Scenario = 3 Then
    opntable.Open "Select sum(ORAmount) as TotalLacking from tblCMS_CDLiquidationRefundOR where RecID='" & RecordID & "' and RDONo='" & ReportNo & "'", opndbaseFMIS, adOpenStatic, adLockOptimistic
ElseIf Scenario = 1 Then
    opntable.Open "Select sum(ORAmount) as TotalLacking from tblCMS_CDLiquiditionRefundForOverCA where RecID='" & RecordID & "' and RDONo='" & ReportNo & "'", opndbaseFMIS, adOpenStatic, adLockOptimistic
End If

If opntable.RecordCount <> 0 Then
    If opntable!TotalLacking = LackingAmt Then
        CheckAmtLacking = True
    Else
        CheckAmtLacking = False
    End If
Else
    CheckAmtLacking = False
End If


End Function
Private Function VerifyFexist(ByVal RecordID As String, ByVal ReportNo As String, ByVal Scenario As Integer) As Boolean
Dim opntable As New ADODB.Recordset

If Scenario = 2 Or Scenario = 3 Then
    opntable.Open "Select * from tblCMS_CDLiquidationRefundOR where RecID='" & RecordID & "' and RDONo='" & ReportNo & "'", opndbaseFMIS, adOpenStatic, adLockOptimistic
ElseIf Scenario = 1 Then
    opntable.Open "Select * from tblCMS_CDLiquiditionRefundForOverCA where RecID='" & RecordID & "' and RDONo='" & ReportNo & "'", opndbaseFMIS, adOpenStatic, adLockOptimistic
End If

If opntable.RecordCount <> 0 Then
    VerifyFexist = True
Else
    VerifyFexist = False
End If

End Function



Private Sub FindLikeLastName(ByVal RecordID As String)
Dim cc As Integer

For cc = 0 To List2.ListCount - 1
    If UCase(List2.List(cc)) Like UCase(RecordID) & "*" Then
        List2.ListIndex = cc
    End If
Next cc
End Sub

Private Sub MSHFlexGrid1_DblClick()
If Len(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 8)) > 0 Then
    ActiveFormCaller = Me.name
    ForTheGridRowNo = MSHFlexGrid1.Row

    If Len(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 10)) <> 0 Then 'Kung Naa nay JEV No
        'frmJEVNumberingAssignment_New.txt_JEVNO.Text = MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 10)
        With frmJEVNumberingAssignment_New
            .IsSaveAccntng = False
            .whatfield = "DVNO"
             .Uno = MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 8)
             .fundcode = cmb_FundType.ItemData(cmb_FundType.ListIndex)
            .FTYPE = cmb_FundType.Text
            .FmisVoucherno = List2.List(List2.ListIndex)
            .checkno = ListView1.ListItems(1).Text
            .Date_ = ListView1.ListItems(1).ListSubItems(2).Text
            .RCI = ListView1.ListItems(1).ListSubItems(3).Text
            .Ttype = 3
            .txt_DVNo.Text = MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 0)
            .Show vbModal
            Call List2_Click
        End With
    Else
    With frmJEVNumberingAssignment_New
            .IsSaveAccntng = False
            .whatfield = "DVNO"
             .Uno = MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 0)
            .fundcode = cmb_FundType.ItemData(cmb_FundType.ListIndex)
            .FTYPE = cmb_FundType.Text
            .FmisVoucherno = List2.List(List2.ListIndex)
            .checkno = ListView1.ListItems(1).Text
            .Date_ = ListView1.ListItems(1).ListSubItems(2).Text
            .RCI = ListView1.ListItems(1).ListSubItems(3).Text
            .Ttype = 3
            .txt_DVNo.Text = MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 8)
           
            .Show vbModal
            Call List2_Click
    End With
    End If
Else
    MsgBox "There is no Voucher Attachment for this Check!" & Chr(13) & Chr(13) & "Please Select a New..", vbInformation, "System Information"
End If

End Sub


Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
    Case 1 'Print
        
    Case 3 'Close
        Unload Me
End Select
End Sub

Private Sub txt_RecordID_Click()
If Len(Trim(txt_RecordID.Text)) <> 0 Then
    txt_RecordID.SelStart = 0
    txt_RecordID.SelLength = Len(txt_RecordID.Text)
    txt_RecordID.SetFocus
Else
    txt_RecordID.SetFocus
End If
End Sub

Private Sub txt_RecordID_KeyDown(KeyCode As Integer, Shift As Integer)
Dim tmpVal As Long

On Error GoTo handler
If KeyCode = 13 Then
    If Len(Trim(txt_RecordID.Text)) <> 0 Then
        If InStr(txt_RecordID.Text, "FMISNo-") <> 0 Then
            tmpVal = GetIndex4ListBox(List2, txt_RecordID.Text)
            If tmpVal <> 0 Then
                List2.ListIndex = tmpVal
            Else
                MsgBox "Record ID Not Found!", vbInformation, "System Information"
                txt_RecordID.SelStart = 0
                txt_RecordID.SelLength = Len(txt_RecordID.Text)
                txt_RecordID.SetFocus
            End If
        Else
            txt_RecordID.Text = "FMISNo-" & val(txt_RecordID.Text)
            tmpVal = GetIndex4ListBox(List2, txt_RecordID.Text)
            If tmpVal <> 0 Then
                List2.ListIndex = tmpVal
            Else
                MsgBox "Record ID Not Found!", vbInformation, "System Information"
                txt_RecordID.SelStart = 0
                txt_RecordID.SelLength = Len(txt_RecordID.Text)
                txt_RecordID.SetFocus
            End If
        End If
    End If
End If
handler:
If err.Number <> 0 Then
    MsgBox err.description
End If
End Sub

Private Sub txt_RecordID_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
    Case 45, 48 To 57, 13
    Case Else
        KeyAscii = 0
End Select
End Sub
Private Function GetBalance(ByVal RowNo As Integer) As Currency
GetBalance = CCur(MSHFlexGrid1.TextMatrix(RowNo, 5)) - CCur(MSHFlexGrid1.TextMatrix(RowNo, 6))
End Function
Private Function GetTotalSelColAmount(ByVal Colno As Integer) As Currency
Dim cc As Integer

For cc = 1 To MSHFlexGrid1.Rows - 1
    GetTotalSelColAmount = GetTotalSelColAmount + CCur(MSHFlexGrid1.TextMatrix(cc, Colno))
Next cc
End Function

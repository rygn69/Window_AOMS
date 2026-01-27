VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{1693405E-2DC9-4248-B52F-4AC9145DA2AF}#1.0#0"; "WINXPC~1.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frm_BeginBeginAccountcodeSub 
   BackColor       =   &H00CBE1E7&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Beginning Balance Utility"
   ClientHeight    =   9630
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10260
   Icon            =   "frm_BeginAccountcodeSub.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "frm_BeginAccountcodeSub.frx":076A
   ScaleHeight     =   9630
   ScaleWidth      =   10260
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab txtsearch3 
      Height          =   9015
      Left            =   0
      TabIndex        =   4
      Top             =   600
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   15901
      _Version        =   393216
      Tabs            =   4
      Tab             =   2
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Beginning Balance Registry"
      TabPicture(0)   =   "frm_BeginAccountcodeSub.frx":AE19
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Picture1"
      Tab(0).Control(1)=   "cmb_FundType"
      Tab(0).Control(2)=   "Timer1"
      Tab(0).Control(3)=   "txtDebit"
      Tab(0).Control(4)=   "txtCredit"
      Tab(0).Control(5)=   "txtsearch"
      Tab(0).Control(6)=   "optCode"
      Tab(0).Control(7)=   "optName"
      Tab(0).Control(8)=   "DTPicker1"
      Tab(0).Control(9)=   "lvButtons_H1"
      Tab(0).Control(10)=   "LstAccountcode"
      Tab(0).Control(11)=   "itb32x32"
      Tab(0).Control(12)=   "WindowsXPC1"
      Tab(0).Control(13)=   "lvButtons_H6"
      Tab(0).Control(14)=   "lvButtons_H7"
      Tab(0).Control(15)=   "Line1"
      Tab(0).Control(16)=   "Label1"
      Tab(0).Control(17)=   "lblException"
      Tab(0).Control(18)=   "Label3"
      Tab(0).Control(19)=   "Label4"
      Tab(0).Control(20)=   "Label5"
      Tab(0).ControlCount=   21
      TabCaption(1)   =   "Estimated Revenue and Receipts"
      TabPicture(1)   =   "frm_BeginAccountcodeSub.frx":AE35
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Picture2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Cmd_RRR_Fundtype"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "txt_RRR_Credit"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "txt_RRR_search"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Opt_RRR_Code"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Opt_RRR_Name"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "DTP_RRR_Year"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "lvl_RRR_Cancel"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "ListView1"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "ImageList1"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "WindowsXPC2"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "lvButtons_H9"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "lvl_RRR_Import"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "Line2"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "Label9"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "Label7"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "Label6"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).ControlCount=   17
      TabCaption(2)   =   "SIE"
      TabPicture(2)   =   "frm_BeginAccountcodeSub.frx":AE51
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Label8"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Label10"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Label11"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Line3"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "Label12"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "lvButtons_H12"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "lvButtons_H11"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "lvButtons_H10"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "ListView2"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).Control(9)=   "lvButtons_H8"
      Tab(2).Control(9).Enabled=   0   'False
      Tab(2).Control(10)=   "DTPicker13"
      Tab(2).Control(10).Enabled=   0   'False
      Tab(2).Control(11)=   "Picture3"
      Tab(2).Control(11).Enabled=   0   'False
      Tab(2).Control(12)=   "cmb_FundType3"
      Tab(2).Control(12).Enabled=   0   'False
      Tab(2).Control(13)=   "Text2"
      Tab(2).Control(13).Enabled=   0   'False
      Tab(2).Control(14)=   "Text3"
      Tab(2).Control(14).Enabled=   0   'False
      Tab(2).Control(15)=   "optcode3"
      Tab(2).Control(15).Enabled=   0   'False
      Tab(2).Control(16)=   "optName3"
      Tab(2).Control(16).Enabled=   0   'False
      Tab(2).ControlCount=   17
      TabCaption(3)   =   "Cash Flow"
      TabPicture(3)   =   "frm_BeginAccountcodeSub.frx":AE6D
      Tab(3).ControlEnabled=   0   'False
      Tab(3).ControlCount=   0
      Begin VB.OptionButton optName3 
         Caption         =   "Name"
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
         Left            =   9240
         TabIndex        =   48
         Top             =   420
         Value           =   -1  'True
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.OptionButton optcode3 
         Caption         =   "Code"
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
         Left            =   8400
         TabIndex        =   47
         Top             =   420
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
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
         Left            =   7800
         TabIndex        =   46
         Top             =   870
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
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
         Left            =   6000
         TabIndex        =   45
         Top             =   8460
         Width           =   2775
      End
      Begin VB.ComboBox cmb_FundType3 
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
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   44
         Top             =   900
         Width           =   3375
      End
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   6855
         Left            =   120
         ScaleHeight     =   6825
         ScaleWidth      =   9945
         TabIndex        =   41
         Top             =   1380
         Width           =   9975
         Begin VB.TextBox txt_SIE_Entry 
            Appearance      =   0  'Flat
            BackColor       =   &H0000FFFF&
            BorderStyle     =   0  'None
            ForeColor       =   &H00000000&
            Height          =   405
            Left            =   6720
            TabIndex        =   42
            Top             =   2160
            Visible         =   0   'False
            Width           =   1545
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid2 
            Height          =   6855
            Left            =   0
            TabIndex        =   43
            Top             =   0
            Width           =   9975
            _ExtentX        =   17595
            _ExtentY        =   12091
            _Version        =   393216
            ScrollTrack     =   -1  'True
            AllowUserResizing=   1
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
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   6855
         Left            =   -74880
         ScaleHeight     =   6825
         ScaleWidth      =   9945
         TabIndex        =   31
         Top             =   1380
         Width           =   9975
         Begin VB.TextBox txt_RRR_Entry 
            Appearance      =   0  'Flat
            BackColor       =   &H0000FFFF&
            BorderStyle     =   0  'None
            ForeColor       =   &H00000000&
            Height          =   405
            Left            =   5520
            TabIndex        =   32
            Top             =   1800
            Visible         =   0   'False
            Width           =   1545
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSH_RRR_Grid 
            Height          =   6855
            Left            =   0
            TabIndex        =   33
            Top             =   0
            Width           =   9975
            _ExtentX        =   17595
            _ExtentY        =   12091
            _Version        =   393216
            ScrollTrack     =   -1  'True
            AllowUserResizing=   1
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
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
      End
      Begin VB.ComboBox Cmd_RRR_Fundtype 
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
         Left            =   -73800
         Style           =   2  'Dropdown List
         TabIndex        =   30
         Top             =   900
         Width           =   3375
      End
      Begin VB.TextBox txt_RRR_Credit 
         Alignment       =   1  'Right Justify
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
         Left            =   -69000
         TabIndex        =   29
         Top             =   8460
         Width           =   2775
      End
      Begin VB.TextBox txt_RRR_search 
         Alignment       =   1  'Right Justify
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
         Left            =   -68040
         TabIndex        =   28
         Top             =   870
         Width           =   3135
      End
      Begin VB.OptionButton Opt_RRR_Code 
         BackColor       =   &H80000016&
         Caption         =   "Code"
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
         Left            =   -66840
         TabIndex        =   27
         Top             =   540
         Width           =   855
      End
      Begin VB.OptionButton Opt_RRR_Name 
         BackColor       =   &H80000016&
         Caption         =   "Name"
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
         Left            =   -65880
         TabIndex        =   26
         Top             =   540
         Width           =   975
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   6855
         Left            =   -74880
         ScaleHeight     =   6825
         ScaleWidth      =   9945
         TabIndex        =   12
         Top             =   1260
         Width           =   9975
         Begin VB.TextBox txt_entry 
            Appearance      =   0  'Flat
            BackColor       =   &H0000FFFF&
            BorderStyle     =   0  'None
            ForeColor       =   &H00000000&
            Height          =   405
            Left            =   6720
            TabIndex        =   13
            Top             =   2160
            Visible         =   0   'False
            Width           =   1545
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
            Height          =   6855
            Left            =   0
            TabIndex        =   14
            Top             =   0
            Width           =   9975
            _ExtentX        =   17595
            _ExtentY        =   12091
            _Version        =   393216
            ScrollTrack     =   -1  'True
            AllowUserResizing=   1
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
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
      End
      Begin VB.ComboBox cmb_FundType 
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
         Left            =   -73800
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   780
         Width           =   3375
      End
      Begin VB.Timer Timer1 
         Interval        =   300
         Left            =   -75000
         Top             =   300
      End
      Begin VB.TextBox txtDebit 
         Alignment       =   1  'Right Justify
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
         Left            =   -72720
         TabIndex        =   10
         Top             =   8340
         Width           =   2775
      End
      Begin VB.TextBox txtCredit 
         Alignment       =   1  'Right Justify
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
         Left            =   -69000
         TabIndex        =   9
         Top             =   8340
         Width           =   2775
      End
      Begin VB.TextBox txtsearch 
         Alignment       =   1  'Right Justify
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
         Left            =   -68040
         TabIndex        =   8
         Top             =   750
         Width           =   3135
      End
      Begin VB.OptionButton optCode 
         BackColor       =   &H80000016&
         Caption         =   "Code"
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
         Left            =   -66840
         TabIndex        =   7
         Top             =   420
         Width           =   855
      End
      Begin VB.OptionButton optName 
         BackColor       =   &H80000016&
         Caption         =   "Name"
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
         Left            =   -65880
         TabIndex        =   6
         Top             =   420
         Width           =   975
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   -70320
         TabIndex        =   5
         Top             =   780
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
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
         CustomFormat    =   "yyyy"
         Format          =   184614915
         UpDown          =   -1  'True
         CurrentDate     =   40976
      End
      Begin lvButton.lvButtons_H lvButtons_H1 
         Height          =   495
         Left            =   -66120
         TabIndex        =   15
         Top             =   8340
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
         Caption         =   "&Cancel"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         CapStyle        =   1
         Mode            =   0
         Value           =   0   'False
         Image           =   "frm_BeginAccountcodeSub.frx":AE89
         ImgSize         =   24
         cBack           =   -2147483633
      End
      Begin MSComctlLib.ListView LstAccountcode 
         Height          =   5775
         Left            =   -74880
         TabIndex        =   16
         Top             =   2340
         Width           =   9975
         _ExtentX        =   17595
         _ExtentY        =   10186
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
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Accountcode"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Accountname"
            Object.Width           =   12347
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "ID"
            Object.Width           =   0
         EndProperty
      End
      Begin MSComctlLib.ImageList itb32x32 
         Left            =   -65520
         Top             =   2700
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
               Picture         =   "frm_BeginAccountcodeSub.frx":E993
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm_BeginAccountcodeSub.frx":10325
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm_BeginAccountcodeSub.frx":11CB7
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm_BeginAccountcodeSub.frx":13649
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm_BeginAccountcodeSub.frx":14FDB
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm_BeginAccountcodeSub.frx":1696D
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm_BeginAccountcodeSub.frx":182FF
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm_BeginAccountcodeSub.frx":19C91
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm_BeginAccountcodeSub.frx":1B623
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm_BeginAccountcodeSub.frx":1CFB7
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm_BeginAccountcodeSub.frx":1DC93
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm_BeginAccountcodeSub.frx":1E573
               Key             =   ""
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm_BeginAccountcodeSub.frx":1F24F
               Key             =   ""
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm_BeginAccountcodeSub.frx":1FF2B
               Key             =   ""
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm_BeginAccountcodeSub.frx":20C07
               Key             =   ""
            EndProperty
            BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm_BeginAccountcodeSub.frx":218E3
               Key             =   ""
            EndProperty
            BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm_BeginAccountcodeSub.frx":225BF
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin WinXPC_Engine.WindowsXPC WindowsXPC1 
         Left            =   -68760
         Top             =   1260
         _ExtentX        =   6588
         _ExtentY        =   1085
         ColorScheme     =   2
         EngineStarted   =   -1  'True
         Common_Dialog   =   0   'False
      End
      Begin lvButton.lvButtons_H lvButtons_H6 
         Height          =   375
         Left            =   -69120
         TabIndex        =   17
         Top             =   780
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         Caption         =   "&Load"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFore           =   0
         cFHover         =   33023
         cBhover         =   8438015
         LockHover       =   3
         cGradient       =   33023
         Gradient        =   3
         CapStyle        =   1
         Mode            =   0
         Value           =   0   'False
         Image           =   "frm_BeginAccountcodeSub.frx":22E9B
         cBack           =   16777215
      End
      Begin lvButton.lvButtons_H lvButtons_H7 
         Height          =   375
         Left            =   -74880
         TabIndex        =   18
         Top             =   8340
         Visible         =   0   'False
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         Caption         =   "&Import"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFore           =   0
         cFHover         =   33023
         cBhover         =   8438015
         LockHover       =   3
         cGradient       =   33023
         Gradient        =   3
         CapStyle        =   1
         Mode            =   0
         Value           =   0   'False
         Image           =   "frm_BeginAccountcodeSub.frx":269A5
         cBack           =   16777215
      End
      Begin MSComCtl2.DTPicker DTP_RRR_Year 
         Height          =   375
         Left            =   -70320
         TabIndex        =   25
         Top             =   900
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
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
         CustomFormat    =   "yyyy"
         Format          =   184614915
         UpDown          =   -1  'True
         CurrentDate     =   40976
      End
      Begin lvButton.lvButtons_H lvl_RRR_Cancel 
         Height          =   495
         Left            =   -66120
         TabIndex        =   34
         Top             =   8460
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
         Caption         =   "&Cancel"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         CapStyle        =   1
         Mode            =   0
         Value           =   0   'False
         Image           =   "frm_BeginAccountcodeSub.frx":2A4AF
         ImgSize         =   24
         cBack           =   -2147483633
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   5775
         Left            =   -74880
         TabIndex        =   35
         Top             =   2460
         Width           =   9975
         _ExtentX        =   17595
         _ExtentY        =   10186
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
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Accountcode"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Accountname"
            Object.Width           =   12347
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "ID"
            Object.Width           =   0
         EndProperty
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   -65520
         Top             =   2460
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
               Picture         =   "frm_BeginAccountcodeSub.frx":2DFB9
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm_BeginAccountcodeSub.frx":2F94B
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm_BeginAccountcodeSub.frx":312DD
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm_BeginAccountcodeSub.frx":32C6F
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm_BeginAccountcodeSub.frx":34601
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm_BeginAccountcodeSub.frx":35F93
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm_BeginAccountcodeSub.frx":37925
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm_BeginAccountcodeSub.frx":392B7
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm_BeginAccountcodeSub.frx":3AC49
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm_BeginAccountcodeSub.frx":3C5DD
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm_BeginAccountcodeSub.frx":3D2B9
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm_BeginAccountcodeSub.frx":3DB99
               Key             =   ""
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm_BeginAccountcodeSub.frx":3E875
               Key             =   ""
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm_BeginAccountcodeSub.frx":3F551
               Key             =   ""
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm_BeginAccountcodeSub.frx":4022D
               Key             =   ""
            EndProperty
            BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm_BeginAccountcodeSub.frx":40F09
               Key             =   ""
            EndProperty
            BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm_BeginAccountcodeSub.frx":41BE5
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin WinXPC_Engine.WindowsXPC WindowsXPC2 
         Left            =   -68760
         Top             =   1380
         _ExtentX        =   6588
         _ExtentY        =   1085
         ColorScheme     =   2
         EngineStarted   =   -1  'True
         Common_Dialog   =   0   'False
      End
      Begin lvButton.lvButtons_H lvButtons_H9 
         Height          =   375
         Left            =   -69120
         TabIndex        =   36
         Top             =   900
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         Caption         =   "&Load"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFore           =   0
         cFHover         =   33023
         cBhover         =   8438015
         LockHover       =   3
         cGradient       =   33023
         Gradient        =   3
         CapStyle        =   1
         Mode            =   0
         Value           =   0   'False
         Image           =   "frm_BeginAccountcodeSub.frx":424C1
         cBack           =   16777215
      End
      Begin lvButton.lvButtons_H lvl_RRR_Import 
         Height          =   375
         Left            =   -74880
         TabIndex        =   37
         Top             =   8460
         Visible         =   0   'False
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         Caption         =   "&Import"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFore           =   0
         cFHover         =   33023
         cBhover         =   8438015
         LockHover       =   3
         cGradient       =   33023
         Gradient        =   3
         CapStyle        =   1
         Mode            =   0
         Value           =   0   'False
         Image           =   "frm_BeginAccountcodeSub.frx":45FCB
         cBack           =   16777215
      End
      Begin MSComCtl2.DTPicker DTPicker13 
         Height          =   375
         Left            =   4680
         TabIndex        =   49
         Top             =   900
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
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
         CustomFormat    =   "yyyy"
         Format          =   184614915
         UpDown          =   -1  'True
         CurrentDate     =   40976
      End
      Begin lvButton.lvButtons_H lvButtons_H8 
         Height          =   495
         Left            =   8880
         TabIndex        =   50
         Top             =   8460
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
         Caption         =   "&Cancel"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         CapStyle        =   1
         Mode            =   0
         Value           =   0   'False
         Image           =   "frm_BeginAccountcodeSub.frx":49AD5
         ImgSize         =   24
         cBack           =   -2147483633
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   5775
         Left            =   120
         TabIndex        =   51
         Top             =   2460
         Width           =   9975
         _ExtentX        =   17595
         _ExtentY        =   10186
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
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Accountcode"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Accountname"
            Object.Width           =   12347
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "ID"
            Object.Width           =   0
         EndProperty
      End
      Begin lvButton.lvButtons_H lvButtons_H10 
         Height          =   375
         Left            =   5680
         TabIndex        =   52
         Top             =   900
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         Caption         =   "&Load"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFore           =   0
         cFHover         =   33023
         cBhover         =   8438015
         LockHover       =   3
         cGradient       =   33023
         Gradient        =   3
         CapStyle        =   1
         Mode            =   0
         Value           =   0   'False
         Image           =   "frm_BeginAccountcodeSub.frx":4D5DF
         cBack           =   16777215
      End
      Begin lvButton.lvButtons_H lvButtons_H11 
         Height          =   375
         Left            =   120
         TabIndex        =   53
         Top             =   8460
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         Caption         =   "&Import"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFore           =   0
         cFHover         =   33023
         cBhover         =   8438015
         LockHover       =   3
         cGradient       =   33023
         Gradient        =   3
         CapStyle        =   1
         Mode            =   0
         Value           =   0   'False
         Image           =   "frm_BeginAccountcodeSub.frx":510E9
         cBack           =   16777215
      End
      Begin lvButton.lvButtons_H lvButtons_H12 
         Height          =   375
         Left            =   6720
         TabIndex        =   58
         Top             =   900
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         Caption         =   "&Import"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFore           =   0
         cFHover         =   33023
         cBhover         =   8438015
         LockHover       =   3
         cGradient       =   33023
         Gradient        =   3
         CapStyle        =   1
         Mode            =   0
         Value           =   0   'False
         Image           =   "frm_BeginAccountcodeSub.frx":54BF3
         cBack           =   16777215
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Total:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5160
         TabIndex        =   57
         Top             =   8490
         Width           =   855
      End
      Begin VB.Line Line3 
         X1              =   120
         X2              =   10080
         Y1              =   8340
         Y2              =   8340
      End
      Begin VB.Label Label11 
         Caption         =   "Beginnig Balance for Statement of Income and Expenses"
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
         Left            =   120
         TabIndex        =   56
         Top             =   420
         Width           =   6975
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Search By:"
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
         Left            =   7320
         TabIndex        =   55
         Top             =   420
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Fund Type:"
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
         TabIndex        =   54
         Top             =   900
         Width           =   1335
      End
      Begin VB.Line Line2 
         X1              =   -74880
         X2              =   -64920
         Y1              =   8340
         Y2              =   8340
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Fund Type:"
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
         Left            =   -74880
         TabIndex        =   40
         Top             =   900
         Width           =   1335
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Total:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -69840
         TabIndex        =   39
         Top             =   8490
         Width           =   855
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Search By:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -68040
         TabIndex        =   38
         Top             =   540
         Width           =   1455
      End
      Begin VB.Line Line1 
         X1              =   -74880
         X2              =   -64920
         Y1              =   8220
         Y2              =   8220
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Fund Type:"
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
         Left            =   -74880
         TabIndex        =   23
         Top             =   780
         Width           =   1335
      End
      Begin VB.Label lblException 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   -74880
         TabIndex        =   22
         Top             =   660
         Width           =   60
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Debit:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -73560
         TabIndex        =   21
         Top             =   8370
         Width           =   855
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Credit:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -69840
         TabIndex        =   20
         Top             =   8370
         Width           =   855
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Search By:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -68040
         TabIndex        =   19
         Top             =   420
         Width           =   1455
      End
   End
   Begin lvButton.lvButtons_H lvButtons_H2 
      Height          =   495
      Left            =   4440
      TabIndex        =   0
      Top             =   10080
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      Caption         =   "&Export"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      CapStyle        =   1
      Mode            =   0
      Value           =   0   'False
      Image           =   "frm_BeginAccountcodeSub.frx":554CD
      ImgSize         =   24
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H lvButtons_H3 
      Height          =   495
      Left            =   5880
      TabIndex        =   1
      Top             =   10080
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      Caption         =   "&Add"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      CapStyle        =   1
      Mode            =   0
      Value           =   0   'False
      Image           =   "frm_BeginAccountcodeSub.frx":58FD7
      ImgSize         =   24
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H lvButtons_H4 
      Height          =   495
      Left            =   7320
      TabIndex        =   2
      Top             =   10080
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      Caption         =   "&Generate"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      CapStyle        =   1
      Mode            =   0
      Value           =   0   'False
      Image           =   "frm_BeginAccountcodeSub.frx":5CAE1
      ImgSize         =   24
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H lvButtons_H5 
      Height          =   495
      Left            =   480
      TabIndex        =   3
      Top             =   10200
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      Caption         =   "&Export"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      CapStyle        =   1
      Mode            =   0
      Value           =   0   'False
      Image           =   "frm_BeginAccountcodeSub.frx":605EB
      ImgSize         =   24
      cBack           =   -2147483633
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Define criteria prior and click load button to display the details."
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
      Left            =   120
      TabIndex        =   24
      Top             =   120
      Width           =   5475
   End
End
Attribute VB_Name = "frm_BeginBeginAccountcodeSub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub cmb_FundType_Change()
MSHFlexGrid1.Clear
End Sub

Private Sub cmb_FundType_Click()
MSHFlexGrid1.Cols = 6
MSHFlexGrid1.Rows = 2
        MSHFlexGrid1.TextMatrix(0, 1) = "Code"
        MSHFlexGrid1.TextMatrix(0, 2) = "Explanation"
        MSHFlexGrid1.TextMatrix(0, 3) = "Debit"
        MSHFlexGrid1.TextMatrix(0, 4) = "Credit"
        
        MSHFlexGrid1.ColWidth(0) = 0
        MSHFlexGrid1.ColWidth(1) = 700
        MSHFlexGrid1.ColWidth(2) = 5800
        MSHFlexGrid1.ColWidth(3) = 1500
        MSHFlexGrid1.ColWidth(4) = 1500
        MSHFlexGrid1.ColWidth(5) = 0
        MSHFlexGrid1.TextMatrix(1, 1) = ""
        MSHFlexGrid1.TextMatrix(1, 2) = ""
        MSHFlexGrid1.TextMatrix(1, 3) = ""
        MSHFlexGrid1.TextMatrix(1, 4) = ""
End Sub
Private Sub Cmd_RRR_Fundtype_Change()
Call lvButtons_H9_Click
End Sub
Private Sub Cmd_RRR_Fundtype_Click()
Call lvButtons_H9_Click
End Sub
Private Sub Form_Load()
Call LoadFundType(Cmd_RRR_Fundtype)
Call LoadFundType(cmb_FundType)
Call LoadFundType(cmb_FundType3)
End Sub
Public Function GetAccountNamebyorder(ByVal lst As ListView, ByVal Condition As String)
'Dim rec As New ADODB.Recordset
'Dim x
'Dim z As Integer
''Condition = Replace(Condition, "'", "")
'rec.Open "Select Accountcode,Accountname from tblREF_AIS_ChartOfAccountsMother order by " & Condition & "", opndbaseFMIS, adOpenStatic, adLockOptimistic
'    lst.ListItems.Clear
'    If rec.RecordCount > 0 Then
'        Set MSHFlexGrid1.DataSource = rec
'    End If
'    Call SetGrid
'rec.Close
'Set rec = Nothing

End Function
Private Sub GetSum()
On Error Resume Next
Dim Debit As Currency
Dim Credit As Currency
Dim x As Long
For x = 1 To MSHFlexGrid1.Rows - 1
    If Trim(MSHFlexGrid1.TextMatrix(x, 3)) <> "" Then
    Debit = CDbl(Debit) + CDbl(MSHFlexGrid1.TextMatrix(x, 3))
    End If
    If Trim(MSHFlexGrid1.TextMatrix(x, 4)) <> "" Then
    Credit = CDbl(Credit) + CDbl(MSHFlexGrid1.TextMatrix(x, 4))
    End If
    DoEvents
Next x
txtDebit.Text = Format(Debit, "#,##0.00")
txtCredit.Text = Format(Credit, "#,##0.00")
End Sub
Private Sub GetSumRRR()
On Error Resume Next
Dim Debit As Currency
Dim Credit As Currency
Dim x As Long
For x = 1 To MSH_RRR_Grid.Rows - 1
    If Trim(MSH_RRR_Grid.TextMatrix(x, 3)) <> "" Then
    Debit = CDbl(Debit) + CDbl(MSH_RRR_Grid.TextMatrix(x, 3))
    End If
    DoEvents
Next x
txt_RRR_Credit.Text = Format(Debit, "#,##0.00")
End Sub
Private Sub GetSumSIE()
'On Error Resume Next
Dim Debit As Currency
Dim Credit As Currency
Dim x As Long
For x = 1 To MSHFlexGrid2.Rows - 1
    If Trim(MSHFlexGrid2.TextMatrix(x, 3)) <> "" Then
    Debit = CDbl(Debit) + CDbl(MSHFlexGrid2.TextMatrix(x, 3))
    End If
    DoEvents
Next x
Text2.Text = Format(Debit, "#,##0.00")
End Sub
Public Sub SetGrid()
Dim x As Long

        MSHFlexGrid1.Cols = 6
        MSHFlexGrid1.TextMatrix(0, 1) = "Code"
        MSHFlexGrid1.TextMatrix(0, 2) = "Explanation"
        MSHFlexGrid1.TextMatrix(0, 3) = "Debit"
        MSHFlexGrid1.TextMatrix(0, 4) = "Credit"
        
        MSHFlexGrid1.ColWidth(0) = 0
        MSHFlexGrid1.ColWidth(1) = 700
        MSHFlexGrid1.ColWidth(2) = 5800
        MSHFlexGrid1.ColWidth(3) = 1500
        MSHFlexGrid1.ColWidth(4) = 1500
        MSHFlexGrid1.ColWidth(5) = 0
For x = 1 To MSHFlexGrid1.Rows - 1
    If MSHFlexGrid1.TextMatrix(x, 3) <> "" And Val(MSHFlexGrid1.TextMatrix(x, 3)) <> 0 Then
    MSHFlexGrid1.TextMatrix(x, 3) = Format(MSHFlexGrid1.TextMatrix(x, 3), "#,###.00")
    
    Else
    MSHFlexGrid1.TextMatrix(x, 3) = ""
    End If
    
    If MSHFlexGrid1.TextMatrix(x, 4) <> "" And Val(MSHFlexGrid1.TextMatrix(x, 4)) <> 0 Then
    MSHFlexGrid1.TextMatrix(x, 4) = Format(MSHFlexGrid1.TextMatrix(x, 4), "#,###.00")
    Else
    MSHFlexGrid1.TextMatrix(x, 4) = ""
    End If
    'MSHFlexGrid1.TextMatrix(x, 3) = IIf((MSHFlexGrid1.TextMatrix(x, 3)) = "", "", Format(MSHFlexGrid1.TextMatrix(x, 3), "#,###.00"))
    'MSHFlexGrid1.TextMatrix(x, 4) = IIf((MSHFlexGrid1.TextMatrix(x, 4)) = "", "", Format(MSHFlexGrid1.TextMatrix(x, 4), "#,###.00"))
Next x
End Sub

Public Sub SetGridRRR()
Dim x As Long

        MSH_RRR_Grid.Cols = 5
        MSH_RRR_Grid.TextMatrix(0, 1) = "Code"
        MSH_RRR_Grid.TextMatrix(0, 2) = "Explanation"
        MSH_RRR_Grid.TextMatrix(0, 3) = "Estimated Amount"
        
        MSH_RRR_Grid.ColWidth(0) = 0
        MSH_RRR_Grid.ColWidth(1) = 700
        MSH_RRR_Grid.ColWidth(2) = 6300
        MSH_RRR_Grid.ColWidth(3) = 2500
        
        MSH_RRR_Grid.ColWidth(4) = 0
For x = 1 To MSH_RRR_Grid.Rows - 1
    If MSH_RRR_Grid.TextMatrix(x, 3) <> "" And Val(MSH_RRR_Grid.TextMatrix(x, 3)) <> 0 Then
    MSH_RRR_Grid.TextMatrix(x, 3) = Format(MSH_RRR_Grid.TextMatrix(x, 3), "#,###.00")
    
    Else
    MSH_RRR_Grid.TextMatrix(x, 3) = ""
    End If
Next x
End Sub

Public Sub SetGridSIE()
Dim x As Long

        MSHFlexGrid2.Cols = 5
        MSHFlexGrid2.TextMatrix(0, 1) = "Code"
        MSHFlexGrid2.TextMatrix(0, 2) = "Explanation"
        MSHFlexGrid2.TextMatrix(0, 3) = "Estimated Amount"
        
        MSHFlexGrid2.ColWidth(0) = 0
        MSHFlexGrid2.ColWidth(1) = 700
        MSHFlexGrid2.ColWidth(2) = 6300
        MSHFlexGrid2.ColWidth(3) = 2500
        
        MSHFlexGrid2.ColWidth(4) = 0
For x = 1 To MSHFlexGrid2.Rows - 1
    If MSHFlexGrid2.TextMatrix(x, 3) <> "" And Val(MSHFlexGrid2.TextMatrix(x, 3)) <> 0 Then
    MSHFlexGrid2.TextMatrix(x, 3) = Format(MSHFlexGrid2.TextMatrix(x, 3), "#,###.00")
    
    Else
    MSHFlexGrid2.TextMatrix(x, 3) = ""
    End If
Next x
End Sub

Private Sub LstAccountcode_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
If ColumnHeader = "Accountname" Then
Call GetAccountNamebyorder(LstAccountcode, "Accountname")
ElseIf ColumnHeader = "Accountcode" Then
Call GetAccountNamebyorder(LstAccountcode, "Accountcode")
End If
End Sub

Private Sub LstAccountcode_DblClick()
With frm_BeginAccountcodeSub1
.Address = Trim(LstAccountcode.SelectedItem.SubItems(1))
.col = 2
.accountcode = Trim(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 1))
.Subcode1 = Trim(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 1))
.Subdesc1 = Trim(LstAccountcode.SelectedItem.SubItems(1))
.Condition = "Subcode1 ='" & Trim(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 1)) & "'"
.Show 1
End With
End Sub
Private Sub lvButtons_H1_Click()
Unload Me
End Sub

Private Sub lvButtons_H10_Click()
'On Error GoTo bad
Dim rec As New ADODB.Recordset
Dim x
Dim z As Integer
Set rec = opndbaseFMIS.Execute("SELECT  accountcode,[AccountName],Amount FROM [fmis].[dbo].[MPfunc_SIE_AccountcodeView](" & cmb_FundType3.ItemData(cmb_FundType3.ListIndex) & "," & DTPicker13.Year & ")  order by accountcode")
   If rec.RecordCount > 0 Then
        Set MSHFlexGrid2.DataSource = rec
    Else
    MSHFlexGrid2.Clear
    MSHFlexGrid2.FixedRows = 1
    MSHFlexGrid2.Rows = 2
   End If
   Call SetGridSIE
    Call GetSumSIE
rec.Close
Set rec = Nothing
Exit Sub
bad:
MsgBox "Noted" & err.description
End Sub

Private Sub lvButtons_H11_Click()
With frm_SIE_Import
    .fundcode = cmb_FundType3.ItemData(cmb_FundType3.ListIndex)
    .YEAR_ = DTPicker13.Year
.Show 1
Call lvButtons_H10_Click
End With
End Sub

Private Sub lvButtons_H12_Click()
frm_SIE_ImportFromSIE.Show 1
End Sub

Private Sub lvButtons_H4_Click()
Dim x As Integer
For x = 1 To MSHFlexGrid1.Rows - 1
opndbaseFMIS.Execute "Insert into tblReff_CodeClassification(subcode1,subdesc1) values('" & LstAccountcode.ListItems(x).Text & "','" & Replace(LstAccountcode.ListItems(x).ListSubItems(1).Text, "'", "") & "') "
Next x
MsgBox "Successfully save", vbInformation, "System Message"
End Sub

Private Sub lvButtons_H6_Click()
On Error GoTo bad
Dim rec As New ADODB.Recordset
Dim x
Dim z As Integer
'Set rec = opndbaseFMIS.Execute("Select left(subcode1,3) as Accountcode,subdesc1,Sum(Sdebit),Sum(Scredit),max(lvl) from vw_MP_BeginnigBalance where (fundcode = " & cmb_FundType.ItemData(cmb_FundType.ListIndex) & " or fundcode is null) and  (actioncode = 0 or actioncode is null) group by left(subcode1,3),subdesc1 order by left(subcode1,3)")
Set rec = opndbaseFMIS.Execute("SELECT  accountcode,[AccountName],sum([Sumdebit]) as debit,sum([Sumcredit]) as credit FROM [fmis].[dbo].[vw_MP_BeginBal]where year_ = '" & DTPicker1.Year & "' and (fundcode = " & cmb_FundType.ItemData(cmb_FundType.ListIndex) & ")  group by accountcode,Accountname order by accountcode")
   If rec.RecordCount > 0 Then
        Set MSHFlexGrid1.DataSource = rec
    Else
    MSHFlexGrid1.Clear
    MSHFlexGrid1.FixedRows = 1
    MSHFlexGrid1.Rows = 2
   End If
   Call SetGrid
    Call GetSum
rec.Close
Set rec = Nothing
Exit Sub
bad:
MsgBox "Noted" & err.description
End Sub

Private Sub lvButtons_H7_Click()
If cmb_FundType.Text <> "" And DTPicker1.Year <> "" Then
medll.centerme frm_AccntsPayImport
frm_AccntsPayImport.fundcode = cmb_FundType.ItemData(cmb_FundType.ListIndex)
frm_AccntsPayImport.YEAR_ = DTPicker1.Year
frm_AccntsPayImport.Show 1
End If
End Sub

Private Sub lvButtons_H9_Click()
On Error GoTo bad
Dim rec As New ADODB.Recordset
Dim x
Dim z As Integer
'Set rec = opndbaseFMIS.Execute("Select left(subcode1,3) as Accountcode,subdesc1,Sum(Sdebit),Sum(Scredit),max(lvl) from vw_MP_BeginnigBalance where (fundcode = " & cmb_FundType.ItemData(cmb_FundType.ListIndex) & " or fundcode is null) and  (actioncode = 0 or actioncode is null) group by left(subcode1,3),subdesc1 order by left(subcode1,3)")
Set rec = opndbaseFMIS.Execute("SELECT  Code,[SourceName],EstimatedAmount FROM [fmis].[dbo].[vw_MP_RRR_FullEstimatedDesc] where year_ = '" & DTP_RRR_Year.Year & "' and (fundcode = " & Cmd_RRR_Fundtype.ItemData(Cmd_RRR_Fundtype.ListIndex) & ") order by Sourcename")
   If rec.RecordCount > 0 Then
        Set MSH_RRR_Grid.DataSource = rec
    Else
    MSH_RRR_Grid.Clear
    MSH_RRR_Grid.FixedRows = 1
    MSH_RRR_Grid.Rows = 2
   End If
   Call SetGridRRR
    Call GetSumRRR
rec.Close
Set rec = Nothing
Exit Sub
bad:
MsgBox "Noted" & err.description
End Sub

Private Sub MSH_RRR_Grid_Click()
On Error GoTo bad

        Select Case MSH_RRR_Grid.col
        Case 3 To 4 'Debit/Credit
            txt_RRR_Entry.Move MSH_RRR_Grid.CellLeft, MSH_RRR_Grid.CellTop, MSH_RRR_Grid.CellWidth, MSH_RRR_Grid.CellHeight
            txt_RRR_Entry.Visible = True
            If Len(Trim(MSH_RRR_Grid.Text)) <> 0 Then
                txt_RRR_Entry.Text = MSH_RRR_Grid.Text
                txt_RRR_Entry.SelStart = 0
                txt_RRR_Entry.SelLength = Len(txt_RRR_Entry.Text)
            Else
                txt_RRR_Entry.Text = ""
            End If
            txt_RRR_Entry.SetFocus
        Case Else
            txt_RRR_Entry.Visible = False
        End Select
Exit Sub
bad:
MsgBox err.description
End Sub

Private Sub MSHFlexGrid1_Click()
On Error GoTo bad

        Select Case MSHFlexGrid1.col
        Case 3 To 4 'Debit/Credit
        
        If ExecFunction("SELECT [fmis].[dbo].[Mpfunc_ChckIfHaveSub] ('" & Trim(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 1)) & "',0)") > 1 Then
        txt_entry.Visible = False
        Call MSHFlexGrid1_DblClick
        Else
            txt_entry.Move MSHFlexGrid1.CellLeft, MSHFlexGrid1.CellTop, MSHFlexGrid1.CellWidth, MSHFlexGrid1.CellHeight
            txt_entry.Visible = True
            If Len(Trim(MSHFlexGrid1.Text)) <> 0 Then
                txt_entry.Text = MSHFlexGrid1.Text
                txt_entry.SelStart = 0
                txt_entry.SelLength = Len(txt_entry.Text)
            Else
                txt_entry.Text = ""
            End If
            txt_entry.SetFocus
        End If
        
        Case Else
            txt_entry.Visible = False
        End Select
Exit Sub
bad:
MsgBox err.description
End Sub

Private Sub MSHFlexGrid1_DblClick()
If Trim(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 1)) <> "" Then
    With frm_BeginAccountcodeSub1
    .Address = Trim(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 2))
    .col = 2
    .accountcode = Trim(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 1))
    .Subcode1 = Trim(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 1))
    .Subdesc1 = Trim(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 2))
    '.Condition = "Subcode1 ='" & Trim(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 1)) & "' and  subcode2 is not null and (fundcode = " & cmb_FundType.ItemData(cmb_FundType.ListIndex) & " or fundcode is null)"
    .Condition = "Exec [MPproc_LoadJEVfromBegenning] @fundcode = '" & Trim(cmb_FundType.ItemData(cmb_FundType.ListIndex)) & "',@Accountcode = '" & Trim(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 1)) & "',@lvl =2,@year = '" & DTPicker1.Year & "'"
    .Caption = Trim(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 1))
    .fundcode = cmb_FundType.ItemData(cmb_FundType.ListIndex)
    .YEAR_ = DTPicker1.Year
    .Show 1
    Call lvButtons_H6_Click
    End With
End If
End Sub

Private Sub MSHFlexGrid2_Click()
On Error GoTo bad
Select Case MSHFlexGrid2.col
Case 3 To 4 'Debit/Credit
    txt_SIE_Entry.Move MSHFlexGrid2.CellLeft, MSHFlexGrid2.CellTop, MSHFlexGrid2.CellWidth, MSHFlexGrid2.CellHeight
    txt_SIE_Entry.Visible = True
    If Len(Trim(MSHFlexGrid2.Text)) <> 0 Then
        txt_SIE_Entry.Text = MSHFlexGrid2.Text
        txt_SIE_Entry.SelStart = 0
        txt_SIE_Entry.SelLength = Len(txt_SIE_Entry.Text)
    Else
        txt_SIE_Entry.Text = ""
    End If
    txt_SIE_Entry.SetFocus
Case Else
    txt_SIE_Entry.Visible = False
End Select
Exit Sub
bad:
MsgBox err.description
End Sub

Private Sub Timer1_Timer()
DoEvents
End Sub

Private Sub txt_entry_KeyPress(KeyAscii As Integer)
 On Error GoTo bad
    If KeyAscii = 13 Then
            If IsNumeric(txt_entry.Text) = False And txt_entry.Text <> "" Then
                MsgBox "None Numeric Entry, Please Check Your Entry", vbCritical, "System Message"
                Exit Sub
            End If
            MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, MSHFlexGrid1.col) = Format((txt_entry.Text), "#,##0.00")
                If MSHFlexGrid1.col = 3 Then
                    If Trim(txt_entry.Text) <> "" Then
                        MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 4) = ""
                    Else
                        MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 3) = ""
                    End If
                
                ElseIf MSHFlexGrid1.col <> 5 Then
                    
                    If Trim(txt_entry.Text) <> "" Then
                        MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 3) = ""
                    Else
                        MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 4) = ""
                    End If
                End If
                txt_entry.Visible = False
                If MSHFlexGrid1.col = 5 Then
                    If txt_entry.Text = "1" Or txt_entry.Text = "5" Then
                    MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, MSHFlexGrid1.col) = txt_entry.Text
                    Else
                    MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, MSHFlexGrid1.col) = "1"
                    End If
                End If
                Call SaveAmount(IIf((MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 3)) = "", 0, MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 3)), IIf((MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 4)) = "", 0, MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 4)))
                Call GetSum
    End If
Exit Sub
bad:
    Call LoadErr(err.Number, err.Source & ", " & Me.name & ", " & Me.Caption, err.description)
End Sub
Public Function SaveAmount(ByVal Debit As Currency, ByVal Credit As Currency)
Dim rec As New ADODB.Recordset

rec.Open "Select Accountcode from tblAMIS_Begeningbalance where accountcode = '" & Trim(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 1)) & "' and fundcode = '" & cmb_FundType.ItemData(cmb_FundType.ListIndex) & "' AND YEAR_ = '" & DTPicker1.Year & "'", opndbaseFMIS, adOpenStatic
    If rec.RecordCount > 0 Then
        opndbaseFMIS.Execute "Update tblAMIS_Begeningbalance set debit = '" & Debit & "',credit = '" & Credit & "' where accountcode = '" & Trim(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 1)) & "' and fundcode = '" & cmb_FundType.ItemData(cmb_FundType.ListIndex) & "' AND YEAR_ = '" & DTPicker1.Year & "' "
    Else
        opndbaseFMIS.Execute "Insert into tblAMIS_Begeningbalance (accountcode,debit,Credit,fundcode,actioncode,YEAR_) values ('" & Trim(Me.Caption) & "-" & Trim(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 1)) & "','" & Debit & "','" & Credit & "','" & cmb_FundType.ItemData(cmb_FundType.ListIndex) & "',1,'" & DTPicker1.Year & "')"
    End If
rec.Close
End Function
Public Function SaveAmountRRR(ByVal EstimatedAmount As Currency)
Dim rec As New ADODB.Recordset

rec.Open "Select rrrtrnno from [tblAMIS_RRREstimated] where rrrtrnno = '" & Trim(MSH_RRR_Grid.TextMatrix(MSH_RRR_Grid.Row, 1)) & "' and fundcode = '" & Cmd_RRR_Fundtype.ItemData(Cmd_RRR_Fundtype.ListIndex) & "' AND YEAR_ = '" & DTP_RRR_Year.Year & "'", opndbaseFMIS, adOpenStatic
    If rec.RecordCount > 0 Then
        opndbaseFMIS.Execute "Update [tblAMIS_RRREstimated] set EstimatedAmount = '" & EstimatedAmount & "' where rrrtrnno = '" & Trim(MSH_RRR_Grid.TextMatrix(MSH_RRR_Grid.Row, 1)) & "' and fundcode = '" & Cmd_RRR_Fundtype.ItemData(Cmd_RRR_Fundtype.ListIndex) & "' AND YEAR_ = '" & DTP_RRR_Year.Year & "' "
    Else
        opndbaseFMIS.Execute "Insert into [tblAMIS_RRREstimated] (rrrtrnno,[EstimatedAmount],fundcode,actioncode,YEAR_) values ('" & Trim(MSH_RRR_Grid.TextMatrix(MSH_RRR_Grid.Row, 1)) & "','" & EstimatedAmount & "','" & Cmd_RRR_Fundtype.ItemData(Cmd_RRR_Fundtype.ListIndex) & "',1,'" & DTP_RRR_Year.Year & "')"
    End If
rec.Close
End Function
Public Function SaveAmountSIE(ByVal amount As Currency)
Dim rec As New ADODB.Recordset
rec.Open "Select trnno from [tblAMIS_BegeningbalanceSIE] where Accountcode = '" & Trim(MSHFlexGrid2.TextMatrix(MSHFlexGrid2.Row, 1)) & "' and fundcode = '" & cmb_FundType3.ItemData(cmb_FundType3.ListIndex) & "' AND YEAR_ = '" & DTPicker13.Year & "'", opndbaseFMIS, adOpenStatic
    If rec.RecordCount > 0 Then
        opndbaseFMIS.Execute "Update [tblAMIS_BegeningbalanceSIE] set Amount = '" & amount & "' where Accountcode = '" & Trim(MSHFlexGrid2.TextMatrix(MSHFlexGrid2.Row, 1)) & "' and fundcode = '" & cmb_FundType3.ItemData(cmb_FundType3.ListIndex) & "' AND YEAR_ = '" & DTPicker13.Year & "' "
    Else
        opndbaseFMIS.Execute "Insert into [tblAMIS_BegeningbalanceSIE] (Accountcode,[Amount],fundcode,actioncode,YEAR_) values ('" & Trim(MSHFlexGrid2.TextMatrix(MSHFlexGrid2.Row, 1)) & "','" & amount & "','" & cmb_FundType3.ItemData(cmb_FundType3.ListIndex) & "',1,'" & DTPicker13.Year & "')"
    End If
rec.Close
End Function
Private Sub txt_RRR_Entry_KeyPress(KeyAscii As Integer)
 On Error GoTo bad
    If KeyAscii = 13 Then
            If IsNumeric(txt_RRR_Entry.Text) = False And txt_RRR_Entry.Text <> "" Then
                MsgBox "None Numeric Entry, Please Check Your Entry", vbCritical, "System Message"
                Exit Sub
            End If
            MSH_RRR_Grid.TextMatrix(MSH_RRR_Grid.Row, MSH_RRR_Grid.col) = Format((txt_RRR_Entry.Text), "#,##0.00")
                If MSH_RRR_Grid.col = 3 Then
                    If Trim(txt_RRR_Entry.Text) <> "" Then
                        MSH_RRR_Grid.TextMatrix(MSH_RRR_Grid.Row, 4) = ""
                    Else
                        MSH_RRR_Grid.TextMatrix(MSH_RRR_Grid.Row, 3) = ""
                    End If
                
'                ElseIf MSH_RRR_Grid.col <> 5 Then
'
'                    If Trim(txt_RRR_Entry.Text) <> "" Then
'                        MSH_RRR_Grid.TextMatrix(MSH_RRR_Grid.Row, 3) = ""
'                    Else
'                        MSH_RRR_Grid.TextMatrix(MSH_RRR_Grid.Row, 4) = ""
'                    End If
                End If
                txt_RRR_Entry.Visible = False
                Call SaveAmountRRR(IIf((MSH_RRR_Grid.TextMatrix(MSH_RRR_Grid.Row, 3)) = "", 0, MSH_RRR_Grid.TextMatrix(MSH_RRR_Grid.Row, 3)))
                Call GetSumRRR
    End If
Exit Sub
bad:
    Call LoadErr(err.Number, err.Source & ", " & Me.name & ", " & Me.Caption, err.description)
End Sub
Private Sub txt_SIE_Entry_KeyPress(KeyAscii As Integer)
 On Error GoTo bad
    If KeyAscii = 13 Then
            If IsNumeric(txt_SIE_Entry.Text) = False And txt_SIE_Entry.Text <> "" Then
                MsgBox "None Numeric Entry, Please Check Your Entry", vbCritical, "System Message"
                Exit Sub
            End If
            MSHFlexGrid2.TextMatrix(MSHFlexGrid2.Row, MSHFlexGrid2.col) = Format((txt_SIE_Entry.Text), "#,##0.00")
                If MSHFlexGrid2.col = 3 Then
                    If Trim(txt_SIE_Entry.Text) <> "" Then
                        MSHFlexGrid2.TextMatrix(MSHFlexGrid2.Row, 4) = ""
                    Else
                        MSHFlexGrid2.TextMatrix(MSHFlexGrid2.Row, 3) = ""
                    End If
                
'                ElseIf MSHFlexGrid2.col <> 5 Then
'
'                    If Trim(txt_SIE_Entry.Text) <> "" Then
'                        MSHFlexGrid2.TextMatrix(MSHFlexGrid2.Row, 3) = ""
'                    Else
'                        MSHFlexGrid2.TextMatrix(MSHFlexGrid2.Row, 4) = ""
'                    End If
                End If
                txt_SIE_Entry.Visible = False
                Call SaveAmountSIE(IIf((MSHFlexGrid2.TextMatrix(MSHFlexGrid2.Row, 3)) = "", 0, MSHFlexGrid2.TextMatrix(MSHFlexGrid2.Row, 3)))
                Call GetSumSIE
    End If
Exit Sub
bad:
    Call LoadErr(err.Number, err.Source & ", " & Me.name & ", " & Me.Caption, err.description)
End Sub

Private Sub txt_RRR_search_KeyPress(KeyAscii As Integer)
On Error GoTo bad
If KeyAscii = 13 Then
Dim rec As New ADODB.Recordset
Dim x
Dim z As Integer
Dim sql As String
    If Opt_RRR_Code.Value = True Then
       'sql = "Select left(subcode1,3) as Accountcode,subdesc1,Sum(Sdebit),Sum(Scredit),max(lvl) from vw_MP_BeginnigBalance where (fundcode = " & cmb_FundType.ItemData(cmb_FundType.ListIndex) & " or fundcode is null) and  left(subcode1,3) like '" & Trim(txtSearch.Text) & "%' and (actioncode = 0 or actioncode is null) group by left(subcode1,3),subdesc1 order by Subdesc1"
       sql = "SELECT  code,[Sourcename],EstimatedAmount FROM [fmis].[dbo].[vw_MP_RRR_FullEstimatedDesc] where year_ = '" & DTP_RRR_Year.Year & "' and  (fundcode = " & Cmd_RRR_Fundtype.ItemData(Cmd_RRR_Fundtype.ListIndex) & " or fundcode is null) AND code  like '" & Trim(txt_RRR_search.Text) & "%'  order by Sourcename"
    Else
        'sql = "Select left(subcode1,3) as Accountcode,subdesc1,Sum(Sdebit),Sum(Scredit),max(lvl) from vw_MP_BeginnigBalance where (fundcode = " & cmb_FundType.ItemData(cmb_FundType.ListIndex) & " or fundcode is null) and subdesc1 like '%" & Trim(txtSearch.Text) & "%' and (actioncode = 0 or actioncode is null) group by left(subcode1,3),subdesc1 order by Subdesc1"
        sql = "SELECT  code,[Sourcename],EstimatedAmount FROM [fmis].[dbo].[vw_MP_RRR_FullEstimatedDesc]where year_ = '" & DTP_RRR_Year.Year & "' and  (fundcode = " & Cmd_RRR_Fundtype.ItemData(Cmd_RRR_Fundtype.ListIndex) & " or fundcode is null) AND sOURCEname  like '" & Trim(txt_RRR_search.Text) & "%' order by Sourcename"
    End If
    
    Set rec = opndbaseFMIS.Execute(sql)
        If rec.RecordCount > 0 Then
            Set MSH_RRR_Grid.DataSource = rec
            Call SetGridRRR
        End If
        Call GetSumRRR
    rec.Close
Set rec = Nothing
End If
Exit Sub
bad:
MsgBox "Noted: " & err.description
End Sub

Private Sub txtSearch_KeyPress(KeyAscii As Integer)
On Error GoTo bad
If KeyAscii = 13 Then
Dim rec As New ADODB.Recordset
Dim x
Dim z As Integer
Dim sql As String
    If optCode.Value = True Then
       'sql = "Select left(subcode1,3) as Accountcode,subdesc1,Sum(Sdebit),Sum(Scredit),max(lvl) from vw_MP_BeginnigBalance where (fundcode = " & cmb_FundType.ItemData(cmb_FundType.ListIndex) & " or fundcode is null) and  left(subcode1,3) like '" & Trim(txtSearch.Text) & "%' and (actioncode = 0 or actioncode is null) group by left(subcode1,3),subdesc1 order by Subdesc1"
       sql = "SELECT  accountcode,[AccountName],sum([Sumdebit]) as debit,sum([Sumcredit]) as credit FROM [fmis].[dbo].[vw_MP_BeginBal] where year_ = '" & DTPicker1.Year & "' and  (fundcode = " & cmb_FundType.ItemData(cmb_FundType.ListIndex) & " or fundcode is null) AND accountcode  like '" & Trim(txtsearch.Text) & "%' group by accountcode,Accountname order by accountcode"
    Else
        'sql = "Select left(subcode1,3) as Accountcode,subdesc1,Sum(Sdebit),Sum(Scredit),max(lvl) from vw_MP_BeginnigBalance where (fundcode = " & cmb_FundType.ItemData(cmb_FundType.ListIndex) & " or fundcode is null) and subdesc1 like '%" & Trim(txtSearch.Text) & "%' and (actioncode = 0 or actioncode is null) group by left(subcode1,3),subdesc1 order by Subdesc1"
        sql = "SELECT  accountcode,[AccountName],sum([Sumdebit]) as debit,sum([Sumcredit]) as credit FROM [fmis].[dbo].[vw_MP_BeginBal]where year_ = '" & DTPicker1.Year & "' and  (fundcode = " & cmb_FundType.ItemData(cmb_FundType.ListIndex) & " or fundcode is null) AND accountname  like '" & Trim(txtsearch.Text) & "%' group by accountcode,Accountname order by accountcode"
    End If
    
    Set rec = opndbaseFMIS.Execute(sql)
        If rec.RecordCount > 0 Then
            Set MSHFlexGrid1.DataSource = rec
            Call SetGrid
        End If
        Call GetSum
    rec.Close
Set rec = Nothing
End If
Exit Sub
bad:
MsgBox "Noted: " & err.description
End Sub


VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{1693405E-2DC9-4248-B52F-4AC9145DA2AF}#1.0#0"; "WINXPC~1.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MShflxgd.ocx"
Begin VB.Form frm_GeneralJournalJevNumbering 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "JEV Numbering for General Journal through DVNO"
   ClientHeight    =   9450
   ClientLeft      =   660
   ClientTop       =   1140
   ClientWidth     =   14925
   Icon            =   "frm_GeneralJournalJevNumbering.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9450
   ScaleWidth      =   14925
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
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
      Left            =   11805
      Locked          =   -1  'True
      TabIndex        =   36
      Top             =   8880
      Width           =   3060
   End
   Begin VB.TextBox txtDVNo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   135
      TabIndex        =   34
      Top             =   1335
      Width           =   4845
   End
   Begin VB.TextBox txtAlobs 
      Appearance      =   0  'Flat
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
      Left            =   5160
      Locked          =   -1  'True
      TabIndex        =   29
      Top             =   1320
      Width           =   3660
   End
   Begin VB.TextBox txtParticular 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Left            =   5160
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   28
      Top             =   2280
      Width           =   9690
   End
   Begin VB.TextBox txtAmount 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
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
      Left            =   11760
      Locked          =   -1  'True
      TabIndex        =   27
      Top             =   1320
      Width           =   3060
   End
   Begin VB.TextBox txtFund 
      Appearance      =   0  'Flat
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
      Left            =   9000
      Locked          =   -1  'True
      TabIndex        =   26
      Top             =   1320
      Width           =   2580
   End
   Begin VB.TextBox txt_RecordID 
      Appearance      =   0  'Flat
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
      Height          =   540
      Left            =   120
      TabIndex        =   20
      ToolTipText     =   "Type only CN No. then press Enter"
      Top             =   2280
      Width           =   3765
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&ADD"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   3960
      Picture         =   "frm_GeneralJournalJevNumbering.frx":0E42
      TabIndex        =   3
      Top             =   2280
      Width           =   1005
   End
   Begin WinXPC_Engine.WindowsXPC WindowsXPC1 
      Left            =   9360
      Top             =   9720
      _ExtentX        =   6588
      _ExtentY        =   1085
      ColorScheme     =   2
      Common_Dialog   =   0   'False
      PictureControl  =   0   'False
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
            Picture         =   "frm_GeneralJournalJevNumbering.frx":493C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_GeneralJournalJevNumbering.frx":59BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_GeneralJournalJevNumbering.frx":7AF8
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_GeneralJournalJevNumbering.frx":8D7A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_GeneralJournalJevNumbering.frx":BC04
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   840
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   14925
      _ExtentX        =   26326
      _ExtentY        =   1482
      ButtonWidth     =   1323
      ButtonHeight    =   1429
      Appearance      =   1
      Style           =   1
      ImageList       =   "itb32x32"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Delete"
            ImageIndex      =   4
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
         Left            =   7440
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
               Picture         =   "frm_GeneralJournalJevNumbering.frx":11966
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm_GeneralJournalJevNumbering.frx":132F8
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm_GeneralJournalJevNumbering.frx":14C8A
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm_GeneralJournalJevNumbering.frx":1661C
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm_GeneralJournalJevNumbering.frx":17FAE
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm_GeneralJournalJevNumbering.frx":19940
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm_GeneralJournalJevNumbering.frx":1B2D2
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm_GeneralJournalJevNumbering.frx":1CC64
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm_GeneralJournalJevNumbering.frx":1E5F6
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm_GeneralJournalJevNumbering.frx":1FF8A
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm_GeneralJournalJevNumbering.frx":20C66
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm_GeneralJournalJevNumbering.frx":21546
               Key             =   ""
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm_GeneralJournalJevNumbering.frx":22222
               Key             =   ""
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm_GeneralJournalJevNumbering.frx":22EFE
               Key             =   ""
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm_GeneralJournalJevNumbering.frx":23BDA
               Key             =   ""
            EndProperty
            BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm_GeneralJournalJevNumbering.frx":248B6
               Key             =   ""
            EndProperty
            BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm_GeneralJournalJevNumbering.frx":25592
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComCtl2.Animation Animation1 
         Height          =   450
         Left            =   11760
         TabIndex        =   19
         Top             =   120
         Width           =   480
         _ExtentX        =   847
         _ExtentY        =   794
         _Version        =   393216
         FullWidth       =   32
         FullHeight      =   30
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grd_details 
      Height          =   5010
      Left            =   120
      TabIndex        =   25
      Top             =   3720
      Width           =   14700
      _ExtentX        =   25929
      _ExtentY        =   8837
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
   Begin VB.Frame Frame5 
      BackColor       =   &H00000000&
      Height          =   1410
      Left            =   6120
      TabIndex        =   2
      Top             =   3840
      Visible         =   0   'False
      Width           =   3675
      Begin VB.CommandButton cmd_post 
         Caption         =   "Post (JEV No.)"
         Height          =   1005
         Left            =   1800
         Picture         =   "frm_GeneralJournalJevNumbering.frx":25E6E
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   240
         Width           =   1755
      End
      Begin VB.CommandButton cmd_Mass 
         Caption         =   "Mass JEV Nos."
         Height          =   1005
         Left            =   120
         Picture         =   "frm_GeneralJournalJevNumbering.frx":29968
         Style           =   1  'Graphical
         TabIndex        =   17
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
         TabIndex        =   6
         Top             =   -1320
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
            TabIndex        =   15
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
            TabIndex        =   8
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
            TabIndex        =   7
            Top             =   750
            Width           =   2565
         End
         Begin MSComCtl2.DTPicker DTPCNdate 
            CausesValidation=   0   'False
            Height          =   390
            Left            =   5340
            TabIndex        =   9
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
            Format          =   173211649
            UpDown          =   -1  'True
            CurrentDate     =   38240
         End
         Begin MSComCtl2.DTPicker DTPRdate 
            CausesValidation=   0   'False
            Height          =   390
            Left            =   5340
            TabIndex        =   10
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
            Format          =   173277185
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
            TabIndex        =   16
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
            TabIndex        =   14
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
            TabIndex        =   13
            Top             =   315
            Width           =   1170
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Account Name"
            Height          =   195
            Left            =   3600
            TabIndex        =   12
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
            TabIndex        =   11
            Top             =   795
            Width           =   1725
         End
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
   Begin MSComCtl2.DTPicker DTPicker1 
      CausesValidation=   0   'False
      Height          =   360
      Left            =   11640
      TabIndex        =   23
      Top             =   4560
      Visible         =   0   'False
      Width           =   3045
      _ExtentX        =   5371
      _ExtentY        =   635
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
      Format          =   173277187
      UpDown          =   -1  'True
      CurrentDate     =   38240
   End
   Begin VB.ComboBox cmb_FundType 
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
      Left            =   11640
      TabIndex        =   21
      Top             =   3960
      Visible         =   0   'False
      Width           =   3060
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000012&
      Caption         =   "Note: Double Click the list below to post the transaction"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   120
      TabIndex        =   38
      Top             =   3120
      Width           =   5655
   End
   Begin VB.Label Label17 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Amount:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   285
      Left            =   10200
      TabIndex        =   37
      Top             =   9000
      Width           =   1560
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1. Enter PTV  Number and press ENTER:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   120
      TabIndex        =   35
      Top             =   960
      Width           =   4350
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Report No:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   285
      Left            =   5160
      TabIndex        =   33
      Top             =   960
      Width           =   1170
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Particular:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   285
      Left            =   5160
      TabIndex        =   32
      Top             =   1920
      Width           =   1080
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Amount:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   285
      Left            =   12570
      TabIndex        =   31
      Top             =   960
      Width           =   945
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fund Type:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   285
      Left            =   8970
      TabIndex        =   30
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Transactiondate:"
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
      Height          =   255
      Left            =   9840
      TabIndex        =   24
      Top             =   4605
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Special Accounts:"
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
      Height          =   255
      Left            =   9840
      TabIndex        =   22
      Top             =   4005
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label14"
      Height          =   195
      Left            =   13755
      TabIndex        =   5
      Top             =   9135
      Width           =   570
   End
   Begin VB.Label Label13 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "2. Enter DVNO and click ADD button"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   90
      TabIndex        =   4
      Top             =   1920
      Width           =   4845
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label8"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   105
      TabIndex        =   1
      Top             =   9060
      Width           =   480
   End
End
Attribute VB_Name = "frm_GeneralJournalJevNumbering"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub LoadLOCbySQL()
Dim rec As New ADODB.Recordset
Dim Frec As New ADODB.Recordset
Dim jevno As String
Dim x
If checkIfAlreadyInTheLIst(txt_RecordID.Text) = True Then
    MsgBox "Transaction Already on the List..", vbCriticalm, "System Information"
    Exit Sub
End If
If txtFund.Text = "" Then
    MsgBox "Invalid PTVno...", vbCriticalm, "System Information"
    Exit Sub
End If
    Set rec = opndbaseFMIS.Execute("Select ObrNo,Name, Particular,OfficeMedium, [Gross Amount],PTV as JEVNO,[FundCode],[FundName],DVNO from [MPfunc_GetTranThroughDVNO] ('" & txt_RecordID.Text & "')")
    If rec.RecordCount > 0 Then
        With grd_details
           x = grd_details.Rows - 1
            .TextMatrix(x, 0) = rec!obrno
            .TextMatrix(x, 1) = rec!name
            .TextMatrix(x, 2) = rec!Particular
            .TextMatrix(x, 3) = rec!OfficeMedium
            .TextMatrix(x, 4) = rec![Gross Amount]
            .TextMatrix(x, 5) = getJEVNO(txt_RecordID.Text)
            .TextMatrix(x, 6) = rec!fundcode
            .TextMatrix(x, 7) = rec!FundName
            .TextMatrix(x, 8) = rec!dvno
            .Rows = grd_details.Rows + 1
        End With
    End If
    rec.Close
Call GEtTotal
End Sub
Function getJEVNO(ByVal dvno) As String
Dim rec As New ADODB.Recordset
getJEVNO = ""
Set Frec = opndbaseFMIS.Execute("SELECT JEVNO FROM [fmis].[dbo].[tblAMIS_FinalJEV] where actioncode = 1 and Dvno = '" & txt_RecordID.Text & "'")
If Frec.RecordCount > 0 Then
getJEVNO = Trim(Frec!jevno)
End If
Frec.Close
Set Frec = Nothing
End Function

Private Sub Command1_Click()
Animation1.Visible = True
Animation1.Open App.path & AViLocation & "\horizontaloading.avi"
Animation1.Play
    Call LoadLOCbySQL
Animation1.Stop
Animation1.Close
Animation1.Visible = False
End Sub
Public Sub GEtTotal()
Dim x As Long
Dim cur As Currency
For x = 1 To grd_details.Rows - 1
    If Trim(grd_details.TextMatrix(x, 4)) <> "" Then
        cur = cur + CCur(grd_details.TextMatrix(x, 4))
    End If
Next x
Text1.Text = Format(cur, "#,##0.00")
End Sub
Function checkIfAlreadyInTheLIst(dvno As String) As Boolean
Dim x As Long
For x = 1 To grd_details.Rows - 1
    If grd_details.TextMatrix(x, 8) = dvno Then
        checkIfAlreadyInTheLIst = True
        Exit For
    End If
Next x
End Function
Private Sub Form_Load()
Call LoadFundType(cmb_FundType)
DTPicker1.Value = Now
Call SetGrid
End Sub
Private Sub LoadFund()
Dim Frec As New ADODB.Recordset
Dim x As Integer
cmb_FundType.Clear
Frec.Open ("Select * from tblRefBMS_Funds Order By FundMedium"), opndbaseFMIS, adOpenStatic, adLockOptimistic
If Frec.RecordCount > 0 Then
    For x = 1 To Frec.RecordCount
        cmb_FundType.AddItem Frec!fundmedium
        cmb_FundType.ItemData(cmb_FundType.NewIndex) = Frec!fundcode
        Frec.MoveNext
    Next x
End If
Frec.Close
Set Frec = Nothing

End Sub
Public Sub SetGrid()
With grd_details
    .Cols = 9
    .Rows = 2
    .FixedRows = 1
    .TextMatrix(0, 0) = "OBRNO"
    .TextMatrix(0, 1) = "Name"
    .TextMatrix(0, 2) = "Particular"
    .TextMatrix(0, 3) = "RC"
    .TextMatrix(0, 4) = "Amount"
    .TextMatrix(0, 5) = "JEVNO"
    .TextMatrix(0, 6) = "OBRNO"
    .TextMatrix(0, 7) = "Fundcode"
    .TextMatrix(0, 8) = "Fundtype"
    .ColWidth(0) = 1500 'OBRNO
    .ColWidth(1) = 2000 'Name
    .ColWidth(2) = 5500 'Particular
    .ColWidth(3) = 2000 'office medium
    .ColWidth(4) = 1500 'Amount
    .ColWidth(5) = 1500 'JEVNO
    .ColWidth(6) = 0    'Fundcode
    .ColWidth(7) = 0    'Fundtype
    .ColWidth(8) = 0    'dvno
End With
End Sub



Private Sub grd_details_DblClick()
On Error GoTo bad
If Trim(grd_details.TextMatrix(grd_details.Row, 5)) <> "" Then
    MsgBox "This transation is Already Posted..", vbInformation, "System Message"
    Exit Sub
End If
If Trim(grd_details.TextMatrix(grd_details.Row, 6)) <> "" Then
ActiveFormCaller = Me.name
    ForTheGridRowNo = grd_details.Row
        With frmJEVNumberingAssignment_New
        .IsSaveAccntng = False
        .ptv = txtDVNo.Text
        .ptvNo = txtDVNo.Text
        .Ttype = 4
        .fundcode = grd_details.TextMatrix(grd_details.Row, 6)
        .FundType = grd_details.TextMatrix(grd_details.Row, 7)
        .FTYPE = grd_details.TextMatrix(grd_details.Row, 7)
        .Transtype = 4
        .Gamount = Trim(grd_details.TextMatrix(grd_details.Row, 4))
        .whatfield = "DVNO"
        .txt_DVNo.Text = grd_details.TextMatrix(grd_details.Row, 8)
        .txt_DVNo.Locked = True
        .Show 1
        getJEVNO (txt_RecordID.Text)
            

'            .Date_ = grd_details.TextMatrix(grd_details.Row, 2)
'            .RCI = List1.List(List1.ListIndex)
'            .checkno = grd_details.TextMatrix(grd_details.Row, 3)
'            .Particular = grd_details.TextMatrix(grd_details.Row, 8)
'            .FTYPE = cmb_FundType.Text
'            .ptvNo = grd_details.TextMatrix(grd_details.Row, 5)
'            .whatfield = "DVNO"
'            .Uno = grd_details.TextMatrix(grd_details.Row, 0)
'            .Show 1
        
        End With
Else
    'MsgBox "NO PTV no...Please Enter PTVNO first..", vbCritical, "System Message"
End If
Exit Sub
bad:
Call LoadErr(err.Number, Me.name, err.description)
End Sub


Private Sub grd_details_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDelete Then
    If grd_details.Rows = 2 Then
    With grd_details
        .TextMatrix(grd_details.Row, 0) = ""
        .TextMatrix(grd_details.Row, 1) = ""
        .TextMatrix(grd_details.Row, 2) = ""
        .TextMatrix(grd_details.Row, 3) = ""
        .TextMatrix(grd_details.Row, 4) = ""
        .TextMatrix(grd_details.Row, 5) = ""
        .TextMatrix(grd_details.Row, 6) = ""
        .TextMatrix(grd_details.Row, 7) = ""
        .TextMatrix(grd_details.Row, 8) = ""
    End With
    Else
        grd_details.RemoveItem (grd_details.Row)
    End If
    Call GEtTotal
End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
 Select Case Button:
    Case "&Delete":
        If MsgBox("Are you sure do you want to remove selected transation?", vbInformation + vbYesNo, "System Confirmation") = vbYes Then
        grd_details.RemoveItem (grd_details.Row)
        End If
    Case "&Close"
        Unload Me
    End Select
End Sub

Private Sub TxtDvno_Change()
    txtAlobs.Text = ""
    txtParticular.Text = ""
    txtFund.Text = ""
    txtAmount.Text = ""
End Sub

Private Sub txtDVNo_KeyPress(KeyAscii As Integer)
On Error GoTo bad
Dim DVRec As New ADODB.Recordset
If KeyAscii = 13 Then
    Set DVRec = opndbaseFMIS.Execute("Select * FRom tblCMS_CDCheckBook where DVNo='" & txtDVNo.Text & "' and (ActionCode=1)")
        If DVRec.RecordCount > 0 Then
            txtAlobs.Text = Trim(DVRec!chknumber)
            txtParticular.Text = Trim(DVRec!Particular)
            txtFund.Text = GetSFNameByCode(Left(txtDVNo.Text, 3))
            txtAmount.Text = Format(DVRec!amount, "#,##0.00")
        Else
            MsgBox "No record Found", vbCritical, "System Information"
            txtAlobs.Text = ""
            txtParticular.Text = ""
            txtFund.Text = ""
            txtAmount.Text = ""
        End If
    DVRec.Close
    Set DVRec = Nothing
End If
Exit Sub
bad:
MsgBox err.description
End Sub

VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{1693405E-2DC9-4248-B52F-4AC9145DA2AF}#1.0#0"; "WINXPC~1.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmCDCashReceiptsJevNumbering 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "JEV Numbering for Cash Receipts and General Journal through PTV number"
   ClientHeight    =   10125
   ClientLeft      =   720
   ClientTop       =   1200
   ClientWidth     =   14505
   Icon            =   "frmCDCashReceiptsJevNumbering.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   10125
   ScaleWidth      =   14505
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   840
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   14505
      _ExtentX        =   25585
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
            Description     =   "7"
            ImageIndex      =   7
         EndProperty
      EndProperty
      Begin MSComCtl2.Animation Animation1 
         Height          =   450
         Left            =   11880
         TabIndex        =   1
         Top             =   120
         Width           =   480
         _ExtentX        =   847
         _ExtentY        =   794
         _Version        =   393216
         FullWidth       =   32
         FullHeight      =   30
      End
      Begin MSComctlLib.ImageList itb32x32 
         Left            =   13200
         Top             =   360
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
               Picture         =   "frmCDCashReceiptsJevNumbering.frx":0E42
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCDCashReceiptsJevNumbering.frx":27D4
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCDCashReceiptsJevNumbering.frx":4166
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCDCashReceiptsJevNumbering.frx":5AF8
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCDCashReceiptsJevNumbering.frx":748A
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCDCashReceiptsJevNumbering.frx":8E1C
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCDCashReceiptsJevNumbering.frx":A7AE
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCDCashReceiptsJevNumbering.frx":C140
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCDCashReceiptsJevNumbering.frx":DAD2
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCDCashReceiptsJevNumbering.frx":F466
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCDCashReceiptsJevNumbering.frx":10142
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCDCashReceiptsJevNumbering.frx":10A22
               Key             =   ""
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCDCashReceiptsJevNumbering.frx":116FE
               Key             =   ""
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCDCashReceiptsJevNumbering.frx":123DA
               Key             =   ""
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCDCashReceiptsJevNumbering.frx":130B6
               Key             =   ""
            EndProperty
            BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCDCashReceiptsJevNumbering.frx":13D92
               Key             =   ""
            EndProperty
            BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCDCashReceiptsJevNumbering.frx":14A6E
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "Note: To Remove the selected item,please check the ckeckbox on the list                and press Del"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1020
         Left            =   5040
         TabIndex        =   44
         Top             =   -360
         Width           =   7980
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   10815
      Left            =   -600
      TabIndex        =   2
      Top             =   0
      Width           =   15855
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         ForeColor       =   &H80000008&
         Height          =   6240
         Left            =   3165
         ScaleHeight     =   6210
         ScaleWidth      =   11820
         TabIndex        =   35
         Top             =   3450
         Width           =   11850
         Begin VB.CheckBox Check1 
            BackColor       =   &H8000000A&
            Caption         =   "Select All"
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
            Left            =   45
            TabIndex        =   37
            Top             =   120
            Width           =   1215
         End
         Begin VB.CheckBox Check2 
            BackColor       =   &H8000000A&
            Caption         =   "Deselect All"
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
            Left            =   1320
            TabIndex        =   36
            Top             =   120
            Width           =   1455
         End
         Begin MSComctlLib.ListView lstDetails 
            Height          =   5655
            Left            =   0
            TabIndex        =   38
            Top             =   480
            Width           =   11775
            _ExtentX        =   20770
            _ExtentY        =   9975
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            Checkboxes      =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   0
            BackColor       =   -2147483643
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
            NumItems        =   16
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "TransNo"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Particular"
               Object.Width           =   5292
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Transactiondate"
               Object.Width           =   2822
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Ref"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "Code"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   5
               Text            =   "Amount"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   6
               Text            =   "Balance Amt"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   7
               Text            =   "ReconcilingSeqNo"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   8
               Text            =   "DVNo"
               Object.Width           =   3528
            EndProperty
            BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   9
               Text            =   "AlreadySaved"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   10
               Text            =   "JEVNo"
               Object.Width           =   4410
            EndProperty
            BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   11
               Text            =   "RDno"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   12
               Text            =   "fundtype"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   13
               Text            =   "bankid"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   14
               Text            =   "banckaccountno"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   15
               Text            =   "accountname"
               Object.Width           =   0
            EndProperty
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "For the Period"
         Height          =   975
         Left            =   690
         TabIndex        =   32
         Top             =   1815
         Width           =   2265
         Begin MSComCtl2.DTPicker DTPicker1 
            CausesValidation=   0   'False
            Height          =   390
            Left            =   90
            TabIndex        =   33
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
            Format          =   112721923
            UpDown          =   -1  'True
            CurrentDate     =   38240
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Label1"
            Height          =   195
            Left            =   1530
            TabIndex        =   34
            Top             =   1125
            Visible         =   0   'False
            Width           =   480
         End
      End
      Begin VB.Frame Frame6 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "PTV Number"
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   690
         TabIndex        =   31
         Top             =   4425
         Width           =   2280
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
         Left            =   705
         TabIndex        =   30
         ToolTipText     =   "Double Click PTV number to Add on JEV numbering list."
         Top             =   4740
         Width           =   2265
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   50
         Left            =   9630
         Top             =   2535
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
         Left            =   690
         TabIndex        =   29
         ToolTipText     =   "Type only PTV number then press ENTER"
         Top             =   3870
         Width           =   2280
      End
      Begin VB.Frame Frame2 
         Caption         =   "Special Account"
         Height          =   780
         Left            =   705
         TabIndex        =   27
         Top             =   930
         Width           =   2250
         Begin VB.ComboBox cmb_FundType 
            Height          =   315
            Left            =   75
            Style           =   2  'Dropdown List
            TabIndex        =   28
            Top             =   300
            Width           =   2100
         End
      End
      Begin VB.Frame Frame5 
         Height          =   2610
         Left            =   3165
         TabIndex        =   4
         Top             =   780
         Width           =   11835
         Begin VB.Frame Frame4 
            Caption         =   "Transaction type"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1095
            Left            =   120
            TabIndex        =   45
            Top             =   1440
            Width           =   7815
            Begin VB.OptionButton Opt_CRJ 
               Caption         =   "Cash Receipts Journal"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Left            =   120
               TabIndex        =   47
               Tag             =   "1"
               Top             =   480
               Width           =   3615
            End
            Begin VB.OptionButton Opt_GJ 
               Caption         =   "General Journal"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Left            =   4080
               TabIndex        =   46
               Tag             =   "4"
               Top             =   480
               Width           =   2895
            End
         End
         Begin VB.Frame Frame3 
            Caption         =   "Source of Fund"
            Height          =   1245
            Left            =   60
            TabIndex        =   8
            Top             =   75
            Width           =   11490
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
               Left            =   4920
               Style           =   2  'Dropdown List
               TabIndex        =   13
               Top             =   1320
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
               Left            =   2010
               TabIndex        =   12
               Top             =   315
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
               Left            =   7770
               TabIndex        =   11
               Top             =   315
               Width           =   3525
            End
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
               Left            =   2010
               TabIndex        =   10
               Top             =   750
               Width           =   3525
            End
            Begin VB.TextBox txt_RDNo 
               Alignment       =   2  'Center
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   15.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   7770
               Locked          =   -1  'True
               TabIndex        =   9
               Top             =   720
               Width           =   3525
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Fund Type"
               Height          =   195
               Left            =   180
               TabIndex        =   18
               Top             =   1575
               Width           =   765
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Drawee Bank"
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
               TabIndex        =   17
               Top             =   435
               Width           =   1440
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Account Name"
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
               Left            =   6060
               TabIndex        =   16
               Top             =   435
               Width           =   1575
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Account Number"
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
               TabIndex        =   15
               Top             =   855
               Width           =   1785
            End
            Begin VB.Label Label15 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Report Number:"
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
               Left            =   6000
               TabIndex        =   14
               Top             =   840
               Width           =   1695
            End
         End
         Begin VB.CommandButton cmd_post 
            Caption         =   "Post (JEV No.)"
            Height          =   1005
            Left            =   9840
            Picture         =   "frmCDCashReceiptsJevNumbering.frx":1534A
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   1440
            Visible         =   0   'False
            Width           =   1755
         End
         Begin VB.CommandButton cmd_Mass 
            Caption         =   "Mass JEV Nos."
            Height          =   1005
            Left            =   8160
            Picture         =   "frmCDCashReceiptsJevNumbering.frx":18E44
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   1440
            Visible         =   0   'False
            Width           =   1635
         End
         Begin MSFlexGridLib.MSFlexGrid MSHFlexGrid1 
            Height          =   960
            Left            =   3840
            TabIndex        =   7
            Top             =   3360
            Width           =   630
            _ExtentX        =   1111
            _ExtentY        =   1693
            _Version        =   393216
            FixedCols       =   0
            ForeColorFixed  =   4210688
            BackColorBkg    =   0
            ScrollTrack     =   -1  'True
            FocusRect       =   0
            SelectionMode   =   1
            AllowUserResizing=   1
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Narrow"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total Lacking Amount :"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   7965
            TabIndex        =   26
            Top             =   870
            Width           =   1650
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total Liquidated Amount :"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   7965
            TabIndex        =   25
            Top             =   420
            Width           =   1815
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total Cash Advance :"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   4380
            TabIndex        =   24
            Top             =   870
            Width           =   1545
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
            TabIndex        =   23
            Top             =   870
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
            TabIndex        =   22
            Top             =   420
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
            TabIndex        =   21
            Top             =   870
            Width           =   1890
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
            TabIndex        =   20
            Top             =   420
            Width           =   1890
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total Check Amount:"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   4380
            TabIndex        =   19
            Top             =   420
            Width           =   1500
         End
         Begin VB.Shape Shape3 
            BackColor       =   &H00000000&
            BackStyle       =   1  'Opaque
            Height          =   1275
            Left            =   8040
            Top             =   1365
            Width           =   3810
         End
         Begin VB.Shape Shape2 
            BackColor       =   &H00000000&
            BackStyle       =   1  'Opaque
            Height          =   2385
            Left            =   -135
            Top             =   -1020
            Width           =   12000
         End
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Load Reports"
         Height          =   435
         Left            =   1830
         TabIndex        =   3
         Top             =   2910
         Width           =   1125
      End
      Begin WinXPC_Engine.WindowsXPC WindowsXPC1 
         Left            =   5400
         Top             =   9240
         _ExtentX        =   6588
         _ExtentY        =   1085
         ColorScheme     =   2
         Common_Dialog   =   0   'False
         PictureControl  =   0   'False
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   1995
         Top             =   9150
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
               Picture         =   "frmCDCashReceiptsJevNumbering.frx":1C93E
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCDCashReceiptsJevNumbering.frx":1D9C0
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCDCashReceiptsJevNumbering.frx":1FAFA
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCDCashReceiptsJevNumbering.frx":20D7C
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCDCashReceiptsJevNumbering.frx":23C06
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00000000&
         BackStyle       =   1  'Opaque
         Height          =   9540
         Left            =   600
         Top             =   720
         Width           =   2445
      End
      Begin VB.Label Label7 
         Caption         =   "Label7"
         Height          =   210
         Left            =   4170
         TabIndex        =   43
         Top             =   3195
         Visible         =   0   'False
         Width           =   1395
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label8"
         ForeColor       =   &H8000000E&
         Height          =   195
         Left            =   735
         TabIndex        =   42
         Top             =   9195
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Search PTV No.:"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   690
         TabIndex        =   41
         Top             =   3600
         Width           =   1215
      End
      Begin VB.Label Label13 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "Label13"
         ForeColor       =   &H0000FF00&
         Height          =   420
         Left            =   720
         TabIndex        =   40
         Top             =   2910
         Width           =   1020
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label14"
         Height          =   195
         Left            =   14385
         TabIndex        =   39
         Top             =   9750
         Width           =   570
      End
   End
End
Attribute VB_Name = "frmCDCashReceiptsJevNumbering"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim tmpAccName As String
Dim FMISNo As String


Private Sub Check1_Click()
If Check1.Value = 1 Then
CkeckOption ("select")
Check2.Value = 0
End If
End Sub

Private Sub Check2_Click()
If Check2.Value = 1 Then
CkeckOption ("deselect")
Check1.Value = 0
End If
End Sub


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



Private Sub cmb_FundType_Click()
cmb_accnumber.Text = ""
cmb_AccountName.Text = ""
cmb_bank.Text = ""
lstDetails.ListItems.Clear
txt_RDNo.Text = ""
List2.Clear
End Sub

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
Private Sub JEVMassNumbering(ByVal FundType As String)
Dim opnJEV As New ADODB.Recordset
Dim rec As New ADODB.Recordset
Dim sql As String
Dim cc As Integer
Dim dvno As String
Dim LastJEVSNno As Long
Dim TYP As Integer

If Opt_CRJ.Value = True Then: TYP = Opt_CRJ.Tag
If opt_GJ.Value = True Then: TYP = opt_GJ.Tag

    rec.Open ("EXEC [dbo].[Proc_GetMaxJevSeries_New] @transtype = " & TYP & ",@jevyeardate = '" & DatePost & "' ,@fundtype = '" & cmb_fundtype.Text & "'"), opndbaseFMIS, adOpenStatic, adLockOptimistic
    LastJEVSNno = rec.Fields!MAXJEVSERIES
    rec.Close
    
For cc = 1 To lstDetails.ListItems.Count

    If Len(lstDetails.ListItems(cc).ListSubItems(8).Text) > 0 Then
    
        sql = "SELECT tblCMS_CDCheckBook.FundCode, tblAMIS_COllectionDepositt.TransType as TransType, tblAMIS_COllectionDepositt.ptvno as DVNo, " & _
                "          tblAMIS_COllectionDepositt.TransDate as TransDate, tblAMIS_COllectionDepositt.JEVSeriesNo as JEVSeriesNo " & _
                " FROM tblCMS_CDCheckBook INNER JOIN " & _
                "          tblAMIS_COllectionDepositt ON tblCMS_CDCheckBook.DVNo = tblAMIS_COllectionDepositt.PTVno " & _
                " Where (tblAMIS_COllectionDepositt.ActionCode = 1) And (tblCMS_CDCheckBook.ActionCode = 1) " & _
                " GROUP BY tblCMS_CDCheckBook.Fundcode, tblAMIS_COllectionDepositt.TransType, tblAMIS_COllectionDepositt.PTVno, " & _
                "          tblAMIS_COllectionDepositt.TransDate , tblAMIS_COllectionDepositt.JEVSeriesNo " & _
                " HAVING   tblAMIS_COllectionDepositt.PTVno ='" & lstDetails.ListItems(cc).ListSubItems(8).Text & "'"
    
        opnJEV.Open sql, opndbaseFMIS, adOpenStatic, adLockOptimistic
        If opnJEV.RecordCount <> 0 Then
             lstDetails.ListItems(cc).ListSubItems(10).Text = cmb_fundtype.ItemData(cmb_fundtype.ListIndex) & "-" & Right(Year(DatePost), 2) & "-" & Format(Month(DatePost), "00") & "-" & Format(TYP, "00") & "-" & Format(LastJEVSNno, "0000")
            LastJEVSNno = LastJEVSNno + 1
        Else 'No REcord Found yet in the AMIS
            lstDetails.ListItems(cc).ListSubItems(10).Text = "000-00-00-00-xxxxx"
        End If
        opnJEV.Close
        Set opnJEV = Nothing
    End If
Next cc




End Sub

Private Sub cmd_post_Click()
Dim cc As Integer
Dim Ttype As Integer

If Opt_CRJ.Value = True Then: Ttype = Opt_CRJ.Tag
If opt_GJ.Value = True Then: Ttype = opt_GJ.Tag

If MsgBox("Save JEV Nos.?", vbQuestion + vbYesNo, "System Confirmation") = vbYes Then
Animation1.Stop
Animation1.Close
Animation1.Visible = False
            For cc = 1 To lstDetails.ListItems.Count
                
                    If Len(Trim(lstDetails.ListItems(cc).ListSubItems(10).Text)) > 0 Then
                        If IsFormatCorrect(lstDetails.ListItems(cc).ListSubItems(10).Text) = True Then
                            
                            Call GEtCompleteJEVDetails(lstDetails.ListItems(cc).ListSubItems(8).Text, "PTV", lstDetails.ListItems(cc).ListSubItems(2).Text, "", "" _
                            , lstDetails.ListItems(cc).ListSubItems(1).Text, lstDetails.ListItems(cc).ListSubItems(10).Text, "", "", "0", "0", "0", Ttype, "", "", "", cmb_fundtype.Text, "", "", "", "", ExtractJEVSNo(lstDetails.ListItems(cc).ListSubItems(10).Text), DatePost, lstDetails.ListItems(cc).ListSubItems(8).Text)
    
                            'Updating table from PTO....
                            opndbaseFMIS.Execute "Update tblCMS_CDcheckbook set AlreadySaved2JEV=1 where trnno=" & lstDetails.ListItems(cc).Text & ""
                            
                            'Updating Accounting REcord...
                            opndbaseFMIS.Execute "update tblAMIS_COllectionDepositt set JEVNo='" & lstDetails.ListItems(cc).ListSubItems(10).Text & "',JEVSeriesNo=" & ExtractJEVSNo(lstDetails.ListItems(cc).ListSubItems(10).Text) & ",JEVBy='" & ActiveUserID & "',JEVDate='" & Date & "' where ptvno='" & lstDetails.ListItems(cc).ListSubItems(8).Text & "'"
                        
                        End If
                    End If
               
            Next cc
Animation1.Stop
Animation1.Close
Animation1.Visible = False
MsgBox "Posting to JEV, Successful!", vbInformation, "System Information"
lstDetails.ListItems.Clear
Command1_Click 'Loading Back Active Cash Disbursement Numbers...
List2.ListIndex = GetIndex4ListBox(List2, FMISNo)
End If

End Sub

Private Sub Command1_Click()
If Opt_CRJ.Value = True Or opt_GJ.Value = True Then
Else
MsgBox "Please Select Transtype first", vbCritical, "System Message"
Call cmb_FundType_Click
Exit Sub
End If
Animation1.Visible = True
Animation1.Open App.path & AViLocation & "\horizontaloading.avi"
Animation1.Play
Label13.Caption = "Loading, Please wait.."
Label13.Refresh
Call LoadSavedReport(ActiveUserID, DTPicker1.Year, DTPicker1.Month, cmb_fundtype.Text)
Label13.Caption = ""
Animation1.Stop
Animation1.Close
Animation1.Visible = False
End Sub

Private Sub Command2_Click()

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

Private Sub ClearCmb()
cmb_Fund.Clear
cmb_bank.Clear
cmb_accnumber.Clear
cmb_AccountName.Clear

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
MSHFlexGrid1.TextMatrix(0, 1) = "Particular"
MSHFlexGrid1.TextMatrix(0, 2) = "Transaction date"
MSHFlexGrid1.TextMatrix(0, 3) = "Ref."
MSHFlexGrid1.TextMatrix(0, 4) = "Code"
MSHFlexGrid1.TextMatrix(0, 5) = "Amount"
MSHFlexGrid1.TextMatrix(0, 6) = "Balance Amt."
MSHFlexGrid1.TextMatrix(0, 7) = "ReconcilingSeqNo"
MSHFlexGrid1.TextMatrix(0, 8) = "DVNo"
MSHFlexGrid1.TextMatrix(0, 9) = "AlreadySaved"
MSHFlexGrid1.TextMatrix(0, 10) = "JEVNo"



MSHFlexGrid1.ColWidth(0) = 0
MSHFlexGrid1.ColWidth(1) = 2300
MSHFlexGrid1.ColWidth(2) = 1400
MSHFlexGrid1.ColWidth(3) = 0
MSHFlexGrid1.ColWidth(4) = 0
MSHFlexGrid1.ColWidth(5) = 1300
MSHFlexGrid1.ColWidth(6) = 1300
MSHFlexGrid1.ColWidth(7) = 1800
MSHFlexGrid1.ColWidth(8) = 2000
MSHFlexGrid1.ColWidth(9) = 0
MSHFlexGrid1.ColWidth(10) = 2000

For cc = 0 To MSHFlexGrid1.Cols - 1
    MSHFlexGrid1.Row = 0
    MSHFlexGrid1.col = cc
    MSHFlexGrid1.CellAlignment = 4
Next cc
End Sub

Private Function GEtTotal() As Currency
Dim cc As Integer

For cc = 1 To lstDetails.ListItems.Count
    If GEtTotal <> 0 Then
        GEtTotal = GEtTotal + CCur(lstDetails.ListItems(cc).ListSubItems(5))
    Else
        GEtTotal = CCur(lstDetails.ListItems(cc).ListSubItems(5))
    End If
Next cc
End Function


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
Dim Ttype As Integer

If Opt_CRJ.Value = True Then: Ttype = Opt_CRJ.Tag
If opt_GJ.Value = True Then: Ttype = opt_GJ.Tag

If fund = "Provincial Learning Center" Then: SFCOde = "2473,36052,36055"
If fund = "Agricultural Resource Center" Then: SFCOde = "2476"
If fund = "Patin-ay Waterworks System" Then: SFCOde = "2475,36056"
If fund = "Eco-ASERBAC" Then: SFCOde = "2477"

If fund = "Provincial Learning Center" Or fund = "Agricultural Resource Center" Or fund = "Patin-ay Waterworks System" Or fund = "Eco-ASERBAC" Then
fund = "Economic Enterprises"
End If


'SFCOde = ""
'If fund = "Eco-PTC" Then: SFCOde = "2473,36052,36055"
'If fund = "Eco-PNB" Then: SFCOde = "2476"
'If fund = "Eco-WATERWORKS" Then: SFCOde = "2475,36056"
'If fund = "Eco-ASERBAC" Then: SFCOde = "2477"
'
'If fund = "Eco-PTC" Or fund = "Eco-PNB" Or fund = "Eco-WATERWORKS" Or fund = "Eco-ASERBAC" Then
'fund = "Economic Enterprises"
'End If

Select Case Ttype
    Case 1
     Select Case (fund):
        Case "Economic Enterprises"
        
        sql = " SELECT  tblCMS_CDCheckBook.dvno, tblCMS_CDCheckBook.CompositionCode, " & _
                              " tblCMS_CDCheckBook.Amount, tblCMS_CDCheckBook.ChkNumber, tblCMS_CDCheckBook.UserID, " & _
                              " tblCMS_CDCheckBook.DebitCredit, vw_DepositoryBank.FundType FROM  tblCMS_CDCheckBook INNER JOIN " & _
                              " vw_DepositoryBank ON tblCMS_CDCheckBook.CompositionCode = vw_DepositoryBank.FMISAccountCode " & _
                              " WHERE  (tblCMS_CDCheckBook.Actioncode = 1) AND " & _
                              " (tblCMS_CDCheckBook.ChkNUmber IS NOT NULL) and (right(left([ChkNumber],8),1)) = '-' and tblCMS_CDCheckBook.transactiontype = 'DS' AND (tblCMS_CDCheckBook.CompositionCode in (" & SFCOde & ")) AND " & _
                              " (YEAR(tblCMS_CDCheckBook.transactiondate) = " & TrnYear & ") AND (MONTH(tblCMS_CDCheckBook.transactiondate) = '" & trnMonth & "' AND tblCMS_CDCheckBook.alreadysaved2jev = 0) OR " & _
                              " (tblCMS_CDCheckBook.Actioncode = 1) AND (tblCMS_CDCheckBook.DebitCredit = 1) AND " & _
                              "  (tblCMS_CDCheckBook.CompositionCode in (" & SFCOde & ")) AND " & _
                              " (year(tblCMS_CDCheckBook.transactiondate) = " & TrnYear & ") AND tblCMS_CDCheckBook.alreadysaved2jev = 0 AND (MONTH(tblCMS_CDCheckBook.transactiondate) = '" & trnMonth & "') and (tblCMS_CDCheckBook.transactiontype = 'DS')  and ( (right(left([ChkNumber],8),1)) = '-' )" & _
                             " ORDER BY SUBSTRING(tblCMS_CDCheckBook.dvno, 16, 6),SUBSTRING(tblCMS_CDCheckBook.dvno, 1, 15)"
        Case Else
        sql = " SELECT  tblCMS_CDCheckBook.dvno, tblCMS_CDCheckBook.CompositionCode, " & _
                              " tblCMS_CDCheckBook.Amount, tblCMS_CDCheckBook.ChkNumber, tblCMS_CDCheckBook.UserID, " & _
                              " tblCMS_CDCheckBook.DebitCredit, vw_DepositoryBank.FundType FROM  tblCMS_CDCheckBook INNER JOIN " & _
                              " vw_DepositoryBank ON tblCMS_CDCheckBook.CompositionCode = vw_DepositoryBank.FMISAccountCode " & _
                              " WHERE  (tblCMS_CDCheckBook.Actioncode = 1) AND " & _
                              " (tblCMS_CDCheckBook.ChkNUmber IS NOT NULL) and (right(left([ChkNumber],8),1)) = '-' and (tblCMS_CDCheckBook.transactiontype = 'DS' ) AND (vw_DepositoryBank.FundType = '" & fund & "') AND " & _
                              " (YEAR(tblCMS_CDCheckBook.transactiondate) = " & TrnYear & ") AND (MONTH(tblCMS_CDCheckBook.transactiondate) = '" & trnMonth & "' AND tblCMS_CDCheckBook.alreadysaved2jev = 0) OR " & _
                              " (tblCMS_CDCheckBook.Actioncode = 1) AND (tblCMS_CDCheckBook.DebitCredit = 1) AND " & _
                              "  (vw_DepositoryBank.FundType = '" & fund & "') AND " & _
                              " (year(tblCMS_CDCheckBook.transactiondate) = " & TrnYear & ") AND tblCMS_CDCheckBook.alreadysaved2jev = 0 AND (MONTH(tblCMS_CDCheckBook.transactiondate) = '" & trnMonth & "') and (tblCMS_CDCheckBook.transactiontype = 'DS') and ((right(left([ChkNumber],8),1)) = '-') " & _
                             " ORDER BY SUBSTRING(tblCMS_CDCheckBook.dvno, 16, 6),SUBSTRING(tblCMS_CDCheckBook.dvno, 1, 15)"
        End Select
    Case 4
     Select Case (fund):
        Case "Economic Enterprises"
        
        sql = " SELECT  tblCMS_CDCheckBook.dvno, tblCMS_CDCheckBook.CompositionCode, " & _
                              " tblCMS_CDCheckBook.Amount, tblCMS_CDCheckBook.ChkNumber, tblCMS_CDCheckBook.UserID, " & _
                              " tblCMS_CDCheckBook.DebitCredit, vw_DepositoryBank.FundType FROM  tblCMS_CDCheckBook INNER JOIN " & _
                              " vw_DepositoryBank ON tblCMS_CDCheckBook.CompositionCode = vw_DepositoryBank.FMISAccountCode " & _
                              " WHERE  (tblCMS_CDCheckBook.Actioncode = 1) AND " & _
                              " (tblCMS_CDCheckBook.ChkNUmber IS NOT NULL) and (right(left([ChkNumber],8),1)) = '-' AND (tblCMS_CDCheckBook.CompositionCode in (" & SFCOde & ")) AND " & _
                              " (YEAR(tblCMS_CDCheckBook.transactiondate) = " & TrnYear & ") AND (MONTH(tblCMS_CDCheckBook.transactiondate) = '" & trnMonth & "' AND tblCMS_CDCheckBook.alreadysaved2jev = 0) OR " & _
                              " (tblCMS_CDCheckBook.Actioncode = 1) AND (tblCMS_CDCheckBook.DebitCredit = 1) AND " & _
                              "  (tblCMS_CDCheckBook.CompositionCode in (" & SFCOde & ")) AND " & _
                              " (year(tblCMS_CDCheckBook.transactiondate) = " & TrnYear & ") AND tblCMS_CDCheckBook.alreadysaved2jev = 0 AND (MONTH(tblCMS_CDCheckBook.transactiondate) = '" & trnMonth & "') and ( tblCMS_CDCheckBook.ChkNumber = 'JV')" & _
                             " ORDER BY SUBSTRING(tblCMS_CDCheckBook.dvno, 16, 6),SUBSTRING(tblCMS_CDCheckBook.dvno, 1, 15)"
        Case Else
        sql = " SELECT  tblCMS_CDCheckBook.dvno, tblCMS_CDCheckBook.CompositionCode " & _
                              "  FROM  tblCMS_CDCheckBook INNER JOIN " & _
                              " vw_DepositoryBank ON tblCMS_CDCheckBook.CompositionCode = vw_DepositoryBank.FMISAccountCode " & _
                              " WHERE  (tblCMS_CDCheckBook.Actioncode = 1) AND " & _
                              " (tblCMS_CDCheckBook.ChkNUmber IS NOT NULL) and (right(left([ChkNumber],8),1)) = '-' and ( tblCMS_CDCheckBook.ChkNumber = 'JV') AND (vw_DepositoryBank.FundType = '" & fund & "') AND " & _
                              " (YEAR(tblCMS_CDCheckBook.transactiondate) = " & TrnYear & ") AND (MONTH(tblCMS_CDCheckBook.transactiondate) = '" & trnMonth & "' AND tblCMS_CDCheckBook.alreadysaved2jev = 0) OR " & _
                              " (tblCMS_CDCheckBook.Actioncode = 1) AND (tblCMS_CDCheckBook.DebitCredit = 1) AND " & _
                              "  (vw_DepositoryBank.FundType = '" & fund & "') AND " & _
                              " (year(tblCMS_CDCheckBook.transactiondate) = " & TrnYear & ") AND tblCMS_CDCheckBook.alreadysaved2jev = 0 AND (MONTH(tblCMS_CDCheckBook.transactiondate) = '" & trnMonth & "') and ( tblCMS_CDCheckBook.ChkNumber = 'JV') " & _
                             " ORDER BY SUBSTRING(tblCMS_CDCheckBook.dvno, 16, 6),SUBSTRING(tblCMS_CDCheckBook.dvno, 1, 15)"
        
        End Select
    End Select
'opnvoucher.Open "Select RecordID,compositioncode from vw_CDCreateRDONo where RDONo is not null and year(checkdate)=" & TrnYear & " and month(checkdate)=" & trnMonth & " and userid='" & UserID & "' or len(RDONo)<>0 and year(checkdate)=" & TrnYear & " and month(checkdate)=" & trnMonth & " order by RecordID", opndbaseFMIS, adOpenStatic, adLockOptimistic
'Debug.Print sql
'MsgBox sql
opnvoucher.Open sql, opndbaseFMIS, adOpenStatic, adLockOptimistic



If opnvoucher.RecordCount <> 0 Then
    List2.Clear
    Do Until opnvoucher.EOF
        List2.AddItem (opnvoucher!dvno)
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



Private Sub grd_details_Click()

End Sub

Private Sub LoadBackBreakdown(ByVal ptv As String)
Dim opnvoucher As New ADODB.Recordset
Dim sql As String
Dim x


sql = "Select * from vw_MP_CashReceiptsAdvanceBreakdown   where dvno='" & ptv & "' and CompositionCode=" & List2.ItemData(List2.ListIndex) & " and actioncode = 1  order by trnno"

'Debug.Print sql

opnvoucher.Open sql, opndbaseFMIS, adOpenStatic, adLockOptimistic

If opnvoucher.RecordCount <> 0 Then
    
    '---------------------------------------------
    
    'Call SetGrid
   ' MSHFlexGrid1.Rows = opnvoucher.RecordCount + 1
   ' Do Until opnvoucher.EOF
   
        Set x = lstDetails.ListItems.Add(, , IIf(IsNull(opnvoucher!Trnno), 0, opnvoucher!Trnno))
        x.SubItems(1) = IIf(IsNull(opnvoucher!Particular), "", opnvoucher!Particular)
        x.SubItems(2) = IIf(IsNull(opnvoucher!TransactionDate), "", opnvoucher!TransactionDate)
'       X.SubItems(4) = IIf(IsNull(opnvoucher!RefNo), "", opnvoucher!RefNo)
        x.SubItems(4) = opnvoucher!fundcode & "-" & opnvoucher!MotherFund
        x.SubItems(5) = Format(opnvoucher!amount, "###,##0.00") 'Cash Advanced Amount
        x.SubItems(6) = IIf(IsNull(opnvoucher!balanceamt), "", Format(opnvoucher!balanceamt, "###,##0.00")) 'Liquiditing Amount
        x.SubItems(7) = IIf(IsNull(opnvoucher!ReconcilingSeqNo), "", opnvoucher!ReconcilingSeqNo) 'Normal Balance
        x.SubItems(8) = opnvoucher!dvno
        'X.SubItems(9) = IIf(IsNull(opnvoucher!RefundORNo), 0, 1)
        
        x.SubItems(11) = IIf(IsNull(opnvoucher!chknumber), "", opnvoucher!chknumber)
        x.SubItems(12) = IIf(IsNull(opnvoucher!FundType), "", opnvoucher!FundType)
        x.SubItems(13) = IIf(IsNull(opnvoucher!BankID), "", opnvoucher!BankID)
        x.SubItems(14) = IIf(IsNull(opnvoucher!BankAccountNo), "", opnvoucher!BankAccountNo)
        x.SubItems(15) = IIf(IsNull(opnvoucher!Accountname), "", opnvoucher!Accountname)
        
        
       ' opnvoucher.MoveNext
   ' Loop
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
End Sub

Private Sub Frame1_Click()
If Opt_CRJ.Value = False Or opt_GJ.Value = False Then
MsgBox "Please Select Transaction type..", vbInformation, "System Message"
Exit Sub
End If
End Sub

Private Sub List2_DblClick()
If IfExist = True Then
     Label13.Caption = "Loading Details..."
    Label13.Refresh
    FMISNo = List2.Text
    Call LoadBackBreakdown(FMISNo)
    txt_RecordID.Text = FMISNo
    Label13.Caption = ""
    Label14.Caption = (MSHFlexGrid1.Rows - 1) & " Voucher/s Found..."
Else
    MsgBox "PTV No Already in the List", vbCritical, "System Message"
End If
End Sub

Private Sub List2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Call List2_DblClick
End If
End Sub

Private Sub lstDetails_Click()
If lstDetails.ListItems.Count <> 0 Then
    Call ClearCmb
    txt_RDNo.Text = ""
    Call LoadFundType(cmb_Fund)
    txt_RDNo.Text = lstDetails.SelectedItem.SubItems(11)
    cmb_Fund.ListIndex = GetIndex(cmb_Fund, lstDetails.SelectedItem.SubItems(12))
    cmb_bank.ListIndex = GetIndex(cmb_bank, lstDetails.SelectedItem.SubItems(13))
'    cmb_accnumber.ListIndex = GetIndex(cmb_accnumber, lstDetails.SelectedItem.SubItems(14))
'    cmb_AccountName.ListIndex = GetIndex(cmb_AccountName, lstDetails.SelectedItem.SubItems(15))
End If
Check1.Value = 0
Check2.Value = 0
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




Private Sub lstDetails_DblClick()
If Len(lstDetails.SelectedItem.SubItems(8)) > 0 Then
    ActiveFormCaller = Me.name
    'ForTheGridRowNo = MSHFlexGrid1.Row

    If IFOK(lstDetails.SelectedItem.SubItems(8)) = True Then 'Kung Naa nay JEV No
    Animation1.Visible = True
    Animation1.Open App.path & AViLocation & "\horizontaloading.avi"
    Animation1.Play
'        frmJEVPreparationforColection_New.txt_Jevno = lstDetails.SelectedItem.SubItems(10)
         frmJEVPreparationforColection_New.txtDVNo = lstDetails.SelectedItem.SubItems(8)
'        frmJEVPreparationforColection_New.txt_AlobsNo = lstDetails.SelectedItem.SubItems(11)
'        frmJEVPreparationforColection_New.txt_particular = lstDetails.SelectedItem.SubItems(1)
'        frmJEVPreparationforColection_New.txt_Amount = lstDetails.SelectedItem.SubItems(5)
'        frmJEVPreparationforColection_New.txt_FundType = lstDetails.SelectedItem.SubItems(12)
'        frmJEVPreparationforColection_New.LoadAccountsByFund (frmJEVPreparationforColection_New.txt_FundType)
        frmJEVPreparationforColection_New.Show
    Animation1.Stop
    Animation1.Close
    Animation1.Visible = False
    Else
'        If Opt_CRJ.Value = True Then
'            frmJEVPreparationforColection_New.txtDVNo.Text = lstDetails.SelectedItem.SubItems(8)
'            'frmJEVPreparationforColection_New.loaddt
'
'            frmJEVPreparationforColection_New.Show vbModal
'        ElseIf Opt_GJ.Value = True Then
'            frmJEVPreparationforGeneralJournal_New.txtDVNo.Text = lstDetails.SelectedItem.SubItems(8)
'           ' frmJEVPreparationforGeneralJournal.loaddt
'
'            frmJEVPreparationforGeneralJournal_New.Show vbModal
'        End If
            frmJEVPreparationforColection_New.txtDVNo.Text = lstDetails.SelectedItem.SubItems(8)
           ' frmJEVPreparationforGeneralJournal.loaddt
           frmJEVPreparationforColection_New.Show
    End If
Else
    MsgBox "There is no Voucher Attachment for this Check!" & Chr(13) & Chr(13) & "Please Select a New..", vbInformation, "System Information"
End If
End Sub
Private Function IFOK(ByVal ptvNo As String) As Boolean
Dim rec As New ADODB.Recordset
IFOK = False
rec.Open "Select top 1 ptvno from tblAMIS_COllectionDepositt where PTVno = '" & ptvNo & "'", opndbaseFMIS, adOpenStatic, adLockOptimistic
If rec.RecordCount <> 0 Then
    IFOK = True
End If
rec.Close
Set rec = Nothing
End Function


Private Sub lstDetails_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo bad
Dim x As Integer
Dim z As Integer
If KeyCode = vbKeyDelete Then
        For x = 1 To lstDetails.ListItems.Count
                If lstDetails.ListItems(x).Checked = True Then
                    lstDetails.ListItems.Remove (x)
                    x = x - 1
                End If
        Next x
End If
Exit Sub
bad:
    If err.Number = 35600 Then
    Exit Sub
    End If
End Sub
Private Function IfExist() As Boolean
Dim x As Integer
IfExist = True
If lstDetails.ListItems.Count = 0 Then
IfExist = True
Else
        For x = 1 To lstDetails.ListItems.Count
            If List2.Text = lstDetails.ListItems(x).SubItems(8) Then
                IfExist = False
            End If
        Next x
End If
End Function
Private Function CkeckOption(ByVal lect As String)
Dim x As Integer
For x = 1 To lstDetails.ListItems.Count
        Select Case lect:
        Case "deselect"
            lstDetails.ListItems(x).Checked = False
        Case Else
            lstDetails.ListItems(x).Checked = True
        End Select
Next x
End Function

Private Sub MSHFlexGrid1_DblClick()
If Len(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 8)) > 0 Then
    ActiveFormCaller = Me.name
    ForTheGridRowNo = MSHFlexGrid1.Row

    If MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 10) <> "000-00-00-00-xxxxx" Then 'Kung Naa nay JEV No
        'frmJEVPreparationforColection_New.txt_JEVNO.Text = MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 10)
        frmJEVPreparationforColection_New.txtDVNo.Text = MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 8)
        'frmJEVPreparationforColection_New.txtAlobsNo.Text = txt_RDNo.Text
        'frmJEVPreparationforColection_New.txtParticular.Text = MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 1)
        'frmJEVPreparationforColection_New.txtAmount.Text = MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 5)
        'frmJEVPreparationforColection_New.txtFund.Text = cmb_Fund.Text
        'frmJEVPreparationforColection_New.LoadAccountsByFund (frmJEVPreparationforColection_New.txt_FundType.Text)
        frmJEVPreparationforColection_New.Show vbModal
    Else
        frmJEVPreparationforColection_New.txtDVNo.Text = MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 8)
        'frmJEVPreparationforColection_New.loaddt
        
        frmJEVPreparationforColection_New.Show vbModal
        Call cmd_Mass_Click
    End If
Else
    MsgBox "There is no Voucher Attachment for this Check!" & Chr(13) & Chr(13) & "Please Select a New..", vbInformation, "System Information"
End If

End Sub

Private Sub Opt_CRJ_Click()
Opt_CRJ.Value = True
Call cmb_FundType_Click
Call Command1_Click
End Sub

Private Sub Opt_GJ_Click()
opt_GJ.Value = True
Call cmb_FundType_Click
Call Command1_Click
End Sub

Private Sub Timer1_Timer()
Call SetGrid
Call LoadFundType(cmb_fundtype)
'Call LoadSavedReport(ActiveUserID, DTPicker1.Year, DTPicker1.Month, cmb_FundType.Text)
Timer1.Enabled = False
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
    Case 1 'Print
        
    Case 3 'Close
        Unload Me
End Select
End Sub


Private Sub txt_RecordID_Click()
'If Len(Trim(txt_RecordID.Text)) <> "" Then
'    txt_RecordID.SelStart = 0
'    txt_RecordID.SelLength = Len(txt_RecordID.Text)
'    txt_RecordID.SetFocus
'Else
'    txt_RecordID.SetFocus
'End If
End Sub

Private Sub txt_RecordID_KeyDown(KeyCode As Integer, Shift As Integer)
Dim tmpVal As Long

On Error GoTo handler
If KeyCode = 13 Then
    If Len(Trim(txt_RecordID.Text)) <> 0 Then
        If txt_RecordID.Text <> "" Then
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
           ' txt_RecordID.Text = "FMISNo-" & Val(txt_RecordID.Text)
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
'Select Case KeyAscii
'    Case 45, 48 To 57, 13
'    Case Else
'        KeyAscii = 0
'End Select
End Sub
Private Function GetBalance(ByVal RowNo As Integer) As Currency
GetBalance = CCur(MSHFlexGrid1.TextMatrix(RowNo, 5)) - CCur(MSHFlexGrid1.TextMatrix(RowNo, 6))
End Function
Private Function GetTotalSelColAmount(ByVal Colno As Integer) As Currency
Dim cc As Integer

For cc = 1 To MSHFlexGrid1.Rows - 1
'    GetTotalSelColAmount = GetTotalSelColAmount + CCur(MSHFlexGrid1.TextMatrix(cc, Colno))
Next cc
End Function

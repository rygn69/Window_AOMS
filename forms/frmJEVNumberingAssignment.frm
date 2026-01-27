VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{1693405E-2DC9-4248-B52F-4AC9145DA2AF}#1.0#0"; "WinXPC Engine.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmJEVNumberingAssignment 
   Caption         =   "JEV Numbering Assignment"
   ClientHeight    =   9165
   ClientLeft      =   285
   ClientTop       =   720
   ClientWidth     =   14505
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmJEVNumberingAssignment.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9165
   ScaleWidth      =   14505
   Begin VB.TextBox txt_Jevno 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   9720
      MaxLength       =   18
      TabIndex        =   58
      Top             =   3720
      Width           =   3420
   End
   Begin VB.Frame frmTrans 
      Caption         =   "Transaction type"
      Enabled         =   0   'False
      ForeColor       =   &H80000017&
      Height          =   735
      Left            =   7320
      TabIndex        =   34
      Top             =   120
      Width           =   2775
      Begin VB.OptionButton optNonObR 
         Caption         =   "Non-ObR"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1440
         TabIndex        =   36
         Tag             =   "1"
         Top             =   360
         Width           =   1215
      End
      Begin VB.OptionButton optObR 
         Caption         =   "With ObR"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   35
         Tag             =   "0"
         Top             =   360
         Value           =   -1  'True
         Width           =   1215
      End
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
      Height          =   2190
      Left            =   225
      TabIndex        =   0
      Top             =   1065
      Width           =   14115
      Begin VB.ComboBox cmbNonAlobs 
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
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   37
         Top             =   585
         Width           =   4380
      End
      Begin VB.ComboBox cmbrc 
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
         Left            =   5160
         Locked          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   29
         Top             =   600
         Width           =   4335
      End
      Begin VB.CommandButton cmd 
         Caption         =   "...."
         Height          =   375
         Left            =   4560
         TabIndex        =   33
         Top             =   1440
         Width           =   375
      End
      Begin VB.CommandButton Command3 
         Caption         =   "&e"
         Height          =   255
         Left            =   9240
         TabIndex        =   31
         Top             =   1080
         Width           =   255
      End
      Begin VB.TextBox txt_Claimant 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Top             =   1425
         Width           =   4380
      End
      Begin VB.TextBox txt_particular 
         Height          =   720
         Left            =   5160
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Top             =   1425
         Width           =   4290
      End
      Begin VB.TextBox txt_Amount 
         Alignment       =   2  'Center
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
         Left            =   9870
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   1485
         Width           =   4020
      End
      Begin VB.TextBox txt_FundType 
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
         Left            =   9870
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   600
         Width           =   4020
      End
      Begin VB.TextBox txt_AlobsNo 
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
         Left            =   120
         TabIndex        =   4
         Top             =   585
         Width           =   4380
      End
      Begin VB.TextBox txtclaimantcode 
         Height          =   360
         Left            =   1080
         TabIndex        =   32
         Top             =   1440
         Visible         =   0   'False
         Width           =   570
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
         Left            =   5100
         TabIndex        =   11
         Top             =   270
         Width           =   2235
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Claimant"
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
         TabIndex        =   10
         Top             =   1050
         Width           =   930
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Alobs/OBR No:"
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
         TabIndex        =   9
         Top             =   270
         Width           =   1605
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Particular"
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
         Left            =   5100
         TabIndex        =   8
         Top             =   1050
         Width           =   990
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Amount (Gross)"
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
         Left            =   9900
         TabIndex        =   7
         Top             =   1170
         Width           =   1695
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fund Type"
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
         Left            =   9900
         TabIndex        =   6
         Top             =   270
         Width           =   1125
      End
   End
   Begin VB.TextBox txtDate 
      Height          =   360
      Left            =   11880
      Locked          =   -1  'True
      TabIndex        =   27
      Top             =   465
      Width           =   2565
   End
   Begin VB.CommandButton cmdsave 
      Caption         =   "&Update"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5760
      Picture         =   "frmJEVNumberingAssignment.frx":076A
      TabIndex        =   25
      ToolTipText     =   "Saves to Journal Entry"
      Top             =   120
      Width           =   1335
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6600
      Top             =   360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJEVNumberingAssignment.frx":4264
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJEVNumberingAssignment.frx":46B6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Assign No. to JEV"
      Height          =   840
      Left            =   13335
      TabIndex        =   22
      Top             =   3540
      Width           =   990
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
      Height          =   4575
      Left            =   210
      ScaleHeight     =   4545
      ScaleWidth      =   14070
      TabIndex        =   17
      Top             =   4560
      Width           =   14100
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
         TabIndex        =   38
         Top             =   0
         Visible         =   0   'False
         Width           =   14175
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
            Left            =   5400
            Locked          =   -1  'True
            TabIndex        =   48
            Top             =   1320
            Width           =   2655
         End
         Begin VB.CommandButton Command6 
            Caption         =   "&Hide"
            Height          =   495
            Left            =   12960
            Picture         =   "frmJEVNumberingAssignment.frx":49D0
            TabIndex        =   47
            Top             =   240
            Width           =   975
         End
         Begin VB.TextBox txtCClaimant 
            BackColor       =   &H80000000&
            Height          =   360
            Left            =   5400
            Locked          =   -1  'True
            TabIndex        =   46
            Top             =   360
            Width           =   2655
         End
         Begin VB.TextBox txtCamount 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000000&
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   5400
            Locked          =   -1  'True
            TabIndex        =   45
            Top             =   840
            Width           =   2655
         End
         Begin VB.TextBox txtCParticular 
            BackColor       =   &H80000000&
            Height          =   600
            Left            =   1200
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   44
            Top             =   1800
            Width           =   6855
         End
         Begin VB.TextBox txtCChecdate 
            Alignment       =   2  'Center
            BackColor       =   &H80000000&
            Height          =   360
            Left            =   1200
            Locked          =   -1  'True
            TabIndex        =   43
            Top             =   1320
            Width           =   2535
         End
         Begin VB.TextBox txtCCheckno 
            Alignment       =   2  'Center
            BackColor       =   &H80000000&
            Height          =   360
            Left            =   1200
            Locked          =   -1  'True
            TabIndex        =   42
            Top             =   840
            Width           =   2535
         End
         Begin VB.CommandButton cmdadd 
            Caption         =   "&Add"
            Height          =   375
            Left            =   3840
            TabIndex        =   41
            Top             =   360
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.TextBox txtCDvno 
            Alignment       =   2  'Center
            BackColor       =   &H80000000&
            Height          =   360
            Left            =   1200
            Locked          =   -1  'True
            TabIndex        =   40
            Top             =   360
            Width           =   2535
         End
         Begin VB.CommandButton Command4 
            Caption         =   "&Edit List"
            Height          =   495
            Left            =   8640
            Picture         =   "frmJEVNumberingAssignment.frx":529A
            TabIndex        =   39
            Top             =   240
            Visible         =   0   'False
            Width           =   975
         End
         Begin MSComctlLib.ListView ListView1 
            Height          =   1935
            Left            =   120
            TabIndex        =   49
            Top             =   2520
            Width           =   13815
            _ExtentX        =   24368
            _ExtentY        =   3413
            View            =   3
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
            NumItems        =   7
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "trnno"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   1
               Text            =   "Voucher No."
               Object.Width           =   4410
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   2
               Text            =   "checkno"
               Object.Width           =   3528
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
               Object.Width           =   7056
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "claimant"
               Object.Width           =   3528
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   6
               Text            =   "amount"
               Object.Width           =   3528
            EndProperty
         End
         Begin VB.Label Label21 
            Caption         =   "Total Amount:"
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
            Left            =   3840
            TabIndex        =   56
            Top             =   1440
            Width           =   1575
         End
         Begin VB.Label Label20 
            Caption         =   "Claimant:"
            Height          =   375
            Left            =   4440
            TabIndex        =   55
            Top             =   480
            Width           =   975
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
            Left            =   4440
            TabIndex        =   54
            Top             =   960
            Width           =   975
         End
         Begin VB.Label Label18 
            Caption         =   "Particular:"
            Height          =   375
            Left            =   120
            TabIndex        =   53
            Top             =   1920
            Width           =   1095
         End
         Begin VB.Label Label17 
            Caption         =   "Checkdate:"
            Height          =   375
            Left            =   120
            TabIndex        =   52
            Top             =   1320
            Width           =   1095
         End
         Begin VB.Label Label16 
            Caption         =   "Checkno:"
            Height          =   375
            Left            =   120
            TabIndex        =   51
            Top             =   840
            Width           =   975
         End
         Begin VB.Label Label15 
            Caption         =   "Dvno:"
            Height          =   375
            Left            =   120
            TabIndex        =   50
            Top             =   360
            Width           =   615
         End
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
         Left            =   240
         TabIndex        =   24
         Top             =   0
         Visible         =   0   'False
         Width           =   1665
      End
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
         Left            =   240
         TabIndex        =   23
         Text            =   "cmbEntry"
         Top             =   600
         Visible         =   0   'False
         Width           =   1575
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   4200
         Left            =   0
         TabIndex        =   18
         Top             =   0
         Width           =   14070
         _ExtentX        =   24818
         _ExtentY        =   7408
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
      Height          =   720
      Left            =   225
      TabIndex        =   12
      Top             =   3255
      Width           =   8190
      Begin VB.OptionButton opn_Coll 
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
         Left            =   240
         TabIndex        =   16
         Tag             =   "01"
         Top             =   885
         Width           =   1260
      End
      Begin VB.OptionButton opn_CheckDisb 
         Caption         =   "Check Disbursement"
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
         Left            =   165
         TabIndex        =   15
         Tag             =   "02"
         Top             =   300
         Width           =   2580
      End
      Begin VB.OptionButton opn_CashDisb 
         Caption         =   "Cash Disbursement"
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
         Left            =   3000
         TabIndex        =   14
         Tag             =   "03"
         Top             =   300
         Width           =   2580
      End
      Begin VB.OptionButton opn_Other 
         Caption         =   "General Journal"
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
         Left            =   5925
         TabIndex        =   13
         Tag             =   "04"
         Top             =   300
         Width           =   2070
      End
   End
   Begin VB.TextBox txt_DVNo 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   420
      TabIndex        =   20
      Top             =   360
      Width           =   4845
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Update Accountcode"
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
      Left            =   240
      TabIndex        =   26
      Top             =   4800
      Width           =   2175
   End
   Begin WinXPC_Engine.WindowsXPC WindowsXPC1 
      Left            =   8160
      Top             =   5880
      _ExtentX        =   6588
      _ExtentY        =   1085
      ColorScheme     =   2
      Common_Dialog   =   0   'False
      PictureControl  =   0   'False
   End
   Begin VB.Label Label22 
      Caption         =   "<<--Cash Advance Details"
      Height          =   255
      Left            =   5640
      TabIndex        =   57
      Top             =   4200
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Label Label2 
      Caption         =   $"frmJEVNumberingAssignment.frx":5B64
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   30
      Top             =   4080
      Width           =   4335
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
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   11880
      TabIndex        =   28
      Top             =   240
      Width           =   1035
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DV Number:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   180
      TabIndex        =   21
      Top             =   75
      Width           =   870
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Assigned JEV No."
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
      Left            =   8865
      TabIndex        =   19
      Top             =   3495
      Width           =   1590
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   1110
      Left            =   8760
      Top             =   3360
      Width           =   6735
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   960
      Left            =   0
      Top             =   0
      Width           =   5700
   End
End
Attribute VB_Name = "frmJEVNumberingAssignment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
'
'Option Explicit
'
'Dim Edited As Boolean
'Dim xDebit As Currency
'Dim xCredit As Currency
'Dim xObR As String
'Dim xNAcode, DVN As String
'Dim CUFlag As Boolean           'Claimant Update Flag
'Dim XFlag As Boolean
'Dim rcedit As Boolean
'Dim CAedit As Boolean
'Public IfNew As Boolean
'Public isfrom_jevNumbering As Boolean
'Dim not_coloraly_total_debit, not_coloraly_total_credit, coloraly_total_debit, coloraly_total_credit As Double
'Public ptv, Approvedid, Datetimeapproved, Logoutby, Logoutdatetime, Logoutremark, Continuing As String
'Private Sub LoadBackDVDetails(ByVal DVNo As String)
'Dim opnDV As New ADODB.Recordset
'Dim y As String
'opnDV.Open "Select top 1 * from tblAMIS_IncomingDVTrns where DVNo='" & DVNo & "' and actioncode=1 order by trnno desc", opndbaseFMIS, adOpenStatic, adLockOptimistic
'If opnDV.RecordCount <> 0 Then
'    LoadTrans
'    txt_FundType.Text = opnDV!FundType
'    txtclaimantcode.Text = IIf(IsNull(opnDV!ClaimantCode), "N/A", opnDV!ClaimantCode)
'    cmbrc.Text = GetOfficeName(opnDV![RCenter], "OfficeMedium")
'
'    txt_Claimant.Text = IIf(IsNull(opnDV!ClaimantCode), getClaimantBYdvno(DVNo), GetClaimantDetails(IIf(IsNull(opnDV!ClaimantCode), "", opnDV!ClaimantCode), "Name"))
'   y = getClaimantBYdvno(DVNo)
'    txt_particular.Text = opnDV!Particular
'    txt_Amount.Text = Format(opnDV!Gamount, "#,##0.00")
'    If opnDV!NonAlobs = 1 Then
'    optNonObR.Value = True
'    cmbNonAlobs.Text = GetNonAlobsName(opnDV!obrno)
'    Else
'    optObR.Value = True
'    txt_AlobsNo.Text = opnDV!obrno
'    End If
'    If Trim(cmbNonAlobs.Text) = "Liquidation of Cash Advance" Then
'        Call AllLoadCAdetails(ListView1, txt_DVNo.Text, txtctotalAmnt)
'        Label22.Visible = True
'        fmeCA.Visible = True
'
'    Else
'        fmeCA.Visible = False
'        Label22.Visible = False
'    End If
'Else
'    txt_AlobsNo.Text = ""
'    txt_FundType.Text = ""
'    txt_Claimant.Text = ""
'    txt_particular.Text = ""
'    txt_Amount.Text = ""
'    optNonObR.Value = False
'    optObR.Value = False
'    Call SetGrid
'End If
'opnDV.Close
'Set opnDV = Nothing
'
'Call LoadOtherDetails(DVNo)
''Call LoadAccountsByFund(Trim(Me.txt_FundType))
'End Sub
'Public Sub LoadAccountsByFund(ByVal fundmedium As String)
'Dim ARec As New ADODB.Recordset
'Dim x As Integer
'Dim FundName As String
'
'    cmbEntry.Clear
'    cmbEntry.Visible = False
'    FundName = GetFundName(fundmedium)
'    ARec.Open ("Select distinct * from [tblREF_AIS_ChartofAccounts] Where [Active]=1 and [FundType]='" & FundName & "' Order by [ChildAccountCode]"), opndbaseFMIS, adOpenStatic, adLockOptimistic
'    If ARec.RecordCount > 0 Then
'        For x = 1 To ARec.RecordCount
'            cmbEntry.AddItem ARec![childaccountcode]
'            cmbEntry.ItemData(cmbEntry.NewIndex) = ARec![FmisAccountcode]
'            ARec.MoveNext
'        Next x
'    End If
'    ARec.Close
'    Set ARec = Nothing
'
'End Sub
'
'
'Public Function LoadAcctngEntries(ByVal DVNo As String)
'Dim DRec As New ADODB.Recordset
'Dim rec As New ADODB.Recordset
'Dim x As Integer
'    DRec.Open ("Select ChildAccountcode,Debit ,Credit From tblAMIS_AccoutingEntries Where [reffno]='" & DVNo & "' And (ActionCode=1) "), opndbaseFMIS, adOpenStatic, adLockOptimistic
'    If DRec.RecordCount > 0 Then
'        If EditCount = False Then
'        EditCount = True
'            rec.Open "Select dvno from tblAMIs_tmpjournal where dvno = '" & DVNo & "'", opndbaseFMIS, adOpenStatic
'            If rec.RecordCount > 0 Then
'                    If MsgBox("This transaction Have a temporary Accounting Entries, do you want to Delete?", vbCritical + vbYesNo, "System Information") = vbYes Then
'                        opndbaseFMIS.Execute "Delete from tblAMIs_tmpjournal where Dvno = '" & TxtDvno.Text & "'"
'                        For x = 1 To DRec.RecordCount
'                        DoEvents
'                            opndbaseFMIS.Execute "Insert into tblAMIs_tmpjournal (Dvno,Accountcode,Debit,Credit) values ('" & Trim(TxtDvno.Text) & "','" & Trim(DRec!childaccountcode) & "'," & DRec!Debit & "," & DRec!Credit & ")"
'                            DRec.MoveNext
'                        Next x
'                    End If
'            Else
'            For x = 1 To DRec.RecordCount
'                        DoEvents
'                            opndbaseFMIS.Execute "Insert into tblAMIs_tmpjournal (Dvno,Accountcode,Debit,Credit) values ('" & Trim(TxtDvno.Text) & "','" & Trim(DRec!childaccountcode) & "'," & DRec!Debit & "," & DRec!Credit & ")"
'                            DRec.MoveNext
'                        Next x
'            End If
'            rec.Close
'        End If
'    End If
'    DRec.Close
'    Set DRec = Nothing
'End Function
'Private Sub SetGrid()
'Dim cc As Integer
'
'MSFlexGrid1.Clear
'MSFlexGrid1.Cols = 7
'MSFlexGrid1.Rows = 2
'
'    MSFlexGrid1.TextMatrix(0, 0) = "trnno"
'    MSFlexGrid1.TextMatrix(0, 1) = "FMISCode"
'    MSFlexGrid1.TextMatrix(0, 2) = "Account Code"
'    MSFlexGrid1.TextMatrix(0, 3) = "Accounts and Explanation"
'    MSFlexGrid1.TextMatrix(0, 4) = "Debit"
'    MSFlexGrid1.TextMatrix(0, 5) = "Credit"
'    MSFlexGrid1.TextMatrix(0, 6) = "Actioncode"
'
'MSFlexGrid1.ColWidth(0) = 0
'MSFlexGrid1.ColWidth(1) = 0
'MSFlexGrid1.ColWidth(2) = 3000
'MSFlexGrid1.ColWidth(3) = 5000
'MSFlexGrid1.ColWidth(4) = 2000
'MSFlexGrid1.ColWidth(5) = 2000
'MSFlexGrid1.ColWidth(6) = 1500
'
'For cc = 0 To MSFlexGrid1.Cols - 1
'    MSFlexGrid1.Row = 0
'    MSFlexGrid1.col = cc
'    MSFlexGrid1.CellAlignment = 4
'Next cc
'End Sub
'Private Sub LoadOtherDetails(ByVal DV As String)
'Dim opnDV As New ADODB.Recordset
'
'opnDV.Open "Select * from tblAMIS_JournalEntry where DVNo='" & DV & "' and (actioncode=1 or actioncode=5) ", opndbaseFMIS, adOpenStatic, adLockOptimistic
'
'MSFlexGrid1.TextMatrix(0, 0) = "trnno"
'    MSFlexGrid1.TextMatrix(0, 1) = "FMISCode"
'    MSFlexGrid1.TextMatrix(0, 2) = "Account Code"
'    MSFlexGrid1.TextMatrix(0, 3) = "Accounts and Explanation"
'    MSFlexGrid1.TextMatrix(0, 4) = "Debit"
'    MSFlexGrid1.TextMatrix(0, 5) = "Credit"
'    MSFlexGrid1.TextMatrix(0, 6) = "Actioncode"
'
'If opnDV.RecordCount <> 0 Then
'    Call SelectTrnType(opnDV!Transtype)
'
'    Approvedid = IIf(IsNull(opnDV.Fields!ApprovedByID), "", opnDV.Fields!ApprovedByID)
'    Datetimeapproved = IIf(IsNull(opnDV.Fields!Datetimeapproved), "", opnDV.Fields!Datetimeapproved)
'    Logoutby = IIf(IsNull(opnDV.Fields!Logoutby), "", opnDV.Fields!Logoutby)
'    Logoutremark = IIf(IsNull(opnDV.Fields!Logoutremark), "", opnDV.Fields!Logoutremark)
'    Logoutdatetime = IIf(IsNull(opnDV.Fields!Logoutdatetime), "", opnDV.Fields!Logoutdatetime)
'    Continuing = IIf(IsNull(opnDV.Fields!Continuing), "", opnDV.Fields!Continuing)
'    txtDate.Text = IIf(IsNull(opnDV.Fields!datetimeentered), "", opnDV.Fields!datetimeentered)
''    Do Until opnDV.EOF
''
''
''    MSFlexGrid1.TextMatrix(MSFlexGrid1.Rows - 1, 0) = opnDV!Trnno
''    MSFlexGrid1.TextMatrix(MSFlexGrid1.Rows - 1, 1) = opnDV!FmisAccntCode
''    MSFlexGrid1.TextMatrix(MSFlexGrid1.Rows - 1, 2) = GetAccntDescription(opnDV!FmisAccntCode, "ACCT_CODE")
''    MSFlexGrid1.TextMatrix(MSFlexGrid1.Rows - 1, 3) = GetAccntDescription(opnDV!FmisAccntCode, "ACCT_ENTRIES")
''
''    If opnDV!DebitCredit = 1 Then
''        MSFlexGrid1.TextMatrix(MSFlexGrid1.Rows - 1, 4) = Format(opnDV!Amount, "#,##0.00")
''    ElseIf opnDV!DebitCredit = 0 Then
''        MSFlexGrid1.TextMatrix(MSFlexGrid1.Rows - 1, 5) = Format(opnDV!Amount, "#,##0.00")
''    End If
''    MSFlexGrid1.TextMatrix(MSFlexGrid1.Rows - 1, 6) = opnDV!ActionCode
''    MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1
''    opnDV.MoveNext
''    Loop
'    Call GetSum
'Else
'    Approvedid = ""
'    Datetimeapproved = ""
'    Logoutby = ""
'    Logoutremark = ""
'    Logoutdatetime = ""
'    Continuing = ""
'    Call ClearAllOption
'End If
'opnDV.Close
'Set opnDV = Nothing
'
'End Sub
'
'Private Function GetAccntDescription(ByVal FMISCode As Long, ByVal NeedFld As String) As String
'Dim opnDesc As New ADODB.Recordset
'
'opnDesc.Open "Select * from tblREF_AIS_ChartofAccounts where FMISAccountCode=" & FMISCode & " and active=1", opndbaseFMIS, adOpenStatic, adLockOptimistic
'If opnDesc.RecordCount <> 0 Then
'    Select Case NeedFld
'        Case "ACCT_ENTRIES"
'            If opnDesc!Accountname = opnDesc!AccountNamefull Then
'                GetAccntDescription = opnDesc!Accountname
'            Else
'                GetAccntDescription = opnDesc!Accountname & "-" & opnDesc!AccountNamefull
'            End If
'
'        Case "ACCT_CODE"
'            GetAccntDescription = opnDesc!childaccountcode
'    End Select
'End If
'opnDesc.Close
'
'End Function
'Private Sub ClearAllOption()
'opn_Coll.Value = False
'opn_CheckDisb.Value = False
'opn_CashDisb.Value = False
'opn_Other.Value = False
'End Sub
'Private Sub SelectTrnType(ByVal TransCode As String)
'Select Case TransCode
'    Case 1
'        opn_Coll.Value = True
'    Case 2
'        opn_CheckDisb.Value = True
'    Case 3
'        opn_CashDisb.Value = True
'    Case 4
'        opn_Other.Value = True
'End Select
'End Sub
'
'Private Sub cmbEntry_Click()
'cmbEntry.SetFocus
'End Sub
'
'Private Sub cmbEntry_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        If cmbEntry.ListIndex <> -1 Then
'            MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2) = cmbEntry.Text
'            MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 0) = cmbEntry.ItemData(cmbEntry.ListIndex)
'            If MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2) = "TOTAL" Then
'                MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 3) = ""
'                MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 4) = ""
'            ElseIf MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 4) = "" And MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 5) = "" Then
'                MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 5) = ""
'                MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 4) = ""
'            ElseIf Val(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 4)) > 0 And Val(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 5)) > 0 Then
'                MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 5) = ""
'                MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 4) = ""
'            End If
'            MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 3) = GetAccountNameByFMISAccountCode(cmbEntry.ItemData(cmbEntry.ListIndex))
'            MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 1) = cmbEntry.ItemData(cmbEntry.ListIndex)
'            MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 6) = "1"
'
'        ElseIf cmbEntry.Text = "" Then
'            MSFlexGrid1.RemoveItem (MSFlexGrid1.Row)
'        Else
'        End If
'        cmbEntry.Visible = False
'        MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1
'        Call GetSum
'        MSFlexGrid1.SetFocus
'        Edited = True
'    Else
'        KeyAscii = AutoFind(cmbEntry, KeyAscii, True)
'    End If
'
'End Sub
'Private Sub GetSum1()
'On Error GoTo bad
'Dim x As Integer
'
'    xDebit = 0
'    xCredit = 0
'    For x = 1 To MSFlexGrid1.Rows - 1
'        If MSFlexGrid1.TextMatrix(x, 0) <> "" Then
'            xDebit = xDebit + CCur(IIf(MSFlexGrid1.TextMatrix(x, 3) = "", 0, MSFlexGrid1.TextMatrix(x, 3)))
'            xCredit = xCredit + CCur(IIf(MSFlexGrid1.TextMatrix(x, 4) = "", 0, MSFlexGrid1.TextMatrix(x, 4)))
'        Else
'            MSFlexGrid1.TextMatrix(x, 2) = "TOTAL"
'            MSFlexGrid1.TextMatrix(x, 3) = xDebit
'            MSFlexGrid1.TextMatrix(x, 4) = xCredit
'            Exit For
'        End If
'    Next x
'Exit Sub
'bad:
'MsgBox err.Description
'End Sub
'
'Private Sub LoadNonAlobs()
'Dim NRec As New ADODB.Recordset
'Dim x As Integer
'
'cmbNonAlobs.Clear
'
'NRec.Open ("Select * From tblCMS_CDNoneAlobs Order By NonAlobs"), opndbaseFMIS, adOpenStatic, adLockOptimistic
'If NRec.RecordCount > 0 Then
'    For x = 1 To NRec.RecordCount
'        cmbNonAlobs.AddItem NRec!NonAlobs
'        cmbNonAlobs.ItemData(cmbNonAlobs.NewIndex) = NRec!Trnno
'        NRec.MoveNext
'    Next x
'End If
'NRec.Close
'Set NRec = Nothing
'
'End Sub
'
'Private Sub cmbNonAlobs_Change()
'Ifliquidition
'End Sub
'
'Private Sub cmbRC_Click()
'
''    If Trim(cmbRC.Text) <> "" Then
''        txt_RCenter.Text = Trim(cmbRC.Text)
''        txt_RCenter.Visible = True
''        cmbRC.Visible = False
''    End If
'End Sub
'
'Private Sub cmd_Click()
'CUFlag = True
'ActiveFormCaller = "frmJEVNumberingAssignment"
'frmCDClaimantRegistry.Show 1
'End Sub
'
'Private Sub cmdadd_Click()
'Dim x
'If IfexistDv(txtCDvno.Text) = False Then
'    If txtCDvno.Text <> "" And txtCCheckno.Text <> "" And txtCamount.Text <> "" Then
'        Set x = ListView1.ListItems.Add(, , "")
'            x.SubItems(1) = txtCDvno.Text
'            x.SubItems(2) = txtCCheckno.Text
'            x.SubItems(3) = txtCChecdate.Text
'            x.SubItems(4) = txtCParticular.Text
'            x.SubItems(5) = txtCClaimant.Text
'            x.SubItems(6) = txtCamount.Text
'            txtctotalAmnt.Text = Format(GetCATotalamount(ListView1), "#,##0.00")
'    Else
'    MsgBox "Please check your entry", vbInformation, "System Message"
'    End If
'Else
'    MsgBox "Dvno Already on the List", vbInformation, "System Message"
'End If
'End Sub
'Public Function IfexistDv(ByVal DVNo As String) As Boolean
'Dim y As Integer
'IfexistDv = False
'If ListView1.ListItems.Count <> 0 Then
'    For y = 1 To ListView1.ListItems.Count
'        If DVNo = ListView1.ListItems(y).SubItems(1) Then
'            IfexistDv = True
'        End If
'    Next y
'End If
'End Function
'
'Private Sub Command1_Click()
'If Edited = True Then
'    If MsgBox("Save Change?", vbInformation + vbYesNo, "System Message") = vbYes Then
'        Call Cmdsave_Click
'    End If
'End If
'
'    If CheckIfExistInFinalJEV(txt_Jevno.Text) = False Then
'        Select Case ActiveFormCaller
'            Case "frmJEVNumberingThruRCI"
'                frmJEVNumberingThruRCI.grd_details.TextMatrix(ForTheGridRowNo, 14) = txt_Jevno.Text
'                Unload Me
'            Case "frmCDCashDisbursedReport"
'                frmCDCashDisbursedReport.MSHFlexGrid1.TextMatrix(ForTheGridRowNo, 10) = txt_Jevno.Text
'                Unload Me
'            Case "frmGeneralJournalJevNumbering"
'            frmGeneralJournalJevNumbering.grd_details.TextMatrix(ForTheGridRowNo, 4) = txt_Jevno.Text
'                Unload Me
'        End Select
'    Else
'    MsgBox "JEV number already exist in the Database..", vbInformation, "System Message"
'    End If
'End Sub
'Private Function ChkEntry() As Boolean
'
'    ChkEntry = False
'    If Trim(txt_DVNo.Text) <> "" Or (txt_AlobsNo.Text <> "" Or cmbNonAlobs.Text <> "") And txt_Claimant.Text <> "" And cmbrc.Text <> "" And txt_particular.Text <> "" And txt_FundType.Text <> "" And txt_Amount.Text <> "" Then
'        If xDebit = xCredit And xDebit > 0 Then
'            If Format(xDebit, "###,##0.00") = Format(txt_Amount.Text, "###,##0.00") Then
'
'                ChkEntry = True
'            ElseIf Format(xDebit, "###,##0.00") < Format(txt_Amount.Text, "###,##0.00") Then
'                If MsgBox("Ops..! Your total Entry is less than to your Gross Amount, Are you Sure do you want to proceed?", vbInformation + vbYesNo, "System Messag") = vbYes Then
'                    ChkEntry = True
'                Else
'                    ChkEntry = False
'                End If
'            End If
'        End If
'    End If
'
'End Function
'Private Function coloraly() As Boolean
'Dim x As Integer
'    For x = 1 To MSFlexGrid1.Rows - 1
'        If MSFlexGrid1.TextMatrix(x, 2) <> "TOTAL" Then
'            If MSFlexGrid1.TextMatrix(x, 5) <> "" Then
'                If MSFlexGrid1.TextMatrix(x, 5) = "5" Then
'                    coloraly = True
'                    Exit Function
'                End If
'            End If
'        Else
'            Exit For
'        End If
'    Next x
'End Function
'
'Private Sub Cmdsave_Click()
''On Error GoTo bad
'
'Dim DRec As New ADODB.Recordset
'Dim xType As Integer
'Dim x As Integer
'    If optObR.Value = True Or optNonObR.Value = True Then
'          If ChkEntry = True Then
'            If MsgBox("Are you sure you want to save this transaction?", vbQuestion + vbYesNo) = vbYes Then
'
'
'                If opn_Coll.Value = True Then xType = CInt(opn_Coll.Tag)
'                If opn_CashDisb.Value = True Then xType = CInt(opn_CashDisb.Tag)
'                If opn_CheckDisb.Value = True Then xType = CInt(opn_CheckDisb.Tag)
'                If opn_Other.Value = True Then xType = CInt(opn_Other.Tag)
'
'
'                If CUFlag = True Then
'                    opndbaseFMIS.Execute "Update [tblAMIS_IncomingDVTrns] set [Particular]='" & Trim(Replace(txt_particular.Text, "'", "''")) & "', Claimantcode = '" & txtclaimantcode.Text & "' Where DVNo='" & Trim(txt_DVNo.Text) & "' And ActionCode=1"
'                End If
'
'                DRec.Open ("Select * FRom tblAMIS_IncomingDVTrns where DVNo='" & txt_DVNo.Text & "' and ActionCode=1"), opndbaseFMIS, adOpenStatic, adLockOptimistic
'                        If DRec.RecordCount = 0 Then
'                            If optNonObR.Value = True Then ' NONE OBR
'                                opndbaseFMIS.Execute "Insert Into tblAMIS_IncomingDVTrns (DVNo,ObrNo,FundType,RCenter,RCenterCode,Particular,GAmount,TransactionDate,UserID,Actioncode,DateTimeEntered,Continuing,NonAlobs,Claimantcode,PAout) Values ('" & txt_DVNo.Text & "','" & GetNonAlobsCode(cmbNonAlobs.ItemData(cmbNonAlobs.ListIndex)) & "'," & _
'                                "'" & txt_FundType.Text & "'," & cmbrc.ItemData(cmbrc.ListIndex) & ",0,'" & Trim(Replace(txt_particular, "'", "''")) & "'," & CCur(txt_Amount.Text) & ",'" & Format(txtDate.Text, "mm/dd/yyyy") & "','" & ActiveUserID & "',1,'" & Format(Now, "mm/dd/yyyy hh:mm:ss AMPM") & "'," & IIf(XFlag, 1, 0) & "," & optNonObR.Tag & ",'" & txtclaimantcode.Text & "',1)"
'                            Else 'WITH OBR
'                                    opndbaseFMIS.Execute "Insert Into tblAMIS_IncomingDVTrns (DVNo,ObrNo,FundType,RCenter,RCenterCode,Particular,GAmount,TransactionDate,UserID,Actioncode,DateTimeEntered,Continuing,NonALobs,claimantcode,PAout) Values ('" & txt_DVNo.Text & "','" & Trim(txt_AlobsNo.Text) & "'," & _
'                                "'" & txt_FundType.Text & "'," & cmbrc.ItemData(cmbrc.ListIndex) & ",'" & Mid(txt_AlobsNo.Text, 5, 4) & "','" & Trim(Replace(txt_particular, "'", "''")) & "'," & CCur(txt_Amount.Text) & ",'" & Format(txtDate.Text, "mm/dd/yyyy") & "','" & Trim(ActiveUserID) & "',1,'" & Format(Now, "mm/dd/yyyy hh:mm:ss AMPM") & "'," & IIf(XFlag, 1, 0) & "," & optObR.Tag & ",'" & txtclaimantcode.Text & "',1)"
'                            End If
'                        End If
'                DRec.Close
'                Set DRec = Nothing
'
'                If xNAcode <> "" Then
'                    xObR = xNAcode
'                End If
'                opndbaseFMIS.Execute "Update tblAMIS_JournalEntry set ActionCode=2, UserID=UserID + '," & ActiveUserID & "', DateTimeEntered=DateTimeEntered + '," & Format(Now, "yyyy/mm/dd hh:mm:ss AMPM") & "' Where DVNo='" & Me.txt_DVNo.Text & "' And ActionCode=1"
'                For x = 1 To MSFlexGrid1.Rows - 1
'                    If MSFlexGrid1.TextMatrix(x, 3) <> "TOTAL" Then
'                        If MSFlexGrid1.TextMatrix(x, 1) <> "" Then
'                            If MSFlexGrid1.TextMatrix(x, 4) <> "" Or MSFlexGrid1.TextMatrix(x, 5) <> "" Then
'opndbaseFMIS.Execute "Insert Into tblAMIS_JournalEntry (TransType,DVNo,ObrNo,FmisAccntCode,Amount,DebitCredit,TransDate,UserID,Actioncode,DateTimeEntered,ApprovedByID,DateTimeApproved,LogOutBy,LogOutDateTime,LogOutRemark,Continuing) values (" & xType & ",'" & txt_DVNo.Text & "','" & txt_AlobsNo.Text & "'," & CLng(MSFlexGrid1.TextMatrix(x, 1)) & "," & CCur(IIf(IsNumeric(MSFlexGrid1.TextMatrix(x, 4)), MSFlexGrid1.TextMatrix(x, 4), 0)) + CCur(IIf(IsNumeric(MSFlexGrid1.TextMatrix(x, 5)), MSFlexGrid1.TextMatrix(x, 5), 0)) & "," & IIf(Trim(MSFlexGrid1.TextMatrix(x, 4)) = "", 0, 1) & ",'" & Format(Now, "yyyy/mm/dd hh:mm:ss AMPM") & "','" & Trim(ActiveUserID) & "'," & MSFlexGrid1.TextMatrix(x, 6) & ",'" & Format(Now, "yyyy/mm/dd hh:mm:ss AMPM") & "','" & Trim(Approvedid) & "', '" & Trim(Datetimeapproved) & "', '" & Trim(Logoutby) & "', '" & Trim(Logoutdatetime) & "', '" & Trim(Logoutremark) & "', '" & Trim(Continuing) & "')"
'                            End If
'                        End If
'                    Else
'                        Exit For
'                    End If
'                Next x
'                MsgBox "Successfully updated...!", vbInformation, "System Message"
'                Edited = False
'                'Call Toolbar1_ButtonClick(Toolbar1.Buttons.Item(1))
'            End If
'        Else
'            MsgBox "Save operation cancelled!" & vbCrLf & vbCrLf & "Please check your entry.", vbExclamation + vbOKOnly
'        End If
'    Else
'    MsgBox "Please Select Transaction type", vbInformation, "System Message"
'    End If
'Exit Sub
'bad:
'MsgBox err.Description
'End Sub
'Private Sub LoadClaimantCODE(ByVal Claimant As Variant)
''Dim opnDetails As New ADODB.Recordset
''Dim opnDetails1 As New ADODB.Recordset
''
''Select Case Classification
''Case "Individual", "Company", "National", "BarangayTreasurer", "MunicipalTreasurer"
''opnDetails.Open "Select * from tblCMS_CDClaimantDetails where lastname = '" & Claimant & "'%", opndbaseFMIS, adOpenStatic, adLockOptimistic
''If opnDetails.RecordCount <> 0 Then
''    LoadClaimantCODE = opnDetails!ClaimantCode
''Else
''    opnDetails1.Open "Select * from employee where firstname like '" & Claimant & "%'", opndbasePMIS, adOpenStatic, adLockOptimistic
''    If opnDetails.RecordCount <> 0 Then
''    LoadClaimantCODE = opnDetails1!SwipEmployeeID
''    End If
''End If
''opnDetails.Close
''Set opnDetails = Nothing
'End Sub
'
'Private Sub Command3_Click()
'CUFlag = True
'End Sub
'
'Private Sub Command4_Click()
'CAedit = True
'IfCAedit
'End Sub
'
'Private Sub Command6_Click()
'fmeCA.Visible = False
'End Sub
'
'Private Sub ctlblink1_Blinked()
'
'End Sub
'
'Private Sub fmeCA_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'Label22.FontBold = False
'Label22.FontUnderline = False
'End Sub
'
''Private Sub Command2_Click()
''Dim xType As Integer
''        If MsgBox("Are you sure you want to save this transaction?", vbQuestion + vbYesNo) = vbYes Then
''
''            If opn_Coll.Value = True Then xType = CInt(opn_Coll.Tag)
''            If opn_CashDisb.Value = True Then xType = CInt(opn_CashDisb.Tag)
''            If opn_CheckDisb.Value = True Then xType = CInt(opn_CheckDisb.Tag)
''            If opn_Other.Value = True Then xType = CInt(opn_Other.Tag)
''
''            For x = 1 To MSFlexGrid1.Rows - 1
''                If MSFlexGrid1.TextMatrix(x, 0) <> "" Then
''                    If MSFlexGrid1.TextMatrix(x, 1) <> "" Then
''
''                            opndbaseFMIS.Execute "update tblAMIS_JournalEntry set FmisAccntCode = '" & MSFlexGrid1.TextMatrix(x, 1) & "',transtype = " & xType & " where trnno = '" & MSFlexGrid1.TextMatrix(x, 0) & "'"
''
''                    End If
''                Else
''                    Exit For
''                End If
''            Next x
''        End If
''        MsgBox "Successfully Update", vbInformation, "System Message"
''    End Sub
'
'Private Sub Form_Load()
'Me.Left = (Screen.Width - Me.Width) / 2
'Me.Top = (Screen.Height - Me.Height) / 2
'Call SetGrid
'LoadOffice
'LoadTrans
'End Sub
'Private Function CAClear()
'txtCamount.Text = ""
'txtCChecdate.Text = ""
''txtCCheckno.Text = ""
'txtCClaimant.Text = ""
'txtCParticular.Text = ""
'End Function
'
'Private Function IfCAedit()
'If CAedit = True Then
'    txtCDvno.BackColor = &HFFFFFF
'    txtCDvno.Locked = False
'    'txtctotalAmnt.Text = ""
'    txtCCheckno.BackColor = &HFFFFFF
'    txtCCheckno.Locked = False
'    CAClear
'Else
'    txtCCheckno.BackColor = &H80000004
'    txtCCheckno.Locked = True
'    txtCDvno.BackColor = &H80000004
'    txtCDvno.Locked = True
'End If
'End Function
'Public Function Ifliquidition()
'If Trim(cmbNonAlobs.Text) = "Liquidation of Cash Advance" Then
'    Label22.Visible = True
'    ListView1.Visible = True
'    Call AllLoadCAdetails(ListView1, txt_DVNo.Text, txtctotalAmnt)
'End If
'End Function
'Public Function LoadTrans()
''If IfNew = True Then
'   If optNonObR.Value = True Then
'       cmbNonAlobs.Visible = True
'   Else
'   cmbNonAlobs.Visible = False
'    End If
''Else
''     cmbNonAlobs.Visible = False
''     End If
'LoadNonAlobs
'End Function
'
'
'Private Sub GetSum()
'Dim x As Integer
'    not_coloraly_total_debit = 0
'    not_coloraly_total_credit = 0
'     coloraly_total_credit = 0
'     coloraly_total_debit = 0
'
'    xDebit = 0
'    xCredit = 0
'    For x = 1 To MSFlexGrid1.Rows - 1
'        If MSFlexGrid1.TextMatrix(x, 1) <> "" Then
'            xDebit = xDebit + CCur(IIf(MSFlexGrid1.TextMatrix(x, 4) = "", 0, MSFlexGrid1.TextMatrix(x, 4)))
'            xCredit = xCredit + CCur(IIf(MSFlexGrid1.TextMatrix(x, 5) = "", 0, MSFlexGrid1.TextMatrix(x, 5)))
'                If Trim(MSFlexGrid1.TextMatrix(x, 6)) <> 5 Then
'                    not_coloraly_total_debit = not_coloraly_total_debit + CCur(IIf(MSFlexGrid1.TextMatrix(x, 4) = "", 0, MSFlexGrid1.TextMatrix(x, 4)))
'                    not_coloraly_total_credit = not_coloraly_total_credit + CCur(IIf(MSFlexGrid1.TextMatrix(x, 5) = "", 0, MSFlexGrid1.TextMatrix(x, 5)))
'                Else
'                    coloraly_total_debit = coloraly_total_debit + CCur(IIf(MSFlexGrid1.TextMatrix(x, 4) = "", 0, MSFlexGrid1.TextMatrix(x, 4)))
'                    coloraly_total_credit = coloraly_total_credit + CCur(IIf(MSFlexGrid1.TextMatrix(x, 5) = "", 0, MSFlexGrid1.TextMatrix(x, 5)))
'                End If
'        Else
'
'            MSFlexGrid1.TextMatrix(x, 3) = "TOTAL"
'            MSFlexGrid1.TextMatrix(x, 4) = Format(xDebit, "#,##0.00")
'            MSFlexGrid1.TextMatrix(x, 5) = Format(xCredit, "#,##0.00")
'            Exit For
'        End If
'    Next x
'
'End Sub
'Public Sub LoadOffice()
'Dim OREc As New ADODB.Recordset
'Dim x As Integer
'
'cmbrc.Clear
'
'OREc.Open ("Select distinct * FRom tblREF_AIS_Offices Order By [OfficeMedium]"), opndbaseFMIS, adOpenStatic, adLockOptimistic
'If OREc.RecordCount > 0 Then
'    For x = 1 To OREc.RecordCount
'        cmbrc.AddItem OREc![OfficeMedium]
'        cmbrc.ItemData(cmbrc.NewIndex) = OREc!fmisofficeid
'        OREc.MoveNext
'    Next x
'End If
'OREc.Close
'Set OREc = Nothing
'
'End Sub
'
'
'
'Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'Label22.FontBold = False
'Label22.FontUnderline = False
'End Sub
'
'Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'IfNew = False
'End Sub
'
'Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'Label22.FontBold = False
'Label22.FontUnderline = False
'End Sub
'
'Private Sub Label22_Click()
'fmeCA.Visible = True
'End Sub
'
'Private Sub Label22_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'Label22.FontBold = True
'Label22.FontUnderline = True
'End Sub
'
'Private Sub MaskEdBox1_Change()
'
'End Sub
'
'Private Sub MaskEdBox1_LostFocus()
'
'End Sub
'Private Sub MSFlexGrid1_Click()
'On Error GoTo bad
'    Select Case MSFlexGrid1.col
'    Case 2 'AccntCode
'        txt_entry.Visible = False
'        cmbEntry.Move MSFlexGrid1.CellLeft, MSFlexGrid1.CellTop, MSFlexGrid1.CellWidth
'        cmbEntry.Visible = True
'        If Len(Trim(MSFlexGrid1.Text)) <> 0 Then
'            cmbEntry.Text = MSFlexGrid1.Text
'        Else
'            cmbEntry.ListIndex = -1
'        End If
'    cmbEntry.SetFocus
'    Case 4 To 6 'Debit/Credit/actioncode
'        cmbEntry.Visible = False
'        txt_entry.Move MSFlexGrid1.CellLeft, MSFlexGrid1.CellTop, MSFlexGrid1.CellWidth, MSFlexGrid1.CellHeight
'        txt_entry.Visible = True
'        If Len(Trim(MSFlexGrid1.Text)) <> 0 Then
'            txt_entry.Text = MSFlexGrid1.Text
'            txt_entry.SelStart = 0
'            txt_entry.SelLength = Len(txt_entry.Text)
'        Else
'            txt_entry.Text = ""
'        End If
'        txt_entry.SetFocus
'    Case Else
'        txt_entry.Visible = False
'        cmbEntry.Visible = False
'    End Select
'    'MSFlexGrid1.SetFocus
'Exit Sub
'bad:
'    MsgBox err.Description
'End Sub
'
'
'Private Sub MSFlexGrid1_KeyPress(KeyAscii As Integer)
'Call MSFlexGrid1_Click
'End Sub
'
'Private Sub MSFlexGrid1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'Label22.FontBold = False
'Label22.FontUnderline = False
'End Sub
'
'Private Sub optNonObR_Click()
'LoadTrans
'End Sub
'
'Private Sub optObR_Click()
'LoadTrans
'End Sub
'
'Private Sub txt_DVNo_Change()
''If Len(Trim(txt_JEVNo.Text)) = 0 Then
''    txt_JEVNo.Text = SetNewJEVNo(txt_DVNo.Text, frmJEVNumberingThruRCI.DTPicker1.Year, frmJEVNumberingThruRCI.DTPicker1.Month)
''End If
'Call LoadBackDVDetails(txt_DVNo.Text)
'End Sub
'
'
'Private Sub txt_entry_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, MSFlexGrid1.col) = txt_entry.Text
'        If MSFlexGrid1.col = 4 Then
'            If Trim(txt_entry.Text) <> "" Then
'                MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 5) = ""
'            Else
'                MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 4) = ""
'            End If
'        Else
'            If Trim(txt_entry.Text) <> "" Then
'                MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 4) = ""
'            Else
'                MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 5) = ""
'            End If
'        End If
'        txt_entry.Visible = False
'
'        Call GetSum
'        txt_entry.Text = ""
'        MSFlexGrid1.SetFocus
'    End If
'
'End Sub
'
'Private Sub txt_RCenter_Click()
''cmbRC.Visible = True
'End Sub
'

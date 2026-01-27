VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{1693405E-2DC9-4248-B52F-4AC9145DA2AF}#1.0#0"; "WinXPC Engine.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Begin VB.Form frmJEVNumberingAssignment_New 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "JEV Numbering Assignment"
   ClientHeight    =   8955
   ClientLeft      =   435
   ClientTop       =   930
   ClientWidth     =   13665
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmJEVNumberingAssignment_New.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8955
   ScaleWidth      =   13665
   Begin VB.CheckBox Chk_haveDoc 
      BackColor       =   &H80000012&
      Caption         =   "No Documents attach"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   8640
      TabIndex        =   67
      Top             =   3480
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   11880
      Picture         =   "frmJEVNumberingAssignment_New.frx":076A
      Style           =   1  'Graphical
      TabIndex        =   66
      Top             =   3840
      Width           =   495
   End
   Begin VB.TextBox txt_Jevno 
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
      Left            =   8640
      MaxLength       =   18
      TabIndex        =   65
      Top             =   3840
      Width           =   3180
   End
   Begin VB.CommandButton cmd_post 
      Caption         =   "Post"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1125
      Left            =   12480
      Picture         =   "frmJEVNumberingAssignment_New.frx":08B4
      Style           =   1  'Graphical
      TabIndex        =   64
      Top             =   3480
      Width           =   1035
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
      Left            =   120
      TabIndex        =   53
      Top             =   4200
      Width           =   2055
   End
   Begin VB.Frame frmTrans 
      BackColor       =   &H80000012&
      Caption         =   "Transaction type"
      Enabled         =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   6840
      TabIndex        =   31
      Top             =   120
      Width           =   2775
      Begin VB.OptionButton optNonObR 
         BackColor       =   &H80000012&
         Caption         =   "Non-ObR"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1440
         TabIndex        =   33
         Tag             =   "1"
         Top             =   360
         Width           =   1215
      End
      Begin VB.OptionButton optObR 
         BackColor       =   &H80000012&
         Caption         =   "With ObR"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Tag             =   "0"
         Top             =   360
         Value           =   -1  'True
         Width           =   1215
      End
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
      Height          =   2190
      Left            =   105
      TabIndex        =   0
      Top             =   1065
      Width           =   13395
      Begin VB.CommandButton Command9 
         Caption         =   "..."
         Height          =   375
         Left            =   12840
         TabIndex        =   62
         ToolTipText     =   "Click here to change Fundtype"
         Top             =   570
         Width           =   375
      End
      Begin VB.CommandButton Command8 
         Caption         =   "..."
         Height          =   375
         Left            =   12840
         TabIndex        =   61
         ToolTipText     =   "Click here to edit particulars..."
         Top             =   1410
         Width           =   375
      End
      Begin VB.ComboBox cmbrc 
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
         Left            =   5160
         Style           =   2  'Dropdown List
         TabIndex        =   27
         Top             =   585
         Width           =   4335
      End
      Begin VB.CommandButton cmd 
         Caption         =   "...."
         Height          =   375
         Left            =   4560
         TabIndex        =   30
         Top             =   1320
         Width           =   375
      End
      Begin VB.TextBox txt_Claimant 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
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
         Top             =   1305
         Width           =   4380
      End
      Begin VB.TextBox txt_particular 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
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
         TabIndex        =   3
         Top             =   1305
         Width           =   4290
      End
      Begin VB.TextBox txt_Amount 
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
         Left            =   9870
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   1425
         Width           =   2940
      End
      Begin VB.TextBox txtclaimantcode 
         Height          =   360
         Left            =   1080
         TabIndex        =   29
         Top             =   1800
         Visible         =   0   'False
         Width           =   570
      End
      Begin VB.CommandButton Command3 
         Caption         =   "&e"
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
         Left            =   9000
         TabIndex        =   28
         Top             =   1380
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.ComboBox cmbNonAlobs 
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
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   34
         Top             =   585
         Width           =   4380
      End
      Begin VB.TextBox txt_AlobsNo 
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
         Left            =   120
         TabIndex        =   4
         Top             =   585
         Width           =   4380
      End
      Begin VB.ComboBox cmb_fundtype 
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
         Left            =   9840
         Style           =   2  'Dropdown List
         TabIndex        =   63
         Top             =   600
         Visible         =   0   'False
         Width           =   3015
      End
      Begin VB.TextBox txt_FundType 
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
         Left            =   9870
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   585
         Width           =   2940
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Responsibility Center"
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
         Left            =   5220
         TabIndex        =   11
         Top             =   270
         Width           =   1785
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Claimant"
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
         Left            =   180
         TabIndex        =   10
         Top             =   1050
         Width           =   870
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Alobs/OBR No:"
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
         Left            =   180
         TabIndex        =   9
         Top             =   270
         Width           =   1260
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Particular"
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
         Left            =   5220
         TabIndex        =   8
         Top             =   1050
         Width           =   1155
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Amount (Gross)"
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
         Left            =   9900
         TabIndex        =   7
         Top             =   1170
         Width           =   1350
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fund Type"
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
         Left            =   9900
         TabIndex        =   6
         Top             =   270
         Width           =   900
      End
   End
   Begin VB.TextBox txtDate 
      Appearance      =   0  'Flat
      Height          =   360
      Left            =   9720
      Locked          =   -1  'True
      TabIndex        =   25
      Top             =   465
      Width           =   2565
   End
   Begin VB.CommandButton cmdsave 
      Caption         =   "&Update"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5400
      Picture         =   "frmJEVNumberingAssignment_New.frx":43AE
      Style           =   1  'Graphical
      TabIndex        =   23
      ToolTipText     =   "Saves to Journal Entry"
      Top             =   240
      Width           =   975
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7800
      Top             =   960
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
            Picture         =   "frmJEVNumberingAssignment_New.frx":46F0
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJEVNumberingAssignment_New.frx":4B42
            Key             =   ""
         EndProperty
      EndProperty
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
      Height          =   4215
      Left            =   90
      ScaleHeight     =   4185
      ScaleWidth      =   12135
      TabIndex        =   17
      Top             =   4560
      Width           =   12165
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   4200
         Left            =   0
         TabIndex        =   18
         Top             =   0
         Width           =   12150
         _ExtentX        =   21431
         _ExtentY        =   7408
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
         TabIndex        =   22
         Top             =   0
         Visible         =   0   'False
         Width           =   1305
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
         TabIndex        =   21
         Text            =   "cmbEntry"
         Top             =   600
         Visible         =   0   'False
         Width           =   1575
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
         Height          =   4215
         Left            =   0
         TabIndex        =   35
         Top             =   0
         Visible         =   0   'False
         Width           =   12135
         Begin VB.TextBox txtCclaimantcode 
            Height          =   375
            Left            =   3960
            TabIndex        =   60
            Top             =   840
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.TextBox txtCObrno 
            Alignment       =   2  'Center
            BackColor       =   &H80000000&
            Height          =   360
            Left            =   1200
            Locked          =   -1  'True
            TabIndex        =   58
            Top             =   1800
            Width           =   2535
         End
         Begin VB.CommandButton Command7 
            Caption         =   "New"
            Height          =   735
            Left            =   9600
            Picture         =   "frmJEVNumberingAssignment_New.frx":4E5C
            Style           =   1  'Graphical
            TabIndex        =   56
            Top             =   240
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.CommandButton Command5 
            Caption         =   "Edit"
            Height          =   735
            Left            =   10440
            Picture         =   "frmJEVNumberingAssignment_New.frx":529E
            Style           =   1  'Graphical
            TabIndex        =   55
            Top             =   240
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.CommandButton Command4 
            Caption         =   "&Hide"
            Height          =   735
            Left            =   11280
            Picture         =   "frmJEVNumberingAssignment_New.frx":56E0
            Style           =   1  'Graphical
            TabIndex        =   54
            Top             =   240
            Width           =   735
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
            Left            =   6840
            Locked          =   -1  'True
            TabIndex        =   44
            Top             =   1320
            Width           =   2655
         End
         Begin VB.CommandButton Command6 
            Caption         =   "&Hide"
            Height          =   495
            Left            =   12960
            Picture         =   "frmJEVNumberingAssignment_New.frx":5C6A
            TabIndex        =   43
            Top             =   240
            Width           =   975
         End
         Begin VB.TextBox txtCClaimant 
            BackColor       =   &H80000000&
            Height          =   360
            Left            =   6840
            Locked          =   -1  'True
            TabIndex        =   42
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
            Left            =   6840
            Locked          =   -1  'True
            TabIndex        =   41
            Top             =   840
            Width           =   2655
         End
         Begin VB.TextBox txtCParticular 
            BackColor       =   &H80000000&
            Height          =   600
            Left            =   5040
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   40
            Top             =   1800
            Width           =   6855
         End
         Begin VB.TextBox txtCChecdate 
            Alignment       =   2  'Center
            BackColor       =   &H80000000&
            Height          =   360
            Left            =   1200
            Locked          =   -1  'True
            TabIndex        =   39
            Top             =   1320
            Width           =   2535
         End
         Begin VB.TextBox txtCCheckno 
            Alignment       =   2  'Center
            BackColor       =   &H80000000&
            Height          =   360
            Left            =   1200
            Locked          =   -1  'True
            TabIndex        =   38
            Top             =   840
            Width           =   2535
         End
         Begin VB.CommandButton cmdadd 
            Caption         =   "&Add"
            Height          =   375
            Left            =   3840
            TabIndex        =   37
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
            TabIndex        =   36
            Top             =   360
            Width           =   2535
         End
         Begin MSComctlLib.ListView ListView1 
            Height          =   1575
            Left            =   120
            TabIndex        =   57
            Top             =   2520
            Width           =   11895
            _ExtentX        =   20981
            _ExtentY        =   2778
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
         Begin VB.Label Label2 
            Caption         =   "Obr No."
            Height          =   375
            Left            =   120
            TabIndex        =   59
            Top             =   1800
            Width           =   1095
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
            Left            =   5280
            TabIndex        =   51
            Top             =   1440
            Width           =   1575
         End
         Begin VB.Label Label20 
            Caption         =   "Claimant:"
            Height          =   375
            Left            =   5880
            TabIndex        =   50
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
            Left            =   5880
            TabIndex        =   49
            Top             =   960
            Width           =   975
         End
         Begin VB.Label Label18 
            Caption         =   "Particular:"
            Height          =   375
            Left            =   3960
            TabIndex        =   48
            Top             =   1920
            Width           =   1095
         End
         Begin VB.Label Label17 
            Caption         =   "Checkdate:"
            Height          =   375
            Left            =   120
            TabIndex        =   47
            Top             =   1320
            Width           =   1095
         End
         Begin VB.Label Label16 
            Caption         =   "Checkno:"
            Height          =   375
            Left            =   120
            TabIndex        =   46
            Top             =   840
            Width           =   975
         End
         Begin VB.Label Label15 
            Caption         =   "Dvno:"
            Height          =   375
            Left            =   120
            TabIndex        =   45
            Top             =   360
            Width           =   615
         End
      End
   End
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
      Left            =   105
      TabIndex        =   12
      Top             =   3255
      Width           =   8070
      Begin VB.OptionButton opn_CheckDisb 
         Caption         =   "Check Disbursement"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   165
         Style           =   1  'Graphical
         TabIndex        =   15
         Tag             =   "02"
         Top             =   300
         UseMaskColor    =   -1  'True
         Width           =   2580
      End
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
      Begin VB.OptionButton opn_CashDisb 
         Caption         =   "Cash Disbursement"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   2760
         Style           =   1  'Graphical
         TabIndex        =   14
         Tag             =   "03"
         Top             =   300
         Width           =   2580
      End
      Begin VB.OptionButton opn_Other 
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
         Height          =   420
         Left            =   5325
         Style           =   1  'Graphical
         TabIndex        =   13
         Tag             =   "04"
         Top             =   300
         Width           =   2580
      End
   End
   Begin VB.TextBox txt_DVNo 
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
      Left            =   180
      TabIndex        =   19
      Top             =   360
      Width           =   5085
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
      Left            =   120
      TabIndex        =   24
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
   Begin VB.Shape Shape3 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   4695
      Left            =   12360
      Top             =   4320
      Width           =   1335
   End
   Begin VB.Label Label22 
      Caption         =   "<<--Cash Advance Details"
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
      Left            =   5280
      TabIndex        =   52
      Top             =   4200
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date Prepared"
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   9720
      TabIndex        =   26
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DV Number:"
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
      Left            =   180
      TabIndex        =   20
      Top             =   75
      Width           =   1050
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   1110
      Left            =   8520
      Top             =   3360
      Width           =   6615
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   960
      Left            =   0
      Top             =   0
      Width           =   13620
   End
End
Attribute VB_Name = "frmJEVNumberingAssignment_New"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Option Explicit

Dim Edited As Boolean
Dim xDebit As Currency
Dim xCredit As Currency
Dim xObR As String
Dim xNAcode, DVN As String
Dim CUFlag As Boolean           'Claimant Update Flag
Dim XFlag As Boolean
Dim rcedit As Boolean
Dim CAedit As Boolean
Public IfNew As Boolean
Public isfrom_jevNumbering As Boolean
Dim ifsaveamount As Boolean
Dim ifColoraly, SaveOk As Boolean
Public Ttype As Integer
Public fundcode As Long
Public FundType, claimantname As String
Public EditCount, IsSaveAccntng As Boolean
Dim not_coloraly_total_debit, not_coloraly_total_credit, coloraly_total_debit, coloraly_total_credit As Double
Public ptv, Approvedid, Datetimeapproved, Logoutby, Logoutdatetime, Logoutremark, Continuing As String

Public whatfield As String, Date_ As String, RCI As String, checkno As String, Particular As String, jevno As String, _
 ClaimantCode As String, FmisAccountcode As String, Gamount As Currency, Debit As Currency _
, Credit As Currency, Transtype As Integer, FmisVoucherno As String, dvno As String, obrno As String, FTYPE As String, _
 RCenter As String, OOE As String, RDOno As String, RefNo As String, Jevseries As Long, jevdate As Date, ptvNo As String, Uno As Long, ishaveDOC As Integer

Private Sub LoadBackDVDetails(ByVal dvno As String)
Dim opnDV As New ADODB.Recordset
Dim y As String
opnDV.Open "Select top 1 * from tblAMIS_IncomingDVTrns where DVNo='" & dvno & "' and actioncode=1 order by trnno desc", opndbaseFMIS, adOpenStatic, adLockOptimistic
If opnDV.RecordCount <> 0 Then
    LoadTrans
    txt_FundType.Text = opnDV!FundType
    txtclaimantcode.Text = IIf(IsNull(opnDV!ClaimantCode), "N/A", opnDV!ClaimantCode)
    cmbrc.Text = GetOfficeName(opnDV![RCenter], "OfficeMedium")
    
    txt_Claimant.Text = IIf(IsNull(opnDV!ClaimantCode), getClaimantBYdvno(dvno), GetClaimantDetails(IIf(IsNull(opnDV!ClaimantCode), "", opnDV!ClaimantCode), "Name"))
   y = getClaimantBYdvno(dvno)
    txt_particular.Text = opnDV!Particular
    txt_Amount.Text = Format(opnDV!Gamount, "#,##0.00")
    xNAcode = opnDV!obrno
    Chk_haveDoc.Value = ishaveDOC
    If opnDV!NonAlobs = 1 Then
    optNonObR.Value = True
    cmbNonAlobs.Text = GetNonAlobsName(opnDV!obrno)
    
    Else
    optObR.Value = True
    txt_AlobsNo.Text = opnDV!obrno
    End If
    If Trim(cmbNonAlobs.Text) = "Liquidation of Cash Advance" Then
        Call AllLoadCAdetails(ListView1, txt_DVNo.Text, txtctotalAmnt)
        Label22.Visible = True
        fmeCA.Visible = True
        
    Else
        fmeCA.Visible = False
        Label22.Visible = False
    End If
    SelectTrnType (Ttype)
Else
    txt_AlobsNo.Text = ""
    txt_FundType.Text = ""
    txt_Claimant.Text = ""
    txt_particular.Text = ""
    txt_Amount.Text = ""
    optNonObR.Value = False
    optObR.Value = False
    Call SetGrid
End If
opnDV.Close
Set opnDV = Nothing
Call LoadOtherDetails(dvno)
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
Private Sub LoadOtherDetails(ByVal DV As String)
Dim opnDV As New ADODB.Recordset

opnDV.Open "Select * from tblAMIS_JournalEntry where DVNo='" & DV & "' and (actioncode=1 or actioncode=5) ", opndbaseFMIS, adOpenStatic, adLockOptimistic
MSFlexGrid1.Cols = 7
MSFlexGrid1.TextMatrix(0, 0) = "trnno"
    MSFlexGrid1.TextMatrix(0, 1) = "FMISCode"
    MSFlexGrid1.TextMatrix(0, 2) = "Account Code"
    MSFlexGrid1.TextMatrix(0, 3) = "Accounts and Explanation"
    MSFlexGrid1.TextMatrix(0, 4) = "Debit"
    MSFlexGrid1.TextMatrix(0, 5) = "Credit"
'    MSFlexGrid1.TextMatrix(0, 6) = "Actioncode"
    
If opnDV.RecordCount <> 0 Then
    Call SelectTrnType(IIf(IsNull(Ttype), 3, Ttype))
    
    Approvedid = IIf(IsNull(opnDV.Fields!ApprovedByID), "", opnDV.Fields!ApprovedByID)
    Datetimeapproved = IIf(IsNull(opnDV.Fields!Datetimeapproved), "", opnDV.Fields!Datetimeapproved)
    Logoutby = IIf(IsNull(opnDV.Fields!Logoutby), "", opnDV.Fields!Logoutby)
    Logoutremark = IIf(IsNull(opnDV.Fields!Logoutremark), "", opnDV.Fields!Logoutremark)
    Logoutdatetime = IIf(IsNull(opnDV.Fields!Logoutdatetime), "", opnDV.Fields!Logoutdatetime)
    Continuing = IIf(IsNull(opnDV.Fields!Continuing), "", opnDV.Fields!Continuing)
    txtDate.Text = IIf(IsNull(opnDV.Fields!datetimeentered), "", opnDV.Fields!datetimeentered)
    'Call LoadAcctngEntries(txt_DVNo.Text)
    Call GetAccntngEntries
    Call GetSum
Else
    Approvedid = ""
    Datetimeapproved = ""
    Logoutby = ""
    Logoutremark = ""
    Logoutdatetime = ""
    Continuing = ""
    Call GetAccntngEntries
    Call GetSum
    'Call ClearAllOption
End If
opnDV.Close
Set opnDV = Nothing

End Sub
'Private Sub SetGrid()
'Dim cc As Integer
'
'    MSFlexGrid1.Clear
'    MSFlexGrid1.Rows = 50
'    MSFlexGrid1.Cols = 6 ' IIf(LCase(Trim(lblMode)) = "edit", 6, 5)
'
'    MSFlexGrid1.TextMatrix(0, 1) = "Account Code"
'    MSFlexGrid1.TextMatrix(0, 2) = "Accounts and Explanation"
'    MSFlexGrid1.TextMatrix(0, 3) = "Debit"
'    MSFlexGrid1.TextMatrix(0, 4) = "Credit"
'    MSFlexGrid1.TextMatrix(0, 5) = "ActionCode"
'
'    MSFlexGrid1.ColWidth(0) = 0
'    MSFlexGrid1.ColWidth(1) = 1500
'    MSFlexGrid1.ColWidth(2) = 6550
'    MSFlexGrid1.ColWidth(3) = 1500
'    MSFlexGrid1.ColWidth(4) = 1500
'    MSFlexGrid1.ColWidth(5) = 0
'    For cc = 0 To MSFlexGrid1.Cols - 1
'        MSFlexGrid1.Row = 0
'        MSFlexGrid1.col = cc
'        MSFlexGrid1.CellAlignment = 4
'    Next cc
'End Sub
'Public Function LoadAcctngEntries(ByVal DVNo As String)
'Dim DRec As New ADODB.Recordset
'Dim x As Integer
'Call SetGrid
'    'DRec.Close
'    DRec.Open ("Select left(ChildAccountcode,3) as childcode,sum(Debit) as sumdebit,sum(credit) as sumcredit From tblAMIS_AccoutingEntries Where [reffno]='" & DVNo & "' And (ActionCode=1) group by reffno,actioncode,left(ChildAccountcode,3)"), opndbaseFMIS, adOpenStatic, adLockOptimistic
'    If DRec.RecordCount > 0 Then
'        For x = 1 To DRec.RecordCount
''            MSFlexGrid1.TextMatrix(x, 0) = DRec![Trnno]
'            MSFlexGrid1.TextMatrix(x, 1) = DRec!childcode
'            MSFlexGrid1.TextMatrix(x, 2) = GetAccountNameByAccountcode(DRec!childcode)
'            MSFlexGrid1.TextMatrix(x, 4) = IIf((Format(DRec!sumCredit, "#,##0.00") = "0.00"), "", Format(DRec!sumCredit, "#,##0.00"))
'            MSFlexGrid1.TextMatrix(x, 3) = IIf((Format(DRec!sumDebit, "#,##0.00") = "0.00"), "", Format(DRec!sumDebit, "#,##0.00"))
'
'           ' If LCase(Trim(lblMode.Caption)) = "edit" Then MSFlexGrid1.TextMatrix(x, 5) = DRec!ActionCode  ' for coloraly purpose
'            DRec.MoveNext
'        Next x
'        Call GetSum
'    End If
'    DRec.Close
'    Set DRec = Nothing
'End Function
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
Public Sub GetAccntngEntries()
Dim Drec As New ADODB.Recordset
Dim x As Integer
Call SetGrid
    'DRec.Close
    If IsSaveAccntng = False Then
        Set Drec = opndbaseFMIS.Execute("Select left(ChildAccountcode,3) as childcode,sum(Debit) as sumdebit,sum(credit) as sumcredit From tblAMIS_AccoutingEntries Where [reffno]='" & txt_DVNo.Text & "' And (ActionCode=1) group by reffno,actioncode,left(ChildAccountcode,3)")
        If Drec.RecordCount > 0 Then
            For x = 1 To Drec.RecordCount
    '            MSFlexGrid1.TextMatrix(x, 0) = DRec![Trnno]
                MSFlexGrid1.TextMatrix(x, 1) = Drec!childcode
                MSFlexGrid1.TextMatrix(x, 2) = GetAccountNameByAccountcode(Drec!childcode)
                MSFlexGrid1.TextMatrix(x, 4) = IIf((Format(Drec!sumCredit, "#,##0.00") = "0.00"), "", Format(Drec!sumCredit, "#,##0.00"))
                MSFlexGrid1.TextMatrix(x, 3) = IIf((Format(Drec!sumDebit, "#,##0.00") = "0.00"), "", Format(Drec!sumDebit, "#,##0.00"))
              MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1
               ' If LCase(Trim(lblMode.Caption)) = "edit" Then MSFlexGrid1.TextMatrix(x, 5) = DRec!ActionCode  ' for coloraly purpose
                Drec.MoveNext
            Next x
            
        End If
    Else
        Set Drec = opndbaseFMIS.Execute("Select left(Accountcode,3) as childcode,sum(Debit) as sumdebit,sum(credit) as sumcredit From tblAMIS_tmpjournal Where [dvno]='" & txt_DVNo.Text & "' group by Dvno,left(Accountcode,3)")
    If Drec.RecordCount > 0 Then
        For x = 1 To Drec.RecordCount
            'MSFlexGrid1.TextMatrix(x, 0) = DRec![Trnno]
            
            MSFlexGrid1.TextMatrix(x, 1) = Drec!childcode
            MSFlexGrid1.TextMatrix(x, 2) = GetAccountNameByAccountcode(Drec!childcode)
            MSFlexGrid1.TextMatrix(x, 4) = IIf((Format(Drec!sumCredit, "#,##0.00") = "0.00"), "", Format(Drec!sumCredit, "#,##0.00"))
            MSFlexGrid1.TextMatrix(x, 3) = IIf((Format(Drec!sumDebit, "#,##0.00") = "0.00"), "", Format(Drec!sumDebit, "#,##0.00"))
            MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1
            'If LCase(Trim(lblMode.Caption)) = "edit" Then MSFlexGrid1.TextMatrix(x, 5) = DRec!ActionCode  ' for coloraly purpose
            Drec.MoveNext
        Next x
    End If
    End If
    Call GetSum
    Drec.Close
    Set Drec = Nothing
End Sub
Private Function GetAccntDescription(ByVal FMISCode As Long, ByVal NeedFld As String) As String
Dim opnDesc As New ADODB.Recordset

opnDesc.Open "Select * from tblREF_AIS_ChartofAccounts where FMISAccountCode=" & FMISCode & " and active=1", opndbaseFMIS, adOpenStatic, adLockOptimistic
If opnDesc.RecordCount <> 0 Then
    Select Case NeedFld
        Case "ACCT_ENTRIES"
            If opnDesc!Accountname = opnDesc!AccountNamefull Then
                GetAccntDescription = opnDesc!Accountname
            Else
                GetAccntDescription = opnDesc!Accountname & "-" & opnDesc!AccountNamefull
            End If
        
        Case "ACCT_CODE"
            GetAccntDescription = opnDesc!childaccountcode
    End Select
End If
opnDesc.Close

End Function
Private Sub ClearAllOption()
opn_Coll.Value = False
opn_CheckDisb.Value = False
opn_CashDisb.Value = False
opn_Other.Value = False
End Sub
Private Sub SelectTrnType(ByVal TransCode As String)
Select Case TransCode
    Case 1
        opn_Coll.Value = True
    Case 2
        opn_CheckDisb.Value = True
    Case 3
        opn_CashDisb.Value = True
    Case 4
        opn_Other.Value = True
End Select
End Sub

Private Sub cmb_FundType_Click()
txt_FundType.Text = GetFundMedium(cmb_fundtype.ItemData(cmb_fundtype.ListIndex))
FTYPE = cmb_fundtype.Text
FundType = cmb_fundtype.Text
fundcode = cmb_fundtype.ItemData(cmb_fundtype.ListIndex)
'cmb_fundtype.Visible = False
End Sub

Private Sub cmbEntry_Click()
cmbEntry.SetFocus
End Sub

Private Sub cmbEntry_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If cmbEntry.ListIndex <> -1 Then
            MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2) = cmbEntry.Text
            MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 0) = cmbEntry.ItemData(cmbEntry.ListIndex)
            If MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2) = "TOTAL" Then
                MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 3) = ""
                MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 4) = ""
            ElseIf MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 4) = "" And MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 5) = "" Then
                MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 5) = ""
                MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 4) = ""
            ElseIf Val(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 4)) > 0 And Val(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 5)) > 0 Then
                MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 5) = ""
                MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 4) = ""
            End If
            MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 3) = GetAccountNameByFMISAccountCode(cmbEntry.ItemData(cmbEntry.ListIndex))
            MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 1) = cmbEntry.ItemData(cmbEntry.ListIndex)
            MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 6) = "1"
            
        ElseIf cmbEntry.Text = "" Then
            MSFlexGrid1.RemoveItem (MSFlexGrid1.Row)
        Else
        End If
        cmbEntry.Visible = False
        MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1
        Call GetSum
        MSFlexGrid1.SetFocus
        Edited = True
    Else
        KeyAscii = AutoFind(cmbEntry, KeyAscii, True)
    End If

End Sub

Private Sub LoadNonAlobs()
Dim NRec As New ADODB.Recordset
Dim x As Integer

cmbNonAlobs.Clear

NRec.Open ("Select * From tblCMS_CDNoneAlobs Order By NonAlobs"), opndbaseFMIS, adOpenStatic, adLockOptimistic
If NRec.RecordCount > 0 Then
    For x = 1 To NRec.RecordCount
        cmbNonAlobs.AddItem NRec!NonAlobs
        cmbNonAlobs.ItemData(cmbNonAlobs.NewIndex) = NRec!Trnno
        NRec.MoveNext
    Next x
End If
NRec.Close
Set NRec = Nothing

End Sub

Private Sub cmbNonAlobs_Change()
Ifliquidition
End Sub

Private Sub cmbRC_Click()

'    If Trim(cmbRC.Text) <> "" Then
'        txt_RCenter.Text = Trim(cmbRC.Text)
'        txt_RCenter.Visible = True
'        cmbRC.Visible = False
'    End If
End Sub

Private Sub cmd_Click()
CUFlag = True
ActiveFormCaller = "frmJEVNumberingAssignment_New"
frmCDClaimantRegistry.Show 1
txt_Claimant.Text = CM
txtclaimantcode.Text = cc
End Sub

Private Sub cmd_post_Click()
Dim cc, tmp As Integer
If MsgBox("Save first?", vbQuestion + vbYesNo, "System Confirmation") = vbYes Then
    Call cmdSave_Click
End If
If SaveOk = True Then
    If MsgBox("Are you sure do you want to Post the transaction.?", vbQuestion + vbYesNo, "System Confirmation") = vbYes Then
        If JevOk = True Then
                tmp = ExtractJEVSNo(txt_Jevno.Text)
                jevdate = DatePost
                Jevseries = tmp
                    If Len(Trim(txt_Jevno)) > 0 Then
                             If IsFormatCorrect(txt_Jevno.Text) = True Then
                                Call GEtCompleteJEVDetails(Trim(txt_DVNo.Text), whatfield, Date_, RCI, checkno _
                                , txt_particular.Text, txt_Jevno.Text, ClaimantCode, FmisAccountcode, Gamount, Debit, Credit, Ttype, FmisVoucherno, dvno, obrno, FTYPE, RCenter _
                                , OOE, RDOno, RefNo, Jevseries, jevdate, ptvNo)
                                If Chk_haveDoc.Value = 1 Then
                                    opndbaseFMIS.Execute "update tblAMIS_FinalJEV set HaveDoc = 1 where jevno = '" & txt_Jevno.Text & "' and actioncode = 1"
                                Else
                                opndbaseFMIS.Execute "update tblAMIS_FinalJEV set HaveDoc = 0 where jevno = '" & txt_Jevno.Text & "' and actioncode = 1"
                                End If
                                'Updating table from PTO....
                                If Ttype = 2 Then
                                opndbaseFMIS.Execute "Update tblCMS_CDRCIReport set AlreadySaved2JEV=1,DatePostedtoJEV='" & Date & "',PostedtoJEVUserid='" & Trim(ActiveUserID) & "' where trnno=" & Uno & ""
                                ElseIf Ttype = 3 Then
                                    opndbaseFMIS.Execute "Update tblCMS_CDCashBook set AlreadySaved2JEV=1,DatePostedtoJEV='" & Date & "',PostedtoJEVUserid='" & Trim(ActiveUserID) & "' where trnno=" & Uno & ""
                                End If
                                'Updating Accounting REcord...
                                opndbaseFMIS.Execute "update tblAMIS_JournalEntry set JEVNo='" & txt_Jevno.Text & "', " & _
                                    " JEVSeriesNo=" & tmp & ",JEVBy='" & ActiveUserID & "', " & _
                                    " JEVDate='" & DatePost & "',transtype = " & Ttype & " where DVNo='" & txt_DVNo.Text & "'"
                            End If
                    End If
        JevOk = False
        Else
            Exit Sub
        End If
    'MsgBox "Posting to JEV, Successful!", vbInformation, "System Information"
    SaveOk = False
    Unload Me
    End If
End If
End Sub

Private Sub cmdadd_Click()
Dim x
If IfexistDv(txtCDvno.Text) = False Then
    If txtCDvno.Text <> "" And txtCCheckno.Text <> "" And txtCamount.Text <> "" Then
        Set x = ListView1.ListItems.Add(, , "")
            x.SubItems(1) = txtCDvno.Text
            x.SubItems(2) = txtCCheckno.Text
            x.SubItems(3) = txtCChecdate.Text
            x.SubItems(4) = txtCParticular.Text
            x.SubItems(5) = txtCClaimant.Text
            x.SubItems(6) = txtCamount.Text
            txtctotalAmnt.Text = Format(GetCATotalamount(ListView1), "#,##0.00")
    Else
    MsgBox "Please check your entry", vbInformation, "System Message"
    End If
Else
    MsgBox "Dvno Already on the List", vbInformation, "System Message"
End If
End Sub
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
End Function

Private Sub Command1_Click()
Dim LastJEVNo As Long
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
Dim rec As New ADODB.Recordset
frmPOstdate.Show 1
If JevOk = True Then
    Set rec = opndbaseFMIS.Execute("EXEC [dbo].[Proc_GetMaxJevSeries_new] @transtype = " & Ttype & ",@jevyeardate = '" & DatePost & "' ,@fundtype = '" & FTYPE & "'")
    LastJEVNo = rec.Fields!MAXJEVSERIES
    rec.Close
    txt_Jevno.Text = fundcode & "-" & Right(Year(DatePost), 2) & "-" & Format(Month(DatePost), "00") & "-" & Format(Ttype, "00") & "-" & Format(LastJEVNo, "0000")
Else
MsgBox "Cannot Generate the System JEV Number,If you cancel to Set the Date", vbInformation, "System Message"
End If
End Sub
Private Function ChkEntry() As Boolean

    ChkEntry = False
    If Left(txt_Jevno.Text, 3) = "119" Then
        MsgBox "Please Specify what kind of fundtype of economic enterprises..", vbCritical, "System Message"
        cmb_fundtype.Visible = True
        Exit Function
    End If
    If CheckIfExistPTVinFinalJEV(txt_DVNo.Text) = True Then 'check if EXIST DV number  in Final JEV
        MsgBox "DV number is Already Posted...Please Check your Entry", vbCritical, "System Message"
        Exit Function
    End If
    If Trim(txt_DVNo.Text) <> "" And (txt_AlobsNo.Text <> "" Or cmbNonAlobs.Text <> "") And txt_Claimant.Text <> "" And cmbrc.Text <> "" And txt_particular.Text <> "" And txt_FundType.Text <> "" And txt_Amount.Text <> "" Then
        If CheckIfExistInFinalJEV(txt_Jevno.Text) = False Then 'check if EXIST JEV no  in Final JEV
            If xDebit = xCredit And xDebit > 0 Then
                If Format(xDebit, "###,##0.00") = Format(txt_Amount.Text, "###,##0.00") Then
                    
                    ChkEntry = True
                ElseIf Format(xDebit, "###,##0.00") < Format(txt_Amount.Text, "###,##0.00") Then
                    If MsgBox("Ops..! Your total Entry is less than to your Gross Amount, Are you Sure do you want to proceed?", vbInformation + vbYesNo, "System Messag") = vbYes Then
                        ChkEntry = True
                    Else
                        ChkEntry = False
                    End If
                ElseIf Format(xDebit, "###,##0.00") > Format(txt_Amount.Text, "###,##0.00") Then
                    If MsgBox("Ops..! Your total Entry is Greater than to your Gross Amount, Are you Sure this transaction is have a Corolary Entry?", vbInformation + vbYesNo, "System Messag") = vbYes Then
                        ChkEntry = True
                    Else
                        ChkEntry = False
                    End If
                
                End If
            Else
            MsgBox "Total Debit and total Credit are not Equal, Please Check The Entry...!", vbCritical, "System Information"
            End If
        Else
            MsgBox "JEV number already exist in the Database..", vbInformation, "System Message"
                If MsgBox("Do you want to Generate JEV Number?", vbInformation + vbYesNo, "System Confirmation") = vbYes Then
                  Call Command1_Click
                End If
        End If
    Else
    MsgBox "Some fields are Empty,Please Check it", vbInformation, "System Message"
    End If
    
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

Private Sub cmdSave_Click()
'On Error GoTo bad
'txt_FundType.Locked = False
Dim Drec As New ADODB.Recordset
Dim xType As Integer
Dim x As Integer
    If optObR.Value = True Or optNonObR.Value = True Then
          If ChkEntry = True Then
            If MsgBox("Are you sure you want to save this transaction?", vbQuestion + vbYesNo) = vbYes Then
                If opn_Coll.Value = True Then xType = CInt(opn_Coll.Tag)
                If opn_CashDisb.Value = True Then xType = CInt(opn_CashDisb.Tag)
                If opn_CheckDisb.Value = True Then xType = CInt(opn_CheckDisb.Tag)
                If opn_Other.Value = True Then xType = CInt(opn_Other.Tag)
               
                
                If CUFlag = True Then
                    opndbaseFMIS.Execute "Update [tblAMIS_IncomingDVTrns] set [Particular]='" & Trim(Replace(txt_particular.Text, "'", "''")) & "', Claimantcode = '" & txtclaimantcode.Text & "' Where DVNo='" & Trim(txt_DVNo.Text) & "' And ActionCode=1"
                End If
                
                Drec.Open ("Select * FRom tblAMIS_IncomingDVTrns where DVNo='" & txt_DVNo.Text & "' and ActionCode=1"), opndbaseFMIS, adOpenStatic, adLockOptimistic
                        If Drec.RecordCount = 0 Then
                            If optNonObR.Value = True Then ' NONE OBR
                                opndbaseFMIS.Execute "Insert Into tblAMIS_IncomingDVTrns (DVNo,ObrNo,FundType,RCenter,RCenterCode,Particular,GAmount,TransactionDate,UserID,Actioncode,DateTimeEntered,Continuing,NonAlobs,Claimantcode,PAout) Values ('" & txt_DVNo.Text & "','" & GetNonAlobsCode(cmbNonAlobs.ItemData(cmbNonAlobs.ListIndex)) & "'," & _
                                "'" & txt_FundType.Text & "'," & cmbrc.ItemData(cmbrc.ListIndex) & ",0,'" & Trim(Replace(txt_particular, "'", "''")) & "'," & CCur(txt_Amount.Text) & ",'" & Format(txtDate.Text, "mm/dd/yyyy") & "','" & Trim(ActiveUserID) & "',1,'" & Format(Now, "mm/dd/yyyy hh:mm:ss AMPM") & "'," & IIf(XFlag, 1, 0) & "," & optNonObR.Tag & ",'" & txtclaimantcode.Text & "',1)"
                            Else 'WITH OBR
                                    opndbaseFMIS.Execute "Insert Into tblAMIS_IncomingDVTrns (DVNo,ObrNo,FundType,RCenter,RCenterCode,Particular,GAmount,TransactionDate,UserID,Actioncode,DateTimeEntered,Continuing,NonALobs,claimantcode,PAout) Values ('" & txt_DVNo.Text & "','" & Trim(txt_AlobsNo.Text) & "'," & _
                                "'" & txt_FundType.Text & "'," & cmbrc.ItemData(cmbrc.ListIndex) & ",'" & Mid(txt_AlobsNo.Text, 5, 4) & "','" & Trim(Replace(txt_particular, "'", "''")) & "'," & CCur(txt_Amount.Text) & ",'" & Format(txtDate.Text, "mm/dd/yyyy") & "','" & Trim(ActiveUserID) & "',1,'" & Format(Now, "mm/dd/yyyy hh:mm:ss AMPM") & "'," & IIf(XFlag, 1, 0) & "," & optObR.Tag & ",'" & txtclaimantcode.Text & "',1)"
                            End If
                        End If
                Drec.Close
                
'                Call GEtCompleteJEVDetails(Trim(txt_DVNo.Text), "TempPost", Date_, RCI, checkno _
'                                , txt_particular.Text, "", ClaimantCode, FmisAccountcode, Gamount, Debit, Credit, Ttype, FmisVoucherno, dvno, obrno, FTYPE, RCenter _
'                                , OOE, RDOno, RefNo, 0, Now, ptvNo)
                                
                Set Drec = Nothing
                
                If xNAcode <> "" Then
                    xObR = xNAcode
                End If
                'opndbaseFMIS.Execute "Update tblAMIS_JournalEntry set ActionCode=2, UserID=UserID + '," & ActiveUserID & "', DateTimeEntered=DateTimeEntered + '," & Format(Now, "yyyy/mm/dd hh:mm:ss AMPM") & "' Where DVNo='" & Me.txt_DVNo.Text & "' And ActionCode=1"
                
                'opndbaseFMIS.Execute "Insert Into tblAMIS_JournalEntry (TransType,DVNo,ObrNo,FmisAccntCode,Amount,DebitCredit,TransDate,UserID,Actioncode,DateTimeEntered,ApprovedByID,DateTimeApproved,LogOutBy,LogOutDateTime,LogOutRemark,Continuing) values (" & xType & ",'" & txt_DVNo.Text & "','" & xObR & "',0,0,0,'" & Format(Now, "yyyy/mm/dd hh:mm:ss AMPM") & "','" & Trim(ActiveUserID) & "',1,'" & Format(Now, "yyyy/mm/dd hh:mm:ss AMPM") & "','" & Trim(Approvedid) & "', '" & Trim(Datetimeapproved) & "', '" & Trim(Logoutby) & "', '" & Trim(Logoutdatetime) & "', '" & Trim(Logoutremark) & "', '" & Trim(Continuing) & "')"
                If IsSaveAccntng = True Then
                    Call SaveAcctngEntries(txt_DVNo.Text)
                End If
                SaveOk = True
                'MsgBox "Successfully updated...!", vbInformation, "System Message"
                Edited = False
                'Call Toolbar1_ButtonClick(Toolbar1.Buttons.Item(1))
            End If
        Else
            'MsgBox "Save operation cancelled!" & vbCrLf & vbCrLf & "Please check your entry.", vbExclamation + vbOKOnly
        End If
    Else
    MsgBox "Please Select Transaction type", vbInformation, "System Message"
    End If
Exit Sub
bad:
MsgBox err.description
End Sub
Public Function SaveAcctngEntries(ByVal dvno As String)
Dim Drec As New ADODB.Recordset
Dim rec As New ADODB.Recordset
Dim x As Integer

    Drec.Open ("Select Accountcode,sum(Debit) as Debit ,sum(Credit) as Credit From tblAMIs_tmpjournal Where [dvno]='" & dvno & "' group by accountcode"), opndbaseFMIS, adOpenStatic, adLockOptimistic
    If Drec.RecordCount > 0 Then
        opndbaseFMIS.Execute "update tblAMIS_AccoutingEntries set actioncode =2 where reffno = '" & dvno & "' and actioncode =1" ', datetimeentered = rtrim(ltrim(DateTimeEntered)) +'," & Format(Now, "yyyy/mm/dd hh:mm:ss AMPM") & "',userid = UserID + '," & Trim(ActiveUserID) & "'
        For x = 1 To Drec.RecordCount
            DoEvents
            opndbaseFMIS.Execute "Insert into tblAMIS_AccoutingEntries (reffNo,ChildAccountcode,debit,credit,actioncode,datetimeentered,transtype,userid) values " & _
            "('" & Trim(dvno) & "','" & Trim(Drec!accountcode) & "'," & Drec!Debit & "," & Drec!Credit & ",1,'" & Format(Now, "yyyy/mm/dd hh:mm:ss AMPM") & "'," & Ttype & ",'" & Trim(ActiveUserID) & "')"
            Drec.MoveNext
        Next x
        opndbaseFMIS.Execute "delete from tblAMIs_tmpjournal where dvno = '" & dvno & "'"
    End If
    Drec.Close
    Set Drec = Nothing
End Function
Private Sub LoadClaimantCODE(ByVal Claimant As Variant)
'Dim opnDetails As New ADODB.Recordset
'Dim opnDetails1 As New ADODB.Recordset
'
'Select Case Classification
'Case "Individual", "Company", "National", "BarangayTreasurer", "MunicipalTreasurer"
'opnDetails.Open "Select * from tblCMS_CDClaimantDetails where lastname = '" & Claimant & "'%", opndbaseFMIS, adOpenStatic, adLockOptimistic
'If opnDetails.RecordCount <> 0 Then
'    LoadClaimantCODE = opnDetails!ClaimantCode
'Else
'    opnDetails1.Open "Select * from employee where firstname like '" & Claimant & "%'", opndbasePMIS, adOpenStatic, adLockOptimistic
'    If opnDetails.RecordCount <> 0 Then
'    LoadClaimantCODE = opnDetails1!SwipEmployeeID
'    End If
'End If
'opnDetails.Close
'Set opnDetails = Nothing
End Sub

Private Sub Command3_Click()
CUFlag = True
End Sub

Private Sub Command4_Click()
fmeCA.Visible = False
End Sub

Private Sub Command5_Click()
CAedit = True
IfCAedit
End Sub

Private Sub Command6_Click()
fmeCA.Visible = False
End Sub

Private Sub ctlblink1_Blinked()

End Sub

Private Sub Command8_Click()
ifsaveamount = True
txt_Amount.Locked = False
End Sub

Private Sub Command9_Click()
cmb_fundtype.Top = txt_FundType.Top
cmb_fundtype.Left = txt_FundType.Left
cmb_fundtype.Width = txt_FundType.Width
cmb_fundtype.Visible = True
End Sub

Private Sub fmeCA_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Label22.FontBold = False
Label22.FontUnderline = False
End Sub

'Private Sub Command2_Click()
'Dim xType As Integer
'        If MsgBox("Are you sure you want to save this transaction?", vbQuestion + vbYesNo) = vbYes Then
'
'            If opn_Coll.Value = True Then xType = CInt(opn_Coll.Tag)
'            If opn_CashDisb.Value = True Then xType = CInt(opn_CashDisb.Tag)
'            If opn_CheckDisb.Value = True Then xType = CInt(opn_CheckDisb.Tag)
'            If opn_Other.Value = True Then xType = CInt(opn_Other.Tag)
'
'            For x = 1 To MSFlexGrid1.Rows - 1
'                If MSFlexGrid1.TextMatrix(x, 0) <> "" Then
'                    If MSFlexGrid1.TextMatrix(x, 1) <> "" Then
'
'                            opndbaseFMIS.Execute "update tblAMIS_JournalEntry set FmisAccntCode = '" & MSFlexGrid1.TextMatrix(x, 1) & "',transtype = " & xType & " where trnno = '" & MSFlexGrid1.TextMatrix(x, 0) & "'"
'
'                    End If
'                Else
'                    Exit For
'                End If
'            Next x
'        End If
'        MsgBox "Successfully Update", vbInformation, "System Message"
'    End Sub

Private Sub Form_Load()
Me.Left = (Screen.Width - Me.Width) / 2
Me.Top = (Screen.Height - Me.Height) / 2
Call SetGrid
Call LoadFundType(cmb_fundtype)
LoadOffice
LoadTrans
SaveOk = False
EditCount = False
End Sub


Private Function CAClear()
txtCamount.Text = ""
txtCChecdate.Text = ""
'txtCCheckno.Text = ""
txtCClaimant.Text = ""
txtCParticular.Text = ""
End Function

Private Function IfCAedit()
If CAedit = True Then
    txtCDvno.BackColor = &HFFFFFF
    txtCDvno.Locked = False
    'txtctotalAmnt.Text = ""
    txtCCheckno.BackColor = &HFFFFFF
    txtCCheckno.Locked = False
    CAClear
Else
    txtCCheckno.BackColor = &H80000004
    txtCCheckno.Locked = True
    txtCDvno.BackColor = &H80000004
    txtCDvno.Locked = True
End If
End Function
Public Function Ifliquidition()
If Trim(cmbNonAlobs.Text) = "Liquidation of Cash Advance" Then
    Label22.Visible = True
    ListView1.Visible = True
    Call AllLoadCAdetails(ListView1, txt_DVNo.Text, txtctotalAmnt)
End If
End Function
Public Function LoadTrans()
'If IfNew = True Then
   If optNonObR.Value = True Then
       cmbNonAlobs.Visible = True
   Else
   cmbNonAlobs.Visible = False
    End If
'Else
'     cmbNonAlobs.Visible = False
'     End If
LoadNonAlobs
End Function
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
        Else
            MSFlexGrid1.TextMatrix(x, 2) = "TOTAL"
            MSFlexGrid1.TextMatrix(x, 3) = Format(xDebit, "#,##0.00")
            MSFlexGrid1.TextMatrix(x, 4) = Format(xCredit, "#,##0.00")
            Exit For
        End If
    Next x
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



Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Label22.FontBold = False
Label22.FontUnderline = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
IfNew = False
End Sub

Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
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

Private Sub MSFlexGrid1_Click()
If chkSC.Value = 1 Then
Call MSFlexGrid1_DblClick
End If
End Sub


Private Sub MSFlexGrid1_DblClick()
    With frmSub3
        .REFF = txt_DVNo.Text
        .Gamount = txt_Amount.Text
        .CName = UCase(txt_Claimant.Text)
        .isEdit = True
        'EditCount = False
        Set .frm = Me
        Call LoadAcctngEntries(txt_DVNo.Text)
        .Show 1
        Call GetAccntngEntries
    End With
End Sub
Public Function LoadAcctngEntries(ByVal dvno As String)
Dim Drec As New ADODB.Recordset
Dim rec As New ADODB.Recordset
Dim rec1 As New ADODB.Recordset
Dim x As Integer
    Drec.Open ("Select ChildAccountcode,Debit ,Credit From tblAMIS_AccoutingEntries Where [reffno]='" & dvno & "' And (ActionCode=1) "), opndbaseFMIS, adOpenStatic, adLockOptimistic
    If Drec.RecordCount > 0 Then
        If EditCount = False Then
            EditCount = True
            rec.Open "Select dvno from tblAMIs_tmpjournal where dvno = '" & dvno & "'", opndbaseFMIS, adOpenStatic
            If rec.RecordCount > 0 Then
                    If MsgBox("This transaction Have a temporary Accounting Entries, do you want to Delete?", vbCritical + vbYesNo, "System Information") = vbYes Then
                        opndbaseFMIS.Execute "Delete from tblAMIs_tmpjournal where Dvno = '" & dvno & "'"
                        For x = 1 To Drec.RecordCount
                        DoEvents
                            opndbaseFMIS.Execute "Insert into tblAMIs_tmpjournal (Dvno,Accountcode,Debit,Credit) values ('" & Trim(dvno) & "','" & Trim(Drec!childaccountcode) & "'," & Drec!Debit & "," & Drec!Credit & ")"
                            Drec.MoveNext
                        Next x
                    End If
            Else
'                rec1.Open "Select dvno from tblAMIs_tmpjournal where dvno = '" & DVNo & "'", opndbaseFMIS, adOpenStatic
'                    If rec1.RecordCount > 0 Then
'                        opndbaseFMIS.Execute "Delete from tblAMIs_tmpjournal where Dvno = '" & DVNo & "'"
'                    End If
'                rec1.Close
                For x = 1 To Drec.RecordCount
                DoEvents
                    opndbaseFMIS.Execute "Insert into tblAMIs_tmpjournal (Dvno,Accountcode,Debit,Credit) values ('" & Trim(dvno) & "','" & Trim(Drec!childaccountcode) & "'," & Drec!Debit & "," & Drec!Credit & ")"
                    Drec.MoveNext
                Next x
            End If
            rec.Close
        Else
'            For x = 1 To DRec.RecordCount
'            DoEvents
'                opndbaseFMIS.Execute "Insert into tblAMIs_tmpjournal (Dvno,Accountcode,Debit,Credit) values ('" & Trim(DVNo) & "','" & Trim(DRec!childaccountcode) & "'," & DRec!Debit & "," & DRec!Credit & ")"
'                DRec.MoveNext
'            Next x
        End If
    End If
    Drec.Close
    Set Drec = Nothing
End Function


Private Sub MSFlexGrid1_KeyPress(KeyAscii As Integer)
'Call MSFlexGrid1_Click
End Sub

Private Sub MSFlexGrid1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Label22.FontBold = False
Label22.FontUnderline = False
End Sub

Private Sub optNonObR_Click()
LoadTrans
End Sub

Private Sub optObR_Click()
LoadTrans
End Sub

Private Sub txt_amount_LostFocus()
txt_Amount.Locked = True
txt_Amount.Text = Format(txt_Amount.Text, "#,##0.00")
End Sub

Private Sub txt_DVNo_Change()
Call LoadBackDVDetails(txt_DVNo.Text)
End Sub


Private Sub txt_entry_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, MSFlexGrid1.col) = txt_entry.Text
        If MSFlexGrid1.col = 4 Then
            If Trim(txt_entry.Text) <> "" Then
                MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 5) = ""
            Else
                MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 4) = ""
            End If
        Else
            If Trim(txt_entry.Text) <> "" Then
                MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 4) = ""
            Else
                MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 5) = ""
            End If
        End If
        txt_entry.Visible = False
        
        Call GetSum
        txt_entry.Text = ""
        MSFlexGrid1.SetFocus
    End If
End Sub
Private Sub txt_RCenter_Click()
'cmbRC.Visible = True
End Sub


VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmIncomingTrn3 
   Caption         =   "Incoming Transaction (DV Numbering)"
   ClientHeight    =   7665
   ClientLeft      =   5070
   ClientTop       =   4695
   ClientWidth     =   13320
   Icon            =   "frmIncomingTrn.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7665
   ScaleWidth      =   13320
   Begin VB.Frame fmeDetails 
      Caption         =   "Transaction Details"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3525
      Left            =   495
      TabIndex        =   9
      Top             =   1920
      Width           =   10350
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
         Height          =   2655
         Left            =   0
         TabIndex        =   39
         Top             =   0
         Visible         =   0   'False
         Width           =   10335
         Begin VB.TextBox txtctotalAmnt 
            Alignment       =   1  'Right Justify
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
            Left            =   7560
            Locked          =   -1  'True
            TabIndex        =   57
            Top             =   1320
            Width           =   2655
         End
         Begin MSComctlLib.ListView ListView1 
            Height          =   2175
            Left            =   120
            TabIndex        =   56
            Top             =   360
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   3836
            View            =   3
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
            NumItems        =   7
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "trnno"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Voucher No."
               Object.Width           =   3528
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "checkno"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "checkdate"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "particular"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "claimant"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   6
               Text            =   "amount"
               Object.Width           =   0
            EndProperty
         End
         Begin VB.CommandButton Command4 
            Caption         =   "&Next"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   9240
            TabIndex        =   55
            Top             =   1920
            Width           =   975
         End
         Begin VB.TextBox txtCClaimant 
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
            Left            =   7560
            Locked          =   -1  'True
            TabIndex        =   53
            Top             =   360
            Width           =   2655
         End
         Begin VB.TextBox txtCamount 
            Alignment       =   1  'Right Justify
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
            Left            =   7560
            Locked          =   -1  'True
            TabIndex        =   51
            Top             =   840
            Width           =   2655
         End
         Begin VB.TextBox txtCParticular 
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
            Left            =   3360
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   49
            Top             =   1800
            Width           =   5775
         End
         Begin VB.TextBox txtCChecdate 
            Alignment       =   2  'Center
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
            Left            =   3360
            Locked          =   -1  'True
            TabIndex        =   47
            Top             =   1320
            Width           =   2535
         End
         Begin VB.TextBox txtCCheckno 
            Alignment       =   2  'Center
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
            Left            =   3360
            TabIndex        =   45
            Top             =   840
            Width           =   2535
         End
         Begin VB.CommandButton Command3 
            Caption         =   "&Add"
            Height          =   375
            Left            =   6000
            TabIndex        =   44
            Top             =   360
            Width           =   495
         End
         Begin VB.TextBox txtCDvno 
            Alignment       =   2  'Center
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
            Left            =   3360
            TabIndex        =   42
            Top             =   360
            Width           =   2535
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
            Left            =   6000
            TabIndex        =   58
            Top             =   1440
            Width           =   1575
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
            Left            =   6600
            TabIndex        =   54
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
            Left            =   6600
            TabIndex        =   52
            Top             =   960
            Width           =   975
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
            Left            =   2280
            TabIndex        =   50
            Top             =   1920
            Width           =   1095
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
            Left            =   2280
            TabIndex        =   48
            Top             =   1320
            Width           =   1095
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
            Left            =   2280
            TabIndex        =   46
            Top             =   840
            Width           =   975
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
            Left            =   2280
            TabIndex        =   43
            Top             =   360
            Width           =   615
         End
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&R"
         Height          =   255
         Left            =   9960
         TabIndex        =   40
         Top             =   3000
         Width           =   255
      End
      Begin VB.ComboBox cmbRC 
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
         Left            =   2880
         Style           =   2  'Dropdown List
         TabIndex        =   32
         Top             =   2160
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.ComboBox cmbFund 
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
         Left            =   2880
         Style           =   2  'Dropdown List
         TabIndex        =   31
         Top             =   1440
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.ComboBox cmbOOE 
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
         ItemData        =   "frmIncomingTrn.frx":076A
         Left            =   2880
         List            =   "frmIncomingTrn.frx":076C
         Style           =   2  'Dropdown List
         TabIndex        =   30
         Top             =   3000
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CommandButton btnClaimant 
         Caption         =   "..."
         Height          =   255
         Left            =   4920
         TabIndex        =   1
         ToolTipText     =   "Click here to select claimant..."
         Top             =   720
         Width           =   255
      End
      Begin VB.TextBox txtDetail 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1875
         Left            =   5505
         MultiLine       =   -1  'True
         TabIndex        =   5
         Top             =   660
         Width           =   4560
      End
      Begin VB.TextBox txtOOE 
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
         Left            =   225
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   2962
         Width           =   5160
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
         Left            =   210
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   1470
         Width           =   4560
      End
      Begin VB.TextBox txtOffice 
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
         Left            =   210
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   2265
         Width           =   4560
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
         Left            =   210
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   660
         Width           =   4560
      End
      Begin VB.TextBox txtAmount 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   5505
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   2940
         Width           =   4440
      End
      Begin VB.Label Label14 
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
         Left            =   7200
         TabIndex        =   41
         Top             =   240
         Width           =   3015
      End
      Begin VB.Label Label10 
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
         Left            =   5505
         TabIndex        =   20
         Top             =   2670
         Width           =   1380
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Particular"
         Height          =   195
         Left            =   5505
         TabIndex        =   19
         Top             =   435
         Width           =   660
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Object of Expenditure"
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
         Left            =   225
         TabIndex        =   18
         Top             =   2670
         Width           =   1905
      End
      Begin VB.Label Label7 
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
         Left            =   210
         TabIndex        =   17
         Top             =   1230
         Width           =   975
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
         Left            =   210
         TabIndex        =   16
         Top             =   1935
         Width           =   1905
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
         Left            =   210
         TabIndex        =   15
         Top             =   435
         Width           =   780
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Log Out JEV"
      Height          =   495
      Left            =   12120
      TabIndex        =   38
      Top             =   7080
      Width           =   1065
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
      ItemData        =   "frmIncomingTrn.frx":076E
      Left            =   11790
      List            =   "frmIncomingTrn.frx":0770
      TabIndex        =   35
      Top             =   2280
      Width           =   1230
   End
   Begin VB.CommandButton btnSearch 
      Caption         =   "Search"
      Height          =   495
      Left            =   10920
      TabIndex        =   34
      Top             =   7080
      Width           =   1065
   End
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
      Left            =   4800
      Style           =   2  'Dropdown List
      TabIndex        =   29
      Top             =   1200
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtRC 
      Height          =   195
      Left            =   120
      TabIndex        =   28
      Top             =   4200
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox txtOfficeCode 
      Height          =   195
      Left            =   120
      TabIndex        =   27
      Top             =   3360
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox txtClaimantCode 
      Height          =   285
      Left            =   120
      TabIndex        =   26
      Top             =   2880
      Visible         =   0   'False
      Width           =   150
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
      Height          =   435
      Left            =   8445
      Locked          =   -1  'True
      TabIndex        =   22
      Top             =   1200
      Width           =   2280
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
      ItemData        =   "frmIncomingTrn.frx":0772
      Left            =   11790
      List            =   "frmIncomingTrn.frx":0774
      TabIndex        =   14
      Top             =   1845
      Width           =   1230
   End
   Begin VB.TextBox txtDVNo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   600
      Left            =   1185
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   6510
      Width           =   8415
   End
   Begin VB.ListBox List1 
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
      Height          =   3660
      Left            =   10920
      TabIndex        =   10
      Top             =   3240
      Width           =   2235
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   840
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   13320
      _ExtentX        =   23495
      _ExtentY        =   1482
      ButtonWidth     =   1323
      ButtonHeight    =   1429
      Appearance      =   1
      Style           =   1
      ImageList       =   "itb32x32"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   11
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
            Object.Visible         =   0   'False
            Caption         =   "Edit"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Delete"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Log Out"
            ImageIndex      =   15
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Close"
            ImageIndex      =   7
         EndProperty
      EndProperty
      Begin MSComctlLib.ImageList itb32x32 
         Left            =   6000
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
               Picture         =   "frmIncomingTrn.frx":0776
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmIncomingTrn.frx":2108
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmIncomingTrn.frx":3A9A
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmIncomingTrn.frx":542C
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmIncomingTrn.frx":6DBE
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmIncomingTrn.frx":8750
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmIncomingTrn.frx":A0E2
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmIncomingTrn.frx":BA74
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmIncomingTrn.frx":D406
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmIncomingTrn.frx":ED9A
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmIncomingTrn.frx":FA76
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmIncomingTrn.frx":10356
               Key             =   ""
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmIncomingTrn.frx":11032
               Key             =   ""
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmIncomingTrn.frx":11D0E
               Key             =   ""
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmIncomingTrn.frx":129EA
               Key             =   ""
            EndProperty
            BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmIncomingTrn.frx":136C6
               Key             =   ""
            EndProperty
            BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmIncomingTrn.frx":143A2
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.TextBox txtObR 
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
      Left            =   465
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1200
      Width           =   4845
   End
   Begin VB.Frame frmTrans 
      BackColor       =   &H80000007&
      Caption         =   "Transaction type"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   735
      Left            =   5520
      TabIndex        =   59
      Top             =   960
      Width           =   2775
      Begin VB.OptionButton optObR 
         BackColor       =   &H00000000&
         Caption         =   "With ObR"
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
         Left            =   120
         TabIndex        =   61
         Top             =   360
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton optNonObR 
         BackColor       =   &H00000000&
         Caption         =   "Non-ObR"
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
         Left            =   1440
         TabIndex        =   60
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Year of:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   11160
      TabIndex        =   37
      Top             =   1920
      Width           =   555
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Month of:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   11040
      TabIndex        =   36
      Top             =   2355
      Width           =   675
   End
   Begin VB.Label lblRefresh 
      BackStyle       =   0  'Transparent
      Caption         =   "Refresh"
      Height          =   495
      Left            =   120
      TabIndex        =   33
      Top             =   5400
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date Recieved"
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
      Left            =   8445
      TabIndex        =   23
      Top             =   945
      Width           =   1365
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
      Left            =   12075
      TabIndex        =   25
      Top             =   1170
      Width           =   615
   End
   Begin VB.Label Label12 
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
      Left            =   11130
      TabIndex        =   24
      Top             =   1170
      Width           =   825
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   975
      Left            =   10920
      Top             =   1785
      Width           =   2235
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Assigned Disbursement Voucher No :"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   420
      TabIndex        =   12
      Top             =   6045
      Width           =   2640
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00000000&
      FillStyle       =   0  'Solid
      Height          =   1725
      Left            =   -855
      Top             =   5850
      Width           =   11655
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Entered Transactions"
      Height          =   195
      Left            =   10920
      TabIndex        =   11
      Top             =   2940
      Width           =   1515
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Alobs/OBR No."
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   450
      TabIndex        =   7
      Top             =   945
      Width           =   1530
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   960
      Left            =   -855
      Top             =   840
      Width           =   11775
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   960
      Left            =   10920
      Top             =   825
      Width           =   2235
   End
End
Attribute VB_Name = "frmIncomingTrn3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Edited As Boolean
Dim DTE As String
Dim UID As String
Dim EditedDV As String
Dim XFlag As Boolean
Dim CATotalamount As Currency


Private Sub btnClaimant_Click()
    ActiveFormCaller = "frmIncomingTrn"
    frmCDClaimantRegistry.Show 1
End Sub

Private Sub btnSearch_Click()
    frmDVSearch.Show 1
End Sub

Private Sub cmb_trnYear_Click()
    Call LoadPrevTrans(cmb_trnYear.Text)
End Sub

Private Sub LoadPrevTrans(ByVal Year_ As Integer)
Dim PRec As New ADODB.Recordset
Dim x As Integer

List1.Clear
List1.Enabled = False

PRec.Open ("Select * From tblAMIS_IncomingDVTrns Where TransactionDate like '%" & Year_ & "' And ActionCode=1 And [PAout]=0 Order by trnno desc"), opndbaseFMIS, adOpenStatic, adLockOptimistic
If PRec.RecordCount > 0 Then
    For x = 1 To PRec.RecordCount
        List1.AddItem PRec!DVNo
        List1.ItemData(List1.NewIndex) = PRec!Trnno
        PRec.MoveNext
    Next x
    List1.Enabled = True
End If
PRec.Close
Set PRec = Nothing

End Sub

Private Sub cmbFund_Click()
    If lblMode.Caption = "NEW" Then
        TxtDvno.Text = GetNewDVNumber(cmbfund.Text)
    End If

End Sub

Private Sub cmbNonAlobs_Change()
Call cmbNonAlobs_Click
End Sub

'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'+++++ Purpose / Description    : To allow DV No input in the txtDVNo text box when the selected itemdata=21.
'+++++ Input                    : None
'+++++ Return                   : None
'+++++ Date Created             : April 27, 2010
'+++++ Programmer               : Eduard Emmanuel D. Gatong
'+++++ UPDATES +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'+++++ Purpose / Description    :
'+++++ Date Updated             :
'+++++ Programmer               :
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'+++++ Purpose / Description    :
'+++++ Date Updated             :
'+++++ Programmer               :
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Private Sub cmbNonAlobs_Click()
If Trim(cmbNonAlobs.Text) = "Liquidation of Cash Advance" Then
fmeCA.Visible = True
Else
fmeCA.Visible = False
End If
End Sub
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Private Sub Command1_Click()
If Edited = True Then
    DVNoOut = EditedDV
    frmDVLogOut.Show 1
End If

End Sub

Private Sub Command2_Click()
frmLock.ShowForm
If Iflock = True Then
    If Edited = False Then
    txtAmount.Text = Format(GetRemainingAmnt(txtObR.Text), "#,##0.00")
    Else
        txtAmount.Text = Format((GetRemainingAmntInBUDGET(txtObR.Text)), "#,##0.00")
    End If
    txtAmount.Locked = False
End If
End Sub

Private Sub Command3_Click()
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
Public Function IfexistDv(ByVal DVNo As String) As Boolean
Dim y As Integer
IfexistDv = False
If ListView1.ListItems.Count <> 0 Then
    For y = 1 To ListView1.ListItems.Count
        If DVNo = ListView1.ListItems(y).SubItems(1) Then
            IfexistDv = True
        End If
    Next y
End If
End Function


Private Sub Command4_Click()
fmeCA.Visible = False
If txtctotalAmnt.Text <> "" Then
txtAmount.Text = txtctotalAmnt.Text
End If
End Sub
Private Function CAClear()
txtCamount.Text = ""
txtCChecdate.Text = ""
'txtCCheckno.Text = ""
txtCClaimant.Text = ""
txtCParticular.Text = ""
End Function
Private Sub fmeDetails_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Label14.FontBold = False
Label14.FontUnderline = False
End Sub

Private Sub Form_Activate()
    If lblRefresh.Caption = "True" Then
        Call Toolbar1_ButtonClick(Toolbar1.Buttons(1))
        lblRefresh.Caption = "False"
    End If
End Sub



'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'+++++ Purpose / Description    : To lock/unlock the txtDVNo textbox on runtime.
'+++++ Input                    : None
'+++++ Return                   : None
'+++++ Date Created             : April 27, 2010
'+++++ Programmer               : Eduard Emmanuel D. Gatong
'+++++ UPDATES +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'+++++ Purpose / Description    :
'+++++ Date Updated             :
'+++++ Programmer               :
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'+++++ Purpose / Description    :
'+++++ Date Updated             :
'+++++ Programmer               :
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If Shift = vbCtrlMask + vbShiftMask Then
        If KeyCode = vbKeyF8 Then
            If TxtDvno.Locked = True Then
                TxtDvno.Locked = False
                MsgBox "Unlocked!"
            Else
                TxtDvno.Locked = True
                MsgBox "Locked!"
            End If
        End If
    End If

End Sub
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Private Sub Form_Load()

Me.Top = (Screen.Height - Me.Height) / 2
Me.Left = (Screen.Width - Me.Width) / 2

ActiveUserID = Trim(ActiveUserID)
Call Toolbar1_ButtonClick(Toolbar1.Buttons(1))


Edited = False
End Sub


Private Sub Label14_Click()
fmeCA.Visible = True
End Sub

Private Sub Label14_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Label14.FontBold = True
Label14.FontUnderline = True
End Sub

Private Sub List1_Click()
    Call ReLoadDetail(List1.Text)
    'fmeCA.Visible = False
End Sub

'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'+++++ Purpose / Description    : Reloading of the saved transaction data.
'+++++ Input                    : (String) DV Number
'+++++ Return                   : None
'+++++ Date Created             : April 14, 2010
'+++++ Programmer               : Eduard Emmanuel D. Gatong
'+++++ UPDATES +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'+++++ Purpose / Description    :
'+++++ Date Updated             :
'+++++ Programmer               :
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'+++++ Purpose / Description    :
'+++++ Date Updated             :
'+++++ Programmer               :
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Private Sub ReLoadDetail(ByVal DVNumber As String)
Dim DVRec As New ADODB.Recordset
    On Error Resume Next
    XFlag = False
    
    DVRec.Open ("Select * from tblAMIS_IncomingDVTrns where DVNo='" & DVNumber & "' and Actioncode=1"), opndbaseFMIS, adOpenStatic, adLockOptimistic
    If DVRec.RecordCount > 0 Then
        Edited = True
        frmTrans.Enabled = False
        If DVRec!Continuing = 1 Then
            XFlag = True
        End If
        EditedDV = DVNumber
        lblMode.Caption = "EDIT"
        fmeCA.Visible = False
        Label14.Visible = False
        txtctotalAmnt.Text = ""
        If DVRec!NonAlobs = 0 Then
            optObR.Value = True
            txtObR.Text = DVRec![obrno]
            txtOfficeCode.Text = DVRec![RCenter]
            txtOffice.Text = GetOfficeName(DVRec![RCenter], "OfficeMedium")
            txtfund.Text = DVRec![FundType]
            txtooe.Text = IIf(IsNull(DVRec![OOE]), "", (DVRec![OOE]))
            
        Else
            optNonObR.Value = True
            cmbNonAlobs.Text = GetNonAlobsName(DVRec![obrno])
            cmbrc.Text = GetOfficeName(DVRec![RCenter], "OfficeMedium")
            cmbfund.Text = DVRec![FundType]
            cmbOOE.Text = DVRec![OOE]
            
            If Trim(cmbNonAlobs.Text) = "Liquidation of Cash Advance" Then
           ' txtCDvno.Text = getCADvnoByLdvno(DVRec![dvno])
                Call AllLoadCAdetails(ListView1, DVNumber, txtctotalAmnt)
                fmeCA.Visible = True
                Label14.Visible = True
                
            End If
        End If
        txtclaimantcode = IIf(IsNull(DVRec!ClaimantCode), "", DVRec!ClaimantCode)
        txtclaimant.Text = getClaimant(IIf(IsNull(DVRec!ClaimantCode), "", DVRec!ClaimantCode))
        txtDetail.Text = DVRec![Particular]
        txtAmount.Text = Format(DVRec![Gamount], "#,###.00")
        TxtDvno.Text = DVRec![DVNo]
        txtDate.Text = Format(DVRec![TransactionDate], "mmmm dd, yyyy")
        DTE = DVRec![datetimeentered]
        UID = DVRec![UserID]
    Else
        MsgBox "Invalid DV Number!", vbExclamation + vbOKOnly, "System Security"
        Call Toolbar1_ButtonClick(Toolbar1.Buttons(1))
    End If
    DVRec.Close
    Set DVRec = Nothing
    
    'opndbaseFMIS.Execute "Insert Into tblAMIS_IncomingDVTrns (DVNo,ObrNo,FundType,RCenter,RCenterCode,OOE,ClaimantCode,Particular,GAmount,TransactionDate,UserID,Actioncode,DateTimeEntered) Values ('" & txtDVNo.Text & "','" & Trim(txtObR.Text) & "','" & txtFund.Text & "'," & txtOfficeCode.Text & "," & Mid(txtObR.Text, 5, 4) & ",'" & txtOOE.Text & "','" & txtClaimantCode.Text & "','" & txtDetail.Text & "'," & CCur(txtAmount.Text) & ",'" & Format(txtDate.Text, "mm/dd/yyyy") & "','" & ActiveUserID & "',1,'" & Format(Now, "mm/dd/yyyy hh:mm:ss AMPM") & "')"
    
End Sub
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Public Function getCADvnoByLdvno(ByVal liqDvno As String)
Dim rec As New ADODB.Recordset
rec.Open "Select * from tblAMIS_LiquiditionOfCA where liquiDvno = '" & liqDvno & "' and actioncode = 1", opndbaseFMIS, adOpenStatic, adLockOptimistic
    If rec.RecordCount <> 0 Then
        getCADvnoByLdvno = Trim(rec.Fields!cadvno)
    End If
rec.Close
End Function


Private Sub ListView1_Click()
If ListView1.ListItems.Count <> 0 Then
            txtCDvno.Text = ListView1.SelectedItem.SubItems(1)
             txtCCheckno.Text = ListView1.SelectedItem.SubItems(2)
             txtCChecdate.Text = ListView1.SelectedItem.SubItems(3)
             txtCParticular.Text = ListView1.SelectedItem.SubItems(4)
             txtCClaimant.Text = ListView1.SelectedItem.SubItems(5)
            txtCamount.Text = ListView1.SelectedItem.SubItems(6)
End If
End Sub
Private Sub ListView1_KeyPress(KeyAscii As Integer)
If vbKeyDelete Then
ListView1.ListItems.Remove (ListView1.SelectedItem.Index)
txtctotalAmnt.Text = Format(GetCATotalamount(ListView1), "#,##0.00")
End If
End Sub

Private Sub optNonObR_Click()
If optNonObR.Value = True Then
    txtObR.Visible = False
    txtOffice.Visible = False
    txtfund.Visible = False
    txtooe.Visible = False
    Command2.Visible = False
    
    cmbNonAlobs.Width = txtObR.Width
    cmbNonAlobs.Left = txtObR.Left
    cmbNonAlobs.Top = txtObR.Top
    cmbNonAlobs.Visible = True
    
    cmbrc.Width = txtOffice.Width
    cmbrc.Left = txtOffice.Left
    cmbrc.Top = txtOffice.Top
    cmbrc.Visible = True

    cmbfund.Width = txtfund.Width
    cmbfund.Left = txtfund.Left
    cmbfund.Top = txtfund.Top
    cmbfund.Visible = True

    cmbOOE.Width = txtooe.Width
    cmbOOE.Left = txtooe.Left
    cmbOOE.Top = txtooe.Top
    cmbOOE.Visible = True
    Label14.Visible = True
    
    txtDetail.Locked = False
    txtAmount.Locked = False
    CAClear
    txtObR.Text = ""
    txtclaimant.Text = ""
    txtclaimantcode.Text = ""
    txtOffice.Text = ""
    txtfund.Text = ""
    txtooe.Text = ""
    txtDetail.Text = ""
    txtAmount.Text = ""
    TxtDvno.Text = ""
End If
End Sub

Private Sub optObR_Click()
   If optObR.Value = True Then
    cmbNonAlobs.Visible = False
    cmbrc.Visible = False
    cmbfund.Visible = False
    cmbOOE.Visible = False
    fmeCA.Visible = False
    
    txtObR.Visible = True
    txtOffice.Visible = True
    txtfund.Visible = True
    txtooe.Visible = True
    Command2.Visible = True
    Label14.Visible = False
    
    'txtDetail.Locked = True
   ' txtAmount.Locked = True
    CAClear
    txtObR.Text = ""
    txtclaimant.Text = ""
    txtclaimantcode.Text = ""
    txtOffice.Text = ""
    txtfund.Text = ""
    txtooe.Text = ""
    txtDetail.Text = ""
    txtAmount.Text = ""
    TxtDvno.Text = ""
    'Call Toolbar1_ButtonClick(Toolbar1.Buttons(1))
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


Private Sub LoadFund()
Dim Frec As New ADODB.Recordset
Dim x As Integer

cmbfund.Clear

Frec.Open ("Select * from tblRefBMS_Funds Order By FundMedium"), opndbaseFMIS, adOpenStatic, adLockOptimistic
If Frec.RecordCount > 0 Then
    For x = 1 To Frec.RecordCount
        cmbfund.AddItem Frec!fundmedium
        cmbfund.ItemData(cmbfund.NewIndex) = Frec!fundcode
        Frec.MoveNext
    Next x
End If
Frec.Close
Set Frec = Nothing

End Sub

Private Sub LoadOOE()
Dim OREc As New ADODB.Recordset
Dim x As Integer

cmbOOE.Clear

OREc.Open ("Select * From tblBMS_ObjectOfExpenditures Order By OOEName"), opndbaseFMIS, adOpenStatic, adLockOptimistic
If OREc.RecordCount > 0 Then
    For x = 1 To OREc.RecordCount
        cmbOOE.AddItem OREc!OOEName
        cmbOOE.ItemData(cmbOOE.NewIndex) = OREc!OOECode
        OREc.MoveNext
    Next x
End If
OREc.Close
Set OREc = Nothing

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

    Select Case Button
    Case "Close":
                    If MsgBox("Are you sure you want to close this form?", vbQuestion + vbYesNo, "System Security") = vbYes Then
                        Unload Me
                    End If
    Case "New":
                    XFlag = False
                    optObR.Value = True
                    UID = ""
                    DTE = ""
                    Edited = False
                    EditedDV = ""
                    lblMode.Caption = "NEW"
                    Call LoadTrnYear(cmb_trnYear)
                    txtDate.Text = Format(Now, "mmmm dd, yyyy")
                    txtObR.Text = ""
                    txtclaimant.Text = ""
                    txtclaimantcode.Text = ""
                    txtOffice.Text = ""
                    txtfund.Text = ""
                    txtooe.Text = ""
                    txtDetail.Text = ""
                    txtAmount.Text = ""
                    TxtDvno.Text = ""
                    TxtDvno.Locked = True
                    
                    txtObR.Visible = True
                    txtOffice.Visible = True
                    txtfund.Visible = True
                    txtooe.Visible = True

                    cmbrc.Visible = False
                    cmbOOE.Visible = False
                    cmbfund.Visible = False
                    cmbNonAlobs.Visible = False
                    frmTrans.Enabled = True
                    'txtDetail.Locked = True
                   ' txtAmount.Locked = True
                    
                    Call LoadNonAlobs
                    Call LoadOffice
                    Call LoadFund
                    Call LoadOOE
                    CAClear
                    ListView1.ListItems.Clear

    Case "Save":
                        
                        If Trim(cmbNonAlobs.Text) = "Liquidation of Cash Advance" Then
                                If ListView1.ListItems.Count <> 0 Then
                                    If txtAmount.Text <> txtctotalAmnt.Text Then
                                       If MsgBox("Gross Amount not Equal to your Total Cash Advance..!" & vbNewLine & "Are You Sure the Cash Advance Details Manually Operated past year?", vbCritical + vbYesNo, "System Message") = vbYes Then
                                       'do nothing
                                       Else
                                        'stop transaction
                                         Exit Sub
                                        End If
                                    End If
                                 Else
                                        If MsgBox("Cash Advance Details Empty..!" & vbNewLine & "Are You Sure the Cash Advance Details Manually Operated past year?", vbCritical + vbYesNo, "System Message") = vbYes Then
                                            'do nothing
                                            Else
                                            ' stop transaction
                                              Exit Sub
                                         End If
                                 End If
                            End If
                   
                        If txtObR.Text <> "" Then
                            If Edited = False Then
                                If ISAlobsAmtOkAgaintsVoucher(txtObR.Text, CCur(txtAmount.Text), GetTotalTrnsactedAmt(txtObR.Text, "tblAMIS_IncomingDVTrns", "GAmount", "ObrNo"), True) = False Then
                                    Exit Sub
                                End If
                            End If
                        End If
                    
                          If MsgBox("Are you sure you want to save this transaction?", vbQuestion + vbYesNo, "System Security") = vbYes Then
                        '  MsgBox cmbNonAlobs.ListIndex
                              If ChkEntry = True Then
                                      Dim xChange As String
                                      
                                      xChange = TxtDvno.Text
                                      
                                      If Edited = True Then
                                          opndbaseFMIS.Execute "Update tblAMIS_IncomingDVTrns set UserID='" & UID & "," & ActiveUserID & "',Actioncode=2,DateTimeEntered='" & DTE & "," & Format(Now, "mm/dd/yyyy hh:mm:ss AMPM") & "'  Where DVNo='" & TxtDvno.Text & "' and Actioncode=1"
                                          opndbaseFMIS.Execute "Update tblAMIS_LiquiditionOfCA set Actioncode=2  Where liquidvno='" & TxtDvno.Text & "' and Actioncode=1"
                                      End If
                                      
                                      If Edited = False Then
                                          If ChkDVExist(TxtDvno.Text) = True Then
                                              If optObR.Value = True Then
                                                  TxtDvno.Text = GetNewDVNumber(txtfund.Text)
                                              Else
                                                  TxtDvno.Text = GetNewDVNumber(cmbfund.Text)
                                              End If
                                          End If
                                      End If
                                      
                                      If optObR.Value = True Then 'with obr
                                          
                                          If XFlag = True And txtOffice.Text = "" And txtooe.Text = "" Then 'continuing
                                              opndbaseFMIS.Execute "Insert Into tblAMIS_IncomingDVTrns (DVNo,ObrNo,FundType,RCenter,RCenterCode,OOE,ClaimantCode,Particular,GAmount,TransactionDate,UserID,Actioncode,DateTimeEntered,Continuing) Values ('" & TxtDvno.Text & "','" & Trim(txtObR.Text) & "','" & txtfund.Text & "'," & cmbrc.ItemData(cmbrc.ListIndex) & "," & Mid(txtObR.Text, 5, 4) & ",'" & cmbOOE.Text & "','" & txtclaimantcode.Text & "','" & Trim(Replace(txtDetail.Text, "'", "''")) & "'," & CCur(txtAmount.Text) & ",'" & Format(txtDate.Text, "mm/dd/yyyy") & "','" & ActiveUserID & "',1,'" & Format(Now, "mm/dd/yyyy hh:mm:ss AMPM") & "'," & IIf(XFlag, 1, 0) & ")"
                                          Else 'current
                                              opndbaseFMIS.Execute "Insert Into tblAMIS_IncomingDVTrns (DVNo,ObrNo,FundType,RCenter,RCenterCode,OOE,ClaimantCode,Particular,GAmount,TransactionDate,UserID,Actioncode,DateTimeEntered,Continuing) Values ('" & TxtDvno.Text & "','" & Trim(txtObR.Text) & "','" & txtfund.Text & "'," & txtOfficeCode.Text & "," & Mid(txtObR.Text, 5, 4) & ",'" & txtooe.Text & "','" & txtclaimantcode.Text & "','" & Trim(Replace(txtDetail.Text, "'", "''")) & "'," & CCur(txtAmount.Text) & ",'" & Format(txtDate.Text, "mm/dd/yyyy") & "','" & ActiveUserID & "',1,'" & Format(Now, "mm/dd/yyyy hh:mm:ss AMPM") & "'," & IIf(XFlag, 1, 0) & ")"
                                          End If
                                      Else 'non alobs
                                          Dim z As Integer
                                          If Trim(cmbNonAlobs.Text) = "Liquidation of Cash Advance" Then
                                             If ListView1.ListItems.Count <> 0 Then
                                                 
                                                    For z = 1 To ListView1.ListItems.Count
                                                        opndbaseFMIS.Execute "Insert into tblAMIS_LiquiditionOfCA ([liquiDvno],[CADvno],[checkno],[checkdate],[status],[actioncode],[amount]) " & _
                                                                                " values ('" & TxtDvno.Text & "' , '" & ListView1.ListItems(z).SubItems(1) & "','" & ListView1.ListItems(z).SubItems(2) & "','" & ListView1.ListItems(z).SubItems(3) & "',0,1, " & CCur(ListView1.ListItems(z).SubItems(6)) & ") "
                                                    Next z
                                              End If
                                            opndbaseFMIS.Execute "Insert Into tblAMIS_IncomingDVTrns (DVNo,ObrNo,FundType,RCenter,RCenterCode,OOE,ClaimantCode,Particular,GAmount,TransactionDate,UserID,Actioncode,DateTimeEntered,NonAlobs,Continuing) Values ('" & TxtDvno.Text & "','" & GetNonAlobsCode(cmbNonAlobs.ItemData(cmbNonAlobs.ListIndex)) & "','" & cmbfund.Text & "'," & cmbrc.ItemData(cmbrc.ListIndex) & ",0,'" & cmbOOE.Text & "','" & txtclaimantcode.Text & "','" & Trim(Replace(txtDetail.Text, "'", "''")) & "'," & CCur(txtAmount.Text) & ",'" & Format(txtDate.Text, "mm/dd/yyyy") & "','" & ActiveUserID & "',1,'" & Format(Now, "mm/dd/yyyy hh:mm:ss AMPM") & "',1," & IIf(XFlag, 1, 0) & ")"
                                          Else
                                              opndbaseFMIS.Execute "Insert Into tblAMIS_IncomingDVTrns (DVNo,ObrNo,FundType,RCenter,RCenterCode,OOE,ClaimantCode,Particular,GAmount,TransactionDate,UserID,Actioncode,DateTimeEntered,NonAlobs,Continuing) Values ('" & TxtDvno.Text & "','" & GetNonAlobsCode(cmbNonAlobs.ItemData(cmbNonAlobs.ListIndex)) & "','" & cmbfund.Text & "'," & cmbrc.ItemData(cmbrc.ListIndex) & ",0,'" & cmbOOE.Text & "','" & txtclaimantcode.Text & "','" & Trim(Replace(txtDetail.Text, "'", "''")) & "'," & CCur(txtAmount.Text) & ",'" & Format(txtDate.Text, "mm/dd/yyyy") & "','" & ActiveUserID & "',1,'" & Format(Now, "mm/dd/yyyy hh:mm:ss AMPM") & "',1," & IIf(XFlag, 1, 0) & ")"
                                          End If
                                          
                                      End If
                                      
                                      If xChange <> TxtDvno.Text Then
                                          MsgBox "Disbursement Voucher number changed from " & xChange & " to " & TxtDvno.Text & vbCrLf & vbCrLf & "Transaction successfully saved!", vbInformation + vbOKOnly, "System Security"
                                      Else
                                          MsgBox "Transaction successfully saved!", vbInformation + vbOKOnly, "System Security"
                                      End If
                                      
                                      Call Toolbar1_ButtonClick(Toolbar1.Buttons(1))
                              Else
                                  MsgBox "Save operation cancelled!" & vbCrLf & vbCrLf & "Please check your entry.", vbExclamation + vbOKOnly, "System Security"
                              End If
                          End If
                        
    Case "Delete":
                    If Edited = True Then
                        If MsgBox("Are you sure you want to delete this transaction?", vbQuestion + vbYesNo) = vbYes Then
                            opndbaseFMIS.Execute "Update tblAMIS_IncomingDVTrns set UserID='" & UID & "," & ActiveUserID & "',Actioncode=3,DateTimeEntered='" & DTE & "," & Format(Now, "mm/dd/yyyy hh:mm:ss AMPM") & "'  Where DVNo='" & TxtDvno.Text & "' and Actioncode=1"
                            Call Toolbar1_ButtonClick(Toolbar1.Buttons(1))
                        End If
                    End If
    Case "Log Out":
                    If Edited = True Then
                        DVNoOut = EditedDV
                        frmDVLogOut.Show 1
                    End If
    End Select

End Sub
Public Function SaveLiquidit()
opndbaseFMIS.Execute "Insert into tblAMIS_LiquiditionOfCA ([liquiDvno],[CADvno],[checkno],[checkdate],[status],[actioncode]),amount " & _
                    " values ('" & TxtDvno.Text & "' , '" & txtCDvno.Text & "','" & txtCCheckno.Text & "','" & txtCChecdate.Text & "',0,1," & txtCamount.Text & ") "

End Function
Private Function ChkEntry() As Boolean

ChkEntry = False

If Trim(txtclaimant.Text) <> "" And Trim(txtclaimantcode.Text) <> "" And Trim(TxtDvno.Text) <> "" Then
    
    If cmbfund.Visible = True Then 'non obr
        If cmbNonAlobs.ListIndex <> -1 And cmbfund.Text <> "" And cmbrc.ListIndex <> -1 And cmbOOE.Text <> "" Then
            ChkEntry = True
        Else
            ChkEntry = False
        End If
    Else 'with obr
        If cmbrc.Visible = True Then 'continuing
            If txtfund.Text <> "" And cmbrc.ListIndex <> -1 And cmbOOE.Text <> "" Then
                ChkEntry = True
            Else
                ChkEntry = False
            End If
        Else 'current
            If txtfund.Text <> "" And txtOfficeCode.Text <> "" And txtooe.Text <> "" And txtclaimantcode.Text <> "" Then
                ChkEntry = True
            Else
                ChkEntry = False
            End If
        End If
    End If
Else
    ChkEntry = True
End If

End Function

Private Function GetNewDVNumber(ByVal FundName As String) As String
Dim Frec As New ADODB.Recordset
Dim FCode As String
Dim Remake As Boolean

GetNewDVNumber = 1

FCode = GetFundCODE(FundName)

Frec.Open ("Select * From tblAMIS_IncomingDVTrns Where substring(DVNo,1,10)='" & FCode & "-" & Format(Now, "yy-mm") & "-' Order by DVNo desc"), opndbaseFMIS, adOpenStatic, adLockOptimistic
'FRec.Open ("Select * From tblAMIS_IncomingDVTrns Where substring(DVNo,1,10)='" & FCode & "-" & Format(Now, "yy-mm") & "-' Order by trnno desc"), opndbaseFMIS, adOpenStatic, adLockOptimistic
If Frec.RecordCount > 0 Then
    GetNewDVNumber = Val(Right(Frec!DVNo, 4)) + 1
End If
Frec.Close
Set Frec = Nothing

GetNewDVNumber = FCode & "-" & Format(Now, "yy-mm") & "-" & Format(GetNewDVNumber, "000#")

xRemake:

Remake = False

Frec.Open ("Select * From tblAMIS_IncomingDVTrns Where DVNo='" & GetNewDVNumber & "'"), opndbaseFMIS, adOpenStatic, adLockOptimistic
If Frec.RecordCount > 0 Then
    Remake = True
    GetNewDVNumber = Mid(GetNewDVNumber, 1, 10) & Format(Val(Mid(GetNewDVNumber, 11, 4)) + 1, "000#")
End If
Frec.Close
Set Frec = Nothing

If Remake = True Then GoTo xRemake

End Function

Public Function LoadCAdetails(ByVal DVNo As String)
Dim rec As New ADODB.Recordset
Dim rec1 As New ADODB.Recordset
    rec.Open "SELECT top 1 percent    a.DVNo, b.CheckNo, b.CheckDate, a.Particular, b.ClaimantName,a.GAmount " & _
            "FROM dbo.tblAMIS_IncomingDVTrns AS a inner join tblCMS_CDNewFMISVoucher as c on a.dvno = left(c.newcontrolno,14) inner join  " & _
            "dbo.tblCMS_CDPreparedCheck AS b ON c.fmisvoucherno = b.MixCode   " & _
            "Where (a.ActionCode = 1) And (b.ActionCode = 1) And a.Dvno = '" & DVNo & "'", opndbaseFMIS, adOpenStatic, adLockOptimistic
    If rec.RecordCount <> 0 Then
    
        txtCCheckno.Text = rec.Fields!checkno
        txtCChecdate.Text = rec.Fields!CheckDate
        txtCClaimant.Text = rec.Fields!ClaimantName
        txtCamount.Text = Format(rec.Fields!Gamount, "#,##0.00")
        txtCParticular.Text = rec.Fields!Particular
    Else
        rec.Close
        Set rec = Nothing
        rec1.Open "SELECT top 1 percent    a.DVNo, b.CheckNo, b.CheckDate, a.Particular, b.ClaimantName,a.GAmount " & _
            "FROM dbo.tblAMIS_IncomingDVTrns AS a inner join " & _
            "dbo.tblCMS_CDPreparedCheck AS b ON a.dvno = b.MixCode   " & _
            "Where (a.ActionCode = 1) And (b.ActionCode = 1) And a.Dvno = '" & DVNo & "'", opndbaseFMIS, adOpenStatic, adLockOptimistic
        If rec1.RecordCount > 0 Then
        
        txtCCheckno.Text = rec1.Fields!checkno
        txtCChecdate.Text = rec1.Fields!CheckDate
        txtCClaimant.Text = rec1.Fields!ClaimantName
        txtCamount.Text = Format(rec1.Fields!Gamount, "#,##0.00")
        txtCParticular.Text = rec1.Fields!Particular
      
        Else
        MsgBox "Dvno Not Found", vbInformation, "System Message"
        End If
          rec1.Close
        Set rec1 = Nothing
    End If
End Function
Public Function IFLiquidit(ByVal DVNo As String) As Boolean
Dim rec As New ADODB.Recordset
IFLiquidit = False
rec.Open "Select * from tblAMIS_LiquiditionOfCA where cadvno = '" & DVNo & "' and actioncode = 1", opndbaseFMIS, adOpenStatic, adLockOptimistic
    If rec.RecordCount <> 0 Then
        IFLiquidit = True
    End If
rec.Close
End Function

Private Sub txtAmount_LostFocus()
Iflock = False
txtAmount.Locked = True
txtAmount.Text = Format(txtAmount.Text, "#,##0.00")
End Sub

Private Sub txtCCheckno_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txtCDvno.Text = GetDVNobyChkNo(txtCCheckno.Text)
    If txtCDvno.Text <> "" Then
         If IFLiquidit(txtCDvno.Text) = True Then
         MsgBox "This Transaction Already Liquidit, Cannot Proccess the Trasaction", vbInformation, "System Message"
         CAClear
         Else
         LoadCAdetails (txtCDvno.Text)
         End If
     End If
Else
CAClear
End If
End Sub

Private Sub txtCDvno_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If txtCDvno.Text <> "" Then
        If IFLiquidit(txtCDvno.Text) = True Then
        MsgBox "This Transaction Already Liquidit, Cannot Proccess the Trasaction", vbInformation, "System Message"
        CAClear
        Else
        LoadCAdetails (txtCDvno.Text)
        End If
    End If
Else
CAClear
End If
End Sub
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'+++++ Purpose / Description    : Validation of inputed ObR and retrieval of ObR details.
'+++++ Date Created             : January 18, 2010
'+++++ Programmer               : Eduard Emmanuel D. Gatong
'+++++ UPDATES +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'+++++ Purpose / Description    :
'+++++ Date Updated             :
'+++++ Programmer               :
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'+++++ Purpose / Description    :
'+++++ Date Updated             :
'+++++ Programmer               :
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Private Sub txtObR_KeyPress(KeyAscii As Integer)
Dim sDVNo As String

    If KeyAscii = 13 Then
        lblMode.Caption = "NEW"
        lblMode.ForeColor = &HFF0000
        If Trim(txtObR.Text) <> "" Then
            txtObR.Text = Trim(Replace(txtObR.Text, "-", ""))
            
            If Len(txtObR.Text) = 15 Then
                txtObR.Text = Format(txtObR.Text, "###-####-##-##-####")
                If ValidObR(txtObR.Text) Then
'                    sDVNo = DVNumbered(txtObR.Text)
'                    If sDVNo <> "" Then
'                        If AlreadyOut(sDVNo) Then
'                            'MsgBox "This ObR number is already out with DV Number " & sDVNo, vbExclamation + vbOKOnly, "System Securty"
'                            'Call Toolbar1_ButtonClick(Toolbar1.Buttons(1))
'
'                        Else
'                            lblMode.Caption = "EDIT"
'                            lblMode.ForeColor = &HFF&
'                            Call ReLoadDetail(sDVNo)
'                        End If
'                    Else
                        lblMode.Caption = "NEW"
                        frmTrans.Enabled = False
                        lblMode.ForeColor = &HFF0000
                        Call GetObRData(txtObR.Text)
                        txtAmount.Text = Format(GetRemainingAmnt(txtObR.Text), "#,##0.00")
                        If txtAmount.Text = "0.00" Then
                            MsgBox "Obr\Alobs Number have 0 balance,Cannot Procces the Transaction..", vbInformation, "System Message"
                            Call Toolbar1_ButtonClick(Toolbar1.Buttons(1))
                           ' txtAmount.Locked = True
                            Exit Sub
                        End If
                        TxtDvno.Text = GetNewDVNumber(txtfund.Text)
                    'End If
                Else
                    MsgBox "Invalid ObR!", vbExclamation + vbOKOnly, "System Securty"
                End If
                
            ElseIf Len(txtObR.Text) = 11 Then
                txtObR.Text = Format(txtObR.Text, "###-##-##-####")
                If AlreadyOut(txtObR.Text) Then
                    MsgBox "This DV number is already out!", vbExclamation + vbOKOnly, "System Securty"
                    Call Toolbar1_ButtonClick(Toolbar1.Buttons(1))
                Else
                    lblMode.Caption = "EDIT"
                    lblMode.ForeColor = &HFF&
                    Call ReLoadDetail(txtObR.Text)
'                    AllLoadCAdetails (txtDVNo.Text)
                End If
            
            Else
                MsgBox "Invalid ObR / DV No!", vbExclamation + vbOKOnly, "System Securty"
            End If
        Else
            MsgBox "Input ObR first!", vbExclamation + vbOKOnly, "System Securty"
        End If
    End If
End Sub
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Private Function AlreadyOut(ByVal DVNo As String) As Boolean
Dim AORec As New ADODB.Recordset

    AlreadyOut = False
    AORec.Open ("Select outby From [tblAMIS_IncomingDVTrns] where DVNo='" & DVNo & "' and ActionCode=1 and ReturnFlag=0"), opndbaseFMIS, adOpenStatic, adLockOptimistic
    If AORec.RecordCount > 0 Then
        If Trim(AORec![OutBy]) <> "" Then
            AlreadyOut = True
        End If
    End If
    AORec.Close
    Set AORec = Nothing
    
End Function

'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'+++++ Purpose / Description    : Retrieval of ObR details.
'+++++ Input                    : (String) Alobs no.
'+++++ Output                   : None
'+++++ Date Created             : January 18, 2010
'+++++ Programmer               : Eduard Emmanuel D. Gatong
'+++++ UPDATES +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'+++++ Purpose / Description    :
'+++++ Date Updated             :
'+++++ Programmer               :
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'+++++ Purpose / Description    :
'+++++ Date Updated             :
'+++++ Programmer               :
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Private Sub GetObRData(ByVal ObR As String)
Dim OREc As New ADODB.Recordset
Dim OName As String
Dim OCode As Integer
Dim OOE As String

XFlag = False

OREc.Open ("Select * From tblFMIS_Transaction Where AlobsNo='" & ObR & "' And ActionCode=1"), opndbaseFMIS, adOpenStatic, adLockOptimistic
If OREc.RecordCount > 0 Then
    txtOffice.Text = GetOfficeName(OREc!FMISOfficeCode, "OfficeMedium")
    txtOfficeCode.Text = OREc!FMISOfficeCode
    If Mid(ObR, 1, 3) = "118" Then
        txtfund.Text = "20% DF"
    ElseIf Mid(ObR, 1, 3) = "101" Then
        txtfund.Text = "GF-Proper"
    Else
        txtfund.Text = OREc!FundType
    End If
    txtooe.Text = OREc!OOE
    txtDetail.Text = OREc!Particulars
    'txtAmount.Text = Format(OREc!Amount, "###,##0.00")
End If
OREc.Close
Set OREc = Nothing

OREc.Open ("Select * From [tblBMS_ExcessControl] Where AlobsNo='" & ObR & "' And ActionCode=1"), opndbaseFMIS, adOpenStatic, adLockOptimistic
If OREc.RecordCount > 0 Then
    XFlag = True
    Call GetObROffice(OREc!AlobsNoCharge, OName, OCode, OOE)
    txtOffice.Text = OName
    txtOfficeCode.Text = OCode
    If Mid(ObR, 1, 3) = "118" Then
        txtfund.Text = "20% DF"
    ElseIf Mid(ObR, 1, 3) = "101" Then
        txtfund.Text = "GF-Proper"
    Else
        txtfund.Text = GetFundMedium(Mid(ObR, 1, 3))
    End If
    txtooe.Text = OOE
    txtDetail.Text = OREc![Details]
    'txtAmount.Text = Format(OREc![Amount], "###,##0.00")
End If
OREc.Close
Set OREc = Nothing

If XFlag = True And txtOffice.Text = "" And txtooe.Text = "" Then
    
    txtOffice.Visible = False
    cmbrc.Width = txtOffice.Width
    cmbrc.Left = txtOffice.Left
    cmbrc.Top = txtOffice.Top
    cmbrc.Visible = True

    txtooe.Visible = False
    cmbOOE.Width = txtooe.Width
    cmbOOE.Left = txtooe.Left
    cmbOOE.Top = txtooe.Top
    cmbOOE.Visible = True

End If

End Sub
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Private Sub GetObROffice(ByVal ObR As String, OName As String, OCode As Integer, OOE As String)
Dim GORec As New ADODB.Recordset

    OName = ""
    OCode = 0
    OOE = ""
    GORec.Open ("Select * From tblFMIS_Transaction Where AlobsNo='" & ObR & "' And ActionCode=1"), opndbaseFMIS, adOpenStatic, adLockOptimistic
    If GORec.RecordCount > 0 Then
        OName = GetOfficeName(GORec!FMISOfficeCode, "OfficeMedium")
        OCode = GORec!FMISOfficeCode
        OOE = GORec!OOE
    End If
    GORec.Close
    Set GORec = Nothing
    
End Sub
Private Function LoadOffice()
Dim OREc As New ADODB.Recordset
Dim x As Integer

cmbrc.Clear
        OREc.Open ("Select * FRom tblREF_AIS_Offices Order By [OfficeMedium]"), opndbaseFMIS, adOpenStatic, adLockOptimistic
        If OREc.RecordCount > 0 Then
            For x = 1 To OREc.RecordCount
                cmbrc.AddItem OREc![OfficeMedium]
                cmbrc.ItemData(cmbrc.NewIndex) = OREc!fmisofficeid
                OREc.MoveNext
            Next x
        End If
        OREc.Close
        Set OREc = Nothing
End Function


'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'+++++ Purpose / Description    : Validation of ObR
'+++++ Input                    : (String) Alobs no.
'+++++ Output                   : (Boolean) True if already numbered, false otherwise.
'+++++ Date Created             : January 18, 2010
'+++++ Programmer               : Eduard Emmanuel D. Gatong
'+++++ UPDATES +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'+++++ Purpose / Description    :
'+++++ Date Updated             :
'+++++ Programmer               :
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'+++++ Purpose / Description    :
'+++++ Date Updated             :
'+++++ Programmer               :
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Private Function DVNumbered(ByVal ObR As String) As String
Dim DRec As New ADODB.Recordset

    DVNumbered = ""
    
    DRec.Open ("Select * From tblAMIS_IncomingDVTrns Where [ObrNo]='" & ObR & "' And ActionCode=1"), opndbaseFMIS, adOpenStatic, adLockOptimistic
    If DRec.RecordCount > 0 Then
        DVNumbered = DRec!DVNo
    End If
    DRec.Close
    Set DRec = Nothing
    
End Function
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++



VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{1693405E-2DC9-4248-B52F-4AC9145DA2AF}#1.0#0"; "WinXPC Engine.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MShflxgd.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmSub2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Import from Other Database"
   ClientHeight    =   9555
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9510
   Icon            =   "frmSub2.frx":0000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9555
   ScaleWidth      =   9510
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "NB"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   1
      Left            =   5760
      TabIndex        =   21
      Top             =   1080
      Width           =   1095
      Begin VB.OptionButton OptCredit 
         Caption         =   "Credit"
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
         TabIndex        =   23
         Top             =   480
         Width           =   855
      End
      Begin VB.OptionButton OptDebit 
         Caption         =   "Debit"
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
         TabIndex        =   22
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.TextBox txtaccountcode 
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
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   18
      Top             =   120
      Width           =   2175
   End
   Begin VB.Frame fmePayroll 
      Height          =   855
      Left            =   1560
      TabIndex        =   13
      Top             =   1080
      Width           =   4095
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Left            =   2400
         TabIndex        =   14
         Top             =   360
         Width           =   1575
         _ExtentX        =   2778
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
         CustomFormat    =   "MMMM"
         Format          =   157024259
         UpDown          =   -1  'True
         CurrentDate     =   40764
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   720
         TabIndex        =   15
         Top             =   360
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
         Format          =   156958723
         UpDown          =   -1  'True
         CurrentDate     =   40764
      End
      Begin VB.Label Label4 
         Caption         =   "Month:"
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
         Left            =   1800
         TabIndex        =   17
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Year:"
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
         TabIndex        =   16
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.ComboBox Combo1 
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
      ItemData        =   "frmSub2.frx":076A
      Left            =   7080
      List            =   "frmSub2.frx":0783
      TabIndex        =   11
      Top             =   120
      Width           =   2175
   End
   Begin VB.TextBox txtgamount 
      Alignment       =   1  'Right Justify
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
      Left            =   4560
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   9000
      Width           =   3255
   End
   Begin WinXPC_Engine.WindowsXPC WindowsXPC1 
      Left            =   3960
      Top             =   10200
      _ExtentX        =   6588
      _ExtentY        =   1085
      ColorScheme     =   2
      Common_Dialog   =   0   'False
   End
   Begin lvButton.lvButtons_H lvButtons_H4 
      Height          =   495
      Left            =   8160
      TabIndex        =   0
      Top             =   9000
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Caption         =   "&Back"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
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
      Image           =   "frmSub2.frx":07C6
      ImgSize         =   24
      cBack           =   -2147483633
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3630
      Left            =   120
      ScaleHeight     =   3600
      ScaleWidth      =   9225
      TabIndex        =   1
      Top             =   5280
      Width           =   9255
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   3600
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   9360
         _ExtentX        =   16510
         _ExtentY        =   6350
         _Version        =   393216
         BackColor       =   16777215
         BackColorSel    =   8454143
         ForeColorSel    =   0
         GridLinesUnpopulated=   1
         SelectionMode   =   1
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
   Begin VB.PictureBox freeSizer1 
      Height          =   480
      Left            =   11160
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   8
      Top             =   10200
      Width           =   1200
   End
   Begin VB.Frame Frame2 
      Height          =   2535
      Left            =   7320
      TabIndex        =   5
      Top             =   9840
      Visible         =   0   'False
      Width           =   4815
      Begin VB.TextBox Text1 
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
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   360
         Width           =   3135
      End
      Begin VB.Label Label8 
         Caption         =   "Enter ARE no."
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
         TabIndex        =   6
         Top             =   360
         Width           =   1455
      End
   End
   Begin lvButton.lvButtons_H lvButtons_H3 
      Height          =   375
      Left            =   8880
      TabIndex        =   12
      Top             =   600
      Visible         =   0   'False
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      Caption         =   "...."
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H lvButtons_H1 
      Height          =   615
      Left            =   8160
      TabIndex        =   19
      Top             =   1320
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1085
      Caption         =   "&Add to Journal"
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
      Image           =   "frmSub2.frx":42D0
      cBack           =   16777215
   End
   Begin lvButton.lvButtons_H lvButtons_H2 
      Height          =   615
      Left            =   6960
      TabIndex        =   24
      Top             =   1320
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1085
      Caption         =   "&Load"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
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
      Image           =   "frmSub2.frx":4622
      cBack           =   16777215
   End
   Begin lvButton.lvButtons_H lvButtons_H5 
      Height          =   375
      Left            =   3840
      TabIndex        =   25
      Top             =   120
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      Caption         =   "...."
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin VB.TextBox txttitle 
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
      Left            =   1560
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   20
      Top             =   600
      Width           =   7695
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3135
      Left            =   120
      TabIndex        =   29
      Top             =   2040
      Width           =   9300
      _ExtentX        =   16404
      _ExtentY        =   5530
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "List"
      TabPicture(0)   =   "frmSub2.frx":812C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame4"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Search batch No. Through Lastname"
      TabPicture(1)   =   "frmSub2.frx":8148
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Combo2"
      Tab(1).Control(1)=   "DTPicker3"
      Tab(1).Control(2)=   "lstresult"
      Tab(1).Control(3)=   "txtfind"
      Tab(1).Control(4)=   "List1"
      Tab(1).Control(5)=   "Label12"
      Tab(1).Control(6)=   "Label11"
      Tab(1).Control(7)=   "Label10"
      Tab(1).Control(8)=   "Label9"
      Tab(1).ControlCount=   9
      Begin VB.ComboBox Combo2 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "frmSub2.frx":8164
         Left            =   -69720
         List            =   "frmSub2.frx":816E
         Style           =   2  'Dropdown List
         TabIndex        =   46
         Top             =   720
         Visible         =   0   'False
         Width           =   1575
      End
      Begin MSComCtl2.DTPicker DTPicker3 
         Height          =   330
         Left            =   -71880
         TabIndex        =   45
         Top             =   720
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   582
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "MMMM, yyyyy"
         Format          =   184418307
         CurrentDate     =   41360
      End
      Begin MSComctlLib.ListView lstresult 
         Height          =   1935
         Left            =   -71880
         TabIndex        =   42
         Top             =   1080
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   3413
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Batch NO."
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Period"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.TextBox txtfind 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -74880
         TabIndex        =   40
         Top             =   720
         Width           =   2895
      End
      Begin VB.Frame Frame4 
         Height          =   2655
         Left            =   120
         TabIndex        =   30
         Top             =   360
         Width           =   9015
         Begin VB.ListBox lstOffices 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFC0&
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   2025
            ItemData        =   "frmSub2.frx":8178
            Left            =   120
            List            =   "frmSub2.frx":817A
            Style           =   1  'Checkbox
            TabIndex        =   34
            Top             =   555
            Width           =   8805
         End
         Begin VB.TextBox txtquery 
            Appearance      =   0  'Flat
            Height          =   1725
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   33
            Text            =   "frmSub2.frx":817C
            Top             =   675
            Width           =   8775
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
            Left            =   4320
            Style           =   2  'Dropdown List
            TabIndex        =   32
            Top             =   1155
            Visible         =   0   'False
            Width           =   3375
         End
         Begin VB.CheckBox Check1 
            Height          =   375
            Left            =   8520
            TabIndex        =   31
            Top             =   1515
            Visible         =   0   'False
            Width           =   375
         End
         Begin lvButton.lvButtons_H lvButtons_H6 
            Height          =   375
            Left            =   2400
            TabIndex        =   35
            Top             =   120
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   661
            Caption         =   "Select all"
            CapAlign        =   2
            BackStyle       =   4
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
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
            cBack           =   -2147483633
         End
         Begin lvButton.lvButtons_H lvButtons_H7 
            Height          =   375
            Left            =   3480
            TabIndex        =   36
            Top             =   120
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   661
            Caption         =   "Deselect all"
            CapAlign        =   2
            BackStyle       =   4
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
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
            cBack           =   -2147483633
         End
         Begin lvButton.lvButtons_H lvButtons_H8 
            Height          =   375
            Left            =   120
            TabIndex        =   38
            Top             =   120
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   661
            Caption         =   "View batch"
            CapAlign        =   2
            BackStyle       =   4
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
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
            cBack           =   -2147483633
         End
         Begin MSComctlLib.ProgressBar ProgressBar1 
            Height          =   135
            Left            =   4560
            TabIndex        =   39
            Top             =   240
            Visible         =   0   'False
            Width           =   4335
            _ExtentX        =   7646
            _ExtentY        =   238
            _Version        =   393216
            BorderStyle     =   1
            Appearance      =   0
            Scrolling       =   1
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            Caption         =   "Responsibilty Center"
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
            Left            =   120
            TabIndex        =   37
            Top             =   195
            Width           =   1935
         End
      End
      Begin MSComctlLib.ListView List1 
         Height          =   1935
         Left            =   -74880
         TabIndex        =   44
         Top             =   1080
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   3413
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Name"
            Object.Width           =   7056
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "ID"
            Object.Width           =   0
         EndProperty
      End
      Begin VB.Label Label12 
         Caption         =   "Period"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -69720
         TabIndex        =   48
         Top             =   480
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label Label11 
         Caption         =   "Month and Year"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -71880
         TabIndex        =   47
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label Label10 
         Caption         =   "Search Result"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -66960
         TabIndex        =   43
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label9 
         Caption         =   "Search Lastname"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74880
         TabIndex        =   41
         Top             =   480
         Width           =   1575
      End
   End
   Begin VB.Label lblID 
      Caption         =   "Label9"
      Height          =   375
      Left            =   4320
      TabIndex        =   28
      Top             =   120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Account Title:"
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
      TabIndex        =   27
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Account Code:"
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
      TabIndex        =   26
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "Select Criteria:"
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
      Left            =   5040
      TabIndex        =   10
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label lblResult 
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
      TabIndex        =   9
      Top             =   9000
      Width           =   1935
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "Total Amount"
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
      Left            =   2280
      TabIndex        =   4
      Top             =   9000
      Width           =   2055
   End
End
Attribute VB_Name = "frmSub2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public accountcode, Accountname, REFF As String
Public RC As Long

Private Sub Check1_Click()
Call lvButtons_H3_Click
End Sub

Private Sub cmbrc_Change()
Call lvButtons_H3_Click
End Sub

Private Sub cmbRC_Click()
Call lvButtons_H3_Click
End Sub

Private Sub Combo1_Change()
On Error Resume Next
Call SetGrid
If Combo1.Text = "Regular" Then
    Label7.Caption = "Responsibility Center"
    Call AddOfficebyPeriod(lstOffices)
    lvButtons_H8.Visible = False
ElseIf Combo1.Text = "Casual" Then
    Call AddBatch(lstOffices)
    Label7.Caption = "Batches"
    lvButtons_H8.Visible = True
Else
'MsgBox "NOt Applicable"
End If
End Sub
Public Sub AddOfficebyPeriod(ByVal lstbox As ListBox)
    Dim OfficeTbl As New ADODB.Recordset
    Dim recCounter As Integer
    Dim LoopCounter As Integer
    OfficeTbl.Open "select officecode,OfficeName,officeid from pmis.dbo.OfficeDescription as a inner join pmis.epay.tbl_t_Payroll as b on b.PayOfficeID = a.OfficeID group by officecode,OfficeName,officeid order by OfficeName", opndbaseFMIS, adOpenStatic
    If OfficeTbl.RecordCount > 0 Then
        OfficeTbl.MoveFirst
        lstbox.Clear
        For recCounter = 1 To OfficeTbl.RecordCount
            If Len(OfficeTbl!officecode) > 0 Then
                lstbox.AddItem OfficeTbl!Officename ' & Space(50) & "'" & OfficeTbl!officeid
                lstbox.ItemData(LoopCounter) = OfficeTbl!OfficeID
                LoopCounter = LoopCounter + 1
            End If
        OfficeTbl.MoveNext
        Next
    End If
    OfficeTbl.Close
    Set OfficeTbl = Nothing
End Sub
Private Sub Combo1_Click()
Call Combo1_Change
End Sub

Private Sub DTPicker1_Change()
If Combo1.Text = "Casual" Then
Call AddBatch(lstOffices)
End If
End Sub

Private Sub DTPicker2_Change()
If Combo1.Text = "Casual" Then
Call AddBatch(lstOffices)
End If
End Sub

Private Sub DTPicker3_Change()
 Call List1_Click
End Sub

Private Sub Form_Load()
'On Error GoTo bad
txtaccountcode.Text = accountcode
txttitle.Text = Accountname
'Call LoadOffice
'Load cmbrc
'cmbrc.ListIndex = 1
DTPicker1.Value = Now
DTPicker2.Value = Now
Combo1.ListIndex = 0
Exit Sub
bad:
MsgBox err.description
End Sub

Private Function LoadOffice()
Dim OREc As New ADODB.Recordset
Dim x As Integer
cmbrc.Clear
        OREc.Open ("Select OfficeMedium, [pmisOfficeID] FRom tblREF_AIS_Offices Order By [OfficeMedium]"), opndbaseFMIS, adOpenStatic, adLockOptimistic
        If OREc.RecordCount > 0 Then
            For x = 1 To OREc.RecordCount
                cmbrc.AddItem OREc!OfficeMedium
                cmbrc.ItemData(cmbrc.NewIndex) = OREc!PMISOfficeID
                OREc.MoveNext
            Next x
        End If
        OREc.Close
        Set OREc = Nothing
End Function

Private Sub ListView1_BeforeLabelEdit(Cancel As Integer)

End Sub

Private Sub List1_Click()
Dim Rrec  As New ADODB.Recordset
Dim x
Dim y As Long
lstresult.ListItems.Clear
Set Rrec = opndbasePMIS.Execute("select BatchNo,period from dbo.ztblCasualEmployeePayroll_Transaction where [SID] = '" & List1.SelectedItem.ListSubItems(1).Text & "' and Month_ = " & DTPicker3.Month & " and year_ = " & DTPicker3.Year & "")
If Rrec.RecordCount > 0 Then
    For y = 1 To Rrec.RecordCount
        Set x = lstresult.ListItems.Add(, , Rrec!batchno)
            x.SubItems(1) = Rrec!Period
        Rrec.MoveNext
    Next y
End If
Rrec.Close
Set rec = Nothing
End Sub

Private Sub lvButtons_H1_Click()
Dim x As Integer
Dim xx As Variant
Dim str() As String
Dim lvl As Integer
Dim Code As Long
Dim childcode As String
Dim Debit, Credit As Currency
Dim z
    
    ProgressBar1.Visible = True
    ProgressBar1.Max = MSHFlexGrid1.Rows - 1
    If MsgBox("Are you sure do you want to add into Journal?", vbInformation + vbYesNo, "System Message") = vbYes Then
    'Me.Visible = False
        For x = 1 To MSHFlexGrid1.Rows - 1
        childcode = Trim(txtaccountcode.Text) & "-" & Trim(MSHFlexGrid1.TextMatrix(x, 1))
        xx = Split(Trim(childcode), "-")
        str() = Split(Trim(childcode), "-", -1, vbTextCompare)
        lvl = UBound(xx) + 1
    
        
            If Trim(MSHFlexGrid1.TextMatrix(x, 1)) <> "" Or Trim(MSHFlexGrid1.TextMatrix(x, 2)) <> "" Then
                opndbaseFMIS.Execute ("Exec [fmis].[dbo].[Proc_CheckIfExistSub] @lvl = " & lvl & ",@childcode = '" & childcode & "',@accountcode = '" & Trim(txtaccountcode.Text) & "',@subcode = '" & Trim(MSHFlexGrid1.TextMatrix(x, 1)) & "',@subdesc = '" & Trim(MSHFlexGrid1.TextMatrix(x, 2)) & "'")
                With frmSub3
                    .Picture2.Visible = False
                    .cmbEntry.Visible = False
                    If OptCredit.Value = True Then
                    Credit = Trim(MSHFlexGrid1.TextMatrix(x, 3))
                    Debit = 0
                    Else
                    Credit = 0
                    Debit = Trim(MSHFlexGrid1.TextMatrix(x, 3))
                    End If
                    opndbaseFMIS.Execute "Insert into tblAMIs_tmpjournal (Dvno,Accountcode,Debit,Credit) values ('" & Trim(REFF) & "','" & Trim(childcode) & "','" & Debit & "','" & Credit & "')"
                End With
                DoEvents
                ProgressBar1.Value = x
            End If
        Next x
    End If
    'Call frmSub3.GetSum
    ProgressBar1.Visible = False
    If MsgBox("Import more?", vbInformation + vbYesNo, "System Confirmation") = vbNo Then
    Unload Me
    End If
End Sub
Private Function LoadAccountsByName(ByVal accountcode As String, ByVal Condition As String)
Dim ARec As New ADODB.Recordset
Dim x As Integer
    ARec.Open "exec Proc_getNamebychildCode @childaccountcode = '" & accountcode & "', @Condition = '" & Condition & "'", opndbaseFMIS, adOpenStatic
        If ARec.RecordCount > 0 Then
            LoadAccountsByName = ARec!Accountfullname
        inRec = True
        End If
    ARec.Close
    Set ARec = Nothing
End Function
Private Sub lvButtons_H2_Click()
Dim rec As New ADODB.Recordset
On Error GoTo bad
'rec.Open "Exec MPproc_LoadSubEntries @whatsystem = 'Payroll',@accountcode = '" & Trim(txtaccountcode.Text) & "',@year = " & DTPicker1.Year & ",@month= " & DTPicker2.Month & ",@office = " & cmbrc.ItemData(cmbrc.ListIndex) & "", opndbaseFMIS, adOpenStatic, adLockOptimistic
Call lvButtons_H3_Click
'MsgBox txtquery.Text
rec.Open txtquery.Text, opndbaseFMIS, adOpenStatic, adLockOptimistic
    MSHFlexGrid1.Clear
    MSHFlexGrid1.Rows = 2
    MSHFlexGrid1.Cols = 4
    If rec.RecordCount > 0 Then
    Set MSHFlexGrid1.DataSource = rec
    End If
    lblResult.Caption = rec.RecordCount & " Record(s) Found"
    Call SetGrid
    Call GettotalAMount
    rec.Close
Exit Sub
bad:
If err.Number = 3704 Then
    MsgBox "No Record Found", vbInformation, "System Message"
Else
MsgBox "Noted: " & err.description, , "System Information"
End If
End Sub
    
Private Sub SetGrid()
On Error Resume Next
Dim cc As Integer
    'MSHFlexGrid1.Clear
     ' IIf(LCase(Trim(lblMode)) = "edit", 6, 5)
    'Name
    
    MSHFlexGrid1.TextMatrix(0, 0) = "ID"
    MSHFlexGrid1.TextMatrix(0, 1) = "ID"
    MSHFlexGrid1.TextMatrix(0, 2) = "Name"
    MSHFlexGrid1.TextMatrix(0, 3) = "Amount"
    
    
    MSHFlexGrid1.ColWidth(0) = 0
    MSHFlexGrid1.ColWidth(1) = 1700
    MSHFlexGrid1.ColWidth(2) = 4000
    MSHFlexGrid1.ColWidth(3) = 1500
    MSHFlexGrid1.ColAlignment(1) = 1
End Sub
Private Function GettotalAMount()
Dim x As Integer
Dim Gamount As Currency
Gamount = 0
For x = 1 To MSHFlexGrid1.Rows - 1
    Gamount = Gamount + IIf((MSHFlexGrid1.TextMatrix(x, 3) = ""), 0, MSHFlexGrid1.TextMatrix(x, 3))
    MSHFlexGrid1.TextMatrix(x, 3) = Format(CCur(IIf((MSHFlexGrid1.TextMatrix(x, 3) = ""), 0, MSHFlexGrid1.TextMatrix(x, 3))), "#,##0.00")
Next x
txtgamount.Text = Format(Gamount, "#,##0.00")
End Function

Private Sub lvButtons_H3_Click()
Dim rec As New ADODB.Recordset
Dim brec As New ADODB.Recordset
On Error GoTo bad
Dim tmpquery, batchquery As String

Set rec = opndbaseFMIS.Execute("Select Query from tblAMIS_Qrygenerator4COA where acountcode = '" & txtaccountcode.Text & "' and trnno = '" & lblID.Caption & "' and actioncode = 1")
    If rec.RecordCount > 0 Then
        
        
        tmpquery = IIf(IsNull(rec.Fields!query), "", rec.Fields!query)
'         MsgBox tmpquery
        'replace Year
        tmpquery = Replace(tmpquery, "@year", DTPicker1.Year)
        'replace month
        tmpquery = Replace(tmpquery, "@Month", DTPicker2.Month)
        'Replace office
        'If Combo1.Text = "Regular" Then
        tmpquery = Replace(tmpquery, "= @office", " in (" & GetallOfficeID & ")")
        'ElseIf Combo1.Text = "Casual" Then
        
        'End If
        tmpquery = Replace(tmpquery, "@AccountCode", "'" & txtaccountcode.Text & "'")
        txtquery.Text = Trim(tmpquery)
        
      ' MsgBox tmpquery
        'execute query
    End If
Set rec = Nothing
Exit Sub
bad:
End Sub
Public Function AddBatch(ByVal lstbox As ListBox)
    Dim OfficeTbl As New ADODB.Recordset
    Dim brec As New ADODB.Recordset
    Dim recCounter As Long
    Dim LoopCounter As Long
    Dim batchquery As String
    Dim sql As String
    lstbox.Clear
    Set brec = opndbaseFMIS.Execute("select * from tblAMIS_SqlBatchno where accountcode = '" & txtaccountcode.Text & "' and actioncode =1")
    If brec.RecordCount > 0 Then
        batchquery = brec!query
        batchquery = Replace(batchquery, "@year", DTPicker1.Year)
        'replace month
        batchquery = Replace(batchquery, "@Month", DTPicker2.Month)
        
        OfficeTbl.Open Replace(batchquery, "'", "''"), opndbaseFMIS, adOpenStatic
    If OfficeTbl.RecordCount > 0 Then
        For recCounter = 1 To OfficeTbl.RecordCount
            If Len(OfficeTbl!batchno) > 0 Then
                lstbox.AddItem OfficeTbl!batchno ' & Space(50) & "'" & OfficeTbl!officeid
                LoopCounter = LoopCounter + 1
            End If
        OfficeTbl.MoveNext
        Next
    End If
    OfficeTbl.Close
    Set OfficeTbl = Nothing
    End If
End Function

Private Function GetallOfficeID() As String
Dim x As Long
Dim officeall As String
For x = 0 To lstOffices.ListCount - 1
    If lstOffices.Selected(x) = True Then
        If Combo1.Text = "Regular" Then
            If officeall <> "" Then
            officeall = officeall & "," & lstOffices.ItemData(x)
            Else
            officeall = lstOffices.ItemData(x)
            End If
        ElseIf Combo1.Text = "Casual" Then
            If officeall <> "" Then
                officeall = officeall & "," & lstOffices.List(x)
            Else
            officeall = lstOffices.List(x)
            End If
        ElseIf Combo1.Text = "JOB Order" Then
            If officeall <> "" Then
                officeall = officeall & "," & lstOffices.List(x)
            Else
            officeall = lstOffices.List(x)
            End If
        ElseIf Combo1.Text = "Contractual" Then
            If officeall <> "" Then
                officeall = officeall & "," & lstOffices.List(x)
            Else
            officeall = lstOffices.List(x)
            End If
         ElseIf Combo1.Text = "SP Scholar" Then
            If officeall <> "" Then
                officeall = officeall & "," & lstOffices.List(x)
            Else
            officeall = lstOffices.List(x)
            End If
        End If
    End If
Next x
GetallOfficeID = officeall
End Function
Private Sub lvButtons_H5_Click()
If Combo1.Text <> "" Then
Set frm_COAQueryGenerator_load.frm = Me
frm_COAQueryGenerator_load.TYP = Combo1.Text
centerme frm_COAQueryGenerator_load

frm_COAQueryGenerator_load.Show 1
If Combo1.Text = "Casual" Then
'Call AddBatch(lstOffices)
End If
Call lvButtons_H3_Click

Else
MsgBox "Oop..!Please Select Employee type..!", vbCritical, "System Message"
End If
End Sub

Private Sub lvButtons_H6_Click()
Dim www As Integer
   ' If chkSelected.Value = 1 Then Exit Sub
    For www = 0 To lstOffices.ListCount - 1
        lstOffices.Selected(www) = True
    Next www
    lstOffices.ListIndex = 0
End Sub

Private Sub lvButtons_H7_Click()
 Dim www As Integer
'    If chkSelected.Value = 1 Then Exit Sub
    For www = 0 To lstOffices.ListCount - 1
        lstOffices.Selected(www) = False
    Next www
    lstOffices.ListIndex = 0
End Sub

Private Sub lvButtons_H8_Click()
Call AddBatch(lstOffices)
End Sub
Private Function centerme(ByVal frm As Form)
Dim H, w, FW, FFW, FH, FFH, x, y As Long
frm.ScaleMode = 5
H = MDIFrm_MAIN.Height
FH = frm.Height
x = frm.ScaleHeight / 2
FFH = (H - FH) / x

w = MDIFrm_MAIN.Width
y = frm.ScaleWidth / 2
FW = frm.Width
FFW = (w - FW)

frm.Top = FFH / 2
frm.Left = FFW / 2
End Function

Private Sub txtfind_Change()
Dim rec As New ADODB.Recordset
Dim x As Long
Dim y
Set rec = opndbasePMIS.Execute("select SwipEmployeeID as ID,rtrim(Lastname) + ',' + rtrim(firstname) + ' ' + mi as Name from Employee where lastname like '%" & txtfind.Text & "%'")
List1.ListItems.Clear
If rec.RecordCount > 0 Then
    For x = 1 To rec.RecordCount
        Set y = List1.ListItems.Add(, , Trim(rec!name))
        y.SubItems(1) = IIf(IsNull(rec!id), "", rec!id)
        rec.MoveNext
    Next x
End If
rec.Close
Set rec = Nothing
End Sub

VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{1693405E-2DC9-4248-B52F-4AC9145DA2AF}#1.0#0"; "WINXPC~1.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frm_coa 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8775
   ClientLeft      =   4005
   ClientTop       =   1965
   ClientWidth     =   13455
   Icon            =   "frm_coa.frx":0000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8775
   ScaleWidth      =   13455
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chksave 
      Caption         =   "Save and Close"
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
      Left            =   10320
      TabIndex        =   12
      Top             =   240
      Width           =   1815
   End
   Begin VB.TextBox Text1 
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
      Height          =   375
      Left            =   11280
      TabIndex        =   9
      Top             =   9240
      Visible         =   0   'False
      Width           =   1935
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
      Height          =   375
      Left            =   8400
      TabIndex        =   8
      Top             =   9240
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   0
      Top             =   0
   End
   Begin WinXPC_Engine.WindowsXPC WindowsXPC1 
      Left            =   6240
      Top             =   9360
      _ExtentX        =   6588
      _ExtentY        =   1085
      ColorScheme     =   2
      Common_Dialog   =   0   'False
   End
   Begin VB.PictureBox freeSizer1 
      Height          =   480
      Left            =   5040
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   7
      Top             =   9120
      Width           =   1200
   End
   Begin lvButton.lvButtons_H lvButtons_H2 
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   9120
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   1085
      Caption         =   "&Auto Generate"
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
      Image           =   "frm_coa.frx":076A
      cBack           =   16777215
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7695
      Left            =   120
      TabIndex        =   13
      Top             =   960
      Width           =   13215
      _ExtentX        =   23310
      _ExtentY        =   13573
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "JEV Entry"
      TabPicture(0)   =   "frm_coa.frx":4274
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Picture1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Cash Flow Entry"
      TabPicture(1)   =   "frm_coa.frx":4290
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label11"
      Tab(1).Control(1)=   "Picture3"
      Tab(1).ControlCount=   2
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   6615
         Left            =   -74880
         ScaleHeight     =   6585
         ScaleWidth      =   12945
         TabIndex        =   29
         Top             =   840
         Width           =   12975
         Begin VB.PictureBox Picture4 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   6255
            Left            =   1200
            ScaleHeight     =   6225
            ScaleWidth      =   8610
            TabIndex        =   32
            Top             =   300
            Visible         =   0   'False
            Width           =   8635
            Begin VB.CheckBox Check2 
               BackColor       =   &H80000005&
               Caption         =   "Many"
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
               TabIndex        =   35
               Top             =   6720
               Visible         =   0   'False
               Width           =   855
            End
            Begin VB.TextBox txtdetails2 
               Appearance      =   0  'Flat
               Height          =   495
               Left            =   840
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   34
               Top             =   75
               Width           =   7215
            End
            Begin VB.TextBox txtfind2 
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
               Left            =   3240
               TabIndex        =   33
               Top             =   5760
               Width           =   3375
            End
            Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid2 
               Height          =   5055
               Left            =   120
               TabIndex        =   36
               Top             =   600
               Width           =   8385
               _ExtentX        =   14790
               _ExtentY        =   8916
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
            Begin lvButton.lvButtons_H lvButtons_H4 
               Height          =   375
               Left            =   8160
               TabIndex        =   37
               Top             =   120
               Width           =   375
               _ExtentX        =   661
               _ExtentY        =   661
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
               Image           =   "frm_coa.frx":42AC
               cBack           =   16777215
            End
            Begin VB.Label Label10 
               BackColor       =   &H80000005&
               Caption         =   "Details:"
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
               TabIndex        =   40
               Top             =   120
               Width           =   855
            End
            Begin VB.Label Label9 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "Press ENTER "
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   495
               Left            =   6840
               TabIndex        =   39
               Top             =   5715
               Width           =   1215
            End
            Begin VB.Label Label8 
               BackColor       =   &H80000005&
               Caption         =   "Search Name:"
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
               Left            =   1920
               TabIndex        =   38
               Top             =   5805
               Width           =   1335
            End
         End
         Begin VB.TextBox txt_Entry2 
            Appearance      =   0  'Flat
            BackColor       =   &H0080FFFF&
            BorderStyle     =   0  'None
            Height          =   525
            Left            =   10320
            TabIndex        =   31
            Top             =   3600
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.ComboBox cmbEntry2 
            BackColor       =   &H0080FFFF&
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
            Left            =   10320
            TabIndex        =   30
            Text            =   "cmbEntry"
            Top             =   1920
            Visible         =   0   'False
            Width           =   1695
         End
         Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
            Height          =   6840
            Left            =   0
            TabIndex        =   42
            Top             =   0
            Width           =   12960
            _ExtentX        =   22860
            _ExtentY        =   12065
            _Version        =   393216
            FixedCols       =   0
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
         End
         Begin MSComctlLib.ListView ListView1 
            Height          =   6495
            Left            =   2280
            TabIndex        =   41
            Top             =   420
            Visible         =   0   'False
            Width           =   9375
            _ExtentX        =   16536
            _ExtentY        =   11456
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
            NumItems        =   2
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Code"
               Object.Width           =   1235
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Explaination"
               Object.Width           =   14111
            EndProperty
         End
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   7095
         Left            =   120
         ScaleHeight     =   7065
         ScaleWidth      =   12945
         TabIndex        =   14
         Top             =   480
         Width           =   12975
         Begin VB.ComboBox cmbEntry 
            BackColor       =   &H0080FFFF&
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
            Left            =   10320
            TabIndex        =   26
            Text            =   "cmbEntry"
            Top             =   1920
            Visible         =   0   'False
            Width           =   1695
         End
         Begin VB.TextBox txt_entry 
            Appearance      =   0  'Flat
            BackColor       =   &H0080FFFF&
            BorderStyle     =   0  'None
            Height          =   525
            Left            =   10320
            TabIndex        =   25
            Top             =   3600
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.PictureBox Picture2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   6615
            Left            =   1200
            ScaleHeight     =   6585
            ScaleWidth      =   8610
            TabIndex        =   15
            Top             =   300
            Visible         =   0   'False
            Width           =   8635
            Begin VB.TextBox txtfind 
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
               TabIndex        =   18
               Top             =   6120
               Width           =   3375
            End
            Begin VB.TextBox txtdetails 
               Appearance      =   0  'Flat
               Height          =   495
               Left            =   840
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   17
               Top             =   75
               Width           =   7215
            End
            Begin VB.CheckBox Check1 
               BackColor       =   &H80000005&
               Caption         =   "Many"
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
               TabIndex        =   16
               Top             =   6720
               Visible         =   0   'False
               Width           =   855
            End
            Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
               Height          =   5415
               Left            =   120
               TabIndex        =   19
               Top             =   600
               Width           =   8385
               _ExtentX        =   14790
               _ExtentY        =   9551
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
            Begin lvButton.lvButtons_H lvButtons_H3 
               Height          =   375
               Left            =   8160
               TabIndex        =   20
               Top             =   120
               Width           =   375
               _ExtentX        =   661
               _ExtentY        =   661
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
               Image           =   "frm_coa.frx":4406
               cBack           =   16777215
            End
            Begin lvButton.lvButtons_H lvButtons_H5 
               Height          =   375
               Left            =   120
               TabIndex        =   21
               Top             =   6120
               Width           =   855
               _ExtentX        =   1508
               _ExtentY        =   661
               Caption         =   "Import"
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
               cFore           =   0
               cFHover         =   33023
               cBhover         =   8438015
               LockHover       =   3
               cGradient       =   33023
               Gradient        =   3
               CapStyle        =   1
               Mode            =   0
               Value           =   0   'False
               Image           =   "frm_coa.frx":4560
               cBack           =   16777215
            End
            Begin lvButton.lvButtons_H lvButtons_H6 
               Height          =   375
               Left            =   1080
               TabIndex        =   45
               Top             =   6120
               Width           =   1695
               _ExtentX        =   2990
               _ExtentY        =   661
               Caption         =   "Generate Payable"
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
               cFore           =   0
               cFHover         =   33023
               cBhover         =   8438015
               LockHover       =   3
               cGradient       =   33023
               Gradient        =   3
               CapStyle        =   1
               Mode            =   0
               Value           =   0   'False
               Image           =   "frm_coa.frx":57E2
               cBack           =   16777215
            End
            Begin VB.Label Label2 
               BackColor       =   &H80000005&
               Caption         =   "Search Name:"
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
               Left            =   2880
               TabIndex        =   24
               Top             =   6165
               Width           =   1335
            End
            Begin VB.Label Label3 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "Press ENTER "
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   495
               Left            =   7800
               TabIndex        =   23
               Top             =   6075
               Width           =   1215
            End
            Begin VB.Label Label4 
               BackColor       =   &H80000005&
               Caption         =   "Details:"
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
               TabIndex        =   22
               Top             =   120
               Width           =   855
            End
         End
         Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
            Height          =   7080
            Left            =   0
            TabIndex        =   28
            Top             =   0
            Width           =   12960
            _ExtentX        =   22860
            _ExtentY        =   12488
            _Version        =   393216
            FixedCols       =   0
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
         End
         Begin MSComctlLib.ListView ListView2 
            Height          =   6495
            Left            =   2400
            TabIndex        =   27
            Top             =   300
            Visible         =   0   'False
            Width           =   9375
            _ExtentX        =   16536
            _ExtentY        =   11456
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
            NumItems        =   2
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Code"
               Object.Width           =   1235
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Explaination"
               Object.Width           =   14111
            EndProperty
         End
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Cash flow Entry"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   -74880
         TabIndex        =   43
         Top             =   410
         Width           =   12855
      End
   End
   Begin VB.TextBox txtformula 
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
      Left            =   4440
      TabIndex        =   5
      Top             =   2160
      Visible         =   0   'False
      Width           =   7575
   End
   Begin lvButton.lvButtons_H lvButtons_H1 
      Height          =   495
      Left            =   12240
      TabIndex        =   44
      Top             =   240
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   873
      Caption         =   "&OK"
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
      Image           =   "frm_coa.frx":92EC
      cBack           =   16777215
   End
   Begin VB.Label Label7 
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
      Left            =   10440
      TabIndex        =   11
      Top             =   9270
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label6 
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
      Left            =   7560
      TabIndex        =   10
      Top             =   9270
      Visible         =   0   'False
      Width           =   855
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
      Left            =   4200
      TabIndex        =   6
      Top             =   1320
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label lblname 
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
      Left            =   1680
      TabIndex        =   4
      Top             =   120
      Width           =   6255
   End
   Begin VB.Label Label5 
      Caption         =   "Claimant Name:"
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
      TabIndex        =   3
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label lblamount 
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
      Left            =   1680
      TabIndex        =   2
      Top             =   480
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "Gross Amount:"
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
      TabIndex        =   1
      Top             =   480
      Width           =   1575
   End
   Begin VB.Menu popup 
      Caption         =   "popup"
      Visible         =   0   'False
      Begin VB.Menu Payroll 
         Caption         =   "Payroll"
      End
      Begin VB.Menu Property 
         Caption         =   "Property"
      End
      Begin VB.Menu AP 
         Caption         =   "Accounts Payable"
      End
   End
   Begin VB.Menu EQ 
      Caption         =   "Execute Query"
      Begin VB.Menu APs 
         Caption         =   "Accounts Payable"
      End
      Begin VB.Menu PC 
         Caption         =   "Post Closing Entry"
      End
   End
End
Attribute VB_Name = "frm_coa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public REFF, datetimeentered, UserID, Accountname, CName As String
Public Damount, Camount, Gamount As Currency
Public Damount2, Camount2, Gamount2 As Currency
Public isEdit, inRec, ifCMB, IfEdit, IfNew, insert, delete, bypass As Boolean
Public Transtype, WhatTab As Integer '
Public tabindex As Integer
Public GetCash As Boolean
Public isPOSTED As Boolean
Public frm As Form

Private Sub SetGrid()
Dim cc As Integer

    MSFlexGrid1.Clear
    MSFlexGrid1.Rows = 2
    MSFlexGrid1.Cols = 6 ' IIf(LCase(Trim(lblMode)) = "edit", 6, 5)
    
    'Name
    MSFlexGrid1.TextMatrix(0, 0) = "ID"
    MSFlexGrid1.TextMatrix(0, 1) = "AccountCode"
    MSFlexGrid1.TextMatrix(0, 2) = "Accounts and Explanation"
    MSFlexGrid1.TextMatrix(0, 3) = "Debit"
    MSFlexGrid1.TextMatrix(0, 4) = "Credit"
    MSFlexGrid1.TextMatrix(0, 5) = "ActionCode"
    
    MSFlexGrid1.ColWidth(0) = 0
    MSFlexGrid1.ColWidth(1) = 1700
    MSFlexGrid1.ColWidth(2) = 8000
    MSFlexGrid1.ColWidth(3) = 1500
    MSFlexGrid1.ColWidth(4) = 1500
    MSFlexGrid1.ColWidth(5) = 0
    MSFlexGrid1.ColAlignment(1) = 1
    
End Sub
Private Sub SetGrid2()
Dim cc As Integer

    MSFlexGrid2.Clear
    MSFlexGrid2.Rows = 2
    MSFlexGrid2.Cols = 6 ' IIf(LCase(Trim(lblMode)) = "edit", 6, 5)
    
    'Name
    MSFlexGrid2.TextMatrix(0, 0) = "ID"
    MSFlexGrid2.TextMatrix(0, 1) = "AccountCode"
    MSFlexGrid2.TextMatrix(0, 2) = "Accounts and Explanation"
    MSFlexGrid2.TextMatrix(0, 3) = "Debit"
    MSFlexGrid2.TextMatrix(0, 4) = "Credit"
    MSFlexGrid2.TextMatrix(0, 5) = "ActionCode"
    
    MSFlexGrid2.ColWidth(0) = 0
    MSFlexGrid2.ColWidth(1) = 1700
    MSFlexGrid2.ColWidth(2) = 8000
    MSFlexGrid2.ColWidth(3) = 1500
    MSFlexGrid2.ColWidth(4) = 1500
    MSFlexGrid2.ColWidth(5) = 0
    MSFlexGrid2.ColAlignment(1) = 1
    
End Sub


Private Sub GotonextCell()
'txt_entry.Move MSFlexGrid1.CellLeft(MSHFlexGrid1.Row, 3), MSFlexGrid1.CellTop, MSFlexGrid1.CellWidth, MSFlexGrid1.CellHeight
End Sub

Public Function GetAccountNamebyorder(ByVal Condition As String)
Dim rec As New ADODB.Recordset
Dim x
Dim z As Integer
rec.Open "SELECT  [code],[AccountChildName] as Explanation,[ChartAccountChildID] FROM [fmis].[Accounting].[tbl_l_ChartOfAccountsChild] where [AccountChildParentID] is null and code like '" & cmbEntry.Text & "%' and AccountChildName like '" & Trim(txtfind.Text) & "%' order by AccountChildName", opndbaseFMIS, adOpenStatic, adLockOptimistic
    'lst.ListItems.Clear
        MSHFlexGrid1.Clear
        MSHFlexGrid1.Rows = 2
    If rec.RecordCount > 0 Then
    
    Set MSHFlexGrid1.DataSource = rec
        MSHFlexGrid1.Cols = 4
        MSHFlexGrid1.TextMatrix(0, 0) = "ChartAccountChildID"
        MSHFlexGrid1.TextMatrix(0, 1) = "Code"
        MSHFlexGrid1.TextMatrix(0, 2) = "Explanation"
        MSHFlexGrid1.TextMatrix(0, 3) = "Formula"
        
        MSHFlexGrid1.ColWidth(0) = 0
        MSHFlexGrid1.ColWidth(1) = 1000
        MSHFlexGrid1.ColWidth(2) = 8000
        MSHFlexGrid1.ColWidth(3) = 0
        
        
    End If
'rec.Close
Set rec = Nothing
End Function

Public Function GetAccountNamebyorder2(ByVal Condition As String)
Dim rec As New ADODB.Recordset
Dim x
Dim z As Integer
rec.Open "Select Accountcode,Accountname from tblREF_AIS_ChartOfAccountsMother where accountcode like '" & cmbEntry2.Text & "%' and accountname like '" & Trim(txtfind2.Text) & "%' order by Accountname", opndbaseFMIS, adOpenStatic, adLockOptimistic
    'lst.ListItems.Clear
        MSHFlexGrid2.Clear
        MSHFlexGrid2.Rows = 2
    If rec.RecordCount > 0 Then
    
'        For z = 1 To rec.RecordCount
'                    Set x = lst.ListItems.Add(, , rec.Fields!Accountcode)
'                    x.SubItems(1) = Trim(rec.Fields!Accountname)
'            rec.MoveNext
'        Next z
    
    Set MSHFlexGrid2.DataSource = rec
        MSHFlexGrid2.Cols = 4
        MSHFlexGrid2.TextMatrix(0, 1) = "Code"
        MSHFlexGrid2.TextMatrix(0, 2) = "Explanation"
        MSHFlexGrid2.TextMatrix(0, 3) = "Formula"
        
        MSHFlexGrid2.ColWidth(0) = 0
        MSHFlexGrid2.ColWidth(1) = 700
        MSHFlexGrid2.ColWidth(2) = 8000
        MSHFlexGrid2.ColWidth(3) = 0
        
        
    End If
'rec.Close
Set rec = Nothing
End Function

Private Sub AP_Click()
With frm_AP_ImportInJEV
.REFF = REFF
.Show 1
GetAccntngEntries (REFF)
End With
'centerme (frm_AP_ImportInJEV)
End Sub

Private Sub cmbEntry_Change()
Call GetAccountNamebyorder("Accountcode")
End Sub
Private Sub cmbEntry2_Change()
Call GetAccountNamebyorder2("Accountcode")
End Sub

Private Sub cmbEntry_Click()
    Call GetAccountNamebyorder("Accountname")
End Sub
Private Sub cmbEntry2_Click()
    Call GetAccountNamebyorder2("Accountname")
End Sub

Private Sub cmbEntry_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       ' If cmbEntry.ListIndex <> -1 Then
            inRec = False
            Accountname = LoadAccountsByName(cmbEntry.Text, "Summary")
            If Trim(cmbEntry.Text) <> "" Then
                If inRec = False Then
                    If cmbEntry.Text <> MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 1) Then
                        MsgBox "Invalid Accountcode Please Select Another Account..!", vbCritical, "System Information"
                    Exit Sub
                    End If
                End If
                ifCMB = True
                If Chckentry = False Then
                Exit Sub
                End If
                ifCMB = False
            End If
            
            
            
            
        If cmbEntry.Text = "" Then
            If MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2) <> "TOTAL" Then
                If MsgBox("Are you sure do you want do remove this Account and its Contain?", vbCritical + vbYesNo, "System Message") = vbYes Then
                opndbaseFMIS.Execute "Delete from tblAMIS_tmpjournal where accountcode like '" & Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 1)) & "%' and dvno = '" & REFF & "'"
                    
                    MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 1) = cmbEntry.Text
                    
                    If MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2) = "TOTAL" Then
                            MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 3) = ""
                            MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 4) = ""
                            
                    Else
                        If MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2) = "" Then
                        MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1
                        End If
                        MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2) = Accountname
                    End If
                    
                    If MSFlexGrid1.TextMatrix(MSFlexGrid1.Rows - 1, 2) <> "TOTAL" Then
                        MSFlexGrid1.Rows = MSFlexGrid1.Rows - 1
                    End If
                    MSFlexGrid1.RemoveItem (MSFlexGrid1.Row)
                    
                End If
            
                'MSFlexGrid1.Rows = MSFlexGrid1.Rows - 1
            End If
        Else
            If cmbEntry.Text <> Left(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 1), 3) Then
                If MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 1) <> "" Then
                    If MsgBox("Are you sure do you want do Change the Account " & MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 1) & " to " & cmbEntry.Text & "?", vbCritical + vbYesNo, "System Message") = vbYes Then
                    opndbaseFMIS.Execute "Delete from tblAMIS_tmpjournal where accountcode like '" & Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 1)) & "%' and dvno = '" & REFF & "'"
                    Else
                    Exit Sub
                    End If
                End If
            End If
            MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 1) = cmbEntry.Text
                    
                    If MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2) = "TOTAL" Then
                            MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 3) = ""
                            MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 4) = ""
                            
                    Else
                        If MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2) = "" Then
                        MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1
                        End If
                        MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2) = Accountname
                    End If
            
            If MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2) <> "TOTAL" Then
                MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2) = Accountname
                    If MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2) = "" Then
                    MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1
                    End If
                MSFlexGrid1.col = 3
                Call MSFlexGrid1_Click
            Else
            MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1
            MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2) = Accountname
            
            
            MSFlexGrid1.col = 3
            Call MSFlexGrid1_Click
            End If
            
        End If
        
            
            
        cmbEntry.Visible = False
        ListView2.Visible = False
        Picture2.Visible = False
        Call GetSum
        MSFlexGrid1.SetFocus
        isEdit = True
    Else
'       Call GetAccountNamebyorder("Accountcode")
        txtfind.Text = ""
        Picture2.Move MSFlexGrid1.CellLeft + cmbEntry.Width
        Picture2.Visible = True
    End If
End Sub
Private Function ExistsInJEVEntry(accountcode As String) As Boolean
Dim cnt As Integer
cnt = MSFlexGrid1.Rows - 1
ExistsInJEVEntry = False
For x = 0 To cnt
    If accountcode = MSFlexGrid1.TextMatrix(x, 1) Then
    ExistsInJEVEntry = True
        Exit Function
    End If
Next x
End Function
Private Sub cmbEntry2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       ' If cmbEntry.ListIndex <> -1 Then
       If isPOSTED = True Then
            MsgBox "Unable to edit the entry, the transaction is already approved..", vbInformation, "System Information"
            txt_Entry2.Visible = False
            Exit Sub
        End If
            inRec = False
            Accountname = LoadAccountsByName(cmbEntry2.Text, "Summary")
            If Trim(cmbEntry2.Text) <> "" Then
                If inRec = False Then
                    If cmbEntry2.Text <> MSFlexGrid2.TextMatrix(MSFlexGrid2.Row, 1) Then
                    MsgBox "Invalid Accountcode Please Select Another Account..!", vbCritical, "System Information"
                    Exit Sub
                    End If
                End If
                ifCMB = True
                If Chckentry = False Then
                    Exit Sub
                End If
                If ExistsInJEVEntry(cmbEntry2.Text) = False Then
                    MsgBox "Invalid Accountcode Please Select Another Account..!", vbCritical, "System Information"
                    Exit Sub
                End If
                ifCMB = False
            End If
            
            
            
            
        If cmbEntry2.Text = "" Then
            If MSFlexGrid2.TextMatrix(MSFlexGrid2.Row, 2) <> "TOTAL" Then
                If MsgBox("Are you sure do you want do remove this Account?", vbCritical + vbYesNo, "System Message") = vbYes Then
                opndbaseFMIS.Execute "Delete from [tblAMIS_PostedJEVforCashflow] where accountcode ='" & Trim(MSFlexGrid2.TextMatrix(MSFlexGrid2.Row, 1)) & "' and Jevno = '" & REFF & "'"
                    
                    MSFlexGrid2.TextMatrix(MSFlexGrid2.Row, 1) = cmbEntry2.Text
                    
                    If MSFlexGrid2.TextMatrix(MSFlexGrid2.Row, 2) = "TOTAL" Then
                            MSFlexGrid2.TextMatrix(MSFlexGrid2.Row, 3) = ""
                            MSFlexGrid2.TextMatrix(MSFlexGrid2.Row, 4) = ""
                            
                    Else
                        If MSFlexGrid2.TextMatrix(MSFlexGrid2.Row, 2) = "" Then
                        MSFlexGrid2.Rows = MSFlexGrid2.Rows + 1
                        End If
                        MSFlexGrid2.TextMatrix(MSFlexGrid2.Row, 2) = Accountname
                    End If
                    
                    If MSFlexGrid2.TextMatrix(MSFlexGrid2.Rows - 1, 2) <> "TOTAL" Then
                        MSFlexGrid2.Rows = MSFlexGrid2.Rows - 1
                    End If
                    MSFlexGrid2.RemoveItem (MSFlexGrid2.Row)
                    
                End If
            
                'MSFlexGrid2.Rows = MSFlexGrid2.Rows - 1
            End If
        Else
            If cmbEntry2.Text <> Left(MSFlexGrid2.TextMatrix(MSFlexGrid2.Row, 1), 3) Then
                If MSFlexGrid2.TextMatrix(MSFlexGrid2.Row, 1) <> "" Then
                    If MsgBox("Are you sure do you want do Change the Account " & MSFlexGrid2.TextMatrix(MSFlexGrid2.Row, 1) & " to " & cmbEntry2.Text & "?", vbCritical + vbYesNo, "System Message") = vbYes Then
                    opndbaseFMIS.Execute "Delete from [tblAMIS_PostedJEVforCashflow] where accountcode = '" & Trim(MSFlexGrid2.TextMatrix(MSFlexGrid2.Row, 1)) & "' and jevno = '" & REFF & "'"
                    Else
                    Exit Sub
                    End If
                End If
            End If
            MSFlexGrid2.TextMatrix(MSFlexGrid2.Row, 1) = cmbEntry2.Text
                    
                    If MSFlexGrid2.TextMatrix(MSFlexGrid2.Row, 2) = "TOTAL" Then
                        MSFlexGrid2.TextMatrix(MSFlexGrid2.Row, 3) = ""
                        MSFlexGrid2.TextMatrix(MSFlexGrid2.Row, 4) = ""
                    Else
                        If MSFlexGrid2.TextMatrix(MSFlexGrid2.Row, 2) = "" Then
                        MSFlexGrid2.Rows = MSFlexGrid2.Rows + 1
                        End If
                        MSFlexGrid2.TextMatrix(MSFlexGrid2.Row, 2) = Accountname
                    End If
            
            If MSFlexGrid2.TextMatrix(MSFlexGrid2.Row, 2) <> "TOTAL" Then
                MSFlexGrid2.TextMatrix(MSFlexGrid2.Row, 2) = Accountname
                    If MSFlexGrid2.TextMatrix(MSFlexGrid2.Row, 2) = "" Then
                    MSFlexGrid2.Rows = MSFlexGrid2.Rows + 1
                    End If
                MSFlexGrid2.col = 3
            Else
            MSFlexGrid2.Rows = MSFlexGrid2.Rows + 1
            MSFlexGrid2.TextMatrix(MSFlexGrid2.Row, 2) = Accountname
            
            
            MSFlexGrid2.col = 3
            End If
            
        End If
        
            
            
        cmbEntry2.Visible = False
        ListView2.Visible = False
        Picture4.Visible = False
        Call GetSum2
        MSFlexGrid2.SetFocus
        isEdit = True
    Else
'       Call GetAccountNamebyorder("Accountcode")
        txtfind.Text = ""
        Picture4.Move MSFlexGrid2.CellLeft + cmbEntry2.Width
        Picture4.Visible = True
    End If
End Sub
Private Sub Form_Load()
SSTab1.Tab = WhatTab
Call SetGrid
lblname.Caption = CName
lblamount.Caption = Format(Gamount, "#,##0.00")
If isEdit = True Then
    Call GetAccntngEntries(REFF)
    Call GetAccntngEntries2(REFF)
End If
'Call LoadAccountsByFund(Accountcode)
End Sub
'Private Function LoadDetails(ByVal reff As String)
'Dim DRec As New ADODB.Recordset
'DRec.Open ("Select trnno ,ChildAccountcode, Debit,credit,actioncode,datetimeentered,userid From tblAMIS_AccoutingEntries Where [reffno]='" & reff & "' And (ActionCode=1)"), opndbaseFMIS, adOpenStatic, adLockOptimistic
'    Call SetGrid
'    If DRec.RecordCount > 0 Then
'
'        datetimeentered = DRec!datetimeentered
'        UserID = DRec!UserID
'        For x = 1 To DRec.RecordCount
'            MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1
'            MSFlexGrid1.TextMatrix(x, 0) = DRec![Trnno]
'            MSFlexGrid1.TextMatrix(x, 1) = DRec!childaccountcode
'            MSFlexGrid1.TextMatrix(x, 2) = LoadAccountsByName(DRec!childaccountcode, "Summary")
'            MSFlexGrid1.TextMatrix(x, 4) = IIf((Format(DRec!Credit, "#,##0.00") = "0.00"), "", Format(DRec!Credit, "#,##0.00"))
'            MSFlexGrid1.TextMatrix(x, 3) = IIf((Format(DRec!Debit, "#,##0.00") = "0.00"), "", Format(DRec!Debit, "#,##0.00"))
'            'MSFlexGrid1.TextMatrix(x, 1) = DRec!ActionCode
'            DRec.MoveNext
'        Next x
'        Call GetSum
'    Else
'    MSFlexGrid1.TextMatrix(1, 2) = "TOTAL"
'    End If
'    DRec.Close
'    Set DRec = Nothing
'
'End Function
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
    y = CCur(y) + CCur(str(x))
 Next x
 Exit Function
sum:
 If err.Number = 9 Then
 sumAmount = y
Else
MsgBox "Incorrect Format", vbInformation, "System Message"
End If
End Function
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If isEdit = True Then
    If ChkEntry = True Then
        If chksave.Value = 1 Then
            If CheckIfADMIN(ActiveUserID) = True Then
                If CheckIfHaveCFentry(REFF) = True Then
                    If ChkEntry2 = False Then
                        If MsgBox("CASH FLOW ENTRY is UNBALANCE or EMPTY.." & vbNewLine & "DO you want to ignore and DELETE the Cash Flow Entry?" & vbNewLine & "YES - to Ignore and DELETE" & vbNewLine & "NO to Edit the Cash Flow Entry", vbCritical + vbYesNo, "System Confirmation") = vbYes Then
                            opndbaseFMIS.Execute "Delete from dbo.tblAMIS_PostedJEVforCashflow where jevno = '" & REFF & "'"
                            opndbaseFMIS.Execute "Update dbo.tblAMIS_FinalJEV set filterInCashflow = 0 where jevno = '" & REFF & "' and actioncode = 1"
                        Else
                        Cancel = 1
                        Exit Sub
                        End If
                    Else
                        opndbaseFMIS.Execute "Update dbo.tblAMIS_FinalJEV set filterInCashflow = 1 where jevno = '" & REFF & "' and actioncode = 1"
                    End If
                Else
                    opndbaseFMIS.Execute "Update dbo.tblAMIS_FinalJEV set filterInCashflow = 1 where jevno = '" & REFF & "' and actioncode = 1"
                End If
            End If
            frm.IsSaveAccntng = True
            
        Else
            If CheckIfADMIN(ActiveUserID) = True Then
                If CheckIfHaveCFentry(REFF) = True Then
                    If ChkEntry2 = True Then
                        opndbaseFMIS.Execute "Update fmis.dbo.tblAMIS_FinalJEV set filterInCashflow = 1 where actioncode = 1 and jevno = '" & REFF & "'"
                    Else
                        MsgBox "CASH FLOW ENTRY is UNBALANCE or EMPTY.., Please check your Entry..!", vbInformation, "System Confirmation"
                        opndbaseFMIS.Execute "Update fmis.dbo.tblAMIS_FinalJEV set filterInCashflow = 0 where actioncode = 1 and jevno = '" & REFF & "'"
                        Cancel = 1
                        Exit Sub
                    End If
                End If
            Else
                opndbaseFMIS.Execute "Update dbo.tblAMIS_FinalJEV set filterInCashflow = 1 where jevno = '" & REFF & "' and actioncode = 1"
            End If
            frm.IsSaveAccntng = True
        End If
    Else
        If chksave.Value = 1 Then
            frm.IsSaveAccntng = True
        Else
'            If ChkEntry2 = False Then
'                MsgBox "cash flow is unbalanced"
'            End If
            If MsgBox("Debit and Credit Amounts must be equal." & vbNewLine & "Yes = Edit AccountingEntries" & vbNewLine & "No = Cancel Editing", vbCritical + vbYesNo, "System Confirmation") = vbYes Then
                Cancel = 1
            Else
                If MsgBox("Are you sure do you want to Cancel Editing Accounting Entries" & vbNewLine & "Your Current Changes will be erase if you cancel Editing.", vbCritical + vbYesNo, "System Message") = vbYes Then
                    opndbaseFMIS.Execute "Delete from tblAMIS_tmpJournal where dvno = '" & REFF & "'"
                    frm.EditCount = False
                    frm.IsSaveAccntng = False
                Else
                Cancel = 1
                End If
            End If
        End If
    End If
End If
End Sub

Private Sub lvButtons_H1_Click()
Unload Me
End Sub
Private Function SaveEntry()
Dim x As Integer
'opndbaseFMIS.Execute "Update tblAMIS_AccoutingEntries set actioncode = 2,datetimeentered = '" & Trim(datetimeentered) & "," & Format(Now, "yyyy/mm/dd hh:mm:ss AMPM") & "',userid = '" & Trim(UserID) & "," & Trim(ActiveUserID) & "' where reffno = '" & reff & "'"
'     For x = 1 To MSFlexGrid1.Rows - 1
'        If MSFlexGrid1.TextMatrix(x, 2) <> "TOTAL" Then
'            If MSFlexGrid1.TextMatrix(x, 1) <> "" Then
'                If MSFlexGrid1.TextMatrix(x, 3) <> "" Or MSFlexGrid1.TextMatrix(x, 4) <> "" Then
'                    opndbaseFMIS.Execute "Insert Into [tblAMIS_AccoutingEntries] (reffNo,debit,credit,ChildAccountcode,actioncode,datetimeentered,userid,transtype) values ('" & reff & "'," & CDbl(IIf(IsNumeric(MSFlexGrid1.TextMatrix(x, 3)) = True, MSFlexGrid1.TextMatrix(x, 3), 0)) & "," & CDbl(IIf(IsNumeric(MSFlexGrid1.TextMatrix(x, 4)) = True, MSFlexGrid1.TextMatrix(x, 4), 0)) & "," & _
'                    "'" & MSFlexGrid1.TextMatrix(x, 1) & "',1,'" & Format(Now, "yyyy/mm/dd hh:mm:ss AMPM") & "','" & ActiveUserID & "'," & Transtype & ")"
'                End If
'            End If
'        Else
'            Exit For
'        End If
'    Next x
End Function
Private Function ChkEntry() As Boolean
    ChkEntry = False
        If Damount = Camount And Camount > 0 Then
                ChkEntry = True
        End If
End Function
Private Function ChkEntry2() As Boolean
    ChkEntry2 = False
    If Damount2 = Camount2 And Camount2 > 0 Then
            ChkEntry2 = True
    End If
End Function

Private Sub lvButtons_H3_Click()
Picture2.Visible = False
cmbEntry.Visible = False
End Sub
Private Sub lvButtons_H4_Click()
Picture4.Visible = False
cmbEntry2.Visible = False
End Sub
Private Sub lvButtons_H5_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
PopupMenu popup
End Sub

Private Sub MSFlexGrid1_Click()
On Error GoTo bad
If bypass = True Then
bypass = False
Exit Sub
End If
    Select Case MSFlexGrid1.col
    Case 1 'AccntCode
        txt_entry.Visible = False
        cmbEntry.Move MSFlexGrid1.CellLeft, MSFlexGrid1.CellTop, MSFlexGrid1.CellWidth
        Picture2.Move MSFlexGrid1.CellLeft + cmbEntry.Width
        cmbEntry.Visible = True
        Picture2.Visible = True
        If Len(Trim(MSFlexGrid1.Text)) <> 0 Then
            cmbEntry.Text = MSFlexGrid1.Text
            cmbEntry.SetFocus
        Else
            cmbEntry.Text = ""
            Call GetAccountNamebyorder("Accountname")
            cmbEntry.SetFocus
        End If
        txtdetails.Text = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2)
    Case 3 To 5 'Debit/Credit
        cmbEntry.Visible = False
        Picture2.Visible = False
        If ExecFunction("SELECT [fmis].Accounting.[ufn_CheckIfHaveSub]  ('" & Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 0)) & "',0)") > 1 Then 'go to the subsidiary if have sub account
            txt_entry.Visible = False
            If Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 0)) <> "" Then
            With frm_coa_sub
                .Address = Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2))
                .col = 2
                .Gamount = Gamount
                .isPOSTED = isPOSTED
                .accountcode = Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 1))
                .Subcode1 = Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 1))
                .Subdesc1 = Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2))
                .Accntname = Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2))
                '.Condition = "Subcode1 ='" & Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 1)) & "' and  subcode2 is not null and (dvno = '" & reff & "' or dvno is null)"
                .Condition = "Exec [MPproc_LoadJEVfromtmp] @dvno = '" & Trim(REFF) & "',@Accountcode = '" & Trim(Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 1))) & "',@lvl =2"
                .Fullcode = Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 1))
                 .dvno = Trim(REFF)
                 Set .frm = Me
            .Show 1
            
            Call GetAccntngEntries(REFF)
            End With
        End If
        Else 'give the value of the account
            txt_entry.Move MSFlexGrid1.CellLeft, MSFlexGrid1.CellTop, MSFlexGrid1.CellWidth, MSFlexGrid1.CellHeight
            txt_entry.Visible = True
            If Len(Trim(MSFlexGrid1.Text)) <> 0 Then
                txt_entry.Text = MSFlexGrid1.Text
                txt_entry.SelStart = 0
                txt_entry.SelLength = Len(txt_entry.Text)
            Else
                txt_entry.Text = Format(lblamount.Caption, "#,##0.00")
            End If
            txt_entry.SetFocus
        End If
    Case Else
        txt_entry.Visible = False
        cmbEntry.Visible = False
        Picture2.Visible = False
    End Select
Exit Sub
bad:
MsgBox err.description
End Sub
Private Sub MSFlexGrid2_Click()
On Error GoTo bad
If CheckIfADMIN(ActiveUserID) = False Then
    MsgBox "You Are not Allowed to entry the Cashflow Entry.", vbCritical, "System Message"
    Exit Sub
End If
If bypass = True Then
bypass = False
Exit Sub
End If
    Select Case MSFlexGrid2.col
    Case 1 'AccntCode
        txt_Entry2.Visible = False
        cmbEntry2.Move MSFlexGrid2.CellLeft, MSFlexGrid2.CellTop, MSFlexGrid2.CellWidth
        Picture4.Move MSFlexGrid2.CellLeft + cmbEntry2.Width
        cmbEntry2.Visible = True
        Picture4.Visible = True
        If Len(Trim(MSFlexGrid2.Text)) <> 0 Then
            cmbEntry2.Text = MSFlexGrid2.Text
            cmbEntry2.SetFocus
        Else
            cmbEntry2.Text = ""
            Call GetAccountNamebyorder2("Accountname")
            cmbEntry2.SetFocus
        End If
        txtdetails2.Text = MSFlexGrid2.TextMatrix(MSFlexGrid2.Row, 2)
    Case 3 To 5 'Debit/Credit
        cmbEntry2.Visible = False
        Picture4.Visible = False
         'give the value of the account
            txt_Entry2.Move MSFlexGrid2.CellLeft, MSFlexGrid2.CellTop, MSFlexGrid2.CellWidth, MSFlexGrid2.CellHeight
            txt_Entry2.Visible = True
            If Len(Trim(MSFlexGrid2.Text)) <> 0 Then
                txt_Entry2.Text = MSFlexGrid2.Text
                txt_Entry2.SelStart = 0
                txt_Entry2.SelLength = Len(txt_Entry2.Text)
            Else
                txt_Entry2.Text = Format(lblamount.Caption, "#,##0.00")
            End If
            txt_Entry2.SetFocus
    Case Else
        txt_Entry2.Visible = False
        cmbEntry2.Visible = False
        Picture4.Visible = False
    End Select
Exit Sub
bad:
MsgBox err.description
End Sub
Private Function LoadAccountsByFund(ByVal accountcode As String)
Dim ARec As New ADODB.Recordset
Dim x As Integer

    cmbEntry.Clear
    cmbEntry.Visible = False
    FundName = GetFundName(fundmedium)
    ARec.Open ("exec Proc_CodeExplaination @accountcode = '" & accountcode & "'"), opndbaseFMIS, adOpenStatic, adLockOptimistic
    If ARec.RecordCount > 0 Then
        For x = 1 To ARec.RecordCount
            cmbEntry.AddItem ARec![childaccountcode]
            cmbEntry.ItemData(cmbEntry.NewIndex) = ARec!Subcode1
            ARec.MoveNext
        Next x
    End If
    ARec.Close
    Set ARec = Nothing
End Function
Public Function LoadAccountsbySub(ByVal ParentID As Long)
Dim ARec As New ADODB.Recordset
Dim x As Integer
Dim xx As Variant
Dim str() As String
Dim lvl As Integer
Dim Code As Long
Dim childcode As String
Dim z
   
    
    ARec.Open ("Exec Accounting.[usp_jev_motherCoa] @accountname = '" & Trim(txtfind.Text) & "' , @parentID = " & ParentID & ""), opndbaseFMIS, adOpenStatic, adLockOptimistic
        MSHFlexGrid1.Clear
        MSHFlexGrid1.Cols = 3
        MSHFlexGrid1.Rows = 2
        If ARec.RecordCount > 0 Then
            Set MSHFlexGrid1.DataSource = ARec
        End If
        MSHFlexGrid1.TextMatrix(0, 1) = "Code"
        MSHFlexGrid1.TextMatrix(0, 2) = "Explanation"
        MSHFlexGrid1.ColWidth(0) = 0
        MSHFlexGrid1.ColWidth(1) = 900
        MSHFlexGrid1.ColWidth(2) = 6000
    ARec.Close
    Set ARec = Nothing
End Function
Private Function LoadAccountsByName(ByVal accountcode As String, ByVal Condition As String)
Dim ARec As New ADODB.Recordset
Dim x As Integer
    ARec.Open "exec Accounting.[usp_getNamebychildCode]  @childaccountcode = '" & accountcode & "', @Condition = '" & Condition & "'", opndbaseFMIS, adOpenStatic
        If ARec.RecordCount > 0 Then
            LoadAccountsByName = ARec!Accountfullname
        inRec = True
        End If
    ARec.Close
    Set ARec = Nothing
End Function

Private Sub MSFlexGrid1_DblClick()
If MSHFlexGrid1.col = 2 Then
   ' If ExecFunction("SELECT [fmis].[dbo].[Mpfunc_ChckIfHaveSub] ('" & Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 1)) & "',0)") > 1 Then
            txt_entry.Visible = False
        If Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 1)) <> "" Then
            With frm_SubAccountcode
            .Address = Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2))
            .col = 2
            .isPOSTED = isPOSTED
            .accountcode = Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 1))
            .Subcode1 = Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 1))
            .Subdesc1 = Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2))
            '.Condition = "Subcode1 ='" & Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 1)) & "' and  subcode2 is not null and (dvno = '" & reff & "' or dvno is null)"
            .Condition = "Exec [MPproc_LoadJEVfromtmp] @dvno = '" & Trim(REFF) & "',@Accountcode = '" & Trim(Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 1))) & "',@lvl =2"
            .Fullcode = Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 1))
            
             .dvno = Trim(REFF)
            .Show 1
            Call GetAccntngEntries(REFF)
            End With
        End If
    'Else
    'MsgBox "End of The Accounts", vbInformation, "System Message"
    ' End If
End If
End Sub

Private Sub MSFlexGrid1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call lvButtons_H1_Click
End If
End Sub

Private Sub MSFlexGrid1_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo bad
If KeyCode = vbKeyDelete Then
    Dim i, x As Integer
   With MSFlexGrid1
    If .Row <> .RowSel Then
        If MsgBox("Are you sure do you want do remove Selected Accounts and its Contain?", vbCritical + vbYesNo, "System Message") = vbYes Then
            If .RowSel < .Row Then
                x = 0
                For i = .RowSel To .Row
                    x = x + 1
                Next i
                
                For i = 1 To x
                    If .TextMatrix(.RowSel, 2) <> "TOTAL" Then
                    opndbaseFMIS.Execute "Delete from tblAMIS_tmpjournal where accountcode like '" & Trim(.TextMatrix(.Row, 1)) & "%' and dvno = '" & REFF & "'"
                    .RemoveItem (.RowSel)
                    End If
                Next i
            Else
                x = 0
                For i = .Row To .RowSel
                    x = x + 1
                Next i
                
                For i = 1 To x
                    If .TextMatrix(.Row, 2) <> "TOTAL" Then
                    opndbaseFMIS.Execute "Delete from tblAMIS_tmpjournal where accountcode like '" & Trim(.TextMatrix(.Row, 1)) & "%' and dvno = '" & REFF & "'"
                    .RemoveItem (.Row)
                    End If
                Next i
            End If
        End If
    Else
    If .TextMatrix(.Row, 2) <> "TOTAL" Then
        If MsgBox("Are you sure do you want do remove this Account and its Contain?", vbCritical + vbYesNo, "System Message") = vbYes Then
            opndbaseFMIS.Execute "Delete from tblAMIS_tmpjournal where accountcode like '" & Trim(.TextMatrix(.Row, 1)) & "%' and dvno = '" & REFF & "'"
            .RemoveItem (.Row)
        End If
    End If
    End If
    Call GetSum
End With
Exit Sub
bad:
If err.Number = 30015 Then

Else
MsgBox err.description
End If
End If
End Sub

Private Sub MSFlexGrid1_Scroll()
txt_entry.Visible = False
End Sub

Private Sub MSHFlexGrid1_DblClick()
 If MSHFlexGrid1.Rows > 1 Then
    If Trim(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 1)) <> "" Then
            cmbEntry.Text = Trim(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 1))
             inRec = False
            Accountname = Trim(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 2))
            If Trim(cmbEntry.Text) <> "" Then
                ifCMB = True
                If Chckentry = False Then
                Exit Sub
                End If
                ifCMB = False
            End If
            MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 1) = cmbEntry.Text
            
            If MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2) = "TOTAL" Then
                    MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 3) = ""
                    MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 4) = ""
                    
            Else
                If MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2) = "" Then
                MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1
                End If
                MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2) = Accountname
            End If
            
            
        
            If MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2) <> "TOTAL" Then
                MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2) = Accountname
                    If MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2) = "" Then
                    MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1
                    End If
            Else
            MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1
            MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2) = Accountname
            End If
        cmbEntry.Visible = False
        ListView2.Visible = False
        Picture2.Visible = False
        Call GetSum
        MSFlexGrid1.SetFocus
        isEdit = True
        bypass = True
    End If
End If
End Sub

Private Sub MSHFlexGrid1_KeyPress(KeyAscii As Integer)
'Dim accountcode As String
'If KeyAscii = 13 Then
'    ifCMB = False
'    If Chckentry() = True Then
'         If Trim(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 1)) <> "" Then
'            If Right(Trim(cmbEntry.Text), 1) = "-" Then
'                MSFlexGrid1.TextMatrix((MSFlexGrid1.Rows - 1), 1) = cmbEntry.Text & Trim(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 1))
'                MSFlexGrid1.TextMatrix((MSFlexGrid1.Rows - 1), 2) = LoadAccountsByName(cmbEntry.Text & Trim(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 1)), "Summary")
'            Else
'               If Len(Trim(cmbEntry.Text)) < 3 Then
'                MSFlexGrid1.TextMatrix((MSFlexGrid1.Rows - 1), 1) = Trim(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 1))
'                MSFlexGrid1.TextMatrix((MSFlexGrid1.Rows - 1), 2) = Trim(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 2))
'                Else
'                MSFlexGrid1.TextMatrix((MSFlexGrid1.Rows - 1), 1) = cmbEntry.Text & "-" & Trim(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 1))
'                MSFlexGrid1.TextMatrix((MSFlexGrid1.Rows - 1), 2) = LoadAccountsByName(cmbEntry.Text & "-" & Trim(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 1)), "Summary")
'                End If
'            End If
'            MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1
'            cmbEntry.Move MSFlexGrid1.CellLeft, ((MSFlexGrid1.Rows - 1) * MSFlexGrid1.CellHeight), MSFlexGrid1.CellWidth
'            Call GetSum
'        End If
'    End If
'End If
End Sub
Public Function Chckentry() As Boolean
Dim x As Integer
Dim accountcode As String
Chckentry = True

    If cmbEntry.Text <> MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 1) Then
            If ifCMB = True Then
                accountcode = Trim(cmbEntry.Text)
            Else
                If Right(Trim(cmbEntry.Text), 1) = "-" Then
                    accountcode = cmbEntry.Text & Trim(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 1))
                 
                Else
                   If Len(Trim(cmbEntry.Text)) < 3 Then
                        accountcode = Trim(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 1))
                
                   Else
                       accountcode = cmbEntry.Text & "-" & Trim(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 1))
                   End If
                End If
            End If
        For x = 1 To MSFlexGrid1.Rows - 1
            If Trim(MSFlexGrid1.TextMatrix(x, 1)) = accountcode Then
               ' MsgBox "Acocuntcode and Explaination Already on the List..!", vbInformation, "System Message"
                'Chckentry = False
                'Exit For
            End If
        Next x
    End If
End Function

Private Sub Payroll_Click()
'On Error GoTo bad
Dim name As String

If Right(Trim(cmbEntry.Text), 1) = "-" Then
   name = LoadAccountsByName(Left(Trim(cmbEntry.Text), Len(Trim(cmbEntry.Text)) - 1), "Fullname")
Else
   name = LoadAccountsByName(Trim(cmbEntry.Text), "Fullname")
End If
    
'If Trim(name) = "" Then
'    MsgBox "Invalid Accountcode Please Select Another Account..!", vbCritical, "System Information"
'    Exit Sub
'End If
With frmSub2
.accountcode = cmbEntry.Text
.Accountname = LoadAccountsByName(cmbEntry.Text, "Fullname")
.REFF = Trim(REFF)
centerme frmSub2
.Show 1
GetAccntngEntries (REFF)
End With
Exit Sub
bad:
    MsgBox "Note: " & err.description, vbInformation, "System Message"
End Sub

Private Sub PC_Click()
frm_PostclosingSIEentry.REFF = REFF
frm_PostclosingSIEentry.Show 1
GetAccntngEntries (REFF)
End Sub

Private Sub Property_Click()
frmProperty.REFF = REFF
frmProperty.Show 1
GetAccntngEntries (REFF)
End Sub

Private Sub txt_entry_Change()
txtformula.Text = txt_entry.Text
End Sub

Private Sub txt_entry_KeyPress(KeyAscii As Integer)
Dim tamount As Currency
 On Error GoTo bad
    If KeyAscii = 13 Then
            If isPOSTED = True Then
                MsgBox "Unable to Edit the Entry, the Transaction is Already generate the report", vbInformation, "System Information"
                txt_entry.Visible = False
                Exit Sub
            End If
            If IsNumeric(txt_entry.Text) = False And txt_entry.Text <> "" Then
                If InStr(1, txt_entry.Text, "+") = 0 Then
                    MsgBox "None Numeric Entry, Please Check Your Entry", vbCritical, "System Message"
                    Exit Sub
                End If
            End If
            If txt_entry.Text <> "" Then
                tamount = sumAmount(txt_entry.Text)
            End If
                MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 5) = txt_entry.Text
                If txt_entry.Text <> "" Then
                MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, MSFlexGrid1.col) = IIf((tamount = 0), "", Format((tamount), "#,##0.00"))
                Else
                 MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, MSFlexGrid1.col) = ""
                End If
                txt_entry.Visible = False
            Call SaveAmount(IIf((MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 3)) = "", 0, MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 3)), IIf((MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 4)) = "", 0, MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 4)))
        Call GetSum
        isEdit = True
    End If
Exit Sub
bad:
   ' Call LoadErr(err.Number, err.Source & ", " & Me.name & ", " & Me.Caption, err.Description)
End Sub
Private Sub txt_entry2_KeyPress(KeyAscii As Integer)
Dim tamount As Currency
 On Error GoTo bad
    If KeyAscii = 13 Then
            If isPOSTED = True Then
                MsgBox "Unable to edit the entry, the transaction is already approved..", vbInformation, "System Information"
                txt_Entry2.Visible = False
                Exit Sub
            End If
            If IsNumeric(txt_Entry2.Text) = False And txt_Entry2.Text <> "" Then
                If InStr(1, txt_Entry2.Text, "+") = 0 Then
                    MsgBox "None Numeric Entry, Please Check Your Entry", vbCritical, "System Message"
                    Exit Sub
                End If
            End If
            If txt_Entry2.Text <> "" Then
                tamount = sumAmount(txt_Entry2.Text)
            End If
                MSFlexGrid2.TextMatrix(MSFlexGrid2.Row, 5) = txt_Entry2.Text
                If txt_Entry2.Text <> "" Then
                MSFlexGrid2.TextMatrix(MSFlexGrid2.Row, MSFlexGrid2.col) = IIf((tamount = 0), "", Format((tamount), "#,##0.00"))
                Else
                 MSFlexGrid2.TextMatrix(MSFlexGrid2.Row, MSFlexGrid2.col) = ""
                End If
                txt_Entry2.Visible = False
            Call SaveAmount2(IIf((MSFlexGrid2.TextMatrix(MSFlexGrid2.Row, 3)) = "", 0, MSFlexGrid2.TextMatrix(MSFlexGrid2.Row, 3)), IIf((MSFlexGrid2.TextMatrix(MSFlexGrid2.Row, 4)) = "", 0, MSFlexGrid2.TextMatrix(MSFlexGrid2.Row, 4)))
        Call GetSum2
        isEdit = True
    End If
Exit Sub
bad:
   ' Call LoadErr(err.Number, err.Source & ", " & Me.name & ", " & Me.Caption, err.Description)
End Sub
Public Function SaveAmount(ByVal Debit As Currency, ByVal Credit As Currency)
Dim rec As New ADODB.Recordset

If Debit = 0 And Credit = 0 Then
opndbaseFMIS.Execute "Delete from tblAMIS_tmpJournal where accountcode = '" & Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 1)) & "' and dvno = '" & REFF & "' "
Else
    Set rec = opndbaseFMIS.Execute("Select Accountcode from tblAMIS_tmpjournal where accountcode = '" & Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 1)) & "' and dvno = '" & REFF & "'")
        If rec.RecordCount > 0 Then
            opndbaseFMIS.Execute "Update tblAMIS_tmpJournal set debit = '" & Debit & "',credit = '" & Credit & "' where accountcode = '" & Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 1)) & "' and dvno = '" & REFF & "' "
        Else
            opndbaseFMIS.Execute "Insert into tblAMIS_tmpJournal (accountcode,debit,Credit,dvno) values ('" & Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 1)) & "','" & Debit & "','" & Credit & "','" & Trim(REFF) & "')"
        End If
        GetSum
    rec.Close
End If
End Function
Public Function SaveAmount2(ByVal Debit As Currency, ByVal Credit As Currency)
Dim rec As New ADODB.Recordset

If Debit = 0 And Credit = 0 Then
opndbaseFMIS.Execute "Delete from tblAMIS_PostedJEVforCashflow where accountcode = '" & Trim(MSFlexGrid2.TextMatrix(MSFlexGrid2.Row, 1)) & "' and Jevno = '" & REFF & "' "
Else
    Set rec = opndbaseFMIS.Execute("Select Accountcode from tblAMIS_PostedJEVforCashflow where accountcode = '" & Trim(MSFlexGrid2.TextMatrix(MSFlexGrid2.Row, 1)) & "' and jevno = '" & REFF & "'")
        If Trim(MSFlexGrid2.TextMatrix(MSFlexGrid2.Row, 1)) <> "" Then
                If rec.RecordCount > 0 Then
                    opndbaseFMIS.Execute "Update tblAMIS_PostedJEVforCashflow set debit = '" & Debit & "',credit = '" & Credit & "' where accountcode = '" & Trim(MSFlexGrid2.TextMatrix(MSFlexGrid2.Row, 1)) & "' and jevno = '" & REFF & "' "
                Else
                    opndbaseFMIS.Execute "Insert into tblAMIS_PostedJEVforCashflow (accountcode,debit,Credit,jevno) values ('" & Trim(MSFlexGrid2.TextMatrix(MSFlexGrid2.Row, 1)) & "','" & Debit & "','" & Credit & "','" & Trim(REFF) & "')"
                End If
        End If
        GetSum2
    rec.Close
End If
End Function

Public Sub GetSum()
Dim x As Integer
    Damount = 0
    Camount = 0
    For x = 1 To MSFlexGrid1.Rows - 1
        If MSFlexGrid1.TextMatrix(x, 1) <> "" Then
            Damount = Damount + CCur(IIf(MSFlexGrid1.TextMatrix(x, 3) = "", 0, MSFlexGrid1.TextMatrix(x, 3)))
            Camount = Camount + CCur(IIf(MSFlexGrid1.TextMatrix(x, 4) = "", 0, MSFlexGrid1.TextMatrix(x, 4)))
        Else

           ' MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1
            MSFlexGrid1.TextMatrix(x, 2) = "TOTAL"
            MSFlexGrid1.TextMatrix(x, 3) = IIf((Damount = 0), "", Format(Damount, "#,##0.00"))
            MSFlexGrid1.TextMatrix(x, 4) = IIf((Camount = 0), "", Format(Camount, "#,##0.00"))
            Exit For
        End If
    Next x
End Sub
Public Sub GetSum2()
Dim x As Integer
    Damount2 = 0
    Camount2 = 0
    For x = 1 To MSFlexGrid2.Rows - 1
        If MSFlexGrid2.TextMatrix(x, 1) <> "" Then
            Damount2 = Damount2 + CCur(IIf(MSFlexGrid2.TextMatrix(x, 3) = "", 0, MSFlexGrid2.TextMatrix(x, 3)))
            Camount2 = Camount2 + CCur(IIf(MSFlexGrid2.TextMatrix(x, 4) = "", 0, MSFlexGrid2.TextMatrix(x, 4)))
        Else

           ' MSFlexGrid2.Rows = MSFlexGrid2.Rows + 1
            MSFlexGrid2.TextMatrix(x, 2) = "TOTAL"
            MSFlexGrid2.TextMatrix(x, 3) = IIf((Damount2 = 0), "", Format(Damount2, "#,##0.00"))
            MSFlexGrid2.TextMatrix(x, 4) = IIf((Camount2 = 0), "", Format(Camount2, "#,##0.00"))
            Exit For
        End If
    Next x
End Sub
Private Function whatAction()

End Function

Private Sub txtfind_Change()
    If Len(Trim(cmbEntry.Text)) >= 3 Then
        LoadAccountsbySub (cmbEntry.Text)
        txtdetails.Text = LoadAccountsByName(cmbEntry.Text, "Fullname")
    Else
        Call GetAccountNamebyorder("Accountcode")
    End If
End Sub
Private Sub txtfind2_Change()
    If Len(Trim(cmbEntry2.Text)) >= 3 Then
        LoadAccountsbySub (cmbEntry2.Text)
        txtdetails2.Text = LoadAccountsByName(cmbEntry2.Text, "Fullname")
    Else
    Call GetAccountNamebyorder2("Accountcode")
    End If
End Sub

Private Sub txtfind_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    ifCMB = False
    If Chckentry() = True Then
         If Trim(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 1)) <> "" Then
            If Right(Trim(cmbEntry.Text), 1) = "-" Then
                MSFlexGrid1.TextMatrix((MSFlexGrid1.Rows - 1), 1) = cmbEntry.Text & Trim(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 1))
                MSFlexGrid1.TextMatrix((MSFlexGrid1.Rows - 1), 2) = LoadAccountsByName(cmbEntry.Text & Trim(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 1)), "Summary")
            Else
               If Len(Trim(cmbEntry.Text)) < 3 Then
                MSFlexGrid1.TextMatrix((MSFlexGrid1.Rows - 1), 1) = Trim(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 1))
                MSFlexGrid1.TextMatrix((MSFlexGrid1.Rows - 1), 2) = Trim(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 2))
                Else
                MSFlexGrid1.TextMatrix((MSFlexGrid1.Rows - 1), 1) = cmbEntry.Text & "-" & Trim(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 1))
                MSFlexGrid1.TextMatrix((MSFlexGrid1.Rows - 1), 2) = LoadAccountsByName(cmbEntry.Text & "-" & Trim(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 1)), "Summary")
                End If
            End If
            MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1
            cmbEntry.Move MSFlexGrid1.CellLeft, ((MSFlexGrid1.Rows - 1) * MSFlexGrid1.CellHeight), MSFlexGrid1.CellWidth
            Call GetSum
        End If
    End If
End If
End Sub
Private Sub txtfind2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    ifCMB = False
    If Chckentry() = True Then
         If Trim(MSFlexGrid2.TextMatrix(MSFlexGrid2.Row, 1)) <> "" Then
            If Right(Trim(cmbEntry.Text), 1) = "-" Then
                MSFlexGrid1.TextMatrix((MSFlexGrid1.Rows - 1), 1) = cmbEntry.Text & Trim(MSFlexGrid2.TextMatrix(MSFlexGrid2.Row, 1))
                MSFlexGrid1.TextMatrix((MSFlexGrid1.Rows - 1), 2) = LoadAccountsByName(cmbEntry.Text & Trim(MSFlexGrid2.TextMatrix(MSFlexGrid2.Row, 1)), "Summary")
            Else
               If Len(Trim(cmbEntry.Text)) < 3 Then
                MSFlexGrid1.TextMatrix((MSFlexGrid1.Rows - 1), 1) = Trim(MSFlexGrid2.TextMatrix(MSFlexGrid2.Row, 1))
                MSFlexGrid1.TextMatrix((MSFlexGrid1.Rows - 1), 2) = Trim(MSFlexGrid2.TextMatrix(MSFlexGrid2.Row, 2))
                Else
                MSFlexGrid1.TextMatrix((MSFlexGrid1.Rows - 1), 1) = cmbEntry.Text & "-" & Trim(MSFlexGrid2.TextMatrix(MSFlexGrid2.Row, 1))
                MSFlexGrid1.TextMatrix((MSFlexGrid1.Rows - 1), 2) = LoadAccountsByName(cmbEntry.Text & "-" & Trim(MSFlexGrid2.TextMatrix(MSFlexGrid2.Row, 1)), "Summary")
                End If
            End If
            MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1
            cmbEntry.Move MSFlexGrid1.CellLeft, ((MSFlexGrid1.Rows - 1) * MSFlexGrid1.CellHeight), MSFlexGrid1.CellWidth
            Call GetSum2
        End If
    End If
End If
End Sub

Private Sub txtformula_Change()
txt_entry.Text = txtformula.Text
End Sub

Private Sub txtformula_KeyPress(KeyAscii As Integer)
Dim tamount As Currency
 On Error GoTo bad
    If KeyAscii = 13 Then
            If IsNumeric(txt_entry.Text) = False And txt_entry.Text <> "" Then
                 If InStr(1, txt_entry.Text, "+") = 0 Then
                    MsgBox "None Numeric Entry, Please Check Your Entry", vbCritical, "System Message"
                    Exit Sub
                End If
            End If
                tamount = sumAmount(txt_entry.Text)
                MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 5) = txt_entry.Text
                MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, MSFlexGrid1.col) = Format((tamount), "#,##0.00")
                txt_entry.Visible = False
            
            If MSFlexGrid1.col = 3 Then
                If MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 3) <> "" Then
                    MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 4) = ""
                End If
            ElseIf MSFlexGrid1.col = 4 Then
                If MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 4) <> "" Then
                    MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 3) = ""
                End If
            End If
        Call GetSum
        isEdit = True
    End If
Exit Sub
bad:
    Call LoadErr(err.Number, err.Source & ", " & Me.name & ", " & Me.Caption, err.description)
End Sub
Public Function GetAccntngEntries(ByVal dvno As String)
On Error Resume Next
Dim Drec As New ADODB.Recordset
Dim x As Integer
    Call SetGrid
    Set Drec = opndbaseFMIS.Execute("Select left(Accountcode,3) as childcode,sum(Debit) as sumdebit,sum(credit) as sumcredit From tblAMIS_tmpjournal Where [dvno]='" & dvno & "' group by Dvno,left(Accountcode,3) order by sum(debit) desc")
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
        Call GetSum
    End If
    Drec.Close
    Set Drec = Nothing
End Function

Public Function GetAccntngEntries2(ByVal dvno As String)
On Error Resume Next
Dim Drec As New ADODB.Recordset
Dim x As Integer
    Call SetGrid2
    Set Drec = opndbaseFMIS.Execute("Select left(Accountcode,3) as childcode,sum(Debit) as sumdebit,sum(credit) as sumcredit From tblAMIS_PostedJEVforCashflow Where jevno='" & dvno & "' group by jevno,left(Accountcode,3) order by sum(debit) desc")
    If Drec.RecordCount > 0 Then
        For x = 1 To Drec.RecordCount
            'MSFlexGrid1.TextMatrix(x, 0) = DRec![Trnno]
            
            MSFlexGrid2.TextMatrix(x, 1) = Drec!childcode
            MSFlexGrid2.TextMatrix(x, 2) = GetAccountNameByAccountcode(Drec!childcode)
            MSFlexGrid2.TextMatrix(x, 4) = IIf((Format(Drec!sumCredit, "#,##0.00") = "0.00"), "", Format(Drec!sumCredit, "#,##0.00"))
            MSFlexGrid2.TextMatrix(x, 3) = IIf((Format(Drec!sumDebit, "#,##0.00") = "0.00"), "", Format(Drec!sumDebit, "#,##0.00"))
            MSFlexGrid2.Rows = MSFlexGrid2.Rows + 1
            'If LCase(Trim(lblMode.Caption)) = "edit" Then MSFlexGrid2.TextMatrix(x, 5) = DRec!ActionCode  ' for coloraly purpose
            Drec.MoveNext
        Next x
        Call GetSum2
    End If
    Drec.Close
    Set Drec = Nothing
End Function
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

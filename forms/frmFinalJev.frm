VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MShflxgd.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmFinalJev 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Journal Entry Voucher"
   ClientHeight    =   9735
   ClientLeft      =   -2115
   ClientTop       =   225
   ClientWidth     =   15180
   Icon            =   "frmFinalJev.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   2  'Custom
   Picture         =   "frmFinalJev.frx":076A
   ScaleHeight     =   9735
   ScaleWidth      =   15180
   Begin VB.Frame Frame1 
      Caption         =   "JEV Details"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4050
      Left            =   3480
      TabIndex        =   38
      Top             =   840
      Width           =   11535
      Begin VB.CheckBox Check6 
         Caption         =   "Continuing"
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
         Left            =   10080
         TabIndex        =   77
         Tag             =   "1"
         Top             =   2760
         Width           =   1335
      End
      Begin VB.CommandButton Command4 
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
         Left            =   6840
         MaskColor       =   &H0000FF00&
         Picture         =   "frmFinalJev.frx":0CF4
         Style           =   1  'Graphical
         TabIndex        =   61
         ToolTipText     =   "Click to Generate JEV number"
         Top             =   350
         Width           =   375
      End
      Begin VB.TextBox txtjevno 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         TabIndex        =   59
         Top             =   405
         Width           =   2415
      End
      Begin VB.TextBox txtdvno 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         TabIndex        =   58
         Top             =   800
         Width           =   2415
      End
      Begin VB.TextBox txtobrno 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         TabIndex        =   57
         Top             =   1190
         Width           =   2415
      End
      Begin VB.TextBox txtcheckno 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         TabIndex        =   56
         Top             =   1570
         Width           =   2415
      End
      Begin VB.TextBox txtrcino 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         TabIndex        =   55
         Top             =   1970
         Width           =   2415
      End
      Begin VB.TextBox txtrdono 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         TabIndex        =   54
         Top             =   2350
         Width           =   3255
      End
      Begin VB.TextBox txtjevdate 
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "MMMM yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
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
         TabIndex        =   53
         Top             =   360
         Width           =   1815
      End
      Begin VB.TextBox txtclaimant 
         BackColor       =   &H80000016&
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
         Left            =   6720
         Locked          =   -1  'True
         TabIndex        =   52
         Top             =   1320
         Width           =   4215
      End
      Begin VB.TextBox txtparticular 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   1200
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   51
         Top             =   3150
         Width           =   8895
      End
      Begin VB.TextBox txtptvno 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         TabIndex        =   50
         Top             =   2750
         Width           =   3255
      End
      Begin VB.TextBox txtcheckdate 
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
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
         Left            =   9000
         Locked          =   -1  'True
         TabIndex        =   49
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox txtgamount 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   6720
         TabIndex        =   48
         Top             =   2745
         Width           =   3135
      End
      Begin VB.CommandButton cmd_Browse 
         Height          =   360
         Left            =   11040
         MaskColor       =   &H0000FF00&
         Picture         =   "frmFinalJev.frx":1036
         Style           =   1  'Graphical
         TabIndex        =   47
         Top             =   1320
         Width           =   375
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FF8080&
         Caption         =   "..."
         Height          =   375
         Left            =   11000
         MaskColor       =   &H0000FF00&
         Picture         =   "frmFinalJev.frx":22A8
         TabIndex        =   46
         Top             =   2280
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00FF8080&
         Caption         =   "..."
         Height          =   375
         Left            =   11000
         MaskColor       =   &H0000FF00&
         Picture         =   "frmFinalJev.frx":5DA2
         TabIndex        =   45
         Top             =   1800
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.ComboBox txtfundtype 
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         ItemData        =   "frmFinalJev.frx":989C
         Left            =   6720
         List            =   "frmFinalJev.frx":98AC
         Style           =   2  'Dropdown List
         TabIndex        =   44
         Top             =   840
         Width           =   4695
      End
      Begin VB.ComboBox txtooe 
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         ItemData        =   "frmFinalJev.frx":98CD
         Left            =   6720
         List            =   "frmFinalJev.frx":98DD
         Style           =   2  'Dropdown List
         TabIndex        =   43
         Top             =   1800
         Width           =   4695
      End
      Begin VB.ComboBox txtrc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         ItemData        =   "frmFinalJev.frx":98FE
         Left            =   6720
         List            =   "frmFinalJev.frx":990E
         Style           =   2  'Dropdown List
         TabIndex        =   42
         Top             =   2280
         Width           =   4695
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Post Closing"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   10440
         TabIndex        =   40
         Top             =   240
         Width           =   975
      End
      Begin VB.CheckBox Check5 
         Caption         =   "No Documents attached"
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
         Left            =   10080
         TabIndex        =   39
         Tag             =   "1"
         Top             =   3120
         Width           =   1335
      End
      Begin TabDlg.SSTab SSTab1 
         Height          =   300
         Left            =   0
         TabIndex        =   41
         Top             =   4080
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   529
         _Version        =   393216
         TabHeight       =   520
         TabCaption(0)   =   "Tab 0"
         TabPicture(0)   =   "frmFinalJev.frx":992F
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).ControlCount=   0
         TabCaption(1)   =   "Tab 1"
         Tab(1).ControlEnabled=   0   'False
         Tab(1).ControlCount=   0
         TabCaption(2)   =   "Tab 2"
         Tab(2).ControlEnabled=   0   'False
         Tab(2).ControlCount=   0
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   360
         Left            =   10200
         TabIndex        =   60
         Top             =   360
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   635
         _Version        =   393216
         Format          =   157614081
         CurrentDate     =   40680
      End
      Begin VB.Label Label2 
         Caption         =   "JEV No.:"
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
         TabIndex        =   76
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "DVNo.:"
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
         TabIndex        =   75
         Top             =   890
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Obr No.:"
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
         TabIndex        =   74
         Top             =   1275
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "Check No.:"
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
         TabIndex        =   73
         Top             =   1635
         Width           =   1215
      End
      Begin VB.Label Label6 
         Caption         =   "RCI No.:"
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
         TabIndex        =   72
         Top             =   2070
         Width           =   1095
      End
      Begin VB.Label Label7 
         Caption         =   "RDO No.:"
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
         TabIndex        =   71
         Top             =   2450
         Width           =   1095
      End
      Begin VB.Label Label8 
         Caption         =   "JEV Date.:"
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
         Left            =   3840
         TabIndex        =   70
         Top             =   435
         Width           =   1215
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
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
         Height          =   375
         Left            =   4800
         TabIndex        =   69
         Top             =   1440
         Width           =   1815
      End
      Begin VB.Label gh 
         Alignment       =   1  'Right Justify
         Caption         =   "Object of Expenditure:"
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
         Left            =   4080
         TabIndex        =   68
         Top             =   1920
         Width           =   2535
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
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
         Height          =   255
         Left            =   -120
         TabIndex        =   67
         Top             =   3240
         Width           =   1095
      End
      Begin VB.Label Label13 
         Caption         =   "PTV No.:"
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
         TabIndex        =   66
         Top             =   2800
         Width           =   1095
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Caption         =   "Check/ Deposit/ Transaction date:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   7200
         TabIndex        =   65
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label sd 
         Alignment       =   1  'Right Justify
         Caption         =   "Special Account:"
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
         Left            =   4800
         TabIndex        =   64
         Top             =   915
         Width           =   1815
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         Caption         =   "Responsibilty Center:"
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
         Left            =   4320
         TabIndex        =   63
         Top             =   2355
         Width           =   2295
      End
      Begin VB.Label Label11 
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
         Height          =   375
         Left            =   5160
         TabIndex        =   62
         Top             =   2760
         Width           =   1815
      End
   End
   Begin MSComctlLib.ProgressBar progStat 
      Height          =   150
      Left            =   3600
      TabIndex        =   26
      Top             =   9405
      Visible         =   0   'False
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   265
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Max             =   1000
      Scrolling       =   1
   End
   Begin TabDlg.SSTab SSTab2 
      Height          =   4695
      Left            =   3480
      TabIndex        =   27
      Top             =   4920
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   8281
      _Version        =   393216
      Style           =   1
      Tabs            =   6
      TabsPerRow      =   6
      TabHeight       =   697
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Accounting Entries"
      TabPicture(0)   =   "frmFinalJev.frx":994B
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblmsg"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "MSFlexGrid1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Cash Flow Entry"
      TabPicture(1)   =   "frmFinalJev.frx":9967
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "MSFlexGrid2"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "JEV  Entry log"
      TabPicture(2)   =   "frmFinalJev.frx":9983
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "MSHFlexGrid1"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "OR Details"
      TabPicture(3)   =   "frmFinalJev.frx":999F
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "MSHFlexGrid2"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "Consolidated JEV Entry "
      TabPicture(4)   =   "frmFinalJev.frx":99BB
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Label18"
      Tab(4).Control(1)=   "lblstat"
      Tab(4).Control(2)=   "lvButtons_H2"
      Tab(4).Control(3)=   "MSHFlexGrid3"
      Tab(4).Control(4)=   "lvButtons_H1"
      Tab(4).ControlCount=   5
      TabCaption(5)   =   "Tab 5"
      TabPicture(5)   =   "frmFinalJev.frx":99D7
      Tab(5).ControlEnabled=   0   'False
      Tab(5).ControlCount=   0
      Begin lvButton.lvButtons_H lvButtons_H1 
         Height          =   450
         Left            =   -74880
         TabIndex        =   31
         Top             =   510
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   794
         Caption         =   "Approve"
         CapAlign        =   2
         BackStyle       =   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFHover         =   16711680
         cBhover         =   8421504
         LockHover       =   3
         cGradient       =   12632256
         Gradient        =   4
         Mode            =   0
         Value           =   0   'False
         Image           =   "frmFinalJev.frx":99F3
         ImgSize         =   24
         cBack           =   -2147483633
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   3855
         Left            =   -74880
         TabIndex        =   28
         Top             =   510
         Width           =   11295
         _ExtentX        =   19923
         _ExtentY        =   6800
         _Version        =   393216
         BackColorBkg    =   16777088
         BandDisplay     =   1
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
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid2 
         Height          =   3855
         Left            =   -74880
         TabIndex        =   29
         Top             =   510
         Width           =   11295
         _ExtentX        =   19923
         _ExtentY        =   6800
         _Version        =   393216
         BackColorBkg    =   12632256
         BandDisplay     =   1
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
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid3 
         Height          =   3495
         Left            =   -74880
         TabIndex        =   30
         Tag             =   "2"
         Top             =   990
         Width           =   11295
         _ExtentX        =   19923
         _ExtentY        =   6165
         _Version        =   393216
         Cols            =   3
         BackColorBkg    =   16761024
         ScrollTrack     =   -1  'True
         AllowUserResizing=   1
         BandDisplay     =   1
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
         _Band(0).BandIndent=   5
         _Band(0).Cols   =   3
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin lvButton.lvButtons_H lvButtons_H2 
         Height          =   450
         Left            =   -73080
         TabIndex        =   32
         Top             =   510
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   794
         Caption         =   "Disapprove"
         CapAlign        =   2
         BackStyle       =   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFHover         =   16711680
         cBhover         =   8421504
         LockHover       =   3
         cGradient       =   12632256
         Gradient        =   4
         Mode            =   0
         Value           =   0   'False
         Image           =   "frmFinalJev.frx":AA45
         cBack           =   -2147483633
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   3840
         Left            =   120
         TabIndex        =   78
         Top             =   480
         Width           =   11280
         _ExtentX        =   19897
         _ExtentY        =   6773
         _Version        =   393216
         FixedCols       =   0
         ScrollTrack     =   -1  'True
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
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
         Height          =   3840
         Left            =   -74880
         TabIndex        =   79
         Top             =   480
         Width           =   11280
         _ExtentX        =   19897
         _ExtentY        =   6773
         _Version        =   393216
         FixedCols       =   0
         BackColorBkg    =   8421631
         ScrollTrack     =   -1  'True
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
      Begin VB.Label lblmsg 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
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
         Left            =   7920
         TabIndex        =   37
         Top             =   510
         Visible         =   0   'False
         Width           =   3495
      End
      Begin VB.Label lblstat 
         AutoSize        =   -1  'True
         Caption         =   "NOT APPROVED"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   -70440
         TabIndex        =   34
         Top             =   630
         Width           =   1605
      End
      Begin VB.Label Label18 
         Caption         =   "Status:"
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
         Left            =   -71160
         TabIndex        =   33
         Top             =   630
         Width           =   735
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Search"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8775
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   3255
      Begin VB.CheckBox Check4 
         Caption         =   "No Documents attached"
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
         TabIndex        =   36
         Tag             =   "1"
         Top             =   4440
         Width           =   2895
      End
      Begin VB.CheckBox Check3 
         Caption         =   "For Approval"
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
         Tag             =   "1"
         Top             =   4130
         Width           =   2535
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Search through OBR number"
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
         TabIndex        =   24
         Top             =   3855
         Width           =   3015
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Go"
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
         Left            =   2640
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   4850
         Width           =   495
      End
      Begin VB.Frame Frame4 
         Caption         =   "Transaction Type"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1695
         Left            =   120
         TabIndex        =   8
         Top             =   720
         Width           =   3015
         Begin VB.OptionButton Option6 
            Caption         =   "Memo Entry"
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
            Left            =   240
            TabIndex        =   25
            Tag             =   "5"
            Top             =   1320
            Width           =   2175
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Check Disbursement"
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
            Left            =   240
            TabIndex        =   12
            Tag             =   "2"
            Top             =   480
            Width           =   2175
         End
         Begin VB.OptionButton Option3 
            Caption         =   "Cash Disbursement"
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
            Left            =   240
            TabIndex        =   11
            Tag             =   "3"
            Top             =   780
            Width           =   2175
         End
         Begin VB.OptionButton Option4 
            Caption         =   "General Journal"
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
            Left            =   240
            TabIndex        =   10
            Tag             =   "4"
            Top             =   1080
            Width           =   2175
         End
         Begin VB.OptionButton Option5 
            Caption         =   "Cash Receipts"
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
            Left            =   240
            TabIndex        =   9
            Tag             =   "1"
            Top             =   240
            Value           =   -1  'True
            Width           =   2175
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Date Post"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1380
         Left            =   120
         TabIndex        =   6
         Top             =   2400
         Width           =   3015
         Begin VB.OptionButton Option7 
            Caption         =   "Month-Year"
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
            Left            =   240
            TabIndex        =   14
            Tag             =   "3"
            Top             =   680
            Width           =   1335
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Year"
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
            Left            =   240
            TabIndex        =   13
            Tag             =   "1"
            Top             =   360
            Value           =   -1  'True
            Width           =   855
         End
         Begin MSComCtl2.DTPicker DTPYear 
            Height          =   375
            Left            =   1080
            TabIndex        =   7
            Top             =   240
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
            CustomFormat    =   "yyyy"
            Format          =   155844611
            UpDown          =   -1  'True
            CurrentDate     =   40651
         End
         Begin MSComCtl2.DTPicker DTpMY 
            Height          =   375
            Left            =   480
            TabIndex        =   15
            Top             =   915
            Width           =   2175
            _ExtentX        =   3836
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
            CustomFormat    =   "MMMM yyyy"
            Format          =   155844611
            UpDown          =   -1  'True
            CurrentDate     =   40651
         End
      End
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3030
         Left            =   120
         TabIndex        =   5
         Top             =   5280
         Width           =   3015
      End
      Begin VB.TextBox txtcondition 
         Appearance      =   0  'Flat
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
         ToolTipText     =   "Type JEVno/Dvno/Checkno/PTVno and Press Enter"
         Top             =   4850
         Width           =   2500
      End
      Begin VB.ComboBox cmb_Field 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         ItemData        =   "frmFinalJev.frx":E54F
         Left            =   1320
         List            =   "frmFinalJev.frx":E562
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   285
         Width           =   1815
      End
      Begin VB.Label lblcount 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
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
         Left            =   -120
         TabIndex        =   17
         Top             =   8445
         Width           =   3255
      End
      Begin VB.Label Label1 
         Caption         =   "Select Field:"
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
         TabIndex        =   2
         Top             =   360
         Width           =   1215
      End
   End
   Begin MSComctlLib.ImageList itb32x32 
      Left            =   12600
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
            Picture         =   "frmFinalJev.frx":E58A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFinalJev.frx":FF1C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFinalJev.frx":118AE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFinalJev.frx":13240
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFinalJev.frx":14BD2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFinalJev.frx":16564
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFinalJev.frx":17EF6
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFinalJev.frx":19888
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFinalJev.frx":1B21A
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFinalJev.frx":1CBAE
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFinalJev.frx":1D88A
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFinalJev.frx":1E16A
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFinalJev.frx":1EE46
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFinalJev.frx":1FB22
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFinalJev.frx":207FE
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFinalJev.frx":214DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFinalJev.frx":221B6
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
      Width           =   15180
      _ExtentX        =   26776
      _ExtentY        =   1482
      ButtonWidth     =   1826
      ButtonHeight    =   1429
      Appearance      =   1
      Style           =   1
      ImageList       =   "itb32x32"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   13
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "New"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
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
            Caption         =   "Edit"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
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
            Caption         =   "Adjustment"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Print JEV"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Close"
            ImageIndex      =   7
         EndProperty
      EndProperty
      Begin VB.Frame Frame6 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   13320
         TabIndex        =   21
         Top             =   0
         Width           =   1815
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Mode:"
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
            Left            =   120
            TabIndex        =   23
            Top             =   240
            Width           =   750
         End
         Begin VB.Label lblMode 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "New"
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
            Height          =   300
            Left            =   960
            TabIndex        =   22
            Top             =   240
            Width           =   645
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Note:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   7800
         TabIndex        =   19
         Top             =   0
         Width           =   5415
         Begin VB.Label Label16 
            Caption         =   "Leave the field if Not Applicable, Proceed to the next Field"
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
            TabIndex        =   20
            Top             =   240
            Width           =   5175
         End
      End
      Begin VB.Timer Timer1 
         Interval        =   50
         Left            =   13560
         Top             =   120
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H00FFFFFF&
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   15
         Left            =   0
         Top             =   0
         Width           =   1755
      End
   End
   Begin VB.Label Label14 
      BackColor       =   &H80000012&
      Caption         =   "akot"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   0
      TabIndex        =   18
      Top             =   0
      Width           =   6615
   End
End
Attribute VB_Name = "frmFinalJev"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim isEdit, IfNew As Boolean
Dim IfEdit As String
Public ClaimantCode, RCcode, FmisVoucherno, ref As String
Dim Jevseries   As Long
Dim trans As Integer
Dim xDebit As Currency
Dim xCredit2 As Currency
Dim xDebit2 As Currency
Dim xCredit As Currency
Dim ifColoraly As Boolean

Dim ifsaveamount As Boolean
Dim SaveOk As Boolean
Public Ttype, PClosinG As Integer
Public Fundcode As Long
Public FundType As String
Public EditCount, IsSaveAccntng As Boolean
Public isPOSTED As Boolean
Dim not_coloraly_total_debit, not_coloraly_total_credit, coloraly_total_debit, coloraly_total_credit As Double
Dim Jdate As Date
Private Function txt(ByVal whatconDition As String)
Select Case (whatconDition)

Case "Disable"
            Frame1.Enabled = False
            'Picture1.Enabled = False
Case "Enable"
            Frame1.Enabled = True
            'Picture1.Enabled = True
Case "Clear"
            txtjevdate.Text = ""
            txtCheckno.Text = ""
            txtcheckdate.Text = ""
            txtClaimant.Text = ""
            'txtcondition.Text = ""
            txtDVNo.Text = ""
            txtfundtype.ListIndex = 0
            txtgamount.Text = ""
            txtJEVNo.Text = ""
            txtobrno.Text = ""
            txtOOE.ListIndex = 0
            txtParticular.Text = ""
            txtptvno.Text = ""
            txtRC.ListIndex = 0
            txtrcino.Text = ""
            txtrdono.Text = ""
            ClaimantCode = ""
            RCcode = ""
            FmisVoucherno = ""
            ref = ""
            Call SetGrid
            
End Select
End Function
Private Function Condition(ByVal whatconDition As String)
Select Case whatconDition
Case "ifNEW"
cmd_Browse.Visible = True
txt (Clear)
txt (Enable)
Frame2.Enabled = False
Case "IfEdit"
txt (Enable)
Load
End Select
End Function

Private Sub Check2_Click()
If CheckIfADMIN(ActiveUserID) = False Then
    MsgBox "You are not allowed to entry Post Closing...", vbCritical, "System Message"
    Check2.Value = 0
End If
End Sub

Private Sub cmb_Field_Change()
Call LoadFinalTrans(cmb_Field.Text, txtcondition.Text)
Check1.Caption = "Search " & cmb_Field.Text & " through Obrno"
End Sub

Private Sub cmb_Field_Click()
Call LoadFinalTrans(cmb_Field.Text, txtcondition.Text)
Check1.Caption = "Search " & cmb_Field.Text & " through Obrno"
End Sub



Private Sub Command1_Click()
Call LoadFinalTrans(cmb_Field.Text, txtcondition.Text)
End Sub

Private Sub cmd_Browse_Click()
CUFlag = True
ActiveFormCaller = "frmFinalJEV"
frmCDClaimantRegistry.Show 1

End Sub

Private Sub Command2_Click()
With Form3
.Move (Me.Left + txtRC.Left) + 50, txtRC.Top + Me.Top + txtRC.Height + 400, txtRC.Width
.Show 1
End With
End Sub
Private Sub Command3_Click()
With Form3
.Move (Me.Left + txtOOE.Left) + Frame1.Left + 50, txtOOE.Top + Frame1.Top + Me.Top + txtOOE.Height + 400, txtOOE.Width
.ListView1.Width = .Width - 30
.Show 1
End With
End Sub
Private Sub loadField()
On Error GoTo bad
Dim rec As New ADODB.Recordset
Dim x As Integer

Set rec = opndbaseFMIS.Execute("Select * from tblAMIS_Field")
If rec.RecordCount > 0 Then
cmb_Field.Clear
    For x = 1 To rec.RecordCount
    cmb_Field.AddItem (rec!Fieldname)
    rec.MoveNext
    Next x
End If
rec.Close
cmb_Field.ListIndex = 0
Exit Sub
bad:
MsgBox err.description
End Sub
Private Sub Command4_Click()
Dim rec As New ADODB.Recordset
Dim Lastno As Double
Dim TYP As Integer
'    If Option7.Value = False Then
'        MsgBox "Please Specify the date of your Post date..", vbInformation, "System Message"
'        Exit Sub
'    End If
    
        If Trim(txtfundtype.Text) <> "" Then
        If Option2.Value = True Then: TYP = Option2.Tag
        If Option3.Value = True Then: TYP = Option3.Tag
        If Option4.Value = True Then: TYP = Option4.Tag
        If Option5.Value = True Then: TYP = Option5.Tag
        
        frmPOstdate.Show 1
        If JevOk = True Then
            Set rec = opndbaseFMIS.Execute("EXEC [dbo].[Proc_GetMaxJevSeries_new] @transtype = " & TYP & ",@jevyeardate = '" & DatePost & "' ,@fundtype = '" & txtfundtype.Text & "'")
            LastJEVNo = rec.Fields!MAXJEVSERIES
            rec.Close
            txtJEVNo.Text = txtfundtype.ItemData(txtfundtype.ListIndex) & "-" & Right(Year(DatePost), 2) & "-" & Format(Month(DatePost), "00") & "-" & Format(TYP, "00") & "-" & Format(LastJEVNo, "0000")
            txtjevdate.Text = Format(DatePost, "MMMM dd, yyyy")
            DTpMY.Value = DatePost
        Else
        MsgBox "Cannot Generate the System JEV Number,If you cancel to Set the Date", vbInformation, "System Message"
        End If
        
'
'
'            rec.Open ("EXEC [dbo].[Proc_GetMaxJevSeries_New] @transtype = " & TYP & ",@jevyeardate = '" & DTpMY.Value & "' ,@fundtype = '" & txtfundtype.Text & "'"), opndbaseFMIS, adOpenStatic, adLockOptimistic
'                Lastno = rec.Fields!MAXJEVSERIES
'            rec.Close
'            txtjevno.Text = txtfundtype.ItemData(txtfundtype.ListIndex) & "-" & Right(Year(DTpMY), 2) & "-" & Format(Month(DTpMY), "00") & "-" & Format(TYP, "00") & "-" & Format(Lastno, "0000")
'            txtjevdate.Text = Format(DTpMY, "MMMM yyyy")
        Else
            MsgBox "Opps..!Please Specify Special Accounts First to Proceed the Transaction", vbInformation, "System Message"
        End If
End Sub

Private Sub DTPicker2_CallbackKeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer, ByVal CallbackField As String, CallbackDate As Date)
txtcheckdate.Text = Format(DTPicker2.Value, "MM/dd/yyyy")
End Sub

Private Sub DTPicker2_Change()
txtcheckdate.Text = Format(DTPicker2.Value, "MM/dd/yyyy")
End Sub

Private Sub DTPicker2_Click()
txtcheckdate.Text = Format(DTPicker2.Value, "MM/dd/yyyy")
End Sub

Private Sub Form_Activate()
'ErrDll.CenterMe frmFinalJev
End Sub

Private Sub Form_Load()

cmb_Field.Text = "JEVNo"
Call LoadFinalTrans(cmb_Field.Text, txtcondition.Text)
isEdit = False
Load
Call LoadFundType(txtfundtype)
Call LoadOOE
Call LoadOffice
Call loadField
DTPYear.Value = Now
DTpMY.Value = Now
txt ("Disable")
SSTab2.Tab = 0
If CheckIfADMIN(ActiveUserID) = False Then
    Check2.Visible = False
End If
'ErrDll.CenterMe frmFinalJev
End Sub
Private Function LoadOffice()
Dim OREc As New ADODB.Recordset
Dim x As Integer

txtRC.Clear
        OREc.Open ("Select * FRom tblREF_AIS_Offices Order By [OfficeMedium]"), opndbaseFMIS, adOpenStatic, adLockOptimistic
        If OREc.RecordCount > 0 Then
        txtRC.AddItem ""
            For x = 1 To OREc.RecordCount
                txtRC.AddItem OREc![OfficeMedium]
                txtRC.ItemData(txtRC.NewIndex) = OREc!fmisofficeid
                OREc.MoveNext
            Next x
        End If
        OREc.Close
        Set OREc = Nothing
End Function
Private Sub LoadOOE()
Dim OREc As New ADODB.Recordset
Dim x As Integer

txtOOE.Clear

OREc.Open ("Select * From tblBMS_ObjectOfExpenditures Order By OOEName"), opndbaseFMIS, adOpenStatic, adLockOptimistic
If OREc.RecordCount > 0 Then
        txtOOE.AddItem ""
        'txtooe.ItemData(1) = OREc!OOECode
    For x = 1 To OREc.RecordCount
        txtOOE.AddItem OREc!OOEName
        txtOOE.ItemData(txtOOE.NewIndex) = OREc!OOECode
        OREc.MoveNext
    Next x
End If
OREc.Close
Set OREc = Nothing

End Sub

Private Function LoadFinalTrans(ByVal field As String, ByVal Condition As String)
'On Error GoTo bad
Dim PRec As New ADODB.Recordset
Dim x As Integer
Dim Transtype As String
Dim Postdate As String
Dim year_ As Long
Dim month_, posteD As Integer
Dim whatfield As Integer
Dim orderby, sqlPosted As String
Transtype = ""
If Option2.Value = True Then: Transtype = Option2.Tag
If Option3.Value = True Then: Transtype = Option3.Tag
If Option4.Value = True Then: Transtype = Option4.Tag
If Option5.Value = True Then: Transtype = Option5.Tag
If Option6.Value = True Then: Transtype = Option6.Tag

Postdate = ""
If Option1.Value = True Then: Postdate = "year(jevdate) = " & DTPYear.Year & ""
If Option7.Value = True Then: Postdate = "year(jevdate) = " & DTpMY.Year & " and month(jevdate) = " & DTpMY.Month & ""

If Option1.Value = True Then
year_ = DTPYear.Year
whatfield = 2
Else
year_ = DTpMY.Year
month_ = DTpMY.Month
whatfield = 1
End If


orderby = field
If field = "JEVno" Then: orderby = "substring(Jevno,14,7)"

    List1.Clear
    posteD = 0
    If Check3.Value = 1 Then
    sqlPosted = "and posted = 0"
    ElseIf Check4.Value = 1 Then
    sqlPosted = "and HaveDoc = 1"
    Else
    sqlPosted = ""
    End If
    If Check1.Value = 1 Then
    PRec.Open ("Exec MPproc_FindpostedTransThrougObno @OBRNO ='" & txtcondition.Text & "',@transtype =  " & Transtype & ",@year = " & year_ & ",@month = '" & month_ & "',@what = " & whatfield & ""), opndbaseFMIS, adOpenStatic, adLockOptimistic
    Else
    'MsgBox "Select " & field & " as field From tblAMIS_FinalJEV Where " & field & " like '" & Condition & "%' and " & Postdate & " and transtype in (" & Transtype & ") and ltrim(" & field & ") <> '' and Actioncode=1 " & sqlPosted & "  Group By " & field & " order by " & orderby & ""
    PRec.Open ("Select " & field & " as field From tblAMIS_FinalJEV Where " & field & " like '" & Condition & "%' and " & Postdate & " and transtype in (" & Transtype & ") and ltrim(" & field & ") <> '' and Actioncode=1 " & sqlPosted & "  Group By " & field & " order by " & orderby & ""), opndbaseFMIS, adOpenStatic, adLockOptimistic
    End If
    If PRec.RecordCount > 0 Then
    lblcount.Caption = PRec.RecordCount & " Record(s) found"
        For x = 1 To PRec.RecordCount
            List1.AddItem PRec!field
            PRec.MoveNext
            DoEvents
        Next x
        List1.Enabled = True
    End If
    PRec.Close
    Set PRec = Nothing
Exit Function
bad:
MsgBox err.description
End Function

Private Sub List1_Click()
isEdit = True
IsSaveAccntng = False
Call Loaddetails(cmb_Field.Text, List1.Text)

End Sub
Private Function Loaddetails(ByVal field As String, ByVal Condition As String)
Dim rec As New ADODB.Recordset
On Error Resume Next
Dim Transtype As String
Transtype = ""
If Option2.Value = True Then: Transtype = Option2.Tag
If Option3.Value = True Then: Transtype = Option3.Tag
If Option4.Value = True Then: Transtype = Option4.Tag
If Option5.Value = True Then: Transtype = Option5.Tag
If Option6.Value = True Then: Transtype = Option6.Tag

rec.Open "Select * from tblAMIS_FinalJEV where " & field & " = '" & Condition & "' and  actioncode = 1 and [Transtype] = '" & Transtype & "'", opndbaseFMIS, adOpenStatic, adLockOptimistic
    If rec.RecordCount <> 0 Then
       ' Clrtxt
        txt ("Clear")
        txtcheckdate.Text = Trim(rec.Fields!Date_)
        txtCheckno.Text = Trim(rec.Fields!checkno)
        txtClaimant.Text = getClaimant(rec.Fields!ClaimantCode)
        ClaimantCode = Trim(rec.Fields!ClaimantCode)
        txtDVNo.Text = Trim(rec.Fields!dvno)
        txtfundtype = Trim(rec.Fields!FundType)
        txtjevdate.Text = Format(Trim(rec.Fields!jevdate), "MMMM yyyy")
        DTpMY.Value = rec.Fields!jevdate
        Jdate = rec.Fields!jevdate
        txtJEVNo.Text = Trim(rec.Fields!jevno)
        trans = rec.Fields!Transtype
        txtobrno.Text = IIf((Left(Trim(rec.Fields!obrno), 2) = "NA"), GetNonAlobsName(rec.Fields!obrno), rec.Fields!obrno)
        'If Left(Trim(rec.Fields!obrno), 2) = "NA" Then: txtobrno.Text = GetNonAlobsName(rec.Fields!obrno)
        txtOOE.Text = (FindField("ooe", "tblAMIS_IncomingDVTrns", "dvno", rec.Fields!dvno, "actioncode = 1"))
        txtParticular.Text = Trim(rec.Fields!Particular)
        txtptvno.Text = Trim(rec.Fields!ptvNo)
        txtgamount.Text = Format(rec.Fields!Gamount, "#,##0.00")
        txtRC.Text = GetOfficeName(IIf(Trim(rec!RCenter) = "", "0", Trim(rec!RCenter)), "OfficeMedium")
        FmisVoucherno = rec.Fields!FmisVoucherno
        txtrcino.Text = rec.Fields!RCI
        txtrdono.Text = rec.Fields!RDOno
        Check2.Value = IIf((rec.Fields!PClosinG = True), 1, 0)
        ref = IIf(IsNull(rec.Fields!RefNo), "", (rec.Fields!RefNo))
        isPOSTED = IIf((rec.Fields!posteD = 1), True, False)
        Jevseries = rec.Fields!jevseriesno
        Check5.Value = IIf((rec.Fields!HaveDoc = 1), 1, 0)
            Call LoadEntry(SSTab2.Tab)
            Call GetSum
    End If
        
rec.Close
Set rec = Nothing
End Function
Public Sub LoadEntry(id As Integer)
If SSTab2.Tab = 0 Then
    Call GetAccntngEntries
ElseIf SSTab2.Tab = 1 Then
    Call GetCashFlowEntries
ElseIf SSTab2.Tab = 2 Then
    Call LoadJEVLogEntry
ElseIf SSTab2.Tab = 3 Then
    Call LoadEntryInGrid(MSHFlexGrid2, 2, txtrdono.Text, 3)
ElseIf SSTab2.Tab = 4 Then
    Call LoadEntryInGrid(MSHFlexGrid3, 3, txtJEVNo.Text, 4)
    lblstat.Caption = GetStatOfPostedTransaction(txtJEVNo.Text)
End If
End Sub
Public Sub GetFinalStat(ByVal jevno As String)
Dim rec As New ADODB.Recordset
    
End Sub
Private Function ChkEntry() As Boolean
Dim Gamount As Double
    ChkEntry = False
        If xDebit = xCredit And xDebit > 0 Then
            If Format(xDebit, "###,##0.00") <= Format(txtgamount.Text, "###,##0.00") Then
            ChkEntry = True
            ElseIf Format(xDebit, "###,##0.00") > Format(txtgamount.Text, "###,##0.00") Then
                If MsgBox("Your Gross Amount is Less than to your total Debit or Credit Amount" & vbNewLine & "Are you sure the transaction have Corolary entry?", vbCritical + vbYesNo, "System Information") = vbYes Then
                   ChkEntry = True
                Else
                   MsgBox "Saving CANCEL", vbInformation, "System Message"
                End If
            End If
        End If
End Function

Private Function Clrtxt()
        txtcheckdate.Text = ""
        txtCheckno.Text = ""
        txtClaimant.Text = ""
        txtDVNo.Text = ""
       'txtfundtype.Text = ""
        txtjevdate.Text = ""
        txtJEVNo.Text = ""
        txtobrno.Text = ""
'        txtooe.Text = ""
        txtParticular.Text = ""
        txtptvno.Text = ""
'        txtrc.Text = ""
        txtrcino.Text = ""
        txtrdono.Text = ""
        ref = ""
        Jevseries = 0
            Call SetGrid
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
Private Sub GetSum2()

Dim x As Integer
    xDebit2 = 0
    xCredit2 = 0
    For x = 1 To MSFlexGrid2.Rows - 1
        If MSFlexGrid2.TextMatrix(x, 1) <> "" Then
            xDebit2 = xDebit2 + CCur(IIf(MSFlexGrid2.TextMatrix(x, 3) = "", 0, MSFlexGrid2.TextMatrix(x, 3)))
            xCredit2 = xCredit2 + CCur(IIf(MSFlexGrid2.TextMatrix(x, 4) = "", 0, MSFlexGrid2.TextMatrix(x, 4)))
        Else
            MSFlexGrid2.TextMatrix(x, 2) = "TOTAL"
            MSFlexGrid2.TextMatrix(x, 3) = Format(xDebit2, "#,##0.00")
            MSFlexGrid2.TextMatrix(x, 4) = Format(xCredit2, "#,##0.00")
            Exit For
        End If
    Next x
End Sub
Public Sub LoadAccountsByFund(ByVal FundName As String)
Dim ARec As New ADODB.Recordset
Dim x As Integer
If Left(FundName, 3) = "Eco" Then
FundName = "Economic Enterprises"
End If
    cmbEntry.Clear
    cmbEntry.Visible = False
    ARec.Open ("Exec GetAccountcode @fundtype = '" & FundName & "'"), opndbaseFMIS, adOpenStatic, adLockOptimistic
    If ARec.RecordCount > 0 Then
        Do Until ARec.EOF
            cmbEntry.AddItem ARec![childaccountcode]
            cmbEntry.ItemData(cmbEntry.NewIndex) = ARec![FmisAccountcode]
            ARec.MoveNext
       Loop
    End If
    ARec.Close
    Set ARec = Nothing
    
End Sub

Private Sub lvButtons_H1_Click()
Dim rec As New ADODB.Recordset
If CheckIfADMIN(ActiveUserID) = True Then
    If MsgBox("JEV approval will trigger posting to GL and SL. Approve this JEV?", vbInformation + vbYesNo, "Syetem Confirmation") = vbYes Then
        Select Case DisApprovedAndApprove(txtJEVNo.Text, "", 1)
        Case 0
            MsgBox "No transaction Selected..", vbInformation, "System Information"
        Case 1
            MsgBox "Transaction Approved..", vbInformation, "System Messgae"
            isPOSTED = True
        Case 2
            MsgBox "Unable to filter in Cash Flow, Please Filter manualy..", vbInformation, "System Messgae"
            Call MSFlexGrid2_Click
        Case 3
            MsgBox "Unidentified transaction..Please Contact System Administrator", vbInformation, "System Messgae"
        Case 4
            MsgBox "Transaction Disapproved..", vbInformation, "System Messgae"
            isPOSTED = False
        End Select
        lblstat.Caption = GetStatOfPostedTransaction(txtJEVNo.Text)
    End If
Else
MsgBox "This is for Administrator Privilege..", vbInformation, "System Message"
End If
End Sub

Private Sub lvButtons_H2_Click()
Dim rec As New ADODB.Recordset
If CheckIfADMIN(ActiveUserID) = True Then
If MsgBox("Are you sure do you want to DISAPPROVE this transation?", vbInformation + vbYesNo, "Syetem Confirmation") = vbYes Then
    Select Case DisApprovedAndApprove(txtJEVNo.Text, "", 2)
    Case 0
        MsgBox "No transaction Selected..", vbInformation, "System Information"
    Case 1
        MsgBox "Transaction Approved..", vbInformation, "System Messgae"
    Case 2
        MsgBox "Unable to filter in Cash Flow, Please Filter manualy..", vbInformation, "System Messgae"
        Call MSFlexGrid2_Click
    Case 3
        MsgBox "Unidentified transaction..Please Contact System Administrator", vbInformation, "System Messgae"
    Case 4
        MsgBox "Transaction Disapproved..", vbInformation, "System Messgae"
        frmSub3.isPOSTED = False
        isPOSTED = False
    End Select
    lblstat.Caption = GetStatOfPostedTransaction(txtJEVNo.Text)
End If
Else
MsgBox "This is for Administrator Privilege..", vbInformation, "System Message"
End If
End Sub

Private Sub MSFlexGrid1_DblClick()
If IsNumeric(txtgamount.Text) = True Then
    With frmSub3
        .isPOSTED = isPOSTED
        .REFF = txtJEVNo.Text
        .Gamount = txtgamount.Text
        .CName = UCase(txtClaimant.Text)
        .WhatTab = 0
        .isEdit = True
        EditCount = False
        Set .frm = Me
        Call LoadAcctngEntries(txtJEVNo.Text)
        .Show 1
        Call GetAccntngEntries
        Call GetCashFlowEntries
    End With
End If
End Sub
Public Function LoadAcctngEntries(ByVal jevno As String)
Dim Drec As New ADODB.Recordset
Dim rec As New ADODB.Recordset
Dim x As Integer
    Set Drec = opndbaseFMIS.Execute("Select Accountcode,Debit ,Credit From tblAMIS_postedJEV Where [JEVno]='" & jevno & "' And (ActionCode=1) ")
    If Drec.RecordCount > 0 Then
        If EditCount = False Then
            EditCount = True
            rec.Open "Select dvno from tblAMIs_tmpjournal where dvno = '" & jevno & "'", opndbaseFMIS, adOpenStatic
            If rec.RecordCount > 0 Then
                If MsgBox("This transaction Have a temporary Accounting Entries, do you want to Delete?", vbCritical + vbYesNo, "System Information") = vbYes Then
                   opndbaseFMIS.Execute ("EXECUTE  [fmis].[dbo].[MPproc_InsertJEVEntry] @Field = '" & txtJEVNo.Text & "',@whatField = 'JEVNO'")
                End If
            Else
                opndbaseFMIS.Execute ("EXECUTE  [fmis].[dbo].[MPproc_InsertJEVEntry] @Field = '" & txtJEVNo.Text & "',@whatField = 'JEVNO'")
            End If
            rec.Close
        Else
            opndbaseFMIS.Execute ("EXECUTE  [fmis].[dbo].[MPproc_InsertJEVEntry] @Field = '" & txtJEVNo.Text & "',@whatField = 'JEVNO'")
        End If
    End If
    Drec.Close
    Set Drec = Nothing
End Function



Private Function Load()
If Option1.Value = True Then
    DTPYear.Enabled = True
    DTpMY.Enabled = False
Else
DTPYear.Enabled = False
    DTpMY.Enabled = True
End If

End Function

Private Sub MSFlexGrid2_Click()
If IsNumeric(txtgamount.Text) = True Then
    With frmSub3
        .isPOSTED = isPOSTED
        .REFF = txtJEVNo.Text
        .Gamount = txtgamount.Text
        .CName = UCase(txtClaimant.Text)
        .isEdit = True
        .WhatTab = 1
        EditCount = False
        Set .frm = Me
        Call LoadAcctngEntries(txtJEVNo.Text)
        .Show 1
        Call GetAccntngEntries
        Call GetCashFlowEntries
    End With
End If
End Sub

Private Sub MSHFlexGrid3_DblClick()
frm_JevViewer.jevno = txtJEVNo.Text
frm_JevViewer.Show 1
End Sub

Private Sub Option1_Click()
Load
End Sub

Private Sub Option6_Click()
Call LoadFinalTrans(cmb_Field.Text, txtcondition.Text)
End Sub

Private Sub Option7_Click()
Load
End Sub

Private Sub TabStrip1_Click()

End Sub

Private Sub SSTab2_Click(PreviousTab As Integer)
Call LoadEntry(SSTab2.Tab)
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim rec As New ADODB.Recordset
Dim TYP As Integer
'On Error GoTo bad

If Option2.Value = True Then: TYP = Option2.Tag
If Option3.Value = True Then: TYP = Option3.Tag
If Option4.Value = True Then: TYP = Option4.Tag
If Option5.Value = True Then: TYP = Option5.Tag
If Option6.Value = True Then: TYP = Option6.Tag

Select Case Button

    Case "New"
                isPOSTED = False
                Call SetGrid
                Option7.Value = True
                Load
                Check2.Value = 0
                txt ("Clear")
                txt ("Enable")
                lblMode.Caption = "NEW"
                Toolbar1.Buttons(13).Caption = "Cancel"
                Toolbar1.Buttons(5).Visible = False
                isEdit = False
                Frame2.Enabled = False
                'DTpMY.Value = Now
    Case "Save"
                If isPOSTED = True Then
                    MsgBox "This transaction is Already Generate the report, Unable to EDIT the transaction...!", vbInformation, "System Message"
                    Exit Sub
                End If
                If txtJEVNo.Text <> "" And txtjevdate.Text <> "" And txtfundtype.Text <> "" And txtcheckdate.Text <> "" And txtParticular.Text <> "" And txtgamount.Text <> "" Then
                    If ChkEntry = True Then
                        Call SaveFinalJEV
                    Else
                        MsgBox "Total Debit and Total Credit Amount is not Balance" & vbNewLine & "Please Check Your Entry", vbInformation, "System Message"
                    End If
                    
                Else
                    MsgBox "Complete the Priority field Such as:" & vbNewLine & "(JEV No.,JEV date, Special Account,Checkdate/Deposit date/Transaction date,Particular and Gross Amount", vbInformation, "System Message"
                    Exit Sub
                End If
    Case "Edit"
                If isPOSTED = True Then
                    MsgBox "This transaction is Already Generate the report, Unable to EDIT the transaction...!", vbInformation, "System Message"
                    Exit Sub
                End If
                isEdit = True
                Toolbar1.Buttons(13).Caption = "Cancel"
'                Picture1.Enabled = True
                lblMode.Caption = "EDIT"
                Frame2.Enabled = False
                Frame1.Enabled = True
               ' Call LoadAccountsByFund(txtfundtype.Text)
    Case "Delete"
                If isPOSTED = True Then
                    MsgBox "This transaction is Already Generate the report, Unable to EDIT the transaction...!", vbInformation, "System Message"
                    Exit Sub
                End If
                If MsgBox("Are you sure do you want to delete from the Report?", vbInformation + vbYesNo, "System Message") = vbYes Then
                    Select Case (TYP)
                    Case 1
                        opndbaseFMIS.Execute "Update tblCMS_CDCheckBook set [AlreadySaved2JEV] = 0 where dvno = '" & txtptvno.Text & "'"
                    Case 2
                        opndbaseFMIS.Execute "Update tblCMS_CDRCIReport set [AlreadySaved2JEV] = 0 where Checkno = '" & txtCheckno.Text & "'"
                    Case 3
                        opndbaseFMIS.Execute "Update tblCMS_CDCashBook set [AlreadySaved2JEV] = 0 where alobsno = '" & txtDVNo.Text & "'"
                    Case 4
                        opndbaseFMIS.Execute "Update tblAMIS_CreditNotice set AlreadySaved2JEV = 0 where dvno = '" & txtDVNo.Text & "'"
                    Case 4
                        opndbaseFMIS.Execute "Update tblAMIS_CreditNotice set AlreadySaved2JEV = 0 where dvno = '" & txtDVNo.Text & "'"
                    Case Else
                        opndbaseFMIS.Execute "Update tblCMS_CDCheckBook set [AlreadySaved2JEV] = 0 where dvno = '" & txtptvno.Text & "'"
                        opndbaseFMIS.Execute "Update tblCMS_CDRCIReport set [AlreadySaved2JEV] = 0 where dvno = '" & txtptvno.Text & "'"
                        opndbaseFMIS.Execute "Update tblCMS_CDCashBook set [AlreadySaved2JEV] = 0 where alobsno = '" & txtDVNo.Text & "'"
                        opndbaseFMIS.Execute "Update tblAMIS_CreditNotice set AlreadySaved2JEV = 0 where dvno = '" & txtDVNo.Text & "'"
                    End Select
                    
                    opndbaseFMIS.Execute "Update tblAMIS_PostedJEV set actioncode = 3,[Datetimeentered] = rtrim(ltrim(datetimeentered)) + '" & Now & "',[UserId] = rtrim(ltrim(userid)) + '" & ActiveUserID & "'  where jevno = '" & txtJEVNo.Text & "' and actioncode = 1"
                    opndbaseFMIS.Execute "Delete from tblAMIS_PostedJEVforCashflow where jevno = '" & txtJEVNo.Text & "'"
                    opndbaseFMIS.Execute "Update tblAMIS_FinalJEV set actioncode = 3 where jevno = '" & txtJEVNo.Text & "' and actioncode = 1"
                    
                    If List1.ListCount > 0 Then
                    List1.RemoveItem (List1.ListIndex)
                    End If
                    Clrtxt
                End If
    Case "Adjustment"
                If Trim(txtJEVNo.Text) <> "" Then
                centerme frmJEVPreparationforAjustment_new
                    With frmJEVPreparationforAjustment_new
                    .txtDVNo = txtJEVNo.Text
                    .txtFund.Text = txtfundtype.Text
                    .ClaimantCode = ClaimantCode
                    .RCenter = txtRC.ItemData(txtRC.ListIndex)
                    .cmbrc.Text = txtRC.Text
                    .Show
                    End With
                Else
                MsgBox "No Supporting/Refference to Make Adjustment", vbInformation, "Systen Message"
                End If
    Case "Close"
                If MsgBox("Are you sure do want to close this Form?", vbCritical + vbYesNo, "System Information") = vbYes Then
                Unload Me
                End If
    Case "Print JEV"
    PrintJEV
    Case "Cancel"
                isEdit = False
                txt ("Disable")
                txt ("Clear")
                Call SetGrid
                Frame2.Enabled = True
                Toolbar1.Buttons(13).Caption = "Close"
                Toolbar1.Buttons(5).Visible = True
                Frame1.Enabled = False
End Select
Exit Sub
bad:
Call LoadErr(err.Number, err.Source & ", " & Me.name & ", " & Me.Caption & ", " & "Toolbar1_ButtonClick-", err.description)
End Sub
Private Sub PrintJEV()
Dim sql As String
    sql = "Exec Proc_JevPostedPrinting @JEVno = '" & Trim(txtJEVNo.Text) & "'"
    ReportName = "JEVNEW"
    rptJEVNew.txtClaimDesc.SetText txtParticular.Text & ", " & txtClaimant.Text & ", " & txtobrno.Text
    rptJEVNew.txtRC.SetText txtRC.Text
    rptJEVNew.txtClerk.SetText getUserName(ActiveUserID, "FullName")
    rptJEVNew.Text23.SetText GetEmpPosition(ActiveUserID)
    rptJEVNew.txtJEVNo.SetText txtJEVNo.Text
    rptJEVNew.txtDate.SetText Format(Jdate, "MM/dd/yyyy")
    
    
    rptJEVNew.Trantype = 1
    
'    If chkSTP.Value = 1 Then
'        rptJEVNew.Line1.Suppress = True
'        rptJEVNew.Line2.Suppress = True
'        rptJEVNew.Line3.Suppress = True
'        rptJEVNew.Line4.Suppress = True
'        rptJEVNew.Line5.Suppress = True
'        rptJEVNew.Line6.Suppress = True
'        rptJEVNew.Line8.Suppress = True
'        rptJEVNew.Line9.Suppress = True
'        rptJEVNew.Line10.Suppress = True
'        rptJEVNew.Line11.Suppress = True
'        rptJEVNew.Line12.Suppress = True
'        rptJEVNew.Line13.Suppress = True
'        rptJEVNew.Line14.Suppress = True
'        rptJEVNew.Line15.Suppress = True
'        rptJEVNew.Line16.Suppress = True
'        rptJEVNew.Line17.Suppress = True
'
'        rptJEVNew.Line19.Suppress = True
'
'        rptJEVNew.Text1.Suppress = True
'        rptJEVNew.Text2.Suppress = True
'        rptJEVNew.Text3.Suppress = True
'        rptJEVNew.Text4.Suppress = True
'        rptJEVNew.Text8.Suppress = True
'        rptJEVNew.Text9.Suppress = True
'        rptJEVNew.Text12.Suppress = True
'        rptJEVNew.Text13.Suppress = True
'        rptJEVNew.Text15.Suppress = True
'        rptJEVNew.Text16.Suppress = True
'        rptJEVNew.Text17.Suppress = True
'        rptJEVNew.Text18.Suppress = True
'        rptJEVNew.Text19.Suppress = True
'        rptJEVNew.Text20.Suppress = True
'        rptJEVNew.Text21.Suppress = True
'        rptJEVNew.Text22.Suppress = True
'        rptJEVNew.Text25.Suppress = True
'
'    End If
    rptJEVNew.DiscardSavedData
    rptJEVNew.Database.SetDataSource opndbaseFMIS.Execute(sql)
    rptJEVNew.Database.Verify
     If Option5.Value = True Then rptJEVNew.Trantype = 1
    If Option2.Value = True Then rptJEVNew.Trantype = 2
    If Option3.Value = True Then rptJEVNew.Trantype = 3
    If Option4.Value = True Then rptJEVNew.Trantype = 4
    If Option6.Value = True Then rptJEVNew.Trantype = 4
   frmViewer.Show 1
End Sub
Public Function SaveFinalJEV()
Dim Credit As Currency
Dim Debit As Currency
Dim tmp As Long
PClosinG = 0
If Option2.Value = True Then: trans = Option2.Tag
If Option3.Value = True Then: trans = Option3.Tag
If Option4.Value = True Then: trans = Option4.Tag
If Option5.Value = True Then: trans = Option5.Tag
If Check2.Value = 1 Then: PClosinG = 1
Jevseries = ExtractJEVSNo(txtJEVNo)

If MsgBox("Are You Sure Do you want to Save/Update these entry?", vbInformation + vbYesNo) = vbYes Then
    If isEdit = True Then
        opndbaseFMIS.Execute "update tblAMIS_FinalJEV set actioncode = 2 where jevno = '" & txtJEVNo.Text & "' and actioncode= 1"
    Else
    
        If CheckIfExistInFinalJEV(txtJEVNo.Text) = True Then
                If MsgBox("JEV Number Already exist on the database,Do you want the system Generate the JEV number?" & vbNewLine & "Click YES to generate,NO to cancel..!", vbInformation + vbYesNo, "System Message") = vbYes Then
                Call Command4_Click
                Else
                Exit Function
                End If
        End If
    End If
                Call Saved2FinalJEV_forFinalJEV(txtcheckdate.Text, Trim(txtrcino.Text), txtCheckno.Text, txtParticular.Text, txtJEVNo.Text, ClaimantCode, 0, txtgamount.Text, 0, 0, trans, FmisVoucherno, txtDVNo.Text, txtobrno.Text, txtfundtype.Text, RCcode, txtOOE.Text, txtrdono.Text, ref, Jevseries, DTpMY.Value, txtptvno.Text, PClosinG)
                If Check5.Value = 1 Then
                    opndbaseFMIS.Execute "update tblAMIS_FinalJEV set HaveDoc = 1 where jevno = '" & txtJEVNo.Text & "' and actioncode = 1"
                Else
                    opndbaseFMIS.Execute "update tblAMIS_FinalJEV set HaveDoc = 0 where jevno = '" & txtJEVNo.Text & "' and actioncode = 1"
                End If
                
                If IsSaveAccntng = True Then
                    Call SaveAcctngEntries(txtJEVNo.Text)
                End If
                Call LoadFinalTrans(cmb_Field.Text, txtcondition.Text)
                Call Toolbar1_ButtonClick(Toolbar1.Buttons.Item(1))
End If
End Function
Public Function SaveAcctngEntries(ByVal jevno As String)
Dim Drec As New ADODB.Recordset
Dim rec As New ADODB.Recordset
Dim x As Integer

    Drec.Open ("Select Accountcode,sum(Debit) as Debit ,sum(Credit) as Credit From tblAMIs_tmpjournal Where [dvno]='" & jevno & "' group by accountcode"), opndbaseFMIS, adOpenStatic, adLockOptimistic
    If Drec.RecordCount > 0 Then
        opndbaseFMIS.Execute "update tblAMIS_postedJEV set actioncode =2 where JEVNO = '" & jevno & "' and actioncode =1" ', datetimeentered = rtrim(ltrim(DateTimeEntered)) +'," & Format(Now, "yyyy/mm/dd hh:mm:ss AMPM") & "',userid = rtrim(ltrim(UserID)) + '," & Trim(ActiveUserID) & "'
            progStat.Max = Drec.RecordCount
            lblmsg.Caption = "Saving...." & Drec.RecordCount & "/" & Drec.RecordCount
            lblmsg.Visible = True
            progStat.Visible = True
        For x = 1 To Drec.RecordCount
            DoEvents
            opndbaseFMIS.Execute "Insert into tblAMIS_PostedJEV (JEVNO,Accountcode,debit,credit,actioncode,datetimeentered,userid) values " & _
            "('" & Trim(jevno) & "','" & Trim(Drec!accountcode) & "'," & Drec!Debit & "," & Drec!Credit & ",1,'" & Format(Now, "yyyy/mm/dd hh:mm:ss AMPM") & "','" & Trim(ActiveUserID) & "')"
            Drec.MoveNext
            lblmsg.Caption = "Saving...." & x & "/" & Drec.RecordCount
            progStat.Value = x
        Next x
        opndbaseFMIS.Execute "delete from tblAMIs_tmpjournal where dvno = '" & jevno & "'"
    End If
    Drec.Close
    progStat.Visible = False
    lblmsg.Visible = False
    Set Drec = Nothing
End Function
Private Sub txt_entry_KeyPress(KeyAscii As Integer)
 On Error GoTo bad
    If KeyAscii = 13 Then
            If IsNumeric(txt_entry.Text) = False And txt_entry.Text <> "" Then
                MsgBox "None Numeric Entry, Please Check Your Entry", vbCritical, "System Message"
                Exit Sub
            End If
            MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, MSFlexGrid1.col) = Format((txt_entry.Text), "#,##0.00")
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
                If MSFlexGrid1.col = 5 Then
                    If txt_entry.Text = "1" Or txt_entry.Text = "5" Then
                    MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, MSFlexGrid1.col) = txt_entry.Text
                    Else
                    MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, MSFlexGrid1.col) = "1"
                    End If
                End If
        Call GetSum
    End If
Exit Sub
bad:
    Call LoadErr(err.Number, err.Source & ", " & Me.name & ", " & Me.Caption, err.description)
End Sub


Private Sub Option2_Click()
Call LoadFinalTrans(cmb_Field.Text, txtcondition.Text)
End Sub

Private Sub Option3_Click()
Call LoadFinalTrans(cmb_Field.Text, txtcondition.Text)
End Sub

Private Sub Option4_Click()
Call LoadFinalTrans(cmb_Field.Text, txtcondition.Text)
End Sub

Private Sub Option5_Click()
Call LoadFinalTrans(cmb_Field.Text, txtcondition.Text)
End Sub

Private Sub txtcondition_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Call LoadFinalTrans(cmb_Field.Text, txtcondition.Text)
End If
End Sub

Private Sub txtfundtype_Change()
'If isEdit = False Then
'Call LoadAccountsByFund(txtfundtype.Text)
'End If
End Sub
Private Sub txtfundtype_Click()
'If isEdit = False Then
'Call LoadAccountsByFund(txtfundtype.Text)
'End If
End Sub

Private Sub txtgamount_LostFocus()
txtgamount.Text = Format(txtgamount.Text, "#,###0.00")
End Sub

Private Sub txtRC_Click()
If txtRC.ListIndex <> -1 Then
RCcode = txtRC.ItemData(txtRC.ListIndex)
End If
End Sub
Public Sub GetAccntngEntries()
Dim Drec As New ADODB.Recordset
Dim x As Integer
Call SetGrid
    'DRec.Close
    If IsSaveAccntng = False Then
        Set Drec = opndbaseFMIS.Execute("Select left(Accountcode,3) as childcode,sum(Debit) as sumdebit,sum(credit) as sumcredit From tblAMIS_POSTEDJEV Where [JEVno]='" & txtJEVNo.Text & "' And (ActionCode=1) group by jevno,actioncode,left(Accountcode,3) order by sumdebit desc")
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
        Set Drec = opndbaseFMIS.Execute("Select left(Accountcode,3) as childcode,sum(Debit) as sumdebit,sum(credit) as sumcredit From tblAMIS_tmpjournal Where [dvno]='" & txtJEVNo.Text & "' group by Dvno,left(Accountcode,3)")
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
Public Sub GetCashFlowEntries()
Dim Drec As New ADODB.Recordset
Dim x As Integer
Call SetGrid2
    'DRec.Close
    If IsSaveAccntng = False Then
        'Set DRec = opndbaseFMIS.Execute("Select left(Accountcode,3) as childcode,sum(Debit) as sumdebit,sum(credit) as sumcredit From tblAMIS_POSTEDJEV Where [JEVno]='" & txtJEVNo.Text & "' And (ActionCode=1) group by jevno,actioncode,left(Accountcode,3) order by sumdebit desc")
        Set Drec = opndbaseFMIS.Execute("SELECT  sum([Debit]) as Sumdebit,sum([Credit]) as Sumcredit, [Accountcode] as childcode FROM [fmis].[dbo].[tblAMIS_PostedJEVforCashflow] where [JEVno]='" & txtJEVNo.Text & "'  group by accountcode order by sumdebit desc")
        If Drec.RecordCount > 0 Then
            For x = 1 To Drec.RecordCount
    '            MSFlexGrid1.TextMatrix(x, 0) = DRec![Trnno]
                MSFlexGrid2.TextMatrix(x, 1) = Drec!childcode
                MSFlexGrid2.TextMatrix(x, 2) = GetAccountNameByAccountcode(Drec!childcode)
                MSFlexGrid2.TextMatrix(x, 4) = IIf((Format(Drec!sumCredit, "#,##0.00") = "0.00"), "", Format(Drec!sumCredit, "#,##0.00"))
                MSFlexGrid2.TextMatrix(x, 3) = IIf((Format(Drec!sumDebit, "#,##0.00") = "0.00"), "", Format(Drec!sumDebit, "#,##0.00"))
              MSFlexGrid2.Rows = MSFlexGrid2.Rows + 1
               ' If LCase(Trim(lblMode.Caption)) = "edit" Then MSFlexGrid2.TextMatrix(x, 5) = DRec!ActionCode  ' for coloraly purpose
                Drec.MoveNext
            Next x
            
        End If
    Else
        'Set DRec = opndbaseFMIS.Execute("Select left(Accountcode,3) as childcode,sum(Debit) as sumdebit,sum(credit) as sumcredit From tblAMIS_tmpjournal Where [dvno]='" & txtJEVNo.Text & "' group by Dvno,left(Accountcode,3)")
        Set Drec = opndbaseFMIS.Execute("SELECT  sum([Debit]) as Sumdebit,sum([Credit]) as Sumcredit, [Accountcode] as childcode FROM [fmis].[dbo].[tblAMIS_PostedJEVforCashflow] where [JEVno]='" & txtJEVNo.Text & "'  group by accountcode")
    If Drec.RecordCount > 0 Then
        For x = 1 To Drec.RecordCount
            'MSFlexGrid2.TextMatrix(x, 0) = DRec![Trnno]
            
            MSFlexGrid2.TextMatrix(x, 1) = Drec!childcode
            MSFlexGrid2.TextMatrix(x, 2) = GetAccountNameByAccountcode(Drec!childcode)
            MSFlexGrid2.TextMatrix(x, 4) = IIf((Format(Drec!sumCredit, "#,##0.00") = "0.00"), "", Format(Drec!sumCredit, "#,##0.00"))
            MSFlexGrid2.TextMatrix(x, 3) = IIf((Format(Drec!sumDebit, "#,##0.00") = "0.00"), "", Format(Drec!sumDebit, "#,##0.00"))
            MSFlexGrid2.Rows = MSFlexGrid2.Rows + 1
            'If LCase(Trim(lblMode.Caption)) = "edit" Then MSFlexGrid2.TextMatrix(x, 5) = DRec!ActionCode  ' for coloraly purpose
            Drec.MoveNext
        Next x
    End If
    End If
    Call GetSum2
    Drec.Close
    Set Drec = Nothing
End Sub
Public Sub LoadJEVLogEntry()
Dim Drec As New ADODB.Recordset
Set Drec = opndbaseFMIS.Execute("Exec dbo.[MPproc_LoadLogJEVEntry] @jevno = '" & Trim(txtJEVNo.Text) & "'")
If Drec.RecordCount > 0 Then
Set MSHFlexGrid1.Recordset = Drec
    Call SetMSHGrid(MSHFlexGrid1, 1)
End If
Drec.Close
End Sub
Private Sub SetGrid()
Dim cc As Integer

    MSFlexGrid1.Clear
    MSFlexGrid1.Rows = 2
    MSFlexGrid1.Cols = 6 ' IIf(LCase(Trim(lblMode)) = "edit", 6, 5)
    
    MSFlexGrid1.TextMatrix(0, 1) = "Account Code"
    MSFlexGrid1.TextMatrix(0, 2) = "Accounts and Explanation"
    MSFlexGrid1.TextMatrix(0, 3) = "Debit"
    MSFlexGrid1.TextMatrix(0, 4) = "Credit"
    MSFlexGrid1.TextMatrix(0, 5) = "ActionCode"
    
    MSFlexGrid1.ColWidth(0) = 0
    MSFlexGrid1.ColWidth(1) = 1500
    MSFlexGrid1.ColWidth(2) = 6550
    MSFlexGrid1.ColWidth(3) = 1500
    MSFlexGrid1.ColWidth(4) = 1500
    
    MSFlexGrid1.ColWidth(5) = 0
    
    
    For cc = 0 To MSFlexGrid1.Cols - 1
        MSFlexGrid1.Row = 0
        MSFlexGrid1.col = cc
        MSFlexGrid1.CellAlignment = 4
    Next cc
End Sub

Private Sub SetGrid2()
Dim cc As Integer

    MSFlexGrid2.Clear
    MSFlexGrid2.Rows = 2
    MSFlexGrid2.Cols = 6 ' IIf(LCase(Trim(lblMode)) = "edit", 6, 5)
    
    MSFlexGrid2.TextMatrix(0, 1) = "Account Code"
    MSFlexGrid2.TextMatrix(0, 2) = "Accounts and Explanation"
    MSFlexGrid2.TextMatrix(0, 3) = "Debit"
    MSFlexGrid2.TextMatrix(0, 4) = "Credit"
    
    MSFlexGrid2.ColWidth(0) = 0
    MSFlexGrid2.ColWidth(1) = 1500
    MSFlexGrid2.ColWidth(2) = 6550
    MSFlexGrid2.ColWidth(3) = 1500
    MSFlexGrid2.ColWidth(4) = 1500
    MSFlexGrid2.TextMatrix(0, 5) = "ActionCode"
    'If LCase(Trim(lblMode)) = "Edit" Then
       ' MSFlexGrid2.ColWidth(5) = 1500
    'Else
       MSFlexGrid2.ColWidth(5) = 0
    'End If
    
    
    For cc = 0 To MSFlexGrid2.Cols - 1
        MSFlexGrid2.Row = 0
        MSFlexGrid2.col = cc
        MSFlexGrid2.CellAlignment = 4
    Next cc
End Sub


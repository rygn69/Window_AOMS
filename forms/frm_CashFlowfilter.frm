VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frm_CashFlowfilter 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   9915
   ClientLeft      =   -2115
   ClientTop       =   225
   ClientWidth     =   15375
   Icon            =   "frm_CashFlowfilter.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   2  'Custom
   Picture         =   "frm_CashFlowfilter.frx":076A
   ScaleHeight     =   9915
   ScaleWidth      =   15375
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2535
      Left            =   3840
      ScaleHeight     =   2505
      ScaleWidth      =   11280
      TabIndex        =   44
      Top             =   7320
      Width           =   11310
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
         Height          =   2520
         Left            =   0
         TabIndex        =   45
         Top             =   0
         Width           =   11280
         _ExtentX        =   19897
         _ExtentY        =   4445
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
      Height          =   5415
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   3495
      Begin VB.CommandButton Command1 
         BackColor       =   &H8000000A&
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
         Style           =   1  'Graphical
         TabIndex        =   41
         Top             =   4920
         Width           =   735
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
         Height          =   1575
         Left            =   3960
         TabIndex        =   33
         Top             =   -2880
         Width           =   3255
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
            TabIndex        =   37
            Tag             =   "2"
            Top             =   600
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
            TabIndex        =   36
            Tag             =   "3"
            Top             =   900
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
            TabIndex        =   35
            Tag             =   "4"
            Top             =   1200
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
            TabIndex        =   34
            Tag             =   "1"
            Top             =   360
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
         Height          =   1575
         Left            =   3960
         TabIndex        =   31
         Top             =   -1320
         Width           =   3255
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
            TabIndex        =   39
            Tag             =   "3"
            Top             =   720
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
            TabIndex        =   38
            Tag             =   "1"
            Top             =   360
            Value           =   -1  'True
            Width           =   855
         End
         Begin MSComCtl2.DTPicker DTPYear 
            Height          =   375
            Left            =   1200
            TabIndex        =   32
            Top             =   240
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   661
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "yyyy"
            Format          =   270794755
            UpDown          =   -1  'True
            CurrentDate     =   40651
         End
         Begin MSComCtl2.DTPicker DTpMY 
            Height          =   375
            Left            =   240
            TabIndex        =   40
            Top             =   1080
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   661
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "MMMM yyyy"
            Format          =   270794755
            UpDown          =   -1  'True
            CurrentDate     =   40651
         End
      End
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3960
         Left            =   120
         TabIndex        =   5
         Top             =   840
         Width           =   3255
      End
      Begin VB.TextBox txtcondition 
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
         Left            =   120
         TabIndex        =   4
         ToolTipText     =   "Type JEVno/Dvno/Checkno/PTVno and Press Enter"
         Top             =   360
         Width           =   2415
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
         ItemData        =   "frm_CashFlowfilter.frx":0CF4
         Left            =   5160
         List            =   "frm_CashFlowfilter.frx":0D04
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   -3315
         Width           =   2055
      End
      Begin VB.Label lblcount 
         Alignment       =   1  'Right Justify
         Height          =   255
         Left            =   1680
         TabIndex        =   42
         Top             =   8520
         Width           =   1695
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
         Left            =   3960
         TabIndex        =   2
         Top             =   -3240
         Width           =   1215
      End
   End
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
      Height          =   6135
      Left            =   3720
      TabIndex        =   0
      Top             =   840
      Width           =   11535
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   2895
         Left            =   120
         ScaleHeight     =   2865
         ScaleWidth      =   11280
         TabIndex        =   52
         Top             =   3120
         Width           =   11310
         Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
            Height          =   2880
            Left            =   0
            TabIndex        =   53
            Top             =   0
            Width           =   11280
            _ExtentX        =   19897
            _ExtentY        =   5080
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
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   5640
         TabIndex        =   51
         Top             =   600
         Width           =   5775
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   5640
         TabIndex        =   50
         Top             =   960
         Width           =   5775
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   5640
         TabIndex        =   49
         Top             =   1320
         Width           =   5775
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   5640
         TabIndex        =   48
         Top             =   1680
         Width           =   5775
      End
      Begin VB.TextBox txtjevno 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   1080
         TabIndex        =   46
         Top             =   285
         Width           =   2415
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00FF8080&
         Height          =   375
         Left            =   6480
         MaskColor       =   &H0000FF00&
         Picture         =   "frm_CashFlowfilter.frx":0D25
         Style           =   1  'Graphical
         TabIndex        =   43
         ToolTipText     =   "Click to Generate JEV number"
         Top             =   -480
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtgamount 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   8040
         TabIndex        =   29
         Top             =   2760
         Width           =   3135
      End
      Begin VB.TextBox txtcheckdate 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   9240
         TabIndex        =   26
         Top             =   270
         Width           =   2175
      End
      Begin VB.TextBox txtptvno 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   1080
         TabIndex        =   24
         Top             =   2400
         Width           =   2415
      End
      Begin VB.TextBox txtparticular 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   5640
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   22
         Top             =   2040
         Width           =   5775
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
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5640
         TabIndex        =   18
         Top             =   270
         Width           =   2055
      End
      Begin VB.TextBox txtrdono 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   1080
         TabIndex        =   16
         Top             =   2040
         Width           =   2415
      End
      Begin VB.TextBox txtrcino 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   1080
         TabIndex        =   14
         Top             =   1680
         Width           =   2415
      End
      Begin VB.TextBox txtcheckno 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   1080
         TabIndex        =   12
         Top             =   1320
         Width           =   2415
      End
      Begin VB.TextBox txtobrno 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   1080
         TabIndex        =   10
         Top             =   960
         Width           =   2415
      End
      Begin VB.TextBox txtdvno 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   1080
         TabIndex        =   8
         Top             =   615
         Width           =   2415
      End
      Begin VB.Label Label14 
         Caption         =   "Journal Entry"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   54
         Top             =   2880
         Width           =   1815
      End
      Begin VB.Label Label11 
         Caption         =   "Gross Amount:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6600
         TabIndex        =   30
         Top             =   2760
         Width           =   1455
      End
      Begin VB.Label Label15 
         Caption         =   "Responsibilty Center:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3840
         TabIndex        =   28
         Top             =   1680
         Width           =   2295
      End
      Begin VB.Label sd 
         Caption         =   "Special Account:"
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
         Left            =   4200
         TabIndex        =   27
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Label Label9 
         Caption         =   "Check/ Deposit:"
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
         Left            =   7800
         TabIndex        =   25
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label13 
         Caption         =   "PTV No.:"
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
         Left            =   120
         TabIndex        =   23
         Top             =   2505
         Width           =   1095
      End
      Begin VB.Label Label12 
         Caption         =   "Particular:"
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
         Left            =   4800
         TabIndex        =   21
         Top             =   2160
         Width           =   1095
      End
      Begin VB.Label gh 
         Caption         =   "Object of Expenditure:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3675
         TabIndex        =   20
         Top             =   1395
         Width           =   1935
      End
      Begin VB.Label Label10 
         Caption         =   "Claimant Name:"
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
         Left            =   4320
         TabIndex        =   19
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label Label8 
         Caption         =   "JEV Date.:"
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
         Left            =   4680
         TabIndex        =   17
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label7 
         Caption         =   "RDO No.:"
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
         Left            =   120
         TabIndex        =   15
         Top             =   2070
         Width           =   1095
      End
      Begin VB.Label Label6 
         Caption         =   "RCI No.:"
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
         Left            =   120
         TabIndex        =   13
         Top             =   1710
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "Check No.:"
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
         Left            =   120
         TabIndex        =   11
         Top             =   1395
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "Obr No.:"
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
         Left            =   120
         TabIndex        =   9
         Top             =   1035
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "DVNo.:"
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
         Left            =   120
         TabIndex        =   7
         Top             =   675
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "JEV No.:"
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
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   1095
      End
   End
   Begin MSComctlLib.ImageList itb32x32 
      Left            =   12960
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
            Picture         =   "frm_CashFlowfilter.frx":1067
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_CashFlowfilter.frx":29F9
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_CashFlowfilter.frx":438B
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_CashFlowfilter.frx":5D1D
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_CashFlowfilter.frx":76AF
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_CashFlowfilter.frx":9041
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_CashFlowfilter.frx":A9D3
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_CashFlowfilter.frx":C365
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_CashFlowfilter.frx":DCF7
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_CashFlowfilter.frx":F68B
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_CashFlowfilter.frx":10367
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_CashFlowfilter.frx":10C47
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_CashFlowfilter.frx":11923
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_CashFlowfilter.frx":125FF
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_CashFlowfilter.frx":132DB
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_CashFlowfilter.frx":13FB7
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_CashFlowfilter.frx":14C93
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   840
      Left            =   0
      TabIndex        =   55
      Top             =   0
      Width           =   15375
      _ExtentX        =   27120
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
   End
   Begin VB.Label Label16 
      Caption         =   "Cash Flow entry"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3840
      TabIndex        =   47
      Top             =   7080
      Width           =   1815
   End
End
Attribute VB_Name = "frm_CashFlowfilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim isEdit, IfNew As Boolean
Dim IfEdit As String
Public ClaimantCode, RCcode, FmisVoucherno, ref As String
Dim Jevseries As Long
Dim trans As Integer
Dim xDebit As Currency
Dim xCredit As Currency
Dim ifColoraly As Boolean

Dim ifsaveamount As Boolean
Dim SaveOk As Boolean
Public Ttype As Integer
Public fundcode As Long
Public FundType As String
Public EditCount, IsSaveAccntng As Boolean

Dim not_coloraly_total_debit, not_coloraly_total_credit, coloraly_total_debit, coloraly_total_credit As Double
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
            txtcheckno.Text = ""
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
Private Sub cmb_Field_Change()
Call LoadFinalTrans(cmb_Field.Text, txtcondition.Text)
End Sub

Private Sub cmb_Field_Click()
Call LoadFinalTrans(cmb_Field.Text, txtcondition.Text)
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

Private Sub Command4_Click()
Dim rec As New ADODB.Recordset
Dim Lastno As Double
Dim TYP As Integer
    If Option7.Value = False Then
        MsgBox "Please Specify the date of your Post date..", vbInformation, "System Message"
        Exit Sub
    End If
    
        If Trim(txtfundtype.Text) <> "" Then
        If Option2.Value = True Then: TYP = Option2.Tag
        If Option3.Value = True Then: TYP = Option3.Tag
        If Option4.Value = True Then: TYP = Option4.Tag
        If Option5.Value = True Then: TYP = Option5.Tag
            rec.Open ("EXEC [dbo].[Proc_GetMaxJevSeries_New] @transtype = " & TYP & ",@jevyeardate = '" & Year(DTpMY) & "' ,@fundtype = '" & txtfundtype.Text & "'"), opndbaseFMIS, adOpenStatic, adLockOptimistic
                Lastno = rec.Fields!MAXJEVSERIES
            rec.Close
            txtJEVNo.Text = txtfundtype.ItemData(txtfundtype.ListIndex) & "-" & Right(Year(DTpMY), 2) & "-" & Format(Month(DTpMY), "00") & "-" & Format(TYP, "00") & "-" & Format(Lastno, "0000")
            txtjevdate.Text = Format(DTpMY, "MMMM yyyy")
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
txt ("Disable")
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
Dim PRec As New ADODB.Recordset
Dim x As Integer
Dim Transtype As String
Dim Postdate As String
Dim orderby As String
Transtype = ""
If Option2.Value = True Then: Transtype = Option2.Tag
If Option3.Value = True Then: Transtype = Option3.Tag
If Option4.Value = True Then: Transtype = Option4.Tag
If Option5.Value = True Then: Transtype = Option5.Tag

Postdate = ""
If Option1.Value = True Then: Postdate = "year(jevdate) = " & DTPYear.Year & ""
If Option7.Value = True Then: Postdate = "year(jevdate) = " & DTpMY.Year & " and month(jevdate) = " & DTpMY.Month & ""
orderby = field
If field = "JEVno" Then: orderby = "substring(Jevno,14,7)"

    List1.Clear
    PRec.Open ("Select " & field & " as field From tblAMIS_FinalJEV Where " & field & " like '" & Condition & "%' and " & Postdate & " and transtype in (" & Transtype & ") and ltrim(" & field & ") <> '' and Actioncode=1  Group By " & field & " order by " & orderby & ""), opndbaseFMIS, adOpenStatic, adLockOptimistic
    If PRec.RecordCount > 0 Then
    lblcount.Caption = PRec.RecordCount & " Record found"
        For x = 1 To PRec.RecordCount
            List1.AddItem PRec!field
            PRec.MoveNext
            DoEvents
        Next x
        List1.Enabled = True
    End If
    PRec.Close
    Set PRec = Nothing
End Function

Private Sub List1_Click()
isEdit = True
IsSaveAccntng = False
Call Loaddetails(cmb_Field.Text, List1.Text)
End Sub
Private Function Loaddetails(ByVal field As String, ByVal Condition As String)
Dim rec As New ADODB.Recordset
On Error Resume Next
rec.Open "Select * from tblAMIS_FinalJEV where " & field & " = '" & Condition & "' and  actioncode = 1", opndbaseFMIS, adOpenStatic, adLockOptimistic
    If rec.RecordCount <> 0 Then
       ' Clrtxt
        txt ("Clear")
        txtcheckdate.Text = Trim(rec.Fields!Date_)
        txtcheckno.Text = Trim(rec.Fields!checkno)
        txtClaimant.Text = getClaimant(rec.Fields!ClaimantCode)
        ClaimantCode = Trim(rec.Fields!ClaimantCode)
        txtDVNo.Text = Trim(rec.Fields!dvno)
        txtfundtype = Trim(rec.Fields!FundType)
        txtjevdate.Text = Format(Trim(rec.Fields!jevdate), "MMMM yyyy")
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
        ref = IIf(IsNull(rec.Fields!RefNo), "", (rec.Fields!RefNo))
        Jevseries = rec.Fields!jevseriesno
            Call GetAccntngEntries
            Call GetSum
    End If
        
rec.Close
Set rec = Nothing
End Function
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
        txtcheckno.Text = ""
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

Private Sub MSFlexGrid1_DblClick()
    With frmSub3
        .reff = txtJEVNo.Text
        .Gamount = txtgamount.Text
        .CName = UCase(txtClaimant.Text)
        .isEdit = True
        Set .frm = Me
        Call LoadAcctngEntries(txtJEVNo.Text)
        .Show 1
        Call GetAccntngEntries
    End With
End Sub
Public Function LoadAcctngEntries(ByVal jevno As String)
Dim DRec As New ADODB.Recordset
Dim rec As New ADODB.Recordset
Dim x As Integer
    Set DRec = opndbaseFMIS.Execute("Select Accountcode,Debit ,Credit From tblAMIS_postedJEV Where [JEVno]='" & jevno & "' And (ActionCode=1) ")
    If DRec.RecordCount > 0 Then
        If EditCount = False Then
        EditCount = True
            rec.Open "Select dvno from tblAMIs_tmpjournal where dvno = '" & jevno & "'", opndbaseFMIS, adOpenStatic
            If rec.RecordCount > 0 Then
                    If MsgBox("This transaction Have a temporary Accounting Entries, do you want to Delete?", vbCritical + vbYesNo, "System Information") = vbYes Then
                        opndbaseFMIS.Execute "Delete from tblAMIs_tmpjournal where dvno = '" & txtJEVNo.Text & "'"
                        For x = 1 To DRec.RecordCount
                        DoEvents
                            opndbaseFMIS.Execute "Insert into tblAMIs_tmpjournal (Dvno,Accountcode,Debit,Credit) values ('" & Trim(jevno) & "','" & Trim(DRec!accountcode) & "'," & DRec!Debit & "," & DRec!Credit & ")"
                            DRec.MoveNext
                        Next x
                    End If
            Else
            For x = 1 To DRec.RecordCount
                        DoEvents
                            opndbaseFMIS.Execute "Insert into tblAMIs_tmpjournal (Dvno,Accountcode,Debit,Credit) values ('" & Trim(jevno) & "','" & Trim(DRec!accountcode) & "'," & DRec!Debit & "," & DRec!Credit & ")"
                            DRec.MoveNext
                        Next x
            End If
            rec.Close
        Else
        For x = 1 To DRec.RecordCount
                        DoEvents
                            opndbaseFMIS.Execute "Insert into tblAMIs_tmpjournal (Dvno,Accountcode,Debit,Credit) values ('" & Trim(jevno) & "','" & Trim(DRec!accountcode) & "'," & DRec!Debit & "," & DRec!Credit & ")"
                            DRec.MoveNext
                        Next x
            
        End If
    End If
    DRec.Close
    Set DRec = Nothing
End Function

Private Sub cmbEntry_KeyPress(KeyAscii As Integer)
    
    
    If KeyAscii = 13 Then
        If cmbEntry.ListIndex <> -1 Then
            MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 1) = cmbEntry.Text
            MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 0) = cmbEntry.ItemData(cmbEntry.ListIndex)
            MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 5) = "1"
            If MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2) = "TOTAL" Then
                MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 3) = ""
                MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 4) = ""
            End If
            MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2) = GetAccountNameByFMISAccountCode(cmbEntry.ItemData(cmbEntry.ListIndex))
           ' SET MSFlexGrid1.SetFocus(MSFlexGrid1.Row, 3)
        ElseIf cmbEntry.Text = "" Then
            MSFlexGrid1.RemoveItem (MSFlexGrid1.Row)

        End If
        cmbEntry.Visible = False
        Call GetSum
    MSFlexGrid1.SetFocus
    Else
       ' KeyAscii = AutoFind(cmbEntry, KeyAscii, True)
    End If
    
End Sub

Private Function Load()
If Option1.Value = True Then
    DTPYear.Enabled = True
    DTpMY.Enabled = False
Else
DTPYear.Enabled = False
    DTpMY.Enabled = True
End If

End Function

Private Sub Option1_Click()
Load
End Sub

Private Sub Option7_Click()
Load
End Sub

Public Function SaveFinalJEV()
Dim Credit As Currency
Dim Debit As Currency
Dim tmp As Long
If Option2.Value = True Then: trans = Option2.Tag
If Option3.Value = True Then: trans = Option3.Tag
If Option4.Value = True Then: trans = Option4.Tag
If Option5.Value = True Then: trans = Option5.Tag
Jevseries = ExtractJEVSNo(txtJEVNo)

If MsgBox("Are You Sure Do you want to Save/Update these entry?", vbInformation + vbYesNo) = vbYes Then
    If isEdit = True Then
        opndbaseFMIS.Execute "delete from tblAMIS_FinalJEV where jevno = '" & txtJEVNo.Text & "'"
    Else
    
        If CheckIfExistInFinalJEV(txtJEVNo.Text) = True Then
                If MsgBox("JEV Number Already exist on the database,Do you want the system Generate the JEV number?" & vbNewLine & "Click YES to generate,NO to cancel..!", vbInformation + vbYesNo, "System Message") = vbYes Then
                Call Command4_Click
                Else
                Exit Function
                End If
        End If
    End If
                
                Call Saved2FinalJEV(txtcheckdate.Text, Trim(txtrcino.Text), txtcheckno.Text, txtParticular.Text, txtJEVNo.Text, ClaimantCode, 0, txtgamount.Text, 0, 0, trans, FmisVoucherno, txtDVNo.Text, txtobrno.Text, txtfundtype.Text, RCcode, txtOOE.Text, txtrdono.Text, ref, Jevseries, txtjevdate.Text, txtptvno.Text)
                If IsSaveAccntng = True Then
                    Call SaveAcctngEntries(txtJEVNo.Text)
                End If
                Call LoadFinalTrans(cmb_Field.Text, txtcondition.Text)
                
End If
End Function
Public Function SaveAcctngEntries(ByVal jevno As String)
Dim DRec As New ADODB.Recordset
Dim rec As New ADODB.Recordset
Dim x As Integer

    DRec.Open ("Select Accountcode,sum(Debit) as Debit ,sum(Credit) as Credit From tblAMIs_tmpjournal Where [dvno]='" & jevno & "' group by accountcode"), opndbaseFMIS, adOpenStatic, adLockOptimistic
    If DRec.RecordCount > 0 Then
        opndbaseFMIS.Execute "update tblAMIS_postedJEV set actioncode =2 where JEVNO = '" & jevno & "' and actioncode =1" ', datetimeentered = rtrim(ltrim(DateTimeEntered)) +'," & Format(Now, "yyyy/mm/dd hh:mm:ss AMPM") & "',userid = rtrim(ltrim(UserID)) + '," & Trim(ActiveUserID) & "'
        For x = 1 To DRec.RecordCount
            DoEvents
            opndbaseFMIS.Execute "Insert into tblAMIS_PostedJEV (JEVNO,Accountcode,debit,credit,actioncode,datetimeentered,userid) values " & _
            "('" & Trim(jevno) & "','" & Trim(DRec!accountcode) & "'," & DRec!Debit & "," & DRec!Credit & ",1,'" & Format(Now, "yyyy/mm/dd hh:mm:ss AMPM") & "','" & Trim(ActiveUserID) & "')"
            DRec.MoveNext
        Next x
        opndbaseFMIS.Execute "delete from tblAMIs_tmpjournal where dvno = '" & jevno & "'"
    End If
    DRec.Close
    Set DRec = Nothing
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
If isEdit = False Then
Call LoadAccountsByFund(txtfundtype.Text)
End If
End Sub
Private Sub txtfundtype_Click()
If isEdit = False Then
Call LoadAccountsByFund(txtfundtype.Text)
End If
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
Dim DRec As New ADODB.Recordset
Dim x As Integer
Call SetGrid
    'DRec.Close
    If IsSaveAccntng = False Then
        Set DRec = opndbaseFMIS.Execute("Select left(Accountcode,3) as childcode,sum(Debit) as sumdebit,sum(credit) as sumcredit From tblAMIS_POSTEDJEV Where [JEVno]='" & txtJEVNo.Text & "' And (ActionCode=1) group by jevno,actioncode,left(Accountcode,3)")
        If DRec.RecordCount > 0 Then
            For x = 1 To DRec.RecordCount
    '            MSFlexGrid1.TextMatrix(x, 0) = DRec![Trnno]
                MSFlexGrid1.TextMatrix(x, 1) = DRec!childcode
                MSFlexGrid1.TextMatrix(x, 2) = GetAccountNameByAccountcode(DRec!childcode)
                MSFlexGrid1.TextMatrix(x, 4) = IIf((Format(DRec!sumCredit, "#,##0.00") = "0.00"), "", Format(DRec!sumCredit, "#,##0.00"))
                MSFlexGrid1.TextMatrix(x, 3) = IIf((Format(DRec!sumDebit, "#,##0.00") = "0.00"), "", Format(DRec!sumDebit, "#,##0.00"))
              MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1
               ' If LCase(Trim(lblMode.Caption)) = "edit" Then MSFlexGrid1.TextMatrix(x, 5) = DRec!ActionCode  ' for coloraly purpose
                DRec.MoveNext
            Next x
            
        End If
    Else
        Set DRec = opndbaseFMIS.Execute("Select left(Accountcode,3) as childcode,sum(Debit) as sumdebit,sum(credit) as sumcredit From tblAMIS_tmpjournal Where [dvno]='" & txtJEVNo.Text & "' group by Dvno,left(Accountcode,3)")
    If DRec.RecordCount > 0 Then
        For x = 1 To DRec.RecordCount
            'MSFlexGrid1.TextMatrix(x, 0) = DRec![Trnno]
            
            MSFlexGrid1.TextMatrix(x, 1) = DRec!childcode
            MSFlexGrid1.TextMatrix(x, 2) = GetAccountNameByAccountcode(DRec!childcode)
            MSFlexGrid1.TextMatrix(x, 4) = IIf((Format(DRec!sumCredit, "#,##0.00") = "0.00"), "", Format(DRec!sumCredit, "#,##0.00"))
            MSFlexGrid1.TextMatrix(x, 3) = IIf((Format(DRec!sumDebit, "#,##0.00") = "0.00"), "", Format(DRec!sumDebit, "#,##0.00"))
            MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1
            'If LCase(Trim(lblMode.Caption)) = "edit" Then MSFlexGrid1.TextMatrix(x, 5) = DRec!ActionCode  ' for coloraly purpose
            DRec.MoveNext
        Next x
    End If
    End If
    Call GetSum
    DRec.Close
    Set DRec = Nothing
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


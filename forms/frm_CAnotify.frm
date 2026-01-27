VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MShflxgd.ocx"
Begin VB.Form frm_CAnotify 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Un-liquidated cash advance notification"
   ClientHeight    =   9180
   ClientLeft      =   5055
   ClientTop       =   4980
   ClientWidth     =   15720
   Icon            =   "frm_CAnotify.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9180
   ScaleWidth      =   15720
   Begin TabDlg.SSTab SSTab1 
      Height          =   9015
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   15495
      _ExtentX        =   27331
      _ExtentY        =   15901
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Cash advance Mangement"
      TabPicture(0)   =   "frm_CAnotify.frx":076A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label14"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label12"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label11"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label3"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label2"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label7"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label8"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label9"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label5"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label4"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label6"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label10"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "lblCAlist"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label19"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "lbl_ready"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "lblsent"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "grid_Ready"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "grid_sent"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "grid_Verify"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "txtEmail"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Command6"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "Command5"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "Command4"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "txtDaysPass"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "Command2"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "Command3"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "Command1"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "txtCheckno"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "txtcheckdate"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "txtDetail"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "txtOffice"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "txtClaimant"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "txtAmount"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "txtDVNo"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "txtObR"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "Command8"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).ControlCount=   37
      TabCaption(1)   =   "List of Liquidated Cash Advanced"
      TabPicture(1)   =   "frm_CAnotify.frx":0786
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "Settings"
      TabPicture(2)   =   "frm_CAnotify.frx":07A2
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame1"
      Tab(2).ControlCount=   1
      Begin VB.Frame Frame1 
         Height          =   4095
         Left            =   -74760
         TabIndex        =   35
         Top             =   600
         Width           =   6495
         Begin VB.CheckBox Check1 
            Caption         =   "Test"
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
            Left            =   1680
            TabIndex        =   51
            Top             =   4320
            Width           =   1335
         End
         Begin VB.CheckBox chk_Disable 
            Caption         =   "Disable auto email"
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
            Left            =   1680
            TabIndex        =   41
            Top             =   240
            Width           =   2175
         End
         Begin VB.CommandButton Command7 
            Caption         =   "Apply"
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
            Left            =   5400
            TabIndex        =   40
            Top             =   240
            Width           =   735
         End
         Begin VB.TextBox txtSemail 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
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
            Left            =   1680
            TabIndex        =   39
            TabStop         =   0   'False
            Top             =   1320
            Width           =   3510
         End
         Begin VB.TextBox txtSname 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   360
            Left            =   1680
            TabIndex        =   38
            Top             =   840
            Width           =   3495
         End
         Begin VB.TextBox txtSubject 
            Appearance      =   0  'Flat
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
            Left            =   1680
            TabIndex        =   37
            Top             =   1800
            Width           =   4440
         End
         Begin VB.TextBox txtMessage 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1155
            Left            =   1680
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   36
            Top             =   2280
            Width           =   4680
         End
         Begin VB.Frame Frame2 
            Caption         =   "Frame2"
            Height          =   1335
            Left            =   1680
            TabIndex        =   52
            Top             =   5400
            Width           =   4455
            Begin VB.TextBox Text1 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
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
               Left            =   240
               TabIndex        =   53
               TabStop         =   0   'False
               Top             =   720
               Width           =   3990
            End
            Begin VB.Label Label21 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Send to:"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   240
               Left            =   195
               TabIndex        =   54
               Top             =   360
               Width           =   735
            End
         End
         Begin VB.Label Label15 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Subject:"
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
            Left            =   765
            TabIndex        =   45
            Top             =   1800
            Width           =   720
         End
         Begin VB.Label Label16 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Message with  Parameter:"
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
            Left            =   480
            TabIndex        =   44
            Top             =   2280
            Width           =   1050
         End
         Begin VB.Label Label17 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Name:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   1035
            TabIndex        =   43
            Top             =   945
            Width           =   570
         End
         Begin VB.Label Label18 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Email Address:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   240
            TabIndex        =   42
            Top             =   1320
            Width           =   1290
         End
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Open Command"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   10080
         TabIndex        =   34
         Top             =   960
         Width           =   855
      End
      Begin VB.TextBox txtObR 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Left            =   6480
         Locked          =   -1  'True
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   1440
         Width           =   3510
      End
      Begin VB.TextBox txtDVNo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   6480
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   960
         Width           =   3495
      End
      Begin VB.TextBox txtAmount 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   6480
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   5160
         Width           =   1800
      End
      Begin VB.TextBox txtClaimant 
         Appearance      =   0  'Flat
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
         Left            =   6480
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   1920
         Width           =   4440
      End
      Begin VB.TextBox txtOffice 
         Appearance      =   0  'Flat
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
         Left            =   6480
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   2400
         Width           =   4440
      End
      Begin VB.TextBox txtDetail 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1155
         Left            =   6480
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   14
         Top             =   2880
         Width           =   4680
      End
      Begin VB.TextBox txtcheckdate 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   9360
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   4680
         Width           =   1560
      End
      Begin VB.TextBox txtCheckno 
         Appearance      =   0  'Flat
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
         Left            =   6480
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   4680
         Width           =   1800
      End
      Begin VB.CommandButton Command1 
         Caption         =   "<"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   17.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   7440
         TabIndex        =   9
         Top             =   480
         Width           =   735
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Clear"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   8160
         TabIndex        =   8
         Top             =   480
         Width           =   735
      End
      Begin VB.CommandButton Command2 
         Caption         =   ">"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   17.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   8880
         TabIndex        =   7
         Top             =   480
         Width           =   735
      End
      Begin VB.TextBox txtDaysPass 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Left            =   9360
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   5160
         Width           =   1560
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Refresh"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   14400
         TabIndex        =   5
         Top             =   600
         Width           =   975
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Refresh"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3360
         TabIndex        =   4
         Top             =   600
         Width           =   975
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Refresh"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   14400
         TabIndex        =   3
         Top             =   5640
         Width           =   975
      End
      Begin VB.TextBox txtEmail 
         Appearance      =   0  'Flat
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
         Left            =   6480
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   4200
         Width           =   4440
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grid_Verify 
         Height          =   4575
         Left            =   120
         TabIndex        =   10
         Top             =   960
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   8070
         _Version        =   393216
         FixedCols       =   0
         ScrollTrack     =   -1  'True
         AllowUserResizing=   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grid_sent 
         Height          =   2895
         Left            =   120
         TabIndex        =   11
         Top             =   6000
         Width           =   15255
         _ExtentX        =   26908
         _ExtentY        =   5106
         _Version        =   393216
         FixedCols       =   0
         ScrollTrack     =   -1  'True
         AllowUserResizing=   1
         Appearance      =   0
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grid_Ready 
         Height          =   4575
         Left            =   11160
         TabIndex        =   20
         Top             =   1080
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   8070
         _Version        =   393216
         FixedCols       =   0
         ScrollTrack     =   -1  'True
         AllowUserResizing=   1
         Appearance      =   0
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.Label lblsent 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   1560
         TabIndex        =   50
         Top             =   5640
         Width           =   420
      End
      Begin VB.Label lbl_ready 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   13080
         TabIndex        =   49
         Top             =   600
         Width           =   420
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   0
         TabIndex        =   47
         Top             =   0
         Width           =   60
      End
      Begin VB.Label lblCAlist 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   1920
         TabIndex        =   46
         Top             =   600
         Width           =   420
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   5040
         TabIndex        =   33
         Top             =   5160
         Width           =   1380
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "OBR No."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   4425
         TabIndex        =   32
         Top             =   1440
         Width           =   1935
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Claimant name:"
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
         Left            =   4320
         TabIndex        =   31
         Top             =   1920
         Width           =   1965
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
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
         Left            =   4320
         TabIndex        =   30
         Top             =   2400
         Width           =   2025
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Particular:"
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
         Left            =   5400
         TabIndex        =   29
         Top             =   2880
         Width           =   885
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Checkdate:"
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
         Left            =   8325
         TabIndex        =   28
         Top             =   4680
         Width           =   960
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Checkno:"
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
         Left            =   5505
         TabIndex        =   27
         Top             =   4680
         Width           =   795
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DVNO:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   4620
         TabIndex        =   26
         Top             =   1065
         Width           =   1785
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ready for Sending.."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   11280
         TabIndex        =   25
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cash advance list."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   120
         TabIndex        =   24
         Top             =   600
         Width           =   1545
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sent email list"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   120
         TabIndex        =   23
         Top             =   5640
         Width           =   1455
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Days Pass:"
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
         Left            =   8400
         TabIndex        =   22
         Top             =   5160
         Width           =   930
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Email Address:"
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
         Left            =   4995
         TabIndex        =   21
         Top             =   4200
         Width           =   1290
      End
   End
   Begin VB.Timer Tdoevents 
      Interval        =   1000
      Left            =   0
      Top             =   0
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   0
      TabIndex        =   48
      Top             =   0
      Width           =   420
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Year of:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   11400
      TabIndex        =   0
      Top             =   -120
      Width           =   555
   End
   Begin VB.Menu command 
      Caption         =   "Command"
      Begin VB.Menu notCA 
         Caption         =   "This is not a Cash Advance"
      End
      Begin VB.Menu AL 
         Caption         =   "Already Liquidated"
      End
   End
End
Attribute VB_Name = "frm_CAnotify"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public IsEditParticular As Boolean
Private Sub LoadGrid(ByVal grd As MSHFlexGrid, sqlsquery As String)
Dim PRec As New ADODB.Recordset
Dim x As Integer
grd.Clear
Set PRec = opndbaseFMIS.Execute(sqlsquery)
If PRec.RecordCount > 0 Then
    Set grd.DataSource = PRec
    Call SetMSHGrid(grd, 7)
    If grd.name = "grid_Verify" Then
        lblCAlist.Caption = (grd.Rows - 1) & " Record(s) found"
    ElseIf grd.name = "grid_Ready" Then
         lbl_ready.Caption = (grd.Rows - 1) & " Record(s) found"
    ElseIf grd.name = "grid_sent" Then
         lblsent.Caption = (grd.Rows - 1) & " Record(s) found"
    End If
Else
    If grd.name = "grid_Verify" Then
        lblCAlist.Caption = ""
    ElseIf grd.name = "grid_Ready" Then
         lbl_ready.Caption = ""
    ElseIf grd.name = "grid_sent" Then
         lblsent.Caption = ""
    End If
End If
PRec.Close
Set PRec = Nothing
End Sub

Private Sub AL_Click()
IsOktoClear = False
frm_CAAndLiquidation.dvno = txtDVNo.Text
frm_CAAndLiquidation.Show 1
If IsOktoClear = True Then
    Call LoadGrid(grid_Verify, "EXECUTE  [fmis].[dbo].[MPproc_RPT_Query]    @from = '1/31/2013',@to = '1/31/2013',@reports='ListVerifyCA'")
    Call clr
End If
End Sub

Private Sub Command1_Click()
If txtObR.Text = "" Then
    MsgBox "Invalid Execution..", vbInformation, "System Information"
    Exit Sub
End If
    If MsgBox("Are you sure do you want to remove in list of ready to email?", vbInformation + vbYesNo, "System Confirmation") = vbYes Then
    opndbaseFMIS.Execute "delete from [fmis].[dbo].[tblAMIS_CashAdvanceNotify] where [Obrno]='" & txtObR.Text & "'"
    Call LoadGrid(grid_Ready, "EXECUTE  [fmis].[dbo].[MPproc_RPT_Query]    @from = '1/31/2013',@to = '1/31/2013',@reports='ListCAforSEND'")
    Call LoadGrid(grid_Verify, "EXECUTE  [fmis].[dbo].[MPproc_RPT_Query]    @from = '1/31/2013',@to = '1/31/2013',@reports='ListVerifyCA'")
    Call clr
End If
End Sub

Private Sub Command2_Click()
If txtObR.Text = "" Then
    MsgBox "No details to add in the list.", vbInformation, "System Information"
    Exit Sub
End If
If MsgBox("Are you sure do you want to add in the List for sending email?", vbInformation + vbYesNo, "System Confirmation") = vbYes Then
    opndbaseFMIS.Execute "insert into [fmis].[dbo].[tblAMIS_CashAdvanceNotify]([Dvno],[Obrno],[Name],[OfficeName],[Checkno],[Checkdate],[Amount],[Particular],[Email],[UserVerify],DateTimeVerify,[IsSent]) values " & _
    "('" & txtDVNo.Text & "','" & txtObR.Text & "','" & txtClaimant.Text & "','" & Replace(txtOffice.Text, "'", "''") & "','" & txtCheckno.Text & "','" & txtcheckdate.Text & "','" & txtAmount.Text & "','" & Replace(txtDetail.Text, "'", "''") & "','" & txtEmail.Text & "','" & ActiveUserID & "','" & Now & "',0)"
    Call LoadGrid(grid_Ready, "EXECUTE  [fmis].[dbo].[MPproc_RPT_Query]    @from = '1/31/2013',@to = '1/31/2013',@reports='ListCAforSEND'")
    Call LoadGrid(grid_Verify, "EXECUTE  [fmis].[dbo].[MPproc_RPT_Query]    @from = '1/31/2013',@to = '1/31/2013',@reports='ListVerifyCA'")
    If IsEditParticular = True Then
        opndbaseFMIS.Execute "update [fmis].[dbo].[tblAMIS_IncomingDVTrns] set [Particular] = '" & txtDetail.Text & "' where dvno = '" & txtDVNo.Text & "' and actioncode = 1"
        IsEditParticular = False
    End If
    Call clr
End If
End Sub
Private Sub Command3_Click()
Call clr
Call Loaddetails
End Sub
Private Sub clr()
txtAmount.Text = ""
txtClaimant.Text = ""
txtDetail.Text = ""
txtDVNo.Text = ""
txtCheckno.Text = ""
txtObR.Text = ""
txtOffice.Text = ""
txtcheckdate.Text = ""
txtAmount.Text = ""
txtDaysPass.Text = ""
txtEmail.Text = ""
End Sub
Private Sub Command4_Click()
Call LoadGrid(grid_Ready, "EXECUTE  [fmis].[dbo].[MPproc_RPT_Query]    @from = '1/31/2013',@to = '1/31/2013',@reports='ListCAforSEND'")
End Sub

Private Sub Command5_Click()
Call LoadGrid(grid_Verify, "EXECUTE  [fmis].[dbo].[MPproc_RPT_Query]    @from = '1/31/2013',@to = '1/31/2013',@reports='ListVerifyCA'")
End Sub

Private Sub Command6_Click()
Call LoadGrid(grid_sent, "EXECUTE  [fmis].[dbo].[MPproc_RPT_Query]    @from = '1/31/2013',@to = '1/31/2013',@reports='LisCASENT'")
End Sub

Private Sub Command7_Click()
Dim disable As Integer
If MsgBox("Are you sure do you want to Update the entry?", vbInformation + vbYesNo) = vbYes Then
    disable = IIf((chk_Disable.Value = 1), "2", 1)
    opndbaseFMIS.Execute "Update [dbo].[tblAMIS_FMISEmailSender] Set [IsStopSending] = " & disable & ",[EmailAddress] = '" & txtSemail.Text & "'" & _
    ",[Subject] = '" & txtSubject.Text & "',Message = '" & txtMessage.Text & "',userid = '" & ActiveUserID & "',[datetimeEntered] = '" & Now & "' where [ForWhat] = 'CANotify' and IsActive = 1"
End If
End Sub

Private Sub Command8_Click()
PopupMenu command
End Sub

Private Sub Form_Load()
Call Loaddetails
Call LoadSenderDetails
End Sub
Private Sub Loaddetails()
Call LoadGrid(grid_Verify, "EXECUTE  [fmis].[dbo].[MPproc_RPT_Query]    @from = '1/31/2013',@to = '1/31/2013',@reports='ListVerifyCA'")
Call LoadGrid(grid_Ready, "EXECUTE  [fmis].[dbo].[MPproc_RPT_Query]    @from = '1/31/2013',@to = '1/31/2013',@reports='ListCAforSEND'")
Call LoadGrid(grid_sent, "EXECUTE  [fmis].[dbo].[MPproc_RPT_Query]    @from = '1/31/2013',@to = '1/31/2013',@reports='LisCASENT'")
End Sub

Private Sub grid_Ready_Click()
Call SelectReady
End Sub

Private Sub SelectReady()
On Error Resume Next
If grid_Ready.TextMatrix(grid_Ready.Row, 1) <> "" Then
    txtObR.Text = grid_Ready.TextMatrix(grid_Ready.Row, 3)
    txtAmount.Text = grid_Ready.TextMatrix(grid_Ready.Row, 8)
    txtClaimant.Text = grid_Ready.TextMatrix(grid_Ready.Row, 0)
    txtDetail.Text = grid_Ready.TextMatrix(grid_Ready.Row, 1)
    txtDVNo.Text = grid_Ready.TextMatrix(grid_Ready.Row, 2)
    If Trim(txtDVNo.Text) <> "" Then
        txtEmail.Text = GetExcuteScalar(1, grid_Ready.TextMatrix(grid_Ready.Row, 2))
    End If
    txtCheckno.Text = grid_Ready.TextMatrix(grid_Ready.Row, 5)
    txtOffice.Text = grid_Ready.TextMatrix(grid_Ready.Row, 4)
    txtcheckdate.Text = grid_Ready.TextMatrix(grid_Ready.Row, 6)
    txtDaysPass.Text = grid_Ready.TextMatrix(grid_Ready.Row, 7)
    Command2.Enabled = False
    Command1.Enabled = True
End If
End Sub

Private Sub grid_sent_DblClick()
Call SelectSent
End Sub

Private Sub grid_Verify_Click()
Call SelectVerify
End Sub

Private Sub grid_Verify_KeyDown(KeyCode As Integer, Shift As Integer)
Call SelectVerify
End Sub
Private Sub SelectVerify()
txtObR.Text = grid_Verify.TextMatrix(grid_Verify.Row, 3)
txtAmount.Text = grid_Verify.TextMatrix(grid_Verify.Row, 8)
txtClaimant.Text = grid_Verify.TextMatrix(grid_Verify.Row, 0)
txtDetail.Text = grid_Verify.TextMatrix(grid_Verify.Row, 1)
txtDVNo.Text = grid_Verify.TextMatrix(grid_Verify.Row, 2)
If Trim(txtDVNo.Text) <> "" Then
txtEmail.Text = GetExcuteScalar(1, grid_Verify.TextMatrix(grid_Verify.Row, 2))
End If
txtCheckno.Text = grid_Verify.TextMatrix(grid_Verify.Row, 5)
txtOffice.Text = grid_Verify.TextMatrix(grid_Verify.Row, 4)
txtcheckdate.Text = grid_Verify.TextMatrix(grid_Verify.Row, 6)
txtDaysPass.Text = grid_Verify.TextMatrix(grid_Verify.Row, 7)
Command2.Enabled = True
Command1.Enabled = False
End Sub
Private Sub SelectSent()
With frm_CAChildDetails
    .frmOk = False
    .txtObR.Text = grid_sent.TextMatrix(grid_sent.Row, 3)
    .txtAmount.Text = grid_sent.TextMatrix(grid_sent.Row, 8)
    .txtClaimant.Text = grid_sent.TextMatrix(grid_sent.Row, 0)
    .txtDetail.Text = grid_sent.TextMatrix(grid_sent.Row, 1)
    .txtDVNo.Text = grid_sent.TextMatrix(grid_sent.Row, 2)
    If Trim(txtDVNo.Text) <> "" Then
    .txtEmail.Text = GetExcuteScalar(1, grid_sent.TextMatrix(grid_sent.Row, 2))
    End If
    .txtCheckno.Text = grid_sent.TextMatrix(grid_sent.Row, 5)
    .txtOffice.Text = grid_sent.TextMatrix(grid_sent.Row, 4)
    .txtcheckdate.Text = grid_sent.TextMatrix(grid_sent.Row, 6)
    .txtDaysPass.Text = grid_sent.TextMatrix(grid_sent.Row, 7)
    .txtLastNotify.Text = GetExcuteScalar(2, grid_sent.TextMatrix(grid_sent.Row, 2))
    .Show 1
    If .frmOk = True Then
        Call Loaddetails
    End If
End With


End Sub
Private Sub LoadSenderDetails()
Dim rec As New ADODB.Recordset
Set rec = opndbaseFMIS.Execute("SELECT [EmailAddress],[Subject],[Message],[IsStopSending] FROM [fmis].[dbo].[tblAMIS_FMISEmailSender] where IsActive = 1")
If rec.RecordCount > 0 Then
    txtSemail.Text = rec!EmailAddress
    txtSname.Text = "Cash Advance Notifier"
    txtSubject.Text = rec!Subject
    txtMessage.Text = rec!Message
    
    If rec!IsStopSending = 1 Then
        chk_Disable.Value = 0
    Else
        chk_Disable.Value = 1
    End If
End If
End Sub

Private Sub notCA_Click()
If txtDVNo.Text <> "" Then
    If MsgBox("Are you sure this transaction is not a cash advance?", vbInformation + vbYesNo, "System Confirmation") = vbYes Then
    opndbaseFMIS.Execute "Update dbo.tblAMIS_IncomingDVTrns set isca = 0 where dvno = '" & txtDVNo.Text & "' and actioncode = 1"
    Call LoadGrid(grid_Verify, "EXECUTE  [fmis].[dbo].[MPproc_RPT_Query]    @from = '1/1/2017',@to = '1/31/2017',@reports='ListVerifyCA'")
    Call clr
    End If
End If
End Sub

Private Sub txtDetail_Change()
IsEditParticular = True
End Sub

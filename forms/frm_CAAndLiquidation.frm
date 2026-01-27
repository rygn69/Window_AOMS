VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frm_CAAndLiquidation 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cash Advance and Liquidation"
   ClientHeight    =   6720
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   14025
   Icon            =   "frm_CAAndLiquidation.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6720
   ScaleWidth      =   14025
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   1815
      Left            =   7080
      TabIndex        =   24
      Top             =   360
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   3201
      _Version        =   393216
      Appearance      =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Frame Frame3 
      Caption         =   "Profile"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   120
      TabIndex        =   17
      Top             =   240
      Width           =   6855
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
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   1320
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
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   840
         Width           =   4440
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
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   360
         Width           =   4440
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
         Left            =   555
         TabIndex        =   23
         Top             =   1320
         Width           =   1290
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
         Left            =   -120
         TabIndex        =   22
         Top             =   840
         Width           =   2025
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
         Left            =   -120
         TabIndex        =   21
         Top             =   360
         Width           =   1965
      End
   End
   Begin lvButton.lvButtons_H cmd_cancel 
      Height          =   495
      Left            =   12840
      TabIndex        =   16
      Top             =   6120
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   873
      Caption         =   "&Cancel"
      CapAlign        =   2
      BackStyle       =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
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
   Begin VB.Frame Frame1 
      Caption         =   "Cash Advance Details"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3735
      Left            =   120
      TabIndex        =   0
      Top             =   2280
      Width           =   6855
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
         Left            =   4920
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   3120
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
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   2640
         Width           =   1800
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
         Left            =   4920
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   2640
         Width           =   1560
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
         Left            =   1560
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Top             =   1320
         Width           =   5160
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
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   3120
         Width           =   1800
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
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   360
         Width           =   3495
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
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   840
         Width           =   3510
      End
      Begin lvButton.lvButtons_H lvButtons_H1 
         Height          =   375
         Left            =   5640
         TabIndex        =   25
         Top             =   360
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   661
         Caption         =   "&Save"
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
         Left            =   3960
         TabIndex        =   14
         Top             =   3120
         Width           =   930
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
         Left            =   -300
         TabIndex        =   13
         Top             =   465
         Width           =   1785
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
         Left            =   585
         TabIndex        =   12
         Top             =   2640
         Width           =   795
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
         Left            =   3885
         TabIndex        =   11
         Top             =   2640
         Width           =   960
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
         Left            =   480
         TabIndex        =   10
         Top             =   1320
         Width           =   885
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
         Left            =   -495
         TabIndex        =   9
         Top             =   840
         Width           =   1935
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
         Left            =   120
         TabIndex        =   8
         Top             =   3120
         Width           =   1380
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Liquidation Details"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3735
      Left            =   7080
      TabIndex        =   15
      Top             =   2280
      Width           =   6855
      Begin VB.OptionButton opt_GJ 
         Caption         =   "General Journal"
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
         Left            =   1800
         TabIndex        =   28
         Top             =   600
         Value           =   -1  'True
         Width           =   1695
      End
      Begin VB.OptionButton opt_CR 
         Caption         =   "Cash Receipt"
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
         Left            =   120
         TabIndex        =   27
         Top             =   600
         Width           =   1695
      End
      Begin lvButton.lvButtons_H lvButtons_H2 
         Height          =   375
         Left            =   5880
         TabIndex        =   26
         Top             =   360
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   661
         Caption         =   "&Save"
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
      Begin TabDlg.SSTab SSTab1 
         Height          =   3015
         Left            =   120
         TabIndex        =   31
         Top             =   600
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   5318
         _Version        =   393216
         Style           =   1
         Tabs            =   2
         Tab             =   1
         TabHeight       =   520
         TabCaption(0)   =   "Cash Receipts"
         TabPicture(0)   =   "frm_CAAndLiquidation.frx":076A
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "Label3"
         Tab(0).Control(1)=   "Label13"
         Tab(0).Control(2)=   "Label17"
         Tab(0).Control(3)=   "txtORNo"
         Tab(0).Control(4)=   "txtORAmount"
         Tab(0).Control(5)=   "DtpOrdate"
         Tab(0).ControlCount=   6
         TabCaption(1)   =   "Tab 1"
         TabPicture(1)   =   "frm_CAAndLiquidation.frx":0786
         Tab(1).ControlEnabled=   -1  'True
         Tab(1).Control(0)=   "Label20"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "Label19"
         Tab(1).Control(1).Enabled=   0   'False
         Tab(1).Control(2)=   "Label16"
         Tab(1).Control(2).Enabled=   0   'False
         Tab(1).Control(3)=   "Label11"
         Tab(1).Control(3).Enabled=   0   'False
         Tab(1).Control(4)=   "txtLJevno"
         Tab(1).Control(4).Enabled=   0   'False
         Tab(1).Control(5)=   "txtLDvno"
         Tab(1).Control(5).Enabled=   0   'False
         Tab(1).Control(6)=   "txtLamount"
         Tab(1).Control(6).Enabled=   0   'False
         Tab(1).Control(7)=   "txtLDetails"
         Tab(1).Control(7).Enabled=   0   'False
         Tab(1).ControlCount=   8
         Begin MSComCtl2.DTPicker DtpOrdate 
            Height          =   375
            Left            =   -73680
            TabIndex        =   45
            Top             =   1320
            Width           =   1935
            _ExtentX        =   3413
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
            Format          =   417267713
            CurrentDate     =   41515
         End
         Begin VB.TextBox txtORAmount 
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
            Left            =   -73680
            TabIndex        =   41
            Top             =   1800
            Width           =   1800
         End
         Begin VB.TextBox txtORNo 
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
            Left            =   -73680
            TabIndex        =   40
            Top             =   840
            Width           =   3495
         End
         Begin VB.TextBox txtLDetails 
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
            Height          =   795
            Left            =   1200
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   35
            Top             =   1440
            Width           =   5280
         End
         Begin VB.TextBox txtLamount 
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
            Left            =   1200
            TabIndex        =   34
            Top             =   2280
            Width           =   1800
         End
         Begin VB.TextBox txtLDvno 
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
            Left            =   1200
            TabIndex        =   33
            Top             =   480
            Width           =   3495
         End
         Begin VB.TextBox txtLJevno 
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
            Left            =   1200
            TabIndex        =   32
            TabStop         =   0   'False
            Top             =   960
            Width           =   3510
         End
         Begin VB.Label Label17 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "OR Number:"
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
            Left            =   -74820
            TabIndex        =   44
            Top             =   960
            Width           =   1065
         End
         Begin VB.Label Label13 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "OR date:"
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
            Left            =   -74565
            TabIndex        =   43
            Top             =   1320
            Width           =   765
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Amount:"
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
            Left            =   -74580
            TabIndex        =   42
            Top             =   1800
            Width           =   720
         End
         Begin VB.Label Label11 
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
            Left            =   555
            TabIndex        =   39
            Top             =   600
            Width           =   570
         End
         Begin VB.Label Label16 
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
            Left            =   240
            TabIndex        =   38
            Top             =   1440
            Width           =   885
         End
         Begin VB.Label Label19 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "JEVNO:"
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
            Left            =   330
            TabIndex        =   37
            Top             =   960
            Width           =   750
         End
         Begin VB.Label Label20 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Amount:"
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
            Left            =   300
            TabIndex        =   36
            Top             =   2280
            Width           =   720
         End
      End
      Begin VB.Label lbl_LiquidationStat 
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
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   840
         TabIndex        =   30
         Top             =   3480
         Width           =   1485
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Status:"
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
         Left            =   3720
         TabIndex        =   29
         Top             =   3120
         Width           =   615
      End
   End
End
Attribute VB_Name = "frm_CAAndLiquidation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public dvno, ClaimantCode As String


Private Sub cmd_cancel_Click()
IsOktoClear = False
Unload Me
End Sub

Private Sub Form_Load()
txtDVNo.Text = dvno
Call loadCA
Call loadLiquidation
Call Liquidetails
End Sub

Private Sub loadCA()
Dim PRec As New ADODB.Recordset
Set PRec = opndbaseFMIS.Execute("EXECUTE  [fmis].[dbo].[MPproc_DVNOtransDetails] @ID = 1,@DVNO = '" & txtDVNo.Text & "'")
    If PRec.RecordCount > 0 Then
        txtClaimant.Text = PRec!obrno
        txtCheckno.Text = PRec!checkno
        txtAmount.Text = PRec!amount
        txtClaimant.Text = PRec!payee
        txtDetail.Text = PRec!Particulars
        txtcheckdate.Text = PRec!CheckDate
        txtDaysPass.Text = PRec!dayspass
        txtOffice.Text = PRec!Officename
        ClaimantCode = PRec!ClaimantCode
        If Trim(txtDVNo.Text) <> "" Then
        txtEmail.Text = GetExcuteScalar(1, txtDVNo.Text)
        End If
        txtObR.Text = PRec!obrno
    End If
PRec.Close
Set PRec = Nothing
End Sub
Private Sub loadLiquidation()
Dim PRec As New ADODB.Recordset
Set PRec = opndbaseFMIS.Execute("EXECUTE  [fmis].[dbo].[MPproc_DVNOtransDetails] @ID = 2,@DVNO = '" & txtDVNo.Text & "'")
    If PRec.RecordCount > 0 Then
        txtLDvno.Text = PRec!dvno
        txtAmount.Text = PRec!amount
        txtLDetails.Text = PRec!Particulars
        txtOffice.Text = PRec!Officename
    End If
PRec.Close
Set PRec = Nothing
End Sub

Private Sub lvButtons_H1_Click()
    If MsgBox("Are you sure do you want to update?", vbInformation + vbYesNo, "System Confirmation") = vbYes Then
         opndbaseFMIS.Execute "update [fmis].[dbo].[tblAMIS_IncomingDVTrns] set [Particular] = '" & txtDetail.Text & "' where dvno = '" & txtDVNo.Text & "' and actioncode = 1"
    End If
End Sub

Private Sub lvButtons_H2_Click()

If MsgBox("Are you sure do you want to update?", vbInformation + vbYesNo, "System Confirmation") = vbYes Then
     'opndbaseFMIS.Execute "update [fmis].[dbo].[tblAMIS_IncomingDVTrns] set [Particular] = '" & txtLDetails.Text & "' where dvno = '" & txtLDvno.Text & "' and actioncode = 1"
    If opt_CR.Value = True Then
        If IsNumeric(txtORAmount.Text) = False Then
            MsgBox "None Numeric entry..", vbCritical, "System Message"
            txtORAmount.SetFocus
            Exit Sub
        End If
        If txtORNo.Text = "" Then
            MsgBox "OR Number is empty..", vbCritical, "System Message"
            txtORNo.SetFocus
            Exit Sub
        End If
        
        If CheckIfExists("SELECT [trnno] FROM [fmis].[dbo].[tblAMIS_LiquidationOfCAinOR] where ORNO = '" & txtORNo.Text & "'") Then
            MsgBox "OR Number Already exist in the database..", vbInformation, "System Information"
        Else
            opndbaseFMIS.Execute "insert into [fmis].[dbo].[tblAMIS_LiquidationOfCAinOR] ([DVNO],[ORNO],[ORdate],[amount],userid,DTE) values ('" & txtDVNo.Text & "','" & txtORNo.Text & "','" & DtpOrdate.Value & "','" & txtORAmount.Text & "','" & Trim(ActiveUserID) & "','" & Now & "')"
            MsgBox "Successfully Save", vbInformation, "System Message"
            IsOktoClear = True
            Unload Me
        End If
        
    ElseIf opt_GJ.Value = True Then
    
        If IsNumeric(txtLamount.Text) = False Then
            MsgBox "None Numeric entry..", vbCritical, "System Message"
            txtLamount.SetFocus
            Exit Sub
        End If
        
        If Trim(txtLamount.Text) = "" Then
            MsgBox "Amount is empty, please fill up the field", vbCritical, "System Message"
            txtLamount.SetFocus
            Exit Sub
        End If
        
        If txtLDetails.Text = "" Then
            MsgBox "Particular is empty, please fill up the field", vbCritical, "System Message"
            txtLDetails.SetFocus
            Exit Sub
        End If
        
        If CheckIfExists("SELECT [trnno] FROM [fmis].[dbo].[tblAMIS_LiquiditionOfCA] where [liquiDvno] = '" & txtLDvno.Text & "' and actioncode =1") Then
            MsgBox "Liquidation DVNO Already exist in the database..", vbInformation, "System Information"
        Else
            opndbaseFMIS.Execute "Insert into tblAMIS_LiquiditionOfCA ([liquiDvno],[CADvno],[checkno],[checkdate],[status],[actioncode],[amount],CAobrno,[CAParticular],[CAclaimantcode]) " & _
            " values ('" & txtLDvno.Text & "' , '" & txtDVNo.Text & "','" & txtCheckno.Text & "','" & txtcheckdate.Text & "',0,1, " & CCur(txtAmount.Text) & ",'" & txtObR.Text & "','" & Trim(Replace(txtDetail.Text, "'", "''")) & "','" & ClaimantCode & "') "
            MsgBox "Successfully Save", vbInformation, "System Message"
            Unload Me
        End If
    End If
End If
End Sub

Private Sub Opt_CR_Click()
Call Liquidetails
End Sub
Private Sub Liquidetails()
If opt_CR.Value = True Then
    SSTab1.Tab = 0
ElseIf opt_GJ.Value = True Then
    SSTab1.Tab = 1
End If
End Sub

Private Sub Opt_GJ_Click()
Call Liquidetails
End Sub

Private Sub txtLDvno_KeyPress(KeyAscii As Integer)
Dim Drec As New ADODB.Recordset
On Error GoTo bad
    If KeyAscii = 13 Then
        If CheckIfExists("SELECT [trnno] FROM [fmis].[dbo].[tblAMIS_LiquiditionOfCA] where [liquiDvno] = '" & txtLDvno.Text & "' and actioncode =1") Then
            MsgBox "Liquidation DVNO Already exist in the database..", vbInformation, "System Information"
                    txtLJevno.Text = ""
                    txtLamount.Text = ""
                    txtLDetails.Text = ""
        Else
            Set Drec = opndbaseFMIS.Execute("SELECT obrno, [DVNo],[Particular],[GAmount],(select top 1 jevno from dbo.[tblAMIS_FinalJEV] where dvno = '" & txtLDvno.Text & "' and actioncode = 1) as JEVNO FROM [fmis].[dbo].[tblAMIS_IncomingDVTrns] where ActionCode = 1 and dvno = '" & txtLDvno.Text & "'")
            If Drec.RecordCount > 0 Then
                If Drec!obrno = "NA-21" Then
                    txtLJevno.Text = IIf(IsNull(Drec!jevno), "", Drec!jevno)
                    txtLamount.Text = Format(Drec!Gamount, "#,##0.00")
                    txtLDetails.Text = Drec!Particular
                Else
                    MsgBox "This Transaction is not a Liquidation..", vbInformation, "System Mesagge"
                    txtLJevno.Text = ""
                    txtLamount.Text = ""
                    txtLDetails.Text = ""
                End If
            Else
                MsgBox "Invalid DVNO number...", vbInformation, "System Message"
            End If
            Drec.Close
            Set Drec = Nothing
        End If
    End If

Exit Sub
bad:
MsgBox err.description
End Sub

Private Sub txtORAmount_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    lvButtons_H2.SetFocus
End If
End Sub

Private Sub txtORAmount_LostFocus()
txtORAmount.Text = Format(txtORAmount.Text, "#,##0.00")
End Sub

Private Sub txtORNo_KeyPress(KeyAscii As Integer)
Dim Drec As New ADODB.Recordset
On Error GoTo bad
    If KeyAscii = 13 Then
        Set Drec = opndbaseFMIS.Execute("SELECT [ORNo],[ORDate],[Amount] FROM [fmis].[dbo].[tblCMS_CMCollectionTransaction] where ActionCode = 1 and ORNo = '" & txtORNo.Text & "'")
        If Drec.RecordCount > 0 Then
            txtORAmount.Text = Format(Drec!amount, "#,##0.00")
            DtpOrdate.Value = Drec!ordate
        Else
            MsgBox "Invalid OR number...", vbInformation, "System Message"
            txtORAmount.Text = ""
        End If
        Drec.Close
    Set Drec = Nothing
    End If
Exit Sub
bad:
MsgBox err.description
End Sub

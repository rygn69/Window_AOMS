VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{1693405E-2DC9-4248-B52F-4AC9145DA2AF}#1.0#0"; "WinXPC Engine.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frm_AccountcodeSub 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Accountcode and Explaination Classification Utility"
   ClientHeight    =   9000
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10395
   Icon            =   "frm_AccountcodeSub.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "frm_AccountcodeSub.frx":076A
   ScaleHeight     =   9000
   ScaleWidth      =   10395
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab SSTab1 
      Height          =   7575
      Left            =   120
      TabIndex        =   9
      Top             =   720
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   13361
      _Version        =   393216
      Tabs            =   4
      Tab             =   1
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Chart of Accounts"
      TabPicture(0)   =   "frm_AccountcodeSub.frx":AE19
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "LstAccountcode"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Schedule Management"
      TabPicture(1)   =   "frm_AccountcodeSub.frx":AE35
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label3"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "lblChildStat"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "lblmainStat"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "lvButtons_H17"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "DTPicker3"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Animation1"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Prog_Main"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "prog_ChildStat"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "lvButtons_H16"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "lvButtons_H9"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "lvButtons_H10"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "lvButtons_H14"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "lvButtons_H13"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "lstShed"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "lvButtons_H12"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "lvButtons_H11"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "cmb_fundtype"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).Control(18)=   "DTPicker1"
      Tab(1).Control(18).Enabled=   0   'False
      Tab(1).Control(19)=   "Check1"
      Tab(1).Control(19).Enabled=   0   'False
      Tab(1).Control(20)=   "lvButtons_H4"
      Tab(1).Control(20).Enabled=   0   'False
      Tab(1).Control(21)=   "Check2"
      Tab(1).Control(21).Enabled=   0   'False
      Tab(1).Control(22)=   "Check3"
      Tab(1).Control(22).Enabled=   0   'False
      Tab(1).Control(23)=   "lvButtons_H21"
      Tab(1).Control(23).Enabled=   0   'False
      Tab(1).ControlCount=   24
      TabCaption(2)   =   "Cash Flow Generation"
      TabPicture(2)   =   "frm_AccountcodeSub.frx":AE51
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lvButtons_H15"
      Tab(2).Control(1)=   "DTPicker2"
      Tab(2).Control(2)=   "ProgStat"
      Tab(2).Control(3)=   "lvButtons_H18"
      Tab(2).Control(4)=   "lvButtons_H19"
      Tab(2).Control(5)=   "Label5"
      Tab(2).Control(6)=   "Label4"
      Tab(2).Control(7)=   "lblStat"
      Tab(2).ControlCount=   8
      TabCaption(3)   =   "Final Posting"
      TabPicture(3)   =   "frm_AccountcodeSub.frx":AE6D
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label6"
      Tab(3).Control(1)=   "Label7"
      Tab(3).Control(2)=   "Label8"
      Tab(3).Control(3)=   "lvButtons_H20"
      Tab(3).Control(4)=   "DTPicker4"
      Tab(3).Control(5)=   "cmb_fundtype1"
      Tab(3).Control(6)=   "ListView1"
      Tab(3).ControlCount=   7
      Begin lvButton.lvButtons_H lvButtons_H21 
         Height          =   495
         Left            =   120
         TabIndex        =   49
         Top             =   765
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   873
         Caption         =   "Select unhide only"
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
      Begin VB.CheckBox Check3 
         Caption         =   "Selected Account only"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   48
         Top             =   360
         Value           =   1  'Checked
         Width           =   2175
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   3975
         Left            =   -74760
         TabIndex        =   47
         Top             =   3360
         Width           =   9735
         _ExtentX        =   17171
         _ExtentY        =   7011
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Fund Type"
            Object.Width           =   6174
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Year"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Month"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Post By"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Date Time Posted"
            Object.Width           =   4762
         EndProperty
      End
      Begin VB.ComboBox cmb_fundtype1 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   -72960
         Style           =   2  'Dropdown List
         TabIndex        =   42
         Top             =   1995
         Width           =   5535
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Default"
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
         Left            =   240
         TabIndex        =   24
         Top             =   1440
         Visible         =   0   'False
         Width           =   1095
      End
      Begin lvButton.lvButtons_H lvButtons_H4 
         Height          =   735
         Left            =   7800
         TabIndex        =   20
         Top             =   420
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   1296
         Caption         =   "&View"
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
         cGradient       =   0
         CapStyle        =   1
         Mode            =   0
         Value           =   0   'False
         Image           =   "frm_AccountcodeSub.frx":AE89
         ImgSize         =   24
         cBack           =   -2147483633
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Consolidated"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6360
         TabIndex        =   23
         Top             =   360
         Width           =   2175
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   3960
         TabIndex        =   18
         Top             =   420
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "MMMM dd, yyyy"
         Format          =   159252483
         UpDown          =   -1  'True
         CurrentDate     =   41047
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
         Left            =   3960
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   855
         Width           =   3255
      End
      Begin MSComctlLib.ListView LstAccountcode 
         Height          =   6975
         Left            =   -74880
         TabIndex        =   10
         Top             =   540
         Width           =   9975
         _ExtentX        =   17595
         _ExtentY        =   12303
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
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
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Accountcode"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Accountname"
            Object.Width           =   12347
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "ID"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "View in Schedule"
            Object.Width           =   0
         EndProperty
      End
      Begin lvButton.lvButtons_H lvButtons_H11 
         Height          =   375
         Left            =   960
         TabIndex        =   13
         Top             =   885
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
      Begin lvButton.lvButtons_H lvButtons_H12 
         Height          =   375
         Left            =   1920
         TabIndex        =   14
         Top             =   885
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
      Begin MSComctlLib.ListView lstShed 
         Height          =   5295
         Left            =   120
         TabIndex        =   15
         Top             =   1320
         Width           =   9975
         _ExtentX        =   17595
         _ExtentY        =   9340
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
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
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Accountcode"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Accountname"
            Object.Width           =   12347
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Hide"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "ID"
            Object.Width           =   0
         EndProperty
      End
      Begin lvButton.lvButtons_H lvButtons_H13 
         Height          =   375
         Left            =   120
         TabIndex        =   21
         Top             =   6615
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         Caption         =   "&Unhide"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
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
         Image           =   "frm_AccountcodeSub.frx":BADB
         ImgSize         =   24
         cBack           =   -2147483633
      End
      Begin lvButton.lvButtons_H lvButtons_H14 
         Height          =   375
         Left            =   120
         TabIndex        =   22
         Top             =   7020
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         Caption         =   "&Hide"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
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
         Image           =   "frm_AccountcodeSub.frx":C72D
         ImgSize         =   24
         cBack           =   -2147483633
      End
      Begin lvButton.lvButtons_H lvButtons_H10 
         Height          =   375
         Left            =   2160
         TabIndex        =   12
         Top             =   1695
         Visible         =   0   'False
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         Caption         =   "&Include in Schedule"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
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
         Image           =   "frm_AccountcodeSub.frx":10237
         ImgSize         =   24
         cBack           =   -2147483633
      End
      Begin lvButton.lvButtons_H lvButtons_H9 
         Height          =   375
         Left            =   240
         TabIndex        =   11
         Top             =   1695
         Visible         =   0   'False
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         Caption         =   "&Exclude in Schedule"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
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
         Image           =   "frm_AccountcodeSub.frx":11289
         ImgSize         =   24
         cBack           =   -2147483633
      End
      Begin lvButton.lvButtons_H lvButtons_H16 
         Height          =   735
         Left            =   8760
         TabIndex        =   26
         Top             =   420
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   1296
         Caption         =   "&Generate"
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
         cGradient       =   0
         CapStyle        =   1
         Mode            =   0
         Value           =   0   'False
         Image           =   "frm_AccountcodeSub.frx":122DB
         ImgSize         =   24
         cBack           =   -2147483633
      End
      Begin MSComctlLib.ProgressBar prog_ChildStat 
         Height          =   165
         Left            =   1920
         TabIndex        =   27
         Top             =   7590
         Visible         =   0   'False
         Width           =   8175
         _ExtentX        =   14420
         _ExtentY        =   291
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Max             =   0.1
         Scrolling       =   1
      End
      Begin MSComctlLib.ProgressBar Prog_Main 
         Height          =   165
         Left            =   1920
         TabIndex        =   29
         Top             =   6945
         Visible         =   0   'False
         Width           =   8175
         _ExtentX        =   14420
         _ExtentY        =   291
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Max             =   0.1
         Scrolling       =   1
      End
      Begin MSComCtl2.Animation Animation1 
         Height          =   495
         Left            =   1320
         TabIndex        =   31
         Top             =   6780
         Visible         =   0   'False
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   873
         _Version        =   393216
         AutoPlay        =   -1  'True
         FullWidth       =   25
         FullHeight      =   33
      End
      Begin MSComCtl2.DTPicker DTPicker3 
         Height          =   375
         Left            =   120
         TabIndex        =   32
         ToolTipText     =   "Select date that you want to Copy"
         Top             =   1500
         Visible         =   0   'False
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "MMMM yyyy"
         Format          =   159252483
         CurrentDate     =   41047
      End
      Begin lvButton.lvButtons_H lvButtons_H17 
         Height          =   375
         Left            =   2085
         TabIndex        =   33
         ToolTipText     =   "Copy the Accountcode to other Schedule Report Format"
         Top             =   1500
         Visible         =   0   'False
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         Caption         =   "&Add New"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
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
         Image           =   "frm_AccountcodeSub.frx":1332D
         ImgSize         =   24
         cBack           =   -2147483633
      End
      Begin lvButton.lvButtons_H lvButtons_H15 
         Height          =   735
         Left            =   -69240
         TabIndex        =   34
         Top             =   1620
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   1296
         Caption         =   "&Generate"
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
         cGradient       =   0
         CapStyle        =   1
         Mode            =   0
         Value           =   0   'False
         Image           =   "frm_AccountcodeSub.frx":13F7F
         ImgSize         =   24
         cBack           =   -2147483633
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   615
         Left            =   -72720
         TabIndex        =   35
         Top             =   1740
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   1085
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "MMMM yyyy"
         Format          =   159055875
         CurrentDate     =   41047
      End
      Begin MSComctlLib.ProgressBar ProgStat 
         Height          =   255
         Left            =   -74760
         TabIndex        =   38
         Top             =   3180
         Visible         =   0   'False
         Width           =   9735
         _ExtentX        =   17171
         _ExtentY        =   450
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
      Begin lvButton.lvButtons_H lvButtons_H18 
         Height          =   735
         Left            =   -66360
         TabIndex        =   39
         Top             =   1620
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   1296
         Caption         =   "&View Unfiltered"
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
         cGradient       =   0
         CapStyle        =   1
         Mode            =   0
         Value           =   0   'False
         Image           =   "frm_AccountcodeSub.frx":14FD1
         ImgSize         =   24
         cBack           =   -2147483633
      End
      Begin lvButton.lvButtons_H lvButtons_H19 
         Height          =   735
         Left            =   -67680
         TabIndex        =   40
         Top             =   1620
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   1296
         Caption         =   "&Cancel"
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
         cGradient       =   0
         CapStyle        =   1
         Mode            =   0
         Value           =   0   'False
         Image           =   "frm_AccountcodeSub.frx":1556B
         ImgSize         =   24
         cBack           =   -2147483633
      End
      Begin MSComCtl2.DTPicker DTPicker4 
         Height          =   495
         Left            =   -72960
         TabIndex        =   41
         Top             =   1320
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   873
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "MMMM yyyy"
         Format          =   159055875
         CurrentDate     =   41047
      End
      Begin lvButton.lvButtons_H lvButtons_H20 
         Height          =   975
         Left            =   -67320
         TabIndex        =   43
         Top             =   1440
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   1720
         Caption         =   "&POST"
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
         cGradient       =   0
         CapStyle        =   1
         Mode            =   0
         Value           =   0   'False
         Image           =   "frm_AccountcodeSub.frx":19075
         ImgSize         =   24
         cBack           =   -2147483633
      End
      Begin VB.Label Label8 
         Caption         =   "Posted History"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74760
         TabIndex        =   46
         Top             =   2880
         Width           =   8295
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fundtype:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74595
         TabIndex        =   45
         Top             =   2040
         Width           =   1410
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Month Year:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74820
         TabIndex        =   44
         Top             =   1320
         Width           =   1710
      End
      Begin VB.Label Label5 
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
         Left            =   -74160
         TabIndex        =   37
         Top             =   2700
         Width           =   7335
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Month Year:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   -74760
         TabIndex        =   36
         Top             =   1740
         Width           =   1965
      End
      Begin VB.Label lblmainStat 
         Alignment       =   2  'Center
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
         Height          =   240
         Left            =   1920
         TabIndex        =   30
         Top             =   6660
         Visible         =   0   'False
         Width           =   8100
      End
      Begin VB.Label lblChildStat 
         Alignment       =   2  'Center
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
         Height          =   240
         Left            =   1950
         TabIndex        =   28
         Top             =   7110
         Visible         =   0   'False
         Width           =   8100
      End
      Begin VB.Label lblStat 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   27.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   630
         Left            =   -74820
         TabIndex        =   25
         Top             =   3300
         Width           =   9735
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Month Year:"
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
         Left            =   2640
         TabIndex        =   19
         Top             =   540
         Width           =   1290
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fundtype:"
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
         Left            =   2640
         TabIndex        =   17
         Top             =   900
         Width           =   1215
      End
   End
   Begin lvButton.lvButtons_H lvButtons_H1 
      Height          =   495
      Left            =   9000
      TabIndex        =   0
      Top             =   8400
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Caption         =   "&Close"
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
      cGradient       =   0
      CapStyle        =   1
      Mode            =   0
      Value           =   0   'False
      Image           =   "frm_AccountcodeSub.frx":1A0C7
      ImgSize         =   24
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H lvButtons_H2 
      Height          =   495
      Left            =   4560
      TabIndex        =   1
      Top             =   11040
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      Caption         =   "&Export"
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
      cGradient       =   0
      CapStyle        =   1
      Mode            =   0
      Value           =   0   'False
      Image           =   "frm_AccountcodeSub.frx":1DBD1
      ImgSize         =   24
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H lvButtons_H3 
      Height          =   495
      Left            =   6000
      TabIndex        =   2
      Top             =   11040
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      Caption         =   "&Add"
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
      cGradient       =   0
      CapStyle        =   1
      Mode            =   0
      Value           =   0   'False
      Image           =   "frm_AccountcodeSub.frx":216DB
      ImgSize         =   24
      cBack           =   -2147483633
   End
   Begin MSComctlLib.ImageList itb32x32 
      Left            =   9120
      Top             =   1200
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
            Picture         =   "frm_AccountcodeSub.frx":251E5
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_AccountcodeSub.frx":26B77
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_AccountcodeSub.frx":28509
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_AccountcodeSub.frx":29E9B
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_AccountcodeSub.frx":2B82D
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_AccountcodeSub.frx":2D1BF
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_AccountcodeSub.frx":2EB51
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_AccountcodeSub.frx":304E3
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_AccountcodeSub.frx":31E75
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_AccountcodeSub.frx":33809
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_AccountcodeSub.frx":344E5
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_AccountcodeSub.frx":34DC5
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_AccountcodeSub.frx":35AA1
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_AccountcodeSub.frx":3677D
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_AccountcodeSub.frx":37459
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_AccountcodeSub.frx":38135
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_AccountcodeSub.frx":38E11
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin WinXPC_Engine.WindowsXPC WindowsXPC1 
      Left            =   6840
      Top             =   2280
      _ExtentX        =   6588
      _ExtentY        =   1085
      ColorScheme     =   2
      EngineStarted   =   -1  'True
      Common_Dialog   =   0   'False
   End
   Begin lvButton.lvButtons_H lvButtons_H5 
      Height          =   495
      Left            =   600
      TabIndex        =   3
      Top             =   10920
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      Caption         =   "&Export"
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
      cGradient       =   0
      CapStyle        =   1
      Mode            =   0
      Value           =   0   'False
      Image           =   "frm_AccountcodeSub.frx":396ED
      ImgSize         =   24
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H lvButtons_H6 
      Height          =   495
      Left            =   4920
      TabIndex        =   4
      Top             =   10200
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      Caption         =   "&New"
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
      Image           =   "frm_AccountcodeSub.frx":3D1F7
      ImgSize         =   24
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H lvButtons_H7 
      Height          =   495
      Left            =   7800
      TabIndex        =   5
      Top             =   10560
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      Caption         =   "&Delete"
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
      Image           =   "frm_AccountcodeSub.frx":40D01
      ImgSize         =   24
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H lvButtons_H8 
      Height          =   495
      Left            =   6480
      TabIndex        =   6
      Top             =   10560
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      Caption         =   "&Edit"
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
      Image           =   "frm_AccountcodeSub.frx":4480B
      ImgSize         =   24
      cBack           =   -2147483633
   End
   Begin VB.Label lblException 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Chart of Accounts"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   1770
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "List of Main chart of Accounts"
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
      Left            =   120
      TabIndex        =   7
      Top             =   390
      Width           =   2520
   End
End
Attribute VB_Name = "frm_AccountcodeSub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public IsCancel, IsCancelForPosted As Boolean

Private Sub Check1_Click()
If Check1.Value = 1 Then
    Call LoadMotherFund(Cmb_fundtype)
Else
    Call LoadFundType(Cmb_fundtype)
End If
End Sub

Private Sub cmb_FundType_Click()
Call LoadSched
End Sub

Private Sub DTPicker1_CallbackKeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer, ByVal CallbackField As String, CallbackDate As Date)
Call LoadSched
End Sub

Private Sub DTPicker1_Change()
Call LoadSched
End Sub

Private Sub DTPicker1_Click()
'Call LoadSched
End Sub

Private Sub Form_Load()
Call GetAccountNamebyorder(LstAccountcode, "Accountcode")
Call LoadFundType(Cmb_fundtype)
Call LoadFundType(cmb_fundtype1)
Call LoadPostedhistory
Cmb_fundtype.ListIndex = 0
SSTab1.Tab = 0
DTPicker1.Value = Now
End Sub
Public Function GetAccountNamebyorder(ByVal lst As ListView, ByVal Condition As String)
Dim rec As New ADODB.Recordset
Dim x
Dim z As Integer
'Condition = Replace(Condition, "'", "")
rec.Open "Select Accountcode,Accountname,[DisplayInSched] from tblREF_AIS_ChartOfAccountsMother order by " & Condition & "", opndbaseFMIS, adOpenStatic, adLockOptimistic
    lst.ListItems.Clear
    If rec.RecordCount > 0 Then
        For z = 1 To rec.RecordCount
                    Set x = lst.ListItems.Add(, , rec.Fields!accountcode)
                    x.SubItems(1) = Trim(rec.Fields!Accountname)
                    x.SubItems(3) = IIf((rec.Fields!DisplayInSched = 1), "True", "False")
            rec.MoveNext
        Next z
    End If
rec.Close
Set rec = Nothing
End Function
Public Sub LoadSched()
On Error GoTo bad
Dim rec As New ADODB.Recordset
Dim x
Dim z As Integer
rec.Open "EXEC dbo.[MPfunc_ViewSched] @fundcode = '" & Cmb_fundtype.ItemData(Cmb_fundtype.ListIndex) & "',@year = " & DTPicker1.Year & ",@month = " & DTPicker1.Month & "", opndbaseFMIS, adOpenStatic, adLockOptimistic
    lstShed.ListItems.Clear
    If rec.RecordCount > 0 Then
        For z = 1 To rec.RecordCount
            Set x = lstShed.ListItems.Add(, , Trim(rec.Fields!accountcode))
            x.SubItems(1) = Trim(rec.Fields!Accountname)
            x.SubItems(2) = Trim(rec.Fields!Hide)
            rec.MoveNext
        Next z
    End If
rec.Close
Set rec = Nothing
Exit Sub
bad:
MsgBox err.description
End Sub
Private Sub LstAccountcode_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
If ColumnHeader = "Accountname" Then
Call GetAccountNamebyorder(LstAccountcode, "Accountname")
ElseIf ColumnHeader = "Accountcode" Then
Call GetAccountNamebyorder(LstAccountcode, "Accountcode")
End If
End Sub
Private Sub LstAccountcode_DblClick()
With frm_AccountcodeSub1
.Address = Trim(LstAccountcode.SelectedItem.SubItems(1))
.col = 2
.accountcode = Trim(LstAccountcode.SelectedItem.Text)
.Subcode1 = Trim(LstAccountcode.SelectedItem.Text)
.Subdesc1 = Trim(LstAccountcode.SelectedItem.SubItems(1))
.Condition = "Subcode1 ='" & Trim(LstAccountcode.SelectedItem.Text) & "'"
.accntcode = Trim(LstAccountcode.SelectedItem.Text)
.Show 1
End With
End Sub
Private Sub lvButtons_H1_Click()
Unload Me
End Sub

Private Sub lvButtons_H10_Click()
Dim x As Long
If MsgBox("Are you sure do you want to INCLUDE the Check item in Schedule?", vbInformation + vbYesNo, "System Confirmation") = vbYes Then
For x = 1 To LstAccountcode.ListItems.Count
If LstAccountcode.ListItems(x).Checked = True Then
    opndbaseFMIS.Execute ("Update [tblREF_AIS_ChartOfAccountsMother] set [DisplayInSched] = 1 where accountcode = '" & LstAccountcode.ListItems(x).Text & "'")
End If
Next x
MsgBox "Successfully Updated...!", vbInformation, "System Message"
Call GetAccountNamebyorder(LstAccountcode, "Accountcode")
End If
End Sub

Private Sub lvButtons_H11_Click()
Dim www As Integer
   ' If chkSelected.Value = 1 Then Exit Sub
    For www = 1 To lstShed.ListItems.Count
        lstShed.ListItems(www).Checked = True
    Next www
End Sub
Private Sub lvButtons_H12_Click()
Dim www As Integer
'    If chkSelected.Value = 1 Then Exit Sub
    For www = 1 To lstShed.ListItems.Count
        lstShed.ListItems(www).Checked = False
    Next www
End Sub

Private Sub lvButtons_H13_Click()
Dim z As Integer
If MsgBox("Are you sure Do you want to UNHIDE all checked account?", vbInformation + vbYesNo, "System Message") = vbYes Then
    For z = 1 To lstShed.ListItems.Count
        If lstShed.ListItems(z).Checked = True Then
            opndbaseFMIS.Execute "Insert into [fmis].[dbo].[tblAMIS_SheduleMngtViewer] (Accountcode,Fundcode,year_,month_) values ('" & Trim(lstShed.ListItems(z).Text) & "','" & Cmb_fundtype.ItemData(Cmb_fundtype.ListIndex) & "'," & DTPicker1.Year & "," & DTPicker1.Month & ")"
        End If
    Next z
End If
Call LoadSched
End Sub

Private Sub lvButtons_H14_Click()
Dim x As Long
If MsgBox("Are you sure do you want to HIDE the Check Account in Schedule?", vbInformation + vbYesNo, "System Confirmation") = vbYes Then
For x = 1 To lstShed.ListItems.Count
If lstShed.ListItems(x).Checked = True Then
    opndbaseFMIS.Execute ("Delete FROM [fmis].[dbo].[tblAMIS_SheduleMngtViewer] where accountcode = '" & lstShed.ListItems(x).Text & "' and year_ = " & DTPicker1.Year & " and month_ = " & DTPicker1.Month & " and Fundcode = '" & Cmb_fundtype.ItemData(Cmb_fundtype.ListIndex) & "'")
End If
Next x
Call LoadSched
End If
End Sub

Private Sub lvButtons_H16_Click()
On Error GoTo bad
Dim rec As New ADODB.Recordset
Dim x, y As Long

Dim objCommand As ADODB.command
Set objCommand = New ADODB.command
objCommand.CommandTimeout = 9999999
objCommand.ActiveConnection = opndbaseFMIS



MDIFrm_MAIN.disableEnabletimer (False)
If lstShed.ListItems.Count = 0 Then
    MsgBox "Please Add Accountcode, to Proceed the Generation", vbCritical, "System Message"
    Exit Sub
End If


If Check3.Value = 1 Then
    generate_Selected_Account
Exit Sub
End If
'MDIFrm_MAIN.tmeConnChck.Enabled = False
lblmainStat.Visible = True
Prog_Main.Visible = True
lblChildStat.Visible = True
prog_ChildStat.Visible = True
Prog_Main.Max = lstShed.ListItems.Count
    If MsgBox("Are you sure do you want to GENERATE for Schedule?. Your Previous Generation will be Lost.", vbCritical + vbYesNo, "System Confirmation") = vbYes Then
    opndbaseFMIS.Execute ("Delete from dbo.tblAMIS_Schedule where  Fundtype = '" & Cmb_fundtype.Text & "' and year_ = '" & DTPicker1.Year & "' and month_= '" & DTPicker1.Month & "'")
    Call PlayAVI(Me.Animation1, "Refresh.avi")
        Animation1.Visible = True
    For x = 1 To lstShed.ListItems.Count
        Prog_Main.Value = x
        lblmainStat.Caption = "Accountcode: " & Trim(lstShed.ListItems(x).Text) & "      " & x & " / " & lstShed.ListItems.Count
        If lstShed.ListItems(x).SubItems(2) = "No" Then
            
            
            objCommand.CommandText = "EXECUTE [dbo].[usp_insert_Schedule_byAccount] @Fundcode = '" & Cmb_fundtype.ItemData(Cmb_fundtype.ListIndex) & "',@date = '" & DTPicker1.Value & "',@fundtype ='" & Cmb_fundtype.Text & "',@Accountcode = '" & Trim(lstShed.ListItems(x).Text) & "'"
            Set rec = objCommand.Execute
            'Set rec = opndbaseFMIS.Execute("EXECUTE [dbo].[usp_insert_Schedule_byAccount] @Fundcode = '" & cmb_FundType.ItemData(cmb_FundType.ListIndex) & "',@date = '" & DTPicker1.Value & "',@fundtype ='" & cmb_FundType.Text & "',@Accountcode = '" & Trim(lstShed.ListItems(x).Text) & "'")
        End If
        DoEvents
    Next x
    Call StopAvi(Me.Animation1)
            Animation1.Visible = False
    End If
    MsgBox "Generated Successfully", vbInformation, "System Information"
    MDIFrm_MAIN.disableEnabletimer (True)
lblmainStat.Visible = False
Prog_Main.Visible = False
lblChildStat.Visible = False
prog_ChildStat.Visible = False
'MDIFrm_MAIN.tmeConnChck.Enabled = True
Exit Sub
bad:
MsgBox err.description & " Error Number: " & err.Number
End Sub
Private Sub generate_Selected_Account()
On Error GoTo bad
Dim rec As New ADODB.Recordset
Dim x, y As Long

Dim objCommand As ADODB.command
Set objCommand = New ADODB.command
objCommand.CommandTimeout = 9999999
objCommand.ActiveConnection = opndbaseFMIS

MDIFrm_MAIN.disableEnabletimer (False)
lblmainStat.Visible = True
Prog_Main.Visible = True
lblChildStat.Visible = True
prog_ChildStat.Visible = True
Prog_Main.Max = lstShed.ListItems.Count
    If MsgBox("Are you sure do you want to generate selected account for Schedule?.", vbCritical + vbYesNo, "System Confirmation") = vbYes Then
   ' opndbaseFMIS.Execute ("Delete from dbo.tblAMIS_Schedule where  Fundtype = '" & cmb_fundtype.Text & "' and year_ = '" & DTPicker1.Year & "' and month_= '" & DTPicker1.Month & "'")
    Call PlayAVI(Me.Animation1, "Refresh.avi")
        Animation1.Visible = True
    For x = 1 To lstShed.ListItems.Count
        Prog_Main.Value = x
        lblmainStat.Caption = "Accountcode: " & Trim(lstShed.ListItems(x).Text) & "      " & x & " / " & lstShed.ListItems.Count
        If lstShed.ListItems(x).Checked = True Then
            objCommand.CommandText = "EXECUTE [dbo].[usp_insert_Schedule_byAccount] @Fundcode = '" & Cmb_fundtype.ItemData(Cmb_fundtype.ListIndex) & "',@date = '" & DTPicker1.Value & "',@fundtype ='" & Cmb_fundtype.Text & "',@Accountcode = '" & Trim(lstShed.ListItems(x).Text) & "'"
            Set rec = objCommand.Execute
            'Set rec = opndbaseFMIS.Execute("EXECUTE [dbo].[usp_insert_Schedule_byAccount] @Fundcode = '" & cmb_FundType.ItemData(cmb_FundType.ListIndex) & "',@date = '" & DTPicker1.Value & "',@fundtype ='" & cmb_FundType.Text & "',@Accountcode = '" & Trim(lstShed.ListItems(x).Text) & "'")
        End If
        DoEvents
    Next x
    Call StopAvi(Me.Animation1)
    Animation1.Visible = False
    End If
    MsgBox "Generated Successfully", vbInformation, "System Information"
    MDIFrm_MAIN.disableEnabletimer (True)
lblmainStat.Visible = False
Prog_Main.Visible = False
lblChildStat.Visible = False
prog_ChildStat.Visible = False
'MDIFrm_MAIN.tmeConnChck.Enabled = True
Exit Sub
bad:
MsgBox err.description & " Error Number: " & err.Number
End Sub
Private Sub lvButtons_H18_Click()
frm_UnfilteredcashFlow.Show
End Sub

Private Sub lvButtons_H19_Click()
If MsgBox("Are you sure do you want to Cancel The Cash flow Generation?", vbInformation + vbYesNo, "System Confirmation") = vbYes Then
    IsCancel = True
End If
End Sub

Private Sub lvButtons_H20_Click()
Dim rec As New ADODB.Recordset
On Error GoTo bad
If MsgBox("Are you sure do you want to Final Post the Covered trasaction?", vbInformation + vbYesNo, "System Message") = vbYes Then
    opndbaseFMIS.Execute "Update [fmis].[dbo].[tblAMIS_FinalJEV] set posted = 1,Postedby = '" & ActiveUserID & "',PostedDTE = '" & Now & "' where fundtype = '" & cmb_fundtype1.Text & "' and year(jevdate) = " & DTPicker4.Year & " and month(jevdate) = " & DTPicker4.Month & " and actioncode = 1 and posted = 0"
    MsgBox cmb_fundtype1.Text & " as of " & Format(DTPicker4.Value, "MMMM dd, yyyy") & "Successfully Posted and Never Change", vbInformation, "System Message"
    Call LoadPostedhistory
End If
Exit Sub
bad:
MsgBox err.description
End Sub

Private Sub lvButtons_H21_Click()
Dim www As Integer
   ' If chkSelected.Value = 1 Then Exit Sub
    For www = 1 To lstShed.ListItems.Count
    If lstShed.ListItems(www).SubItems(2) = "No" Then
        lstShed.ListItems(www).Checked = True
    End If
        
    Next www
End Sub


Private Sub lvButtons_H4_Click()
'Dim x As Integer
'For x = 1 To LstAccountcode.ListItems.Count
'opndbaseFMIS.Execute "Insert into tblReff_CodeClassification(subcode1,subdesc1) values('" & LstAccountcode.ListItems(x).Text & "','" & Replace(LstAccountcode.ListItems(x).ListSubItems(1).Text, "'", "") & "') "
'Next x
'MsgBox "Successfully save", vbInformation, "System Message"
If Cmb_fundtype.Text = "" Then
    MsgBox "Please Specify the Fundtype..!", vbCritical, "System Message"
Else
LoadSched
End If
End Sub

Private Sub lvButtons_H15_Click()
Dim rec As New ADODB.Recordset
'On Error GoTo bad
IsCancel = False
Set rec = opndbaseFMIS.Execute("Select jevno From dbo.tblAMIS_FinalJEV where year(jevdate) = '" & DTPicker2.Year & "' and month(Jevdate) = '" & DTPicker2.Month & "' and  actioncode = 1 and filterInCashflow = 0 order by jevno")
    If rec.RecordCount > 0 Then
    progStat.Visible = True
    Label5.Visible = True
    progStat.Max = rec.RecordCount
        For x = 1 To rec.RecordCount
            If IsCancel = True Then
                Exit For
            Else
                'opndbaseFMIS.Execute "exec fmis.dbo.MPproc_SCFFilter @jevno = '" & rec!jevno & "'"
                opndbaseFMIS.Execute "exec fmis.dbo.MPproc_GetCashFlowEntry_New @jevno = '" & Trim(rec!jevno) & "'"
                progStat.Value = x
                Label5.Caption = rec!jevno & "  " & x & "/" & progStat.Max
                rec.MoveNext
                DoEvents
            End If
        Next x
        Label5.Caption = ""
    progStat.Visible = False
    rec.Close
    End If
    
       Set rec = opndbaseFMIS.Execute("Select jevno From dbo.tblAMIS_FinalJEV where year(jevdate) = '" & DTPicker2.Year & "' and month(Jevdate) = '" & DTPicker2.Month & "' and  actioncode = 1 and filterInCashflow = 0 order by trnno")
        If rec.RecordCount > 0 Then
            If MsgBox(rec.RecordCount & " Transaction cannot Filter the Entry Into Cash Flow. Do you want To view unfiltered Transaction?", vbCritical + vbYesNo, "System Message") = vbYes Then
                frm_UnfilteredcashFlow.Show
            End If
        End If
Exit Sub
bad:
MsgBox err.description
End Sub
Public Function LoadPostedhistory()
'
'Dim rec As New ADODB.Recordset
'Dim x
'Dim z As Integer
''Condition = Replace(Condition, "'", "")
'rec.Open "Select fundtype,Year(jevdate) as Year_,month(jevdate) as month_,Postedby,PostedDTE from tblAMIS_FinalJEV where actioncode = 1 and posted = 1 group by fundtype,Year(jevdate),month(jevdate) ,Postedby,PostedDTE order by Year(jevdate) asc,month(jevdate) asc", opndbaseFMIS, adOpenStatic, adLockOptimistic
'    ListView1.ListItems.Clear
'    If rec.RecordCount > 0 Then
'        For z = 0 To rec.RecordCount
'                    Set x = ListView1.ListItems.Add(, , rec.Fields!FundType)
'                    x.SubItems(1) = Trim(rec.Fields!YEAR_)
'                    x.SubItems(2) = Trim(rec.Fields!month_)
'                    x.SubItems(3) = Trim(rec.Fields!Postedby)
'                    x.SubItems(4) = Trim(rec.Fields!PostedDTE)
'            rec.MoveNext
'        Next z
'    End If
'rec.Close
'Set rec = Nothing
End Function

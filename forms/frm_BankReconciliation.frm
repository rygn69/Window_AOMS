VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{1693405E-2DC9-4248-B52F-4AC9145DA2AF}#1.0#0"; "WinXPC Engine.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frm_BankReconciliation 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bank Reconciliation"
   ClientHeight    =   10005
   ClientLeft      =   2280
   ClientTop       =   1170
   ClientWidth     =   15705
   Icon            =   "frm_BankReconciliation.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   10005
   ScaleWidth      =   15705
   Begin VB.Frame Frame1 
      Caption         =   "Bank Statement"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8415
      Left            =   8520
      TabIndex        =   12
      Top             =   1560
      Width           =   7095
      Begin VB.ComboBox cmbBookClass 
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
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   38
         Top             =   5040
         Width           =   6855
      End
      Begin VB.OptionButton opt_Statement 
         Caption         =   "Statement"
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
         Left            =   4680
         Style           =   1  'Graphical
         TabIndex        =   16
         Tag             =   "4"
         Top             =   360
         UseMaskColor    =   -1  'True
         Width           =   2295
      End
      Begin VB.OptionButton Opt_Withdrawals 
         Caption         =   "Withdrawals"
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
         Left            =   2400
         Style           =   1  'Graphical
         TabIndex        =   15
         Tag             =   "2"
         Top             =   360
         UseMaskColor    =   -1  'True
         Width           =   2295
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   50
         Left            =   2280
         Top             =   2715
      End
      Begin VB.OptionButton opt_Deposits 
         Caption         =   "Deposits"
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
         Style           =   1  'Graphical
         TabIndex        =   13
         Tag             =   "1"
         Top             =   360
         UseMaskColor    =   -1  'True
         Width           =   2295
      End
      Begin MSComctlLib.ListView lst_Bank 
         Height          =   3495
         Left            =   120
         TabIndex        =   14
         Top             =   960
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   6165
         SortKey         =   3
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
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
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Date"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Description"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Key             =   "check"
            Text            =   "Ref./Checkno"
            Object.Width           =   2822
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Amount"
            Object.Width           =   2822
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "ID"
            Object.Width           =   0
         EndProperty
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   2895
         Left            =   120
         TabIndex        =   36
         Top             =   5400
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   5106
         SortKey         =   3
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
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
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Date"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Description"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Key             =   "check"
            Text            =   "Ref./Checkno"
            Object.Width           =   2822
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Amount"
            Object.Width           =   2822
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "ID"
            Object.Width           =   0
         EndProperty
      End
      Begin lvButton.lvButtons_H lvButtons_H1 
         Height          =   375
         Left            =   3120
         TabIndex        =   39
         ToolTipText     =   "Add Check Item"
         Top             =   4560
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         CapAlign        =   2
         BackStyle       =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         Image           =   "frm_BankReconciliation.frx":0E42
         cBack           =   -2147483633
      End
      Begin lvButton.lvButtons_H lvButtons_H2 
         Height          =   375
         Left            =   3600
         TabIndex        =   40
         ToolTipText     =   "Remove Check Item"
         Top             =   4560
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         CapAlign        =   2
         BackStyle       =   4
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
         Mode            =   0
         Value           =   0   'False
         Image           =   "frm_BankReconciliation.frx":494C
         cBack           =   -2147483633
      End
   End
   Begin VB.Frame Frame5 
      Height          =   1545
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   15435
      Begin VB.Frame Frame4 
         Caption         =   "Depository Bank"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   900
         Left            =   4080
         TabIndex        =   10
         Top             =   360
         Width           =   3945
         Begin VB.ComboBox cmb_BankName 
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
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   360
            Width           =   3660
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Bank Account Number"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   900
         Left            =   8160
         TabIndex        =   8
         Top             =   360
         Width           =   3345
         Begin VB.ComboBox cmb_Accountno 
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
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   360
            Width           =   3060
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Special Account"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   900
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Width           =   3690
         Begin VB.ComboBox cmb_FundType 
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
            Left            =   195
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   360
            Width           =   3300
         End
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Load"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   915
         Left            =   14160
         TabIndex        =   5
         Top             =   360
         Width           =   1005
      End
      Begin VB.Frame Frame6 
         Caption         =   "Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   900
         Left            =   11640
         TabIndex        =   3
         Top             =   360
         Width           =   2385
         Begin MSComCtl2.DTPicker DTPicker3 
            CausesValidation=   0   'False
            Height          =   390
            Left            =   120
            TabIndex        =   4
            Top             =   360
            Width           =   2085
            _ExtentX        =   3678
            _ExtentY        =   688
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
            Format          =   180355075
            UpDown          =   -1  'True
            CurrentDate     =   38240
         End
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00000000&
         BackStyle       =   1  'Opaque
         Height          =   1185
         Left            =   120
         Top             =   240
         Width           =   15225
      End
   End
   Begin WinXPC_Engine.WindowsXPC WindowsXPC1 
      Left            =   4770
      Top             =   11265
      _ExtentX        =   6588
      _ExtentY        =   1085
      ColorScheme     =   2
      Common_Dialog   =   0   'False
      PictureControl  =   0   'False
   End
   Begin VB.Frame Frame3 
      Caption         =   "Button"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   10335
      Left            =   120
      TabIndex        =   17
      Top             =   1560
      Width           =   1095
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   8910
         Left            =   120
         TabIndex        =   18
         Top             =   360
         Width           =   810
         _ExtentX        =   1429
         _ExtentY        =   15716
         ButtonWidth     =   1402
         ButtonHeight    =   1429
         Style           =   1
         ImageList       =   "itb32x32"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   11
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&New"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "slash"
               Style           =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Caption         =   "&Edit"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Caption         =   "&Save"
               ImageIndex      =   9
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&Match"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Caption         =   "&Delete"
               ImageIndex      =   12
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "s"
               Style           =   3
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&Cancel"
               ImageIndex      =   7
            EndProperty
         EndProperty
         Begin MSComCtl2.Animation Animation1 
            Height          =   450
            Left            =   11400
            TabIndex        =   19
            Top             =   120
            Width           =   480
            _ExtentX        =   847
            _ExtentY        =   794
            _Version        =   393216
            FullWidth       =   32
            FullHeight      =   30
         End
         Begin MSComctlLib.ImageList itb32x32 
            Left            =   120
            Top             =   5520
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
                  Picture         =   "frm_BankReconciliation.frx":8456
                  Key             =   ""
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frm_BankReconciliation.frx":9DE8
                  Key             =   ""
               EndProperty
               BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frm_BankReconciliation.frx":B77A
                  Key             =   ""
               EndProperty
               BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frm_BankReconciliation.frx":D10C
                  Key             =   ""
               EndProperty
               BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frm_BankReconciliation.frx":EA9E
                  Key             =   ""
               EndProperty
               BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frm_BankReconciliation.frx":10430
                  Key             =   ""
               EndProperty
               BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frm_BankReconciliation.frx":11DC2
                  Key             =   ""
               EndProperty
               BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frm_BankReconciliation.frx":13754
                  Key             =   ""
               EndProperty
               BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frm_BankReconciliation.frx":150E6
                  Key             =   ""
               EndProperty
               BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frm_BankReconciliation.frx":16A7A
                  Key             =   ""
               EndProperty
               BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frm_BankReconciliation.frx":17756
                  Key             =   ""
               EndProperty
               BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frm_BankReconciliation.frx":18036
                  Key             =   ""
               EndProperty
               BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frm_BankReconciliation.frx":18D12
                  Key             =   ""
               EndProperty
               BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frm_BankReconciliation.frx":199EE
                  Key             =   ""
               EndProperty
               BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frm_BankReconciliation.frx":1A6CA
                  Key             =   ""
               EndProperty
               BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frm_BankReconciliation.frx":1B3A6
                  Key             =   ""
               EndProperty
               BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frm_BankReconciliation.frx":1C082
                  Key             =   ""
               EndProperty
            EndProperty
         End
      End
   End
   Begin VB.Frame fme_Details 
      Caption         =   "Add New Entry"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4455
      Index           =   1
      Left            =   1320
      TabIndex        =   21
      Top             =   1560
      Visible         =   0   'False
      Width           =   7095
      Begin VB.TextBox txtID 
         Height          =   285
         Left            =   4800
         TabIndex        =   34
         Top             =   480
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txtAmount 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2040
         ScrollBars      =   2  'Vertical
         TabIndex        =   32
         Top             =   2280
         Width           =   2175
      End
      Begin VB.TextBox txtcheckno 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2040
         ScrollBars      =   2  'Vertical
         TabIndex        =   30
         Top             =   1800
         Width           =   2175
      End
      Begin VB.TextBox txtdescription 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   2040
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   28
         Text            =   "frm_BankReconciliation.frx":1C95E
         Top             =   960
         Width           =   4815
      End
      Begin MSComCtl2.DTPicker dt_checkdate 
         Height          =   375
         Left            =   2040
         TabIndex        =   26
         Top             =   405
         Width           =   2175
         _ExtentX        =   3836
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
         Format          =   180355073
         CurrentDate     =   41035
      End
      Begin VB.Timer Timer2 
         Enabled         =   0   'False
         Index           =   1
         Interval        =   50
         Left            =   2280
         Top             =   5835
      End
      Begin VB.Label Label4 
         Caption         =   "Amount:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   31
         Top             =   2370
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   "Ref./Checkno:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   29
         Top             =   1800
         Width           =   1815
      End
      Begin VB.Label Label2 
         Caption         =   "Description:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   27
         Top             =   960
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Deposit/Check date:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   22
         Top             =   480
         Width           =   1815
      End
   End
   Begin VB.Frame Frame8 
      Caption         =   "Book"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8415
      Index           =   0
      Left            =   1320
      TabIndex        =   20
      Top             =   1560
      Width           =   7095
      Begin VB.ComboBox cmbBankClass 
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
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   37
         Top             =   5040
         Width           =   6855
      End
      Begin VB.OptionButton Opt_GJ 
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
         Height          =   495
         Left            =   4680
         Style           =   1  'Graphical
         TabIndex        =   23
         Tag             =   "4"
         Top             =   360
         UseMaskColor    =   -1  'True
         Width           =   2295
      End
      Begin VB.Timer Timer2 
         Enabled         =   0   'False
         Index           =   0
         Interval        =   50
         Left            =   2280
         Top             =   2715
      End
      Begin VB.OptionButton opt_Check 
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
         Height          =   495
         Left            =   2400
         Style           =   1  'Graphical
         TabIndex        =   24
         Tag             =   "2"
         Top             =   360
         UseMaskColor    =   -1  'True
         Width           =   2295
      End
      Begin VB.OptionButton opt_CR 
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
         Height          =   495
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   25
         Tag             =   "1"
         Top             =   360
         UseMaskColor    =   -1  'True
         Width           =   2295
      End
      Begin MSComctlLib.ListView lst_Journal 
         Height          =   3495
         Left            =   120
         TabIndex        =   33
         Top             =   960
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   6165
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483641
         BackColor       =   16777215
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
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Date"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Description"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Ref./Checkno"
            Object.Width           =   2822
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Amount"
            Object.Width           =   2822
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "ID"
            Object.Width           =   0
         EndProperty
      End
      Begin MSComctlLib.ListView ListView3 
         Height          =   2895
         Left            =   120
         TabIndex        =   35
         Top             =   5400
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   5106
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
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   11
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Date"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Description"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Ref./Checkno"
            Object.Width           =   2822
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Amount"
            Object.Width           =   2822
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "ID"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Date"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Description"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Ref./Checkno"
            Object.Width           =   2822
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "Amount"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "ID"
            Object.Width           =   2540
         EndProperty
      End
      Begin lvButton.lvButtons_H lvButtons_H3 
         Height          =   375
         Left            =   3120
         TabIndex        =   41
         ToolTipText     =   "Add Check Item"
         Top             =   4560
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         CapAlign        =   2
         BackStyle       =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         Image           =   "frm_BankReconciliation.frx":1C964
         cBack           =   -2147483633
      End
      Begin lvButton.lvButtons_H lvButtons_H4 
         Height          =   375
         Left            =   3600
         TabIndex        =   42
         ToolTipText     =   "Remove Check Item"
         Top             =   4560
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         CapAlign        =   2
         BackStyle       =   4
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
         Mode            =   0
         Value           =   0   'False
         Image           =   "frm_BankReconciliation.frx":2046E
         cBack           =   -2147483633
      End
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label14"
      Height          =   195
      Left            =   13755
      TabIndex        =   2
      Top             =   9975
      Width           =   570
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label8"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   105
      TabIndex        =   0
      Top             =   9300
      Width           =   480
   End
End
Attribute VB_Name = "frm_BankReconciliation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim tmpAccName As String
Dim FMISNo As String

Private Sub cmb_BankName_Change()
Call Loadcmb(cmb_Accountno, "EXECUTE [fmis].[dbo].[MPproc_Get_AccountNoByBankname] @fundtype = '" & cmb_fundtype.Text & "', @bankid = " & cmb_BankName.ItemData(cmb_BankName.ListIndex) & "")
End Sub

Private Sub cmb_BankName_Click()
Call Loadcmb(cmb_Accountno, "EXECUTE [fmis].[dbo].[MPproc_Get_AccountNoByBankname] @fundtype = '" & cmb_fundtype.Text & "', @bankid = " & cmb_BankName.ItemData(cmb_BankName.ListIndex) & "")
End Sub

Private Sub Command1_Click()
LoadOPTBook
LoadOPTBank
End Sub

Private Sub Form_Load()
'WindowsXPC1.InitSubClassing

Call LoadFundType(cmb_fundtype)
Call Loadcmb(cmb_BankName, "SELECT [trnno] as field1,[BankName] as field2 FROM [fmis].[dbo].[tblCMS_CDBankLibrary]")
Call Loadcmb(cmbBankClass, "EXECUTE [fmis].[dbo].[MPproc_BankSubClass] @what = 'bank'")
Call Loadcmb(cmbBookClass, "EXECUTE [fmis].[dbo].[MPproc_BankSubClass] @what = 'Book'")
'Call LoadSavedReport(ActiveUserID, DTPicker1.Year, DTPicker1.Month, cmb_FundType.Text)
End Sub



Private Sub lst_Bank_Click()
Dim x As Long
Dim Ok As Boolean
Ok = False
If lst_Bank.ListItems.Count = 0 Then
    Exit Sub
End If
With lst_Journal

    For x = 1 To .ListItems.Count
            .ListItems(x).Bold = False
            .ListItems(x).ListSubItems(1).Bold = False
            .ListItems(x).ListSubItems(2).Bold = False
            .ListItems(x).ListSubItems(3).Bold = False

            .ListItems(x).ForeColor = &H0&
            .ListItems(x).ListSubItems(1).ForeColor = &H0&
            .ListItems(x).ListSubItems(2).ForeColor = &H0&
            .ListItems(x).ListSubItems(3).ForeColor = &H0&
        If lst_Bank.SelectedItem.SubItems(2) = .ListItems(x).SubItems(2) Then
            lst_Journal.HideSelection = True
            .ListItems(x).Selected = True
            
            .ListItems(x).Bold = True
            .ListItems(x).ListSubItems(1).Bold = True
            .ListItems(x).ListSubItems(2).Bold = True
            .ListItems(x).ListSubItems(3).Bold = True
            .ListItems(x).Top = 1
            .ListItems(x).ForeColor = &HFF&
            .ListItems(x).ListSubItems(1).ForeColor = &HFF&
            .ListItems(x).ListSubItems(2).ForeColor = &HFF&
            .ListItems(x).ListSubItems(3).ForeColor = &HFF&
            Ok = True
        End If
    Next x
.Refresh
If Ok = False Then
    MsgBox "No Match Found...!", vbInformation, "System Message"
    lst_Journal.HideSelection = True
End If
End With
If Trim(lst_Bank.SelectedItem.Text) <> "" Then
    Call toolbarSTat("Open")
End If
End Sub





Private Sub lst_Journal_KeyUp(KeyCode As Integer, Shift As Integer)
KeyCode = vbUpArrow
End Sub

Private Sub opt_Check_Click()
LoadOPTBook

End Sub

Private Sub Opt_CR_Click()
LoadOPTBook
End Sub

Private Sub opt_Deposits_Click()
LoadOPTBank
End Sub
Private Sub LoadOPTBank()
If Opt_Withdrawals.Value = True Then
    opt_Check.Value = True
    Call LoadReconData(2, lst_Bank, 2)
ElseIf opt_Deposits.Value = True Then
    opt_CR.Value = True
    Call LoadReconData(1, lst_Bank, 2)
ElseIf opt_Statement.Value = True Then
    opt_GJ.Value = True
    Call LoadReconData(4, lst_Bank, 2)
End If
End Sub
Private Sub LoadOPTBook()
If opt_Check.Value = True Then
    Opt_Withdrawals.Value = True
    Call LoadReconData(2, lst_Journal, 1)
ElseIf opt_CR.Value = True Then
    opt_Deposits.Value = True
    Call LoadReconData(1, lst_Journal, 1)
ElseIf opt_GJ.Value = True Then
    opt_Statement.Value = True
    Call LoadReconData(4, lst_Journal, 1)
End If
End Sub
Private Sub LoadReconData(ByVal TYP As Integer, ByVal lstview As ListView, ByVal What As String)
Dim rec As New ADODB.Recordset
On Error GoTo bad
Dim x As Long
Dim y
lstview.ListItems.Clear
Set rec = opndbaseFMIS.Execute("EXECUTE [fmis].[dbo].[MPproc_LoadJournaForBankRecon] @fundtype = '" & cmb_fundtype.Text & "',@BankID = '" & cmb_BankName.ItemData(cmb_BankName.ListIndex) & "',@Accountno = '" & cmb_Accountno.Text & "',@month = '" & DTPicker3.Month & "',@year = '" & DTPicker3.Year & "',@transtype = " & TYP & ",@what = " & What & "")
If rec.RecordCount > 0 Then
    For x = 1 To rec.RecordCount
        'DoEvents
        Set y = lstview.ListItems.Add(, , Format(rec!Date_, "mm/dd/yyyy"))
        y.SubItems(1) = Trim(rec!Particular)
        y.SubItems(2) = Trim(rec!checkno)
        y.SubItems(3) = Format(rec!amount, "#,##0.00")
        y.SubItems(4) = rec!id
        rec.MoveNext
    Next x
End If
rec.Close
Exit Sub
bad:
If What = 1 Then
    If err.Number = 3704 Then
    MsgBox "Please Iditify the Transaction type...", vbInformation, "System Message"
    End If
End If
End Sub
'Private Sub LoadMatch(ByVal TYP As Integer, ByVal lstview As ListView, ByVal What As String)
'Dim rec As New ADODB.Recordset
'On Error GoTo bad
'Dim x As Long
'Dim y
'lstview.ListItems.Clear
'Set rec = opndbaseFMIS.Execute("EXECUTE [fmis].[dbo].[MPproc_LoadJournaForBankRecon] @fundtype = '" & cmb_FundType.Text & "',@BankID = '" & cmb_BankName.ItemData(cmb_BankName.ListIndex) & "',@Accountno = '" & cmb_Accountno.Text & "',@month = '" & DTPicker3.Month & "',@year = '" & DTPicker3.Year & "',@transtype = " & TYP & ",@what = " & What & "")
'If rec.RecordCount > 0 Then
'    For x = 1 To rec.RecordCount
'        'DoEvents
'                    Set z = ListView3.ListItems.Add(, , lst_Journal.ListItems(x).Text)
'                        z.SubItems(1) = lst_Journal.ListItems(x).SubItems(1)
'                        z.SubItems(2) = lst_Journal.ListItems(x).SubItems(2)
'                        z.SubItems(3) = lst_Journal.ListItems(x).SubItems(3)
'                        z.SubItems(4) = lst_Journal.ListItems(x).SubItems(4)
'
'                        z.SubItems(6) = lst_Bank.ListItems(y).Text
'                        z.SubItems(7) = lst_Bank.ListItems(y).SubItems(1)
'                        z.SubItems(8) = lst_Bank.ListItems(y).SubItems(2)
'                        z.SubItems(9) = lst_Bank.ListItems(y).SubItems(3)
'                        z.SubItems(10) = lst_Bank.ListItems(y).SubItems(4)
'        rec.MoveNext
'    Next x
'End If
'rec.Close
'Exit Sub
'bad:
'If What = 1 Then
'    If err.Number = 3704 Then
'    MsgBox "Please Iditify the Transaction type...", vbInformation, "System Message"
'    End If
'End If
'End Sub
Private Sub Opt_GJ_Click()
LoadOPTBook
End Sub

Private Sub opt_Statement_Click()
LoadOPTBank
End Sub

Private Sub Opt_Withdrawals_Click()
LoadOPTBank
End Sub

Private Sub Option4_Click()

End Sub
Private Sub toolbarSTat(ByVal Stat As String)

If Stat = "New" Then
    Toolbar1.Buttons(5).Caption = "&Save" 'save
    Toolbar1.Buttons(5).Enabled = True
    Toolbar1.Buttons(11).Enabled = True ' Cancel
    
    Toolbar1.Buttons(3).Enabled = False ' Edit
    Toolbar1.Buttons(7).Enabled = False 'match
    Toolbar1.Buttons(9).Enabled = False 'delete
    
    fme_Details(1).Visible = True
    dt_checkdate.Value = Now
    txtAmount.Text = ""
    txtCheckno.Text = ""
    txtdescription.Text = ""
    txtID.Text = ""
ElseIf Stat = "Edit" Then
    Toolbar1.Buttons(5).Caption = "&Update"
    Toolbar1.Buttons(5).Enabled = True
    Toolbar1.Buttons(11).Enabled = True
    
    Toolbar1.Buttons(3).Enabled = True
    Toolbar1.Buttons(7).Enabled = False
    Toolbar1.Buttons(9).Enabled = True
    fme_Details(1).Visible = True
ElseIf Stat = "Open" Then
    Toolbar1.Buttons(5).Enabled = False
    Toolbar1.Buttons(11).Enabled = True
    
    Toolbar1.Buttons(3).Enabled = True
    Toolbar1.Buttons(7).Enabled = True
    Toolbar1.Buttons(9).Enabled = True
ElseIf Stat = "Cancel" Then
    Toolbar1.Buttons(5).Caption = "&Save" 'save
    Toolbar1.Buttons(5).Enabled = False
    Toolbar1.Buttons(11).Enabled = True ' Cancel
    
    Toolbar1.Buttons(3).Enabled = False ' Edit
    Toolbar1.Buttons(7).Enabled = False 'match
    Toolbar1.Buttons(9).Enabled = False 'delete
    fme_Details(1).Visible = False
    dt_checkdate.Value = Now
    txtAmount.Text = ""
    txtCheckno.Text = ""
    txtdescription.Text = ""
    txtID.Text = ""
End If
End Sub

Private Sub Option1_Click()

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim t As Integer
If opt_Check.Value = True Then: t = opt_Check.Tag
If opt_CR.Value = True Then: t = opt_CR.Tag
If opt_GJ.Value = True Then: t = opt_GJ.Tag
Select Case Button:
    Case "&New":
                Call toolbarSTat("New")
     Case "&Edit":
                With lst_Bank
                    dt_checkdate.Value = .SelectedItem.Text
                    txtAmount.Text = .SelectedItem.SubItems(3)
                    txtCheckno.Text = .SelectedItem.SubItems(2)
                    txtdescription.Text = .SelectedItem.SubItems(1)
                    txtID.Text = .SelectedItem.SubItems(4)
                End With
                Call toolbarSTat("Edit")
    Case "&Save":
                If CheckEntry = True Then
                    If MsgBox("Are you sure you want to Save this Entry?", vbQuestion + vbYesNo) = vbYes Then
                        opndbaseFMIS.Execute ("INSERT INTO fmis.dbo.tblAMIS_BankReconCiliation(Fundtype,BankID,BankAccountno,Checkdate,[Description],Checkno,Amount,month_,Year_,datetimeentered,UserID,transtype,actioncode) " & _
                        " VALUES  ('" & cmb_fundtype.Text & "','" & cmb_BankName.ItemData(cmb_BankName.ListIndex) & "','" & cmb_Accountno.Text & "','" & dt_checkdate.Value & "','" & txtdescription.Text & "','" & txtCheckno.Text & "','" & txtAmount.Text & "','" & DTPicker3.Month & "','" & DTPicker3.Year & "','" & Now & "','" & ActiveUserID & "'," & t & ",1)")
                        Call LoadReconData(2, lst_Bank, 2)
                    End If
                End If
    Case "&Update":
                If MsgBox("Are you sure you want to Update this Entry?", vbQuestion + vbYesNo) = vbYes Then
                    opndbaseFMIS.Execute ("Update [tblAMIS_BankReconCiliation] set actioncode = 2,userid = '" & Trim(ActiveUserID) & "',datetimeentered = '" & Now & "' where trnno = " & txtID.Text & "")
                        opndbaseFMIS.Execute ("INSERT INTO fmis.dbo.tblAMIS_BankReconCiliation(Fundtype,BankID,BankAccountno,Checkdate,[Description],Checkno,Amount,month_,Year_,datetimeentered,UserID,transtype,actioncode) " & _
                        " VALUES  ('" & cmb_fundtype.Text & "','" & cmb_BankName.ItemData(cmb_BankName.ListIndex) & "','" & cmb_Accountno.Text & "','" & dt_checkdate.Value & "','" & txtdescription.Text & "','" & txtCheckno.Text & "','" & txtAmount.Text & "','" & DTPicker3.Month & "','" & DTPicker3.Year & "','" & Now & "','" & ActiveUserID & "'," & t & ",1)")
                        Call LoadReconData(2, lst_Bank, 2)
                        Call toolbarSTat("Cancel")
                End If
    Case "&Match":
                If MsgBox("Are you sure you want to System Matching?", vbQuestion + vbYesNo) = vbYes Then
                    Call MatchTrans
                End If
    Case "&Delete":
                If MsgBox("Are you sure you want to delete this Entry?", vbQuestion + vbYesNo) = vbYes Then
                   opndbaseFMIS.Execute ("Update [tblAMIS_BankReconCiliation] set actioncode = 3,userid = '" & Trim(ActiveUserID) & "',datetimeentered = '" & Now & "' where trnno = " & lst_Bank.SelectedItem.SubItems(4) & "")
                   Call LoadReconData(2, lst_Bank, 2)
                End If
    Case "&Cancel":
                If MsgBox("Are you sure you want to Cancel the Entry?", vbQuestion + vbYesNo, "System Confirmation") = vbYes Then
                   Call toolbarSTat("Cancel")
                End If
    End Select
End Sub
Private Sub MatchTrans()
Dim x, y As Long
Dim z
For y = 1 To lst_Bank.ListItems.Count
        With lst_Journal
            For x = 1 To .ListItems.Count
                If lst_Bank.ListItems(y).SubItems(2) = .ListItems(x).SubItems(2) And lst_Bank.ListItems(y).SubItems(3) = .ListItems(x).SubItems(3) Then
                    .HideSelection = False
                    .ListItems(x).Selected = True
                    opndbaseFMIS.Execute "Update [tblAMIS_FinalJEV] set bankmatch = 1 where trnno = " & .ListItems(x).SubItems(4) & " and actioncode = 1"
                    opndbaseFMIS.Execute "Update [tblAMIS_BankReconCiliation] set havematch = 1 where trnno = " & lst_Bank.ListItems(y).SubItems(4) & " and actioncode = 1"
                    Set z = ListView3.ListItems.Add(, , .ListItems(x).Text)
                        z.SubItems(1) = .ListItems(x).SubItems(1)
                        z.SubItems(2) = .ListItems(x).SubItems(2)
                        z.SubItems(3) = .ListItems(x).SubItems(3)
                        z.SubItems(4) = .ListItems(x).SubItems(4)
                        
                        z.SubItems(6) = lst_Bank.ListItems(y).Text
                        z.SubItems(7) = lst_Bank.ListItems(y).SubItems(1)
                        z.SubItems(8) = lst_Bank.ListItems(y).SubItems(2)
                        z.SubItems(9) = lst_Bank.ListItems(y).SubItems(3)
                        z.SubItems(10) = lst_Bank.ListItems(y).SubItems(4)
                        DoEvents
                    Exit For
                End If
            Next x
        End With
Next y
End Sub
Private Function CheckEntry() As Boolean
Dim rec As New ADODB.Recordset
CheckEntry = False
If txtAmount.Text <> "" And txtCheckno.Text <> "" Then
    CheckEntry = True
Else
    MsgBox "Please Specify the checkno and Amount..!", vbInformation, "System Message"
    Exit Function
End If
Set rec = opndbaseFMIS.Execute("Select [Checkno] from [tblAMIS_BankReconCiliation] where checkno = '" & txtCheckno.Text & "' and actioncode = 1")
    If rec.RecordCount > 0 Then
        CheckEntry = False
        MsgBox "Checkno Already Exist on the Database...!", vbInformation, "System Message"
        Exit Function
    Else
        CheckEntry = True
    End If
rec.Close
End Function

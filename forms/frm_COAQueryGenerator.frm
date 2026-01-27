VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frm_COAQueryGenerator 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Query Maintenance"
   ClientHeight    =   9810
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11730
   Icon            =   "frm_COAQueryGenerator.frx":0000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9810
   ScaleWidth      =   11730
   Begin TabDlg.SSTab SSTab1 
      Height          =   9615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   16960
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "JEV Maker"
      TabPicture(0)   =   "frm_COAQueryGenerator.frx":076A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label8"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label15"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lvButtons_H2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lvButtons_H1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdupdate"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lvButtons_H3"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Frame3"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Frame2"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Frame1"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Text1"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Combo2"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).ControlCount=   11
      TabCaption(1)   =   "JEV Maker Classification"
      TabPicture(1)   =   "frm_COAQueryGenerator.frx":0786
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label7"
      Tab(1).Control(1)=   "Label9"
      Tab(1).Control(2)=   "ListView2"
      Tab(1).Control(3)=   "cmbtype"
      Tab(1).Control(4)=   "Check1"
      Tab(1).Control(5)=   "Text2"
      Tab(1).Control(6)=   "Frame4"
      Tab(1).ControlCount=   7
      TabCaption(2)   =   "Sub Account Import Management"
      TabPicture(2)   =   "frm_COAQueryGenerator.frx":07A2
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame5"
      Tab(2).ControlCount=   1
      Begin VB.ComboBox Combo2 
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
         ItemData        =   "frm_COAQueryGenerator.frx":07BE
         Left            =   3840
         List            =   "frm_COAQueryGenerator.frx":07CE
         Style           =   2  'Dropdown List
         TabIndex        =   49
         Top             =   4740
         Width           =   2175
      End
      Begin VB.Frame Frame5 
         Height          =   1815
         Left            =   -74880
         TabIndex        =   39
         Top             =   360
         Width           =   6375
         Begin VB.TextBox Text4 
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
            Left            =   1800
            TabIndex        =   41
            Top             =   720
            Width           =   3975
         End
         Begin VB.TextBox Text3 
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
            Left            =   1800
            Locked          =   -1  'True
            TabIndex        =   40
            Top             =   240
            Width           =   3015
         End
         Begin lvButton.lvButtons_H lvButtons_H9 
            Height          =   375
            Left            =   5880
            TabIndex        =   42
            Top             =   720
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
         Begin lvButton.lvButtons_H lvButtons_H10 
            Height          =   375
            Left            =   4440
            TabIndex        =   43
            Top             =   1320
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   661
            Caption         =   "Save"
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
            Mode            =   0
            Value           =   0   'False
            cBack           =   -2147483633
         End
         Begin lvButton.lvButtons_H lvButtons_H11 
            Height          =   375
            Left            =   5160
            TabIndex        =   44
            Top             =   1320
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   661
            Caption         =   "Del"
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
            Mode            =   0
            Value           =   0   'False
            cBack           =   -2147483633
         End
         Begin VB.Label Label13 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Link Description:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   0
            TabIndex        =   46
            Top             =   600
            Width           =   1575
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "AccountName"
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
            Left            =   -240
            TabIndex        =   45
            Top             =   240
            Width           =   1935
         End
      End
      Begin VB.Frame Frame4 
         Height          =   1815
         Left            =   -71160
         TabIndex        =   30
         Top             =   360
         Width           =   7455
         Begin VB.TextBox txtdescription1 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   900
            Left            =   1680
            Locked          =   -1  'True
            TabIndex        =   34
            Top             =   240
            Width           =   5655
         End
         Begin VB.TextBox txtaccountcode1 
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
            Left            =   1680
            TabIndex        =   31
            Top             =   1320
            Width           =   2895
         End
         Begin lvButton.lvButtons_H lvButtons_H5 
            Height          =   375
            Left            =   4680
            TabIndex        =   32
            Top             =   1320
            Width           =   495
            _ExtentX        =   873
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
         Begin lvButton.lvButtons_H lvButtons_H6 
            Height          =   375
            Left            =   5280
            TabIndex        =   36
            Top             =   1320
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   661
            Caption         =   "Add"
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
            Mode            =   0
            Value           =   0   'False
            cBack           =   -2147483633
         End
         Begin lvButton.lvButtons_H lvButtons_H8 
            Height          =   375
            Left            =   6720
            TabIndex        =   37
            Top             =   1320
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   661
            Caption         =   "Del"
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
            Mode            =   0
            Value           =   0   'False
            cBack           =   -2147483633
         End
         Begin lvButton.lvButtons_H lvButtons_H7 
            Height          =   375
            Left            =   6000
            TabIndex        =   38
            Top             =   1320
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   661
            Caption         =   "Clear"
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
            Mode            =   0
            Value           =   0   'False
            cBack           =   -2147483633
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Description:"
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
            Left            =   -360
            TabIndex        =   35
            Top             =   240
            Width           =   1935
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Accountcode:"
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
            Left            =   -360
            TabIndex        =   33
            Top             =   1320
            Width           =   1935
         End
      End
      Begin VB.TextBox Text2 
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
         Left            =   -74880
         Locked          =   -1  'True
         TabIndex        =   28
         Top             =   1800
         Width           =   3615
      End
      Begin VB.TextBox Text1 
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
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   26
         Top             =   4740
         Width           =   2895
      End
      Begin VB.CheckBox Check1 
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
         Left            =   -74880
         TabIndex        =   24
         Top             =   1200
         Width           =   1455
      End
      Begin VB.ComboBox cmbtype 
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
         Left            =   -74880
         TabIndex        =   22
         Text            =   "Combo1"
         Top             =   720
         Width           =   3615
      End
      Begin VB.Frame Frame1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4095
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   11175
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
            ItemData        =   "frm_COAQueryGenerator.frx":07FB
            Left            =   8160
            List            =   "frm_COAQueryGenerator.frx":080B
            Style           =   2  'Dropdown List
            TabIndex        =   47
            Top             =   1560
            Width           =   2175
         End
         Begin VB.TextBox txtquery 
            Height          =   1620
            IMEMode         =   3  'DISABLE
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   13
            Top             =   2400
            Width           =   10935
         End
         Begin VB.TextBox txtdescription 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1020
            Left            =   1800
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   12
            Top             =   360
            Width           =   9135
         End
         Begin VB.TextBox txtaccountcode 
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
            Left            =   1800
            Locked          =   -1  'True
            TabIndex        =   11
            Top             =   1560
            Width           =   3735
         End
         Begin lvButton.lvButtons_H lvButtons_H4 
            Height          =   375
            Left            =   5760
            TabIndex        =   14
            Top             =   1560
            Width           =   495
            _ExtentX        =   873
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
         Begin VB.Label Label14 
            Alignment       =   1  'Right Justify
            Caption         =   "Type of Employee"
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
            Left            =   6240
            TabIndex        =   48
            Top             =   1560
            Width           =   1815
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            Caption         =   "Query:"
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
            TabIndex        =   17
            Top             =   2040
            Width           =   735
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Description:"
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
            Left            =   -120
            TabIndex        =   16
            Top             =   360
            Width           =   1815
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Accountcode:"
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
            Left            =   -240
            TabIndex        =   15
            Top             =   1560
            Width           =   1935
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Entries"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4335
         Left            =   120
         TabIndex        =   8
         Top             =   5160
         Width           =   11175
         Begin MSComctlLib.ListView ListView1 
            Height          =   3975
            Left            =   120
            TabIndex        =   9
            Top             =   240
            Width           =   10935
            _ExtentX        =   19288
            _ExtentY        =   7011
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
            NumItems        =   5
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "ID"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Description"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Accountcode"
               Object.Width           =   5292
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Query"
               Object.Width           =   14111
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "Type"
               Object.Width           =   1235
            EndProperty
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Update Statement Details"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1935
         Left            =   0
         TabIndex        =   1
         Top             =   10440
         Width           =   11175
         Begin VB.TextBox txtconditions 
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
            Left            =   2280
            TabIndex        =   4
            Top             =   1320
            Width           =   4455
         End
         Begin VB.TextBox txtcolumns 
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
            Left            =   2280
            TabIndex        =   3
            Top             =   840
            Width           =   4455
         End
         Begin VB.TextBox txttable 
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
            Left            =   2280
            TabIndex        =   2
            Top             =   360
            Width           =   4455
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            Caption         =   "Column:"
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
            TabIndex        =   7
            Top             =   840
            Width           =   2055
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            Caption         =   "Condition:"
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
            TabIndex        =   6
            Top             =   1320
            Width           =   2055
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            Caption         =   "Table:"
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
            TabIndex        =   5
            Top             =   360
            Width           =   2055
         End
      End
      Begin lvButton.lvButtons_H lvButtons_H3 
         Height          =   495
         Left            =   7440
         TabIndex        =   18
         Top             =   4560
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
         Caption         =   "&Save"
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
         cBhover         =   16512
         cGradient       =   16512
         Gradient        =   3
         CapStyle        =   1
         Mode            =   0
         Value           =   0   'False
         Image           =   "frm_COAQueryGenerator.frx":0838
         ImgSize         =   24
         cBack           =   -2147483633
      End
      Begin lvButton.lvButtons_H cmdupdate 
         Height          =   495
         Left            =   6120
         TabIndex        =   19
         Top             =   4560
         Width           =   1215
         _ExtentX        =   2143
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
         cBhover         =   16512
         cGradient       =   16512
         Gradient        =   3
         CapStyle        =   1
         Mode            =   0
         Value           =   0   'False
         Image           =   "frm_COAQueryGenerator.frx":0B8A
         ImgSize         =   24
         cBack           =   -2147483633
      End
      Begin lvButton.lvButtons_H lvButtons_H1 
         Height          =   495
         Left            =   8760
         TabIndex        =   20
         Top             =   4560
         Width           =   1215
         _ExtentX        =   2143
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
         cBhover         =   16512
         cGradient       =   16512
         Gradient        =   3
         CapStyle        =   1
         Mode            =   0
         Value           =   0   'False
         Image           =   "frm_COAQueryGenerator.frx":17DC
         ImgSize         =   24
         cBack           =   -2147483633
      End
      Begin lvButton.lvButtons_H lvButtons_H2 
         Height          =   495
         Left            =   10080
         TabIndex        =   21
         Top             =   4560
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
         Caption         =   "&Close"
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
         cBhover         =   16512
         cGradient       =   16512
         Gradient        =   3
         CapStyle        =   1
         Mode            =   0
         Value           =   0   'False
         Image           =   "frm_COAQueryGenerator.frx":52E6
         ImgSize         =   24
         cBack           =   -2147483633
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   7095
         Left            =   -74880
         TabIndex        =   25
         Top             =   2400
         Width           =   11295
         _ExtentX        =   19923
         _ExtentY        =   12515
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
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "ID"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Accountcode"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Description"
            Object.Width           =   14111
         EndProperty
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         Caption         =   "Type of Employee"
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
         Left            =   3600
         TabIndex        =   50
         Top             =   4440
         Width           =   1815
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Find:"
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
         Left            =   -75120
         TabIndex        =   29
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Find:"
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
         Left            =   0
         TabIndex        =   27
         Top             =   4440
         Width           =   735
      End
      Begin VB.Label Label7 
         Caption         =   "Transaction Type"
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
         Left            =   -74880
         TabIndex        =   23
         Top             =   360
         Width           =   1815
      End
   End
End
Attribute VB_Name = "frm_COAQueryGenerator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Trnno As Integer
Public IsOK As Boolean

Private Sub Check1_Click()
LoaddataforEntry
End Sub

Private Sub cmbtype_Click()
LoaddataforEntry
End Sub

Private Sub cmdupdate_Click()
txtaccountcode.Text = ""
txtdescription.Text = ""
txtquery.Text = ""
Trnno = 0
Loaddata
End Sub

Private Sub Combo2_Click()
Loaddata
End Sub

Private Sub Form_Load()
Loaddata
Call LoadAccntgTransType(cmbtype)
LoaddataforEntry
End Sub

Private Sub ListView1_Click()
On Error Resume Next
Trnno = 0
txtaccountcode.Text = ""
txtdescription.Text = ""
txtquery.Text = ""
If ListView1.ListItems.Count <> 0 Then
    With ListView1
        Trnno = .SelectedItem.Text
        txtaccountcode.Text = Trim(.SelectedItem.ListSubItems(2).Text)
        txtdescription.Text = Trim(.SelectedItem.ListSubItems(1).Text)
        txtquery.Text = Trim(.SelectedItem.ListSubItems(3).Text)
        Combo1.Text = Trim(.SelectedItem.ListSubItems(4).Text)
    End With
End If
End Sub

Private Sub ListView2_Click()
txtdescription1.Text = ListView2.SelectedItem.ListSubItems(2).Text
txtaccountcode1.Text = ListView2.SelectedItem.ListSubItems(1).Text
End Sub

Private Sub ListView2_KeyUp(KeyCode As Integer, Shift As Integer)
Call ListView2_Click
End Sub

Private Sub lvButtons_H1_Click()
If MsgBox("Are you sure do you want to Delete this entry?", vbCritical + vbYesNo, "System Confirmation") = vbYes Then
    opndbaseFMIS.Execute "delete from tblAMIS_Qrygenerator4COA where trnno = " & Trnno & ""
    MsgBox "Delete Successfully...!", vbInformation, "System Message"
End If
End Sub

Private Sub lvButtons_H2_Click()
Unload Me
End Sub

Private Sub lvButtons_H3_Click()
Dim rec As New ADODB.Recordset
If Trim(txtaccountcode.Text) = "" Or txtdescription.Text = "" Or txtquery.Text = "" Or Combo1.Text = "" Then
    MsgBox "Complete the Fields to Proceed the transaction", vbInformation, "System Message"
    Exit Sub
End If

rec.Open "Select * from tblAMIS_Qrygenerator4COA where trnno = " & Trnno & "", opndbaseFMIS, adOpenStatic, adLockOptimistic
If rec.RecordCount > 0 Then
    If MsgBox("Are you sure do you want to update the Query?", vbInformation + vbYesNo, "System Confirmation") = vbYes Then
        opndbaseFMIS.Execute "Update tblAMIS_Qrygenerator4COA set Acountcode = '" & Replace(Trim(txtaccountcode.Text), "'", "''") & "',description = '" & Replace(Trim(txtdescription.Text), "'", "''") & "',Query = '" & Replace(Trim(txtquery.Text), "'", "''") & "',actioncode = 1,userid = '" & ActiveUserID & "',datetimeentered = '" & Now & "',type = '" & Combo1.Text & "' where trnno = " & Trnno & ""
        MsgBox "Update Successfully...!", vbInformation, "System Message"
    End If
Else
    If MsgBox("Are you Sure do you want to Save?", vbInformation + vbYesNo, "System Confirmation") = vbYes Then
       opndbaseFMIS.Execute "insert into  tblAMIS_Qrygenerator4COA (Acountcode,description,Query,Actioncode,userid,datetimeentered,type) values ('" & Replace(Trim(txtaccountcode.Text), "'", "''") & "','" & Replace(Trim(txtdescription.Text), "'", "''") & "','" & Replace(Trim(txtquery.Text), "'", "''") & "',1,'" & ActiveUserID & "','" & Now & "','" & Combo1.Text & "')"
       MsgBox "Save Successfully...!", vbInformation, "System Message"
    End If
End If
Loaddata
txtaccountcode.Text = ""
txtdescription.Text = ""
txtquery.Text = ""

rec.Close
Set rec = Nothing
End Sub
Private Function Loaddata()
Dim rec As New ADODB.Recordset
Dim x As Integer
Dim z
rec.Open "Select * from tblAMIS_Qrygenerator4COA where type like '" & Combo2.Text & "%' order by acountcode", opndbaseFMIS, adOpenStatic, adLockOptimistic
    If rec.RecordCount > 0 Then
    ListView1.ListItems.Clear
    With ListView1
        For x = 1 To rec.RecordCount
        
            Set z = .ListItems.Add(, , rec!Trnno)
                z.SubItems(1) = Trim(IIf(IsNull(rec!Description), "", rec!Description))
                z.SubItems(2) = Trim(rec!acountcode)
                z.SubItems(3) = Trim(rec!query)
                z.SubItems(4) = Trim(IIf(IsNull(rec!Type), "", rec!Type))
            rec.MoveNext
        Next x
    End With
    End If
rec.Close
Set rec = Nothing
End Function

Private Sub lvButtons_H4_Click()

    With frmforCOA
    .nme = txtdescription
    'If isOK = True Then
    .accntcode = txtaccountcode.Text
    'End If
    .Trnno = Trnno
    Set .frm = Me
    .Show 1
    End With

End Sub

Private Sub lvButtons_H5_Click()
With frmforCOA
    .nme = txtdescription
    'If isOK = True Then
    .accntcode = txtaccountcode.Text
    'End If
    .Trnno = Trnno
    Set .frm = Me
    .Show 1
    End With
    If SSTab1.Tab = 1 Then
    txtaccountcode1.Text = txtaccountcode.Text
    txtaccountcode.Text = ""
    txtdescription1.Text = txtdescription.Text
    txtdescription.Text = ""
    End If
End Sub

Private Sub lvButtons_H6_Click()
Dim rec As New ADODB.Recordset
Dim x As Integer
x = 0
If Check1.Value = 1 Then
x = 1
End If
rec.Open "Select accountcode from tblAMIS_AccountsTypeAndEntry where accountcode = '" & txtaccountcode1.Text & "' and transtype = '" & cmbtype.Text & "' and debitcredit = " & x & "", opndbaseFMIS, adOpenStatic, adLockReadOnly
If rec.RecordCount > 0 Then
    MsgBox "Already Exist In the Database..!Please Check it..", vbInformation + vbCritical, "System Message"
Else
    If MsgBox("Are you sure do you want to add?", vbInformation + vbYesNo, "System Message") = vbYes Then
     opndbaseFMIS.Execute "Insert into tblAMIS_AccountsTypeAndEntry (accountcode,transtype,debitcredit) values ('" & txtaccountcode1.Text & "','" & cmbtype.Text & "'," & x & ")"
     LoaddataforEntry
    End If
End If
rec.Close
End Sub

Private Sub lvButtons_H7_Click()
txtaccountcode1.Text = ""
txtdescription1.Text = ""
End Sub

Private Sub lvButtons_H8_Click()
Dim x As Integer
x = 0
If Check1.Value = 1 Then
x = 1
End If
If MsgBox("Are you sure do you want to Delete?", vbInformation + vbYesNo, "System Message") Then
 opndbaseFMIS.Execute "Delete from tblAMIS_AccountsTypeAndEntry where accountcode = '" & txtaccountcode1.Text & "' and transtype = '" & cmbtype.Text & "' and debitcredit = " & x & ""
 LoaddataforEntry
End If
End Sub

Private Function LoaddataforEntry()
Dim rec As New ADODB.Recordset
Dim x, q As Integer
Dim z
q = 0
If Check1.Value = 1 Then
q = 1
End If
ListView2.ListItems.Clear
rec.Open "Select * from tblAMIS_AccountsTypeAndEntry where transtype = '" & Trim(cmbtype.Text) & "'  and debitcredit = " & q & " order by accountcode ", opndbaseFMIS, adOpenStatic, adLockOptimistic
    If rec.RecordCount > 0 Then
    ListView2.ListItems.Clear
    With ListView2
        For x = 1 To rec.RecordCount
        
            Set z = .ListItems.Add(, , rec!Trnno)
                z.SubItems(1) = Trim(IIf(IsNull(rec!accountcode), "", rec!accountcode))
                z.SubItems(2) = getQueryDescription(rec!accountcode)
            rec.MoveNext
        Next x
    End With
    End If
rec.Close
Set rec = Nothing
End Function


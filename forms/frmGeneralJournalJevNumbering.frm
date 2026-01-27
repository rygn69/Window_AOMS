VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
Object = "{1693405E-2DC9-4248-B52F-4AC9145DA2AF}#1.0#0"; "WinXPC Engine.ocx"
Begin VB.Form frmGeneralJournalJevNumberingXXX 
   Caption         =   "JEV Numbering for General Journal Report"
   ClientHeight    =   9210
   ClientLeft      =   2295
   ClientTop       =   1185
   ClientWidth     =   14925
   Icon            =   "frmGeneralJournalJevNumbering.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   9210
   ScaleWidth      =   14925
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "&Remove"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2640
      TabIndex        =   35
      Top             =   3000
      Width           =   1035
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00000000&
      Caption         =   "Select All"
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
      Left            =   2590
      MaskColor       =   &H00000000&
      TabIndex        =   34
      Top             =   3360
      Width           =   1215
   End
   Begin VB.TextBox txtSearch 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   80
      TabIndex        =   31
      ToolTipText     =   "Type only number then Enter (""FMISNo-00"") will apear"
      Top             =   990
      Width           =   2280
   End
   Begin WinXPC_Engine.WindowsXPC WindowsXPC1 
      Left            =   6960
      Top             =   6960
      _ExtentX        =   6588
      _ExtentY        =   1085
      ColorScheme     =   4
   End
   Begin VB.Frame Frame5 
      Height          =   2250
      Left            =   2535
      TabIndex        =   6
      Top             =   525
      Width           =   12315
      Begin VB.CommandButton Command2 
         Caption         =   "&Addnew"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   885
         Left            =   240
         TabIndex        =   33
         Top             =   720
         Width           =   1275
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Edit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   900
         Left            =   4680
         TabIndex        =   30
         Top             =   720
         Width           =   1275
      End
      Begin VB.Frame Frame3 
         Caption         =   "Provide the Fields Below"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1725
         Left            =   6120
         TabIndex        =   20
         Top             =   240
         Width           =   5970
         Begin VB.ComboBox Cmb_CnNo 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   2220
            Style           =   1  'Simple Combo
            TabIndex        =   28
            Top             =   270
            Width           =   2565
         End
         Begin VB.ComboBox cmb_AccountName 
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
            Left            =   4950
            TabIndex        =   21
            Top             =   1755
            Width           =   2085
         End
         Begin MSComCtl2.DTPicker DTPCNdate 
            CausesValidation=   0   'False
            Height          =   390
            Left            =   2220
            TabIndex        =   22
            Top             =   720
            Width           =   2565
            _ExtentX        =   4524
            _ExtentY        =   688
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "mm/dd/yyyy"
            Format          =   151388161
            UpDown          =   -1  'True
            CurrentDate     =   38240
         End
         Begin MSComCtl2.DTPicker DTPRdate 
            CausesValidation=   0   'False
            Height          =   390
            Left            =   2220
            TabIndex        =   23
            Top             =   1200
            Width           =   2565
            _ExtentX        =   4524
            _ExtentY        =   688
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "MM/dd/yyyy"
            Format          =   151388161
            UpDown          =   -1  'True
            CurrentDate     =   40583
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Date Received"
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
            Left            =   720
            TabIndex        =   27
            Top             =   1275
            Width           =   1725
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Account Name"
            Height          =   195
            Left            =   3600
            TabIndex        =   26
            Top             =   1875
            Width           =   1065
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "CN Date:"
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
            Left            =   1320
            TabIndex        =   25
            Top             =   795
            Width           =   1170
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "CN No:"
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
            Left            =   1440
            TabIndex        =   24
            Top             =   375
            Width           =   1005
         End
      End
      Begin VB.CommandButton cmd_Delete 
         Caption         =   "&Delete"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   885
         Left            =   3120
         TabIndex        =   18
         Top             =   720
         Width           =   1395
      End
      Begin VB.CommandButton cmd_Save 
         Caption         =   "&Save"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   885
         Left            =   1680
         TabIndex        =   8
         Top             =   720
         Width           =   1275
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   945
         Left            =   3960
         TabIndex        =   7
         Top             =   3900
         Width           =   7800
         _ExtentX        =   13758
         _ExtentY        =   1667
         View            =   3
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Check Number"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   1
            Text            =   "Check Amount"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Check Date"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "RCI No."
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Shape Shape3 
         BackColor       =   &H00000000&
         BackStyle       =   1  'Opaque
         Height          =   195
         Left            =   120
         Top             =   1485
         Width           =   3930
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Check Amount:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   4440
         TabIndex        =   17
         Top             =   2760
         Width           =   1575
      End
      Begin VB.Label lbl_CheckAmount 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "000.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   5880
         TabIndex        =   16
         Top             =   2760
         Width           =   1935
      End
      Begin VB.Label lbl_total 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "000.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   5880
         TabIndex        =   15
         Top             =   2640
         Width           =   1935
      End
      Begin VB.Label lbl_TotalLiqAmount 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "000.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   9840
         TabIndex        =   14
         Top             =   2640
         Width           =   1455
      End
      Begin VB.Label lbl_LackingAmount 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "000.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   9840
         TabIndex        =   13
         Top             =   2880
         Width           =   1575
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Cash Advance :"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   4440
         TabIndex        =   12
         Top             =   3120
         Width           =   1575
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Liquidated Amount :"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   7920
         TabIndex        =   11
         Top             =   2760
         Width           =   1815
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Lacking Amount :"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   7920
         TabIndex        =   10
         Top             =   3120
         Width           =   1695
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00000000&
         BackStyle       =   1  'Opaque
         Height          =   1905
         Left            =   105
         Top             =   180
         Width           =   12120
      End
   End
   Begin VB.TextBox txt_RecordID 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   8880
      TabIndex        =   4
      ToolTipText     =   "Type only number then Enter (""FMISNo-00"") will apear"
      Top             =   3000
      Width           =   5760
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   2400
      Top             =   4035
   End
   Begin VB.ListBox List2 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5550
      Left            =   75
      TabIndex        =   2
      Top             =   1965
      Width           =   2265
   End
   Begin VB.Frame Frame6 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Dvnos"
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   60
      TabIndex        =   1
      Top             =   1650
      Width           =   2280
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   600
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   14925
      _ExtentX        =   26326
      _ExtentY        =   1058
      ButtonWidth     =   2117
      ButtonHeight    =   1005
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Print Report"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "slash"
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Close"
            ImageIndex      =   5
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1560
      Top             =   7800
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGeneralJournalJevNumbering.frx":0E42
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGeneralJournalJevNumbering.frx":1EC4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGeneralJournalJevNumbering.frx":3FFE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGeneralJournalJevNumbering.frx":5280
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGeneralJournalJevNumbering.frx":810A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView ListView2 
      Height          =   5175
      Left            =   2520
      TabIndex        =   29
      Top             =   3840
      Width           =   12255
      _ExtentX        =   21616
      _ExtentY        =   9128
      View            =   3
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
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Dvno"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "OBR NO."
         Object.Width           =   3351
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Particular"
         Object.Width           =   7056
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Claimant"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Amount"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "dvno"
         Object.Width           =   0
      EndProperty
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Search Credit Notice No.:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   120
      TabIndex        =   32
      Top             =   720
      Width           =   1815
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Dvno and press ENTER:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   5160
      TabIndex        =   9
      Top             =   3120
      Width           =   3780
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   825
      Left            =   1680
      Top             =   2880
      Width           =   13065
   End
   Begin VB.Label Label13 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Label13"
      ForeColor       =   &H0000FF00&
      Height          =   420
      Left            =   90
      TabIndex        =   19
      Top             =   1455
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label8"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   105
      TabIndex        =   5
      Top             =   7740
      Width           =   480
   End
   Begin VB.Label Label7 
      Caption         =   "Label7"
      Height          =   210
      Left            =   3540
      TabIndex        =   3
      Top             =   3060
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   9540
      Left            =   -30
      Top             =   585
      Width           =   2445
   End
End
Attribute VB_Name = "frmGeneralJournalJevNumberingXXX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
'Dim tmpAccName As String
'Dim FMISNo As String
'
'
'
'
'
'
'
'
'
'
'
'Private Function GetTotalCheckAmount(ByVal RecordID As String) As Currency
'Dim totalCheck As New ADODB.Recordset
'
'totalCheck.Open " select sum(NetAmount) as Amount from vw_CDCashAdvancedChecks where mixcode='" & RecordID & "'", opndbaseFMIS, adOpenStatic, adLockOptimistic
'    GetTotalCheckAmount = IIf(IsNull(totalCheck!AMOUNT), 0, totalCheck!AMOUNT)
'totalCheck.Close
'Set totalCheck = Nothing
'End Function
'
'
'
'Private Sub cmb_FundType_Change()
'
'End Sub
'
'Private Sub cmb_FundType_Click()
'
'End Sub
'
'
'
'Private Sub cmd_Mass_Click()
'
'End Sub
'
'Private Sub Check1_Click()
'On Error GoTo bad
'Dim x As Integer
'If Check1.Value = 1 Then
'        For x = 1 To ListView2.ListItems.Count
'            ListView2.ListItems(x).Checked = True
'        Next x
'Else
'        For x = 1 To ListView2.ListItems.Count
'            ListView2.ListItems(x).Checked = False
'        Next x
'End If
'Exit Sub
'bad:
'End Sub
'
'Private Sub Command3_Click()
'On Error GoTo bad
'
'Dim x As Integer
'
'For x = 1 To ListView2.ListItems.Count
'    If ListView2.ListItems(x).Checked = True Then
'    ListView2.ListItems.Remove (x)
'    x = x - 1
'    End If
'
'Next x
'Exit Sub
'bad:
'
'
'End Sub
'
''Private Sub cmd_post_Click()
''Dim cc As Integer
''
''If MsgBox("Save JEV Nos.?", vbQuestion + vbYesNo, "System Confirmation") = vbYes Then
''For cc = 1 To MSHFlexGrid1.Rows - 1
''
''        If Len(Trim(MSHFlexGrid1.TextMatrix(cc, 10))) > 0 Then
''            If IsFormatCorrect(MSHFlexGrid1.TextMatrix(cc, 10)) = True Then
''
''                'Updating table from PTO....
''                opndbaseFMIS.Execute "Update tblCMS_CDCashBook set AlreadySaved2JEV=1,DatePostedtoJEV='" & Date & "',PostedtoJEVUserid='" & ActiveUserID & "' where trnno=" & MSHFlexGrid1.TextMatrix(cc, 0) & ""
''
''                'Updating Accounting REcord...
''                opndbaseFMIS.Execute "update tblAMIS_JournalEntry set JEVNo='" & MSHFlexGrid1.TextMatrix(cc, 10) & "',JEVSeriesNo=" & ExtractJEVSNo(MSHFlexGrid1.TextMatrix(cc, 10)) & ",JEVBy='" & ActiveUserID & "',JEVDate='" & Date & "' where DVNo='" & MSHFlexGrid1.TextMatrix(cc, 8) & "'"
''
''            End If
''        End If
''
''Next cc
''MsgBox "Posting to JEV, Successful!", vbInformation, "System Information"
''Command1_Click 'Loading Back Active Cash Disbursement Numbers...
''List2.ListIndex = GetIndex4ListBox(List2, FMISNo)
''End If
''
''End Sub
'
'
'
'
'
'Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'If KeyCode = vbKeyEscape Then
'        Unload Me
'End If
'End Sub
'
'Private Sub Clear()
'Cmb_CnNo.Text = ""
'
'ListView2.ListItems.Clear
'End Sub
'Private Sub Form_Load()
'WindowsXPC1.InitSubClassing
'Me.Top = (Screen.Height - Me.Height) / 2
'Me.Left = (Screen.Width - Me.Width) / 2
'DTPCNdate.Value = Now
'Label8.Caption = ""
'Label13.Caption = ""
'Timer1.Enabled = True
'End Sub
'
'Private Sub Form_Unload(Cancel As Integer)
'WindowsXPC1.EndWinXPCSubClassing
'Set frmCDCashDisbursedReport = Nothing
'End Sub
'
'
'Private Sub LoadCACheck(ByVal RecordID As String)
'Dim opnCACheck As New ADODB.Recordset
'Dim sitem As ListItem
'Dim i As Integer
'
'
'opnCACheck.Open "Select * from vw_CDCashAdvancedChecks where mixcode='" & RecordID & "'", opndbaseFMIS, adOpenStatic, adLockOptimistic
'If opnCACheck.RecordCount <> 0 Then
'    ListView1.ListItems.Clear
'    Do Until opnCACheck.EOF
'            Set sitem = ListView1.ListItems.Add()
'            sitem.Text = opnCACheck!checkno
'            sitem.SubItems(1) = Format(opnCACheck!NetAmount, "###,##0.00")
'            sitem.SubItems(2) = opnCACheck!CheckDate
'            sitem.SubItems(3) = GetRCINoPerCheck(opnCACheck!checkno)
'        opnCACheck.MoveNext
'    Loop
'Else
'    ListView1.ListItems.Clear
'End If
'opnCACheck.Close
'Set opnCACheck = Nothing
'End Sub
'
'Private Sub LoadAccountName(ByVal BankAccntNo As String, ByVal FundType As String, ByVal BankID As String)
'Dim opnAcctname As New ADODB.Recordset
'Dim xx As Long
'
'opnAcctname.Open "Select * from vw_DepositoryBank where BankID='" & BankID & "' and FundType='" & FundType & "' and BankAccountNo='" & BankAccntNo & "' and active=1", opndbaseFMIS, adOpenStatic, adLockOptimistic
'If opnAcctname.RecordCount <> 0 Then
'    cmb_AccountName.Clear
'    Do Until opnAcctname.EOF
'        cmb_AccountName.AddItem (opnAcctname!Accountname)
'        cmb_AccountName.ItemData(xx) = opnAcctname!FmisAccountcode
'        xx = xx + 1
'        opnAcctname.MoveNext
'    Loop
'Else
'    cmb_AccountName.Clear
'End If
'opnAcctname.Close
'Set opnAcctname = Nothing
'End Sub
'
'
'Private Sub LoadSavedReport()
'Dim opnvoucher As New ADODB.Recordset
'Dim cc As Integer
'Dim sql As String
'
'
'sql = " SELECT b.cnno FROM tblAMIS_IncomingDVTrns as a inner join tblAMIS_CreditNotice as b on a.dvno = b.dvno WHERE  b.cnno like '" & txtSearch.Text & "%' AND a.ACTIONCODE = 1 AND b.ACTIONCODE = 1 group by b.cnno,b.datetimeentered order by b.datetimeentered "
'
'opnvoucher.Open sql, opndbaseFMIS, adOpenStatic, adLockOptimistic
'
'
'
'If opnvoucher.RecordCount <> 0 Then
'    List2.Clear
'    Do Until opnvoucher.EOF
'        List2.AddItem (opnvoucher!cnno)
'        'List2.ItemData = (opnvoucher.Fields!trnno)
'        opnvoucher.MoveNext
'    Loop
'Else
'    List2.Clear
'End If
'opnvoucher.Close
'Set opnvoucher = Nothing
'Label8.Caption = List2.ListCount & "Record/s Found"
'End Sub
'
'
'
'Private Sub List2_DblClick()
'Label13.Caption = "Loading Details..."
'Label13.Refresh
'cmd_Save.Enabled = False
'cmd_Delete.Enabled = True
'FMISNo = List2.Text
'ListView2.ListItems.Clear
'Call LoadCNdetails(FMISNo)
'
'
'End Sub
'Private Sub LoadBackBreakdown(ByVal DVNo As String)
'Dim opnvoucher As New ADODB.Recordset
'Dim sql As String
'Dim x
'
'sql = "SELECT * FROM tblAMIS_IncomingDVTrns  WHERE DVNO = '" & DVNo & "' and ACTIONCODE = 1 "
'
''Debug.Print sql
'
'opnvoucher.Open sql, opndbaseFMIS, adOpenStatic, adLockOptimistic
'
'If opnvoucher.RecordCount <> 0 Then
'    If IFalreadyOnDatabase(DVNo) = True Then
'    MsgBox "Dvno Already Saved On the database", vbInformation, "System Message"
'    Else
'        If IFAlready(DVNo) = True Then
'            MsgBox "Dvno Already on the list", vbInformation, "Sytem Message"
'        Else
'            With opnvoucher
'            Set x = ListView2.ListItems.Add(, , .Fields!DVNo)
'            x.SubItems(1) = .Fields!obrno
'            x.SubItems(2) = .Fields!Particular
'            x.SubItems(3) = GetClaimantDetails(IIf(IsNull(!ClaimantCode), "N/A", !ClaimantCode), "Name")
'            x.SubItems(4) = Format(.Fields!Gamount, "#,###.00")
'            x.SubItems(5) = .Fields!DVNo
'            End With
'        End If
'    End If
'Else
'MsgBox "No Record Found On the Database", vbInformation, "System Message"
'End If
'opnvoucher.Close
'Set opnvoucher = Nothing
'End Sub
'Private Sub LoadCNdetails(ByVal cnno As String)
'Dim opnvoucher As New ADODB.Recordset
'Dim sql As String
'Dim x
'
'sql = "SELECT min(b.trnno),a.dvno,a.obrno,a.particular,a.claimantcode,a.gamount,b.cnno,b.cndate,receiveddate  FROM tblAMIS_IncomingDVTrns as a inner join tblAMIS_CreditNotice as b ON a.dvno = b.dvno  WHERE b.cnno = '" & cnno & "' and b.ACTIONCODE = 1 and a.actioncode = 1group by a.obrno,a.particular,a.claimantcode,a.gamount,b.cnno,b.cndate,receiveddate,a.dvno order by a.dvno"
'
''Debug.Print sql
'
'opnvoucher.Open sql, opndbaseFMIS, adOpenStatic, adLockOptimistic
'
'If opnvoucher.RecordCount > 0 Then
'    Cmb_CnNo.Text = opnvoucher.Fields!cnno
'    DTPCNdate.Value = opnvoucher.Fields!cndate
'    DTPRdate.Value = opnvoucher.Fields!receiveddate
'    Do Until opnvoucher.EOF
'            With opnvoucher
'            Set x = ListView2.ListItems.Add(, , .Fields!DVNo)
'            x.SubItems(1) = .Fields!obrno
'            x.SubItems(2) = .Fields!Particular
'            x.SubItems(3) = GetClaimantDetails(IIf(IsNull(!ClaimantCode), "N/A", !ClaimantCode), "Name")
'            x.SubItems(4) = Format(.Fields!Gamount, "#,###.00")
'            x.SubItems(5) = .Fields!DVNo
'            End With
'            opnvoucher.MoveNext
'    Loop
'Else
'MsgBox "No Record Found On the Database", vbInformation, "System Message"
'End If
'opnvoucher.Close
'Set opnvoucher = Nothing
'End Sub
'Private Function IFalreadyOnDatabase(ByVal DVNo As String) As Boolean
'Dim rec As New ADODB.Recordset
'IFalreadyOnDatabase = False
'rec.Open "Select * from tblAMIS_CreditNotice where dvno = '" & DVNo & "' and actioncode = 1", opndbaseFMIS, adOpenKeyset, adLockReadOnly
'    If rec.RecordCount <> 0 Then
'        IFalreadyOnDatabase = True
'    End If
'rec.Close
'Set rec = Nothing
'End Function
'
'Private Function IFAlready(ByVal DVNo As String) As Boolean
'Dim x As Integer
'IFAlready = False
'For x = 1 To ListView2.ListItems.Count
'    If DVNo = Trim(ListView2.ListItems(x).Text) Then
'        IFAlready = True
'        Exit For
'    End If
'
'Next x
'End Function
'Private Function ExistBoth(ByVal RecordID As String) As String
'Dim opnTable1 As New ADODB.Recordset
'Dim existNtable1, existNtable2 As Boolean
'
'opnTable1.Open "Select * from  tblCMS_CDLiquidationRefundOR where recid='" & RecordID & "'", opndbaseFMIS, adOpenStatic, adLockOptimistic
'If opnTable1.RecordCount <> 0 Then
'    existNtable1 = True
'Else
'    existNtable1 = False
'End If
'opnTable1.Close
'Set opnTable1 = Nothing
'
'
'opnTable1.Open "Select * from   tblCMS_CDLiquiditionRefundForOverCA where recid='" & RecordID & "'", opndbaseFMIS, adOpenStatic, adLockOptimistic
'If opnTable1.RecordCount <> 0 Then
'    existNtable2 = True
'Else
'    existNtable2 = False
'End If
'opnTable1.Close
'Set opnTable1 = Nothing
'
'
'If existNtable1 = True And existNtable2 = True Then
'    ExistBoth = "BothExisting"
'ElseIf existNtable1 = True And existNtable2 = False Then
'    ExistBoth = "Table1"
'ElseIf existNtable1 = False And existNtable2 = True Then
'    ExistBoth = "Table2"
'ElseIf existNtable1 = False And existNtable2 = False Then
'    ExistBoth = "NoneExisting"
'End If
'
'End Function
'
'Private Function GetTotalAmtOfReplacement(ByVal FMISNo As String, ByVal REpNo As String) As Currency 'This is for the Amount of OR having replaced for the Over Amount of Check Against the Actual Total Cash Advance (the edited cash advance)
'Dim opnRepAmt As New ADODB.Recordset
'
'opnRepAmt.Open "Select sum(ORAmount) as RepAmt from tblCMS_CDLiquiditionRefundForOverCA where RecID='" & FMISNo & "' and RDONo='" & REpNo & "'", opndbaseFMIS, adOpenStatic, adLockOptimistic
'GetTotalAmtOfReplacement = IIf(IsNull(opnRepAmt!RepAmt), 0, opnRepAmt!RepAmt)
'opnRepAmt.Close
'Set opnRepAmt = Nothing
'
'End Function
'
'
'
'
'Private Function GetBackPrevAmtLacking(ByVal RecordID As String, ByVal ReportNo As String, ByVal Scenario As Integer) As Currency
'Dim opntable As New ADODB.Recordset
'
'
'If Scenario = 2 Or Scenario = 3 Then
'    opntable.Open "Select sum(ORAmount) as TotalLacking from tblCMS_CDLiquidationRefundOR where RecID='" & RecordID & "' and RDONo='" & ReportNo & "'", opndbaseFMIS, adOpenStatic, adLockOptimistic
'ElseIf Scenario = 1 Then
'    opntable.Open "Select sum(ORAmount) as TotalLacking from tblCMS_CDLiquiditionRefundForOverCA where RecID='" & RecordID & "' and RDONo='" & ReportNo & "'", opndbaseFMIS, adOpenStatic, adLockOptimistic
'End If
'
'If opntable.RecordCount <> 0 Then
'    GetBackPrevAmtLacking = opntable!TotalLacking
'End If
'
'
'End Function
'Private Function CheckAmtLacking(ByVal RecordID As String, ByVal ReportNo As String, ByVal Scenario As Integer, ByVal LackingAmt As Currency) As Boolean
'Dim opntable As New ADODB.Recordset
'
'
'If Scenario = 2 Or Scenario = 3 Then
'    opntable.Open "Select sum(ORAmount) as TotalLacking from tblCMS_CDLiquidationRefundOR where RecID='" & RecordID & "' and RDONo='" & ReportNo & "'", opndbaseFMIS, adOpenStatic, adLockOptimistic
'ElseIf Scenario = 1 Then
'    opntable.Open "Select sum(ORAmount) as TotalLacking from tblCMS_CDLiquiditionRefundForOverCA where RecID='" & RecordID & "' and RDONo='" & ReportNo & "'", opndbaseFMIS, adOpenStatic, adLockOptimistic
'End If
'
'If opntable.RecordCount <> 0 Then
'    If opntable!TotalLacking = LackingAmt Then
'        CheckAmtLacking = True
'    Else
'        CheckAmtLacking = False
'    End If
'Else
'    CheckAmtLacking = False
'End If
'
'
'End Function
'Private Function VerifyFexist(ByVal RecordID As String, ByVal ReportNo As String, ByVal Scenario As Integer) As Boolean
'Dim opntable As New ADODB.Recordset
'
'If Scenario = 2 Or Scenario = 3 Then
'    opntable.Open "Select * from tblCMS_CDLiquidationRefundOR where RecID='" & RecordID & "' and RDONo='" & ReportNo & "'", opndbaseFMIS, adOpenStatic, adLockOptimistic
'ElseIf Scenario = 1 Then
'    opntable.Open "Select * from tblCMS_CDLiquiditionRefundForOverCA where RecID='" & RecordID & "' and RDONo='" & ReportNo & "'", opndbaseFMIS, adOpenStatic, adLockOptimistic
'End If
'
'If opntable.RecordCount <> 0 Then
'    VerifyFexist = True
'Else
'    VerifyFexist = False
'End If
'
'End Function
'
'
'
'Private Sub FindLikeLastName(ByVal RecordID As String)
'Dim cc As Integer
'
'For cc = 0 To List2.ListCount - 1
'    If UCase(List2.List(cc)) Like UCase(RecordID) & "*" Then
'        List2.ListIndex = cc
'    End If
'Next cc
'End Sub
'
'
'
'Private Sub ListView2_DblClick()
'If ListView2.ListItems.Count <> 0 Then
'    If Len(Trim(ListView2.SelectedItem.Text)) > 0 Then
'        frmJEVPreparationReview.TxtDvno.Text = ListView2.SelectedItem.Text
'        frmJEVPreparationReview.Show vbModal
'    End If
'Else
'MsgBox "No Data to Save...", vbInformation, "System Message"
'End If
'End Sub
'
'
'Private Sub Timer1_Timer()
'Call LoadSavedReport
'Timer1.Enabled = False
'End Sub
'
'Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
'Select Case Button.Index
'    Case 1 'Print
'
'    Case 3 'Close
'        Unload Me
'End Select
'End Sub
'
'
'Private Sub txtSearch_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then
'    LoadSavedReport
'End If
'End Sub
'

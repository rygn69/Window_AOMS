VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "Mscomctl.ocx"
Begin VB.Form frmCDClaimantRegistry 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Claimant Entry Form"
   ClientHeight    =   7455
   ClientLeft      =   1920
   ClientTop       =   2280
   ClientWidth     =   11055
   Icon            =   "frmCDClaimantRegistry.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7455
   ScaleWidth      =   11055
   Begin MSComctlLib.ListView ListView1 
      Height          =   3735
      Left            =   120
      TabIndex        =   35
      Top             =   2880
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   6588
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Name"
         Object.Width           =   6174
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "ID"
         Object.Width           =   0
      EndProperty
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Select"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   630
      Left            =   120
      TabIndex        =   7
      Top             =   6660
      Width           =   3225
   End
   Begin VB.TextBox txt_search 
      Alignment       =   2  'Center
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
      Left            =   120
      TabIndex        =   0
      Top             =   2355
      Width           =   5250
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3660
      Left            =   375
      TabIndex        =   6
      Top             =   2880
      Width           =   2985
   End
   Begin VB.Frame Frame2 
      Caption         =   "Details"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4515
      Left            =   3405
      TabIndex        =   5
      Top             =   2790
      Width           =   7515
      Begin VB.PictureBox pic_EmployeeIndividual 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   3930
         Left            =   120
         ScaleHeight     =   3900
         ScaleWidth      =   7230
         TabIndex        =   9
         Top             =   360
         Visible         =   0   'False
         Width           =   7260
         Begin VB.TextBox txt_contactNo 
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
            Left            =   375
            TabIndex        =   21
            Top             =   3015
            Width           =   6405
         End
         Begin VB.TextBox txt_address 
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
            Left            =   375
            TabIndex        =   19
            Top             =   2235
            Width           =   6405
         End
         Begin VB.TextBox txt_suffix 
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
            TabIndex        =   17
            Top             =   1440
            Width           =   1830
         End
         Begin VB.TextBox txt_MiddleName 
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
            TabIndex        =   16
            Top             =   675
            Width           =   1830
         End
         Begin VB.TextBox txt_Firstname 
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
            Left            =   375
            TabIndex        =   15
            Top             =   1440
            Width           =   4215
         End
         Begin VB.TextBox txt_LastName 
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
            Left            =   375
            TabIndex        =   14
            Top             =   675
            Width           =   4215
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Contact Nos."
            Height          =   195
            Left            =   375
            TabIndex        =   20
            Top             =   2790
            Width           =   930
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Address"
            Height          =   195
            Left            =   375
            TabIndex        =   18
            Top             =   2025
            Width           =   570
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Middle Initial"
            Height          =   195
            Left            =   4935
            TabIndex        =   13
            Top             =   435
            Width           =   870
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Suffix"
            Height          =   195
            Left            =   4935
            TabIndex        =   12
            Top             =   1215
            Width           =   390
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Firstname"
            Height          =   195
            Left            =   375
            TabIndex        =   11
            Top             =   1200
            Width           =   675
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Lastname"
            Height          =   195
            Left            =   375
            TabIndex        =   10
            Top             =   450
            Width           =   690
         End
      End
      Begin VB.PictureBox Pic_company 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   3960
         Left            =   120
         ScaleHeight     =   3930
         ScaleWidth      =   7230
         TabIndex        =   22
         Top             =   360
         Visible         =   0   'False
         Width           =   7260
         Begin VB.CheckBox Check1 
            Caption         =   "Remittance Institution"
            Height          =   195
            Left            =   435
            TabIndex        =   34
            Top             =   3285
            Width           =   2205
         End
         Begin VB.TextBox txt_company 
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
            Left            =   420
            TabIndex        =   25
            Top             =   990
            Width           =   6405
         End
         Begin VB.TextBox txt_companyAddress 
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
            Left            =   420
            TabIndex        =   24
            Top             =   1785
            Width           =   6405
         End
         Begin VB.TextBox txt_companyContactNo 
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
            Left            =   420
            TabIndex        =   23
            Top             =   2565
            Width           =   6405
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Name of Company"
            Height          =   195
            Left            =   420
            TabIndex        =   28
            Top             =   765
            Width           =   1305
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Address"
            Height          =   195
            Left            =   420
            TabIndex        =   27
            Top             =   1575
            Width           =   570
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Contact Nos."
            Height          =   195
            Left            =   420
            TabIndex        =   26
            Top             =   2340
            Width           =   930
         End
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Claimant Classification"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1080
      Left            =   90
      TabIndex        =   1
      Top             =   855
      Width           =   10845
      Begin VB.OptionButton opn_MT 
         Caption         =   "Municipal Treasurers"
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
         Left            =   2520
         TabIndex        =   33
         Top             =   690
         Width           =   2205
      End
      Begin VB.OptionButton opn_BT 
         Caption         =   "Barangay Treasurers"
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
         Left            =   150
         TabIndex        =   32
         Top             =   690
         Width           =   2325
      End
      Begin VB.OptionButton opn_National 
         Caption         =   "National Offices"
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
         Left            =   6930
         TabIndex        =   31
         Top             =   360
         Width           =   1815
      End
      Begin VB.OptionButton opn_offices 
         Caption         =   "Provincial Offices"
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
         Left            =   4665
         TabIndex        =   29
         Top             =   360
         Width           =   1920
      End
      Begin VB.OptionButton opn_company 
         Caption         =   "Company"
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
         Left            =   9075
         TabIndex        =   4
         Top             =   360
         Width           =   1380
      End
      Begin VB.OptionButton opn_Individual 
         Caption         =   "Other Individual"
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
         Left            =   2520
         TabIndex        =   3
         Top             =   360
         Width           =   1830
      End
      Begin VB.OptionButton opn_Employee 
         Caption         =   "Capitol Employee"
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
         Left            =   150
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   360
         Value           =   -1  'True
         Width           =   1950
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Height          =   810
      Left            =   0
      TabIndex        =   30
      Top             =   0
      Width           =   11250
      _ExtentX        =   19844
      _ExtentY        =   1429
      ButtonWidth     =   1323
      ButtonHeight    =   1429
      Style           =   1
      ImageList       =   "itb32x32"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   8
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&New"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "slash"
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Save"
            ImageIndex      =   9
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "By Claimant"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "By Details"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "slash"
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Edit"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "slash"
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Delete"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "slash"
            Style           =   3
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5760
      Top             =   2040
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCDClaimantRegistry.frx":09EA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCDClaimantRegistry.frx":0A48
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCDClaimantRegistry.frx":0AA6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCDClaimantRegistry.frx":0B04
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList itb32x32 
      Left            =   6360
      Top             =   240
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
            Picture         =   "frmCDClaimantRegistry.frx":0B62
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCDClaimantRegistry.frx":24F4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCDClaimantRegistry.frx":3E86
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCDClaimantRegistry.frx":5818
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCDClaimantRegistry.frx":71AA
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCDClaimantRegistry.frx":8B3C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCDClaimantRegistry.frx":A4CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCDClaimantRegistry.frx":BE60
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCDClaimantRegistry.frx":D7F2
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCDClaimantRegistry.frx":F186
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCDClaimantRegistry.frx":FE62
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCDClaimantRegistry.frx":10742
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCDClaimantRegistry.frx":1141E
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCDClaimantRegistry.frx":120FA
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCDClaimantRegistry.frx":12DD6
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCDClaimantRegistry.frx":13AB2
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCDClaimantRegistry.frx":1478E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Type Lastname/Company name here to Search"
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
      Left            =   105
      TabIndex        =   8
      Top             =   1980
      Width           =   4245
   End
End
Attribute VB_Name = "frmCDClaimantRegistry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim saveflag, transNo As String

Private Sub ClearOption()
opn_Employee.Value = True
opn_Individual.Value = False
opn_company.Value = False
opn_offices.Value = False
End Sub

Private Sub SaveEntry()
Dim xxx As Integer
Dim tmpClaimantCode As String

If opn_Employee.Value = True Or opn_Individual.Value = True Or opn_company.Value = True Or opn_National.Value = True Or opn_BT.Value = True Or opn_MT.Value = True Then
    Select Case Classification
        Case "Employee"
            MsgBox "Adding and Updating of records for capitol employee is restricted!" & Chr(13) & "Please refer this to PMIS System Administrator", vbInformation, "System Information"
        
        Case "Individual"
            If DoScreening = True Then
                If saveflag = 1 Then 'new
                    
                    tmpClaimantCode = CreateClaimantCode("O")
                    opndbaseFMIS.Execute "insert into tblCMS_CDClaimantDetails (ClaimantCode,Firstname,MI,LastName,Suffix,Address,ContactNo) " & _
                        " values ('" & tmpClaimantCode & "','" & txt_Firstname.Text & "','" & IIf(Len(Trim(txt_MiddleName.Text)) = 0, "", UCase(Left(txt_MiddleName.Text, 1))) & "', " & _
                        " '" & txt_LastName.Text & "','" & IIf(Len(Trim(txt_suffix.Text)) = 0, "", txt_suffix.Text) & "','" & txt_address.Text & "','" & txt_contactNo.Text & "')"
                
                ElseIf saveflag = 2 Then 'Edit
                    If Len(transNo) <> 0 Then
                        
                        opndbaseFMIS.Execute "Update tblCMS_CDClaimantDetails set Firstname='" & txt_Firstname.Text & "',MI='" & IIf(Len(Trim(txt_MiddleName.Text)) = 0, "", UCase(Left(txt_MiddleName.Text, 1))) & "', " & _
                            " LastName='" & txt_LastName.Text & "',Suffix='" & IIf(Len(Trim(txt_suffix.Text)) = 0, "", txt_suffix.Text) & "',Address='" & txt_address.Text & "',ContactNo='" & txt_contactNo.Text & "' where CLAIMANTCODE='" & transNo & "' and ClaimantCode='" & ListView1.SelectedItem.SubItems(1) & "'"
                    
                    End If
                End If
                
                        '----------Showing Progress Bar in Status Bar----------------------------
                                       MsgBox "Saving Successful!", vbInformation, "System Information"
                        '------------------------------------------------------------------------
                Frame2.Enabled = False
                saveflag = 0
                transNo = 0
                Call Clear
                Call LoadSelectedClassification
                
            Else
                MsgBox "Complete all the important information needed before saving take effect!", vbInformation, "System Information"
            End If
        
        Case "Company"
            If DoScreening = True Then
                If saveflag = 1 Then 'new
                
                    tmpClaimantCode = CreateClaimantCode("C")
                    opndbaseFMIS.Execute "insert into tblCMS_CDClaimantDetails (ClaimantCode,LastName,Address,ContactNo,UsedInRemittance) " & _
                        " values ('" & tmpClaimantCode & "','" & txt_company.Text & "','" & txt_companyAddress.Text & "','" & txt_companyContactNo.Text & "'," & Check1.Value & ")"
                    
                    
                ElseIf saveflag = 2 Then 'Edit
                    If Len(transNo) <> 0 Then
                    
                        opndbaseFMIS.Execute "Update tblCMS_CDClaimantDetails set LastName='" & txt_company.Text & "',Address='" & txt_companyAddress.Text & "',ContactNo='" & txt_companyContactNo.Text & "',UsedInRemittance=" & Check1.Value & " where ClaimantCode='" & ListView1.SelectedItem.SubItems(1) & "'"
                    
                    
                    End If
                End If
                
                        '----------Showing Progress Bar in Status Bar----------------------------
                                       MsgBox "Saving Successful!", vbInformation, "System Information"
                                       Frame2.Enabled = False
                        '------------------------------------------------------------------------
                
                saveflag = 0
                transNo = 0
                Call Clear
                Call LoadSelectedClassification
                
            Else
                MsgBox "Complete all the important information needed before saving take effect!", vbInformation, "System Information"
            End If
    
        Case "National"
            If DoScreening = True Then
                If saveflag = 1 Then 'new
                
                    tmpClaimantCode = CreateClaimantCode("N")
                    opndbaseFMIS.Execute "insert into tblCMS_CDClaimantDetails (ClaimantCode,LastName,Address,ContactNo) " & _
                        " values ('" & tmpClaimantCode & "','" & txt_company.Text & "','" & txt_companyAddress.Text & "','" & txt_companyContactNo.Text & "')"
                    
                    
                ElseIf saveflag = 2 Then 'Edit
                    If transNo <> 0 Then
                    
                        opndbaseFMIS.Execute "Update tblCMS_CDClaimantDetails set LastName='" & txt_company.Text & "',Address='" & txt_companyAddress.Text & "',ContactNo='" & txt_companyContactNo.Text & "' where trnno=" & transNo & " and ClaimantCode='" & ListView1.SelectedItem.SubItems(1) & "'"
                    
                    End If
                End If
                
                        '----------Showing Progress Bar in Status Bar----------------------------
                                       MsgBox "Saving Successful!", vbInformation, "System Information"
                                       Frame2.Enabled = False
                        '------------------------------------------------------------------------
                
                saveflag = 0
                transNo = 0
                Call Clear
                Call LoadSelectedClassification
                
            Else
                MsgBox "Complete all the important information needed before saving take effect!", vbInformation, "System Information"
            End If
    
        Case "BarangayTreasurer"
            If DoScreening = True Then
                If saveflag = 1 Then 'new
                    tmpClaimantCode = CreateClaimantCode("BT")
                    opndbaseFMIS.Execute "insert into tblCMS_CDClaimantDetails (ClaimantCode,LastName,Address,ContactNo) " & _
                        " values ('" & tmpClaimantCode & "','" & txt_company.Text & "','" & txt_companyAddress.Text & "','" & txt_companyContactNo.Text & "')"
                    
                    
                ElseIf saveflag = 2 Then 'Edit
                    If Len(transNo) <> 0 Then
                    
                        opndbaseFMIS.Execute "Update tblCMS_CDClaimantDetails set LastName='" & txt_company.Text & "',Address='" & txt_companyAddress.Text & "',ContactNo='" & txt_companyContactNo.Text & "' where ClaimantCode='" & ListView1.SelectedItem.SubItems(1) & "'"
                    
                    End If
                End If
                
                        '----------Showing Progress Bar in Status Bar----------------------------
                                       MsgBox "Saving Successful!", vbInformation, "System Information"
                                       Frame2.Enabled = False
                        '------------------------------------------------------------------------
                
                saveflag = 0
                transNo = 0
                Call Clear
                Call LoadSelectedClassification
                
            Else
                MsgBox "Complete all the important information needed before saving take effect!", vbInformation, "System Information"
            End If
    
        Case "MunicipalTreasurer"
            If DoScreening = True Then
                If saveflag = 1 Then 'new
                
                    tmpClaimantCode = CreateClaimantCode("MT")
                    opndbaseFMIS.Execute "insert into tblCMS_CDClaimantDetails (ClaimantCode,LastName,Address,ContactNo) " & _
                        " values ('" & tmpClaimantCode & "','" & txt_company.Text & "','" & txt_companyAddress.Text & "','" & txt_companyContactNo.Text & "')"
                    
                ElseIf saveflag = 2 Then 'Edit
                    If transNo <> 0 Then
                        opndbaseFMIS.Execute "Update tblCMS_CDClaimantDetails set LastName='" & txt_company.Text & "',Address='" & txt_companyAddress.Text & "',ContactNo='" & txt_companyContactNo.Text & "' where trnno=" & transNo & " and ClaimantCode='" & ListView1.SelectedItem.SubItems(1) & "'"
                    End If
                End If
                        '----------Showing Progress Bar in Status Bar----------------------------
                                       MsgBox "Saving Successful!", vbInformation, "System Information"
                                       Frame2.Enabled = False
                        '-----------------------------------------------------------------------
                saveflag = 0
                transNo = 0
                Call Clear
                Call LoadSelectedClassification
                
            Else
                MsgBox "Complete all the important information needed before saving take effect!", vbInformation, "System Information"
            End If
    
    End Select
End If
End Sub
Private Function DoScreening() As Boolean
Select Case Classification
    Case "Individual"
        If Len(Trim(txt_LastName.Text)) <> 0 And Len(Trim(txt_Firstname.Text)) <> 0 And Len(Trim(txt_address.Text)) <> 0 Then
            DoScreening = True
        Else
            DoScreening = False
        End If
    Case "Company", "National", "BarangayTreasurer", "MunicipalTreasurer"
        If Len(Trim(txt_company.Text)) <> 0 And Len(Trim(txt_companyAddress.Text)) <> 0 Then
            DoScreening = True
        Else
            DoScreening = False
        End If
End Select
End Function

Private Sub Command1_Click()
Select Case Classification
    Case "Individual", "Employee"
        If ActiveFormCaller = "frmAccntAdviceSpecial" Then
            frmAccntAdviceSpecial.txt_payee.Text = ListView1.SelectedItem.Text
        
        ElseIf ActiveFormCaller = "frmEXVerifyCashAvailabilityNew" Then
            'frmEXVerifyCashAvailabilityNew.txt_claimant.Text = listview1.SelectedItem.Text '& "#" & listview1.SelectedItem.subitems(1)
            'frmEXVerifyCashAvailabilityNew.lbl_ClaimantCode.Caption = listview1.SelectedItem.subitems(1)
        
        ElseIf ActiveFormCaller = "frmExChangeClaimantUtility" Then
            'frmExChangeClaimantUtility.txt_newClaimant.Text = listview1.SelectedItem.Text '& "#" & listview1.SelectedItem.subitems(1)
            'frmExChangeClaimantUtility.txt_ClaimantAddress.Text = txt_address.Text
            'frmExChangeClaimantUtility.lbl_newClaimantCode.Caption = listview1.SelectedItem.subitems(1)
        
        ElseIf ActiveFormCaller = "frmCDSpecialPostingOfCheck" Then
            'frmCDSpecialPostingOfCheck.lbl_claimantname.Caption = listview1.SelectedItem.Text
            'frmCDSpecialPostingOfCheck.lbl_ClaimantCode.Caption = listview1.SelectedItem.subitems(1)
        
        ElseIf ActiveFormCaller = "frmIncomingTrn" Then
            frmIncomingTrn.txtClaimant.Text = ListView1.SelectedItem.Text
            frmIncomingTrn.txtClaimantCode.Text = ListView1.SelectedItem.SubItems(1)
        ElseIf ActiveFormCaller = "frm_POreg" Then
            frm_POReg.txtClaimant.Text = ListView1.SelectedItem.Text
            frm_POReg.txtClaimantCode.Text = ListView1.SelectedItem.SubItems(1)
        ElseIf ActiveFormCaller = "frmJEVNumberingAssignment_New" Then
            CM = ListView1.SelectedItem.Text
            cc = ListView1.SelectedItem.SubItems(1)
        ElseIf ActiveFormCaller = "frmJEVPreparation" Then
            frmJEVPreparation_New.txtClaimant.Text = ListView1.SelectedItem.Text
            frmJEVPreparation_New.txtClaimantCode.Text = ListView1.SelectedItem.SubItems(1)
         ElseIf ActiveFormCaller = "frmJEVPreparation_liquidation" Then
            frmJEVPreparationfor_Liquidation.txtClaimant.Text = ListView1.SelectedItem.Text
            frmJEVPreparationfor_Liquidation.txtClaimantCode.Text = ListView1.SelectedItem.SubItems(1)
        ElseIf ActiveFormCaller = "CAclaimant" Then
            frmJEVPreparationfor_Liquidation.txtCClaimant.Text = ListView1.SelectedItem.Text
            frmJEVPreparationfor_Liquidation.txtCclaimantcode.Text = ListView1.SelectedItem.SubItems(1)
        ElseIf ActiveFormCaller = "frm_Signatory" Then
            frm_Signatory.txtname.Text = ListView1.SelectedItem.Text
            frm_Signatory.txtID.Text = ListView1.SelectedItem.SubItems(1)
        ElseIf ActiveFormCaller = "frmFinalJEV" Then
            frmFinalJev.txtClaimant.Text = ListView1.SelectedItem.Text
            frmFinalJev.ClaimantCode = ListView1.SelectedItem.SubItems(1)
        ElseIf ActiveFormCaller = "frmFinalJev" Then
           frmFinalJev.txtClaimant.Text = ListView1.SelectedItem.Text
           frmFinalJev.ClaimantCode = ListView1.SelectedItem.SubItems(1)
         ElseIf ActiveFormCaller = "CAclaimantNEW" Then
            frmJEVPreparationfor_Liquidation.txtCClaimant.Text = ListView1.SelectedItem.Text
           frmJEVPreparationfor_Liquidation.txtCclaimantcode.Text = ListView1.SelectedItem.SubItems(1)
        End If
    Case "Company", "National", "BarangayTreasurer", "MunicipalTreasurer"
        If ActiveFormCaller = "frmAccntAdviceSpecial" Then
            frmAccntAdviceSpecial.txt_payee.Text = UCase(txt_company.Text) '& "#" & listview1.SelectedItem.subitems(1)
            'frmCDPreparedCheckRegistry.txt_ClaimantAddress.Text = txt_companyAddress.Text
            'frmCDPreparedCheckRegistry.lbl_ClaimantCode.Caption = listview1.SelectedItem.subitems(1)
        
        ElseIf ActiveFormCaller = "frmEXVerifyCashAvailabilityNew" Then
            'frmEXVerifyCashAvailabilityNew.txt_claimant.Text = UCase(txt_company.Text) '& "#" & listview1.SelectedItem.subitems(1)
            'frmEXVerifyCashAvailabilityNew.lbl_ClaimantCode.Caption = listview1.SelectedItem.subitems(1)
        
        ElseIf ActiveFormCaller = "frmExChangeClaimantUtility" Then
            'frmExChangeClaimantUtility.txt_newClaimant.Text = UCase(txt_company.Text) '& "#" & listview1.SelectedItem.subitems(1)
            'frmExChangeClaimantUtility.txt_ClaimantAddress.Text = txt_companyAddress.Text
            'frmExChangeClaimantUtility.lbl_newClaimantCode.Caption = listview1.SelectedItem.subitems(1)
        
        ElseIf ActiveFormCaller = "frmCDSpecialPostingOfCheck" Then
             'frmCDSpecialPostingOfCheck.lbl_claimantname.Caption = UCase(txt_company.Text)
              'frmCDSpecialPostingOfCheck.lbl_ClaimantCode.Caption = listview1.SelectedItem.subitems(1)
        
        ElseIf ActiveFormCaller = "frmIncomingTrn" Then
            frmIncomingTrn.txtClaimant.Text = UCase(txt_company.Text)
            frmIncomingTrn.txtClaimantCode.Text = ListView1.SelectedItem.SubItems(1)
         ElseIf ActiveFormCaller = "frm_POreg" Then
            frm_POReg.txtClaimant.Text = UCase(txt_company.Text)
            frm_POReg.txtClaimantCode.Text = ListView1.SelectedItem.SubItems(1)
        ElseIf ActiveFormCaller = "frmJEVNumberingAssignment_New" Then
            CM = UCase(txt_company.Text)
            cc = ListView1.SelectedItem.SubItems(1)
         ElseIf ActiveFormCaller = "CAclaimant" Then
             'frmJEVPreparation.txtCClaimant.Text = UCase(txt_company.Text)
            'frmJEVPreparation.txtCclaimantcode.Text = listview1.SelectedItem.subitems(1)
        
        ElseIf ActiveFormCaller = "frmJEVPreparation" Then
            frmJEVPreparation_New.txtClaimant.Text = ListView1.SelectedItem.Text
            frmJEVPreparation_New.txtClaimantCode.Text = ListView1.SelectedItem.SubItems(1)
        ElseIf ActiveFormCaller = "frmFinalJEV" Then
           frmFinalJev.txtClaimant.Text = UCase(txt_company.Text)
           frmFinalJev.ClaimantCode = ListView1.SelectedItem.SubItems(1)
         ElseIf ActiveFormCaller = "CAclaimantNEW" Then
            frmJEVPreparationfor_Liquidation.txtCClaimant.Text = ListView1.SelectedItem.Text
           frmJEVPreparationfor_Liquidation.txtCclaimantcode.Text = ListView1.SelectedItem.SubItems(1)
        End If
    
    Case "Offices"
        If ActiveFormCaller = "frmAccntAdviceSpecial" Then
            frmAccntAdviceSpecial.txt_payee.Text = UCase(txt_company.Text) '& "#" & listview1.SelectedItem.subitems(1)
        
        ElseIf ActiveFormCaller = "frmEXVerifyCashAvailabilityNew" Then
            'frmEXVerifyCashAvailabilityNew.txt_claimant.Text = UCase(txt_company.Text) '& "#" & listview1.SelectedItem.subitems(1)
            'frmEXVerifyCashAvailabilityNew.lbl_ClaimantCode.Caption = ListView1.SelectedItem.SubItems(1)
        
        ElseIf ActiveFormCaller = "frmExChangeClaimantUtility" Then
            'frmExChangeClaimantUtility.txt_newClaimant.Text = UCase(txt_company.Text) '& "#" & listview1.SelectedItem.subitems(1)
            'frmExChangeClaimantUtility.txt_ClaimantAddress.Text = txt_companyAddress.Text
            'frmExChangeClaimantUtility.lbl_newClaimantCode.Caption = ListView1.SelectedItem.SubItems(1)
        
        ElseIf ActiveFormCaller = "frmCDSpecialPostingOfCheck" Then
            'frmCDSpecialPostingOfCheck.lbl_claimantname.Caption = UCase(txt_company.Text)
            'frmCDSpecialPostingOfCheck.lbl_ClaimantCode.Caption = ListView1.SelectedItem.SubItems(1)
            
        ElseIf ActiveFormCaller = "frmIncomingTrn" Then
            frmIncomingTrn.txtClaimant.Text = UCase(txt_company.Text)
            frmIncomingTrn.txtClaimantCode.Text = ListView1.SelectedItem.SubItems(1)
        ElseIf ActiveFormCaller = "frm_POreg" Then
            frm_POReg.txtClaimant.Text = UCase(txt_company.Text)
            frm_POReg.txtClaimantCode.Text = ListView1.SelectedItem.SubItems(1)
        ElseIf ActiveFormCaller = "frmJEVNumberingAssignment_New" Then
            CM = UCase(txt_company.Text)
            cc = ListView1.SelectedItem.SubItems(1)
        ElseIf ActiveFormCaller = "frmJEVPreparation" Then
           frmJEVPreparation_New.txtClaimant.Text = ListView1.SelectedItem.Text
            frmJEVPreparation_New.txtClaimantCode.Text = ListView1.SelectedItem.SubItems(1)
        ElseIf ActiveFormCaller = "CAclaimantNEW" Then
            frmJEVPreparationfor_Liquidation.txtCClaimant.Text = ListView1.SelectedItem.Text
           frmJEVPreparationfor_Liquidation.txtClaimantCode.Text = ListView1.SelectedItem.SubItems(1)
            
            'frmJEVPreparation.txtCClaimant.Text = UCase(txt_company.Text)
            'frmJEVPreparation.txtCclaimantcode.Text = ListView1.SelectedItem.SubItems(1)
        ElseIf ActiveFormCaller = "frmFinalJEV" Then
           frmFinalJev.txtClaimant.Text = ListView1.SelectedItem.Text
           frmFinalJev.ClaimantCode = ListView1.SelectedItem.SubItems(1)
        End If
    
End Select
Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then
    Unload Me
End If
End Sub

Private Sub Form_Load()
opn_Employee.Value = True
'txt_Search.SetFocus
Me.Top = (Screen.Height - Me.Height) / 2
Me.Left = 3650

Call Clear
Call ClearOption
saveflag = 0 'No action
transNo = 0
End Sub
Private Sub LoadOffices()
Dim opnOffice As New ADODB.Recordset
Dim xx As Integer
Dim x
opnOffice.Open "Select * from tblREF_AIS_Offices order by OfficeMedium", opndbaseFMIS, adOpenStatic, adLockOptimistic
If opnOffice.RecordCount <> 0 Then
    ListView1.ListItems.Clear
    Do Until opnOffice.EOF
        Set x = ListView1.ListItems.Add(, , UCase(opnOffice!OfficeMedium))
        x.SubItems(1) = opnOffice!fmisofficeid
        opnOffice.MoveNext

    Loop
Else
    ListView1.ListItems.Clear
End If
opnOffice.Close
Set opnOffice = Nothing
End Sub

Private Sub LoadClaimantDetails(ByVal ClaimantCode As String)
Dim opnDetails As New ADODB.Recordset

Select Case Classification
    Case "Individual", "Company", "National", "BarangayTreasurer", "MunicipalTreasurer"
        opnDetails.Open "Select * from tblCMS_CDClaimantDetails where ClaimantCode = '" & ClaimantCode & "'", opndbaseFMIS, adOpenStatic, adLockOptimistic
    Case "Employee"
        opnDetails.Open "Select * from pmis.dbo.employee where swipemployeeid = '" & ClaimantCode & "'", opndbaseFMIS, adOpenStatic, adLockOptimistic
    Case "Offices"
        opnDetails.Open "Select * from tblREF_AIS_Offices where FMISOfficeID = " & ClaimantCode & "", opndbaseFMIS, adOpenStatic, adLockOptimistic
End Select

If opnDetails.RecordCount <> 0 Then

    Select Case Classification
        Case "Individual"
            txt_LastName.Text = UCase(opnDetails!Lastname)
            txt_Firstname.Text = UCase(opnDetails!Firstname)
            txt_MiddleName.Text = UCase(Left(opnDetails!MI, 1))
            txt_suffix.Text = opnDetails!Suffix
            txt_contactNo.Text = opnDetails!ContactNo
            txt_address.Text = StrConv(opnDetails!Address, vbProperCase)
        
        Case "Company", "National", "BarangayTreasurer", "MunicipalTreasurer"
            txt_company.Text = UCase(opnDetails!Lastname)
            'txt_companyAddress.Text = mydll.ProperCase(opnDetails!Address)
            txt_companyAddress.Text = StrConv(opnDetails!Address, vbProperCase)
            txt_companyContactNo.Text = opnDetails!ContactNo
            If opnDetails!UsedInRemittance Then
                Check1.Value = 1
            Else
                Check1.Value = 0
            End If
        
        Case "Offices"
            txt_company.Text = UCase(opnDetails!Officename)
            If IsNull(opnDetails!Address) Then
                txt_companyAddress.Text = ""
            Else
                txt_companyAddress.Text = StrConv(opnDetails!Address, vbProperCase)
            End If
            txt_companyContactNo.Text = ""
        
        Case "Employee"
            txt_LastName.Text = UCase(opnDetails!Lastname)
            txt_Firstname.Text = UCase(opnDetails!Firstname)
            txt_MiddleName.Text = UCase(Left(opnDetails!MI, 1))
            txt_suffix.Text = opnDetails!Suffix
            txt_contactNo.Text = IIf(IsNull(opnDetails!Telephone), "", opnDetails!Telephone)
            txt_address.Text = StrConv(IIf(IsNull(opnDetails!Address), "", opnDetails!Address), vbProperCase)
    End Select
        
Else
    Call Clear
End If
opnDetails.Close
Set opnDetails = Nothing
End Sub
Private Function CreateClaimantCode(ByVal identifier As String) As String
Dim opnClass As New ADODB.Recordset
Dim xx As Long

opnClass.Open "Select ClaimantCode from tblCMS_CDClaimantDetails where ClaimantCode like '" & identifier & "%' order by trnno desc", opndbaseFMIS, adOpenStatic, adLockOptimistic
If opnClass.RecordCount <> 0 Then
    xx = Mid(opnClass!ClaimantCode, Len(identifier) + 1, Len(opnClass!ClaimantCode) - 1)
    xx = CLng(xx) + 1
    CreateClaimantCode = identifier & Format(xx, "0000")
Else
    CreateClaimantCode = identifier & "0001"
End If
opnClass.Close
Set opnClass = Nothing
End Function
Private Sub Clear()
txt_LastName.Text = ""
txt_Firstname.Text = ""
txt_MiddleName.Text = ""
txt_suffix.Text = ""
txt_contactNo.Text = ""
txt_address.Text = ""

txt_company.Text = ""
txt_companyAddress.Text = ""
txt_companyContactNo.Text = ""

'txt_Search.Text = ""
txt_search.Enabled = True
ListView1.ListItems.Clear
End Sub
Private Function Classification() As String
If opn_Employee.Value = True Then
    Classification = "Employee"
ElseIf opn_Individual.Value = True Then
    Classification = "Individual"
ElseIf opn_company.Value = True Then
    Classification = "Company"
ElseIf opn_offices.Value = True Then
    Classification = "Offices"
ElseIf opn_National.Value = True Then
    Classification = "National"
ElseIf opn_BT.Value = True Then
    Classification = "BarangayTreasurer"
ElseIf opn_MT.Value = True Then
    Classification = "MunicipalTreasurer"
End If

End Function
Private Sub Form_Unload(Cancel As Integer)
Set frmCDClaimantRegistry = Nothing
End Sub
Private Sub LoadEmployeeIDs()
Dim opnemployee As New ADODB.Recordset
Dim x
On Error GoTo bad
If Len(txt_search.Text) <> 0 Then
    'frmConnCheck.Show 1
    opnemployee.Open "Select swipemployeeid,case when rtrim(ltrim(isnull(suffix,' '))) <> '' then lastname + ', ' + firstname + ' '  + isnull(Mi,' ') + ', ' + isnull(suffix,' ') else lastname + ', ' + firstname + ' '  + isnull(Mi, ' ' ) end as Name from pmis.dbo.employee where len(swipemployeeid)<>0 and Lastname like '" & Replace(txt_search.Text, "'", "''") & "%' order by Lastname, firstname", opndbaseFMIS, adOpenStatic, adLockOptimistic
Else
    opnemployee.Open "Select swipemployeeid,case when rtrim(ltrim(isnull(suffix, ' '))) <> '' then lastname + ', ' + firstname + ' '  + isnull(Mi,' ') + ', ' + isnull(suffix,' ') else lastname + ', ' + firstname + ' '  + isnull(Mi, ' ' ) end as Name from pmis.dbo.employee where len(swipemployeeid)<>0 order by lastname, firstname", opndbaseFMIS, adOpenStatic, adLockOptimistic
End If


If opnemployee.RecordCount <> 0 Then
    ListView1.ListItems.Clear
    Do Until opnemployee.EOF
'        List1.AddItem (UCase(opnemployee!name))
'        ListView1.ItemData(List1.NewIndex) = opnemployee!swipemployeeid
        Set x = ListView1.ListItems.Add(, , UCase(opnemployee!name))
        x.SubItems(1) = opnemployee!swipemployeeid
        opnemployee.MoveNext
    Loop
Else
    ListView1.ListItems.Clear
End If
opnemployee.Close
Set opnemployee = Nothing
Exit Sub
bad:
MsgBox err.description
End Sub

Private Sub LoadOtherClaimantIDs(ByVal identifier As String)
Dim opnOther As New ADODB.Recordset
Dim cc As Long
Dim x

If Len(txt_search.Text) <> 0 Then
    opnOther.Open "Select ClaimantCode,trnno,case when left('" & identifier & "',1) = 'O'  then (case when rtrim(ltrim(isnull(suffix,' '))) <> '' then lastname + ', ' + firstname + ' '  + Mi + ', ' + suffix else lastname + ', ' + firstname + ' '  + Mi end)else lastname end as Name  from tblCMS_CDClaimantDetails where ClaimantCode like '" & identifier & "%' and lastname like '" & Replace(txt_search.Text, "'", "''") & "%' order by lastname,firstname", opndbaseFMIS, adOpenStatic, adLockOptimistic
Else
    opnOther.Open "Select ClaimantCode,trnno,case when left('" & identifier & "',1) = 'O' then (case when rtrim(ltrim(suffix)) <> '' then lastname + ', ' + firstname + ' '  + Mi + ', ' + suffix else lastname + ', ' + firstname + ' '  + Mi end) else lastname end as Name from tblCMS_CDClaimantDetails where ClaimantCode like '" & identifier & "%'  order by lastname,firstname", opndbaseFMIS, adOpenStatic, adLockOptimistic
End If

If opnOther.RecordCount <> 0 Then
   ListView1.ListItems.Clear
    Do Until opnOther.EOF
         Set x = ListView1.ListItems.Add(, , UCase(IIf(IsNull(opnOther!name), "", opnOther!name)))
        x.SubItems(1) = opnOther!ClaimantCode
        opnOther.MoveNext
        cc = cc + 1
    Loop
Else
    ListView1.ListItems.Clear
End If
opnOther.Close
Set opnOther = Nothing
Exit Sub

End Sub
Private Sub VerifyButtonState()
Select Case Classification
    Case "Employee", "Offices"
        Toolbar1.Buttons(1).Enabled = False
        Toolbar1.Buttons(3).Enabled = False
        Toolbar1.Buttons(5).Enabled = False
        Toolbar1.Buttons(7).Enabled = False
    Case Else
        Toolbar1.Buttons(1).Enabled = True
        Toolbar1.Buttons(3).Enabled = True
        Toolbar1.Buttons(5).Enabled = True
        Toolbar1.Buttons(7).Enabled = True
End Select
End Sub
Private Sub LoadSelectedClassification()
Select Case Classification
    Case "Employee"
        Call LoadEmployeeIDs
        pic_EmployeeIndividual.Move 180, 270, 7260, 4035
        pic_EmployeeIndividual.Visible = True
        pic_EmployeeIndividual.Enabled = False
        Pic_company.Visible = False
    
    Case "Individual"
        Call LoadOtherClaimantIDs("O")
        pic_EmployeeIndividual.Move 180, 270, 7260, 4035
        pic_EmployeeIndividual.Visible = True
        pic_EmployeeIndividual.Enabled = True
        Pic_company.Visible = False
    
    Case "Company"
        Call LoadOtherClaimantIDs("C")
        pic_EmployeeIndividual.Visible = False
        Pic_company.Visible = True
        Pic_company.Move 180, 270, 7260, 4035
        Pic_company.Enabled = True
        
    Case "Offices"
        Call LoadOffices
        pic_EmployeeIndividual.Visible = False
        Pic_company.Visible = True
        Pic_company.Move 180, 270, 7260, 4035
        Pic_company.Enabled = False
        
    Case "National"
        Call LoadOtherClaimantIDs("N")
        pic_EmployeeIndividual.Visible = False
        Pic_company.Visible = True
        Pic_company.Move 180, 270, 7260, 4035
        Pic_company.Enabled = True
        
    Case "BarangayTreasurer"
        Call LoadOtherClaimantIDs("BT")
        pic_EmployeeIndividual.Visible = False
        Pic_company.Visible = True
        Pic_company.Move 180, 270, 7260, 4035
        Pic_company.Enabled = True
    
    Case "MunicipalTreasurer"
        Call LoadOtherClaimantIDs("MT")
        pic_EmployeeIndividual.Visible = False
        Pic_company.Visible = True
        Pic_company.Move 180, 270, 7260, 4035
        Pic_company.Enabled = True
End Select
End Sub

Private Sub ListView1_Click()
On Error Resume Next
        Call LoadClaimantDetails(ListView1.SelectedItem.SubItems(1))
Frame2.Enabled = False
saveflag = 0
transNo = 0
End Sub

Private Sub ListView1_DblClick()
Call Command1_Click
End Sub

Private Sub opn_BT_Click()
Call Clear
Call LoadSelectedClassification
Call VerifyButtonState
List1.Font = "MS Sans Serif"
Frame2.Enabled = False
txt_search.SetFocus
End Sub

Private Sub opn_company_Click()
Call Clear
Call LoadSelectedClassification
Call VerifyButtonState
List1.Font = "MS Sans Serif"
Frame2.Enabled = False
txt_search.SetFocus
End Sub

Private Sub opn_Employee_Click()
Call Clear
Call LoadSelectedClassification
Call VerifyButtonState
List1.Font = "MS Sans Serif"
Frame2.Enabled = False
txt_search.SetFocus
End Sub

Private Sub opn_Individual_Click()
Call Clear
Call LoadSelectedClassification
Call VerifyButtonState
List1.Font = "MS Sans Serif"
Frame2.Enabled = False
txt_search.SetFocus
End Sub

Private Sub Option1_Click()
Call LoadOffices
txt_search.Visible = False
End Sub

Private Sub opn_MT_Click()
Call Clear
Call LoadSelectedClassification
Call VerifyButtonState
List1.Font = "MS Sans Serif"
Frame2.Enabled = False
txt_search.SetFocus
End Sub

Private Sub opn_National_Click()
Call Clear
Call LoadSelectedClassification
Call VerifyButtonState
List1.Font = "MS Sans Serif"
Frame2.Enabled = False
txt_search.SetFocus
End Sub

Private Sub opn_offices_Click()

Call Clear
txt_search.Enabled = False
Call LoadSelectedClassification
Call VerifyButtonState
List1.Font = "Arial Narrow"
Frame2.Enabled = False
'txt_Search.SetFocus
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
    Case 1 'New
        saveflag = 1 'New
        transNo = 0
        Call Clear
        Frame2.Enabled = True
        'txt_LastName.SetFocus
    Case 3 'Save
        Call SaveEntry
        
    Case 5 'Edit
        If Classification = "Individual" Or Classification = "Company" Or Classification = "National" Or Classification = "BarangayTreasurer" Or Classification = "MunicipalTreasurer" Then
            saveflag = 2 'Edit
            transNo = ListView1.SelectedItem.SubItems(1)
            Frame2.Enabled = True
        End If
    Case 7 'Delete
        Frame2.Enabled = True
        If Classification = "Individual" Or Classification = "Company" Or Classification = "National" Or Classification = "BarangayTreasurer" Or Classification = "MunicipalTreasurer" Then
            transNo = ListView1.SelectedItem.SubItems(1)
            If MsgBox("Are You Sure Want to Delete this Record?", vbQuestion + vbYesNo, "System Confirmation Query") = vbYes Then
                If CodeUsed(transNo) = False Then
                    opndbaseFMIS.Execute "Delete from tblCMS_CDClaimantDetails where ClaimantCode='" & transNo & "'"
                    MsgBox "Deleting Record Successful!", vbInformation, "System Information"
                    Call Clear
                    saveflag = 0 'No action
                    transNo = 0
                Else
                    MsgBox "Claimant is in used!" & Chr(13) & Chr(13) & "Deletion is not Allowed!", vbInformation, "System Information"
                    saveflag = 0 'No action
                    transNo = 0
                End If
            End If
        End If
End Select
End Sub
Private Function CodeUsed(ByVal ClaimantCode As String) As Boolean
Dim opntable As New ADODB.Recordset

'------------------------------------1
opntable.Open "Select [ClaimantCode] from tblCMS_EXCashVerification where claimantcode='" & ClaimantCode & "' and actioncode=1", opndbaseFMIS, adOpenStatic, adLockOptimistic
    If opntable.RecordCount <> 0 Then
        CodeUsed = True
    Else
        CodeUsed = False
    End If
opntable.Close
Set opntable = Nothing

'------------------------------------2
If CodeUsed = False Then
    opntable.Open "Select [ClaimantCode] from tblCMS_CDTransactionDetails where claimantcode='" & ClaimantCode & "' and actioncode=1", opndbaseFMIS, adOpenStatic, adLockOptimistic
        If opntable.RecordCount <> 0 Then
            CodeUsed = True
        Else
            CodeUsed = False
        End If
    opntable.Close
    Set opntable = Nothing
End If

'------------------------------------3
If CodeUsed = False Then
    opntable.Open "Select [ClaimantCode] from tblCMS_CDPreparedCheck where claimantcode='" & ClaimantCode & "' and actioncode=1", opndbaseFMIS, adOpenStatic, adLockOptimistic
        If opntable.RecordCount <> 0 Then
            CodeUsed = True
        Else
            CodeUsed = False
        End If
    opntable.Close
    Set opntable = Nothing
End If

If CodeUsed = False Then
    opntable.Open "Select [ClaimantCode] from tblAMIS_IncomingDVTrns where [ClaimantCode]='" & ClaimantCode & "' and actioncode=1", opndbaseFMIS, adOpenStatic, adLockOptimistic
        If opntable.RecordCount <> 0 Then
            CodeUsed = True
        Else
            CodeUsed = False
        End If
    opntable.Close
    Set opntable = Nothing
End If

End Function
Private Sub txt_address_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    If Len(Trim(txt_contactNo.Text)) <> 0 Then
        txt_contactNo.SelStart = 0
        txt_contactNo.SelLength = Len(txt_contactNo.Text)
        txt_contactNo.SetFocus
    Else
        txt_contactNo.SetFocus
    End If
End If
End Sub

Private Sub txt_company_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    If Len(Trim(txt_companyAddress.Text)) <> 0 Then
        txt_companyAddress.SelStart = 0
        txt_companyAddress.SelLength = Len(txt_companyAddress.Text)
        txt_companyAddress.SetFocus
    Else
        txt_companyAddress.SetFocus
    End If
End If
End Sub
Private Sub txt_companyAddress_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    If Len(Trim(txt_companyContactNo.Text)) <> 0 Then
        txt_companyContactNo.SelStart = 0
        txt_companyContactNo.SelLength = Len(txt_companyContactNo.Text)
        txt_companyContactNo.SetFocus
    Else
        txt_companyContactNo.SetFocus
    End If
End If
End Sub

Private Sub txt_Firstname_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    txt_Firstname.Text = UCase(txt_Firstname.Text)
    If Len(Trim(txt_MiddleName.Text)) <> 0 Then
        txt_MiddleName.SelStart = 0
        txt_MiddleName.SelLength = Len(txt_MiddleName.Text)
        txt_MiddleName.SetFocus
    Else
        txt_MiddleName.SetFocus
    End If
End If
End Sub
Private Sub txt_LastName_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    txt_LastName.Text = UCase(txt_LastName.Text)
    If Len(Trim(txt_suffix.Text)) <> 0 Then
        txt_suffix.SelStart = 0
        txt_suffix.SelLength = Len(txt_suffix.Text)
        txt_suffix.SetFocus
    Else
        txt_Firstname.SetFocus
    End If
End If
End Sub

Private Sub txt_MiddleName_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    If Len(Trim(txt_suffix.Text)) <> 0 Then
        txt_suffix.SelStart = 0
        txt_suffix.SelLength = Len(txt_suffix.Text)
        txt_suffix.SetFocus
    Else
        txt_suffix.SetFocus
    End If
End If
End Sub

Private Sub txt_search_Change()
Call LoadSelectedClassification
End Sub
Private Sub txt_suffix_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    If Len(Trim(txt_address.Text)) <> 0 Then
        txt_address.SelStart = 0
        txt_address.SelLength = Len(txt_address.Text)
        txt_address.SetFocus
    Else
        txt_address.SetFocus
    End If
    
End If
End Sub

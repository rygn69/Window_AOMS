VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frm_POReg 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Incoming Purchase Order"
   ClientHeight    =   7665
   ClientLeft      =   5055
   ClientTop       =   4680
   ClientWidth     =   13140
   Icon            =   "frm_POReg.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7665
   ScaleWidth      =   13140
   Begin VB.TextBox txtDate 
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
      Left            =   11040
      Locked          =   -1  'True
      TabIndex        =   27
      Top             =   1200
      Width           =   1800
   End
   Begin lvButton.lvButtons_H Command2 
      Height          =   495
      Left            =   10470
      TabIndex        =   26
      Top             =   1920
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   873
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
      cBhover         =   33023
      cGradient       =   33023
      Gradient        =   3
      Mode            =   0
      Value           =   0   'False
      ImgAlign        =   4
      Image           =   "frm_POReg.frx":076A
      cBack           =   -2147483633
   End
   Begin VB.TextBox txtAmount 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5640
      Locked          =   -1  'True
      TabIndex        =   25
      Top             =   1920
      Width           =   4800
   End
   Begin VB.CommandButton Command5 
      Caption         =   "View JEV"
      Height          =   495
      Left            =   120
      TabIndex        =   21
      Top             =   7680
      Width           =   1065
   End
   Begin VB.Timer Tdoevents 
      Interval        =   1000
      Left            =   0
      Top             =   0
   End
   Begin VB.CommandButton btnSearch 
      Caption         =   "Search"
      Height          =   495
      Left            =   1320
      TabIndex        =   20
      Top             =   7680
      Width           =   1065
   End
   Begin VB.TextBox txtDVNo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   600
      Left            =   2505
      TabIndex        =   12
      Top             =   6750
      Width           =   5535
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4350
      Left            =   120
      TabIndex        =   9
      Top             =   2760
      Width           =   2235
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   840
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   13140
      _ExtentX        =   23178
      _ExtentY        =   1482
      ButtonWidth     =   1058
      ButtonHeight    =   1429
      Appearance      =   1
      Style           =   1
      ImageList       =   "itb32x32"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   9
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
            Caption         =   "Close"
            ImageIndex      =   7
         EndProperty
      EndProperty
      Begin MSComctlLib.ImageList itb32x32 
         Left            =   6000
         Top             =   0
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
               Picture         =   "frm_POReg.frx":0ABC
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm_POReg.frx":244E
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm_POReg.frx":3DE0
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm_POReg.frx":5772
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm_POReg.frx":7104
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm_POReg.frx":8A96
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm_POReg.frx":A428
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm_POReg.frx":BDBA
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm_POReg.frx":D74C
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm_POReg.frx":F0E0
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm_POReg.frx":FDBC
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm_POReg.frx":1069C
               Key             =   ""
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm_POReg.frx":11378
               Key             =   ""
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm_POReg.frx":12054
               Key             =   ""
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm_POReg.frx":12D30
               Key             =   ""
            EndProperty
            BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm_POReg.frx":13A0C
               Key             =   ""
            EndProperty
            BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm_POReg.frx":146E8
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.TextBox txtObR 
      Alignment       =   2  'Center
      BorderStyle     =   0  'None
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
      Left            =   120
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1920
      Width           =   5430
   End
   Begin VB.Frame fmeDetails 
      Caption         =   "Transaction Details"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3645
      Left            =   2640
      TabIndex        =   8
      Top             =   2640
      Width           =   10350
      Begin VB.TextBox txtClaimantCode 
         Height          =   285
         Left            =   2040
         TabIndex        =   35
         Top             =   240
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.TextBox txtOfficeCode 
         Height          =   195
         Left            =   0
         TabIndex        =   34
         Top             =   360
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox txtRC 
         Height          =   195
         Left            =   0
         TabIndex        =   33
         Top             =   840
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.ComboBox cmbOOE 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         ItemData        =   "frm_POReg.frx":14FC4
         Left            =   360
         List            =   "frm_POReg.frx":14FC6
         Style           =   2  'Dropdown List
         TabIndex        =   32
         Top             =   3075
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.ComboBox cmbFund 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   5640
         Style           =   2  'Dropdown List
         TabIndex        =   31
         Top             =   3075
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.ComboBox cmbRC 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   6360
         Style           =   2  'Dropdown List
         TabIndex        =   30
         Top             =   720
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CommandButton btnClaimant 
         Caption         =   "..."
         Height          =   375
         Left            =   5160
         TabIndex        =   1
         ToolTipText     =   "Click here to select claimant..."
         Top             =   660
         Width           =   375
      End
      Begin VB.TextBox txtDetail 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1155
         Left            =   225
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Top             =   1500
         Width           =   9960
      End
      Begin VB.TextBox txtOOE 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   225
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   3075
         Width           =   5160
      End
      Begin VB.TextBox txtFund 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   5610
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   3075
         Width           =   4560
      End
      Begin VB.TextBox txtOffice 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   5610
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   660
         Width           =   4560
      End
      Begin VB.TextBox txtClaimant 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   210
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   660
         Width           =   4920
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Particular"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   240
         TabIndex        =   17
         Top             =   1200
         Width           =   870
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Object of Expenditure"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   225
         TabIndex        =   16
         Top             =   2745
         Width           =   2130
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fund Type"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   5610
         TabIndex        =   15
         Top             =   2745
         Width           =   1065
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Responsibility Center"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   5640
         TabIndex        =   14
         Top             =   315
         Width           =   1980
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Claimant"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   210
         TabIndex        =   13
         Top             =   315
         Width           =   825
      End
   End
   Begin VB.Label Label23 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PO Registry"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   0
      TabIndex        =   29
      Top             =   960
      Width           =   2205
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
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
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   11040
      TabIndex        =   28
      Top             =   960
      Width           =   1395
   End
   Begin VB.Label Label10 
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
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   5640
      TabIndex        =   24
      Top             =   1680
      Width           =   4500
   End
   Begin VB.Label lblRefresh 
      BackStyle       =   0  'Transparent
      Caption         =   "Refresh"
      Height          =   495
      Left            =   120
      TabIndex        =   19
      Top             =   5640
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Enter PO Number"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   2520
      TabIndex        =   11
      Top             =   6360
      Width           =   2220
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00000000&
      FillStyle       =   0  'Solid
      Height          =   1485
      Left            =   1905
      Top             =   6330
      Width           =   11415
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Entered Transactions"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   240
      Left            =   120
      TabIndex        =   10
      Top             =   2460
      Width           =   1815
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Enter OBR No./DV Number"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   90
      TabIndex        =   6
      Top             =   1665
      Width           =   5085
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   1800
      Left            =   -855
      Top             =   720
      Width           =   11775
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mode :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   11130
      TabIndex        =   23
      Top             =   1665
      Width           =   825
   End
   Begin VB.Label lblMode 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NEW"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   12075
      TabIndex        =   22
      Top             =   1665
      Width           =   570
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   1800
      Left            =   10920
      Top             =   720
      Width           =   2115
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   5295
      Left            =   0
      Top             =   2505
      Width           =   2475
   End
End
Attribute VB_Name = "frm_POReg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Edited As Boolean
Dim msgs As String
Private Sub btnClaimant_Click()
    ActiveFormCaller = "frm_POreg"
    frmCDClaimantRegistry.Show 1
End Sub

Private Sub Form_Load()
Call Toolbar1_ButtonClick(Toolbar1.Buttons(1))
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
 Select Case Button
    Case "Close":
                    If MsgBox("Are you sure you want to close this form?", vbQuestion + vbYesNo, "System Security") = vbYes Then
                        Unload Me
                    End If
    Case "New":
             XFlag = False
             UID = ""
             DTE = ""
             Edited = False
             EditedDV = ""
             lblMode.Caption = "NEW"
            ' Call LoadTrnYear(cmb_trnYear)
             txtDate.Text = Format(Now, "MMMM dd, yyyy")
             txtObR.Text = ""
             txtClaimant.Text = ""
             txtClaimantCode.Text = ""
             txtOffice.Text = ""
             txtFund.Text = ""
             txtOOE.Text = ""
             txtDetail.Text = ""
             txtAmount.Text = ""
             txtDVNo.Text = ""
             'txtDVNo.Locked = True
             'Call LoadPrevTrans(cmb_trnYear.Text)
             txtObR.Visible = True
             txtOffice.Visible = True
             txtFund.Visible = True
             txtOOE.Visible = True

             cmbrc.Visible = False
             cmbOOE.Visible = False
             cmbFund.Visible = False
             'cmbNonAlobs.Visible = False
                          
             'Call LoadOffice
            ' Call LoadFund
             'Call LoadOOE

    Case "Save":
            Dim ct As Variant
            Dim str() As String
            Dim obrno As String
            Dim cnt As Integer
            Dim x As Integer
                'txtObR.Text = Format(txtObR.Text, "###-####-##-##-####,###-####-##-##-####")
                ct = Split(Trim(txtObR.Text), ",")
                str() = Split(Trim(txtObR.Text), ",", -1, vbTextCompare)
                cnt = UBound(ct)
                If Trim(txtClaimantCode.Text) = "" Then
                    MsgBox "Please Specify the Claimant first", vbCritical, "System Information"
                    Exit Sub
                End If
                If IsNumeric(txtAmount.Text) = False Then
                    MsgBox "Your gross Amount is not numeric Entry, Please Check your Entry..", vbCritical, "System Message"
                    Exit Sub
                End If
                       
                   
                        If txtObR.Text <> "" Then
                            If Edited = False Then
                                For x = 0 To cnt
                                    If ISAlobsAmtOkAgaintsVoucher(str(x), GetRemainingAmnt(str(x)), GetTotalTrnsactedAmt(str(x), "tblAMIS_IncomingDVTrns", "GAmount", "ObrNo"), True) = False Then
                                        Exit Sub
                                    End If
                                Next x
                            End If
                        End If
                    
                          If MsgBox("Are you sure you want to save this transaction?", vbQuestion + vbYesNo, "System Security") = vbYes Then
                              If Trim(txtClaimant.Text) <> "" And Trim(txtClaimantCode.Text) <> "" And Trim(txtDVNo.Text) <> "" And txtAmount.Text <> 0 And txtObR.Text <> "" Then
                                      Dim xChange As String
                                      
                                      xChange = txtDVNo.Text
                                      
                                      If Edited = True Then
                                          opndbaseFMIS.Execute "Update tblAMIS_IncomingDVTrns set UserID='" & UID & "," & ActiveUserID & "',Actioncode=2,DateTimeEntered='" & DTE & "," & Format(Now, "mm/dd/yyyy hh:mm:ss AMPM") & "'  Where DVNo='" & txtDVNo.Text & "' and Actioncode=10"
                                      End If

                                          If cnt > 0 Then ' more than 1 obrno
                                            If XFlag = True And txtOffice.Text = "" And txtOOE.Text = "" Then 'continuing
                                                opndbaseFMIS.Execute "Insert Into tblAMIS_IncomingDVTrns (DVNo,ObrNo,FundType,RCenter,RCenterCode,OOE,ClaimantCode,Particular,GAmount,TransactionDate,UserID,Actioncode,DateTimeEntered,Continuing,OBR2,moreobr) Values ('" & txtDVNo.Text & "','" & Left(Trim(txtObR.Text), 19) & "','" & txtFund.Text & "'," & cmbrc.ItemData(cmbrc.ListIndex) & "," & Mid(txtObR.Text, 5, 4) & ",'" & cmbOOE.Text & "','" & txtClaimantCode.Text & "','" & Trim(Replace(txtDetail.Text, "'", "''")) & "'," & CCur(txtAmount.Text) & ",'" & Format(txtDate.Text, "mm/dd/yyyy") & "','" & ActiveUserID & "',10,'" & Format(Now, "mm/dd/yyyy hh:mm:ss AMPM") & "'," & IIf(XFlag, 1, 0) & ",'" & Mid(Trim(txtObR.Text), 21, 2000) & "',1)"
                                            Else 'current
                                                opndbaseFMIS.Execute "Insert Into tblAMIS_IncomingDVTrns (DVNo,ObrNo,FundType,RCenter,RCenterCode,OOE,ClaimantCode,Particular,GAmount,TransactionDate,UserID,Actioncode,DateTimeEntered,Continuing,OBR2,moreobr) Values ('" & txtDVNo.Text & "','" & Left(Trim(txtObR.Text), 19) & "','" & txtFund.Text & "'," & txtOfficeCode.Text & "," & Mid(txtObR.Text, 5, 4) & ",'" & txtOOE.Text & "','" & txtClaimantCode.Text & "','" & Trim(Replace(txtDetail.Text, "'", "''")) & "'," & CCur(txtAmount.Text) & ",'" & Format(txtDate.Text, "mm/dd/yyyy") & "','" & ActiveUserID & "',10,'" & Format(Now, "mm/dd/yyyy hh:mm:ss AMPM") & "'," & IIf(XFlag, 1, 0) & ",'" & Mid(Trim(txtObR.Text), 21, 2000) & "',1)"
                                            End If
                                          Else ' 1 obrno
                                            If XFlag = True And txtOffice.Text = "" And txtOOE.Text = "" Then 'continuing
                                                opndbaseFMIS.Execute "Insert Into tblAMIS_IncomingDVTrns (DVNo,ObrNo,FundType,RCenter,RCenterCode,OOE,ClaimantCode,Particular,GAmount,TransactionDate,UserID,Actioncode,DateTimeEntered,Continuing) Values ('" & txtDVNo.Text & "','" & Trim(txtObR.Text) & "','" & txtFund.Text & "'," & cmbrc.ItemData(cmbrc.ListIndex) & "," & Mid(txtObR.Text, 5, 4) & ",'" & cmbOOE.Text & "','" & txtClaimantCode.Text & "','" & Trim(Replace(txtDetail.Text, "'", "''")) & "'," & CCur(txtAmount.Text) & ",'" & Format(txtDate.Text, "mm/dd/yyyy") & "','" & ActiveUserID & "',10,'" & Format(Now, "mm/dd/yyyy hh:mm:ss AMPM") & "'," & IIf(XFlag, 1, 0) & ")"
                                            Else 'current
                                                opndbaseFMIS.Execute "Insert Into tblAMIS_IncomingDVTrns (DVNo,ObrNo,FundType,RCenter,RCenterCode,OOE,ClaimantCode,Particular,GAmount,TransactionDate,UserID,Actioncode,DateTimeEntered,Continuing) Values ('" & txtDVNo.Text & "','" & Trim(txtObR.Text) & "','" & txtFund.Text & "'," & txtOfficeCode.Text & "," & Mid(txtObR.Text, 5, 4) & ",'" & txtOOE.Text & "','" & txtClaimantCode.Text & "','" & Trim(Replace(txtDetail.Text, "'", "''")) & "'," & CCur(txtAmount.Text) & ",'" & Format(txtDate.Text, "mm/dd/yyyy") & "','" & ActiveUserID & "',10,'" & Format(Now, "mm/dd/yyyy hh:mm:ss AMPM") & "'," & IIf(XFlag, 1, 0) & ")"
                                            End If
                                          End If
                                      
                                      Call Toolbar1_ButtonClick(Toolbar1.Buttons(1))
                              Else
                                  MsgBox "Save operation cancelled!, Please Check your Entry.." & vbCrLf & vbCrLf & "Please check your entry.", vbExclamation + vbOKOnly, "System Security"
                              End If
                          End If
                        
    Case "Delete":
                    If Edited = True Then
                        If MsgBox("Are you sure you want to delete this transaction?", vbQuestion + vbYesNo) = vbYes Then
                            opndbaseFMIS.Execute "Update tblAMIS_IncomingDVTrns set UserID='" & UID & "," & ActiveUserID & "',Actioncode=3,DateTimeEntered='" & DTE & "," & Format(Now, "mm/dd/yyyy hh:mm:ss AMPM") & "'  Where DVNo='" & txtDVNo.Text & "' and Actioncode=10"
                            Call Toolbar1_ButtonClick(Toolbar1.Buttons(1))
                        End If
                    End If
    End Select

End Sub
Private Function ChkEntry() As Boolean
ChkEntry = False

If Trim(txtClaimant.Text) <> "" And Trim(txtClaimantCode.Text) <> "" And Trim(txtDVNo.Text) <> "" Then
    
    If cmbFund.Visible = True Then 'non obr
        If cmbNonAlobs.ListIndex <> -1 And cmbFund.Text <> "" And cmbrc.ListIndex <> -1 And cmbOOE.Text <> "" Then
            ChkEntry = True
        Else
            ChkEntry = False
        End If
    Else 'with obr
        If cmbrc.Visible = True Then 'continuing
            If txtFund.Text <> "" And cmbrc.ListIndex <> -1 And cmbOOE.Text <> "" Then
                ChkEntry = True
            Else
                ChkEntry = False
            End If
        Else 'current
            If txtFund.Text <> "" And txtOfficeCode.Text <> "" And txtOOE.Text <> "" And txtClaimantCode.Text <> "" And txtDVNo.Text <> "" Then
                ChkEntry = True
            Else
                ChkEntry = False
            End If
        End If
    End If
Else
    ChkEntry = True
End If

End Function
Private Sub txtObR_KeyPress(KeyAscii As Integer)
Dim sDVNo As String
Dim rec As New ADODB.Recordset
    If KeyAscii = 13 Then
        If Len(txtObR.Text) = 10 Then
            Set rec = opndbaseFMIS.Execute("SELECT dvno From [fmis].[dbo].[tblAMIS_IncomingDVTrns] where dvno = '" & txtObR.Text & "' and actioncode = 10")
            If rec.RecordCount > 0 Then
                If AlreadyOut(txtObR.Text) Then
                    MsgBox "This DV number is already out!", vbExclamation + vbOKOnly, "System Securty"
                Else
                    lblMode.Caption = "EDIT"
                    lblMode.ForeColor = &HFF&
                    Call ReLoadDetail(txtObR.Text)
'                    AllLoadCAdetails (txtDVNo.Text)
                End If
            Else
                lblMode.Caption = "NEW"
                MsgBox "No record Found..!", vbInformation, "System Message"
            End If
        ElseIf Len(txtObR.Text) = 19 Then
            lblMode.Caption = "NEW"
            Call GetObRData(txtObR.Text)
            txtAmount.Text = Format(GetRemainingAmnt(txtObR.Text), "#,##0.00")
            If txtAmount.Text = "0.00" Then
                MsgBox "Obr\Alobs Number have 0 balance,Cannot Procces the Transaction..", vbInformation, "System Message"
                Exit Sub
            End If
        ElseIf Len(txtObR.Text) > 19 Then
            Dim ct As Variant
            Dim str() As String
            Dim Particular As String
            Dim obrno As String
            Dim cnt As Integer
            Dim x As Integer
                'txtObR.Text = Format(txtObR.Text, "###-####-##-##-####,###-####-##-##-####")
                ct = Split(txtObR.Text, ",")
                str() = Split(txtObR.Text, ",", -1, vbTextCompare)
                cnt = UBound(ct)
                txtAmount.Text = ""
                txtDetail.Text = ""
                Particular = ""
                txtClaimant.Text = ""
                For x = 0 To cnt
                    obrno = Trim(Format(str(x), "###-####-##-##-####"))
                    If ValidObR(obrno) Then
                        lblMode.Caption = "NEW"
                        lblMode.ForeColor = &HFF0000
                        Particular = Trim(txtDetail.Text)
                        Call GetObRData(obrno)
                        
                        If x > 0 Then
                        txtDetail.Text = Particular & " and " & Trim(txtDetail.Text)
                        txtObR.Text = txtObR.Text & "," & obrno
                        Else
                        txtObR.Text = obrno
                        End If
                        txtAmount.Text = Format(CCur(GetRemainingAmnt(obrno)) + CCur(IIf(IsNumeric(txtAmount.Text) = True, txtAmount.Text, "0.00")), "#,##0.00")
                        If txtAmount.Text = "0.00" Then
                            MsgBox "Obr\Alobs Number have 0 balance,Cannot Procces the Transaction..", vbInformation, "System Message"
                           
                            Exit For
                        End If
                    Else
                        MsgBox "Invalid ObR!" & x, vbExclamation + vbOKOnly, "System Securty"
                    End If
                Next x
        End If
    
'
'
'
'
'
'
'
'        lblMode.Caption = "NEW"
'        lblMode.ForeColor = &HFF0000
'        If Trim(txtObR.Text) <> "" Then
'            txtObR.Text = Trim(Replace(txtObR.Text, "-", ""))
'            If Len(txtObR.Text) = 15 Then ' one obr no
'                txtObR.Text = Format(txtObR.Text, "###-####-##-##-####")
'                If ValidObR(txtObR.Text) Then
'                        lblMode.Caption = "NEW"
'                        lblMode.ForeColor = &HFF0000
'                        Call GetObRData(txtObR.Text)
'                        txtAmount.Text = Format(GetRemainingAmnt(txtObR.Text), "#,##0.00")
'
'                    'End If
'                Else
'                    MsgBox "Invalid ObR!", vbExclamation + vbOKOnly, "System Securty"
'                End If
'
'
'
'            ElseIf Len(txtObR.Text) > 15 Then 'more than 1 obrno
'
'
'                    'End If
'
'
'            ElseIf Len(txtObR.Text) = 10 Then 'PO number
'
'            Else
'                MsgBox "Invalid ObR / DV No!", vbExclamation + vbOKOnly, "System Securty"
'            End If
'        Else
'            MsgBox "Input ObR first!", vbExclamation + vbOKOnly, "System Securty"
'        End If
    End If
End Sub
Private Sub ReLoadDetail(ByVal POnumber As String)
Dim DVRec As New ADODB.Recordset
    On Error Resume Next
    XFlag = False
    
    DVRec.Open ("Select * from tblAMIS_IncomingDVTrns where DVNo='" & POnumber & "' and Actioncode=10"), opndbaseFMIS, adOpenStatic, adLockOptimistic
    If DVRec.RecordCount > 0 Then
        Edited = True
        frmTrans.Enabled = False
        If DVRec!Continuing = 1 Then
            XFlag = True
        End If
        EditedDV = POnumber
        lblMode.Caption = "EDIT"
        Label14.Visible = False
            optObR.Value = True
            If DVRec!moreobr = 1 Then
            txtObR.Text = Trim(DVRec![obrno]) & "," & Trim(DVRec!obr2)
            Else
            txtObR.Text = Trim(DVRec![obrno])
            End If
            txtOfficeCode.Text = DVRec![RCenter]
            txtOffice.Text = GetOfficeName(DVRec![RCenter], "OfficeMedium")
            txtFund.Text = DVRec![FundType]
            txtOOE.Text = IIf(IsNull(DVRec![OOE]), "", (DVRec![OOE]))
            
        'End If
        txtClaimantCode = IIf(IsNull(DVRec!ClaimantCode), "", DVRec!ClaimantCode)
        txtClaimant.Text = getClaimant(IIf(IsNull(DVRec!ClaimantCode), "", DVRec!ClaimantCode))
        txtDetail.Text = DVRec![Particular]
        txtAmount.Text = Format(DVRec![Gamount], "#,###.00")
        txtDVNo.Text = DVRec![dvno]
        txtDate.Text = Format(DVRec![TransactionDate], "mmmm dd, yyyy")
        DTE = DVRec![datetimeentered]
        UID = DVRec![UserID]
    Else
        MsgBox "Invalid PO Number!", vbExclamation + vbOKOnly, "System Security"
        Call Toolbar1_ButtonClick(Toolbar1.Buttons(1))
    End If
    DVRec.Close
    Set DVRec = Nothing
    
    'opndbaseFMIS.Execute "Insert Into tblAMIS_IncomingDVTrns (DVNo,ObrNo,FundType,RCenter,RCenterCode,OOE,ClaimantCode,Particular,GAmount,TransactionDate,UserID,Actioncode,DateTimeEntered) Values ('" & txtDVNo.Text & "','" & Trim(txtObR.Text) & "','" & txtFund.Text & "'," & txtOfficeCode.Text & "," & Mid(txtObR.Text, 5, 4) & ",'" & txtOOE.Text & "','" & txtClaimantCode.Text & "','" & txtDetail.Text & "'," & CCur(txtAmount.Text) & ",'" & Format(txtDate.Text, "mm/dd/yyyy") & "','" & ActiveUserID & "',1,'" & Format(Now, "mm/dd/yyyy hh:mm:ss AMPM") & "')"
    
End Sub

Private Function AlreadyOut(ByVal POnumber As String) As Boolean
Dim AORec As New ADODB.Recordset

    AlreadyOut = False
    AORec.Open ("Select outby From [tblAMIS_IncomingDVTrns] where DVNo='" & dvno & "' and ActionCode=10 and ReturnFlag=0"), opndbaseFMIS, adOpenStatic, adLockOptimistic
    If AORec.RecordCount > 0 Then
        If Trim(AORec![OutBy]) <> "" Then
            AlreadyOut = True
        End If
    End If
    AORec.Close
    Set AORec = Nothing
    
End Function
Private Sub GetObRData(ByVal ObR As String)
Dim OREc As New ADODB.Recordset
Dim OName As String
Dim OCode As Integer
Dim OOE As String

XFlag = False

OREc.Open ("Select * From tblFMIS_Transaction Where AlobsNo='" & ObR & "' And ActionCode=1"), opndbaseFMIS, adOpenStatic, adLockOptimistic
If OREc.RecordCount > 0 Then
    txtOffice.Text = GetOfficeName(OREc!FMISOfficeCode, "OfficeMedium")
    txtOfficeCode.Text = OREc!FMISOfficeCode
    If Mid(ObR, 1, 3) = "118" Then
        txtFund.Text = "20% DF"
    ElseIf Mid(ObR, 1, 3) = "101" Then
        txtFund.Text = "GF-Proper"
    Else
        txtFund.Text = OREc!FundType
    End If
    txtOOE.Text = OREc!OOE
    txtDetail.Text = OREc!Particulars
    'txtAmount.Text = Format(OREc!Amount, "###,##0.00")
End If
OREc.Close
Set OREc = Nothing

OREc.Open ("Select * From [tblBMS_ExcessControl] Where AlobsNo='" & ObR & "' And ActionCode=1"), opndbaseFMIS, adOpenStatic, adLockOptimistic
If OREc.RecordCount > 0 Then
    XFlag = True
    Call GetObROffice(OREc!AlobsNoCharge, OName, OCode, OOE)
    txtOffice.Text = OName
    txtOfficeCode.Text = OCode
    If Mid(ObR, 1, 3) = "118" Then
        txtFund.Text = "20% DF"
    ElseIf Mid(ObR, 1, 3) = "101" Then
        txtFund.Text = "GF-Proper"
    Else
        txtFund.Text = GetFundMedium(Mid(ObR, 1, 3))
    End If
    txtOOE.Text = OOE
    txtDetail.Text = OREc![Details]
    'txtAmount.Text = Format(OREc![Amount], "###,##0.00")
End If
OREc.Close
Set OREc = Nothing

If XFlag = True And txtOffice.Text = "" And txtOOE.Text = "" Then
    
    txtOffice.Visible = False
    cmbrc.Width = txtOffice.Width
    cmbrc.Left = txtOffice.Left
    cmbrc.Top = txtOffice.Top
    cmbrc.Visible = True

    txtOOE.Visible = False
    cmbOOE.Width = txtOOE.Width
    cmbOOE.Left = txtOOE.Left
    cmbOOE.Top = txtOOE.Top
    cmbOOE.Visible = True

End If
End Sub

Private Sub GetObROffice(ByVal ObR As String, OName As String, OCode As Integer, OOE As String)
Dim GORec As New ADODB.Recordset

    OName = ""
    OCode = 0
    OOE = ""
    GORec.Open ("Select * From tblFMIS_Transaction Where AlobsNo='" & Replace(ObR, "'", "''") & "' And ActionCode=1"), opndbaseFMIS, adOpenStatic, adLockOptimistic
    If GORec.RecordCount > 0 Then
        OName = GetOfficeName(GORec!FMISOfficeCode, "OfficeMedium")
        OCode = GORec!FMISOfficeCode
        OOE = GORec!OOE
    End If
    GORec.Close
    Set GORec = Nothing
    
End Sub
Private Sub LoadPrevTrans(ByVal YEAR_ As Integer)
Dim PRec As New ADODB.Recordset
Dim x As Integer

List1.Clear
List1.Enabled = False

PRec.Open ("Select trnno,dvno From tblAMIS_IncomingDVTrns Where TransactionDate like '%" & YEAR_ & "' And ActionCode=10 And [PAout]=0 Order by trnno desc"), opndbaseFMIS, adOpenStatic, adLockOptimistic
If PRec.RecordCount > 0 Then
    For x = 1 To PRec.RecordCount
        List1.AddItem PRec!dvno
        List1.ItemData(List1.NewIndex) = PRec!Trnno
        PRec.MoveNext
    Next x
    List1.Enabled = True
End If
PRec.Close
Set PRec = Nothing
End Sub

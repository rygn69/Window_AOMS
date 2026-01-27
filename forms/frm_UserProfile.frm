VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm_UserProfile 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Claimant Entry Form"
   ClientHeight    =   5700
   ClientLeft      =   1920
   ClientTop       =   2280
   ClientWidth     =   8070
   Icon            =   "frm_UserProfile.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5700
   ScaleWidth      =   8070
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
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   7755
      Begin VB.PictureBox pic_EmployeeIndividual 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   3930
         Left            =   240
         ScaleHeight     =   3900
         ScaleWidth      =   7230
         TabIndex        =   1
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
            Left            =   360
            TabIndex        =   11
            Top             =   2520
            Width           =   4245
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
            Left            =   360
            TabIndex        =   9
            Top             =   3240
            Width           =   4245
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
            Left            =   390
            TabIndex        =   7
            Top             =   480
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
            TabIndex        =   6
            Top             =   1800
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
            TabIndex        =   5
            Top             =   1155
            Width           =   4200
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Contact Nos."
            Height          =   195
            Left            =   360
            TabIndex        =   10
            Top             =   3000
            Width           =   930
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Password:"
            Height          =   195
            Left            =   375
            TabIndex        =   8
            Top             =   2280
            Width           =   735
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ID"
            Height          =   195
            Left            =   375
            TabIndex        =   4
            Top             =   255
            Width           =   165
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Position:"
            Height          =   195
            Left            =   375
            TabIndex        =   3
            Top             =   1560
            Width           =   600
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Name:"
            Height          =   195
            Left            =   375
            TabIndex        =   2
            Top             =   930
            Width           =   465
         End
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Height          =   810
      Left            =   0
      TabIndex        =   12
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
            Picture         =   "frm_UserProfile.frx":09EA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_UserProfile.frx":0A48
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_UserProfile.frx":0AA6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_UserProfile.frx":0B04
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
            Picture         =   "frm_UserProfile.frx":0B62
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_UserProfile.frx":24F4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_UserProfile.frx":3E86
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_UserProfile.frx":5818
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_UserProfile.frx":71AA
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_UserProfile.frx":8B3C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_UserProfile.frx":A4CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_UserProfile.frx":BE60
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_UserProfile.frx":D7F2
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_UserProfile.frx":F186
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_UserProfile.frx":FE62
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_UserProfile.frx":10742
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_UserProfile.frx":1141E
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_UserProfile.frx":120FA
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_UserProfile.frx":12DD6
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_UserProfile.frx":13AB2
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_UserProfile.frx":1478E
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frm_UserProfile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
'Select Case Button.Index
'    Case 1 'New
'        saveflag = 1 'New
'        transNo = 0
'
'        Frame2.Enabled = True
'        'txt_LastName.SetFocus
'    Case 3 'Save
'        Call SaveEntry
'
'    Case 5 'Edit
'        If Classification = "Individual" Or Classification = "Company" Or Classification = "National" Or Classification = "BarangayTreasurer" Or Classification = "MunicipalTreasurer" Then
'            saveflag = 2 'Edit
'            transNo = List1.ItemData(List1.ListIndex)
'            Frame2.Enabled = True
'        End If
'    Case 7 'Delete
'        Frame2.Enabled = True
'        If Classification = "Individual" Or Classification = "Company" Or Classification = "National" Or Classification = "BarangayTreasurer" Or Classification = "MunicipalTreasurer" Then
'            transNo = List1.ItemData(List1.ListIndex)
'            If MsgBox("Are You Sure Want to Delete this Record?", vbQuestion + vbYesNo, "System Confirmation Query") = vbYes Then
'                If CodeUsed(List1.List(List1.ListIndex)) = False Then
'                    opndbaseFMIS.Execute "Delete from tblCMS_CDClaimantDetails where trnno=" & transNo & ""
'                    MsgBox "Deleting Record Successful!", vbInformation, "System Information"
'                    Call Clear
'                    saveflag = 0 'No action
'                    transNo = 0
'                Else
'                    MsgBox "Claimant is in used!" & Chr(13) & Chr(13) & "Deletion is not Allowed!", vbInformation, "System Information"
'                    saveflag = 0 'No action
'                    transNo = 0
'                End If
'            End If
'        End If
'End Select
'End Sub

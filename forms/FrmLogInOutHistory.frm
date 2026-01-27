VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{1693405E-2DC9-4248-B52F-4AC9145DA2AF}#1.0#0"; "WINXPC~1.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form FrmLogInOutHistory 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Log In/Out History Form"
   ClientHeight    =   9225
   ClientLeft      =   -1560
   ClientTop       =   3750
   ClientWidth     =   10155
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9225
   ScaleWidth      =   10155
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6360
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmLogInOutHistory.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmLogInOutHistory.frx":005E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin WinXPC_Engine.WindowsXPC WindowsXPC1 
      Left            =   6255
      Top             =   2925
      _ExtentX        =   6588
      _ExtentY        =   1085
      ColorScheme     =   2
      Common_Dialog   =   0   'False
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   8820
      Top             =   180
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   6750
      Left            =   375
      TabIndex        =   0
      Top             =   2295
      Width           =   9435
      _ExtentX        =   16642
      _ExtentY        =   11906
      _Version        =   393216
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   2
      SelectionMode   =   1
      AllowUserResizing=   1
      Appearance      =   0
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   600
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   10155
      _ExtentX        =   17912
      _ExtentY        =   1058
      ButtonWidth     =   1482
      ButtonHeight    =   1005
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Open"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "slash"
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Refresh"
            ImageIndex      =   2
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
      EndProperty
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000009&
      X1              =   2370
      X2              =   2370
      Y1              =   1530
      Y2              =   1845
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000009&
      X1              =   2370
      X2              =   2370
      Y1              =   870
      Y2              =   1185
   End
   Begin VB.Label lbl_Position 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   240
      Left            =   2820
      TabIndex        =   5
      Top             =   1830
      Width           =   570
   End
   Begin VB.Label lbl_Fullname 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   240
      Left            =   2820
      TabIndex        =   4
      Top             =   1155
      Width           =   570
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Position"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   2445
      TabIndex        =   3
      Top             =   1545
      Width           =   555
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fullname"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   2445
      TabIndex        =   2
      Top             =   885
      Width           =   630
   End
   Begin VB.Image Image1 
      Height          =   1425
      Left            =   555
      Stretch         =   -1  'True
      Top             =   795
      Width           =   1320
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H80000001&
      BackStyle       =   1  'Opaque
      Height          =   1605
      Left            =   375
      Top             =   690
      Width           =   9435
   End
End
Attribute VB_Name = "FrmLogInOutHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim LogFileLocation As String

Private Sub LoadLogHistory()
Dim LoadHistory As Variant
Dim sss As Integer

On Error GoTo handler

MSFlexGrid1.Clear
MSFlexGrid1.Cols = 7
MSFlexGrid1.Rows = 500
MSFlexGrid1.FormatString = "^USER ID  |^ACCESS LEVEL               |^COMPUTER NAME       |^OFFICE ID    |^DATE                 |^TIME               |^MODE       "

    MSFlexGrid1.Row = 1
    MSFlexGrid1.col = 0
    MSFlexGrid1.ColSel = MSFlexGrid1.Cols - 1
    MSFlexGrid1.RowSel = MSFlexGrid1.Rows - 1
    MSFlexGrid1.Clip = mydll.OpenTxtFile(LogFileLocation)
    
    '----this is used to clear the character that represent "enter" at the userid-------------------
    For sss = 1 To MSFlexGrid1.Rows - 1
        If Len(Trim(MSFlexGrid1.TextMatrix(sss, 0))) > 4 Then
            MSFlexGrid1.TextMatrix(sss, 0) = Mid(MSFlexGrid1.TextMatrix(sss, 0), 2, Len(MSFlexGrid1.TextMatrix(sss, 0)) - 1)
        ElseIf Len(Trim(MSFlexGrid1.TextMatrix(sss, 0))) = 1 Then
            MSFlexGrid1.TextMatrix(sss, 0) = ""
        End If
    Next sss '---------------------------------------------------------------------------------------


    '------Reformating 3 digit userid to 4 digit------------------------------
    For sss = 1 To MSFlexGrid1.Rows - 1
        If Len(Trim(MSFlexGrid1.TextMatrix(sss, 0))) <> 0 Then 'verify if row has a context
            If Len(Trim(MSFlexGrid1.TextMatrix(sss, 0))) < 4 Then
                MSFlexGrid1.TextMatrix(sss, 0) = Format(MSFlexGrid1.TextMatrix(sss, 0), "0000")
            End If
        End If
    Next sss '-----------------------------------------------------------------


handler:
If err.Number <> 0 Then
    MsgBox "Log History for this Date is not available!", vbInformation, "System Information"
End If
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then
    Unload Me
End If
End Sub


Private Sub Form_Load()
Me.Top = (Screen.Height - Me.Height) / 2
Me.Left = 3650

WindowsXPC1.InitSubClassing
LogFileLocation = LogLocation & "\ActivityLog" & Month(Date) & Year(Date) & ".ini"

Call LoadLogHistory
End Sub

Private Sub Form_Unload(Cancel As Integer)
'frmMother.StatusBar1.Panels(3) = "Active Module :"
Set FrmLogInOutHistory = Nothing
WindowsXPC1.EndWinXPCSubClassing
End Sub

Private Sub MSFlexGrid1_Click()
Dim opnEmpDetails As New ADODB.Recordset

On Error GoTo handler

opnEmpDetails.Open "Select * from pmis.dbo.employee where swipemployeeid='" & MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 0) & "'", opndbaseFMIS, adOpenStatic, adLockOptimistic
If opnEmpDetails.RecordCount <> 0 Then
    lbl_Fullname.Caption = opnEmpDetails!Lastname & ", " & opnEmpDetails!Firstname & IIf(Len(opnEmpDetails!MI) = 0 Or IsNull(opnEmpDetails!MI), " ", " " & UCase(Left(opnEmpDetails!MI, 1)) & ".") & IIf(Len(opnEmpDetails!Suffix) = 0 Or IsNull(opnEmpDetails!Suffix), "", ", " & opnEmpDetails!Suffix)
    lbl_Position.Caption = IIf(IsNull(opnEmpDetails!Position), "", opnEmpDetails!Position)
                    
    If IsNull(opnEmpDetails!photo) Or Len(opnEmpDetails!photo) = 0 Then
        Image1.Visible = False
    Else
        Image1.Visible = True
        Image1.Picture = LoadPicture(PicLocation & opnEmpDetails!photo)
    End If
Else
    lbl_Fullname.Caption = ""
    lbl_Position.Caption = ""
    Image1.Visible = False

End If
opnEmpDetails.Close
Set opnEmpDetails = Nothing

handler:
If err.Number <> 0 Then
    Image1.Visible = False
End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
    Case 1 'Open
        CommonDialog1.Flags = cdlOFNNoChangeDir 'Forcing the dialogue box to return to the current openned directory
        CommonDialog1.Flags = cdlOFNHideReadOnly ' Set the dialogue box to hide the Read Only Check box
        CommonDialog1.Filter = "Windows Ini Files (*.ini)|*.ini|"  ' Set filters
        CommonDialog1.FilterIndex = 1 ' Specify default filter
        CommonDialog1.FileName = LogLocation & "\ActivityLog" & Month(Date) & Year(Date) & ".ini"
        CommonDialog1.ShowOpen ' Display the Open dialog box
        LogFileLocation = CommonDialog1.FileName
        Call LoadLogHistory
    Case 3 'Refresh
        Unload Me
        Sleep 500
        FrmLogInOutHistory.Show 'vbModal
End Select
End Sub

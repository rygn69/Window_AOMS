VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{1693405E-2DC9-4248-B52F-4AC9145DA2AF}#1.0#0"; "WinXPC Engine.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmJEVNumbering 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "JEV Numbering for Check Disbursement Report"
   ClientHeight    =   7440
   ClientLeft      =   150
   ClientTop       =   1380
   ClientWidth     =   15990
   Icon            =   "frmJEVNumbering.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7440
   ScaleWidth      =   15990
   Begin VB.TextBox txt_compositioncode 
      Height          =   285
      Left            =   7560
      TabIndex        =   22
      Top             =   7800
      Width           =   735
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Search By:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1140
      Left            =   120
      TabIndex        =   16
      Top             =   120
      Width           =   3375
      Begin VB.ComboBox Combo2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         ItemData        =   "frmJEVNumbering.frx":0E42
         Left            =   240
         List            =   "frmJEVNumbering.frx":0E4C
         TabIndex        =   17
         Top             =   480
         Width           =   2910
      End
   End
   Begin VB.PictureBox pic_details 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4950
      Left            =   0
      ScaleHeight     =   4920
      ScaleWidth      =   15870
      TabIndex        =   14
      Top             =   2400
      Width           =   15900
      Begin MSFlexGridLib.MSFlexGrid grd_details 
         Height          =   4920
         Left            =   0
         TabIndex        =   15
         Top             =   0
         Width           =   15870
         _ExtentX        =   27993
         _ExtentY        =   8678
         _Version        =   393216
         FixedCols       =   0
         ForeColorFixed  =   4210688
         BackColorBkg    =   0
         ScrollTrack     =   -1  'True
         FocusRect       =   0
         SelectionMode   =   1
         AllowUserResizing=   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Post (JEV No)"
      Height          =   1020
      Left            =   11760
      TabIndex        =   12
      Top             =   1200
      Width           =   2475
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Load"
      Height          =   720
      Left            =   8400
      TabIndex        =   7
      Top             =   8400
      Width           =   1725
   End
   Begin VB.Frame Frame4 
      Caption         =   "Mass JEV Numbering"
      ClipControls    =   0   'False
      Height          =   960
      Left            =   10800
      TabIndex        =   5
      Top             =   8160
      Width           =   2820
      Begin VB.CommandButton Command1 
         Caption         =   "Set JEV Nos."
         Height          =   435
         Left            =   375
         TabIndex        =   6
         Top             =   345
         Width           =   2070
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Source Details"
      Height          =   2250
      Left            =   6000
      TabIndex        =   0
      Top             =   0
      Width           =   5655
      Begin VB.TextBox txt_RCI 
         Height          =   360
         Left            =   375
         TabIndex        =   20
         Top             =   1800
         Width           =   2385
      End
      Begin VB.TextBox txt_acctName 
         Height          =   360
         Left            =   2985
         TabIndex        =   11
         Top             =   1170
         Width           =   2385
      End
      Begin VB.TextBox txt_AcctNo 
         Height          =   360
         Left            =   2985
         TabIndex        =   10
         Top             =   540
         Width           =   2385
      End
      Begin VB.TextBox txt_bank 
         Height          =   360
         Left            =   345
         TabIndex        =   9
         Top             =   1170
         Width           =   2385
      End
      Begin VB.TextBox txt_fund 
         Height          =   360
         Left            =   345
         TabIndex        =   8
         Top             =   555
         Width           =   2385
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "RCI Number"
         Height          =   195
         Left            =   240
         TabIndex        =   21
         Top             =   1560
         Width           =   870
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Account Number"
         Height          =   195
         Left            =   2835
         TabIndex        =   4
         Top             =   315
         Width           =   1200
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Account Name"
         Height          =   195
         Left            =   2835
         TabIndex        =   3
         Top             =   885
         Width           =   1065
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Drawee Bank"
         Height          =   195
         Left            =   210
         TabIndex        =   2
         Top             =   930
         Width           =   975
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Special Account"
         Height          =   195
         Left            =   210
         TabIndex        =   1
         Top             =   330
         Width           =   1170
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   1545
      Top             =   4845
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4920
      Top             =   6120
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
            Picture         =   "frmJEVNumbering.frx":0E5F
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJEVNumbering.frx":1EE1
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJEVNumbering.frx":401B
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJEVNumbering.frx":4079
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJEVNumbering.frx":4393
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin WinXPC_Engine.WindowsXPC WindowsXPC1 
      Left            =   255
      Top             =   9090
      _ExtentX        =   6588
      _ExtentY        =   1085
      ColorScheme     =   2
      Common_Dialog   =   0   'False
      PictureControl  =   0   'False
   End
   Begin VB.Frame Frame6 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1020
      Left            =   120
      TabIndex        =   18
      Top             =   1080
      Width           =   5655
      Begin VB.TextBox txt_Search 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   480
         Left            =   240
         TabIndex        =   19
         Top             =   360
         Width           =   5220
      End
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label14"
      Height          =   195
      Left            =   13725
      TabIndex        =   13
      Top             =   9705
      Width           =   570
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   2235
      Left            =   15
      Top             =   0
      Width           =   5850
   End
End
Attribute VB_Name = "frmJEVNumbering"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim tmpCompositionCode As Long
Dim NewTransaction As Boolean
Dim RCILimit As Integer
Dim TmpRCILimit As Variant
Dim tmpRCINo As String
Dim MsgType As Integer
Dim tmpCheckNo As String
Dim tmpcomposition As String


Private Function ResponsibilityCenterCode(ByVal MixCode As String) As String
Dim opnRC As New ADODB.Recordset

If InStr(Trim(MixCode), "FMISNo-") <> 0 Then 'CAsh Advance------------------
    ResponsibilityCenterCode = "1091" 'Automatic PTO is the Responsibility Center -
Else
    opnRC.Open "Select alobsNo from tblCMS_CDTransactionDetails where ControlNo='" & MixCode & "' and actioncode=1", opndbaseFMIS, adOpenStatic, adLockOptimistic
        If opnRC.RecordCount <> 0 Then
            If InStr(opnRC!AlobsNo, "NA-") <> 0 Then 'None Alobs Transactions-----
                ResponsibilityCenterCode = GetNoneAlobsName(opnRC!AlobsNo)
            ElseIf Trim(opnRC!AlobsNo) = "cash advance" Then 'Special Accomodation for Lalang Dubduban---
                ResponsibilityCenterCode = "1091" 'Automatic PTO is the Responsibility Center -
            ElseIf Trim(opnRC!AlobsNo) = "BDH" Then 'Special Accomodation for Lalang Dubduban---
                ResponsibilityCenterCode = "4421" 'Automatic PTO is the Responsibility Center -
            ElseIf Trim(opnRC!AlobsNo) = "PROVINCIAL AID" Then 'Special Accomodation for Provincial Aid---
                ResponsibilityCenterCode = "9997" 'Automatic PTO is the Responsibility Center -
            Else 'For With Alobs Transactions--------------------------------------
                ResponsibilityCenterCode = GetFinalResCode(opnRC!AlobsNo)
            End If
        End If
    opnRC.Close
    Set opnRC = Nothing
End If

End Function
Private Function GetFinalResCode(ByVal AlobsNo As String) As String
Dim tmpval As Variant

If InStr(Trim(AlobsNo), ",") <> 0 Then 'Multiple Alobs----
    tmpval = Split(AlobsNo, ",")    '----Spliting from multiple
    tmpval = tmpval(0)              '----To Single
    
    tmpval = Split(tmpval, "-")     '----Split into Codes----
    GetFinalResCode = tmpval(1) 'Final Responsibility Code---
Else 'Single Alobs----------------------------------
    tmpval = Split(AlobsNo, "-")     '----Split into Codes----
    GetFinalResCode = tmpval(1) 'Final Responsibility Code---
End If
End Function
Private Function GetNoneAlobsName(ByVal NACode As String) As String
Dim opnNA As New ADODB.Recordset

opnNA.Open "Select NonAlobs from tblCMS_CDNoneAlobs where NACode='" & Trim(NACode) & "'", opndbaseFMIS, adOpenStatic, adLockOptimistic
If opnNA.RecordCount <> 0 Then
    GetNoneAlobsName = opnNA!NonAlobs
End If
opnNA.Close
Set opnNA = Nothing
End Function

Private Sub LoadSavedReportNos(ByVal trnMonth As Integer, ByVal TrnYear As Integer, ByVal fund As String)
Dim opnRCINo As New ADODB.Recordset
Dim SQL As String
Dim cc As Integer

SQL = "SELECT  tblCMS_CDRCIReport.RCINo as RCINo, tblCMS_CDRCIReport.Compositioncode as Compositioncode, vw_DepositoryBank.FundType,tblCMS_CDRCIReport.AlreadySaved2JEV " & _
        " FROM tblCMS_CDRCIReport LEFT OUTER JOIN " & _
        " vw_DepositoryBank ON tblCMS_CDRCIReport.Compositioncode = vw_DepositoryBank.FMISAccountCode " & _
        " Where (tblCMS_CDRCIReport.ActionCode = 1) " & _
        " GROUP BY tblCMS_CDRCIReport.RCINo, tblCMS_CDRCIReport.Compositioncode, vw_DepositoryBank.FundType, " & _
        " tblCMS_CDRCIReport.AlreadySaved2JEV , Year(tblCMS_CDRCIReport.CheckDate), Month(tblCMS_CDRCIReport.CheckDate) " & _
        " HAVING (vw_DepositoryBank.FundType = '" & fund & "') AND (YEAR(tblCMS_CDRCIReport.CheckDate) = " & TrnYear & ") AND " & _
        " (MONTH(tblCMS_CDRCIReport.CheckDate) = " & trnMonth & ") AND (tblCMS_CDRCIReport.AlreadySaved2JEV = 0)" & _
        " ORDER BY tblCMS_CDRCIReport.RCINo "

'opnRCINo.Open "Select RCINo from tblCMS_CDRCIReport where Year(CheckDate)=" & TrnYear & " and Month(CheckDate)=" & trnMonth & " and Compositioncode=" & compositioncode & " group by RCINO order by RCINo ", opndbaseFMIS, adOpenStatic, adLockOptimistic
'Debug.Print sql

opnRCINo.Open SQL, opndbaseFMIS, adOpenStatic, adLockOptimistic



If opnRCINo.RecordCount <> 0 Then
    List1.Clear
    Do Until opnRCINo.EOF
        List1.AddItem (opnRCINo!RCINo)
        List1.ItemData(cc) = opnRCINo!compositioncode
        cc = cc + 1
        opnRCINo.MoveNext
    Loop
Else
    List1.Clear
End If
opnRCINo.Close
Set opnRCINo = Nothing

End Sub
Private Function VerifyReleasedChkUpdate(ByVal checkno As String) As Boolean
Dim opnVerifyAgain As New ADODB.Recordset

opnVerifyAgain.Open "Select Released from tblCMS_CDPreparedCheck where CheckNo='" & checkno & "' and actioncode=1", opndbaseFMIS, adOpenStatic, adLockOptimistic
If opnVerifyAgain.RecordCount <> 0 Then
    If opnVerifyAgain!Released = True Then
        VerifyReleasedChkUpdate = True
    Else
        VerifyReleasedChkUpdate = False
    End If
End If
opnVerifyAgain.Close
Set opnVerifyAgain = Nothing

End Function
Private Sub LoadBackReportedChks(ByVal chckno As String)
Dim opnURChk As New ADODB.Recordset
Dim cc As Long
Dim tmpval As Variant

opnURChk.Open "Select * from tblCMS_CDRCIReport where checkno ='" & chckno & "'", opndbaseFMIS, adOpenStatic, adLockOptimistic
If opnURChk.RecordCount <> 0 Then
    tmpval = Split(chckno, "-")
    
    Call SetGRidDetails
    
    grd_details.Rows = opnURChk.RecordCount + 2
    Do Until opnURChk.EOF
        DoEvents
        grd_details.TextMatrix(opnURChk.AbsolutePosition, 0) = opnURChk!trnno
        grd_details.TextMatrix(opnURChk.AbsolutePosition, 1) = IIf(Val(opnURChk!YearObligated) = 0, "", opnURChk!YearObligated)
        grd_details.TextMatrix(opnURChk.AbsolutePosition, 2) = IIf(Val(opnURChk!YearObligated) = 0, "", Format(opnURChk!CheckDate, "m/d/yy"))
        grd_details.TextMatrix(opnURChk.AbsolutePosition, 3) = opnURChk!checkno
        grd_details.TextMatrix(opnURChk.AbsolutePosition, 4) = opnURChk!Released
        
        grd_details.TextMatrix(opnURChk.AbsolutePosition, 5) = IIf(IsNull(opnURChk!DVNo), "", opnURChk!DVNo)
        grd_details.TextMatrix(opnURChk.AbsolutePosition, 6) = IIf(IsNull(opnURChk!ResCenterCode), "", opnURChk!ResCenterCode) 'Responsibility center Code
        grd_details.TextMatrix(opnURChk.AbsolutePosition, 7) = opnURChk!ClaimantName
        grd_details.TextMatrix(opnURChk.AbsolutePosition, 8) = FilteredData(opnURChk!NatureofPayment, Chr(13), opnURChk!ActionCode)
        grd_details.TextMatrix(opnURChk.AbsolutePosition, 9) = IIf(opnURChk!CheckAmount = 0, "", Format(opnURChk!CheckAmount, "###,##0.00"))
        grd_details.TextMatrix(opnURChk.AbsolutePosition, 10) = opnURChk!Released   'This two(2) were used in Loading back Color
        grd_details.TextMatrix(opnURChk.AbsolutePosition, 11) = opnURChk!ActionCode 'when there is Insert/Delete Activity from the cell
        grd_details.TextMatrix(opnURChk.AbsolutePosition, 12) = opnURChk!OrderNo 'SortOrder is used in the report for purposes of sorting
        grd_details.TextMatrix(opnURChk.AbsolutePosition, 13) = IIf(IsNull(opnURChk!TransmittalReportNo), "", opnURChk!TransmittalReportNo)
        
        opnURChk.MoveNext
    Loop
    grd_details.Row = grd_details.Rows - 1 'for setting the focus to the last row of the grid
Else
    Call SetGRidDetails
    MsgBox "Check Number Not Found, Please Check The Check No. and Try Again", vbInformation, "System Message"
End If

opnURChk.Close
Set opnURChk = Nothing


End Sub

Private Function GetReleasedFlagEquivalent(ByVal CellBackColor As String) As Integer
Select Case CellBackColor
    Case "16711680" 'Blue
        GetReleasedFlagEquivalent = 2
    Case "16777215" 'White
        GetReleasedFlagEquivalent = 0
    Case "65280" 'Green
        GetReleasedFlagEquivalent = 1
    Case "255" 'Red
        GetReleasedFlagEquivalent = 4
End Select
End Function
Private Function GetPTVNumber(ByVal checkno As String) As String
Dim opnPTV As New ADODB.Recordset
Dim tmpPTV As Variant

On Error Resume Next

opnPTV.Open "Select DVNo from  tblCMS_CDCheckBook where chknumber='" & checkno & "' and actioncode=1", opndbaseFMIS, adOpenStatic, adLockOptimistic
If opnPTV.RecordCount <> 0 Then
    tmpPTV = Split(opnPTV!DVNo, "-")
    'Debug.Print "Select DVNo from  tblCMS_CDCheckBook where chknumber='" & CheckNo & "' and actioncode=1"
    GetPTVNumber = tmpPTV(0) & "-" & tmpPTV(1) & "-" & tmpPTV(2) & "-" & tmpPTV(3)
Else
    GetPTVNumber = ""
End If
opnPTV.Close
Set opnPTV = Nothing


End Function
















Private Sub Combo1_Click()
Call LoadFund(cmb_FundType)
End Sub

Private Sub Combo2_click()
Select Case (Combo2.Text)
    Case "Checkno":
    End Select
    
End Sub

Private Sub Command1_Click()
'Label13.Caption = "JEV Numbering..."
'Label13.Refresh

Call JEVMassNumbering
'Label13.Caption = ""


End Sub
Private Sub JEVMassNumbering()
Dim opnJEV As New ADODB.Recordset
Dim rec As New ADODB.Recordset
Dim SQL As String
Dim cc As Integer
Dim DVNo As String
Dim LastJEVSNno As Long
    

'LastJEVSNno = GetLatestSNoForJEV(ConvertFullFundtoMedium(FundType), DTPicker1.Year)

For cc = 1 To grd_details.Rows - 2

    DVNo = GetDVNobyChkNo(grd_details.TextMatrix(cc, 3))
    

    If grd_details.TextMatrix(cc, 4) = 2 Then
    
        SQL = "SELECT tblAMIS_IncomingDVTrns.FundType as FundType, tblAMIS_JournalEntry.TransType as TransType, tblAMIS_JournalEntry.DVNo as DVNo, " & _
                "          tblAMIS_JournalEntry.TransDate as TransDate, tblAMIS_JournalEntry.JEVSeriesNo as JEVSeriesNo,(Select FundCode from tblRefBMS_Funds where FundMedium=tblAMIS_IncomingDVTrns.FundType) as FundCode " & _
                " FROM tblAMIS_IncomingDVTrns INNER JOIN " & _
                "          tblAMIS_JournalEntry ON tblAMIS_IncomingDVTrns.DVNo = tblAMIS_JournalEntry.DVNo " & _
                " Where (tblAMIS_JournalEntry.ActionCode = 1) And (tblAMIS_IncomingDVTrns.ActionCode = 1) " & _
                " GROUP BY tblAMIS_IncomingDVTrns.FundType, tblAMIS_JournalEntry.TransType, tblAMIS_JournalEntry.DVNo, " & _
                "          tblAMIS_JournalEntry.TransDate , tblAMIS_JournalEntry.JEVSeriesNo " & _
                " HAVING   tblAMIS_JournalEntry.DVNo ='" & DVNo & "'"
    
        opnJEV.Open SQL, opndbaseFMIS, adOpenStatic, adLockOptimistic
        If opnJEV.RecordCount <> 0 Then
            grd_details.TextMatrix(cc, 14) = opnJEV!FundCode & "-" & Right(Year(Date), 2) & "-" & Format(Month(Date), "00") & "-" & Format(opnJEV!TransType, "00") & "-" & LastJEVSNno
            grd_details.TextMatrix(cc, 14) = opnJEV!FundCode & "-" & Right(grd_details.TextMatrix(cc, 1), 2) & "-" & Left(grd_details.TextMatrix(cc, 1), 2) & "-" & Format(opnJEV!TransType, "00") & "-" & LastJEVSNno
            LastJEVSNno = LastJEVSNno + 1
            
            grd_details.TextMatrix(cc, 15) = "Ready"
        Else 'No REcord Found yet in the AMIS
            grd_details.TextMatrix(cc, 14) = "000-00-00-00-xxxxx"
            grd_details.TextMatrix(cc, 15) = "Not Ready"
        End If
        opnJEV.Close
        Set opnJEV = Nothing

    Else
        grd_details.TextMatrix(cc, 14) = "Not yet Issued Check"
        grd_details.TextMatrix(cc, 15) = "Not Ready"
        
    End If
Next cc

rec.Open "select * from vw_CDRCIReport where checkno = '" & txt_Search.Text & "' and AlreadySaved2JEV = 1 ", opndbaseFMIS, adOpenStatic, adLockOptimistic
    If rec.RecordCount <> 0 Then
    MsgBox "Already Have Jev number Cannot edit", vbInformation, "System Message"
    Call SetGRidDetails
    End If
rec.Close

End Sub
Private Sub Command2_Click()
Dim cc, tmp As Integer
If Len(Trim(grd_details.TextMatrix(cc, 14))) = "000-00-00-00-xxxxx" Or Len(Trim(grd_details.TextMatrix(cc, 14))) = "" Then
    MsgBox "Complete The Data First..", vbInformation, "System Message"
Exit Sub
Else
    If MsgBox("Save JEV Nos.?", vbQuestion + vbYesNo, "System Confirmation") = vbYes Then
    For cc = 1 To grd_details.Rows - 2
        
            If Len(Trim(grd_details.TextMatrix(cc, 14))) > 0 Then
                If IsFormatCorrect(grd_details.TextMatrix(cc, 14)) = True Then
                    
                    'Updating table from PTO....
                    opndbaseFMIS.Execute "Update tblCMS_CDRCIReport set AlreadySaved2JEV=1,DatePostedtoJEV='" & Date & "',PostedtoJEVUserid='" & ActiveUserID & "' where trnno=" & grd_details.TextMatrix(cc, 0) & ""
                    
                    'Updating Accounting REcord...
                    tmp = ExtractJEVSNo(grd_details.TextMatrix(cc, 14))
                    
                   ' opndbaseFMIS.Execute "update tblAMIS_JournalEntry set JEVNo='" & grd_details.TextMatrix(cc, 14) & "', " & _
                   ' " JEVSeriesNo=" & tmp & ",JEVBy='" & ActiveUserID & "', " & _
                   ' " JEVDate='" & Date & "' where DVNo='" & grd_details.TextMatrix(cc, 5) & "'"
                
                opndbaseFMIS.Execute "update tblAMIS_JournalEntry set JEVNo='" & grd_details.TextMatrix(cc, 14) & "', " & _
                    " JEVSeriesNo=" & tmp & ",JEVBy='" & ActiveUserID & "', " & _
                    " JEVDate='" & Date & "' where DVNo='" & GetDVNobyChkNo(grd_details.TextMatrix(cc, 3)) & "'"
                
                
                
                'GetDVNobyChkNo(grd_details.TextMatrix(grd_details.Row, 3))
                
                
                
                
                End If
            End If
       
    Next cc
    MsgBox "Posting to JEV, Successful!", vbInformation, "System Information"
    Command3_Click 'Loading Back Active RCI Numbers...
    'List1.ListIndex = GetIndex4ListBox(List1, tmpRCINo)
    End If
End If
End Sub


Private Sub Command3_Click()
'Label13.Caption = "Loading, Please wait..."
'Label13.Refresh
'Call LoadSavedReportNos(DTPicker1.Month, DTPicker1.Year, cmb_FundType.Text)
'Label13.Caption = ""

End Sub

Private Sub DTPicker1_Change()


DTPicker1.Value = DTPicker1.Month & "/1/" & DTPicker1.Year
'Call SetGRidDetails

End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then
    Unload Me
End If
End Sub

Private Sub Form_Load()
Me.Top = (Screen.Height - Me.Height) / 2 - 550
Me.Left = (Screen.Width - Me.Width) / 2
WindowsXPC1.InitSubClassing
'DTPicker1.Value = Month(Date) & "/1/" & Year(Date)
'Call LoadFundType(cmb_FundType)
'Label13.Caption = ""
Label14.Caption = ""
Timer1.Enabled = True
End Sub
Public Sub LoadFund(ByVal cmb As ComboBox)
Dim opnfund As New ADODB.Recordset
Dim cc As Integer

Select Case (Combo1)
                Case "General Fund":
                 opnfund.Open "Select fundname,fundcode from  tblRefBMS_Funds where fundname not in('Trust Fund','Special Education Fund') order by fundname", opndbaseFMIS, adOpenStatic, adLockOptimistic
                 Case "Trust Fund":
                 opnfund.Open "Select fundname,fundcode from tblRefBMS_Funds where fundname='Trust Fund' order by fundname", opndbaseFMIS, adOpenStatic, adLockOptimistic
                 Case "Special Education Fund":
                 opnfund.Open "Select fundname,fundcode from tblRefBMS_Funds where fundname='Special Education Fund' order by fundname", opndbaseFMIS, adOpenStatic, adLockOptimistic
                End Select
If opnfund.RecordCount <> 0 Then
    cmb.Clear
    Do Until opnfund.EOF
        cmb.AddItem (opnfund!FundName)
        cmb.ItemData(cc) = opnfund!FundCode
        cc = cc + 1
        opnfund.MoveNext
    Loop
Else
    cmb.Clear
End If
opnfund.Close
Set opnfund = Nothing
End Sub


Private Sub SetGRidDetails()
Dim cc As Integer

grd_details.Clear

grd_details.Cols = 16
grd_details.Rows = 3

grd_details.TextMatrix(0, 0) = "trnno"
grd_details.TextMatrix(0, 1) = "Year"
grd_details.TextMatrix(0, 2) = "Chk Date"
grd_details.TextMatrix(0, 3) = "Check No"
grd_details.TextMatrix(0, 4) = "" 'Status Color Code
grd_details.TextMatrix(0, 5) = "PTV Number"
grd_details.TextMatrix(0, 6) = "R.C."
grd_details.TextMatrix(0, 7) = "Payee"
grd_details.TextMatrix(0, 8) = "Nature of Payment"
grd_details.TextMatrix(0, 9) = "Amount"
grd_details.TextMatrix(0, 10) = "Status"
grd_details.TextMatrix(0, 11) = "ActionCode"
grd_details.TextMatrix(0, 12) = "SortOrder"
grd_details.TextMatrix(0, 13) = "TransmittalNo" 'Actually this has nothing to do with the process...........
grd_details.TextMatrix(0, 14) = "JEV No"
grd_details.TextMatrix(0, 15) = "Status"


grd_details.ColWidth(0) = 0
grd_details.ColWidth(1) = 0
grd_details.ColWidth(2) = 800
grd_details.ColWidth(3) = 1200
grd_details.ColWidth(4) = 0
grd_details.ColWidth(5) = 1300
grd_details.ColWidth(6) = 700
grd_details.ColWidth(7) = 3500
grd_details.ColWidth(8) = 3600
grd_details.ColWidth(9) = 1190
grd_details.ColWidth(10) = 0
grd_details.ColWidth(11) = 0
grd_details.ColWidth(12) = 0
grd_details.ColWidth(13) = 0
grd_details.ColWidth(14) = 1700
grd_details.ColWidth(15) = 1300


For cc = 0 To grd_details.Cols - 1
    grd_details.Row = 0
    grd_details.Col = cc
    grd_details.CellAlignment = 4
Next cc
End Sub
Private Function LoadOtherDetailsOfUnreportedCheck(ByVal compositioncode As Long) As Variant
Dim opnOther As New ADODB.Recordset

opnOther.Open "Select * from vw_DepositoryBank where FMISAccountCode=" & compositioncode & " and active=1", opndbaseFMIS, adOpenStatic, adLockOptimistic
If opnOther.RecordCount <> 0 Then
    LoadOtherDetailsOfUnreportedCheck = opnOther!fundmedium & "," & opnOther!Accountname & "," & opnOther!BankAccountNo
End If
opnOther.Close
Set opnOther = Nothing
End Function


Private Sub Form_Unload(Cancel As Integer)
'Set frmCDPrntCheckIssued = Nothing
WindowsXPC1.EndWinXPCSubClassing
End Sub

Private Sub Load_Offices()
Dim OREc As New ADODB.Recordset
Dim X As Integer

With frmJEVNumberingAssignment
    .cmbOffice.ComboItems.Clear
    OREc.Open ("Select * FRom tblREF_AIS_Offices Order By [OfficeMedium]"), opndbaseFMIS, adOpenStatic, adLockOptimistic
    If OREc.RecordCount > 0 Then
        For X = 1 To OREc.RecordCount
            .cmbOffice.ComboItems.Add .cmbOffice.ComboItems.Count + 1, OREc!FMISOfficeID & "ID", OREc![OfficeMedium], 1
        OREc.MoveNext
        Next X
    .cmbOffice.ComboItems(1).Selected = True
    End If
    OREc.Close
    Set OREc = Nothing
End With
End Sub

Private Sub Load_FundTypes()
Dim OREc As New ADODB.Recordset
Dim X As Integer

With frmJEVNumberingAssignment
    .cmbfundtype.ComboItems.Clear
    OREc.Open ("Select * from tblRefBMS_Funds"), opndbaseFMIS, adOpenStatic, adLockOptimistic
    If OREc.RecordCount > 0 Then
        For X = 1 To OREc.RecordCount
            .cmbfundtype.ComboItems.Add .cmbfundtype.ComboItems.Count + 1, OREc!FundCode & "ID", OREc!FundName, 2
        OREc.MoveNext
        Next X
    .cmbfundtype.ComboItems(1).Selected = True
    End If
    OREc.Close
    Set OREc = Nothing
End With


End Sub
Public Sub getRCIno(ByVal chck As String)
Dim rec As New ADODB.Recordset
rec.Open "select Rcino,compositioncode from tblCMS_CDRCIReport where checkno = '" & chck & "'", opndbaseFMIS, adOpenStatic, adLockBatchOptimistic
If rec.RecordCount <> 0 Then
txt_RCI.Text = rec.Fields!RCINo
tmpcomposition = rec.Fields!compositioncode
End If
rec.Close
Set rec = Nothing
End Sub

'---------------RICHARD-----------------------------
Private Sub Set_4_Old_Trans()
On Error GoTo err
   
    With frmJEVNumberingAssignment
        .cmdsave.Visible = True
        .txt_RCenter.Visible = False
        .txt_FundType.Visible = False
        .cmbOffice.Visible = True
        .cmbOffice.Width = .txt_RCenter.Width
        .cmbOffice.Left = .txt_RCenter.Left
        .cmbOffice.Top = .txt_RCenter.Top
        .cmbfundtype.Visible = True
        .cmbfundtype.Width = .txt_FundType.Width
        .cmbfundtype.Left = .txt_FundType.Left
        .cmbfundtype.Top = .txt_FundType.Top
    End With
Exit Sub
err:
    MsgBox "Error: " & err.Description
End Sub

Private Function fundmedium(fnd As String) As String
Dim opnfund As New ADODB.Recordset

    opnfund.Open "Select fundmedium from tblRefBMS_Funds where fundname='" & Trim(fnd) & "'", opndbaseFMIS, adOpenStatic, adLockOptimistic
    If Not opnfund.EOF Then
        fundmedium = Trim(opnfund(0))
    End If
End Function

'----------------------------------------------------


Private Sub grd_details_DblClick()
On Error GoTo bad
'MsgBox GetDVNobyChkNo(grd_details.TextMatrix(grd_details.Row, 3)), vbInformation, "DV Number"

If grd_details.TextMatrix(grd_details.Row, 4) = 2 Then
    ActiveFormCaller = Me.Name
    ForTheGridRowNo = grd_details.Row
    'Load_Offices   '---loads offices to the OFFICE combo - RICHARD
    'Load_FundTypes '---loads functypes to the FUNDTYPE combo - RICHARD
    If Len(grd_details.TextMatrix(grd_details.Row, 14)) <> 0 Then 'Kung Naa nay JEV No
        frmJEVNumberingAssignment.txt_JEVNo.Text = grd_details.TextMatrix(grd_details.Row, 14)
        frmJEVNumberingAssignment.txt_DVNo.Text = GetDVNobyChkNo(grd_details.TextMatrix(grd_details.Row, 3))
        
       
    Else
'        With frmJEVPreparation
'            .Toolbar1_ButtonClick Toolbar1.Buttons(3)
'            .txtDVNo = GetDVNobyChkNo(grd_details.TextMatrix(grd_details.Row, 3))
'            .Show
'        End With
        'frmJEVNumberingAssignment.txt_DVNo.Text = GetDVNobyChkNo(grd_details.TextMatrix(grd_details.Row, 3))
        
    End If
     
    If Trim(frmJEVNumberingAssignment.txt_RCenter) = "" And Trim(frmJEVNumberingAssignment.txt_FundType) = "" Then
        'Set_4_Old_Trans '---set components for old transacted vouchers
        
        '---passing the transaction to JEV Prepartion Form if Old Transaction---'
        With frmJEVPreparation
            .Toolbar1_ButtonClick .Toolbar1.Buttons.Item(1)
            .txtDVNo = GetDVNobyChkNo(grd_details.TextMatrix(grd_details.Row, 3))
            .txtfund = fundmedium(Trim(txt_fund))
            .txtDate = Format(grd_details.TextMatrix(grd_details.Row, 2), "mmmm dd, yyyy")
            .txtDate.Locked = False
            .LoadAccountsByFund Trim(.txtfund)
            .optCheck.Value = True
            .txtClaimant = Trim(grd_details.TextMatrix(grd_details.Row, 7))
            .txtAmount = Trim(grd_details.TextMatrix(grd_details.Row, 9))
            .txtParticular = Trim(grd_details.TextMatrix(grd_details.Row, 8))
                If Trim(grd_details.TextMatrix(grd_details.Row, 6)) <> "" Then
                    .txtRC = GetOfficeName(Val(Trim(grd_details.TextMatrix(grd_details.Row, 6))), "OfficeMedium")
                End If
            'GetOfficeName(DRec!RCenter, "OfficeMedium")
            frmJEVPreparation.LoadOffice
            frmJEVPreparation.cmbRC.Locked = False
            frmJEVPreparation.cmbRC.Visible = False
            .Show 1
        End With
        
    Else
        frmJEVNumberingAssignment.Show vbModal
    End If
Else
    MsgBox "There is no Voucher Attachment for this Check!" & Chr(13) & Chr(13) & "Please Select a New..", vbInformation, "System Information"
End If
Exit Sub
bad:
MsgBox err.Description
End Sub

Private Sub List1_Click()
Label13.Caption = "Loading Details..."
Label13.Refresh
tmpRCINo = List1.Text



Call LoadBackReportedChks(List1.List(List1.ListIndex), DTPicker1.Month, DTPicker1.Year)
Call LoadBackOthers(List1.ItemData(List1.ListIndex))
Label13.Caption = ""
Label14.Caption = (grd_details.Rows - 2) & " Voucher/s Found..."
End Sub
Private Sub LoadBackOthers(ByVal CompositionID As Long)
Dim opnOtherDetails As New ADODB.Recordset
Dim xx As Integer

opnOtherDetails.Open "Select * from vw_DepositoryBank where FMISAccountCode=" & CompositionID & " and active=1", opndbaseFMIS, adOpenKeyset, adLockOptimistic
If opnOtherDetails.RecordCount <> 0 Then
    txt_fund.Text = opnOtherDetails!FundType
    txt_bank.Text = opnOtherDetails!BankID
    txt_AcctNo.Text = opnOtherDetails!BankAccountNo
    txt_acctName.Text = opnOtherDetails!Accountname
End If
opnOtherDetails.Close
Set opnOtherDetails = Nothing
End Sub


Private Sub Text1_Change()

End Sub

Private Sub Timer1_Timer()
    Call SetGRidDetails

Timer1.Enabled = False
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

On Error GoTo handler

Select Case Button.Index
    Case 1 'Print----------------------------
    
    Case 3 'finding Check No ----------------
        If List1.ListCount <> 0 Then
askchekno:
            tmpCheckNo = InputBox("Enter Check No.:", "Search Check No.")
            
            If Len(Trim(tmpCheckNo)) <> 0 Then
                Call FindCheckNo(tmpCheckNo)
            Else
                Exit Sub
            End If
        
        End If


    Case 5 'Close----------------------------
        Unload Me
End Select

handler:
If err.Number <> 0 Then
    MsgBox err.Description
    Exit Sub
End If
End Sub
Private Sub FindCheckNo(ByVal checkno As String)
Dim cc, ZZ As Integer

For cc = 0 To List1.ListCount - 1
DoEvents
    List1.ListIndex = cc
    For ZZ = 1 To grd_details.Rows - 1
    DoEvents
        grd_details.Col = 3
        grd_details.TopRow = ZZ
        If checkno = grd_details.TextMatrix(ZZ, 3) Then
            grd_details.Row = ZZ
            Exit Sub
        End If
    Next ZZ
Next cc
MsgBox "Check No.: " & checkno & " is not found!", vbInformation, "System Information"
End Sub


Private Sub txt_Search_KeyDown(KeyCode As Integer, Shift As Integer)
Dim tmpval As Long

On Error GoTo handler
    If KeyCode = 13 Then
    
    Call LoadBackReportedChks(txt_Search.Text)
    Call getRCIno(txt_Search.Text)
    Call LoadBackOthers(tmpcomposition)
    Call JEVMassNumbering
    End If
    Exit Sub
handler:
If err.Number <> 0 Then
    MsgBox err.Description
End If
End Sub

VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{1693405E-2DC9-4248-B52F-4AC9145DA2AF}#1.0#0"; "WINXPC~1.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmVw_CheckIssued 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Check Issued"
   ClientHeight    =   10080
   ClientLeft      =   150
   ClientTop       =   1380
   ClientWidth     =   14370
   Icon            =   "frmJEVNumberingForCashReceiptsI.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   10080
   ScaleWidth      =   14370
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame5 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Fund Type"
      ForeColor       =   &H00FFFFFF&
      Height          =   660
      Left            =   195
      TabIndex        =   20
      Top             =   720
      Width           =   3135
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
         ItemData        =   "frmJEVNumberingForCashReceiptsI.frx":0E42
         Left            =   75
         List            =   "frmJEVNumberingForCashReceiptsI.frx":0E4F
         TabIndex        =   21
         Top             =   240
         Width           =   2910
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Load Reports"
      Height          =   480
      Left            =   2070
      TabIndex        =   13
      Top             =   3135
      Width           =   1245
   End
   Begin VB.TextBox txt_Search 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   420
      Left            =   165
      TabIndex        =   11
      Top             =   3690
      Width           =   3180
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Special Accounts:"
      ForeColor       =   &H00FFFFFF&
      Height          =   780
      Left            =   195
      TabIndex        =   9
      Top             =   1440
      Width           =   3135
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
         Left            =   75
         TabIndex        =   10
         Top             =   300
         Width           =   2910
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Source Details"
      Height          =   3330
      Left            =   8760
      TabIndex        =   4
      Top             =   600
      Width           =   3150
      Begin VB.TextBox txt_acctName 
         Height          =   360
         Left            =   345
         TabIndex        =   17
         Top             =   2790
         Width           =   2385
      End
      Begin VB.TextBox txt_AcctNo 
         Height          =   360
         Left            =   345
         TabIndex        =   16
         Top             =   2100
         Width           =   2385
      End
      Begin VB.TextBox txt_bank 
         Height          =   360
         Left            =   345
         TabIndex        =   15
         Top             =   1290
         Width           =   2385
      End
      Begin VB.TextBox txt_fund 
         Height          =   360
         Left            =   345
         TabIndex        =   14
         Top             =   555
         Width           =   2385
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Account Number"
         Height          =   195
         Left            =   195
         TabIndex        =   8
         Top             =   1875
         Width           =   1200
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Account Name"
         Height          =   195
         Left            =   195
         TabIndex        =   7
         Top             =   2565
         Width           =   1065
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Drawee Bank"
         Height          =   195
         Left            =   210
         TabIndex        =   6
         Top             =   1050
         Width           =   975
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fund Type"
         Height          =   195
         Left            =   210
         TabIndex        =   5
         Top             =   330
         Width           =   765
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   1545
      Top             =   4845
   End
   Begin VB.ListBox List1 
      Columns         =   2
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3420
      ItemData        =   "frmJEVNumberingForCashReceiptsI.frx":0E85
      Left            =   3480
      List            =   "frmJEVNumberingForCashReceiptsI.frx":0E87
      TabIndex        =   3
      Top             =   840
      Width           =   4980
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "For the Period"
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   165
      TabIndex        =   1
      Top             =   2265
      Width           =   3180
      Begin MSComCtl2.DTPicker DTPicker1 
         CausesValidation=   0   'False
         Height          =   360
         Left            =   165
         TabIndex        =   2
         Top             =   285
         Width           =   2835
         _ExtentX        =   5001
         _ExtentY        =   635
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "MMMM yyyy"
         Format          =   271581187
         UpDown          =   -1  'True
         CurrentDate     =   38240
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7320
      Top             =   4080
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
            Picture         =   "frmJEVNumberingForCashReceiptsI.frx":0E89
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJEVNumberingForCashReceiptsI.frx":1F0B
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJEVNumberingForCashReceiptsI.frx":4045
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJEVNumberingForCashReceiptsI.frx":40A3
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJEVNumberingForCashReceiptsI.frx":43BD
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin WinXPC_Engine.WindowsXPC WindowsXPC1 
      Left            =   240
      Top             =   9600
      _ExtentX        =   6588
      _ExtentY        =   1085
      ColorScheme     =   2
      Common_Dialog   =   0   'False
      PictureControl  =   0   'False
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   600
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   14370
      _ExtentX        =   25347
      _ExtentY        =   1058
      ButtonWidth     =   2249
      ButtonHeight    =   1005
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Print &Report"
            Object.ToolTipText     =   "Print RCI Report"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "slash"
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Find Check No"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Close"
            ImageIndex      =   4
         EndProperty
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid grd_details 
      Height          =   4920
      Left            =   120
      TabIndex        =   22
      Top             =   4560
      Width           =   14175
      _ExtentX        =   25003
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
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label14"
      Height          =   195
      Left            =   13725
      TabIndex        =   19
      Top             =   9705
      Width           =   570
   End
   Begin VB.Label Label13 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Label13"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   180
      TabIndex        =   18
      Top             =   3075
      Width           =   1800
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Search (RCI No.)"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   165
      TabIndex        =   12
      Top             =   3390
      Width           =   1215
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   3795
      Left            =   -105
      Top             =   645
      Width           =   8730
   End
End
Attribute VB_Name = "frmVw_CheckIssued"
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
Dim tmpVal As Variant

If InStr(Trim(AlobsNo), ",") <> 0 Then 'Multiple Alobs----
    tmpVal = Split(AlobsNo, ",")    '----Spliting from multiple
    tmpVal = tmpVal(0)              '----To Single
    
    tmpVal = Split(tmpVal, "-")     '----Split into Codes----
    GetFinalResCode = tmpVal(1) 'Final Responsibility Code---
Else 'Single Alobs----------------------------------
    tmpVal = Split(AlobsNo, "-")     '----Split into Codes----
    GetFinalResCode = tmpVal(1) 'Final Responsibility Code---
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
Dim sql As String
Dim cc As Integer

sql = "SELECT  tblCMS_CDRCIReport.RCINo as RCINo, tblCMS_CDRCIReport.Compositioncode as Compositioncode, vw_DepositoryBank.FundType,tblCMS_CDRCIReport.AlreadySaved2JEV " & _
        " FROM tblCMS_CDRCIReport LEFT OUTER JOIN " & _
        " vw_DepositoryBank ON tblCMS_CDRCIReport.Compositioncode = vw_DepositoryBank.FMISAccountCode " & _
        " Where (tblCMS_CDRCIReport.ActionCode = 1) " & _
        " GROUP BY tblCMS_CDRCIReport.RCINo, tblCMS_CDRCIReport.Compositioncode, vw_DepositoryBank.FundType, " & _
        " tblCMS_CDRCIReport.AlreadySaved2JEV , Year(tblCMS_CDRCIReport.CheckDate), Month(tblCMS_CDRCIReport.CheckDate) " & _
        " HAVING (vw_DepositoryBank.FundType = '" & fund & "') AND (YEAR(tblCMS_CDRCIReport.CheckDate) = " & TrnYear & ") AND " & _
        " (MONTH(tblCMS_CDRCIReport.CheckDate) = " & trnMonth & ") AND (tblCMS_CDRCIReport.AlreadySaved2JEV = 0)" & _
        " ORDER BY tblCMS_CDRCIReport.RCINo "

'opnRCINo.Open "Select RCINo from tblCMS_CDRCIReport where Year(CheckDate)=" & TrnYear & " and Month(CheckDate)=" & trnMonth & " and Compositioncode=" & compositioncode & " group by RCINO order by RCINo ", opndbaseFMIS, adOpenStatic, adLockOptimistic
Debug.Print sql

opnRCINo.Open sql, opndbaseFMIS, adOpenStatic, adLockOptimistic



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
    If opnVerifyAgain!released = True Then
        VerifyReleasedChkUpdate = True
    Else
        VerifyReleasedChkUpdate = False
    End If
End If
opnVerifyAgain.Close
Set opnVerifyAgain = Nothing

End Function
Private Sub LoadBackReportedChks(ByVal RCINo As String, ByVal trnMonth As Integer, ByVal TrnYear As Integer)
Dim opnURChk As New ADODB.Recordset
Dim cc As Long
Dim tmpVal As Variant

opnURChk.Open "Select * from vw_MP_CheckIssuedview where RCINo ='" & RCINo & "' and year(CheckDate)=" & TrnYear & " and month(CheckDate)=" & trnMonth & "  and AlreadySaved2JEV=0 order by OrderNo", opndbaseFMIS, adOpenStatic, adLockOptimistic
If opnURChk.RecordCount <> 0 Then
    tmpVal = Split(RCINo, "-")
    
    Call SetGRidDetails
    
    grd_details.Rows = opnURChk.RecordCount + 2
    Do Until opnURChk.EOF
        DoEvents
        grd_details.TextMatrix(opnURChk.AbsolutePosition, 0) = opnURChk!Trnno
        grd_details.TextMatrix(opnURChk.AbsolutePosition, 1) = IIf(Val(opnURChk!YearObligated) = 0, "", opnURChk!YearObligated)
        grd_details.TextMatrix(opnURChk.AbsolutePosition, 2) = IIf(Val(opnURChk!YearObligated) = 0, "", Format(opnURChk!CheckDate, "m/d/yy"))
        grd_details.TextMatrix(opnURChk.AbsolutePosition, 3) = opnURChk!checkno
        grd_details.TextMatrix(opnURChk.AbsolutePosition, 4) = opnURChk!released
        
        grd_details.TextMatrix(opnURChk.AbsolutePosition, 5) = IIf(IsNull(opnURChk!dvno), "", opnURChk!dvno)
        grd_details.TextMatrix(opnURChk.AbsolutePosition, 6) = IIf(IsNull(opnURChk!ResCenterCode), "", opnURChk!ResCenterCode) 'Responsibility center Code
        grd_details.TextMatrix(opnURChk.AbsolutePosition, 7) = opnURChk!claimantname
        grd_details.TextMatrix(opnURChk.AbsolutePosition, 8) = FilteredData(opnURChk!NatureofPayment, Chr(13), opnURChk!ActionCode)
        grd_details.TextMatrix(opnURChk.AbsolutePosition, 9) = IIf(opnURChk!CheckAmount = 0, "", Format(opnURChk!CheckAmount, "###,##0.00"))
        grd_details.TextMatrix(opnURChk.AbsolutePosition, 10) = opnURChk!released   'This two(2) were used in Loading back Color
        grd_details.TextMatrix(opnURChk.AbsolutePosition, 11) = opnURChk!ActionCode 'when there is Insert/Delete Activity from the cell
        grd_details.TextMatrix(opnURChk.AbsolutePosition, 12) = opnURChk!OrderNo 'SortOrder is used in the report for purposes of sorting
        grd_details.TextMatrix(opnURChk.AbsolutePosition, 13) = IIf(IsNull(opnURChk!TransmittalReportNo), "", opnURChk!TransmittalReportNo)
        grd_details.TextMatrix(opnURChk.AbsolutePosition, 14) = isIssued(opnURChk!actiontype)
        If grd_details.TextMatrix(opnURChk.AbsolutePosition, 14) = "Check Issued" Then
        grd_details.TextMatrix(opnURChk.AbsolutePosition, 15) = opnURChk!datetimeentered
        Else
        grd_details.TextMatrix(opnURChk.AbsolutePosition, 15) = ""
        End If
        opnURChk.MoveNext
    Loop
    grd_details.Row = grd_details.Rows - 1 'for setting the focus to the last row of the grid
Else
    Call SetGRidDetails
End If

opnURChk.Close
Set opnURChk = Nothing


End Sub
Public Function isIssued(ByVal act As Integer) As String
If act = 6 Then
    isIssued = "Check Issued"
Else
    isIssued = "Not Yet Issued Check"
End If
End Function
'Public Function dateIssued(ByVal act As Integer) As String
'If act = 6 Then
'    isIssued = "Check Issued"
'Else
'    isIssued = "Not Yet Issued Check"
'End If
'End Function

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
    tmpPTV = Split(opnPTV!dvno, "-")
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


Private Sub JEVMassNumbering(ByVal FundType As String)
Dim opnJEV As New ADODB.Recordset
Dim sql As String
Dim cc As Integer
Dim dvno As String
Dim LastJEVSNno As Long
Dim last As String
    

LastJEVSNno = GetLatestSNoForJEV(ConvertFullFundtoMedium(FundType), DTPicker1.Year, DTPicker1.Month)

For cc = 1 To grd_details.Rows - 2

    dvno = GetDVNobyChkNo(grd_details.TextMatrix(cc, 3))
    

    If grd_details.TextMatrix(cc, 4) = 2 Then
    
        sql = "SELECT tblAMIS_IncomingDVTrns.FundType as FundType, tblAMIS_JournalEntry.TransType as TransType, tblAMIS_JournalEntry.DVNo as DVNo, " & _
                "          tblAMIS_JournalEntry.TransDate as TransDate, tblAMIS_JournalEntry.JEVSeriesNo as JEVSeriesNo,(Select FundCode from tblRefBMS_Funds where FundMedium=tblAMIS_IncomingDVTrns.FundType) as FundCode " & _
                " FROM tblAMIS_IncomingDVTrns INNER JOIN " & _
                "          tblAMIS_JournalEntry ON tblAMIS_IncomingDVTrns.DVNo = tblAMIS_JournalEntry.DVNo " & _
                " Where (tblAMIS_JournalEntry.ActionCode = 1) And (tblAMIS_IncomingDVTrns.ActionCode = 1) " & _
                " GROUP BY tblAMIS_IncomingDVTrns.FundType, tblAMIS_JournalEntry.TransType, tblAMIS_JournalEntry.DVNo, " & _
                "          tblAMIS_JournalEntry.TransDate , tblAMIS_JournalEntry.JEVSeriesNo " & _
                " HAVING   tblAMIS_JournalEntry.DVNo ='" & dvno & "'"
    
        
        
        
        
        
        
        
        
        
        opnJEV.Open sql, opndbaseFMIS, adOpenStatic, adLockOptimistic
        If opnJEV.RecordCount <> 0 Then
            'grd_details.TextMatrix(cc, 14) = opnJEV!FundCode & "-" & Right(Year(Date), 2) & "-" & Format(Month(Date), "00") & "-" & Format(opnJEV!TransType, "00") & "-" & LastJEVSNno
            
            grd_details.TextMatrix(cc, 14) = opnJEV!fundcode & "-" & Right(DTPicker1.Year, 2) & "-" & Format((DTPicker1.Month), "00") & "-" & Format(opnJEV!Transtype, "00") & "-" & Format(LastJEVSNno, "0000")
            LastJEVSNno = LastJEVSNno + 1
        Else 'No REcord Found yet in the AMIS
            grd_details.TextMatrix(cc, 14) = "000-00-00-00-xxxxx"
        End If
        opnJEV.Close
        Set opnJEV = Nothing

    Else
        grd_details.TextMatrix(cc, 14) = "Not yet Issued Check"
        
    End If
Next cc



End Sub

Private Sub Command3_Click()
Label13.Caption = "Loading, Please wait..."
Label13.Refresh
Call LoadSavedReportNos(DTPicker1.Month, DTPicker1.Year, cmb_FundType.Text)
Label13.Caption = ""

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
DTPicker1.Value = Month(Date) & "/1/" & Year(Date)
'Call LoadFundType(cmb_FundType)
Label13.Caption = ""
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
        cmb.ItemData(cc) = opnfund!fundcode
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
grd_details.TextMatrix(0, 14) = "Status"
grd_details.TextMatrix(0, 15) = "Date Issued"



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
grd_details.ColWidth(15) = 1800


For cc = 0 To grd_details.Cols - 1
    grd_details.Row = 0
    grd_details.col = cc
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


Private Sub Form_Resize()
On Error Resume Next
  'This will resize the grid when the form is resized
  Me.grd_details.Width = Me.ScaleWidth
  Me.grd_details.Height = Me.ScaleHeight - grd_details.Top ''- 30 - picButtons.Height - picStatBox.Height
 
End Sub

Private Sub Form_Unload(Cancel As Integer)
'Set frmCDPrntCheckIssued = Nothing
WindowsXPC1.EndWinXPCSubClassing
End Sub

'Private Sub Load_Offices()
'Dim OREc As New ADODB.Recordset
'Dim x As Integer
'
'With frmJEVNumberingAssignment_New
'    .cmbOffice.ComboItems.Clear
'    OREc.Open ("Select * FRom tblREF_AIS_Offices Order By [OfficeMedium]"), opndbaseFMIS, adOpenStatic, adLockOptimistic
'    If OREc.RecordCount > 0 Then
'        For x = 1 To OREc.RecordCount
'            .cmbOffice.ComboItems.Add .cmbOffice.ComboItems.Count + 1, OREc!FMISOfficeID & "ID", OREc![officemeDium], 1
'        OREc.MoveNext
'        Next x
'    .cmbOffice.ComboItems(1).Selected = True
'    End If
'    OREc.Close
'    Set OREc = Nothing
'End With
'End Sub

'Private Sub Load_FundTypes()
'Dim OREc As New ADODB.Recordset
'Dim x As Integer
'
'With frmJEVNumberingAssignment_New
'    .cmbfundtype.ComboItems.Clear
'    OREc.Open ("Select * from tblRefBMS_Funds"), opndbaseFMIS, adOpenStatic, adLockOptimistic
'    If OREc.RecordCount > 0 Then
'        For x = 1 To OREc.RecordCount
'            .cmbfundtype.ComboItems.Add .cmbfundtype.ComboItems.Count + 1, OREc!FundCode & "ID", OREc!FundName, 2
'        OREc.MoveNext
'        Next x
'    .cmbfundtype.ComboItems(1).Selected = True
'    End If
'    OREc.Close
'    Set OREc = Nothing
'End With
'
'
'End Sub

'---------------RICHARD-----------------------------
'Private Sub Set_4_Old_Trans()
'On Error GoTo err
'
'    With frmJEVNumberingAssignment_New
'        .cmdsave.Visible = True
'        .txt_RCenter.Visible = False
'        .txt_FundType.Visible = False
'       ' .cmbOffice.Visible = True
'       ' .cmbOffice.Width = .txt_RCenter.Width
'       ' .cmbOffice.Left = .txt_RCenter.Left
'       ' .cmbOffice.Top = .txt_RCenter.Top
'        .cmbfundtype.Visible = True
'        .cmbfundtype.Width = .txt_FundType.Width
'        .cmbfundtype.Left = .txt_FundType.Left
'        .cmbfundtype.Top = .txt_FundType.Top
'    End With
'Exit Sub
'err:
'    MsgBox "Error: " & err.Description
'End Sub

Private Function fundmedium(fnd As String) As String
Dim opnfund As New ADODB.Recordset

    opnfund.Open "Select fundmedium from tblRefBMS_Funds where fundname='" & Trim(fnd) & "'", opndbaseFMIS, adOpenStatic, adLockOptimistic
    If Not opnfund.EOF Then
        fundmedium = Trim(opnfund(0))
    End If
End Function

'----------------------------------------------------


Private Sub grd_details_DblClick()
''MsgBox GetDVNobyChkNo(grd_details.TextMatrix(grd_details.Row, 3)), vbInformation, "DV Number"
'
'If grd_details.TextMatrix(grd_details.Row, 4) = 2 Then
'    ActiveFormCaller = Me.Name
'    ForTheGridRowNo = grd_details.Row
'    'Load_Offices   '---loads offices to the OFFICE combo - RICHARD
'    'Load_FundTypes '---loads functypes to the FUNDTYPE combo - RICHARD
'    If Len(grd_details.TextMatrix(grd_details.Row, 14)) <> 0 Then 'Kung Naa nay JEV No
'        frmJEVNumberingAssignment_New.txt_JEVNo.Text = grd_details.TextMatrix(grd_details.Row, 14)
'        frmJEVNumberingAssignment_New.txt_DVNo.Text = GetDVNobyChkNo(grd_details.TextMatrix(grd_details.Row, 3))
'
'
'    Else
''        With frmJEVPreparation
''            .Toolbar1_ButtonClick Toolbar1.Buttons(3)
''            .txtDVNo = GetDVNobyChkNo(grd_details.TextMatrix(grd_details.Row, 3))
''            .Show
''        End With
'        'frmJEVNumberingAssignment_New.txt_DVNo.Text = GetDVNobyChkNo(grd_details.TextMatrix(grd_details.Row, 3))
'
'    End If
'
'    If Trim(frmJEVNumberingAssignment_New.txt_RCenter) = "" And Trim(frmJEVNumberingAssignment_New.txt_FundType) = "" Then
'        'Set_4_Old_Trans '---set components for old transacted vouchers
'
'        '---passing the transaction to JEV Prepartion Form if Old Transaction---'
'        With frmJEVPreparation
'        .MSFlexGrid1.Clear
'            .Toolbar1_ButtonClick .Toolbar1.Buttons.Item(1)
'            .txtDVNo = GetDVNobyChkNo(grd_details.TextMatrix(grd_details.Row, 3))
'            .txtfund = fundmedium(Trim(cmb_FundType))
'            .txtDate = Format(grd_details.TextMatrix(grd_details.Row, 2), "mmmm dd, yyyy")
'            .txtDate.Locked = False
'            .LoadAccountsByFund Trim(.txtfund)
'            .optCheck.Value = True
'            .txtClaimant = Trim(grd_details.TextMatrix(grd_details.Row, 7))
'            .txtAmount = Trim(grd_details.TextMatrix(grd_details.Row, 9))
'            .txtParticular = Trim(grd_details.TextMatrix(grd_details.Row, 8))
'                If Trim(grd_details.TextMatrix(grd_details.Row, 6)) <> "" Then
'                    .txtRC = GetOfficeName(Val(Trim(grd_details.TextMatrix(grd_details.Row, 6))), "OfficeMedium")
'                End If
'            'GetOfficeName(DRec!RCenter, "OfficeMedium")
'            'frmJEVPreparation.LoadOffice
'            'frmJEVPreparation.cmbRC.Locked = False
'            'frmJEVPreparation.cmbRC.Visible = False
'            .txtAlobs.Locked = False
'            .txtRC.Locked = False
'            .Show 1
'        End With
'
'    Else
'        frmJEVNumberingAssignment_New.Show vbModal
'    End If
'Else
'    MsgBox "There is no Voucher Attachment for this Check!" & Chr(13) & Chr(13) & "Please Select a New..", vbInformation, "System Information"
'End If
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


Private Sub Timer1_Timer()
    Call SetGRidDetails

Timer1.Enabled = False
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

On Error GoTo handler

Select Case Button.Index
    Case 1 'Print----------------------------
            
            Call LoadReportedChks(List1.List(List1.ListIndex), DTPicker1.Month, DTPicker1.Year)
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
    MsgBox err.description
    Exit Sub
End If
End Sub
Private Sub LoadReportedChks(ByVal RCINo As String, ByVal trnMonth As Integer, ByVal TrnYear As Integer)
On Error GoTo bad

frmRptChckIssue.query = "Select * from tblCMS_CDRCIReport where RCINo ='" & RCINo & "' and year(CheckDate)=" & TrnYear & " and month(CheckDate)=" & trnMonth & "  order by OrderNo"
frmRptChckIssue.bnk = Me.txt_bank.Text & "-" & Me.txt_AcctNo
frmRptChckIssue.DTE = "Month: " & Format(DTPicker1.Value, "mmmm") & " " & DTPicker1.Year
frmRptChckIssue.fnd = Me.txt_fund.Text & "-" & Me.txt_acctName
frmRptChckIssue.RCI = RCINo
frmRptChckIssue.Show
Exit Sub
bad:
    MsgBox err.description
End Sub

Private Sub FindCheckNo(ByVal checkno As String)
Dim cc, ZZ As Integer

For cc = 0 To List1.ListCount - 1
DoEvents
    List1.ListIndex = cc
    For ZZ = 1 To grd_details.Rows - 1
    DoEvents
        grd_details.col = 3
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
Dim tmpVal As Long

On Error GoTo handler
If KeyCode = 13 Then
    If Len(Trim(txt_search.Text)) <> 0 Then
        
            tmpVal = GetIndex4ListBox(List1, txt_search.Text)
            If tmpVal <> 0 Then
                List1.ListIndex = tmpVal
            Else
                MsgBox "REport No. Not Found!", vbInformation, "System Information"
                txt_search.SelStart = 0
                txt_search.SelLength = Len(txt_search.Text)
                txt_search.SetFocus
            End If
       
    End If
End If
handler:
If err.Number <> 0 Then
    MsgBox err.description
End If
End Sub

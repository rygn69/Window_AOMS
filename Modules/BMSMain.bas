Attribute VB_Name = "BMSMain"

'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'+++++++++  Title           :   Budget Monitoring System (BMS)
'+++++++++                      A part of the Financial Management &
'+++++++++                      Information System(FMIS)
'+++++++++  Programmer      :   Eduard Emmanuel Dacillo Gatong
'+++++++++  Database Used   :   FMIS, PMIS
'+++++++++  Servers Used    :   FMIS, PMIS, Picture_Server
'+++++++++  Period          :   July 2004 - September 2006
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Public FMISDBLocation, PMISDBLocation, MEMISDBLocation, ReportLocation, PicturePath
Public opndb As New ADODB.Connection
Public opnPMIS As New ADODB.Connection
Public opnMEMIS As New ADODB.Connection
Public ReportDB As New ADODB.Connection
Public OpnReport As New ADODB.Recordset
Public opnRec As New ADODB.Recordset
Public Rec As New ADODB.Recordset
Public OpnOffice As New ADODB.Recordset
Public OfficeID As String
Public DivisionID As String
Public UserID As String
Public UserName As String
Public UserDesignation As String
Public Action As Integer
'Public TransFlagExpenses As Integer
Public TransFlagClaimant As Integer
Public TrackCallFlag As Integer
Public GroupCount As Integer
Public SavedALOBS As String
Public PTOAccts As String
Public ReportName As String
Public UserDiv As Integer
Public ViewerCaption As String
Public ProgChargeCode As Integer
Public ComputerName As String
Public GetTransYear As Integer
Public UpdateLocation As String
Public PGOCode As Integer
Public Indic As Boolean
Public ExConUserID As String
Public uidParameter As String
Public userRights As Boolean
Public strPreparedby As String
Public strPreparedbyPos As String
Public strNotedby As String
Public strNotedbyPos As String
Public strlbl As String
Public strSelOffice As String
'Public EditControlFlag As Integer

Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Const SW_SHOWNORMAL = 1

Public Function IsInFMISTrans(ByVal alobsno As String) As Boolean
Dim IRec As New ADODB.Recordset

IsInFMISTrans = False

IRec.Open ("Select * From tblFMIS_Transaction Where AlobsNo='" & Format(alobsno, "000-0000-00-00-0000") & "' And actioncode=1"), opndb, adOpenStatic, adLockOptimistic
If IRec.RecordCount > 0 Then
    IsInFMISTrans = True
End If
IRec.Close
Set IRec = Nothing

IRec.Open ("Select * From tblBMS_ExcessControl Where AlobsNo='" & Format(alobsno, "000-0000-00-00-0000") & "' And actioncode=1"), opndb, adOpenStatic, adLockOptimistic
If IRec.RecordCount > 0 Then
    IsInFMISTrans = True
End If
IRec.Close
Set IRec = Nothing

End Function

Public Function IsPaidByPTO(ByVal alobsno As String) As Boolean
Dim TRec As New ADODB.Recordset

IsPaidByPTO = False

TRec.Open ("Select * FRom tblCMS_CDTransactionDetails Where (AlobsNo like '%" & alobsno & "%' Or AlobsNo like '%" & Format(alobsno, "000-0000-00-00-####") & "%') and Actioncode=1 and PaidUnPaid=1"), opndb, adOpenStatic, adLockOptimistic
If TRec.RecordCount > 0 Then
    IsPaidByPTO = True
End If
TRec.Close
Set TRec = Nothing

End Function

'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++++ Purpose       : To trap edition of transaction after being processed in PTO.
'++++ Programmer    : Eduard Emmanuel D. Gatong
'++++ Date          : 01/26/2010
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Public Function InPTO(ByVal alobsno As String) As Boolean
Dim PRec As New ADODB.Recordset

InPTO = False

PRec.Open ("Select * From tblCMS_EXCashVerification Where alobsno='" & alobsno & "' and actioncode=1"), opndb, adOpenStatic, adLockOptimistic
If PRec.RecordCount > 0 Then
    InPTO = True
End If
PRec.Close
Set PRec = Nothing

End Function

Sub Main()

On Error GoTo edgeErrHandler

Call GetDBLocation

'Call CheckUpdate

Call ChangeScreen(1024, 768, 16)

frmSplash.Show
frmSplash.Refresh

'Call CreateDummy   ' temp disable

'Call GetDBLocation

opndb.ConnectionTimeout = 60
opndb.CommandTimeout = 60
opndb.CursorLocation = adUseClient

'opndb.Open "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;User ID=sa;Initial Catalog=fmis;Data Source=RYGN\SQLServer2008"
opndb.Open FMISDBLocation

opnPMIS.ConnectionTimeout = 60
opnPMIS.CursorLocation = adUseClient

'opnPMIS.Open "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;User ID=sa;Initial Catalog=pmis;Data Source=RYGN\SQLServer2008"
opnPMIS.Open PMISDBLocation

opnMEMIS.ConnectionTimeout = 60
opnMEMIS.CursorLocation = adUseClient

'opnMEMIS.Open "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;User ID=sa;Initial Catalog=memis;Data Source=RYGN\SQLServer2008"
opnMEMIS.Open MEMISDBLocation

'ReportDB.CursorLocation = adUseClient
'ReportDB.Open ReportLocation

frmLogin.Show

edgeErrHandler:
If Err.Number <> 0 Then
    MsgBox Err.Description & vbCrLf & vbCrLf & "Please check your server connection and try again!", vbExclamation + vbOKOnly, "BMS Security Center"
    frmSplash.Hide
    frmSplashEnd.Show 1
    'End
End If

End Sub

Public Function FindFile(ByVal DirFileName As String) As Integer

On Error GoTo NoFile

FileDateTime (DirFileName)     'test if the file exist
FindFile = 1
Exit Function

NoFile:

    FindFile = 0

End Function


Public Sub CheckUpdate()
Dim XUpdateCount As Integer
Dim SUpdateCount As Integer
Dim xhwd As Long


If UpdateLocation = "" Then
    MsgBox "BMS auto update feature on this computer is off!" & vbCrLf & vbCrLf & "Please contact your system administrator to activate BMS auto update feature.", vbExclamation + vbOKOnly, "BMS Security Center"
Else
    If FindFile(UpdateLocation & "\UpdateLog.edg") = 1 Then
        XUpdateCount = CInt(GetTxtFileData("[Updates]", "UpdateCount", App.Path & "\Common Files\Server.edg"))
        SUpdateCount = CInt(GetTxtFileData("[Updates]", "UpdateCount", UpdateLocation & "\UpdateLog.edg"))
        If XUpdateCount <> SUpdateCount Then
            If MsgBox("New updates are available for your BMS!" & vbCrLf & vbCrLf & "Do you want to update?", vbInformation + vbYesNo, "BMS Security Center") = vbYes Then
                If FindFile(App.Path & "\EDGEUpdate.exe") = 1 Then
                    ShellExecute xhwd, vbNullString, App.Path & "\EDGEUpdater.exe", vbNullString, vbNullString, SW_SHOWNORMAL
                    End
                Else
                    MsgBox "Cannot find BMS Updater!" & vbCrLf & vbCrLf & "Please contact your system administrator.", vbExclamation + vbOKOnly, "BMS Security Center"
                End If
            End If
        End If
    End If
End If

End Sub


'+++++++++++++++++++++++++++++++++++++++++++++++++
'+++++++++++++++++++ --- Note
'   All the Right commands in this program does not refer to
'   the VB default Right function, I made my own Right function
'   beacause I found out that the Right function of VB does not
'   work when you use ScrollingText control. My following Right
'   function solves the problem.
'+++++++++++++++++++ --- edge

Public Function Right(ByVal Txt As String, ByVal Num As Integer) As String
Dim x As Integer
Dim tempRight As String

On Error GoTo errEnd

tempRight = ""
For x = 1 To Num
    tempRight = Mid(Txt, Len(Txt) + 1 - x, 1) & tempRight
Next x

Right = Trim(tempRight)

errEnd:
End Function
'+++++++++++++++++++++++++++++++++++++++++++++++++

Public Function GetDBLocation()
'+++++++++ This is my old way of extracting data from text file. --- edge
'Dim X As Integer

'X = 0
'Open App.Path & "\Common Files\Server.edg" For Input As #1
'Do While Not EOF(1)
    'X = X + 1
    'If X = 1 Then Line Input #1, FMISDBLocation
    'If X = 2 Then Line Input #1, PMISDBLocation
    'If X = 3 Then Line Input #1, PicturePath
'Loop

'Close #1
'+++++++++ This is my new way of extracting data from text file. --- edge
ReportLocation = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Reports\tblBMSDummy.mdb;Persist Security Info=False"
FMISDBLocation = GetTxtFileData("[Server]", "fmis", App.Path & "\Common Files\Server.edg")
PMISDBLocation = GetTxtFileData("[Server]", "pmis", App.Path & "\Common Files\Server.edg")
PicturePath = GetTxtFileData("[Pictures]", "picloc", App.Path & "\Common Files\Server.edg")
OfficeID = GetTxtFileData("[Station]", "Office", App.Path & "\Common Files\Server.edg")
DivisionID = GetTxtFileData("[Station]", "Division", App.Path & "\Common Files\Server.edg")
UpdateLocation = GetTxtFileData("[Updates]", "UpdateLoc", App.Path & "\Common Files\Server.edg")
MEMISDBLocation = GetTxtFileData("[Server]", "memis", App.Path & "\Common Files\Server.edg")

strPreparedby = GetTxtFileData("[Signatories]", "Prepared by", App.Path & "\Common Files\Server.edg")
strPreparedbyPos = GetTxtFileData("[Signatories]", "Prepared by position", App.Path & "\Common Files\Server.edg")
strNotedby = GetTxtFileData("[Signatories]", "Noted by", App.Path & "\Common Files\Server.edg")
strNotedbyPos = GetTxtFileData("[Signatories]", "Noted by position", App.Path & "\Common Files\Server.edg")

'+++++++++ --- edge
End Function


Public Function GetFundName(ByVal ProgramCode As Long) As String
Dim FCode As Integer

opnRec.Open ("Select * From tblRefBMS_BudgetProgram Where FmisProgramCode=" & ProgramCode & " And YearOf=" & GetTransYear & " And ActionCode=1 And ProjectFlag=0"), opndb, adOpenStatic, adLockOptimistic
If opnRec.RecordCount <> 0 Then FCode = opnRec!FundCode

opnRec.Close
Set opnRec = Nothing

opnRec.Open ("Select * From tblRefBMS_Funds where FundCode=" & FCode & ""), opndb, adOpenStatic, adLockOptimistic
If opnRec.RecordCount <> 0 Then GetFundName = opnRec!FundName

opnRec.Close
Set opnRec = Nothing

End Function

Public Function GetFundName_NextYear(ByVal ProgramCode As Long) As String
Dim FCode As Integer

opnRec.Open ("Select * From tblRefBMS_BudgetProgram Where FmisProgramCode=" & ProgramCode & " And YearOf=" & GetTransYear + 1 & " And ActionCode=1 And ProjectFlag=0"), opndb, adOpenStatic, adLockOptimistic
If opnRec.RecordCount <> 0 Then FCode = opnRec!FundCode

opnRec.Close
Set opnRec = Nothing

opnRec.Open ("Select * From tblRefBMS_Funds where FundCode=" & FCode & ""), opndb, adOpenStatic, adLockOptimistic
If opnRec.RecordCount <> 0 Then GetFundName_NextYear = opnRec!FundName

opnRec.Close
Set opnRec = Nothing

End Function

Public Function GetFundCode(ByVal FMISOfficeCode As Long) As Integer

opnRec.Open ("Select FMISOfficeCode, FundCode From tblRefBMS_BudgetProgram Where FMISOfficeCode=" & FMISOfficeCode & " And ActionCode=1 And ProjectFlag=0 And YearOf=" & GetTransYear & " Group By FMISOfficeCode, FundCode"), opndb, adOpenStatic, adLockOptimistic
If opnRec.RecordCount <> 0 Then
    GetFundCode = opnRec!FundCode
End If
opnRec.Close
Set opnRec = Nothing

End Function

Public Function GetFundType(ByVal FundCode As Integer) As String

opnRec.Open ("Select * From tblRefBMS_Funds Where FundCode=" & FundCode & ""), opndb, adOpenStatic, adLockOptimistic
If opnRec.RecordCount <> 0 Then
    GetFundType = opnRec!FundName
End If
opnRec.Close
Set opnRec = Nothing

End Function

Public Function GetFundTypeShort(ByVal FundCode As Integer) As String

opnRec.Open ("Select * From tblRefBMS_Funds Where FundCode=" & FundCode & ""), opndb, adOpenStatic, adLockOptimistic
If opnRec.RecordCount <> 0 Then
    GetFundTypeShort = opnRec!FundMedium
End If
opnRec.Close
Set opnRec = Nothing

End Function

Public Function SumReleasePS(ByVal ProgramCode As Long) As Currency

opnRec.Open ("Select Sum(AmountPS) as SumPS From tblBMS_Releases Where FMISProgramCode=" & ProgramCode & ""), opndb, adOpenStatic, adLockOptimistic
SumReleasePS = opnRec!SumPS
opnRec.Close
Set opnRec = Nothing

End Function

Public Function SumReleaseMOOE(ByVal ProgramCode As Long) As Currency

opnRec.Open ("Select Sum(AmountMOOE) as SumMOOE From tblBMS_Releases Where FMISProgramCode=" & ProgramCode & ""), opndb, adOpenStatic, adLockOptimistic
SumReleaseMOOE = opnRec!SumMOOE
opnRec.Close
Set opnRec = Nothing

End Function

Public Function SumReleaseCO(ByVal ProgramCode As Long) As Currency

opnRec.Open ("Select Sum(AmountCO) as SumCO From tblBMS_Releases Where FMISProgramCode=" & ProgramCode & ""), opndb, adOpenStatic, adLockOptimistic
SumReleaseCO = opnRec!SumCO
opnRec.Close
Set opnRec = Nothing

End Function

'+++++++++++++++++++++++++++++++++++++++++++++++++++
'   Note: the following are the syntax and values to be used in using the
'               CheckRights function.

'           Budget Preparation     :   CheckRights("BudgetPreparation")
'           Budget Monthly Release :   CheckRights("MonthlyRelease")
'           Budget Annual Budget   :   CheckRights("AnnualBudget")
'           Budget Control         :   CheckRights("BudgetControl")
'           User Profiles          :   CheckRights("UserProfile")
'           Budget Program Tool    :   CheckRights("ToolProgram")
'           User Account Tool      :   CheckRights("ToolAccount")
'           User Program Tool Next :   CheckRights("ToolProgramNext")
'           User Account Tool Next :   CheckRights("ToolAccountNext")
'           Override               :   CheckRights("CanOveride")
'           Reports                :   CheckRights("Reports")
'+++++++++++++++++++++++++++++++++++++++++++++++++++

Public Function CheckRights(ByVal MenuID As String) As Long
Dim RightRec As New ADODB.Recordset

RightRec.Open ("Select * From tblBMS_Users Where UserID='" & UserID & "' And ActionCode=1"), opndb, adOpenStatic, adLockOptimistic
If RightRec.RecordCount <> 0 Then
    CheckRights = InStr(RightRec!AccessRights, MenuID)
End If
RightRec.Close
Set RightRec = Nothing

End Function

Public Function EDGEncrypt(ByVal PassWord As String) As String
Dim x As Integer
Dim newPass As String

newPass = ""

For x = 1 To Len(Trim(PassWord))
    newPass = newPass & Chr(Asc(Mid(UCase(PassWord), x, 1)) + Len(Trim(PassWord)) + 1 - x)
Next x

EDGEncrypt = newPass

End Function


Public Function EDGEDecrypt(ByVal PassWord As String) As String
Dim x, y As Integer
Dim newPass As String

newPass = ""

For x = 1 To Len(PassWord)
    y = Len(PassWord) + 1 - x
    newPass = newPass & Chr(Asc(Mid(PassWord, x, 1)) - y)
Next x

EDGEDecrypt = newPass

End Function

Public Function getEmployeeName(ByVal SwipeID As String) As String

opnRec.Open ("Select * From Employee Where SwipEmployeeID='" & SwipeID & "'"), opnPMIS, adOpenStatic, adLockOptimistic
If opnRec.RecordCount <> 0 Then
    getEmployeeName = opnRec!FirstName & " " & opnRec!MI & ". " & opnRec!LastName
End If
opnRec.Close
Set opnRec = Nothing

End Function

Public Function getFunctionCode(ByVal FMISProgramCode As Long) As String
Dim ORec As New ADODB.Recordset

If FMISProgramCode = 43 Or FMISProgramCode = 50 Or FMISProgramCode = 51 Or FMISProgramCode = 52 Or FMISProgramCode = 53 Or FMISProgramCode = 54 Or FMISProgramCode = 55 Or FMISProgramCode = 58 Then
    ORec.Open ("Select FunctionCode as FunctionID  From [tblBMS_NonOfficeCode] Where [ProgCode]=" & FMISProgramCode & " And ActionCode=1 And YearOf=" & GetTransYear & ""), opndb, adOpenStatic, adLockOptimistic
Else
    ORec.Open ("Select * From tblRefBMS_BudgetProgram Where FMISProgramCode=" & FMISProgramCode & " And ActionCode=1 And YearOf=" & GetTransYear & ""), opndb, adOpenStatic, adLockOptimistic
End If
If ORec.RecordCount <> 0 Then
    getFunctionCode = ORec!FunctionID
End If
ORec.Close
Set ORec = Nothing

End Function

Public Function getNonOfficeFunctionCode(ByVal FmisProgCode As Long, ByVal FmisAcctCode As Long) As String
Rec.Open ("Select * from [tblBMS_AnnualBudget_Account] where FMISProgramCode=" & FmisProgCode & " and FMISAccountCode =" & FmisAcctCode & " And ActionCode=1 And YearOf=" & GetTransYear & " "), opndb, adOpenStatic, adLockOptimistic
If Rec.RecordCount <> 0 Then
    getNonOfficeFunctionCode = IIf(IsNull(Rec!NonOfficeFunctionid), 0, Rec!NonOfficeFunctionid)
End If
Rec.Close
Set Rec = Nothing
End Function

'use for nonoffice
'Public Function getObRFunctionCode(ByVal FMISProgramCode As Long, ByVal FMISAccountCode As Long, ByVal Num As Boolean) As String
Public Function getObRFunctionCode(ByVal FMISProgramCode As Long) As String

opnRec.Open ("Select * From tblBMS_ObRFunctionCode Where ProgCode=" & FMISProgramCode & " And YearOf=" & GetTransYear & " and actioncode=1"), opndb, adOpenStatic, adLockOptimistic
If opnRec.RecordCount <> 0 Then
    getObRFunctionCode = opnRec!FunctionCode
Else
    'If Num = True Then
        getObRFunctionCode = getFunctionCode(FMISProgramCode)
    'Else
    '    getObRFunctionCode = getNonOfficeFunctionCode(FMISProgramCode, FMISAccountCode)
    'End If
End If
opnRec.Close
Set opnRec = Nothing

End Function

Public Function getAlobsSequence(ByVal FundCode As String) As Long
'Public Function getAlobsSequence() As Long

If FundCode = 118 Then
    opnRec.Open ("Select * From vwBMS_SortedOR Where YearOf=" & Format(ServerDate, "yy") & " And MonthOf=" & Month(ServerDate) & " and Fund = 118 Order by Series"), opndb, adOpenStatic, adLockOptimistic
ElseIf FundCode = 201 Then
    opnRec.Open ("Select * From vwBMS_SortedOR Where YearOf=" & Format(ServerDate, "yy") & " And MonthOf=" & Month(ServerDate) & " and Fund = 201 Order by Series"), opndb, adOpenStatic, adLockOptimistic
Else
    opnRec.Open ("Select * From vwBMS_SortedOR Where YearOf=" & Format(ServerDate, "yy") & " And MonthOf=" & Month(ServerDate) & " and Fund <> 118 and Fund <> 201  Order by Series"), opndb, adOpenStatic, adLockOptimistic
End If
If opnRec.RecordCount <> 0 Then
    opnRec.MoveLast
    getAlobsSequence = opnRec!Series
Else
    getAlobsSequence = 0
End If
opnRec.Close
Set opnRec = Nothing

End Function

Public Function getTransTypeSeqNo(ByVal TransTypeCode As Long) As Long

opnRec.Open ("Select * From tblFMIS_Transaction Where TransTypeCode=" & TransTypeCode & " And '20'+left(right(AlobsNo,10),2)='" & GetTransYear & "' Order By TransTypeSeqNo Asc"), opndb, adOpenStatic, adLockOptimistic
If opnRec.RecordCount <> 0 Then
    opnRec.MoveLast
    getTransTypeSeqNo = opnRec!TransTypeSeqNo
Else
    getTransTypeSeqNo = 0
End If
opnRec.Close
Set opnRec = Nothing

End Function


Public Function getClaimantName(ByVal ClaimantCode As Long, SourceCode As Long) As String
Dim RecOpn As New ADODB.Recordset

Select Case SourceCode
Case 1: RecOpn.Open ("Select * From tblREF_AIS_Offices Where FMISOfficeID=" & ClaimantCode & ""), opndb, adOpenStatic, adLockOptimistic
Case 2: RecOpn.Open ("Select * From Employee Where SwipEmployeeID=" & ClaimantCode & ""), opnPMIS, adOpenStatic, adLockOptimistic
Case 3: RecOpn.Open ("Select * From tblRefBMS_ClaimantCompany Where ActionCode=1 And ClaimantCode=" & ClaimantCode & ""), opndb, adOpenStatic, adLockOptimistic
Case 4: RecOpn.Open ("Select * From tblRefBMS_ClaimantOtherAgencies Where ActionCode=1 And ClaimantCode=" & ClaimantCode & ""), opndb, adOpenStatic, adLockOptimistic
Case 5: RecOpn.Open ("Select * From tblRefBMS_ClaimantOtherIndividual Where ActionCode=1 And ClaimantCode=" & ClaimantCode & ""), opndb, adOpenStatic, adLockOptimistic
End Select

If RecOpn.RecordCount <> 0 Then
    Select Case SourceCode
    Case 1: getClaimantName = RecOpn!OfficeName
    Case 2: getClaimantName = RecOpn!LastName & ", " & RecOpn!FirstName & " " & RecOpn!MI & "."
    Case 3: getClaimantName = RecOpn!Name
    Case 4: getClaimantName = RecOpn!Name
    Case 5: getClaimantName = RecOpn!LastName & ", " & RecOpn!FirstName & " " & RecOpn!MI & "."
    End Select
End If

RecOpn.Close
Set RecOpn = Nothing

End Function


Public Function getAccountName(ByVal FMISAccountCode As Long) As String

OpnOffice.Open ("Select * From tblREF_AIS_ChartofAccounts Where FMISAccountCode=" & FMISAccountCode & ""), opndb, adOpenStatic, adLockOptimistic
If OpnOffice.RecordCount <> 0 Then
    getAccountName = OpnOffice!AccountNameFull
End If
OpnOffice.Close
Set OpnOffice = Nothing

End Function


Public Function getAccountCode(ByVal AccountNameFull As String) As Long

OpnOffice.Open ("Select * From tblREF_AIS_ChartofAccounts Where AccountNameFull='" & AccountNameFull & "'"), opndb, adOpenStatic, adLockOptimistic
If OpnOffice.RecordCount <> 0 Then
    getAccountCode = OpnOffice!FMISAccountCode
End If
OpnOffice.Close
Set OpnOffice = Nothing

End Function


Public Function getChildAccountCode(ByVal FMISAccountCode As Long) As String

OpnOffice.Open ("Select * From tblREF_AIS_ChartofAccounts Where FMISAccountCode=" & FMISAccountCode & ""), opndb, adOpenStatic, adLockOptimistic
If OpnOffice.RecordCount <> 0 Then
    getChildAccountCode = OpnOffice!ChildAccountCode
End If
OpnOffice.Close
Set OpnOffice = Nothing

End Function


Public Function GetOfficeName(ByVal FMISProgramCode As Long) As String

opnRec.Open ("Select * From vwBMSProgramCode_OfficeName Where FMISProgramCode=" & FMISProgramCode & " And YearOf=" & GetTransYear & ""), opndb, adOpenStatic, adLockOptimistic
If opnRec.RecordCount <> 0 Then
    GetOfficeName = opnRec!OfficeName
End If
opnRec.Close
Set opnRec = Nothing

End Function

Public Function GetOfficeName_NextYear(ByVal FMISProgramCode As Long) As String

opnRec.Open ("Select * From vwBMSProgramCode_OfficeName_NextYear Where FMISProgramCode=" & FMISProgramCode & ""), opndb, adOpenStatic, adLockOptimistic
If opnRec.RecordCount <> 0 Then
    GetOfficeName_NextYear = opnRec!OfficeName
End If
opnRec.Close
Set opnRec = Nothing

End Function


Public Function GetOfficeID(ByVal FMISProgramCode As Long) As Long

opnRec.Open ("Select * From vwBMSProgramCode_OfficeName Where FMISProgramCode=" & FMISProgramCode & " And YearOf=" & GetTransYear & ""), opndb, adOpenStatic, adLockOptimistic
If opnRec.RecordCount <> 0 Then
    GetOfficeID = opnRec!FMISOfficeID
End If
opnRec.Close
Set opnRec = Nothing

End Function

Public Function GetOfficeID_NextYear(ByVal FMISProgramCode As Long) As Long

opnRec.Open ("Select * From vwBMSProgramCode_OfficeName_NextYear Where FMISProgramCode=" & FMISProgramCode & ""), opndb, adOpenStatic, adLockOptimistic
If opnRec.RecordCount <> 0 Then
    GetOfficeID_NextYear = opnRec!FMISOfficeID
End If
opnRec.Close
Set opnRec = Nothing

End Function

'+++++++++++++ for budget purposes only. --- edge
Public Function GetProgramName(ByVal FMISProgramCode As Long) As String

opnRec.Open ("Select * From tblRefBMS_BudgetProgram Where FMISProgramCode=" & FMISProgramCode & " And ActionCode=1 And ProjectFlag=0 And YearOf=" & GetTransYear & ""), opndb, adOpenStatic, adLockOptimistic
If opnRec.RecordCount <> 0 Then
    GetProgramName = opnRec!ProgramDescription
End If
opnRec.Close
Set opnRec = Nothing

End Function

Public Function GetProgramName_NextYear(ByVal FMISProgramCode As Long) As String

opnRec.Open ("Select * From tblRefBMS_BudgetProgram_NextYear Where FMISProgramCode=" & FMISProgramCode & ""), opndb, adOpenStatic, adLockOptimistic
If opnRec.RecordCount <> 0 Then
    GetProgramName_NextYear = opnRec!ProgramDescription
End If
opnRec.Close
Set opnRec = Nothing

End Function

Public Function getOOEName(ByVal OOECode As Long) As String
Dim OpnOOE As New ADODB.Recordset

OpnOOE.Open ("Select * From tblBMS_ObjectOfExpenditures Where OOECode=" & OOECode & ""), opndb, adOpenStatic, adLockOptimistic
If OpnOOE.RecordCount <> 0 Then
    getOOEName = OpnOOE!OOEName
End If
OpnOOE.Close
Set OpnOOE = Nothing

End Function

'+++++ For Budget purposes only. --- edge
Public Function SumAnnualBudgetOOE_Office(ByVal FMISOfficeCode As Long, ByVal OOECode As Long) As Currency
Dim RecNoAcct As New ADODB.Recordset

                    '++++++ This will get the total amount of progams with account. --- edge
opnRec.Open ("SELECT * From vwBMSAnnualBudgetOOE_Office Where FmisOfficeCode=" & FMISOfficeCode & " And OOECode=" & OOECode & " And YearOf=" & GetTransYear & ""), opndb, adOpenStatic, adLockOptimistic
If opnRec.RecordCount <> 0 Then
    SumAnnualBudgetOOE_Office = opnRec!Amount
Else
    SumAnnualBudgetOOE_Office = "0.00"
End If
opnRec.Close
Set opnRec = Nothing

'++++++ Deleted : 08/04/2005 --- edge
                    '++++++ This will get the total amount of progams without account. --- edge
'RecNoAcct.Open ("Select * From vwBMSAnnualBudgetOOE_Office_NoAcct Where FMISOfficeCode=" & FMISOfficeCode & " And OOECode=" & OOECode & ""), opndb, adOpenStatic, adLockOptimistic
'If RecNoAcct.RecordCount <> 0 Then
'    SumAnnualBudgetOOE_Office = SumAnnualBudgetOOE_Office + RecNoAcct!Amount
'End If
'RecNoAcct.Close
'Set RecNoAcct = Nothing
'++++++ --- edge
End Function

Public Function CreateDummy()
FileCopy App.Path & "\Common Files\BMSDummy", App.Path & "\Reports\tblBMSDummy.mdb"
End Function

'++++++++++++ this code was taken from JVBoniza.dll, which is programmed by Engr. Jumar V. Boniza --- edge
'Public Function ConsiderApostrophe(ByVal ggg As String) As String
'    Dim QLocator(0 To 200) As Byte
'    Dim EditedString As String
'    Dim qqq As Byte
'    qqq = 0
'    QLocator(0) = InStr(ggg, "'")
'    If QLocator(0) > 0 Then
'        EditedString = Mid(ggg, 1, QLocator(0)) & "'" & Mid(ggg, QLocator(0) + 1, Len(ggg))
'        Do Until InStr(QLocator(qqq) + 2, EditedString, "'") = 0
'                If InStr(QLocator(qqq) + 2, EditedString, "'") > 0 Then
'                    QLocator(qqq + 1) = InStr(QLocator(qqq) + 2, EditedString, "'")
'                    EditedString = Mid(EditedString, 1, InStr(QLocator(qqq) + 2, EditedString, "'")) & "'" & Mid(EditedString, InStr(QLocator(qqq) + 2, EditedString, "'") + 1, Len(EditedString))
'                End If
'            qqq = qqq + 1
'        Loop
'    Else
'            EditedString = ggg
'    End If
'    ConsiderApostrophe = EditedString
'End Function
'++++++++++++ --- edge


'++++++++++++ This function is my version of reading data from text files --- edge
Public Function GetTxtFileData(ByVal txtTable As String, ByVal txtField As String, TxtFile As String) As String
Dim TableSearch As Integer
Dim DummyData
Dim strPos As Long

On Error GoTo edgeErr

TableSearch = 0

Open TxtFile For Input As #1
Do While Not EOF(1)
    Line Input #1, DummyData
    If Mid(Trim(DummyData), 1, 1) <> ";" Then
        If TableSearch = 0 Then
            If Trim(DummyData) = Trim(txtTable) Then TableSearch = 1
        Else
                strPos = InStr(1, DummyData, txtField, vbTextCompare)
                If strPos <> 0 Then
                    GetTxtFileData = Right(DummyData, (Len(DummyData) - Len(txtField)) - 1)
                    TableSearch = 0
                End If
        End If
    End If
Loop

Close #1

If GetTxtFileData = "" Then MsgBox "No Match Found in Text File!", vbExclamation + vbOKOnly, "BMS Information Center"

GoTo endoffunction

edgeErr:
MsgBox "Error Reading Text File!" & vbCrLf & vbCrLf & Err.Description, vbExclamation + vbOKOnly, "BMS Security Center"

endoffunction:

End Function
'++++++++++++ --- edge


Public Function MediumNameToFundName(ByVal FundMediumName As String) As String
Dim OpnMedium As New ADODB.Recordset

OpnMedium.Open ("Select * From tblRefBMS_Funds Where FundMedium='" & FundMediumName & "'"), opndb, adOpenStatic, adLockOptimistic
If OpnMedium.RecordCount <> 0 Then
    MediumNameToFundName = OpnMedium!FundName
End If
OpnMedium.Close
Set OpnMedium = Nothing

End Function

'++++++ For budget purposes only. --- edge
Public Function getTotalOtherFundsReleased(ByVal OtherFundMedium As String) As Currency
Dim TotalRec As New ADODB.Recordset

TotalRec.Open ("Select * From vwBMS_TotalOtherFunds_Released Where OtherFundMedium='" & OtherFundMedium & "' And YearOf=" & GetTransYear & ""), opndb, adOpenStatic, adLockOptimistic
If TotalRec.RecordCount <> 0 Then
    getTotalOtherFundsReleased = TotalRec!Amount
Else
    getTotalOtherFundsReleased = 0
End If
TotalRec.Close
Set TotalRec = Nothing

End Function


Public Function FundCodeToFundMedium(ByVal FundCode As Long) As String
Dim MedRec As New ADODB.Recordset

MedRec.Open ("Select * From tblRefBMS_Funds Where FundCode=" & FundCode & ""), opndb, adOpenStatic, adLockOptimistic
If MedRec.RecordCount <> 0 Then
    FundCodeToFundMedium = MedRec!FundMedium
Else
    FundCodeToFundMedium = "Not Found"
    MsgBox "Invalid Fund Code!", vbExclamation + vbOKOnly, "BMS Security Center"
End If
MedRec.Close
Set MedRec = Nothing

End Function

Public Function getTypeReleasedAmount(ByVal TypeID As Long, ByVal OOECode As Long) As Currency
Dim ReleaseRec As New ADODB.Recordset

ReleaseRec.Open ("Select * From vwBMS_PlanTotalReleaseType_OOE Where TypeID=" & TypeID & " And YearOf=" & GetTransYear & ""), opndb, adOpenStatic, adLockOptimistic
If ReleaseRec.RecordCount <> 0 Then
    Select Case OOECode
        Case 1:     getTypeReleasedAmount = ReleaseRec!Totalps
        Case 2:     getTypeReleasedAmount = ReleaseRec!totalmooe
        Case 3:     getTypeReleasedAmount = ReleaseRec!totalco
        Case Else:  getTypeReleasedAmount = 0
    End Select
Else
    getTypeReleasedAmount = 0
End If
ReleaseRec.Close
Set ReleaseRec = Nothing

End Function

Public Function GetMonthNum(ByVal MonthOf As String) As Long

Select Case MonthOf
Case "January": GetMonthNum = 1
Case "February": GetMonthNum = 2
Case "March": GetMonthNum = 3
Case "April": GetMonthNum = 4
Case "May": GetMonthNum = 5
Case "June": GetMonthNum = 6
Case "July": GetMonthNum = 7
Case "August": GetMonthNum = 8
Case "September": GetMonthNum = 9
Case "October": GetMonthNum = 10
Case "November": GetMonthNum = 11
Case "December": GetMonthNum = 12
Case Else: GetMonthNum = 0
End Select

End Function

Public Function getLastProgramCode() As Long
Dim RecProg As New ADODB.Recordset

RecProg.Open ("Select * from tblRefBMS_BudgetProgram Order By FmisProgramCode"), opndb, adOpenStatic, adLockOptimistic
If RecProg.RecordCount <> 0 Then
    RecProg.MoveLast
    getLastProgramCode = RecProg!FMISProgramCode
Else
    getLastProgramCode = 0
End If
RecProg.Close
Set RecProg = Nothing

End Function

'====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====
'+++ Function Created        : 04/26/2006
'+++ Created by              : EDGE
'+++ Purpose / Description   : Added to allow per account monthly release.
'+++                           This function returns the Annual budget appropriation of the account and its
'+++                           Object of Expenditure.
'+++ Functions / Subs Used   : None
'+++ Tables / Views Used     : tblBMS_AnnualBudget_Account, tblBMS_AnnualBudget
'+++ Requirements / Inputs   : FMISProgramCode, FMISAccountCode, Integer variable that will hold the OOECode.
'+++ Output                  : Annual Budget Appropriation of the account in currency, Object of Expenditure Code in integer
'+++ Limitations             : This can only be used for budget accounts, not for planning accounts.
'====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====
Public Function getAnnualBudget_Account(ByVal PCode As Long, ByVal ACode As Long, OOECode As Long) As Currency
Dim xREc As New ADODB.Recordset

xREc.Open ("Select * From tblBMS_AnnualBudget_Account Where FMISProgramCode=" & PCode & " and FMISAccountCode=" & ACode & " and ActionCode=1 and ProjectFlag=0 and YearOf=" & GetTransYear & ""), opndb, adOpenStatic, adLockOptimistic
If xREc.RecordCount <> 0 Then
    OOECode = xREc!OOECode
End If
xREc.Close
Set xREc = Nothing

xREc.Open ("Select * From tblBMS_AnnualBudget Where ActionCode=1 and FmisProgramCode=" & PCode & " and FMISAccountCode=" & ACode & " and YearOf=" & GetTransYear & " and ProjectFlag=0"), opndb, adOpenStatic, adLockOptimistic
If xREc.RecordCount <> 0 Then
    getAnnualBudget_Account = xREc!AllotedAmount
Else
    getAnnualBudget_Account = 0
End If
xREc.Close
Set xREc = Nothing

End Function
'EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE

Public Function getAnnualBudget_AccountNew(ByVal PCode As Long, ByVal ACode As Long) As Currency
Dim xREc As New ADODB.Recordset

xREc.Open ("Select * From tblBMS_AnnualBudget Where ActionCode=1 and FmisProgramCode=" & PCode & " and FMISAccountCode=" & ACode & " and YearOf=" & GetTransYear & " and ProjectFlag=0"), opndb, adOpenStatic, adLockOptimistic
If xREc.RecordCount <> 0 Then
    getAnnualBudget_AccountNew = xREc!AllotedAmount
Else
    getAnnualBudget_AccountNew = 0
End If
xREc.Close
Set xREc = Nothing

End Function
'RYGN=======RYGN=======RYGN=======RYGN=======RYGN=======RYGN=======RYGN=======RYGN=======RYGN=======RYGN=======RYGN=======RYGN=======RYGN=======RYGN=======RYGN=======RYGN=======RYGN=======RYGN

'====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====
'+++ Function Created        : 05/02/2006
'+++ Created by              : EDGE
'+++ Purpose / Description   : Added to allow per account monthly release.
'+++                           This function gets the total amount released for the particular budget office account.
'+++ Functions / Subs Used   : None
'+++ Tables / Views Used     : vwBMS_TotalRelease_Account
'+++ Requirements / Inputs   : FMISProgramCode, FMISAccountCode, OOECode
'+++ Output                  : Total Released per Account in currency
'+++ Limitations             : This can only be used for budget accounts, not for planning accounts.
'====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====
Public Function GetTotalRelease_Account(ByVal FMISPCode As Long, ByVal FMISACode As Long, ByVal OOECode As Long) As Currency
Dim CRec As New ADODB.Recordset

CRec.Open ("Select * From vwBMS_TotalRelease_Account Where FMISProgramCode=" & FMISPCode & " And FMISAccountCode=" & FMISACode & " And YearOf=" & GetTransYear & ""), opndb, adOpenStatic, adLockOptimistic
If CRec.RecordCount <> 0 Then
    Select Case OOECode
    Case 1: GetTotalRelease_Account = CRec!AmountPS
    Case 2: GetTotalRelease_Account = CRec!AmountMOOE
    Case 3: GetTotalRelease_Account = CRec!AmountCO
    End Select
Else
    GetTotalRelease_Account = 0
End If
CRec.Close
Set CRec = Nothing

End Function
'EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE

Public Function GetReverseTotalAmountTo(ByVal FMISPCode As Long, ByVal FMISACode As Long, ByVal OOECode As Long) As Currency
Dim Rev As New ADODB.Recordset
    Rev.Open ("Select sum(amount) as Amount from vwBMS_ReversionTotalAmount where [ToProgCode]=" & FMISPCode & " and [ToAccountCode]=" & FMISACode & " and ToOOECode=" & OOECode & " and YearOf=" & GetTransYear & " group by [ToProgCode],[ToAccountCode],YearOf "), opndb, adOpenStatic, adLockOptimistic
    If Not Rev.EOF Then
        GetReverseTotalAmountTo = Rev!Amount
    Else
        GetReverseTotalAmountTo = 0
    End If
Rev.Close
Set Rev = Nothing
End Function

Public Function GetReverseDeductedAmount(ByVal FMISPCode As Long, ByVal FMISACode As Long, ByVal OOECode As Long) As Currency
Dim Rev As New ADODB.Recordset
    Rev.Open ("Select sum(amount) as Amount from [vwBMS_ReversionDeductedAmount] where [FromProgCode]=" & FMISPCode & " and [FromAccountCode]=" & FMISACode & "and FromOOECode=" & OOECode & " and YearOf=" & GetTransYear & " group by [FromProgCode],[FromAccountCode],YearOf "), opndb, adOpenStatic, adLockOptimistic
    If Not Rev.EOF Then
        GetReverseDeductedAmount = Rev!Amount
    Else
        GetReverseDeductedAmount = 0
    End If
Rev.Close
Set Rev = Nothing
End Function

Public Function getObjctExpndtures(ByVal FMISPCode As Long, ByVal FMISACode As Long, ByVal FMISPCode1 As Long, ByVal FMISACode1 As Long, ByVal tRn As Long, ByVal OOECode As Long) As String
Dim xREc As New ADODB.Recordset

xREc.Open ("Select * From tblBMS_Reversion Where FromProgCode=" & FMISPCode & " and FromAccountCode=" & FMISACode & "  and ToProgCode=" & FMISPCode1 & " and ToAccountCode=" & FMISACode1 & " and trnno=" & tRn & " and ToOOECode=" & OOECode & " and ActionCode=1 and YearOf=" & GetTransYear & ""), opndb, adOpenStatic, adLockOptimistic
If xREc.RecordCount > 0 Then
        getObjctExpndtures = getOOE(xREc!ToOOECode)
        OOECode = xREc!ToOOECode
Else
    xREc.Close
    Set xREc = Nothing
    xREc.Open ("Select * From tblBMS_AnnualBudget_Account Where FMISProgramCode=" & FMISPCode1 & " and FMISAccountCode=" & FMISACode1 & " and ActionCode=1 and ProjectFlag=0 and YearOf=" & GetTransYear & ""), opndb, adOpenStatic, adLockOptimistic
    If xREc.RecordCount <> 0 Then
        getObjctExpndtures = getOOE(xREc!OOECode)
        OOECode = xREc!OOECode
    End If
End If
xREc.Close
Set xREc = Nothing
End Function

Public Function getObjectName(ByVal FMISPCode As Long, ByVal FMISACode As Long, ByVal trnN As Long, ByVal OOECode As Long) As String
Dim xREc As New ADODB.Recordset

xREc.Open ("Select * From tblBMS_Reversion Where FromProgCode=" & FMISPCode & " and FromAccountCode=" & FMISACode & " and FromOOECode=" & OOECode & " and trnno=" & trnN & " and ActionCode=1 and YearOf=" & GetTransYear & ""), opndb, adOpenStatic, adLockOptimistic
If xREc.RecordCount > 0 Then
        getObjectName = getOOE(xREc!OOECode)
        OOECode = xREc!OOECode
Else
    xREc.Close
    Set xREc = Nothing
    xREc.Open ("Select * From tblBMS_AnnualBudget_Account Where FMISProgramCode=" & FMISPCode & " and FMISAccountCode=" & FMISACode & " and ActionCode=1 and ProjectFlag=0 and YearOf=" & GetTransYear & ""), opndb, adOpenStatic, adLockOptimistic
    If xREc.RecordCount <> 0 Then
        getObjectName = getOOE(xREc!OOECode)
        OOECode = xREc!OOECode
    End If
End If
xREc.Close
Set xREc = Nothing
End Function

Public Function getOOE(ByVal OOE As Long) As String
Dim recX As New ADODB.Recordset
    recX.Open ("Select * from [tblBMS_ObjectOfExpenditures] where OOECode=" & OOE & " order by OOECode ASC"), opndb, adOpenStatic, adLockOptimistic
    If recX.RecordCount > 0 Then getOOE = recX!OOEName
    recX.Close
    Set recX = Nothing
End Function

'====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====
'+++ Function Created        : 08/03/2007
'+++ Created by              : EDGE
'+++ Purpose / Description   : Added to allow per account monthly release.
'+++                           This function gets the total amount released for the particular budget office account.
'+++ Functions / Subs Used   : None
'+++ Tables / Views Used     : tblBMS_Releases
'+++ Requirements / Inputs   : FMISProgramCode, FMISAccountCode, OOECode, EndMonth
'+++ Output                  : Total Released per Account in currency for the given period
'+++ Limitations             : This can only be used for budget accounts, not for planning accounts.
'====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====
Public Function GetTotalPeriodicRelease_Account(ByVal FMISPCode As Long, ByVal FMISACode As Long, ByVal OOECode As Long, ByVal EndMonth As Long) As Currency
Dim TCRec As New ADODB.Recordset

'If FMISPCode = 44 Or FMISPCode = 54 Then
'    MsgBox ""
'
'End If
TCRec.Open ("Select Sum(AmountPS) as AmountPS, Sum(AmountMOOE) as AmountMOOE, Sum(AmountCO) as AmountCO From tblBMS_Releases Where FMISProgramCode=" & FMISPCode & " And FMISAccountCode=" & FMISACode & " And YearOf=" & GetTransYear & " And ActionCode=1 And MonthOf<=" & EndMonth & ""), opndb, adOpenStatic, adLockOptimistic
If TCRec.RecordCount <> 0 Then
    
    Select Case OOECode
    Case 1: GetTotalPeriodicRelease_Account = IIf(IsNull(TCRec!AmountPS), 0, TCRec!AmountPS)
    Case 2: GetTotalPeriodicRelease_Account = IIf(IsNull(TCRec!AmountMOOE), 0, TCRec!AmountMOOE)
    Case 3: GetTotalPeriodicRelease_Account = IIf(IsNull(TCRec!AmountCO), 0, TCRec!AmountCO)
    End Select
Else
    GetTotalPeriodicRelease_Account = 0
End If
TCRec.Close
Set TCRec = Nothing

End Function
'EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE

'====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====
'+++ Function Created        : 05/03/2006
'+++ Created by              : EDGE
'+++ Purpose / Description   : This function gets the total amount controlled for the particular budget office account.
'+++ Functions / Subs Used   : None
'+++ Tables / Views Used     : vwBMS_TotalControl_Account
'+++ Requirements / Inputs   : FMISProgramCode, FMISAccountCode
'+++ Output                  : Total Control per Account in currency
'+++ Limitations             : This can only be used for budget accounts, not for planning accounts.
'====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====
Public Function GetTotalControl_Account(ByVal FMISProgramCode As Long, ByVal FMISAccountCode As Long) As Currency
Dim CRec As New ADODB.Recordset

'Debug.Print "Select * From vwBMS_TotalControl_Account Where FmisProgramCode=" & FMISProgramCode & " And Budget_AcctCharge=" & FMISAccountCode & " And YearOf=" & GetTransYear & ""
If GetOfficeCode(strSelOffice) = True And strlbl = "EDIT" Then
    CRec.Open ("Select sum(Amount) as Amount From vwBMS_PGO_TotalControl_Account Where FmisProgramCode=" & FMISProgramCode & " And Budget_AcctCharge=" & FMISAccountCode & " And YearOf=" & GetTransYear & ""), opndb, adOpenStatic, adLockOptimistic
Else
    If CheckUserControl(UserID) = True Then
        CRec.Open ("Select sum(Amount) as Amount From vwBMS_PGO_TotalControl_Account Where FmisProgramCode=" & FMISProgramCode & " And Budget_AcctCharge=" & FMISAccountCode & " And YearOf=" & GetTransYear & ""), opndb, adOpenStatic, adLockOptimistic
    Else
        CRec.Open ("Select sum(Amount) as Amount From vwBMS_TotalControl_Account Where FmisProgramCode=" & FMISProgramCode & " And Budget_AcctCharge=" & FMISAccountCode & " And YearOf=" & GetTransYear & ""), opndb, adOpenStatic, adLockOptimistic
    End If
End If
If CRec.RecordCount <> 0 Then
    GetTotalControl_Account = IIf(IsNull(CRec!Amount), 0, CRec!Amount)
Else
    GetTotalControl_Account = 0
End If
CRec.Close
Set CRec = Nothing

End Function
'EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE

Private Function GetOfficeCode(strSelOffice) As Boolean
    If strSelOffice = 1 Or strSelOffice = 43 Then
        GetOfficeCode = True
    Else
        GetOfficeCode = False
    End If
End Function

'====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====
'+++ Function Created        : 08/03/2007
'+++ Created by              : EDGE
'+++ Purpose / Description   : This function gets the total amount controlled for the particular budget office account.
'+++ Functions / Subs Used   : None
'+++ Tables / Views Used     : vwBMS_TotalControlAmount_PerAccountPerMonth
'+++ Requirements / Inputs   : FMISProgramCode, FMISAccountCode, EndMonth
'+++ Output                  : Total Control per Account in currency
'+++ Limitations             : This can only be used for budget accounts, not for planning accounts.
'====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====
Public Function GetTotalPeriodicControl_Account(ByVal FMISProgramCode As Long, ByVal FMISAccountCode As Long, ByVal EndMonth As Long) As Currency
Dim CRec As New ADODB.Recordset
Dim FoundAlobs As New ADODB.Recordset
Dim Alob As String
Dim NewAmount As Currency

If FMISProgramCode <> 54 Then
    
    'If frmOfficeSelectSAAO.byAlobsno.Value = 1 Then
       CRec.Open ("Select Sum(Amount) as Amount From vwBMS_TotalControlAmount_PerAccountPerMonth Where Budget_AcctCharge = " & FMISAccountCode & " and FMISProgramCode = " & FMISProgramCode & " And YearOf = " & GetTransYear & " And MonthOf<=" & EndMonth & ""), opndb, adOpenStatic, adLockOptimistic
    'Else 'query used for alobsno by date time entered
    '  CRec.Open ("Select Sum(Amount) as Amount From vwBMS_TotalControlAmount_PerAccountPerMonth_byDateTimeEntered Where Budget_AcctCharge = " & FMISAccountCode & " and FMISProgramCode = " & FMISProgramCode & " And YearOf = " & GetTransYear & " And MonthOf<=" & EndMonth & " And YearOf2 = " & GetTransYear & " "), opndb, adOpenStatic, adLockOptimistic
    'End If
    
    NewAmount = NewAmount + IIf(IsNull(CRec!Amount), 0, CRec!Amount)
    Alob = ""
    
'    If CheckAuditTrailCode(FMISProgramCode, FMISAccountCode, EndMonth) = 1 Then 'used for Audit Trail
'        CRec.Close
'        Set CRec = Nothing
'        CRec.Open ("Select * from [vwBMS_TotalControlAmount_PerAccountPerMonth_AuditTrail] where Budget_AcctCharge = " & FMISAccountCode & " and FMISProgramCode = " & FMISProgramCode & " and AuditTrail_TrnSeqNo is not null and Yearof=" & Right(GetTransYear, 2) & " and MonthOf <= " & EndMonth & " order by AuditTrail_TrnSeqNo ASC "), opndb, adOpenStatic, adLockOptimistic
'        If CRec.RecordCount <> 0 Then
'            While Not CRec.EOF
'                If CRec!alobsno <> Alob Then
'
'                    Alob = CRec!alobsno
'                    FoundAlobs.Open ("Select * from [vwBMS_TotalControlAmount_PerAccountPerMonth_AuditTrail] where AlobsNo='" & CRec!alobsno & "' and Budget_AcctCharge = " & FMISAccountCode & " and FMISProgramCode = " & FMISProgramCode & " and AuditTrail_TrnSeqNo is not null and Yearof=" & Right(GetTransYear, 2) & " and MonthOf <= " & EndMonth & " order by AuditTrail_TrnSeqNo ASC "), opndb, adOpenStatic, adLockOptimistic
'                    FoundAlobs.MoveLast
'                    NewAmount = NewAmount + FoundAlobs!Amount
'                    FoundAlobs.Close
'                    Set FoundAlobs = Nothing
'
'                End If
'
'                CRec.MoveNext
'
'            Wend
'        End If
'
'    End If
    
    If CRec.RecordCount <> 0 Then
        GetTotalPeriodicControl_Account = IIf(IsNull(NewAmount), 0, NewAmount)
    Else
        GetTotalPeriodicControl_Account = 0
    End If
    CRec.Close
    Set CRec = Nothing
Else
    
        If FMISAccountCode = 2625 Then  'aserbac
            CRec.Open ("Select Sum(Amount) as Amount From tblFMIS_Transaction Where FmisOfficeCode = 39 and Income=0 And cast('20' + substring(AlobsNo,10,2) as int) = " & GetTransYear & " And cast(substring(AlobsNo,13,2) as int)<=" & EndMonth & " And Actioncode=1 and AuditTrailCode=0"), opndb, adOpenStatic, adLockOptimistic
            If CRec.RecordCount <> 0 Then
                GetTotalPeriodicControl_Account = IIf(IsNull(CRec!Amount), 0, CRec!Amount)
            Else
                GetTotalPeriodicControl_Account = 0
            End If
            CRec.Close
            Set CRec = Nothing
        ElseIf FMISAccountCode = 2623 Then  'pnb
            CRec.Open ("Select Sum(Amount) as Amount From tblFMIS_Transaction Where FmisOfficeCode = 41 and Income=0 And cast('20' + substring(AlobsNo,10,2) as int) = " & GetTransYear & " And cast(substring(AlobsNo,13,2) as int)<=" & EndMonth & " And Actioncode=1 and AuditTrailCode=0"), opndb, adOpenStatic, adLockOptimistic
            If CRec.RecordCount <> 0 Then
                GetTotalPeriodicControl_Account = IIf(IsNull(CRec!Amount), 0, CRec!Amount)
            Else
                GetTotalPeriodicControl_Account = 0
            End If
            CRec.Close
            Set CRec = Nothing
        ElseIf FMISAccountCode = 2622 Then  'ptc
            CRec.Open ("Select Sum(Amount) as Amount From tblFMIS_Transaction Where FmisOfficeCode = 37 and Income=0 And cast('20' + substring(AlobsNo,10,2) as int) = " & GetTransYear & " And cast(substring(AlobsNo,13,2) as int)<=" & EndMonth & " And Actioncode=1 and AuditTrailCode=0"), opndb, adOpenStatic, adLockOptimistic
            If CRec.RecordCount <> 0 Then
                GetTotalPeriodicControl_Account = IIf(IsNull(CRec!Amount), 0, CRec!Amount)
            Else
                GetTotalPeriodicControl_Account = 0
            End If
            CRec.Close
            Set CRec = Nothing
        ElseIf FMISAccountCode = 2624 Then  'water
            CRec.Open ("Select Sum(Amount) as Amount From tblFMIS_Transaction Where FmisOfficeCode = 38 and Income=0 And cast('20' + substring(AlobsNo,10,2) as int) = " & GetTransYear & " And cast(substring(AlobsNo,13,2) as int)<=" & EndMonth & " And Actioncode=1 and AuditTrailCode=0"), opndb, adOpenStatic, adLockOptimistic
            If CRec.RecordCount <> 0 Then
                GetTotalPeriodicControl_Account = IIf(IsNull(CRec!Amount), 0, CRec!Amount)
            Else
                GetTotalPeriodicControl_Account = 0
            End If
            CRec.Close
            Set CRec = Nothing
        ElseIf FMISAccountCode = 2626 Then  'dxda
            CRec.Open ("Select Sum(Amount) as Amount From tblFMIS_Transaction Where FmisOfficeCode = 40 and Income=0 And cast('20' + substring(AlobsNo,10,2) as int) = " & GetTransYear & " And cast(substring(AlobsNo,13,2) as int)<=" & EndMonth & " And Actioncode=1 and AuditTrailCode=0"), opndb, adOpenStatic, adLockOptimistic
            If CRec.RecordCount <> 0 Then
                GetTotalPeriodicControl_Account = IIf(IsNull(CRec!Amount), 0, CRec!Amount)
            Else
                GetTotalPeriodicControl_Account = 0
            End If
            CRec.Close
            Set CRec = Nothing
        Else
            GetTotalPeriodicControl_Account = 0
        End If
    
    If CheckAuditTrailCode(FMISProgramCode, FMISAccountCode, EndMonth) = 1 Then ''''''' this part of the code is used for audit trail...  ''''''''''''''
        CRec.Close
        Set CRec = Nothing
        If FMISAccountCode = 2625 Then  'aserbac
            CRec.Open ("Select * from [vwBMS_TotalControlAmount_PerAccountPerMonth_AuditTrail] where AuditTrail_TrnSeqNo is not null and MonthOf <= " & EndMonth & " order by AuditTrail_TrnSeqNo "), opndb, adOpenStatic, adLockOptimistic
            If CRec.RecordCount <> 0 Then
                CRec.MoveLast
                GetTotalPeriodicControl_Account = GetTotalPeriodicControl_Account + IIf(IsNull(CRec!Amount), 0, CRec!Amount)
            End If
            CRec.Close
            Set CRec = Nothing
        ElseIf FMISAccountCode = 2623 Then  'pnb
            CRec.Open ("Select * from [vwBMS_TotalControlAmount_PerAccountPerMonth_AuditTrail] where AuditTrail_TrnSeqNo is not null and MonthOf <= " & EndMonth & " order by AuditTrail_TrnSeqNo "), opndb, adOpenStatic, adLockOptimistic
            If CRec.RecordCount <> 0 Then
               CRec.MoveLast
                GetTotalPeriodicControl_Account = GetTotalPeriodicControl_Account + IIf(IsNull(CRec!Amount), 0, CRec!Amount)
            End If
            CRec.Close
            Set CRec = Nothing
        ElseIf FMISAccountCode = 2622 Then  'ptc
            CRec.Open ("Select * from [vwBMS_TotalControlAmount_PerAccountPerMonth_AuditTrail] where AuditTrail_TrnSeqNo is not null and MonthOf <= " & EndMonth & " order by AuditTrail_TrnSeqNo "), opndb, adOpenStatic, adLockOptimistic
            If CRec.RecordCount <> 0 Then
               CRec.MoveLast
                GetTotalPeriodicControl_Account = GetTotalPeriodicControl_Account + IIf(IsNull(CRec!Amount), 0, CRec!Amount)
            End If
            CRec.Close
            Set CRec = Nothing
        ElseIf FMISAccountCode = 2624 Then  'water
            CRec.Open ("Select * from [vwBMS_TotalControlAmount_PerAccountPerMonth_AuditTrail] where AuditTrail_TrnSeqNo is not null and MonthOf <= " & EndMonth & " order by AuditTrail_TrnSeqNo "), opndb, adOpenStatic, adLockOptimistic
            If CRec.RecordCount <> 0 Then
               CRec.MoveLast
                GetTotalPeriodicControl_Account = GetTotalPeriodicControl_Account + IIf(IsNull(CRec!Amount), 0, CRec!Amount)
            End If
            CRec.Close
            Set CRec = Nothing
        ElseIf FMISAccountCode = 2626 Then  'dxda
            CRec.Open ("Select * from [vwBMS_TotalControlAmount_PerAccountPerMonth_AuditTrail] where AuditTrail_TrnSeqNo is not null and MonthOf <= " & EndMonth & " order by AuditTrail_TrnSeqNo "), opndb, adOpenStatic, adLockOptimistic
            If CRec.RecordCount <> 0 Then
               CRec.MoveLast
                GetTotalPeriodicControl_Account = GetTotalPeriodicControl_Account + IIf(IsNull(CRec!Amount), 0, CRec!Amount)
            End If
            CRec.Close
            Set CRec = Nothing
        Else
            GetTotalPeriodicControl_Account = 0
        End If
    End If
End If

End Function
'EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE

Public Function CheckAuditTrailCode(FProg, FAcct, EndMonth) As Integer
Dim OpnAuditTrail As New ADODB.Recordset
    OpnAuditTrail.Open "Select * from dbo.tblBMS_SubsidiaryLedger where fmisprogramcode=" & FProg & " and budget_acctcharge=" & FAcct & " and MONTH(AuditTrailCode_DateTime) <=" & EndMonth & " and  YEAR(AuditTrailCode_DateTime)=" & GetTransYear & " and audittrailcode > 0 ", opndb, adOpenStatic, adLockOptimistic
    If OpnAuditTrail.RecordCount <> 0 Then
        CheckAuditTrailCode = 1
    Else
        CheckAuditTrailCode = 0
    End If
End Function

'====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====
'+++ Function Created        : 05/29/2006
'+++ Created by              : EDGE
'+++ Purpose / Description   : This function gets the Binded Office Code, Program Code and Account Code of the given
'+++                            Office Code.
'+++ Functions / Subs Used   : None
'+++ Tables / Views Used     : tblBMS_Binding
'+++ Requirements / Inputs   : FMISOfficeCode of the original office, a variable for the binded FMISOfficeCode,
'+++                            a variable for the binded FMISProgramCode, a varialble for the binded FMISAccountCode
'+++ Output                  : Binded FMISOfficeCode, binded FMISProgramCode, binded FMISAccountCode
'+++ Limitations             : This can only be used for budget accounts, not for planning accounts.
'====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====
Public Sub GetBindedOfficeProgramAccount(ByVal FMISOfficeCode As Long, BindOfficeCode As Long, BindProgramCode As Long, BindAccountCode As Long)
Dim getRec As New ADODB.Recordset

getRec.Open ("Select * From tblBMS_Binding Where OfficeCode=" & FMISOfficeCode & " And ActionCode=1 And YearOf=" & GetTransYear & ""), opndb, adOpenStatic, adLockOptimistic
If getRec.RecordCount <> 0 Then
    BindOfficeCode = getRec!OfficeCodeBind
    BindProgramCode = getRec!ProgramCodeBind
    BindAccountCode = getRec!AccountCodeBind
Else
    BindOfficeCode = 0
    BindProgramCode = 0
    BindAccountCode = 0
End If
getRec.Close
Set getRec = Nothing

End Sub
'EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE

'====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====
'+++ Function Created        : 06/16/2006
'+++ Created by              : EDGE
'+++ Purpose / Description   : This sub routine logs the activity of the user in the tblBMS_Log
'+++ Functions / Subs Used   : GetCompName
'+++ Tables / Views Used     : tblBMS_Log
'+++ Requirements / Inputs   : current Form Caption, affected Table Name, User ID, and the activity of the user or the SQL statement.
'+++ Output                  : None
'+++ Limitations             : None
'====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====
Public Sub LogActivity(ByVal FormCaption As String, ByVal TableName As String, ByVal UserID As String, ByVal ActivitySQL As String)
    'opndb.Execute "Insert into tblBMS_Log (UserID,[Transaction],tblname,FormName,ComputerName,datetimeentered) Values ('" & UserID & "','" & ConsiderApostrophe(ActivitySQL) & "','" & TableName & "','" & ConsiderApostrophe(FormCaption) & "','" & GetCompName & "','" & ServerDate & "')"
End Sub
'EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE

'====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====
'+++ Function Created        : 06/16/2006
'+++ Created by              : EDGE
'+++ Purpose / Description   : This function gets the computer name
'+++ Functions / Subs Used   : GetComputerName
'+++ Tables / Views Used     : None
'+++ Requirements / Inputs   : None
'+++ Output                  : Computer Name
'+++ Limitations             : None
'====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====
Public Function GetCompName() As String
Dim compname As String, retval As Long  ' string to use as buffer & return value

compname = Space(255)  ' set a large enough buffer for the computer name
retval = GetComputerName(compname, 255)  ' get the computer's name
' Remove the trailing null character from the strong
compname = Mid(compname, 1, InStr(compname, vbNullChar) - 1)

GetCompName = compname

End Function
'EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE

'====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====
'+++ Function Created        : 06/16/2006
'+++ Created by              : EDGE
'+++ Purpose / Description   : This function formats the apostrophe in every word to be acceptable on tthe SQL Server
'+++ Functions / Subs Used   : None
'+++ Tables / Views Used     : None
'+++ Requirements / Inputs   : String with apostrophe
'+++ Output                  : String with formatted apostrophe
'+++ Limitations             : None
'====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====
Public Function ConsiderApostrophe(ByVal Sentence As String) As String
Dim NewSentence As String
Dim x As Integer

x = 1

Do Until InStr(x, Sentence, "'") = 0
    x = InStr(x, Sentence, "'")
    Sentence = Mid(Sentence, 1, x) & "'" & Mid(Sentence, x + 1)
    x = x + 2
Loop

ConsiderApostrophe = Sentence

End Function
'EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE

'====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====
'+++ Function Created        : 06/16/2006
'+++ Created by              : EDGE
'+++ Purpose / Description   : This function retrieves the date of the server
'+++ Functions / Subs Used   : None
'+++ Tables / Views Used     : None
'+++ Requirements / Inputs   : None
'+++ Output                  : None
'+++ Limitations             : None
'====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====
Public Function ServerDate() As Date
Dim DateRec As New ADODB.Recordset

DateRec.Open ("Select getdate() as serverdate"), opndb, adOpenStatic, adLockOptimistic
ServerDate = DateRec!ServerDate
DateRec.Close
Set DateRec = Nothing

End Function
'EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE

'====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====
'+++ Function Created        : 07/21/2006
'+++ Created by              : EDGE
'+++ Purpose / Description   : This function retrieves office medium name
'+++ Functions / Subs Used   : None
'+++ Tables / Views Used     : tblREF_AIS_Offices
'+++ Requirements / Inputs   : office id
'+++ Output                  : office medium name
'+++ Limitations             : None
'====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====
Public Function GetOfficeMedium(ByVal OfficeID As Long) As String
Dim xREc As New ADODB.Recordset

xREc.Open ("Select * From tblREF_AIS_Offices Where FMISOfficeID=" & OfficeID & ""), opndb, adOpenStatic, adLockOptimistic
If xREc.RecordCount <> 0 Then
    GetOfficeMedium = xREc!OfficeMedium
Else
    GetOfficeMedium = ""
End If
xREc.Close
Set xREc = Nothing

End Function
'EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE

'====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====
'+++ Function Created        : 07/25/2006
'+++ Created by              : EDGE
'+++ Purpose / Description   : This function returns the re-alligned appropriation of the account
'+++ Functions / Subs Used   : None
'+++ Tables / Views Used     : tblBMS_Reallignment
'+++ Requirements / Inputs   : FMISProgramCode, FMISAccountCode
'+++ Output                  : Re-alligned Appropriation of the account in currency
'+++ Limitations             : This can only be used for budget accounts, not for planning accounts.
'====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====
Public Function GetReAllignedAmountTo(ByVal ToProgCode As Long, ByVal ToAcctCode As Long) As Currency
Dim xREc As New ADODB.Recordset

xREc.Open ("Select sum(Amount) as Amount From tblBMS_Reallignment Where ToProgCode=" & ToProgCode & " And ToAccountCode=" & ToAcctCode & " And ActionCode=1 And YearOf=" & GetTransYear & ""), opndb, adOpenStatic, adLockOptimistic
If xREc.RecordCount <> 0 Then
    GetReAllignedAmountTo = IIf(IsNull(xREc!Amount), 0, xREc!Amount)
Else
    GetReAllignedAmountTo = 0
End If
xREc.Close
Set xREc = Nothing

End Function
'EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE


'====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====
'+++ Function Created        : 07/25/2006
'+++ Created by              : EDGE
'+++ Purpose / Description   : This function returns the total supplemental appropriation of the account
'+++ Functions / Subs Used   : None
'+++ Tables / Views Used     : vwBMS_TotalSupplemental_Account
'+++ Requirements / Inputs   : FMISProgramCode, FMISAccountCode
'+++ Output                  : total supplemental appropriation of the account in currency
'+++ Limitations             : This can only be used for budget accounts, not for planning accounts.
'====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====
Public Function GetTotalSupplementalPerAccount(ByVal FMISProgramCode As Long, FMISAccountCode As Long) As Currency
Dim xREc As New ADODB.Recordset

xREc.Open ("Select * From vwBMS_TotalSupplemental_Account Where FMISProgramCode=" & FMISProgramCode & " And FMISAccountCode=" & FMISAccountCode & " And YearOf=" & GetTransYear & ""), opndb, adOpenStatic, adLockOptimistic
If xREc.RecordCount <> 0 Then
    GetTotalSupplementalPerAccount = xREc!Amount
Else
    GetTotalSupplementalPerAccount = 0
End If
xREc.Close
Set xREc = Nothing

End Function
'EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE

'====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====
'+++ Function Created        : 08/3/2007
'+++ Created by              : EDGE
'+++ Purpose / Description   : This function returns the total supplemental appropriation of the account
'+++ Functions / Subs Used   : None
'+++ Tables / Views Used     : tblBMS_SupplementalBudget
'+++ Requirements / Inputs   : FMISProgramCode, FMISAccountCode, EndMonth
'+++ Output                  : total supplemental appropriation of the account in currency for the given period
'+++ Limitations             : This can only be used for budget accounts, not for planning accounts.
'====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====
Public Function GetTotalPeriodicSupplementalPerAccount(ByVal FMISProgramCode As Long, FMISAccountCode As Long, ByVal EndMonth As Long) As Currency
Dim xREc As New ADODB.Recordset

xREc.Open ("Select sum(Amount) as Amount From tblBMS_SupplementalBudget Where FMISProgramCode=" & FMISProgramCode & " And FMISAccountCode=" & FMISAccountCode & " And YearOf=" & GetTransYear & " And MonthOf<=" & EndMonth & " And ActionCode=1"), opndb, adOpenStatic, adLockOptimistic
If xREc.RecordCount <> 0 Then
    GetTotalPeriodicSupplementalPerAccount = IIf(IsNull(xREc!Amount), 0, xREc!Amount)
Else
    GetTotalPeriodicSupplementalPerAccount = 0
End If
xREc.Close
Set xREc = Nothing

End Function
'EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE


'====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====
'+++ Function Created        : 08/10/2006
'+++ Created by              : EDGE
'+++ Purpose / Description   : This function returns the deducted re-alligned appropriation of the account
'+++ Functions / Subs Used   : None
'+++ Tables / Views Used     : vwBMS_TotalDeductedRealignment
'+++ Requirements / Inputs   : FMISProgramCode, FMISAccountCode
'+++ Output                  : Deductged Re-alligned Appropriation of the account in currency
'+++ Limitations             : This can only be used for budget accounts, not for planning accounts.
'====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====
Public Function GetReAlignDeduction(ByVal FMISProgramCode As Long, ByVal FMISAccountCode As Long) As Currency
Dim DRec As New ADODB.Recordset

DRec.Open ("Select * From vwBMS_TotalDeductedRealignment Where FromProgCode=" & FMISProgramCode & " And FromAccountCode=" & FMISAccountCode & " And YearOf=" & GetTransYear & ""), opndb, adOpenStatic, adLockOptimistic
If DRec.RecordCount <> 0 Then
    GetReAlignDeduction = DRec!Amount
Else
    GetReAlignDeduction = 0
End If
DRec.Close
Set DRec = Nothing

End Function
'RYGN===='RYGN===='RYGN===='RYGN===='RYGN===='RYGN===='RYGN===='RYGN===='RYGN===='RYGN===='RYGN===='RYGN===='RYGN===='RYGN====
Public Function GetReversedFrom(ByVal FMISProgramCode As Long, ByVal FMISAccountCode As Long, ByVal OOECode As Long) As Currency
    Dim Rev As New ADODB.Recordset
    
    Rev.Open ("Select sum(amount)as Amount From vwBMS_ReversionDeductedAmount Where FromProgCode=" & FMISProgramCode & " And FromAccountCode=" & FMISAccountCode & " and FromOOECode=" & OOECode & "And YearOf=" & GetTransYear & " group by [FromProgCode],[FromAccountCode],YearOf"), opndb, adOpenStatic, adLockOptimistic
    If Rev.RecordCount <> 0 Then
        GetReversedFrom = Rev!Amount
    Else
        GetReversedFrom = 0
    End If
    Rev.Close
    Set Rev = Nothing
End Function

'RYGN===='RYGN===='RYGN===='RYGN===='RYGN===='RYGN===='RYGN===='RYGN===='RYGN===='RYGN===='RYGN===='RYGN===='RYGN===='RYGN====
Public Function GetReversedTo(ByVal FMISProgramCode As Long, ByVal FMISAccountCode As Long, ByVal OOE As Long) As Currency
    Dim Rev As New ADODB.Recordset
    
    Rev.Open ("Select sum(amount)as Amount From [vwBMS_ReversionTotalAmount] Where ToProgCode=" & FMISProgramCode & " And ToAccountCode=" & FMISAccountCode & " And ToOOECode=" & OOE & " And YearOf=" & GetTransYear & " group by [ToProgCode],[ToAccountCode],YearOf "), opndb, adOpenStatic, adLockOptimistic
    If Rev.RecordCount <> 0 Then
        GetReversedTo = Rev!Amount
    Else
        GetReversedTo = 0
    End If
    Rev.Close
    Set Rev = Nothing
End Function

'EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE

'====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====
'+++ Function Created        : 08/10/2006
'+++ Created by              : EDGE
'+++ Purpose / Description   : This function returns the total released subsidy of the office
'+++ Functions / Subs Used   : None
'+++ Tables / Views Used     : vwBMS_TotalSubsidyRelease
'+++ Requirements / Inputs   : FMISOfficeCode of the economic enterprise
'+++ Output                  : Total subsidy release of the office in currency
'+++ Limitations             : This can only be used for budget accounts, not for planning accounts.
'====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====
Public Function GetTotalSubsidyRelease(ByVal FMISOfficeCode As Long) As Currency
Dim SREc As New ADODB.Recordset

SREc.Open ("Select * From vwBMS_TotalSubsidyRelease Where FMISOfficeCode=" & FMISOfficeCode & " And YearOf=" & GetTransYear & ""), opndb, adOpenStatic, adLockOptimistic
If SREc.RecordCount <> 0 Then
    GetTotalSubsidyRelease = SREc!Amount
Else
    GetTotalSubsidyRelease = 0
End If
SREc.Close
Set SREc = Nothing

End Function
'EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE

'====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====
'+++ Function Created        : 08/19/2006
'+++ Created by              : EDGE
'+++ Purpose / Description   : This function returns the total released subsidy of the office
'+++ Functions / Subs Used   : None
'+++ Tables / Views Used     : vwBMS_TotalIncomeRelease
'+++ Requirements / Inputs   : FMISOfficeCode of the economic enterprise
'+++ Output                  : Total income release of the office in currency
'+++ Limitations             : This can only be used for budget accounts, not for planning accounts.
'====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====
Public Function GetTotalIncomeRelease(FMISOfficeCode As Long) As Currency
Dim IRec As New ADODB.Recordset

IRec.Open ("Select * From vwBMS_TotalIncomeRelease Where FMISOfficeCode=" & FMISOfficeCode & " And YearOf=" & GetTransYear & ""), opndb, adOpenStatic, adLockOptimistic
If IRec.RecordCount <> 0 Then
    GetTotalIncomeRelease = IRec!Amount
Else
    GetTotalIncomeRelease = 0
End If
IRec.Close
Set IRec = Nothing

End Function
'EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE

'====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====
'+++ Function Created        : 01/02/2007
'+++ Created by              : EDGE
'+++ Purpose / Description   : This function returns the Transaction Year
'+++ Functions / Subs Used   : None
'+++ Tables / Views Used     : tblBMS_TransYear
'+++ Requirements / Inputs   : None
'+++ Output                  : Transaction Year
'+++ Limitations             : This can only be used for budget accounts, not for planning accounts.
'====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====
Public Function GetLastTransYear() As Long
Dim TYear As New ADODB.Recordset

TYear.Open ("Select * From tblBMS_TransYear"), opndb, adOpenStatic, adLockOptimistic
If TYear.RecordCount <> 0 Then
    TYear.MoveLast
    GetLastTransYear = TYear!trnYear
End If
TYear.Close
Set TYear = Nothing

End Function
'EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE

'====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====
'+++ Function Created        : 08/18/2006
'+++ Created by              : EDGE
'+++ Purpose / Description   : This function returns the reserve amount of the account
'+++ Functions / Subs Used   : None
'+++ Tables / Views Used     : vwBMS_ReserveAmount_Account
'+++ Requirements / Inputs   : FMISProgramCode, FMISAccountCode
'+++ Output                  : reserve amount of the account in currency
'+++ Limitations             : This can only be used for budget accounts, not for planning accounts.
'====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====
Public Function GetReserveAmount(ByVal FMISProgramCode As Long, ByVal FMISAccountCode As Long) As Currency
Dim ResRec As New ADODB.Recordset

ResRec.Open ("Select * From vwBMS_ReserveAmount_Account Where FMISProgramCode=" & FMISProgramCode & " And FMISAccountCode=" & FMISAccountCode & " And YearOf=" & GetTransYear & ""), opndb, adOpenStatic, adLockOptimistic
If ResRec.RecordCount <> 0 Then
    GetReserveAmount = ResRec!ReserveAmount
Else
    GetReserveAmount = 0
End If
ResRec.Close
Set ResRec = Nothing

End Function
'EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE====EDGE

'RYGN==>>RYGN==>>RYGN==>>RYGN==>>RYGN==>>RYGN==>>RYGN==>>RYGN==>>RYGN==>>RYGN==>>RYGN==>>RYGN==>>RYGN==>>RYGN==>>RYGN==>>RYGN==>>RYGN==>>RYGN==>>RYGN==>RYGN==>>RYGN==>>

Public Function getTempAlobsNo(ByVal yr As Long, ByVal offID As Integer) As String

Dim gno As New ADODB.Recordset
Dim gno1 As New ADODB.Recordset
Dim exno As New ADODB.Recordset
Dim nos As Integer
    
    nos = 0
    gno.Open ("Select top 1 * from [tblFMIS_Transaction] where substring(Alobsno,1,2)='" & Right(yr, 2) & "' and substring(alobsno,3,2)='" & Format(Month(ServerDate), "0#") & "' and PGOActionCode = 1 or PGOActionCode = 4     order by cast(AlobSeqNo as integer) desc"), opndb, adOpenStatic, adLockOptimistic
    If gno.RecordCount = 0 Then
        nos = Right(GetTransYear, 2) & Format(Month(ServerDate), "0#") & "1"
    Else
        nos = gno!AlobSeqNo
    End If
    
    exno.Open ("Select TOP 1 * from [tblBMS_ExcessControl] where substring(Alobsno,1,2)='" & Right(yr, 2) & "' and substring(alobsno,3,2)='" & Format(Month(ServerDate), "0#") & "' and PGOActionCode = 1 or PGOActionCode = 4  order by trnno desc"), opndb, adOpenStatic, adLockOptimistic
    If exno.RecordCount = 0 Then
        getTempAlobsNo = nos
    Else
        nos = nos + exno.RecordCount
    End If
    
    exno.Close
    Set exno = Nothing
    
    While Not gno.EOF
        
        getTempAlobsNo = Right(GetTransYear, 2) & Format(Month(ServerDate), "0#") & IIf(offID = 1, 1, 2) & nos + 1
        gno.Close
        Set gno = Nothing
        Exit Function
    Wend
    
End Function
'RYGN==>>RYGN==>>RYGN==>>RYGN==>>RYGN==>>RYGN==>>RYGN==>>RYGN==>>RYGN==>>RYGN==>>RYGN==>>RYGN==>>RYGN==>>RYGN==>>RYGN==>>RYGN==>>RYGN==>>RYGN==>>RYGN==>RYGN==>>RYGN==>>

'RYGN==>>RYGN==>>RYGN==>>RYGN==>>RYGN==>>RYGN==>>RYGN==>>RYGN==>>RYGN==>>RYGN==>>RYGN==>>RYGN==>>RYGN==>>RYGN==>>RYGN==>>RYGN==>>RYGN==>>RYGN==>>RYGN==>RYGN==>>RYGN==>>

Public Function chkAlobsNo(ByVal alobs As String) As Integer
    Rec.Open "Select * from [dbo].[tblBMS_ExcessControl] where alobsno='" & alobs & "' and actioncode=1", opndb, adOpenStatic, adLockOptimistic
    If Rec.RecordCount <> 0 Then
        'chkAlobsNo = Rec!PGOActionCode
        If Rec!PGOActionCode <> 0 Then
            chkAlobsNo = Rec!PGOActionCode
        Else
            chkAlobsNo = 0
        End If
    Else
        chkAlobsNo = 0
    End If
    Rec.Close
    Set Rec = Nothing
End Function

Public Function CheckSpecificUser(ByVal uid As String) As String
    uidParameter = ""
    If CheckUserControl(uid) = True Then
        uidParameter = " and [PGOUserID] = " & uid & ""
    Else
        uidParameter = " and [BudgetUserID] = " & uid & ""
    End If
End Function

Public Function CheckUserControl(ByVal uid As String) As Boolean
    If uid = "1237" Or uid = "1735" Then
        CheckUserControl = True
    Else
        CheckUserControl = False
    End If
End Function

Public Function getAccClaimantName(ByVal alobs As String) As String
    Rec.Open "Select *  FROM [fmis].[dbo].[tblAMIS_IncomingDVTrns] where ObrNo='" & alobs & "' and  Actioncode =1", opndb, adOpenStatic, adLockOptimistic
    If Rec.RecordCount > 0 Then
        getAccClaimantName = getDetails(Rec!ClaimantCode)
    End If
    Rec.Close
    Set Rec = Nothing
End Function

Public Function getDetails(ByVal cc As String) As String
    opnRec.Open "Select * FROM [fmis].[dbo].[MPfunc_Claimant] () where id='" & cc & "'", opndb, adOpenStatic, adLockOptimistic
    If opnRec.RecordCount > 0 Then
        getDetails = opnRec!Name
    End If
    opnRec.Close
    Set opnRec = Nothing
End Function
'RYGN==>>RYGN==>>RYGN==>>RYGN==>>RYGN==>>RYGN==>>RYGN==>>RYGN==>>RYGN==>>RYGN==>>RYGN==>>RYGN==>>RYGN==>>RYGN==>>RYGN==>>RYGN==>>RYGN==>>RYGN==>>RYGN==>RYGN==>>RYGN==>>


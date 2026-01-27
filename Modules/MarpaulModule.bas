Attribute VB_Name = "MarpaulModule"
'Option Explicit
Public MPDll As New clsBlowfish
Public ErrDll As New errolDLL
Public strReportName As String
Public IsCopy As Boolean
Public Autoclose As Boolean
Public DefaultPost As Date
Public admin As Integer
Public SystemAdmin As Integer
Public SystemVersion As String
Public SystemDescription As String
Public IsOktoClear  As Boolean
Public CM, cc As String

'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'+++++ Purpose / Description    : Dummy the Journal Entry into Final Journal Entry
'+++++ Input                    : Saving JEV Details in one table
'+++++ Return                   :
'+++++ Date Created             : March 16, 2011
'+++++ Programmer               : Mar Paul M. Ajero
'+++++ UPDATES +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Public Criteria As String
      Public Const SWP_NOMOVE = 2
      Public Const SWP_NOSIZE = 1
      Public Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
      Public Const HWND_TOPMOST = -1
      Public Const HWND_NOTOPMOST = -2

      Declare Function SetWindowPos Lib "user32" _
            (ByVal hwnd As Long, _
            ByVal hWndInsertAfter As Long, _
            ByVal x As Long, _
            ByVal y As Long, _
            ByVal cx As Long, _
            ByVal cy As Long, _
            ByVal wFlags As Long) As Long
Public Function FindField(ByVal field As String, ByVal Tblname As String, ByVal FieldCondition As String, ByVal Condition As String, ByVal Lcondition As String) As String
Dim Frec As New ADODB.Recordset
Frec.Open (" Select top 1 " & field & " as field from " & Tblname & " where " & FieldCondition & " = '" & Condition & "' and " & Lcondition & " "), opndbaseFMIS, adOpenStatic, adLockOptimistic
If Frec.RecordCount <> 0 Then
    FindField = IIf(IsNull(Frec.Fields!field), "", Frec.Fields!field)
End If
Frec.Close
Set Frec = Nothing
End Function
Public Function CheckIfExistInFinalJEV(ByVal jevno As String) As Boolean
Dim rec As New ADODB.Recordset
CheckIfExistInFinalJEV = False
rec.Open "Select Jevno from tblAMIS_FinalJEV where jevno = '" & jevno & "' and actioncode = 1", opndbaseFMIS, adOpenStatic, adLockOptimistic
If rec.RecordCount > 0 Then
    CheckIfExistInFinalJEV = True
End If
rec.Close
Set rec = Nothing
End Function
Public Function CheckIfExistPTVinFinalJEV(ByVal ptvNo As String) As Boolean
Dim rec As New ADODB.Recordset
CheckIfExistPTVinFinalJEV = False
rec.Open "Select trnno from tblAMIS_FinalJEV where jevno = '" & ptvNo & "' and actioncode = 1", opndbaseFMIS, adOpenStatic, adLockOptimistic
If rec.RecordCount > 0 Then
    CheckIfExistPTVinFinalJEV = True
End If
rec.Close
Set rec = Nothing
End Function
Public Function Saved2FinalJEV(ByVal Date_ As Date, ByVal RCI As String, ByVal checkno As String, ByVal Particular As String, ByVal jevno As String, _
ByVal ClaimantCode As String, ByVal FmisAccountcode As String, ByVal Gamount As Currency, ByVal Debit As Currency _
, ByVal Credit As Currency, ByVal Transtype As Integer, ByVal FmisVoucherno As String, ByVal dvno As String, ByVal obrno As String, ByVal FundType As String, _
ByVal RCenter As String, ByVal OOE As String, ByVal RDOno As String, ByVal RefNo As String, ByVal Jevseries As Long, ByVal jevdate As Date, ByVal ptvNo As String)
opndbaseFMIS.Execute ("Insert Into tblAMIS_FinalJEV (Date_, RCI, checkno, Particular, JEVNo, ClaimantCode, FmisAccountcode, Gamount, Debit, Credit, Transtype, FmisVoucherno, DVNo, Obrno, FundType, RCenter, OOE, RDOno, RefNo,datetimeentered,jevseriesno,jevdate,ptvno,actioncode,isnew,jevby) " & _
"VALUES ('" & Date_ & "', '" & RCI & "','" & checkno & "','" & Replace(Particular, "'", "''") & "','" & jevno & "','" & ClaimantCode & "','" & FmisAccountcode & "','" & Gamount & "','" & Debit & "','" & Credit & "'," & Transtype & ",'" & FmisVoucherno & "'" & _
",'" & dvno & "','" & obrno & "','" & FundType & "','" & RCenter & "','" & OOE & "','" & RDOno & "','" & RefNo & "','" & Format(Now, "MM/dd/yyyy") & "'," & Jevseries & ",'" & jevdate & "','" & ptvNo & "',1,1,'" & Trim(ActiveUserID) & "')")

opndbaseFMIS.Execute ("update [fmis].[dbo].[tblAMIS_FinalJEV] set y = YEAR(jevdate) ,m = MONTH(jevdate) where  year(Jevdate) >= 2012 and y is null and m is null")
End Function
Public Function Saved2FinalJEV_forFinalJEV(ByVal Date_ As Date, ByVal RCI As String, ByVal checkno As String, ByVal Particular As String, ByVal jevno As String, _
ByVal ClaimantCode As String, ByVal FmisAccountcode As String, ByVal Gamount As Currency, ByVal Debit As Currency _
, ByVal Credit As Currency, ByVal Transtype As Integer, ByVal FmisVoucherno As String, ByVal dvno As String, ByVal obrno As String, ByVal FundType As String, _
ByVal RCenter As String, ByVal OOE As String, ByVal RDOno As String, ByVal RefNo As String, ByVal Jevseries As Long, ByVal jevdate As Date, ByVal ptvNo As String, ByVal PClosinG As Integer)
opndbaseFMIS.Execute ("Insert Into tblAMIS_FinalJEV (Date_, RCI, checkno, Particular, JEVNo, ClaimantCode, FmisAccountcode, Gamount, Debit, Credit, Transtype, FmisVoucherno, DVNo, Obrno, FundType, RCenter, OOE, RDOno, RefNo,datetimeentered,jevseriesno,jevdate,ptvno,actioncode,isnew,jevby,Pclosing) " & _
"VALUES ('" & Date_ & "', '" & RCI & "','" & checkno & "','" & Replace(Particular, "'", "''") & "','" & jevno & "','" & ClaimantCode & "','" & FmisAccountcode & "','" & Gamount & "','" & Debit & "','" & Credit & "'," & Transtype & ",'" & FmisVoucherno & "'" & _
",'" & dvno & "','" & obrno & "','" & FundType & "','" & RCenter & "','" & OOE & "','" & RDOno & "','" & RefNo & "','" & Format(Now, "MM/dd/yyyy") & "'," & Jevseries & ",'" & jevdate & "','" & ptvNo & "',1,1,'" & Trim(ActiveUserID) & "'," & PClosinG & ")")

opndbaseFMIS.Execute ("update [fmis].[dbo].[tblAMIS_FinalJEV] set y = YEAR(jevdate) ,m = MONTH(jevdate) where  year(Jevdate) >= 2012 and y is null and m is null")
End Function
Public Function GEtCompleteJEVDetails(ByVal field As String, ByVal whatfield As String, ByVal Date_ As String, ByVal RCI As String, ByVal checkno As String, ByVal Particular As String, ByVal jevno As String, _
ByVal ClaimantCode As String, ByVal FmisAccountcode As String, ByVal Gamount As Currency, ByVal Debit As Currency _
, ByVal Credit As Currency, ByVal Transtype As Integer, ByVal FmisVoucherno As String, ByVal dvno As String, ByVal obrno As String, ByVal FundType As String, _
ByVal RCenter As String, ByVal OOE As String, ByVal RDOno As String, ByVal RefNo As String, ByVal Jevseries As Long, ByVal jevdate As Date, ByVal ptvNo As String)

opndbaseFMIS.Execute ("Execute Proc_GetTransDetails @whatField = '" & whatfield & "',@Field = '" & field & "',@Date_ ='" & Date_ & "',@RCI= '" & RCI & "',@Checkno = '" & checkno & "'" & _
    ",@Particular='" & Replace(Particular, "'", "''") & "',@JEVno='" & jevno & "',@Claimantcode='" & ClaimantCode & "',@Gamount = '" & Gamount & "',@Debit = '" & Debit & "',@Credit='" & Credit & "',@Transtype = '" & Transtype & "'" & _
    ",@FmisVoucherno='" & FmisVoucherno & "',@Dvno = '" & dvno & "',@Obrno = '" & obrno & "',@Fundtype  = '" & FundType & "',@RCenter ='" & RCenter & "',@DateTimeEntered = '" & Now & "'" & _
    ",@Rdono= '" & RDOno & "',@refno= '" & RefNo & "',@childAccountcode= '',@OOE = '" & OOE & "',@Jevdate = '" & jevdate & "',@actioncode =1,@Ptvno='" & ptvNo & "',@JevSeriesNo= '" & Jevseries & "',@JEVby= '" & Trim(ActiveUserID) & "'")
    
End Function
Public Function GEtCompleteJEVDetails_v1(ByVal field As String, ByVal whatfield As String, ByVal Date_ As String, ByVal RCI As String, ByVal checkno As String, ByVal Particular As String, ByVal jevno As String, _
ByVal ClaimantCode As String, ByVal FmisAccountcode As String, ByVal Gamount As Currency, ByVal Debit As Currency _
, ByVal Credit As Currency, ByVal Transtype As Integer, ByVal FmisVoucherno As String, ByVal dvno As String, ByVal obrno As String, ByVal FundType As String, _
ByVal RCenter As String, ByVal OOE As String, ByVal RDOno As String, ByVal RefNo As String, ByVal Jevseries As Long, ByVal jevdate As Date, ByVal ptvNo As String, ByVal Continuing As Integer, ByVal isAdjustment As Integer)

opndbaseFMIS.Execute ("Execute Proc_GetTransDetails_v1 @whatField = '" & whatfield & "',@Field = '" & field & "',@Date_ ='" & Date_ & "',@RCI= '" & RCI & "',@Checkno = '" & checkno & "'" & _
    ",@Particular='" & Replace(Particular, "'", "''") & "',@JEVno='" & jevno & "',@Claimantcode='" & ClaimantCode & "',@Gamount = '" & Gamount & "',@Debit = '" & Debit & "',@Credit='" & Credit & "',@Transtype = '" & Transtype & "'" & _
    ",@FmisVoucherno='" & FmisVoucherno & "',@Dvno = '" & dvno & "',@Obrno = '" & obrno & "',@Fundtype  = '" & FundType & "',@RCenter ='" & RCenter & "',@DateTimeEntered = '" & Now & "'" & _
    ",@Rdono= '" & RDOno & "',@refno= '" & RefNo & "',@childAccountcode= '',@OOE = '" & OOE & "',@Jevdate = '" & jevdate & "',@actioncode =1,@Ptvno='" & ptvNo & "',@JevSeriesNo= '" & Jevseries & "',@JEVby= '" & Trim(ActiveUserID) & "',@continuing = '" & Continuing & "',@IsAdjustment = " & isAdjustment & "")
End Function
Public Function loadChildAccountcode(ByVal FundType As String, ByVal cmb As ComboBox)
Dim crec As New ADODB.Recordset
Dim x As Integer
crec.Open "Proc_GetAccountcodebyFtype @fundtype = '" & FundType & "'", opndbaseFMIS, adOpenStatic, adLockOptimistic
cmb.Clear
If crec.RecordCount > 0 Then
    For x = 1 To crec.RecordCount
        cmb.AddItem (crec.Fields!childaccountcode)
        crec.MoveNext
    Next x
End If
crec.Close
Set crec = Nothing
End Function
Public Function loadAccountcodefromCOAM(ByVal cmb As ComboBox)
Dim crec As New ADODB.Recordset
Dim x As Integer
crec.Open "Select Accountcode from tblREF_AIS_ChartOfAccountsMother", opndbaseFMIS, adOpenStatic, adLockOptimistic
cmb.Clear
If crec.RecordCount > 0 Then
    For x = 1 To crec.RecordCount
        cmb.AddItem (crec.Fields!accountcode)
        crec.MoveNext
        DoEvents
    Next x
End If
crec.Close
Set crec = Nothing
End Function
Public Function loadChildAccountcodebyfundcode(ByVal FundType As String, ByVal cmb As ComboBox)
Dim crec As New ADODB.Recordset
Dim x As Long
crec.Open "Proc_GetAccountcodebyFcode @fundcode = '" & FundType & "',@condition = ''", opndbaseFMIS, adOpenStatic, adLockOptimistic
cmb.Clear
If crec.RecordCount > 0 Then
    For x = 1 To crec.RecordCount
        cmb.AddItem Trim((crec.Fields!childaccountcode))
        crec.MoveNext
    Next x
End If
crec.Close
Set crec = Nothing
End Function


Public Function loadAccountcode(ByVal FundType As String, ByVal cmb As ComboBox)
Dim crec As New ADODB.Recordset
Dim x As Integer
crec.Open "Proc_AccountcodebyFtype @fundtype = '" & FundType & "'", opndbaseFMIS, adOpenStatic, adLockOptimistic
cmb.Clear
If crec.RecordCount > 0 Then
    For x = 1 To crec.RecordCount
        cmb.AddItem (crec.Fields!accountcode)
        crec.MoveNext
    Next x
End If
crec.Close
Set crec = Nothing
End Function
Public Function GetSignatory(ByVal cmb As ComboBox, ByVal Whattype As String)
Dim rec As New ADODB.Recordset
Dim x As Integer
cmb.Clear
cmb.AddItem ""
rec.Open " Select * from tblReff_Signatory where typsig = '" & Whattype & "'", opndbaseFMIS, adOpenStatic, adLockOptimistic
    If rec.RecordCount > 0 Then
        For x = 1 To rec.RecordCount
            cmb.AddItem rec!FullName
            cmb.ItemData(cmb.NewIndex) = CInt(rec!id)
            rec.MoveNext
        Next x
    End If
rec.Close
End Function
Public Function Loadcmb(ByVal cmb As ComboBox, ByVal sql As String)
Dim rec As New ADODB.Recordset
Dim x As Integer
cmb.Clear
cmb.AddItem ""
rec.Open sql, opndbaseFMIS, adOpenStatic, adLockOptimistic
    If rec.RecordCount > 0 Then
        For x = 1 To rec.RecordCount
            cmb.AddItem Trim(rec!Field2)
            cmb.ItemData(cmb.NewIndex) = CInt(rec!Field1)
            rec.MoveNext
            DoEvents
        Next x
    End If
rec.Close
End Function
Public Function LogApprovedAndAudit(ByVal dvno As String, ByVal Whattype As String, ByVal who As String)
Dim rec As New ADODB.Recordset
Dim dtetype As String
Dim status As Integer

If Whattype = "Auditby" Then
    dtetype = "AuditByDTE"
    status = 3
ElseIf Whattype = "preAudit" Then
    dtetype = "preAuditDTE"
    status = 26
ElseIf Whattype = "Approvedby" Then
    dtetype = "ApprovedbyDTE"
    status = 4
ElseIf Whattype = "UserID" Then
    dtetype = "datetimeentered"
    status = 1
End If

rec.Open "Select * from tblAMIS_LogApprovedAndAudit where dvno = '" & dvno & "'", opndbaseFMIS, adOpenStatic, adLockOptimistic
If rec.RecordCount > 0 Then
    opndbaseFMIS.Execute "Update tblAMIS_LogApprovedAndAudit set " & Whattype & " = '" & who & "'," & dtetype & " = getdate(),TStatus = " & status & " where dvno = '" & dvno & "'"
Else
    opndbaseFMIS.Execute "INSERT INTO tblAMIS_LogApprovedAndAudit (" & Whattype & ",dvno," & dtetype & ",TStatus) VALUES ('" & who & "','" & dvno & "',getdate()," & status & ")"
End If
rec.Close
End Function
Public Function IfHaveSubname(ByVal accountcode As Long) As Boolean
Dim rec As New ADODB.Recordset
IfHaveSubname = False
rec.Open "SELECT Subcode1,max([lvl]) as MaxLvl FROM [fmis].[dbo].[tblReff_CodeClassification] where Subcode1 = " & accountcode & " group by Subcode1", opndbaseFMIS, adOpenStatic, adLockOptimistic
    If rec.RecordCount <> 0 Then
        If rec!maxlvl > 0 Then
            IfHaveSubname = True
        End If
    End If
rec.Close
End Function
Public Function GetAccountNames(ByVal lst As ListView, ByVal TableName As String, ByVal field As String, ByVal Field2 As String, ByVal Condition As String, ByVal Order As String)
Dim rec As New ADODB.Recordset
Dim x
Dim z As Integer
'Condition = Replace(Condition, "'", "")
MDIFrm_MAIN.tmeConnChck.Enabled = False
Set rec = opndbaseFMIS.Execute("Select " & field & " as field," & Field2 & " as field2 from " & TableName & " where " & Condition & "  group by " & field & "," & Field2 & " order by " & Order & "")
    lst.ListItems.Clear
    If rec.RecordCount > 0 Then
        For z = 1 To rec.RecordCount
                If IsNull(rec.Fields!field) = False Then
                    Set x = lst.ListItems.Add(, , rec.Fields!field)
                    x.SubItems(1) = Trim(rec.Fields!Field2)
                End If
            rec.MoveNext
            'DoEvents
        Next z
    End If
rec.Close
Set rec = Nothing
MDIFrm_MAIN.tmeConnChck.Enabled = True
End Function
Public Function GetAccountNameForSetUp(ByVal lst As ListView, ByVal TableName As String, ByVal field As String, ByVal Field2 As String, ByVal Condition As String, ByVal Order As String)
Dim rec As New ADODB.Recordset
Dim x
Dim z As Integer
'Condition = Replace(Condition, "'", "")
MDIFrm_MAIN.tmeConnChck.Enabled = False
Set rec = opndbaseFMIS.Execute("Select " & field & " as field," & Field2 & " as field2,max(actioncode) as Actioncode from " & TableName & " where " & Condition & "  group by " & field & "," & Field2 & " order by " & Order & "")
    lst.ListItems.Clear
    If rec.RecordCount > 0 Then
        For z = 1 To rec.RecordCount
                If IsNull(rec.Fields!field) = False Then
                    Set x = lst.ListItems.Add(, , rec.Fields!field)
                    x.SubItems(1) = Trim(rec.Fields!Field2)
                    x.SubItems(2) = IIf((rec.Fields!ActionCode) = 0, "No", "Yes")
                End If
            rec.MoveNext
            'DoEvents
        Next z
    End If
rec.Close
Set rec = Nothing
MDIFrm_MAIN.tmeConnChck.Enabled = True
End Function
Public Sub SetAnimation(ByVal objAnimation As Animation)
    objAnimation.Visible = True
    objAnimation.Open App.path & "\Avis\refresh.avi"
    objAnimation.Play
End Sub
Public Sub UnsetAnimation(ByVal objAnimation As Animation)
    objAnimation.Stop
    objAnimation.Close
    objAnimation.Visible = False
End Sub
Public Function CheckIfMoreOBR(ByVal dvno As String) As String
Dim rec As New ADODB.Recordset
CheckIfMoreOBR = ""
rec.Open "Select case when moreobr = 1 then rtrim(ltrim(obrno)) + ',' + rtrim(ltrim(obr2)) else obrno end  as OBR from tblAMIS_IncomingDVTrns  where dvno = '" & dvno & "' and actioncode = 1", opndbaseFMIS, adOpenStatic, adLockOptimistic
    If rec.RecordCount > 0 Then
        CheckIfMoreOBR = rec!ObR
    End If
rec.Close
End Function

Public Function GetNewConnection(ByVal ConnID As Long)
Dim rec As New ADODB.Recordset
Dim ConnString As String
rec.Open "Select string from tblreff_ManageConnection where trnno = " & ConnID & "", opndbaseFMIS, adOpenStatic, adLockOptimistic
    If rec.RecordCount > 0 Then
        ConnString = Trim(rec!String)
        If NewConnection.State = 1 Then: NewConnection.Close
        NewConnection.ConnectionTimeout = 120
        NewConnection.CursorLocation = adUseClient
        DoEvents
        NewConnection.Open ConnString
    End If
rec.Close
Set rec = Nothing
End Function
Public Function GetPmisOfficecode(ByVal fmisofficeid As Integer) As Integer
Dim rec As New ADODB.Recordset

GetPmisOfficecode = 0
rec.Open "Select [pmisOfficeID] from [tblREF_AIS_Offices] where  [FMISOfficeID] = " & fmisofficeid & "", opndbaseFMIS, adOpenStatic, adLockOptimistic
    If rec.RecordCount > 0 Then
        GetPmisOfficecode = rec!PMISOfficeID
    End If
rec.Close
End Function
Public Function ExecFunction(ByVal Functions As String) As String
Dim rec As New ADODB.Recordset
ExecFunction = ""
'MsgBox Functions
Set rec = opndbaseFMIS.Execute(Functions & " as Field")
        ExecFunction = rec!field
rec.Close
End Function
Public Function UpdateExtractor(ByVal Tables As String, ByVal columns As String, ByVal field As String, ByVal WhereCon As String, ByVal Conditions As String) As String
Dim rec As New ADODB.Recordset
Dim sql As String
sql = "Update " & Trim(Tables) & " set " & Trim(columns) & "='" & Trim(field) & "' where " & Trim(WhereCon) & "=" & Trim(Conditions)
'MsgBox sql
 opndbaseFMIS.Execute (sql)
End Function
Public Function GetLvlbyCode(ByVal accntcode As String) As Integer
Dim xx As Variant
Dim str() As String
Dim lvl As Integer
    xx = Split(accntcode, "-")
    str() = Split(accntcode, "-", -1, vbTextCompare)
    lvl = UBound(xx) + 1
    If lvl = 1 Then
        lvl = 0
    End If
GetLvlbyCode = lvl
End Function
Public Function CheckIfHavenullAccnt(ByVal REFF As String, ByVal Transtype As Integer) As Boolean
CheckIfHavenullAccnt = False
If ExecFunction("Select [fmis].[dbo].[MPfunc_ChckIfHaveAccntcode] ('" & REFF & "'," & Transtype & ")") = 1 Then
    CheckIfHavenullAccnt = True
End If
End Function
Public Function LoadAccountsByNames(ByVal accountcode As String, ByVal Condition As String)
Dim ARec As New ADODB.Recordset
Dim x As Integer
    Set ARec = opndbaseFMIS.Execute("exec Proc_getNamebychildCode @childaccountcode = '" & accountcode & "', @Condition = '" & Condition & "'")
        If ARec.RecordCount > 0 Then
            LoadAccountsByNames = ARec!Accountfullname
       ' inRec = True
        End If
    ARec.Close
    Set ARec = Nothing
End Function
Public Function CheckWhatENtry(ByVal MSHFlex As MSHFlexGrid) As Integer
Dim x As Long
CheckWhatENtry = 4
For x = 1 To MSHFlex.Rows - 1
    If MSHFlex.TextMatrix(x, 1) = "101" Then
        CheckWhatENtry = 1
        Exit For
    End If
Next x
End Function
Public Function TransactionLogging(ByVal strTransaction As String, ByVal strTablename As String, ByVal strFormName As String, ByVal IP As String)
On Error Resume Next
    If opndbaseFMIS.State = 1 Then
        opndbaseFMIS.Execute "EXECUTE [fmis].[dbo].[MPproc_Loggers] @userid = '" & ActiveUserID & "',@transtype = '" & strTransaction & "',@tblname = '" & strTablename & "',@formname = '" & strFormName & "',@computername = '" & GetPCName & "',@IP = '" & IP & "',@datetimeentered = '" & Now & "'"
    End If
End Function
Public Function OnlineLogging(ByVal UserID As String, ByVal IP As String, ByVal port As Long)
    opndbaseFMIS.Execute "insert into [fmis].[dbo].[tblAMIS_OnlineUser]([UserID],[IP],[port]) values  ('" & UserID & "','" & IP & "','" & port & "')"
End Function
Public Function OnlineDeleteLogging()
    opndbaseFMIS.Execute "delete FROM [fmis].[dbo].[tblAMIS_OnlineUser] where userid = '" & ActiveUserID & "'"
End Function
Public Function LoadAccntgTransType(ByVal cmb As ComboBox)
Dim rec As New ADODB.Recordset
Dim x As Double
cmb.Clear
Set rec = opndbaseFMIS.Execute("Select * from MPFunc_LoadAccntgTransType()")
If rec.RecordCount > 0 Then
For x = 1 To rec.RecordCount
cmb.AddItem rec!name
cmb.ItemData(cmb.NewIndex) = rec!Code
rec.MoveNext
Next x
End If
End Function
Public Function getQueryDescription(ByVal accountcode As String) As String
Dim rec As New ADODB.Recordset
Set rec = opndbaseFMIS.Execute("Select description from tblAMIS_Qrygenerator4COA where acountcode = '" & accountcode & "'")
If rec.RecordCount > 0 Then
    getQueryDescription = Trim(rec!description)
End If
rec.Close
End Function
Public Function LoadMotherFund(ByVal cmb As ComboBox)
Dim rec As New ADODB.Recordset
Set rec = opndbaseFMIS.Execute("Execute [dbo].[Proc_ConsoFund]")
cmb.Clear
If rec.RecordCount > 0 Then
For x = 1 To rec.RecordCount
    cmb.AddItem rec!motherfundtype
    cmb.ItemData(cmb.NewIndex) = rec!fundcode
    rec.MoveNext
Next x
End If
End Function
Public Sub AddOffice(ByVal lstbox As ListBox)
    Dim OfficeTbl As New ADODB.Recordset
    Dim recCounter As Integer
    Dim LoopCounter As Integer
    OfficeTbl.Open "select officecode,OfficeName,officeid from pmis.dbo.OfficeDescription  order by OfficeName", opndbaseFMIS, adOpenStatic
    If OfficeTbl.RecordCount > 0 Then
        OfficeTbl.MoveFirst
        lstbox.Clear
        For recCounter = 1 To OfficeTbl.RecordCount
            If Len(OfficeTbl!officecode) > 0 Then
                lstbox.AddItem OfficeTbl!Officename ' & Space(50) & "'" & OfficeTbl!officeid
                lstbox.ItemData(LoopCounter) = OfficeTbl!OfficeID
                LoopCounter = LoopCounter + 1
            End If
        OfficeTbl.MoveNext
        Next
    End If
    OfficeTbl.Close
    Set OfficeTbl = Nothing
End Sub
Public Function GetEmpPosition(ByVal UserID As String) As String
Dim rec As New ADODB.Recordset
Set rec = opndbaseFMIS.Execute("Select position FROM tblAMIS_UserRegistry where userid= '" & ActiveUserID & "' and actioncode = 1")
If rec.RecordCount > 0 Then
    GetEmpPosition = IIf(IsNull(rec!Position) = True, "Accounting Clerk", Trim(rec!Position))
'    GetEmpPosition = Replace(GetEmpPosition, "(", "")
'    GetEmpPosition = Replace(GetEmpPosition, ")", "")
Else
    GetEmpPosition = "N/A"
End If
Set rec = Nothing
End Function
Public Function GetEmpName(ByVal UserID As String) As String
Dim rec As New ADODB.Recordset
Dim rec1 As New ADODB.Recordset
Set rec = opndbaseFMIS.Execute("Select  case when rtrim(ltrim(suffix)) <> '' then  firstname + ' ' + left(Mi,1) + '. ' + lastname + ', '  + suffix else  firstname + ' ' + left(Mi,1) + '. ' + lastname end as Name from pmis.dbo.employee where [SwipEmployeeID] = '" & ActiveUserID & "'")
If rec.RecordCount > 0 Then
    GetEmpName = rec!name
Else
    GetEmpName = "N/A"
End If
Set rec = Nothing
End Function

Public Sub SetMSHGrid(ByVal MSHgrd As MSHFlexGrid, id As Long)
Dim rec As New ADODB.Recordset

Set rec = opndbaseFMIS.Execute("Select * from tblAMIS_GridProperties where id = " & id & " order by cols")
Dim x As Integer
   ' MSHgrd.Clear
    'MSHgrd.Rows = 2
    MSHgrd.Cols = rec.RecordCount ' IIf(LCase(Trim(lblMode)) = "edit", 6, 5)
    
    For x = 1 To rec.RecordCount
        MSHgrd.ColWidth(rec!Cols) = rec!ColWidth
        MSHgrd.ColAlignment(rec!Cols) = rec!ColAlignment
        rec.MoveNext
    Next x
rec.Close
Set rec = Nothing
End Sub
Public Sub LoadEntryInGrid(ByVal MSHgrd As MSHFlexGrid, TYP As Long, whatfield As String, setGrdID As Long)
Dim Drec As New ADODB.Recordset
MSHgrd.Clear
Set Drec = opndbaseFMIS.Execute("Exec dbo.[MPproc_LoadEntryInGrid] @type = " & TYP & ",@whatfield = '" & Trim(whatfield) & "'")
If Drec.RecordCount > 0 Then
Set MSHgrd.Recordset = Drec
    Call SetMSHGrid(MSHgrd, setGrdID)
End If
Drec.Close
End Sub
Public Sub AddImageToDB(ByVal strFile As String, ByVal id As Integer, ByVal description As String, name As String, Version As String, isID As Integer)
Dim rs As ADODB.Recordset
Dim strStream As ADODB.Stream
    'Add the image to the database
    Set strStream = New ADODB.Stream
    strStream.Type = adTypeBinary
    strStream.Open
    strStream.LoadFromFile strFile
    
    opndbaseFMIS.Execute "Delete FROM [fmis].[dbo].[tblAMIS_SystemUpdate] where ISID = " & isID & ""
    Set rs = New ADODB.Recordset
    With rs
        .ActiveConnection = opndbaseFMIS
        .Source = "Select * FROM [fmis].[dbo].[tblAMIS_SystemUpdate] where ISID = " & isID & ""
        .CursorType = adOpenKeyset
        .LockType = adLockOptimistic
        .Open
    End With
    rs.AddNew

    rs.Fields("Name").Value = name
    rs.Fields("version").Value = Version
    rs.Fields("IS").Value = strStream.Read
    rs.Fields("userid").Value = ActiveUserID
    rs.Fields("DatetimeUpload").Value = Now
    rs.Fields("description").Value = description
    rs.Fields("ISID").Value = isID
    'MsgBox Trim(rs!description) & vbNewLine & description
    rs.update
    strStream.Close
    rs.Close

    'Cleanup
    Set strStream = Nothing
    Set rs = Nothing
    Set cn = Nothing
End Sub

Public Function ViewFromDB(ByVal id As String, ByVal TempPath As String) As Boolean
Dim cn As ADODB.Connection
Dim rs As ADODB.Recordset
Dim strStream As ADODB.Stream
Dim strSQL As String
    
    strSQL = "Select * FROM [fmis].[dbo].[tblAMIS_SystemUpdate] where ISID = 1"
                
    Set rs = New ADODB.Recordset
    With rs
        .ActiveConnection = opndbaseFMIS
        .Source = strSQL
        .Open
    End With
    
    If Not (rs.BOF And rs.EOF) Then
        Set strStream = New ADODB.Stream
        strStream.Type = adTypeBinary
        strStream.Open
    
        strStream.Write rs!IS
    
        strStream.SaveToFile TempPath, adSaveCreateOverWrite
        
        strStream.Close
        Set strStream = Nothing
        
        ViewFromDB = True
    End If
    rs.Close
    Set rs = Nothing
End Function

Public Function CheckTheUpdate(ByVal Version As String) As Boolean
Dim rec As New ADODB.Recordset
Set rec = opndbaseFMIS.Execute("Select version,description from dbo.[tblAMIS_SystemUpdate] where ISID = 1")
CheckTheUpdate = False
If rec.RecordCount > 0 Then
    If Trim(rec!Version) = Trim(Version) Then
    
    ElseIf Trim(Version) > Trim(rec!Version) Then
    
    Else
        CheckTheUpdate = True
        SystemVersion = Trim(rec!Version)
        SystemDescription = Trim(IIf(IsNull(rec!description), "", rec!description))
    End If
End If
End Function
Public Function LoadImageUser()
Dim rs As ADODB.Recordset
Dim strStream As ADODB.Stream
Dim strSQL As String
    
    strSQL = "SELECT Pic FROM [fmis].[dbo].[tblAMIS_UserRegistry] where userid = '" & Trim(ActiveUserID) & "' and actioncode =1"
                
    Set rs = New ADODB.Recordset
    With rs
        .ActiveConnection = opndbaseFMIS
        .Source = strSQL
        .Open
    End With
    
    If Not (rs.BOF And rs.EOF) Then
         Set strStream = New ADODB.Stream
        strStream.Type = adTypeBinary
        strStream.Open
        If IsNull(rs!pic) = False Then
        strStream.Write rs!pic
        strStream.SaveToFile App.path & "\img.bmp", adSaveCreateOverWrite
        strStream.Close
        End If
        Set strStream = Nothing
    End If
    
    rs.Close
    Set rs = Nothing
End Function

      Public Function SetTopMostWindow(hwnd As Long, Topmost As Boolean) _
         As Long

         If Topmost = True Then 'Make the window topmost
            SetTopMostWindow = SetWindowPos(hwnd, HWND_TOPMOST, 0, 0, 0, _
               0, FLAGS)
         Else
            SetTopMostWindow = SetWindowPos(hwnd, HWND_NOTOPMOST, 0, 0, _
               0, 0, FLAGS)
            SetTopMostWindow = False
         End If
      End Function
Public Function LogTrans(ByVal dvno As String, ByVal status As Integer)
opndbaseFMIS.Execute "EXECUTE [fmis].[dbo].[MPproc_SavetransFortracking] @stat = " & status & ",@dvno = '" & dvno & "',@userid = '" & ActiveUserID & "'"
End Function
Public Function IfExist(ByVal sqlstatement As String) As Integer
Dim rec As New ADODB.Recordset
IfExist = 0
Set rec = opndbaseFMIS.Execute(sqlstatement)
If rec.RecordCount > 0 Then
    IfExist = 1
End If
rec.Close
Set rec = Nothing
End Function
Public Function GeTRecord(ByVal sqlstatement As String) As ADODB.Recordset
Set GeTRecord = opndbaseFMIS.Execute(sqlstatement)
If rec.RecordCount = 0 Then
    MsgBox "No Record Found..!", vbInformation, "System Information"
End If
End Function
Public Function DisApprovedAndApprove(ByVal jevno As String, desc As String, TYP As Integer) As Integer
Dim rec As New ADODB.Recordset
Dim rec1 As New ADODB.Recordset
Set rec = opndbaseFMIS.Execute("Select jevno From dbo.tblAMIS_FinalJEV where jevno = '" & jevno & "' and  actioncode = 1 order by jevno")
If rec.RecordCount > 0 Then
    Set rec1 = opndbaseFMIS.Execute("exec fmis.dbo.MPproc_ApproveDisApproveJEV @jevno = '" & jevno & "',@userID = '" & ActiveUserID & "',@description  = '" & desc & "',@typ = " & TYP & "")
    If rec1.RecordCount > 0 Then
        DisApprovedAndApprove = rec1!Stat
    End If
    rec1.Close
    Set rec1 = Nothing
Else
    DisApprovedAndApprove = 0
End If
rec.Close
Set rec = Nothing
End Function

Public Function GetStatOfPostedTransaction(ByVal jevno As String) As String
Dim SRec As New ADODB.Recordset
On Error GoTo bad
GetStatOfPostedTransaction = ""
 Set SRec = opndbaseFMIS.Execute("Select dbo.MPfunc_GetStatOfPostedTransaction('" & jevno & "') as R")
 If SRec.RecordCount > 0 Then
    GetStatOfPostedTransaction = SRec!r
 End If
 SRec.Close
 Set SRec = Nothing
 Exit Function
bad:
 MsgBox err.description
End Function

Public Function CheckIfHaveCFentry(ByVal jevno As String) As Boolean
Dim SRec As New ADODB.Recordset
 Set SRec = opndbaseFMIS.Execute("select dbo.[MPfunc_CheckIfHaveCFentry]('" & jevno & "') as R")
 If SRec!r = 0 Then
 CheckIfHaveCFentry = False
 Else
 CheckIfHaveCFentry = True
 End If
 SRec.Close
 Set SRec = Nothing
End Function
Public Function SystemMaintainance(ByVal RID As Integer) As Boolean
Dim Rrec As New ADODB.Recordset
SystemMaintainance = False
Set Rrec = opndbaseFMIS.Execute("select rid from tblAMIS_ReportMaintenace where rid = " & RID & " and status = 2")
If Rrec.RecordCount > 0 Then
    SystemMaintainance = True
End If
Rrec.Close
Set Rrec = Nothing
End Function
Public Function GetExcuteScalar(What As Integer, p As String) As String
Dim rec As New ADODB.Recordset
GetEmailByDvno = ""
Set rec = opndbaseFMIS.Execute("execute dbo.MPproc_ExecuteScalar @what = " & What & ",@p = '" & p & "'")
If rec.RecordCount > 0 Then
    GetExcuteScalar = IIf(IsNull(rec!field), "", rec!field)
End If
'MsgBox rec!field
rec.Close
Set rec = Nothing
End Function
Public Function ExcuteScalar(qry As String) As String
Dim rec As New ADODB.Recordset
ExcuteScalar = ""
Set rec = opndbaseFMIS.Execute(qry)
If rec.RecordCount > 0 Then
    ExcuteScalar = IIf(IsNull(rec!field), "", rec!field)
End If
'MsgBox rec!field
rec.Close
Set rec = Nothing
End Function
Public Function CheckIfExists(sqlquery As String) As Boolean
Dim rec As New ADODB.Recordset
CheckIfExists = False
Set rec = opndbaseFMIS.Execute(sqlquery)
If rec.RecordCount > 0 Then
    CheckIfExists = True
End If
rec.Close
Set rec = Nothing
End Function
Public Function VerifyCheckNo(ByVal checkno As String) As Integer
Dim opnChk As New ADODB.Recordset
Dim opnChkRoute As New ADODB.Recordset

opnChk.Open "Select chkno from tblAMIS_AccountantAdvice where chkno='" & checkno & "' and actioncode=1", opndbaseFMIS, adOpenStatic, adLockOptimistic
If opnChk.RecordCount <> 0 Then
    VerifyCheckNo = 1 'Already Prepared with Accountant Advice
Else
    opnChkRoute.Open "SELECT * FROM tblCMS_CDCheckRoutine where actioncode=1 and  checkno='" & checkno & "'", opndbaseFMIS, adOpenStatic, adLockOptimistic
        If opnChkRoute.RecordCount <> 0 Then
            VerifyCheckNo = opnChkRoute!actiontype '3='Newly Prepared, 4=Check for Signature, 5=Ready for Release, 6=Already Released
        Else
            VerifyCheckNo = 0 'Not Yet Prepared
        End If
    opnChkRoute.Close
    Set opnChkRoute = Nothing

End If
opnChk.Close
Set opnChk = Nothing

End Function
Public Function centerme(ByVal frm As Form)
On Error Resume Next
Dim H, w, FW, FFW, FH, FFH, x, y As Long
frm.ScaleMode = 5
H = MDIFrm_MAIN.Height
FH = frm.Height
x = frm.ScaleHeight / 2
FFH = (H - FH) / x

w = MDIFrm_MAIN.Width
y = frm.ScaleWidth / 2
FW = frm.Width
FFW = (w - FW)

frm.Top = FFH / 2
frm.Left = FFW / 2
End Function

Public Sub Format_Number(ByVal txt As TextBox)
    If txt.Text = "" Then
        Exit Sub
    End If
    If IsNumeric(txt.Text) = True Then
       txt.Text = Format(txt.Text, "#,##0.00")
    Else
        MsgBox "None numeric entry.", vbCritical + vbInformation, "System Message"
        txt.SetFocus
    End If
End Sub

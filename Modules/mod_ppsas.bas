Attribute VB_Name = "mod_ppsas"

Public Function Get_MaxValue_COA(ByVal AccountChildParentID As Long)
Dim rec As New ADODB.Recordset
Dim sql As String
sql = "select Accounting.fn_getMaxValue_COA(" & AccountChildParentID & ") as maxid"
rec.Open sql, opndbaseFMIS, adOpenStatic, adLockOptimistic
    If rec.RecordCount > 0 Then
        Get_MaxValue_COA = IIf(IsNull(rec!maxid), 0, rec!maxid)
    Else
        Get_MaxValue_COA = 1
    End If
rec.Close
Set rec = Nothing
End Function

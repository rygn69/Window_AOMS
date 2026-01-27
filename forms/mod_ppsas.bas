Attribute VB_Name = "mod_ppsas"
Public Function Get_Account_by_ChartAccountParentID(ByVal lst As ListView, ByVal ChartAccountParentID As Long, ByVal Order As String)
Dim rec As New ADODB.Recordset
Dim x
Dim z As Integer
rec.Open "Select [ChartAccountID],[Accountcode],Accountname from [Accounting].[tbl_l_ChartOfAccountsParent] where ChartAccountParentID = " & ChartAccountParentID & " order by Accountname", opndbaseFMIS, adOpenStatic, adLockOptimistic
    lst.ListItems.Clear
    If rec.RecordCount > 0 Then
        For z = 1 To rec.RecordCount
                    Set x = lst.ListItems.Add(, , rec.Fields!ChartAccountID)
                    x.SubItems(1) = Trim(rec.Fields!accountcode)
                    x.SubItems(2) = Trim(rec.Fields!Accountname)
            rec.MoveNext
        Next z
    End If
rec.Close
Set rec = Nothing
End Function

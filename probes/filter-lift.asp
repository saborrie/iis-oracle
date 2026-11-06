<%
' What ADO does when a Recordset filter is applied and then lifted
' (AxonASP issue #189: lifting the filter there leaves only the previously
' matching rows). Uses a fabricated recordset, so no database is involved
' and the filtering is purely client-side ADO.
'
' ADO constants used numerically so the page is self-contained:
' 200 = adVarChar, 3 = adUseClient / adLockOptimistic.
Option Explicit
Response.ContentType = "text/plain"

Sub Say(s)
	Response.Write s & vbLf
End Sub

Sub Check(label)
	If Err.Number <> 0 Then
		Say label & "  ->  ERR " & Err.Number & " (" & Err.Description & ")"
		Err.Clear
	End If
End Sub

Function Names(rs)
	Dim parts
	parts = ""
	If Not (rs.BOF And rs.EOF) Then rs.MoveFirst
	Do While Not rs.EOF
		parts = parts & rs("name") & ";"
		rs.MoveNext
	Loop
	Names = parts
End Function

Dim rs, n

On Error Resume Next

Set rs = Server.CreateObject("ADODB.Recordset")
rs.CursorLocation = 3
rs.Fields.Append "name", 200, 50
rs.LockType = 3
rs.Open
For Each n In Array("apple", "pear", "plum")
	rs.AddNew
	rs("name") = n
	rs.Update
Next
Check "setup"

Say "opened  ->  count=" & rs.RecordCount & " rows=" & Names(rs)
Check "opened"

rs.Filter = "name = 'pear'"
Check "apply-filter"
Say "filtered  ->  count=" & rs.RecordCount & " rows=" & Names(rs)
Check "filtered"

rs.Filter = ""
Check "lift-with-empty-string"
Say "lifted-empty  ->  count=" & rs.RecordCount & " rows=" & Names(rs)
Check "lifted-empty"

rs.Filter = "name = 'plum'"
rs.Filter = 0
Check "lift-with-adFilterNone"
Say "lifted-zero  ->  count=" & rs.RecordCount & " rows=" & Names(rs)
Check "lifted-zero"

' Where does the cursor sit right after a filter is applied and lifted?
rs.Filter = "name = 'pear'"
Say "position-after-apply  ->  AbsolutePosition=" & rs.AbsolutePosition & " current=" & rs("name")
Check "position-after-apply"
rs.Filter = ""
Say "position-after-lift  ->  AbsolutePosition=" & rs.AbsolutePosition & " current=" & rs("name")
Check "position-after-lift"

Say "end-of-probe"
%>

<%
' Null in the ASP context: Response.Write, and a Null field value read back
' from a fabricated (disconnected) ADODB recordset - the shape every page
' with nullable database columns hits. No database involved.
'
' ADO constants used numerically so the page is self-contained:
' 200 = adVarChar, 32 = adFldIsNullable, 3 = adLockOptimistic / adUseClient.
Option Explicit
Response.ContentType = "text/plain"

Sub Say(s)
	Response.Write s & vbLf
End Sub

Dim v, rs

On Error Resume Next

' A Null inside a & concatenation, assigned before writing.
Err.Clear
v = "[" & Null & "]"
If Err.Number <> 0 Then
	Say "concat-into-variable  ->  ERR " & Err.Number & " (" & Err.Description & ")"
Else
	Say "concat-into-variable  ->  " & v & " (vartype=" & VarType(v) & ")"
End If

' The same through a fabricated recordset field holding Null.
Err.Clear
Set rs = Server.CreateObject("ADODB.Recordset")
rs.CursorLocation = 3
rs.Fields.Append "note", 200, 50, 32
rs.LockType = 3
rs.Open
rs.AddNew
rs("note") = Null
rs.Update
If Err.Number <> 0 Then
	Say "fabricated-null-field setup  ->  ERR " & Err.Number & " (" & Err.Description & ")"
	Err.Clear
Else
	Say "IsNull(rs(""note""))  ->  " & IsNull(rs("note"))
	Err.Clear
	v = "[" & rs("note") & "]"
	If Err.Number <> 0 Then
		Say "concat-null-field  ->  ERR " & Err.Number & " (" & Err.Description & ")"
		Err.Clear
	Else
		Say "concat-null-field  ->  " & v & " (vartype=" & VarType(v) & ")"
	End If
End If

' Writing Null directly. Last, because its output (if any) interleaves.
Say "direct-write-of-Null-follows:"
Err.Clear
Response.Write Null
Say ""
Say "Response.Write Null  ->  err=" & Err.Number & " (" & Err.Description & ")"
Err.Clear

Say "end-of-probe"
%>

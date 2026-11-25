<%@ Language=VBScript %>
<%
' Response.End called inside a function, after Response.Clear, while the page
' (the caller) is under On Error Resume Next: does the page stop, and is what
' was written before Response.Clear gone?
Response.ContentType = "text/plain"
Response.Write "written before, then cleared" & vbLf

Function Stops()
	On Error Resume Next
	Response.Clear
	Response.Status = "503 Service Unavailable"
	Response.Write "stopped inside the function, Err " & Err.Number & vbLf
	Response.End
	Stops = "returned"
End Function

On Error Resume Next
Dim v
v = Stops()
Response.Write "AFTER: the page carried on, v=[" & v & "]" & vbLf
%>

<%@ Language=VBScript %>
<%
' Under Option Explicit: reading and assigning a name that was never declared.
' Each case reports Err.Number/Description under On Error Resume Next.
Option Explicit
Dim x
Response.ContentType = "text/plain"

Sub Report(label)
	If Err.Number <> 0 Then
		Response.Write label & " -> ERR " & Err.Number & " (" & Err.Description & ")" & vbLf
	Else
		Response.Write label & " -> no error; x=[" & x & "] (" & TypeName(x) & ")" & vbLf
	End If
	Err.Clear
End Sub

On Error Resume Next
Err.Clear
x = "unset" : x = NoSuchVar : Report "x = NoSuchVar"
x = "unset" : x = NoSuchVar + 1 : Report "x = NoSuchVar + 1"
x = "unset" : NoSuchVar = 1 : Report "NoSuchVar = 1"
%>

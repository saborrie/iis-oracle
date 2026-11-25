<%@ Language=VBScript %>
<%
' Calling a function or sub that does not exist, with Option Explicit.
' Each case reports Err.Number/Description under On Error Resume Next, and
' what the target variable holds afterwards.
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
x = "unset" : x = NoSuchFunction() : Report "x = NoSuchFunction()"
x = "unset" : x = NoSuchFunction(1, 2) : Report "x = NoSuchFunction(1, 2)"
x = "unset" : Call NoSuchSub : Report "Call NoSuchSub"
x = "unset" : Call NoSuchSub(1) : Report "Call NoSuchSub(1)"
x = "unset" : NoSuchSub : Report "NoSuchSub"
x = "unset" : NoSuchSub 1 : Report "NoSuchSub 1"
x = "unset" : x = NoSuchVar : Report "x = NoSuchVar (control)"
%>

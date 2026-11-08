<%@ Language=VBScript %>
<%
' Harder variant: the commented-out arm re-enters script with real statements
' in the gap (an If closed by End If), the way commented-out code with "%" ">"
' sequences inside the comments produces naked statements between Select Case
' and the first real Case.
Option Explicit
Dim x, fn
x = "b" : fn = "field"
Select Case x
' if false then %><option value="<%=fn%>">x</option><% end if
' %><select id="<%=fn%>"><%
Case "a"
	Response.Write "A"
Case "b"
	Response.Write "B"
End Select
%>
[done]

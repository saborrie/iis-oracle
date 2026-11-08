<%@ Language=VBScript %>
<%
' A commented-out Case arm whose comment contains a tag close: the "%" ">"
' inside the comment ends the script block, so HTML and an expression block
' sit between Select Case and the first real Case. Does IIS compile this,
' and does any of the gap content reach the output?
Option Explicit
Dim x, fn
x = "b" : fn = "field"
Select Case x
' a commented-out arm follows; its comment ends the script block mid-line:
' %><select id="<%=fn%>"><option>-- Choose --</option></select><%
Case "a"
	Response.Write "A"
Case "b"
	Response.Write "B"
End Select
%>
[done]

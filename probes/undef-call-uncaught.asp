<%@ Language=VBScript %>
<%
' An uncaught call to a function that does not exist, without Option Explicit.
Response.Write "before" & vbLf
x = NoSuchFunction()
Response.Write "after; x=[" & x & "] (" & TypeName(x) & ")" & vbLf
%>

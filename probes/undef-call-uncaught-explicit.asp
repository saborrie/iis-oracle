<%@ Language=VBScript %>
<%
' An uncaught call to a function that does not exist, with Option Explicit.
Option Explicit
Dim x
Response.Write "before" & vbLf
x = NoSuchFunction()
Response.Write "after; x=[" & x & "] (" & TypeName(x) & ")" & vbLf
%>

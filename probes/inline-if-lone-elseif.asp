<%@ Language=VBScript %>
<%
' ElseIf variant of inline-if-lone-else.asp: single-line If body ends at the
' tag boundary, an ElseIf follows in the next script block, no enclosing If.
Option Explicit
Dim a, y
a = 1
y = 0
If a = 1 Then y = 1 %>x<% ElseIf a = 2 Then %>z<% End If
%>
[done y=<%=y%>]

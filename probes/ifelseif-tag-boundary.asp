<%@ Language=VBScript %>
<%
' ElseIf variant: block If at a tag boundary, body ends with a single-line If,
' then an ElseIf in the next script block continues the OUTER block If chain.
' Case A drives the If arm (expect "A2" - inner single-line If true writes 2 via y).
' Case B drives the ElseIf arm (expect "B").
Option Explicit
Dim a, b, y
a = 2 : b = 0 : y = 0
If a <> 0 Then %>A<% If a <> 1 Then y = 2 %><% ElseIf b <> 0 Then %>B<% End If
%>
[first y=<%=y%>]
<%
a = 0 : b = 5 : y = 0
If a <> 0 Then %>A<% If a <> 1 Then y = 2 %><% ElseIf b <> 0 Then %>B<% End If
%>
[second y=<%=y%>]

<%@ Language=VBScript %>
<%
' Same question for a same-line If/ElseIf chain: body ends at the tag
' boundary, End If follows in the next script block.
Option Explicit
Dim a, y
a = 2
y = 0
If a = 1 Then y = 1 ElseIf a = 2 Then y = 2 %><% End If %>
[done y=<%=y%>]

<%@ Language=VBScript %>
<%
' A complete same-line If/Else chain whose body ends at the tag boundary,
' followed by End If in the next script block. Is that End If an orphan
' (like the plain single-line If case measured 2026-09-16), or does the
' Else arm change the answer?
Option Explicit
Dim a, y
a = 3
y = 0
If a = 1 Then y = 1 Else y = 2 %><% End If %>
[done y=<%=y%>]

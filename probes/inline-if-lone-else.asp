<%@ Language=VBScript %>
<%
' A single-line If whose body ends at the tag boundary, followed by an Else in
' the next script block, with NO enclosing If. If the inline If is complete at
' the tag boundary (the End If rule measured earlier), this Else is an orphan
' and the page must fail to compile; if IIS instead lets the inline If absorb
' a tag-boundary Else, the page compiles and prints one of the branches.
Option Explicit
Dim a, y
a = 1
y = 0
If a = 1 Then y = 1 %>x<% Else %>z<% End If
%>
[done y=<%=y%>]

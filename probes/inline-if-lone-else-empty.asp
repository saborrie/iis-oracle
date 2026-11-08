<%@ Language=VBScript %>
<%
' Like inline-if-lone-else.asp but with NOTHING between the closing and opening
' tags: a single-line If whose body ends at the tag boundary, immediately
' followed by a new script block starting with Else, no enclosing If.
' Does the empty tag boundary change whether the Else can attach?
Option Explicit
Dim a, y
a = 1
y = 0
If a = 1 Then y = 1 %><% Else %>z<% End If
%>
[done y=<%=y%>]

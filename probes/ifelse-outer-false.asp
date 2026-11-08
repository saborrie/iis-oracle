<%@ Language=VBScript %>
<%
' Same shape as ifelse-outer-true.asp with the outer condition false:
' if the Else binds to the outer block If, output is "z" and y stays 0.
Option Explicit
Dim a, y
a = 3
y = 0
If a = 1 Then %>x<% If a = 2 Then y = 1 %><% Else %>z<% End If
%>
[done y=<%=y%>]

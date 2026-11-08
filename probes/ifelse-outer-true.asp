<%@ Language=VBScript %>
<%
' A block If opened at a tag boundary whose body ends with a single-line If,
' followed by an Else in the next script block. The single-line If is
' complete at the end of its line (it must not absorb what follows - see the
' earlier End If measurement), so the Else must bind to the OUTER block If.
' Outer condition true, inner condition false: expect output "x", y stays 0.
Option Explicit
Dim a, y
a = 1
y = 0
If a = 1 Then %>x<% If a = 2 Then y = 1 %><% Else %>z<% End If
%>
[done y=<%=y%>]

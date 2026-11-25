<%@ Language=VBScript %>
<%
' Follow-up to date-arithmetic-edges.asp: a number minus a Date, the other
' operand orders, comparing Dates either side of 1899-12-30, CDate of
' negative serials, and a For loop over a Date with a numeric limit.
' Each case reports value, TypeName and Err.Number under On Error Resume Next.
Option Explicit
Response.ContentType = "text/plain"

Dim d, e, a, b, v, x, cnt
d = #2027-01-01 10:30:00#
e = #2026-12-31#
a = #1899-12-29 06:00:00#
b = #1899-12-29 18:00:00#

Sub Report(label)
	If Err.Number <> 0 Then
		Response.Write label & " -> ERR " & Err.Number & " (" & Err.Description & ")" & vbLf
	Else
		Response.Write label & " -> " & v & " (" & TypeName(v) & ")" & vbLf
	End If
	Err.Clear
End Sub

On Error Resume Next
Err.Clear
v = 1 - d : Report "1 - d"
v = 46389 - d : Report "46389 - d"
v = "1" + d : Report """1"" + d"
v = d - True : Report "d - True"
v = d - Empty : Report "d - Empty"
v = Empty - d : Report "Empty - d"
v = Empty + d : Report "Empty + d"
v = d * Empty : Report "d * Empty"
v = d + d + d + d + d + d + d + d + d + d + d + d + d + d + d + d + d + d + d + d + d + d + d + d + d + d + d + d + d + d + d + d + d + d + d + d + d + d + d + d + d + d + d + d + d + d + d + d + d + d + d + d + d + d + d + d + d + d + d + d + d + d + d + d + d + d : Report "66 * d by addition"
v = (b < a) : Report "b < a"
v = (b > a) : Report "b > a"
v = CDbl(b) : Report "CDbl(b)"
v = CDate(-0.25) : Report "CDate(-0.25)"
v = CDate(-1.25) : Report "CDate(-1.25)"
v = CDate(-1.75) : Report "CDate(-1.75)"
cnt = 0
For x = e To 46389
	cnt = cnt + 1
Next
v = cnt : Report "For x = e To 46389 count"
v = x : Report "x after numeric-limit loop"
cnt = 0
For x = 46387 To e + 2
	cnt = cnt + 1
Next
v = cnt : Report "For x = 46387 To e + 2 count"
v = x : Report "x after numeric-start loop"
v = Hex(e) : Report "Hex(e)"
v = Sqr(e) : Report "Sqr(e)"
%>

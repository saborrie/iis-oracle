<%@ Language=VBScript %>
<%
' Edges of Date arithmetic not covered by date-arithmetic.asp: dates before
' 1899-12-30 (negative serials), operand types other than numbers, the
' remaining conversions, sub-second fractions, range overflow, Now()
' resolution, and For loops over dates. Each case reports value, TypeName
' and Err.Number under On Error Resume Next.
Option Explicit
Response.ContentType = "text/plain"

Dim d, e, p, q, h, n, v, x, cnt
d = #2027-01-01 10:30:00#
e = #2026-12-31#
p = #1800-01-01 06:00:00#
q = #1899-12-29 06:00:00#
h = #2027-01-01 12:00:00#

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
v = CDbl(p) : Report "CDbl(p)"
v = CDbl(q) : Report "CDbl(q)"
v = p + 0.5 : Report "p + 0.5"
v = p - 0.5 : Report "p - 0.5"
v = p + 1 : Report "p + 1"
v = q + 1 : Report "q + 1"
v = q + 2 : Report "q + 2"
v = Int(p) : Report "Int(p)"
v = Fix(p) : Report "Fix(p)"
v = Round(p) : Report "Round(p)"
v = CLng(p) : Report "CLng(p)"
v = -p : Report "-p"
v = -e : Report "-e"
v = (p < 0) : Report "p < 0"
v = d - p : Report "d - p"
v = d + e : Report "d + e"
v = d + "1" : Report "d + ""1"""
v = d + "0.5" : Report "d + ""0.5"""
v = d - "1" : Report "d - ""1"""
v = d + True : Report "d + True"
v = d + Empty : Report "d + Empty"
v = d + Null : Report "d + Null"
v = d ^ 1 : Report "d ^ 1"
v = e - d : Report "e - d"
v = d + 1 - 1 : Report "d + 1 - 1"
v = (d + 1 - 1 = d) : Report "d + 1 - 1 = d"
v = (d = 46388.4375) : Report "d = 46388.4375"
v = (46388.4375 = d) : Report "46388.4375 = d"
v = CInt(d) : Report "CInt(d)"
v = CInt(#1950-01-01#) : Report "CInt(#1950-01-01#)"
v = CSng(d) : Report "CSng(d)"
v = CCur(d) : Report "CCur(d)"
v = CBool(d) : Report "CBool(d)"
v = CByte(d) : Report "CByte(d)"
v = CLng(d) : Report "CLng(d)"
v = CLng(h) : Report "CLng(h)"
v = Round(h) : Report "Round(h)"
v = Round(d, 2) : Report "Round(d, 2)"
v = Abs(d) : Report "Abs(d)"
v = Abs(-d) : Report "Abs(-d)"
v = Sgn(d) : Report "Sgn(d)"
v = (d And 1) : Report "d And 1"
v = d \ e : Report "d \ e"
v = d Mod e : Report "d Mod e"
v = d + 0.00001 : Report "d + 0.00001"
v = Second(d + 0.000006) : Report "Second(d + 0.000006)"
v = CDate(0.00001) : Report "CDate(0.00001)"
v = d + 3000000 : Report "d + 3000000"
v = d - 1000000 : Report "d - 1000000"
v = CDate(3000000) : Report "CDate(3000000)"
n = Now()
v = CDbl(n) * 86400 - Round(CDbl(n) * 86400) : Report "Now() sub-second residue"
v = (n + 1 - 1 = n) : Report "n + 1 - 1 = n"
cnt = 0
For x = e To e + 2
	cnt = cnt + 1
Next
v = cnt : Report "For x = e To e + 2 count"
v = x : Report "x after loop"
cnt = 0
For x = e To d Step 0.5
	cnt = cnt + 1
Next
v = cnt : Report "For x = e To d Step 0.5 count"
v = x : Report "x after step loop"
%>

<%@ Language=VBScript %>
<%
' Arithmetic and numeric conversion on Date values: VBScript stores a Date as
' a Double (days since 1899-12-30, time as the fraction), so Date + Number
' should give a Date and CDbl(Date) the day serial. Each case reports value,
' TypeName and Err.Number under On Error Resume Next.
Option Explicit
Response.ContentType = "text/plain"

Dim d, e, v
d = #2027-01-01 10:30:00#
e = #2026-12-31#

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
v = d : Report "d"
v = e : Report "e"
v = d + 1 : Report "d + 1"
v = 1 + d : Report "1 + d"
v = d - 1 : Report "d - 1"
v = d + 0.5 : Report "d + 0.5"
v = d - e : Report "d - e"
v = d * 1 : Report "d * 1"
v = d / 1 : Report "d / 1"
v = d \ 1 : Report "d \ 1"
v = d Mod 7 : Report "d Mod 7"
v = -d : Report "-d"
v = Int(d) : Report "Int(d)"
v = Fix(d) : Report "Fix(d)"
v = Round(d) : Report "Round(d)"
v = CDbl(d) : Report "CDbl(d)"
v = CLng(e) : Report "CLng(e)"
v = CDbl(#1899-12-30#) : Report "CDbl(#1899-12-30#)"
v = CDate(46388.4375) : Report "CDate(46388.4375)"
v = CDate(d + 1) : Report "CDate(d + 1)"
v = (d > 46388) : Report "d > 46388"
v = (d < 46389) : Report "d < 46389"
v = (d = e + 1.4375) : Report "d = e + 1.4375"
v = DateAdd("d", 1, d) : Report "DateAdd(""d"", 1, d)"
v = TypeName(Date() - 1) : Report "TypeName(Date() - 1)"
v = TypeName(Now() + 1) : Report "TypeName(Now() + 1)"
v = (Int(Now() + 1) = Date() + 1) : Report "Int(Now() + 1) = Date() + 1"
%>

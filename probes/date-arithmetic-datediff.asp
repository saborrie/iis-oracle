<%@ Language=VBScript %>
<%
' Date functions applied to the results of Date arithmetic: DateDiff, Year,
' Day, DateAdd and FormatDateTime on d + 1 / d - 1 / Now() + 1 / Date() - 1.
' Each case reports value, TypeName and Err.Number under On Error Resume Next.
Option Explicit
Response.ContentType = "text/plain"

Dim d, v
d = #2027-01-01 10:30:00#

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
v = DateDiff("d", d, d + 1) : Report "DateDiff(""d"", d, d + 1)"
v = DateDiff("h", d - 1, d) : Report "DateDiff(""h"", d - 1, d)"
v = DateDiff("d", Date() - 1, Date()) : Report "DateDiff(""d"", Date() - 1, Date())"
v = DateDiff("n", Now(), Now() + 1) : Report "DateDiff(""n"", Now(), Now() + 1)"
v = Year(d + 1) : Report "Year(d + 1)"
v = Day(d + 1) : Report "Day(d + 1)"
v = Hour(d + 0.5) : Report "Hour(d + 0.5)"
v = DateAdd("d", 1, d + 1) : Report "DateAdd(""d"", 1, d + 1)"
v = FormatDateTime(d - 1, 2) : Report "FormatDateTime(d - 1, 2)"
v = (d + 1 > d) : Report "d + 1 > d"
%>

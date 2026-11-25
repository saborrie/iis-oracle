<%@ Language=VBScript %>
<%
' Number to text and text to number under several locales (SetLocale, then
' Response.LCID and Session.LCID), and a locale-free way to write an amount to
' two places from Currency arithmetic.
Option Explicit
Response.ContentType = "text/plain"
Dim lcids, lcid, v

Sub Report(label)
	If Err.Number <> 0 Then
		Response.Write label & " -> ERR " & Err.Number & " (" & Err.Description & ")" & vbLf
	Else
		Response.Write label & " -> [" & v & "] " & TypeName(v) & vbLf
	End If
	Err.Clear
End Sub

' Two places, halves away from zero, digits and a point only.
Function TwoPlaces(x)
	Dim c, neg, cents, whole, frac
	c = CCur(x)
	neg = (c < 0)
	If neg Then c = -c
	cents = Int(c * 100 + 0.5)
	whole = Int(cents / 100)
	frac = cents - whole * 100
	TwoPlaces = CStr(whole) & "." & Right("0" & CStr(frac), 2)
	If neg And cents <> 0 Then TwoPlaces = "-" & TwoPlaces
End Function

Sub Cases(tag)
	v = GetLocale() : Report tag & " GetLocale"
	v = CStr(1234.5) : Report tag & " CStr(1234.5)"
	v = 1234.5 & "" : Report tag & " 1234.5 & """""
	v = CStr(CCur(1234.5)) : Report tag & " CStr(CCur(1234.5))"
	v = CStr(1234567) : Report tag & " CStr(1234567)"
	v = CStr(CDbl(15056)) : Report tag & " CStr(CDbl(15056))"
	v = FormatNumber(1234.5, 2) : Report tag & " FormatNumber(1234.5, 2)"
	v = IsNumeric("150,00") : Report tag & " IsNumeric(""150,00"")"
	v = IsNumeric("150.00") : Report tag & " IsNumeric(""150.00"")"
	v = CDbl("150.00") : Report tag & " CDbl(""150.00"")"
	v = CDbl("150,00") : Report tag & " CDbl(""150,00"")"
	v = CDbl("1.234,5") : Report tag & " CDbl(""1.234,5"")"
	v = CCur("15.00") : Report tag & " CCur(""15.00"")"
	v = TwoPlaces(150.555) : Report tag & " TwoPlaces(150.555)"
	v = TwoPlaces(-150.555) : Report tag & " TwoPlaces(-150.555)"
	v = TwoPlaces(1234567.8) : Report tag & " TwoPlaces(1234567.8)"
	v = TwoPlaces(0.005) : Report tag & " TwoPlaces(0.005)"
	v = TwoPlaces(0.004999) : Report tag & " TwoPlaces(0.004999)"
	v = TwoPlaces(2.675) : Report tag & " TwoPlaces(2.675)"
	v = TwoPlaces(CCur(7.333) * 12.21) : Report tag & " TwoPlaces(CCur(7.333)*12.21)"
	v = CCur(1500) / 100 : Report tag & " CCur(1500) / 100"
	v = Round(2.5) : Report tag & " Round(2.5)"
	v = Int(-0.5) : Report tag & " Int(-0.5)"
	v = CLng("2610") : Report tag & " CLng(""2610"")"
End Sub

On Error Resume Next
lcids = Array(1033, 2057, 1031, 1036)
For Each lcid In lcids
	v = SetLocale(lcid) : Report "SetLocale " & lcid
	Cases "set" & lcid
Next
v = SetLocale(1033)
Response.LCID = 1031 : v = Response.LCID : Report "Response.LCID = 1031"
Cases "resp1031"
Response.LCID = 1033
Session.LCID = 1031 : v = Session.LCID : Report "Session.LCID = 1031"
Cases "sess1031"
%>

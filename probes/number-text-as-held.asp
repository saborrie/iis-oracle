<%@ Language=VBScript %>
<%
' The text VBScript itself gives a number (CStr, and & "", which a page building
' a URL uses), under several locales: how many significant digits a Double gets,
' where exponent form starts, which decimal separator and minus sign are used,
' and what Single, Currency and Decimal give. Then a writer that takes that text
' and swaps the locale's decimal separator for a point, refusing anything that is
' then not digits with an optional minus and point: its output must be the same
' in every locale, and never rounded.
Option Explicit
Response.ContentType = "text/plain"
Dim lcids, lcid, v

function LocalePoint()
	dim sample
	LocalePoint = ""
	on error resume next
	sample = CStr(CDbl(1.5))
	if Err.Number <> 0 then exit function
	if Len(sample) = 3 and Left(sample, 1) = "1" and Right(sample, 1) = "5" then LocalePoint = Mid(sample, 2, 1)
end function

function IsAmountText(text)
	dim body, point, i
	IsAmountText = False
	body = text
	if Left(body, 1) = "-" then body = Mid(body, 2)
	point = InStr(body, ".")
	if point = 1 or point = Len(body) then exit function
	if point > 0 then body = Left(body, point - 1) & Mid(body, point + 1)
	if body = "" then exit function
	for i = 1 to Len(body)
		if InStr(1, "0123456789", Mid(body, i, 1), vbBinaryCompare) = 0 then exit function
	next
	IsAmountText = True
end function

function AsHeld(value)
	dim text, point
	AsHeld = ""
	on error resume next
	select case VarType(value)
	case vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbByte, 14
		text = CStr(value)
		point = LocalePoint()
		if point = "" then exit function
		if point <> "." then
			if InStr(text, ".") > 0 then exit function
			text = Replace(text, point, ".")
		end if
	case vbString
		text = value
	case else
		exit function
	end select
	if Err.Number <> 0 then exit function
	if not IsAmountText(text) then exit function
	AsHeld = text
end function

Sub Show(label, x)
	Dim c, a, w
	c = "" : a = "" : w = ""
	On Error Resume Next
	Err.Clear
	c = CStr(x)
	If Err.Number <> 0 Then c = "ERR " & Err.Number
	Err.Clear
	a = x & ""
	If Err.Number <> 0 Then a = "ERR " & Err.Number
	Err.Clear
	w = AsHeld(x)
	Response.Write label & " (" & TypeName(x) & ") CStr=[" & c & "] &=[" & a & "] AsHeld=[" & w & "]" & vbLf
	Err.Clear
End Sub

Sub Cases(tag)
	Dim rs
	Response.Write "--- " & tag & " GetLocale=" & GetLocale() & " LocalePoint=[" & LocalePoint() & "] FormatNumber(1234.5,2)=[" & FormatNumber(1234.5, 2) & "]" & vbLf
	Show "1.5", 1.5
	Show "0.5", 0.5
	Show "-0.5", -0.5
	Show "-1234.5", -1234.5
	Show "457.875", 457.875
	Show "37.5*12.21", 37.5 * 12.21
	Show "0.1+0.2", 0.1 + 0.2
	Show "1.1*3", 1.1 * 3
	Show "1/3", 1 / 3
	Show "200/3", 200 / 3
	Show "7.333*12.21", 7.333 * 12.21
	Show "39.5*11.44", 39.5 * 11.44
	Show "0.07*100", 0.07 * 100
	Show "1.005*1000", 1.005 * 1000
	Show "CDbl(100)", CDbl(100)
	Show "CDbl(0)", CDbl(0)
	Show "-CDbl(0)", -CDbl(0)
	Show "1234567.891", 1234567.891
	Show "123456789012345", 123456789012345
	Show "999999999999999", 999999999999999
	Show "1234567890123456", 1234567890123456
	Show "10^14", 10 ^ 14
	Show "10^15", 10 ^ 15
	Show "10^16", 10 ^ 16
	Show "2^53", 2 ^ 53
	Show "12345678901234.5", 12345678901234.5
	Show "123456789012345.6", 123456789012345.6
	Show "0.1", 0.1
	Show "0.01", 0.01
	Show "0.001", 0.001
	Show "0.0001", 0.0001
	Show "0.00001", 0.00001
	Show "0.000012345", 0.000012345
	Show "0.123456789012345678", 0.123456789012345678
	Show "0.004999", 0.004999
	Show "CSng(0.1)", CSng(0.1)
	Show "CSng(1234.567)", CSng(1234.567)
	Show "CSng(457.875)", CSng(457.875)
	Show "CSng(16777217)", CSng(16777217)
	Show "CSng(123456789)", CSng(123456789)
	Show "CCur(457.875)", CCur(457.875)
	Show "CCur(40)", CCur(40)
	Show "CCur(0.0001)", CCur(0.0001)
	Show "CCur(-1234.5678)", CCur(-1234.5678)
	Show "CCur(922337203685477)", CCur(922337203685477)
	Show "CCur(7.333)*12.21", CCur(7.333) * 12.21
	Show "CCur(10)/3", CCur(10) / 3
	Show "CInt(7)", CInt(7)
	Show "CLng(-1234567)", CLng(-1234567)
	Show "CByte(7)", CByte(7)
	Show "True", True
	Show "Empty", Empty
	Show "Now-ish date", CDate(2)
	Show "text 150.5", "150.5"
	Show "text 150,5", "150,5"
	Show "text 1,234.56", "1,234.56"
	Show "text 1e3", "1e3"
	Show "text 007.10", "007.10"
	Show "text -3.14159", "-3.14159"
	Show "text .5", ".5"
	Show "text 5.", "5."
	Show "text empty", ""
	' Decimal, as ADO hands a numeric column over.
	On Error Resume Next
	Err.Clear
	Set rs = CreateObject("ADODB.Recordset")
	rs.Fields.Append "n", 131
	rs.Fields("n").Precision = 18
	rs.Fields("n").NumericScale = 4
	rs.Open
	rs.AddNew
	rs("n") = 457.875
	rs.Update
	If Err.Number <> 0 Then
		Response.Write "ADO numeric -> ERR " & Err.Number & " (" & Err.Description & ")" & vbLf
	Else
		Show "ADO numeric(18,4) 457.875", rs("n").Value
		rs("n") = 100
		rs.Update
		Show "ADO numeric(18,4) 100", rs("n").Value
		rs("n") = -0.5
		rs.Update
		Show "ADO numeric(18,4) -0.5", rs("n").Value
	End If
	Err.Clear
End Sub

On Error Resume Next
lcids = Array(1033, 2057, 1031, 1036, 2055, 1049, 1025, 1065)
For Each lcid In lcids
	Err.Clear
	v = SetLocale(lcid)
	If Err.Number <> 0 Then Response.Write "SetLocale " & lcid & " -> ERR " & Err.Number & vbLf
	Err.Clear
	Cases "SetLocale " & lcid
Next
v = SetLocale(1033)
Response.LCID = 1031
Cases "Response.LCID 1031"
Response.LCID = 1033
Session.LCID = 1031
Cases "Session.LCID 1031"
%>

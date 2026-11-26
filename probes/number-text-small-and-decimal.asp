<%@ Language=VBScript %>
<%
' Follows number-text-as-held.asp: where CStr of a Double near zero turns to
' exponent form, and the text of a Decimal as ADO hands a numeric column over
' (does it keep its scale's trailing zeros?), each step reported on its own.
Option Explicit
Response.ContentType = "text/plain"
Dim lcid, v, rs, e, d

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


Sub Step(label)
	If Err.Number <> 0 Then Response.Write label & " -> ERR " & Err.Number & " (" & Err.Description & ")" & vbLf Else Response.Write label & " -> ok" & vbLf
	Err.Clear
End Sub

Sub Cases(tag)
	Response.Write "--- " & tag & " GetLocale=" & GetLocale() & " LocalePoint=[" & LocalePoint() & "]" & vbLf
	d = 1
	For e = 1 To 20
		d = d / 10
		Show "10^-" & e, d
	Next
	Show "1.5E-7 literal", 1.5E-7
	Show "0.00000123456789", 0.00000123456789
	Show "-0.0000001", -0.0000001
	Show "0.1E-15 * 3", 0.1E-15 * 3
	Show "1E-300", 1E-300
	Show "1.7E+308", 1.7E+308
	Show "CSng(0.00001)", CSng(0.00001)
	Show "CSng(0.0000001)", CSng(0.0000001)
	On Error Resume Next
	Err.Clear
	Set rs = CreateObject("ADODB.Recordset") : Step "create recordset"
	rs.Fields.Append "n", 131, 19 : Step "append adNumeric"
	rs.Fields("n").Precision = 18 : Step "precision"
	rs.Fields("n").NumericScale = 4 : Step "scale"
	rs.Fields.Append "c", 6 : Step "append adCurrency"
	rs.Fields.Append "d", 14 : Step "append adDecimal"
	rs.Fields("d").Precision = 18
	rs.Fields("d").NumericScale = 4 : Step "decimal scale"
	rs.Open : Step "open"
	rs.AddNew : Step "addnew"
	rs.Fields("n").Value = 457.875 : Step "set n"
	rs.Fields("c").Value = 457.875 : Step "set c"
	rs.Fields("d").Value = 457.875 : Step "set d"
	rs.Update : Step "update"
	v = rs.Fields("n").Value : Step "read n"
	Show "adNumeric(18,4) 457.875", v
	v = rs.Fields("c").Value : Step "read c"
	Show "adCurrency 457.875", v
	v = rs.Fields("d").Value : Step "read d"
	Show "adDecimal(18,4) 457.875", v
	rs.Fields("n").Value = 100 : rs.Fields("d").Value = -0.5 : rs.Update : Step "set again"
	v = rs.Fields("n").Value
	Show "adNumeric(18,4) 100", v
	v = rs.Fields("d").Value
	Show "adDecimal(18,4) -0.5", v
	rs.Close
	Err.Clear
End Sub

On Error Resume Next
For Each lcid In Array(1033, 1031)
	v = SetLocale(lcid)
	Cases "SetLocale " & lcid
Next
%>

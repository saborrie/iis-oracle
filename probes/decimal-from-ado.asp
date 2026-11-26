<%@ Language=VBScript %>
<%
' What VBScript can do with a Decimal, as ADO hands a numeric column over: each
' operation on its own, since a Sub that took one printed nothing at all in
' number-text-small-and-decimal.asp.
Option Explicit
Response.ContentType = "text/plain"
Dim lcid, v, r, rs

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


Sub Report(label)
	If Err.Number <> 0 Then
		Response.Write label & " -> ERR " & Err.Number & " (" & Err.Description & ")" & vbLf
	Else
		Response.Write label & " -> [" & r & "]" & vbLf
	End If
	Err.Clear
	r = "(unset)"
End Sub

Function TypeOfArg(x)
	TypeOfArg = VarType(x)
End Function

Sub Probe(tag)
	On Error Resume Next
	Err.Clear
	Response.Write "-- " & tag & vbLf
	r = "(unset)"
	r = VarType(v) : Report "VarType"
	r = TypeName(v) : Report "TypeName"
	r = IsNumeric(v) : Report "IsNumeric"
	r = CStr(v) : Report "CStr"
	r = v & "" : Report "& """""
	r = CStr(CDbl(v)) : Report "CStr(CDbl)"
	r = CStr(CCur(v)) : Report "CStr(CCur)"
	r = (v < 0) : Report "v < 0"
	r = v * 2 : Report "v * 2"
	r = TypeName(v * 2) : Report "TypeName(v * 2)"
	r = TypeOfArg(v) : Report "VarType inside a function"
	r = AsHeld(v) : Report "AsHeld"
	r = AsHeld(rs.Fields(0).Value) : Report "AsHeld(field 0 value)"
	r = "gross=" & v : Report """gross="" & v"
	r = CStr(v + 0) : Report "CStr(v + 0)"
	r = (v > 0) : Report "v > 0"
End Sub

Sub Cases(tag)
	Response.Write "--- " & tag & " GetLocale=" & GetLocale() & " LocalePoint=[" & LocalePoint() & "]" & vbLf
	On Error Resume Next
	Err.Clear
	Set rs = CreateObject("ADODB.Recordset")
	rs.Fields.Append "n", 131, 19
	rs.Fields("n").Precision = 18
	rs.Fields("n").NumericScale = 4
	rs.Fields.Append "d", 14
	rs.Fields("d").Precision = 18
	rs.Fields("d").NumericScale = 4
	rs.Fields.Append "b", 20
	rs.Open
	rs.AddNew
	rs.Fields("n").Value = 457.875
	rs.Fields("d").Value = -1234.5
	rs.Fields("b").Value = 150
	rs.Update
	r = "" : Report "recordset made"
	v = rs.Fields("n").Value : Probe "adNumeric(18,4) 457.875"
	v = rs.Fields("d").Value : Probe "adDecimal(18,4) -1234.5"
	v = rs.Fields("b").Value : Probe "adBigInt 150"
	rs.Fields("n").Value = 100
	rs.Update
	v = rs.Fields("n").Value : Probe "adNumeric(18,4) 100"
	rs.Close
	Err.Clear
End Sub

On Error Resume Next
For Each lcid In Array(1033, 1031)
	v = SetLocale(lcid)
	Cases "SetLocale " & lcid
Next
%>

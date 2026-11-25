<%@ Language=VBScript %>
<%
' A locale-free way to write an amount as digits, a point and two places, and to
' percent-encode text as UTF-8, run under several locales: the output must be
' the same in each.
Option Explicit
Response.ContentType = "text/plain"
Dim lcids, lcid, cases, i, v

function AmountText(value)
	dim cents, negative, text, whole, fraction, i, c, point
	AmountText = ""
	on error resume next
	select case VarType(value)
	case vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbByte, 14
		negative = (value < 0)
		cents = Int(Round(Abs(CDbl(value)) * 100, 6) + 0.5)
	case vbString
		text = value
		negative = (Left(text, 1) = "-")
		if negative then text = Mid(text, 2)
		point = InStr(text, ".")
		if point = 0 then
			whole = text
			fraction = ""
		else
			whole = Left(text, point - 1)
			fraction = Mid(text, point + 1)
		end if
		if whole = "" or Len(whole) > 13 or (point > 0 and fraction = "") then exit function
		for i = 1 to Len(whole & fraction)
			c = Mid(whole & fraction, i, 1)
			if InStr(1, "0123456789", c, vbBinaryCompare) = 0 then exit function
		next
		fraction = Left(fraction & "000", 3)
		cents = CDbl(whole) * 100 + CDbl(Left(fraction, 2))
		if CInt(Right(fraction, 1)) >= 5 then cents = cents + 1
	case else
		exit function
	end select
	if Err.Number <> 0 then exit function
	if cents >= 10 ^ 14 then exit function
	whole = Int(cents / 100)
	fraction = cents - whole * 100
	text = CStr(whole) & "." & Right("0" & CStr(fraction), 2)
	if negative and cents <> 0 then text = "-" & text
	if Err.Number <> 0 then exit function
	AmountText = text
end function

function PercentEncode(text)
	dim result, i, c, code, low
	result = ""
	text = CStr(text)
	i = 1
	do while i <= Len(text)
		c = Mid(text, i, 1)
		code = AscW(c)
		if code < 0 then code = code + 65536
		if (code >= 48 and code <= 57) or (code >= 65 and code <= 90) or (code >= 97 and code <= 122) or c = "-" or c = "." or c = "_" or c = "~" then
			result = result & c
		elseif code < 128 then
			result = result & HexByte(code)
		elseif code < 2048 then
			result = result & HexByte(192 + (code \ 64)) & HexByte(128 + (code and 63))
		elseif code >= 55296 and code <= 56319 and i < Len(text) then
			low = AscW(Mid(text, i + 1, 1))
			if low < 0 then low = low + 65536
			code = 65536 + (code - 55296) * 1024 + (low - 56320)
			result = result & HexByte(240 + (code \ 262144)) & HexByte(128 + ((code \ 4096) and 63)) & _
				HexByte(128 + ((code \ 64) and 63)) & HexByte(128 + (code and 63))
			i = i + 1
		else
			result = result & HexByte(224 + (code \ 4096)) & HexByte(128 + ((code \ 64) and 63)) & HexByte(128 + (code and 63))
		end if
		i = i + 1
	loop
	PercentEncode = result
end function

function HexByte(byteValue)
	HexByte = "%" & Right("0" & Hex(byteValue), 2)
end function

lcids = Array(1033, 1031, 1036, 2057)
For Each lcid In lcids
	SetLocale lcid
	Response.Write "--- " & lcid & " (FormatNumber writes " & FormatNumber(1234.5, 2) & ")" & vbLf
	cases = Array(1234.5, CDbl(1234567.891), 150.555, -150.555, 2.675, 0.005, 0.004999, -0.001, CCur(40), CCur(0.125), CLng(7), CSng(0.1), 0, "100", "100.5", "0.125", "-3.14159", "1,234.56", "150,00", "", Empty, Null, True, 10 ^ 15, CDbl(999999999999.99))
	For i = 0 To UBound(cases)
		v = AmountText(cases(i))
		Response.Write i & " -> [" & v & "]" & vbLf
	Next
	Response.Write "escape -> " & PercentEncode("A 1+/" & ChrW(233) & ChrW(8217) & ChrW(55357) & ChrW(56832) & vbTab) & vbLf
Next
%>

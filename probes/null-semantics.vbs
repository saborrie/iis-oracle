' How VBScript treats Null, measured rather than derived.
'
' The question that prompted this (AxonASP issue #188): the & operator is
' documented to treat a lone Null operand as a zero-length string, while +
' propagates Null. Each case below prints the result's VarType/TypeName and
' value, or the error it raises, so an engine can be made to match the whole
' matrix and not just the headline case.

Option Explicit

Function Describe(expr)
	Dim v
	On Error Resume Next
	Err.Clear
	v = Eval(expr)
	If Err.Number <> 0 Then
		Describe = "ERR " & Err.Number & " (" & Err.Description & ")"
		Err.Clear
	ElseIf IsNull(v) Then
		Describe = "Null (vartype=" & VarType(v) & " typename=" & TypeName(v) & ")"
	ElseIf IsEmpty(v) Then
		Describe = "Empty (vartype=" & VarType(v) & " typename=" & TypeName(v) & ")"
	Else
		Describe = "[" & CStr(v) & "] (vartype=" & VarType(v) & " typename=" & TypeName(v) & ")"
	End If
	On Error GoTo 0
End Function

Dim cases, c
cases = Array( _
	"""a"" & Null", _
	"Null & ""a""", _
	"Null & Null", _
	"Null & Empty", _
	"Empty & Null", _
	"1 & Null", _
	"Null & 1", _
	"""a"" + Null", _
	"1 + Null", _
	"Null + Null", _
	"Null = Null", _
	"IsNull(""a"" & Null)", _
	"CStr(Null)", _
	"CBool(Null)", _
	"CInt(Null)", _
	"Len(Null)", _
	"Trim(Null)", _
	"UCase(Null)", _
	"Replace(Null, ""a"", ""b"")", _
	"Replace(""aa"", Null, ""b"")", _
	"Replace(""aa"", ""a"", Null)", _
	"Left(Null, 1)", _
	"Mid(Null, 1)", _
	"Right(Null, 1)", _
	"LCase(Null)", _
	"LTrim(Null)", _
	"RTrim(Null)", _
	"InStr(Null, ""a"")", _
	"InStr(""a"", Null)", _
	"InStr(1, Null, ""a"")", _
	"LenB(Null)", _
	"StrComp(Null, ""a"")", _
	"CLng(Null)", _
	"CDbl(Null)", _
	"CDate(Null)", _
	"""a"" & Null & ""b""", _
	"Len("""" & Null)" _
)
For Each c In cases
	WScript.Echo c & "  ->  " & Describe(c)
Next

' If Null Then is a statement, not an expression, so it cannot go through
' Eval. With On Error Resume Next an erroring If skips both branches.
Dim hit
On Error Resume Next
Err.Clear
hit = "no-branch-ran"
If Null Then hit = "true-branch" Else hit = "false-branch"
If Err.Number <> 0 Then
	WScript.Echo "If Null Then ... Else ...  ->  ERR " & Err.Number & " (" & Err.Description & "); " & hit
Else
	WScript.Echo "If Null Then ... Else ...  ->  " & hit
End If
On Error GoTo 0

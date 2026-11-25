<%@ Language=VBScript %>
<%
' Scripting.Dictionary edges, AscW/ChrW at the surrogate and sign edges, and
' what a caller under On Error Resume Next sees when a function it calls
' raises without a handler of its own.
Option Explicit
Response.ContentType = "text/plain"
Dim d, v, o, k, state

Sub Report(label)
	If Err.Number <> 0 Then
		Response.Write label & " -> ERR " & Err.Number & " (" & Err.Description & ")" & vbLf
	Else
		Response.Write label & " -> [" & v & "] " & TypeName(v) & vbLf
	End If
	Err.Clear
End Sub

Function Raises()
	Raises = "partial"
	state = "before raise"
	Err.Raise vbObjectError + 1, "probe", "raised on purpose"
	state = "after raise"
	Raises = "complete"
End Function

Function CallsRaises()
	CallsRaises = "outer partial"
	CallsRaises = Raises()
	state = state & ", outer after inner"
End Function

Function HandlesItsOwn()
	On Error Resume Next
	Err.Raise vbObjectError + 2, "probe", "inner trapped"
	HandlesItsOwn = "handled " & Err.Number
	Err.Clear
End Function

Function TrapsThenRaisesAfterGoto0()
	On Error Resume Next
	Dim x
	x = 1 / 0
	On Error GoTo 0
	TrapsThenRaisesAfterGoto0 = "assigned"
	Err.Raise vbObjectError + 3, "probe", "after goto 0"
	TrapsThenRaisesAfterGoto0 = "complete"
End Function

On Error Resume Next
Set d = CreateObject("Scripting.Dictionary")
d.Add "a", 1
d.Add "a", 2 : v = d("a") : Report "Dictionary Add duplicate"
v = d.Exists("A") : Report "Exists(""A"") after Add ""a"" (default CompareMode)"
v = d.CompareMode : Report "CompareMode"
v = d.Count : Report "Count before reading a missing key"
v = d("missing") : Report "read d(""missing"")"
v = d.Count : Report "Count after reading a missing key"
v = d.Exists("missing") : Report "Exists(""missing"") after reading it"
d.Add "x" & vbTab & "y", 3 : v = d("x" & vbTab & "y") : Report "key with a tab"
Set o = CreateObject("Scripting.Dictionary")
o.Add "inner", "yes"
Set d("obj") = o : v = d("obj")("inner") : Report "object stored with Set d(k) = o"
v = IsObject(d("obj")) : Report "IsObject(d(""obj""))"
v = UBound(d.Keys) : Report "UBound(d.Keys)"
v = d.Items()(0) : Report "d.Items()(0)"
d.Item("new") = "assigned" : v = d("new") : Report "d.Item(k) = v adds"
d.Add 1, "number one" : v = d.Exists("1") : Report "Exists(""1"") after Add 1"

v = AscW(ChrW(233)) : Report "AscW(ChrW(233))"
v = AscW(ChrW(55357)) : Report "AscW(ChrW(55357))"
v = AscW(ChrW(65535)) : Report "AscW(ChrW(65535))"
v = AscW(ChrW(-10179)) : Report "AscW(ChrW(-10179))"
v = Len(ChrW(55357) & ChrW(56832)) : Report "Len of a surrogate pair"
v = CLng("&HFFFF") : Report "CLng(""&HFFFF"")"
v = Mid("abc", 5, 1) : Report "Mid(""abc"", 5, 1)"
v = Mid("abc", 3) : Report "Mid(""abc"", 3)"
v = InStr(1, "0123456789abcdef", "F", vbBinaryCompare) : Report "InStr binary for ""F"""
v = InStr(1, "0123456789abcdef", "F", vbTextCompare) : Report "InStr text for ""F"""
v = InStr(5, "abc", "c") : Report "InStr(5, ""abc"", ""c"")"
v = Asc("") : Report "Asc("""")"
v = VarType(CCur(1)) & " " & VarType(1.5) & " " & VarType(1) & " " & VarType(100000) & " " & VarType(True) & " " & VarType("1") & " " & VarType(Empty) & " " & VarType(Null) & " " & VarType(Array()) : Report "VarTypes Cur Dbl Int Lng Bool Str Empty Null Array"
v = 2 ^ 53 : Report "2 ^ 53"
v = CDbl("9007199254740993") : Report "CDbl(""9007199254740993"")"
v = 10 ^ 2 : Report "10 ^ 2"

state = "unset" : v = "unset"
v = Raises() : Report "v = Raises() under the caller's OERN"
Response.Write "  v afterwards [" & v & "], state [" & state & "]" & vbLf
state = "unset" : v = "unset"
v = CallsRaises() : Report "v = CallsRaises() (two levels) under the caller's OERN"
Response.Write "  v afterwards [" & v & "], state [" & state & "]" & vbLf
v = HandlesItsOwn() : Report "HandlesItsOwn"
v = Err.Number : Report "caller's Err.Number after HandlesItsOwn"
v = "unset"
v = TrapsThenRaisesAfterGoto0() : Report "TrapsThenRaisesAfterGoto0 under caller's OERN"
Response.Write "  v afterwards [" & v & "]" & vbLf

Randomize
v = Rnd : Report "Rnd after Randomize"
v = Hex(Int(Rnd * 2147483647)) : Report "Hex(Int(Rnd * 2147483647))"
%>

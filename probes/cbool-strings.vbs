' How does the explicit CBool cast coerce a STRING? Companion to the
' if-condition-coercion.asp page: the If condition rejects " true " with
' type mismatch while accepting " 2 " - does CBool draw the same line?
Option Explicit

Function Describe(ByVal s)
    Dim v
    On Error Resume Next
    v = CBool(s)
    If Err.Number <> 0 Then
        Describe = "ERR " & Err.Number & " (" & Err.Description & ")"
    Else
        Describe = CStr(v)
    End If
    On Error GoTo 0
End Function

WScript.Echo "CBool(""0"") -> " & Describe("0")
WScript.Echo "CBool(""1"") -> " & Describe("1")
WScript.Echo "CBool(""-1"") -> " & Describe("-1")
WScript.Echo "CBool(""0.0"") -> " & Describe("0.0")
WScript.Echo "CBool("" 2 "") -> " & Describe(" 2 ")
WScript.Echo "CBool(""true"") -> " & Describe("true")
WScript.Echo "CBool(""False"") -> " & Describe("False")
WScript.Echo "CBool("" true "") -> " & Describe(" true ")
WScript.Echo "CBool(""abc"") -> " & Describe("abc")
WScript.Echo "CBool("""") -> " & Describe("")
WScript.Echo "CBool(""&h10"") -> " & Describe("&h10")
WScript.Echo "CBool(""&h0"") -> " & Describe("&h0")

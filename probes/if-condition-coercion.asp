<%@ Language=VBScript %>
<%
' How does an If condition coerce a STRING? (CBool rules, or something else?)
' Each case reports branch taken and Err.Number under On Error Resume Next.
Option Explicit

Function Cond(ByVal v)
    Dim taken
    On Error Resume Next
    If v Then
        taken = "true-branch"
    Else
        taken = "false-branch"
    End If
    If Err.Number <> 0 Then
        Cond = "ERR " & Err.Number & " (" & Err.Description & ")"
    Else
        Cond = taken
    End If
    On Error GoTo 0
End Function

Response.Write "If ""0"" Then -> " & Cond("0") & vbCrLf
Response.Write "If ""1"" Then -> " & Cond("1") & vbCrLf
Response.Write "If ""-1"" Then -> " & Cond("-1") & vbCrLf
Response.Write "If ""0.0"" Then -> " & Cond("0.0") & vbCrLf
Response.Write "If "" 2 "" Then -> " & Cond(" 2 ") & vbCrLf
Response.Write "If ""true"" Then -> " & Cond("true") & vbCrLf
Response.Write "If ""False"" Then -> " & Cond("False") & vbCrLf
Response.Write "If "" true "" Then -> " & Cond(" true ") & vbCrLf
Response.Write "If ""abc"" Then -> " & Cond("abc") & vbCrLf
Response.Write "If """" Then -> " & Cond("") & vbCrLf
Response.Write "If ""&h10"" Then -> " & Cond("&h10") & vbCrLf
Response.Write "If ""&h0"" Then -> " & Cond("&h0") & vbCrLf
%>

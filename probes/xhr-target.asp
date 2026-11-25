<%@ Language=VBScript %>
<%
' A target for the xhr-*.asp probes: waits ?wait= seconds, answers with
' ?status= (default 200), and echoes what it was sent. With ?utf8=1 it answers
' application/json (no charset) carrying UTF-8 bytes for U+00E9 and U+2019.
Option Explicit
Dim waitFor, t0, n, bytes, i, hexText
waitFor = 0
If IsNumeric(Request.QueryString("wait")) Then waitFor = CDbl(Request.QueryString("wait"))
t0 = Timer
Do While Timer - t0 < waitFor And Timer >= t0
Loop
If Request.QueryString("status") <> "" Then Response.Status = Request.QueryString("status") & " Probe"
If Request.QueryString("utf8") = "1" Then
	Response.ContentType = "application/json"
	Response.CodePage = 65001
	Response.Write "{""s"":""" & ChrW(233) & ChrW(8217) & """}"
	Response.End
End If
Response.ContentType = "text/plain"
Response.Write "method=" & Request.ServerVariables("REQUEST_METHOD") & vbLf
Response.Write "ua=" & Request.ServerVariables("HTTP_USER_AGENT") & vbLf
Response.Write "rid=" & Request.ServerVariables("HTTP_X_REQUEST_ID") & vbLf
Response.Write "ctype=" & Request.ServerVariables("CONTENT_TYPE") & vbLf
n = Request.TotalBytes
Response.Write "bytes=" & n & vbLf
If n > 0 Then
	bytes = Request.BinaryRead(n)
	hexText = ""
	For i = 1 To LenB(bytes)
		hexText = hexText & Right("0" & UCase(Hex(AscB(MidB(bytes, i, 1)))), 2)
	Next
	Response.Write "body-hex=" & hexText & vbLf
End If
%>

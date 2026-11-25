<%@ Language=VBScript %>
<%
' xhr-timeouts.asp again, against a target that waits without spinning (so no
' CPU is taken from the waiting caller), at 127.0.0.1 rather than localhost, each
' setting twice, and the clock's own resolution measured first.
Option Explicit
Response.ContentType = "text/plain"
Server.ScriptTimeout = 600
Dim slow, t0, t1, i
slow = "http://127.0.0.1/probe/xhr-target.asp?hold=12"

t0 = Timer
Do
	t1 = Timer
Loop While t1 = t0
Response.Write "Timer tick -> " & Round((t1 - t0) * 1000, 1) & " ms" & vbLf

Sub Measure(label, url, r, c, s, rv)
	Dim h, t0, e
	On Error Resume Next
	Set h = Server.CreateObject("MSXML2.ServerXMLHTTP")
	h.setTimeouts r, c, s, rv
	h.open "GET", url, False
	t0 = Timer
	h.send
	e = Err.Number
	Response.Write label & " setTimeouts(" & r & "," & c & "," & s & "," & rv & ") -> ERR " & e & ", " & Round(Timer - t0, 2) & " s" & vbLf
	Err.Clear
	Set h = Nothing
End Sub

For i = 1 To 2
	Measure "receive", slow, 5000, 5000, 5000, 500
	Measure "receive", slow, 5000, 5000, 5000, 1000
	Measure "receive", slow, 5000, 5000, 5000, 2000
	Measure "receive", slow, 5000, 5000, 5000, 3000
	Measure "receive", slow, 5000, 5000, 5000, 6000
	Measure "receive", slow, 2000, 2000, 2000, 1000
	Measure "receive", slow, 2000, 2000, 2000, 2000
Next
%>

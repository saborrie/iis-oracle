<%@ Language=VBScript %>
<%
' How long a synchronous ServerXMLHTTP send actually waits, for several
' setTimeouts(resolve, connect, send, receive) settings: against an answer that
' takes 10 s (receive), a port nothing listens on (refused), and an address
' that drops connection attempts (connect).
Option Explicit
Response.ContentType = "text/plain"
Server.ScriptTimeout = 600
Dim slow
slow = "http://localhost/probe/xhr-target.asp?wait=10"

Sub Measure(label, url, r, c, s, rv)
	Dim h, t0, e
	On Error Resume Next
	Set h = Server.CreateObject("MSXML2.ServerXMLHTTP")
	If r >= 0 Then h.setTimeouts r, c, s, rv
	h.open "GET", url, False
	t0 = Timer
	h.send
	e = Err.Number
	Response.Write label & " setTimeouts(" & r & "," & c & "," & s & "," & rv & ") -> ERR " & e & ", " & Round(Timer - t0, 2) & " s" & vbLf
	Err.Clear
	Set h = Nothing
End Sub

Measure "receive", slow, 5000, 5000, 5000, 500
Measure "receive", slow, 5000, 5000, 5000, 1000
Measure "receive", slow, 5000, 5000, 5000, 1500
Measure "receive", slow, 5000, 5000, 5000, 2000
Measure "receive", slow, 5000, 5000, 5000, 3000
Measure "receive", slow, 1000, 1000, 1000, 1000
Measure "receive", slow, 2000, 2000, 2000, 2000
Measure "refused", "http://127.0.0.1:1/", 5000, 5000, 5000, 5000
Measure "refused", "http://127.0.0.1:1/", 5000, 500, 5000, 5000
Measure "refused", "http://127.0.0.1:1/", 5000, 1000, 5000, 5000
Measure "refused", "http://127.0.0.1:1/", -1, 0, 0, 0
Measure "dropped", "http://10.255.255.1/", 5000, 1000, 5000, 5000
Measure "dropped", "http://10.255.255.1/", 5000, 2000, 5000, 5000
%>

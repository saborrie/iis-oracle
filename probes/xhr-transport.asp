<%@ Language=VBScript %>
<%
' MSXML2.ServerXMLHTTP from a page: what send raises when nothing listens, when
' the host does not resolve, and when the answer takes longer than setTimeouts
' allows; what .status and .responseText read afterwards; whether a .timeout
' property exists; non-2xx answers; request headers and a String body as the
' target receives them; and how responseText decodes UTF-8 JSON sent with no
' charset. Calls go to xhr-target.asp on this same server.
Option Explicit
Response.ContentType = "text/plain"
Server.ScriptTimeout = 300
Dim base, h, t0, v
base = "http://localhost/probe/xhr-target.asp"

Sub Line(label, value)
	Response.Write label & " -> " & value & vbLf
End Sub
Sub Report(label)
	If Err.Number <> 0 Then
		Line label, "ERR " & Err.Number & " (" & Err.Description & ")"
	Else
		Line label, "no error"
	End If
	Err.Clear
End Sub
Function Elapsed()
	Elapsed = Round(Timer - t0, 2)
End Function

On Error Resume Next

' A: nothing listening.
Set h = Server.CreateObject("MSXML2.ServerXMLHTTP")
h.setTimeouts 5000, 5000, 5000, 5000 : Report "A setTimeouts"
h.open "GET", "http://127.0.0.1:1/", False : Report "A open"
t0 = Timer
h.send : Report "A send (refused)"
Line "A elapsed", Elapsed()
v = h.status : Report "A read status" : Line "A status value", "[" & v & "] " & TypeName(v)
v = h.responseText : Report "A read responseText" : Line "A responseText value", "[" & v & "]"
v = h.readyState : Report "A read readyState" : Line "A readyState value", v
Set h = Nothing

' B: the answer takes 6 s, every timeout is 1 s.
Set h = Server.CreateObject("MSXML2.ServerXMLHTTP")
h.setTimeouts 1000, 1000, 1000, 1000 : Report "B setTimeouts"
h.open "GET", base & "?wait=6", False
t0 = Timer
h.send : Report "B send (slow answer)"
Line "B elapsed", Elapsed()
v = h.status : Report "B read status" : Line "B status value", "[" & v & "]"
Set h = Nothing

' C: setTimeouts 2000,2000,2000,2000 and a 3 s answer.
Set h = Server.CreateObject("MSXML2.ServerXMLHTTP")
h.setTimeouts 2000, 2000, 2000, 2000
h.open "GET", base & "?wait=3", False
t0 = Timer
h.send : Report "C send (3 s answer, 2 s receive)"
Line "C elapsed", Elapsed()
Set h = Nothing

' D: no setTimeouts at all, a 2 s answer (control: the default waits).
Set h = Server.CreateObject("MSXML2.ServerXMLHTTP")
h.open "GET", base & "?wait=2", False
t0 = Timer
h.send : Report "D send (defaults)"
Line "D elapsed", Elapsed()
Line "D status", h.status
Set h = Nothing

' E: a .timeout property.
Set h = Server.CreateObject("MSXML2.ServerXMLHTTP")
h.timeout = 2 : Report "E assign .timeout = 2"
v = h.timeout : Report "E read .timeout" : Line "E .timeout value", "[" & v & "]"
Set h = Nothing

' F: a host that does not resolve.
Set h = Server.CreateObject("MSXML2.ServerXMLHTTP")
h.setTimeouts 5000, 5000, 5000, 5000
h.open "GET", "http://no-such-host.invalid/", False
h.send : Report "F send (no such host)"
Set h = Nothing

' G: a 503 answer with a body.
Set h = Server.CreateObject("MSXML2.ServerXMLHTTP")
h.open "GET", base & "?status=503", False
h.send : Report "G send (503)"
Line "G status", h.status
Line "G responseText starts", Left(Replace(h.responseText, vbLf, "|"), 40)
Set h = Nothing

' H: request headers and a String body.
Set h = Server.CreateObject("MSXML2.ServerXMLHTTP")
h.open "POST", base, False
h.setRequestHeader "Content-Type", "application/json" : Report "H set Content-Type"
h.setRequestHeader "User-Agent", "probe-client/1.0" : Report "H set User-Agent"
h.setRequestHeader "X-Request-Id", "page-abc.7" : Report "H set X-Request-Id"
h.send "{""a"":""1.50""}" : Report "H send"
Line "H status", h.status
Line "H echo", Replace(h.responseText, vbLf, " | ")
Set h = Nothing

' I: UTF-8 JSON sent with no charset.
Set h = Server.CreateObject("MSXML2.ServerXMLHTTP")
h.open "GET", base & "?utf8=1", False
h.send : Report "I send"
v = h.responseText
Line "I length", Len(v)
Line "I AscW of chars 7 and 8", AscW(Mid(v, 7, 1)) & " " & AscW(Mid(v, 8, 1))
Line "I content-type", h.getResponseHeader("Content-Type")
Set h = Nothing

' J: open after a failed send on the same object, for a second attempt.
Set h = Server.CreateObject("MSXML2.ServerXMLHTTP")
h.open "GET", "http://127.0.0.1:1/", False
h.send : Report "J first send (refused)"
h.open "GET", base, False : Report "J reopen"
h.send : Report "J second send"
Line "J status", h.status
Set h = Nothing
%>

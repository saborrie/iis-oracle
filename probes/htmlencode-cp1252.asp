<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%><%
' Server.HTMLEncode on non-ASCII input under CODEPAGE=1252. Which characters
' become &#N; entities, and which pass through? Inputs are built with ChrW so
' the source file encoding is irrelevant; output is printed as code points so
' the response encoding is irrelevant too.
Option Explicit
Response.ContentType = "text/plain"

Sub Say(s)
	Response.Write s & vbLf
End Sub

Function Points(s)
	Dim i, r
	r = ""
	For i = 1 To Len(s)
		If i > 1 Then r = r & " "
		r = r & AscW(Mid(s, i, 1))
	Next
	Points = r
End Function

Dim cases, c, enc
' 39 apostrophe, 34 quote, 127 DEL, 128 ctrl, 150 ctrl, 160 nbsp, 163 pound,
' 169 copyright, 233 e-acute, 255 y-diaeresis, 256 A-macron, 8211 en dash,
' 8212 em dash, 8217 right quote, 8364 euro, 9731 snowman
cases = Array(39, 34, 127, 128, 150, 160, 163, 169, 233, 255, 256, 8211, 8212, 8217, 8364, 9731)

On Error Resume Next
Say "Session.CodePage=" & Session.CodePage & " Response.CodePage=" & Response.CodePage
For Each c In cases
	Err.Clear
	enc = Server.HTMLEncode("a" & ChrW(c) & "b")
	If Err.Number <> 0 Then
		Say "U+" & Hex(c) & " (" & c & ")  ->  ERR " & Err.Number & " (" & Err.Description & ")"
	Else
		Say "U+" & Hex(c) & " (" & c & ")  ->  [" & Points(enc) & "]"
	End If
Next
%>

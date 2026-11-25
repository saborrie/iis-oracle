<%@ Language=VBScript %>
<%
' The Set-Cookie header IIS emits for Response.Cookies(...).Expires assigned
' from a Date literal, DateAdd, Date/Now arithmetic and a date string.
' Fetched with headers (hdr- prefix). The body prints Now() and the server's
' local time so the Expires values can be checked against them.
Option Explicit
Response.ContentType = "text/plain"

Response.Cookies("lit") = "1"
Response.Cookies("lit").Expires = #2027-01-01#
Response.Cookies("littime") = "2"
Response.Cookies("littime").Expires = #2027-01-01 10:30:00#
Response.Cookies("dateadd") = "3"
Response.Cookies("dateadd").Expires = DateAdd("d", 1, Now())
Response.Cookies("nowplus") = "4"
Response.Cookies("nowplus").Expires = Now() + 1
Response.Cookies("dateminus") = "5"
Response.Cookies("dateminus").Expires = Date() - 1
Response.Cookies("str") = "6"
Response.Cookies("str").Expires = "January 1, 2027"

Response.Write "Now() = " & Now() & vbLf
%>

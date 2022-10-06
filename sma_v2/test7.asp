<%@ LANGUAGE=VBSCRIPT %>
<%
option explicit
on error resume next
%>

<!-- #include file="smaProcs.inc" -->

<%
dim strFrom, strTo, strCC, strBCC, strSubject, strBody

'set parameters
strFrom = "daniel_nica@telus.com"
strTo = "Daniel Nica <daniel_nica@telus.com>"
strCC = "daniel_nica@telus.com"
strBCC = ""
strSubject = "Testing CDO object"
strBody = "Here is the body of this message..."

'send email
SendEmail strFrom, strTo, strCC, strBCC, strSubject, strBody

if err then
	Response.Write err.description
end if
%>
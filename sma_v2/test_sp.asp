<%@ Language=VBScript %>
<!-- #include file="smaconstants.inc" -->
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
</HEAD>
<BODY>

<P>&nbsp;</P>
<%
on error resume next
dim objConn
set objConn = Server.CreateObject("ADODB.Connection")
objConn.ConnectionString = strConstConnectString
objConn.open
if err then 
	Response.Write "Error opening connection.<br>"
	Response.Write err.number & " - " & err.description
	Response.End
end if

dim cmdFirstTest
set cmdFirstTest = server.CreateObject("ADODB.Command")
set cmdFirstTest.ActiveConnection = objConn
cmdFirstTest.CommandType = adCmdStoredProc
cmdFirstTest.CommandText = "sma_userid.FIRST_TEST"
cmdFirstTest.Parameters.Append cmdFirstTest.CreateParameter("param1", adNumeric, adParamInput,,4)
if err then 
	Response.Write "Unexpected error.<br>"
	Response.Write err.number & " - " & err.description
	Response.End
end if
cmdFirstTest.Execute
if objConn.Errors.Count <> 0 then
	Response.Write "<b>Error executing command.</b><br>"
	Response.Write objConn.Errors(0).NativeError & " - " & objConn.Errors(0).Description
	objConn.Errors.Clear
	Response.End
else
	Response.Write "OK"
end if

%>
</BODY>
</HTML>

<%
'check if the user is logged on
'if Request.Cookies("UserAccessLevel")(strConst_Logon) = "" then
if Session(strConst_Logon)="" then
	response.redirect "default.asp?redir=Y"
end if



Dim objConn,StrConnectString
'StrConnectString = Request.Cookies("UserInformation")("ConnectString")
StrConnectString = Session("ConnectString")
set objConn = Server.CreateObject("ADODB.Connection")
objConn.ConnectionString = StrConnectString
objConn.open

'unexpected error, possible a database connection error
if err then
	DisplayError "BACK", "", err.number, "UNEXPECTED ERROR - Possible database connection error", err.description
end if
%>


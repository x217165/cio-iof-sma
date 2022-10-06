<%
'Dim smaConnKenan
'Set smaConnKenan = Server.CreateObject("ADODB.Connection")
'check if the user is logged on
'if Request.Cookies("UserAccessLevel")(strConst_Logon) = "" then
'	response.redirect "default.asp"
'end if

Dim objKenanConn,StrKenanConnectString
'StrKenanConnectString = "DRIVER={Microsoft ODBC for Oracle};UID=APP_SMA2;Password=APP_SMA2;server=KFXDV01"		'connection string
'StrKenanConnectString = "DRIVER={Oracle in OraClient11g_home};UID=readall;PWD=readall;Data Source=KFXDV"
'StrKenanConnectString = "DRIVER={Microsoft ODBC for Oracle};UID=readall;Password=readall;server=KFXRP"		'connection string
'StrKenanConnectString = "DRIVER={Microsoft ODBC for Oracle};UID=APP_SMA2;Password=SvcMgtAdm#2015;server=KFXPR01"
 StrKenanConnectString = "DRIVER={Microsoft ODBC for Oracle};UID=APP_SMA2;Password=SvcMgtAdm#2015;server=KFXPR"
'StrKenanConnectString = "DRIVER={Microsoft ODBC for Oracle};UID=readall;Password=readall;server=KFXPT"		'connection string
set objKenanConn = Server.CreateObject("ADODB.Connection")
objKenanConn.ConnectionString = StrKenanConnectString
objKenanConn.open


'unexpected error, possible a database connection error
if err then
	DisplayError "BACK", "", err.number, "UNEXPECTED ERROR - Possible database connection error", err.description
end if
'smaConnKenan.CursorLocation = 3

'const strKenanConstConnectString = "DRIVER={Microsoft ODBC for Oracle};UID=readall;Password=readall;server=KFXDV"		'connection string
'const ConnKenan.Provider = "MSDAORA;Data Source=KFXDV;Password=readall;User ID=readall"

'smaConnKenan.Open
%>



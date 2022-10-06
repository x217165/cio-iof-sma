<%@ Language=VBScript %>
<% Option Explicit 

Dim objConn,Struser,Strpassword,StrConnectstring
 on Error Resume Next
  Struser=request.form("txtusername")
  Strpassword=request.form("txtpassword")
  
'Setup connection object
 StrConnectstring = "DSN=orad5;uid=" & Struser & ";pwd=" & Strpassword
 'StrConnectstring = "DSN=orad5;uid=sma_userid;pwd=sma2" 
 set objConn = Server.CreateObject("ADODB.Connection")
 objConn.ConnectionString = StrConnectstring
 objConn.open
 
 if (objConn.Errors(0).NativeError <> 0) THEN
 Response.Redirect "loginfailed.htm"
 ELSE
 'Maintains state using cookies
  Response.Cookies("UserInformation")("username") = Struser 
  Response.Cookies("UserInformation")("password") = Strpassword
  Response.Cookies("UserInformation")("ConnectString") = StrConnectstring
  'Response.Cookies("UserInformation").Expires = date-1
  
  Response.Redirect "index.html"
 end if 
 
%>


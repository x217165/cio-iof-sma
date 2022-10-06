<%@ Language=VBScript %>
<% Option Explicit  %>
<% Response.Buffer = True%>
<!-- #include file="smaConstants.inc" -->
<!-- #include file="smaProcs.inc" -->
<%
On Error Resume Next
Dim strUserID, chkUID
Dim objConn, objRS, objEmail
Dim strUserName, strUserPass, strUserEmail
Dim strSQL, strError, lngLDAPerror
Dim intBusinessFunctionID, intAccessLevel

strUserName = Request.Form("txtUserName")
strUserPass = Request.Form("txtPassword")
strUserName="Tester2"
strUserPass = "testing"

Function CheckUser()
On Error Resume Next
	strSQL = "SELECT SEC.USERID, SEC.PASSWORD, CON.EMAIL_ADDRESS " &_
			"FROM MSACCESS.TBLSECURITY SEC, CRP.CONTACT CON " &_
			"WHERE CON.CONTACT_ID = SEC.STAFFID AND " &_
			"SEC.USERID = '" & strUserName & "' AND " &_
			"SEC.PASSWORD = '" & strUserPass & "'"
	
	Set objConn = Server.CreateObject("ADODB.Connection")
	objConn.ConnectionString = strConstConnectString
	objConn.Open
	If err Then
		DisplayError "BACK", "", err.Number, "Here. Cannot connect to database...", err.Description
		Response.End
	End If
	Set objEmail = Server.CreateObject("ADODB.Recordset")
	objEmail.Open strSQL, objConn, adOpenForwardOnly, adLockReadOnly, adCmdText
	If objEmail.EOF Then
		CheckUser = False
	Else
		If objEmail.Fields("EMAIL_ADDRESS").Value <> "" Then strUserEmail = objEmail.Fields("EMAIL_ADDRESS").Value
		CheckUser = True
	End If
	objEmail.Close
	Set objEmail = Nothing
	objConn.Close
	Set objConn = Nothing
End Function

If (strUserPass <> "") And (strUserName <> "") Then
	If Not CheckUser() Then
		DisplayError "BACK", "", 0, "Incorrect User ID, Password, or User not defined in SMA.", "Please contact your System Administrator."
		Response.End 	
	End If
	'authenticate the user
	'check if the user is defined TBLSECURITY
	strSQL = " SELECT nvl(bfa.access_level, 0) access_level" &_
		  " ,      b.business_function_id" &_
		  " FROM ( SELECT a.business_function_id" &_
		  "        ,      a.business_func_access_level_id" &_
		  "        ,      c.access_level" &_
		  "        FROM   msaccess.tblsecurity s" &_
		  "        ,      msaccess.staff_security_role r" &_
		  "        ,      msaccess.security_role t" &_
		  "        ,      msaccess.business_func_security_role a" &_
		  "        ,      msaccess.business_func_access_level c" &_
	 	  "        WHERE  s.userid = '" & routineOraString(strUserName) & "'" &_
		  "        AND    s.staffid = r.staff_id" &_
		  "        AND    r.security_role_id = t.security_role_id" &_
		  "        AND    t.security_role_name LIKE 'SMA%'" &_
		  "        AND    t.security_role_id = a.security_role_id" &_
		  "        AND    a.business_func_access_level_id = c.business_func_access_level_id" &_
		  "      ) bfa" &_
		  " ,      msaccess.business_function b" &_
		  " ,      msaccess.application a" &_
		  " WHERE  bfa.business_function_id (+)= b.business_function_id" &_
		  " AND    b.application_id = a.application_id" &_
		  " AND    a.application_name = 'SMA2' " &_
		  " ORDER BY business_function_id"
	
	Set objConn = Server.CreateObject("ADODB.Connection")
	objConn.ConnectionString = strConstConnectString
	objConn.Open
	If err Then
		DisplayError "BACK", "", err.Number, "Cannot connect to database...", err.Description
		Response.End
	End If

	Set objRS = Server.CreateObject("ADODB.Recordset")
	objRS.Open strSQL, objConn, adOpenForwardOnly, adLockReadOnly, adCmdText
	If err Then
		DisplayError "BACK", "", err.Number, "Cannot connect to database...", err.Description
		Response.End
	End If

	If objRS.EOF Then
		DisplayError "BACK", "", 0, "User not defined in SMA.", "Please contact your System Administrator."
		Response.Write "Logon failed. <a href='default.asp'>Try again.</a>"&"<br>"
		Response.End 
	Else
		'User OK, create cookies and go...
		Response.Cookies("UserInformation")("username") = strUserName 
'			Response.Cookies("UserInformation")("password") = strUserPass
		Response.Cookies("UserInformation")("ConnectString") = strConstConnectString
		'set access rights cookies
		Response.Cookies("UserAccessLevel")(strConst_Logon) = now

		intBusinessFunctionID = 0
		intAccessLevel = 0

		'email address required for sending email messages
		If Len(strUserEmail) <> 0 Then Response.Cookies("UserInformation")("email_address") = strUserEmail
			
		Do While Not objRS.EOF
			If (CLng(intBusinessFunctionID) <> 0) And (CLng(intBusinessFunctionID) <> CLng(objRS.Fields("BUSINESS_FUNCTION_ID").Value)) Then
				Response.Cookies("UserAccessLevel")(CStr(intBusinessFunctionID)) = intAccessLevel
				intAccessLevel = 0
			End If
			intBusinessFunctionID = objRS.Fields("BUSINESS_FUNCTION_ID").Value
			intAccessLevel = intAccessLevel Or objRS.Fields("ACCESS_LEVEL").Value
			objRS.MoveNext			
		Loop
		'store the last one
		If (CLng(intBusinessFunctionID) <> 0) Then
			Response.Cookies("UserAccessLevel")(CStr(intBusinessFunctionID)) = intAccessLevel
		End If
		objRS.close
		Set objRS = Nothing
		'trap unexpected error
		If err Then
			DisplayError "BACK", "", err.Number, "Unexpected error.", err.Description
			Response.End
		End If
		Response.Write "<SCRIPT type=""text/javascript"" language=""javascript"">"
		Response.write "window.open('index.asp','_blank','status=yes,hotkeys=no,toolbar=no,location=no,menubar=no,scrollbars=yes,resizable=yes');"
		Response.Write "document.location.href = ""loggedon.asp"";"
		Response.Write "</SCRIPT>"
		Response.End 
	End If
End If

%>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<META http-equiv="Pragma" content="no-cache">
<SCRIPT type="text/javascript" language="javascript">
	var strAppName = navigator.appName;
	//check browser
	if (strAppName != 'Microsoft Internet Explorer'){document.location="netscape.htm"}
	//check version
	var strVersion = navigator.appVersion;
	intVersion = parseFloat(strVersion.substring(strVersion.indexOf("MSIE")+5,strVersion.lastIndexOf(";")));
	if (intVersion < 5) {document.location="netscape.htm"}
</SCRIPT>
</HEAD>
<SCRIPT type="text/javascript" language="javascript" src="GeneralJavaFunctions.js"></SCRIPT>
<SCRIPT type="text/javascript" language="javascript">

//are cookies enabled on this pc?
SetCookie("test_cookie","the cookie is here");
if (GetCookie("test_cookie") != "the cookie is here") {document.location="enableCookies.htm";}

DeleteCookie("UserInformation");
DeleteCookie("UserAccessLevel");

function Page_onLoad()
{
  if (top.document != document) 
  {
  top.document.location.href = "default.asp"
  return
  }
  if (GetCookie("UserID") != "") {alert();document.frmUserLogon.txtUserName.value = GetCookie("UserID");}
  if (document.frmUserLogon.txtUserName.value == "")
  {
	document.frmUserLogon.txtUserName.focus();
	document.frmUserLogon.txtUserName.select();
  }
  else
  {
	document.frmUserLogon.txtPassword.focus();
	document.frmUserLogon.txtPassword.select();
  }
}

function Form_submit(theForm)
{
  if (theForm.txtUserName.value == "")
  {
    alert("Please enter a value for the \"User ID\" field.");
    theForm.txtUserName.focus();
    theForm.txtUserName.select();
    return (false);
  }

  if (theForm.txtUserName.value.length < 1)
  {
    alert("Please enter at least 1 characters in the \"User ID\" field.");
    theForm.txtUserName.focus();
    theForm.txtUserName.select();
    return (false);
  }

  if (theForm.txtUserName.value.length > 30)
  {
    alert("Please enter at most 30 characters in the \"User ID\" field.");
    theForm.txtUserName.focus();
    theForm.txtUserName.select();
    return (false);
  }
  if (theForm.txtPassword.value == "")
  {
    alert("Please enter a value for the \"Password\" field.");
    theForm.txtPassword.focus();
    theForm.txtPassword.select();
    return (false);
  }

  if (theForm.txtPassword.value.length < 1)
  {
    alert("Please enter at least 1 characters in the \"Password\" field.");
    theForm.txtPassword.focus();
    theForm.txtPassword.select();
    return (false);
  }

  if (theForm.txtPassword.value.length > 30)
  {
    alert("Please enter at most 30 characters in the \"Password\" field.");
    theForm.txtPassword.focus();
    theForm.txtPassword.select();
    return (false);
  }

  //cookie(s)
  SetCookie("UserID", document.frmUserLogon.txtUserName.value);
return (true)
}

function fct_changePassword() {
var strUserName = document.frmUserLogon.txtUserName.value;
var strUserPass = document.frmUserLogon.txtPassword.value;

	if (strUserName.length == 0) {
		alert('Please enter your User ID.');
		document.frmUserLogon.txtUserName.focus();
		return false;
	}
	SetCookie("UserName", strUserName);
	if (strUserPass.length != 0) {
		SetCookie("UserPass", strUserPass);
	}
	window.open("ChangePassword.asp", "Password", "height=200,width=400,left=300,top=300,status=no,toolbar=no,menubar=no,location=no")
	return true;
}
</SCRIPT>
<TITLE>SMA - Logon</TITLE>
<BODY onload="Page_onLoad()">
<TABLE border="0" cellPadding="0" cellSpacing="0" align="center" width="400">
	<TR bgcolor="#ffffff">
		<TD bgcolor="#ffffff" align="middle">
		<IMG src="Images/sma-logo-trans.gif" width="221" height="191"> </TD>
	</TR>
	<TR> 
		<TD colspan="3" align="middle"><STRONG><FONT size="6">Service Management Administration<BR></FONT></STRONG></TD>
	</TR>
</TABLE>

<FORM method="post" action="default.asp" name="frmUserLogon" id="frmUserLogon">
<TABLE align="center" border="0" cellPadding="1" cellSpacing="1" width="400">
	<TR>
		<TD align="right" width="50%"><STRONG>User ID:</STRONG></TD>
		<TD align="left" width="50%"><INPUT id="txtUserName" name="txtUserName" size="10" maxlength="10" value=""></TD>
	</TR>
	<TR>
		<TD align="right" width="50%"><STRONG>Password:</STRONG></TD>
		<TD align="left" width="50%"><INPUT id="txtPassword" name="txtPassword" size="10" maxLength="10" type="password" value="">
		&nbsp;<INPUT id="btnPassword" name="btnPassword" type="button" style="width: 2cm" value="Change" onClick="return fct_changePassword();"></TD>
	</TR>
	<TR>
		<TD align="right" width="50%">&nbsp;</TD>
		<TD align="left" width="50%"><INPUT id="btnSubmit" name="btnSubmit" type="submit" style="width: 2cm" value="Logon"></TD>
	</TR>
</TABLE>
</FORM>

<TABLE border="0" cellPadding="0" cellSpacing="0" align="center" width="400">
	<TR><TD bgcolor="white">&nbsp;</TD></TR>
	<TR bgcolor="#ffffff"><TD bgcolor="#32cd32">&nbsp;</TD></TR>
</TABLE>

</BODY>
</HTML>

<%@ Language=VBScript %>
<%  Option Explicit   %>
<!-- #include file="smaConstants.inc" -->
<!-- #include file="smaProcs.inc" -->
<%
on error resume next
dim strUserID, chkUID
dim rs, rsContact, rsEmail, objConn
dim strUserName, strUserPass
dim sql, strError, lngLDAPerror

function LDAP(strUserName, strUserPass)
on error resume next
dim con
	LDAP = false
	set con = server.CreateObject("ADODB.Connection")
	set rs = server.CreateObject("ADODB.Recordset")
	
	con.Provider = "ADsDSOObject"
	con.Open "Active Directory Provider", "uid=" & strUserName & ",ou=employee,o=BC TEL,c=CA", strUserPass
'	con.Open "Active Directory Provider", "uid=" & strUserName & ",ou=employee,o=TELUS,c=CA", strUserPass
	if con.Errors.Count <> 0 then exit function
	set rs = con.Execute("<LDAP://ldap.bctel.com/uid=" & strUserName & ",ou=employee,o=BC TEL,c=CA>;(objectClass=User);ADsPath;onelevel")
'	set rs = con.Execute("<LDAP://secureca.bctel.com/uid=" & strUserName & ",ou=employee,o=TELUS,c=CA>;(objectClass=User);ADsPath;onelevel")

	if con.Errors.Count = 0 then
		LDAP = true
	end if
end function

strUserName = Request.Form("txtUserName")
strUserPass = Request.Form("txtPassword")

if (strUserPass <> "") and (strUserName <> "") then
	'authenticate the user
	if LDAP(strUserName, strUserPass) then
		Response.Write "<center>Your LDAP authentication was succesfully completed. You will be able to logon to SMA2 when it becomes available.<br>Thank you for your interest!<br><br>You can now close this page."
		Response.end
	else
		Response.Write "Your LDAP authentication failed. You can <a href='checkLDAP.asp'>try again</a> or you can call SPOC and have your LDAP password reset."
		Response.end
	end if
end if

%>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<META http-equiv="Pragma" content="no-cache">
<script type="text/javascript">
	var strAppName = navigator.appName;
	//check browser
	if (strAppName != 'Microsoft Internet Explorer'){document.location="netscape.htm"}
	//check version
	var strVersion = navigator.appVersion;
	intVersion = parseFloat(strVersion.substring(strVersion.indexOf("MSIE")+5,strVersion.lastIndexOf(";")));
	if (intVersion < 5) {document.location="netscape.htm"}
</script>
</HEAD>
<SCRIPT type="text/javascript" language="javascript" src="GeneralJavaFunctions.js"></SCRIPT>
<script type="text/javascript">

//are cookies enabled on this pc?
SetCookie("test_cookie","the cookie is here");
if (GetCookie("test_cookie") != "the cookie is here") {document.location="enableCookies.htm";}

function Page_onLoad()
{
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

</script>

<title>SMA - Logon</title>

<body onload="Page_onLoad()">

<table border="0" cellPadding="0" cellSpacing="0" align="center" width="400" style="WIDTH: 400px">
 <tr bgcolor="#ffffff">
   	<td bgcolor="#ffffff" align="middle">
	  <IMG style="WIDTH: 152px; HEIGHT: 128px" height=191 src="Images/sma-logo-trans.gif" width=221 > 
    </td>
 </tr>
 <tr> 
    <td colspan="3" align="middle">
      <H2>Check your LDAP password</H2></td>
 </tr>
</table>

<form method="post" action="checkLDAP.asp" name="frmUserLogon" id="frmUserLogon">
<table align="center" border="0" cellPadding="1" cellSpacing="1" style="WIDTH: 400px" width="400">
    <tr>
        <td></td>
        <td style="WIDTH: 30%" width="30%">
            <div align="right" ><strong>User ID:</strong></div></td>
        <td style="WIDTH: 50%" width="50%"> 
             <input name="txtUserName" size="30" maxlength="40" value=""></td>
        <td></td></tr>
    <tr>
        <td></td>
        <td style="WIDTH: 30%" width="30%">
            <div align="right" ><strong>Password:</strong></div></td>
        <td style="WIDTH: 50%" width="50%"> 
            <input name="txtPassword" size="30" maxLength="40" type="password" value=""></td>
        <td></td></tr>
    <tr>
        <td></td>
        <td colSpan="2">
            <div align="center"><br>
              <input name="btnSubmit" type="submit" value="Logon"></div> 
        <td></td></tr>
    <tr>
        <td colSpan="4">
            The purpose of this webpage is to let you test your LDAP userid and password. <br>
            For
            the userid use your T###### number. Example:<br>&nbsp;&nbsp; -if your employee number is 805221, 
      enter T805221<BR>&nbsp;&nbsp; -if your employee number is 238, enter 
      T000238<BR>As for the password, use one of the following (in 
      order):<BR>&nbsp;&nbsp; -your SRT2 password (if you have 
      one)<BR>&nbsp;&nbsp; -your TELUS NT domain password<BR>&nbsp;&nbsp; -your 
      email password (if you have one)<BR>&nbsp;&nbsp; -call SPOC if you cannot 
      logon successfully</tr>
                        
</table>
</form>
<p>
<table border="0" cellPadding="0" cellSpacing="0" align="center" width="400">
 <tr>
   <td align=middle><font color=red size=2>&nbsp;</font></td>
 </tr>
 <tr bgcolor="#ffffff">
   	<td bgcolor="#32cd32" align="middle"> 
    &nbsp;
	</td>
 </tr>
</table></p>

</body>
</HTML>

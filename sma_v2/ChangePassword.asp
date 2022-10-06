<%@ Language=VBScript %>
<% Option Explicit  %>
<% Response.Buffer = True%>
<!-- #include file="smaConstants.inc" -->
<!-- #include file="smaProcs.inc" -->
<!--
*********************************************************************************************
* Page name:	ChangePassword.asp							   								*
* Purpose:		To allow user to change their logon password								*
*																							*
* Created by:																				*
*																							*
*********************************************************************************************
        Date		Author			Changes/enhancements made
        -----		------		------------------------------------------------------
       14-Aug-01	 DTy		Ensure password are no longer than 10 characters as
								  it gets truncated when save into msaccess.tblSecurity.
								Set window lenght for password to 10 characters.  Old was 8.
*********************************************************************************************
-->
<HTML>
<HEAD>
<TITLE>Password Change</TITLE>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<META http-equiv="Pragma" content="no-cache">
<LINK rel="stylesheet" type="text/css" href="stylesheets/styles.css">
<SCRIPT type="text/javascript" language="javascript" src="GeneralJavaFunctions.js"></SCRIPT>
<%
'On Error Resume Next
Dim strUserName, strUserPass, strUserPassNew
Dim objConn,objRS, strSQL, lRecordsAffected, strErrMessage

	set objConn = Server.CreateObject("ADODB.Connection")
	objConn.ConnectionString = strConstConnectString
	objConn.open

	strUserName = Request.Cookies("UserName")
	strUserPass = Request.Cookies("UserPass")
	
	If Request.Form("hdnFrmAction") = "SAVE" Then
		strUserName = LCase(Trim(Request.Form("txtUserName")))
		strUserPass = Trim(Request.Form("txtUserPass"))
		strUserPassNew = Trim(Request.Form("txtUserPassNew"))
		
		dim cmdUpdateObj
		set cmdUpdateObj = server.CreateObject("ADODB.Command")
		set cmdUpdateObj.ActiveConnection = objConn
		cmdUpdateObj.CommandType = adCmdStoredProc
		cmdUpdateObj.CommandText = "sma_sp_userid.spk_security_inter.sp_password_update"
		
		'create params 
		cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_update_userid", adVarChar, adParamInput, 20, strUserName)
		cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_userid", adVarChar, adParamInput, 20, strUserName)
		cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_old_password", adVarChar, adParamInput, 10, strUserPass)
		cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_new_password", adVarChar, adParamInput, 10, strUserPassNew)
		
		'call the update stored proc 
		on error resume next
		cmdUpdateObj.Execute
		on error goto 0
		
		if objConn.Errors.Count <> 0 then
			DisplayError "BACK", "", objConn.Errors(0).NativeError, "CANNOT UPDATE OBJECT", objConn.Errors(0).Description
			objConn.Errors.Clear
		else
			Response.Write "<SCRIPT type=""text/javascript"" language=""javascript"">" & vbCrLf
			Response.Write "alert('Password Changed Successfully!');" & vbCrLf
			Response.Write "window.close();" & vbCrLf
			Response.Write "</SCRIPT>" & vbCrLf
			Response.End
		End If
	End If
%>
<SCRIPT type="text/javascript" language="javascript">
<!-- Hide Script
function body_onLoad() {
var strUserName = document.frmChangePassword.txtUserName.value;
var strUserPass = document.frmChangePassword.txtUserPass.value;

	if (strUserName.length > 0) {
		document.frmChangePassword.txtUserPass.focus();
	}
	if (strUserPass.length > 0) {
		document.frmChangePassword.txtUserPassNew.focus();
	}
	DeleteCookie("UserName");
	DeleteCookie("UserPass");
}

function fct_Submit() {
var strNewPassword = document.frmChangePassword.txtUserPassNew.value;
var strConPassword = document.frmChangePassword.txtUserPassCon.value;

	if (strNewPassword.length > 10) {
		alert('Password is limited to 10 character maximum.  Please re-enter.');
		document.frmChangePassword.txtUserPassCon.focus();
		document.frmChangePassword.txtUserPassCon.select();
		return false;
	}

	if (strNewPassword != strConPassword) {
		alert('Confirm password does not match new password.  Please re-enter.');
		document.frmChangePassword.txtUserPassCon.focus();
		document.frmChangePassword.txtUserPassCon.select();
		return false;
	}
	document.frmChangePassword.hdnFrmAction.value = "SAVE";
	document.frmChangePassword.submit();
}

function fct_Cancel() {
	window.close();
}
-->
</SCRIPT>
</HEAD>
<BODY onLoad="body_onLoad();">
<FORM id="frmChangePassword" name="frmChangePassword" action="ChangePassword.asp" method="post">
	<INPUT type="hidden" id="hdnFrmAction" name="hdnFrmAction" value="">
<TABLE align="center" border="0" cellpadding="1" cellspacing="1" width="100%">
<THEAD>
	<TR>
		<TH align="left" colspan="2">Enter Old and New Password</TH>
	</TR>
</THEAD>
<TBODY>
	<TR>
		<TD align="right" width="35%"><STRONG>User ID:</STRONG></TD>
		<TD align="left" width="35%"><INPUT id="txtUserName" name="txtUserName" maxlength="10" size="10" readonly value="<%=strUserName%>"></TD>
	</TR>
	<TR>
		<TD align="right" width="35%"><STRONG>Old Password:</STRONG></TD>
		<TD align="left" width="35%"><INPUT type="password" id="txtUserPass" name="txtUserPass" maxlength="10" size="10" value="<%=strUserPass%>"></TD>
	</TR>
	<TR>
		<TD align="right" width="35%"><STRONG>New Password:</STRONG></TD>
		<TD align="left" width="35%"><INPUT type="password" id="txtUserPassNew" name="txtUserPassNew" maxlength="11" size="11" value=""></TD>
	</TR>
	<TR>
		<TD align="right" width="35%"><STRONG>Confirm Password:</STRONG></TD>
		<TD align="left" width="35%"><INPUT type="password" id="txtUserPassCon" name="txtUserPassCon" maxlength="10" size="10" value=""></TD>
	</TR>
</TBODY>
<TFOOT>
	<TR>
		<TH align="right" colspan="2">
		<INPUT type="button" value="Submit" id="btnSubmit" name="btnSubmit" onClick="return fct_Submit();">
		<INPUT type="button" value="Cancel" id="btnCancel" name="btnCancel" onClick="fct_Cancel();">
		</TH>
	</TR>
</TFOOT>
</TABLE>
</BODY>
</HTML>

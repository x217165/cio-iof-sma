<%@ LANGUAGE=VBSCRIPT   %>
<% option explicit      %>
<% on error resume next %>
<!-- #include file=smaConstants.inc -->
<!-- #include file=smaProcs.inc -->
<!-- #include file=databaseconnect.asp -->
<!--
********************************************************************************************
* Page name:	MakeCriteria.asp
* Purpose:		To dynamically set the criteria to search for an asset make.
*				Results are displayed via MakeList.asp
*
* In Param:		This page reads following cookies
*				UserID
*				LastName
*               FirstName
*
*
* Created by:	Chris Roe Oct. 31, 2000
*
********************************************************************************************
-->

<%
const COOKIE_USERID = "UserID"
const COOKIE_LASTNAME = "LastName"
const COOKIE_FIRSTNAME = "FirstName"
const LIST_PAGE   = "StaffRoleList.asp"
const DETAIL_PAGE = "StaffRoleDetail.asp"

'check user's rights
dim intAccessLevel
intAccessLevel = CInt(CheckLogon(strConst_Security))
if (intAccessLevel and intConst_Access_ReadOnly) <> intConst_Access_ReadOnly then
	DisplayError "BACK", "", 0, "ACCESS DENIED", "You don't have access to security. Please contact your system administrator."
end if

Dim sql
Dim objRS
sql = "SELECT s.security_role_id, s.security_role_name FROM msaccess.security_role s ORDER BY s.security_role_name"
set objRs=objConn.execute(sql)

%>
<HTML>
<HEAD>
<META HTTP-EQUIV="Pragma" CONTENT="no-cache">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<LINK href="stylesheets/styles.css" rel="stylesheet" type="text/css">
<script type="text/javascript" SRC="GeneralJavaFunctions.js"></script>
<script type="text/javascript" SRC="AccessLevels.js"></script>
<script type="text/javascript">
var intAccessLevel = <%=intAccessLevel%>;

//set section title
setPageTitle("SMA - Staff Role");

function fct_onLoad()
{
 	if (document.frmSearch.txtUserID.value != "" ||
		document.frmSearch.txtLastName.value != "" ||
		document.frmSearch.txtFirstName.value != "")
 	{
 		DeleteCookie("<%=COOKIE_USERID%>");
 		DeleteCookie("<%=COOKIE_LASTNAME%>");
 		DeleteCookie("<%=COOKIE_FIRSTNAME%>");
 		DeleteCookie("WinName");
 		document.frmSearch.submit();
 	}
}

function fct_clear()
{

	document.frmSearch.txtUserID.value = "";
	document.frmSearch.txtLastName.value = "";
	document.frmSearch.txtFirstName.value = "";
	document.frmSearch.selRole.seelctedIndex = 0;

}

function btnNew_onclick()
{
	if ((intAccessLevel & intConst_Access_Create) != intConst_Access_Create)
	{
		alert('Access denied. Please contact your system administrator.');
		return false;
	}

	parent.document.location.href ="<%=DETAIL_PAGE%>?NewRecord=NEW" ;

}

function validate(theForm){

	var bolConfirm ;

	if (isWhitespace(theForm.txtUserID.value) && isWhitespace(theForm.txtLastName.value) && isWhitespace(theForm.txtFirstName.value) && isWhitespace(theForm.txtUserID.value) && theForm.selRole.selectedIndex == 0)
	{
		bolConfirm = window.confirm("No Search Criteria have been entered. This search may take a long time..Continue?");
		if (bolConfirm)
		{
			return true;
		}
		else
		{
			return false;
		}

	}
	return true;
}
</script>

</HEAD>
<BODY onLoad="fct_onLoad();">
<form name="frmSearch" action="<%=LIST_PAGE%>" method="post" target="fraResult" onsubmit="return validate(this);">

<INPUT name="hdnWinName"  type="hidden" value="<%=Request.Cookies("WinName")%>">

<table border="0" width="100%">
	<thead>
		<tr>
			<td colspan=6>Staff Role Search</td>
		</tr>
	</thead>
	<tbody>
		<tr>
			<td width=15% align=right>Last Name</td>
			<td width=20% align=left><INPUT type="text" name="txtLastName" value="<%=Request.Cookies(COOKIE_LASTNAME)%>"></td>
			<td width=15%>&nbsp</td>
			<td width=15% align=right>User ID</td>
			<td width=20% align=left><INPUT type="text" name="txtUserID" value="<%=Request.Cookies(COOKIE_USERID)%>"></td>
			<td>&nbsp</td>
		</tr>
		<tr>
			<td width=15% align=right>First Name</td>
			<td width=20% align=left><INPUT type="text" name="txtFirstName" value="<%=Request.Cookies(COOKIE_FIRSTNAME)%>"></td>
			<td width=15%>&nbsp</td>
			<td width=15% align=right>Security Role</td>
			<td width=20% align=left>
				<SELECT name="selRole" style="WIDTH: 153px">
					<OPTION> </OPTION>
				<%while not objRS.EOF%>
					<OPTION value="<%=objRS("security_role_id")%>"><%=objRS("security_role_name")%></OPTION>
					<%objRS.MoveNext%>
				<%wend%>
				<% objRs.close
				   set objRS = Nothing

				   objConn.close
				   set objConn = Nothing
				%>

				</SELECT>
			</td>
			<td>&nbsp</td>
		</tr>
		<tr>
			<td width=15% align=right>Show Security Roles</td>
			<td width=20% align=left><INPUT type="checkbox" name="chkShowRoles"></td>
			<td width=15%>&nbsp</td>
			<td align="right">Active Only</td>
			<td><INPUT name="chkActiveOnly" type="checkbox" checked>
		</tr>
		<tr>
			<td align=right colspan="6">
				<INPUT name=btnClear type=button style="width: 2cm" value=Clear onClick="fct_clear()">&nbsp;&nbsp;
				<INPUT name=btnSubmit type=submit style="width: 2cm" value=Search>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			</td>
		</tr>
	</tbody>
</table>
</form>
</BODY>
</HTML>

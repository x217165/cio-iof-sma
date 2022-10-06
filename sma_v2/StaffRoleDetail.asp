<%@ Language=VBScript %>
<% Option Explicit %>
<% on error resume next %>
<% Response.Buffer = true %>
<!--#include file="SmaConstants.inc"-->
<!--#include file="databaseconnect.asp"-->
<!--#include file="smaProcs.inc"-->
<%
Const ASP_NAME = "StaffRoleDetail.asp"
Const IFRAME_FILE = "StaffSecurityRoleList.asp"
Const IFRAME_DETAIL_FILE = "StaffSecurityRoleDetail.asp"
Const NO_ID = "null"

'check user's rights
Dim intAccessLevel
Dim strNew
Dim strRealUserID
Dim strWinMessage

intAccessLevel = CInt(CheckLogon(strConst_Security))
strRealUserID = Session("username")

if (intAccessLevel and intConst_Access_ReadOnly) <> intConst_Access_ReadOnly then
	DisplayError "BACK", "", 0, "ACCESS DENIED", "You don't have access to security. Please contact your system administrator"
end if

Dim strContactID, StrSql, objRS

strContactID = Request("hdnContactID")
strNew =Request("NewRecord")

if strContactID = "" THEN
	strContactID = NO_ID
END IF

if strContactID <> NO_ID then

	StrSql = " SELECT c.contact_name" &_
			 " ,      c.last_name" &_
			 " ,      c.first_name" &_
			 " ,      c.contact_id" &_
			 " ,      c.userid" &_
			 " FROM   crp.contact c" &_
			 " WHERE  c.contact_id = " & strContactID

	'Create Recordset object
	set objRs = objConn.Execute(StrSql)

	if err then
		DisplayError "BACK", "", err.Number, "CANNOT CREATE RECORDSET", err.Description
	end if
end if

%>
<HTML>
<HEAD>
<META HTTP-EQUIV="Pragma" CONTENT="no-cache">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<LINK REL="stylesheet" TYPE="text/css" HREF="stylesheets/styles.css">
<TITLE></TITLE>
<script type="text/javascript" SRC="GeneralJavaFunctions.js"></script>
<script type="text/javascript" SRC="AccessLevels.js"></script>
<SCRIPT LANGUAGE=javascript>
<!--
var intAccessLevel = <%=intAccessLevel%>;

function iFrame_display()
{
//called whenever a refresh of the iFrame is needed
	document.frames("aifr").document.location.href = '<%=IFRAME_FILE%>?ContactID=' + '<%=strContactID%>';
}

function btn_iFrmAdd()
{

	if ((intAccessLevel & intConst_Access_Create) != intConst_Access_Create)
	{
		alert('Access denied.  Please contact your system administrator.');
		return false;
	}

	var NewWin;
	NewWin=window.open("<%=IFRAME_DETAIL_FILE%>?NewContact=NEW&ContactID=<%=strContactID%>", "NewWin", "toolbar=no,status=no,width=700,height=260,menubar=no resize=no");
	NewWin.focus();
}

function btn_iFrmUpdate()
{


	var NewWin;

	if ((intAccessLevel & intConst_Access_Update) != intConst_Access_Update)
	{
		alert('Access denied.  Please contact your system administrator.');
		return false;
	}

	if (document.frames("aifr").frmIFR.hdnStaffSecurityRoleID.value !="")
	{

		var strSource ="<%=IFRAME_DETAIL_FILE%>?ContactID=<%=strContactID%>&StaffSecurityRoleID="+document.frames("aifr").frmIFR.hdnStaffSecurityRoleID.value;
		NewWin=window.open(strSource ,"NewWin","toolbar=no,status=no,width=700,left=100,top=200,height=260,menubar=no,resize=no");
		NewWin.focus();
	}
	else
	{
		alert('You must select a record to update!');
	}
}

function btn_iFrmDelete()
{

	if ((intAccessLevel & intConst_Access_Delete) != intConst_Access_Delete)
	{
		alert('Access denied.  Please contact your system administrator.');
		return false;
	}

	if (document.frames("aifr").frmIFR.hdnStaffSecurityRoleID.value !="")
	{
		if (confirm('Do you really want to delete this role from this employee?'))
		{
			document.frames("aifr").document.location.href = "<%=IFRAME_FILE%>?txtFrmAction=DELETE&hdnStaffSecurityRoleID=" + document.frames("aifr").frmIFR.hdnStaffSecurityRoleID.value + "&ContactID=<%=strContactID%>" + "&hdnUpdateDateTime=" + document.frames("aifr").frmIFR.hdnUpdateDateTime.value;
		}
	}
	else
	{
		 alert('You must select a record to delete.');
	}
}

//-->
</SCRIPT>
</HEAD>
<BODY onload="return iFrame_display();">
<FORM NAME=frmDetail METHOD=POST ACTION="<%=ASP_NAME%>">
	<INPUT id=hdnContactID      name=hdnContactID      type=hidden value="<%if strContactID <> NO_ID then Response.Write strContactID else Response.Write null end if%>">
	<table width="100%" border=0>
		<thead>
			<TR>
				<TD align=left colspan=2>Staff Role Assignment</TD>
			</tr>
		</thead>
		<tbody>
			<TR>
				<TD ALIGN=RIGHT NOWRAP>Contact Name</TD>
				<TD><INPUT name=txtName disabled size="50" value="<%if strContactID <> NO_ID then  Response.Write routineHtmlString(objRS("contact_name")) else Response.Write null end if%>"></TD>
			</TR>
			<TR>
				<TD ALIGN=RIGHT NOWRAP>First Name</TD>
				<TD><INPUT name=txtFirstName disabled size="20" value="<%if strContactID <> NO_ID then  Response.Write routineHtmlString(objRS("first_name")) else Response.Write null end if%>"></TD>
			</TR>
			<TR>
				<TD ALIGN=RIGHT NOWRAP>Last Name</TD>
				<TD><INPUT name=txtLastName disabled size="20" value="<%if strContactID <> NO_ID then  Response.Write routineHtmlString(objRS("last_name")) else Response.Write null end if%>"></TD>
			</TR>
			<TR>
				<TD ALIGN=RIGHT NOWRAP>User ID</TD>
				<TD><INPUT name=txtLastName disabled size="20" value="<%if strContactID <> NO_ID then  Response.Write routineHtmlString(objRS("userid")) else Response.Write null end if%>"></TD>
			</TR>
			<tr>
				<td>&nbsp;</td>
			</tr>
			<table>
				<thead>
					<TR>
						<td colspan=4 align=left>Staff Security Roles</td>
					</TR>
				</thead>
				<tbody>
					<tr>
						<TD colspan=4>
							<iframe id="aifr" width="100%" height="350" src="" scrolling="yes" marginheight="1" marginwidth="1"></iframe>
						</TD>
					</tr>
					<tr>
						<td>
							<input type="button" value="Delete"  <%if strContactID <> NO_ID then  Response.Write null else Response.Write "DISABLED" end if%>   name="btn_iFrameDelete"  onClick="btn_iFrmDelete();" style="width: 2cm">&nbsp;&nbsp;
							<input type="button" value="Refresh" <%if strContactID <> NO_ID then  Response.Write null else Response.Write "DISABLED" end if%>   name="btn_iFrameRefresh" onClick="iFrame_display();" style="width: 2cm">&nbsp;&nbsp;
							<input type="button" value="New"     <%if strContactID <> NO_ID then  Response.Write null else Response.Write "DISABLED" end if%>   name="btn_iFrameAdd"     onClick="btn_iFrmAdd(); "   style="width: 2cm">&nbsp;&nbsp;
							<input type="button" value="Update"  <%if strContactID <> NO_ID then  Response.Write null else Response.Write "DISABLED" end if%>   name="btn_iFrameupdate"  onClick="btn_iFrmUpdate();" style="width: 2cm">&nbsp;&nbsp;
						</TD>
					</TR>
				</tbody>
			</TABLE>

		</tbody>
	</table>
</FORM>
<%

 'Clean up our ADO objects
if strContactID <> NO_ID then
    objRS.close
    set objRS = Nothing
end if

%>

</BODY>
</HTML>

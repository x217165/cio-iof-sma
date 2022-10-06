<%@ Language=VBScript %>
<% Option Explicit %>
<% 'on error resume next%>
<!-- #include file="smaConstants.inc" -->
<!-- #include file="smaProcs.inc" -->
<!-- #include file="databaseconnect.asp" -->
<!--
******************************************************
* FILE:		StaffSecurityRole.asp
*
* PURPOSE:	Displays all current security roles for a 
*			specific user in StaffRoleDetail.asp
*
* PARAMS:	This page accepts the following variables
*			(either as URL param, or Form varaibles):
*				ContactID - the user to display security roles for
*				hdnhdnStaffSecurityRoleID - specifies a record to delete
*               txtFrmAction - set to "DELETE" if the page is to delete a specific role
*				hdnUpdateDateTime - the date and time that this record was last updated
*
******************************************************
-->
<%
'Get Circuit Id?
Const ASP_NAME = "StaffSecurityRole.asp" 'only need to change this value when changing the filename

dim strContactID, strhdnStaffSecurityRoleID, objRsServiceContact, StrSql
Dim intAccessLevel

intAccessLevel = CInt(CheckLogon(strConst_Security))


strContactID = Request("ContactID")
strhdnStaffSecurityRoleID = Request("hdnStaffSecurityRoleID")

 select case Request("txtFrmAction")
	case "DELETE" 
	 
		if (strhdnStaffSecurityRoleID <> "") then
		    if (intAccessLevel and intConst_Access_Delete) <> intConst_Access_Delete then
				DisplayError "BACK", "", 0, "DELETE DENIED", "You don't have access to delete Service Location Contacts. Please contact your system administrator."
			end if
			
			dim cmdDeleteObj
			set cmdDeleteObj = server.CreateObject("ADODB.Command")
			set cmdDeleteObj.ActiveConnection = objConn

			cmdDeleteObj.CommandType = adCmdStoredProc
			cmdDeleteObj.CommandText = "sma_sp_userid.spk_security_inter.sp_staff_security_role_delete"

			cmdDeleteObj.Parameters.Append cmdDeleteObj.CreateParameter("p_staff_security_role_id", adNumeric, adParamInput, , clng(strhdnStaffSecurityRoleID))					'number(9)	
			cmdDeleteObj.Parameters.Append cmdDeleteObj.CreateParameter("p_last_update", adDBTimeStamp, adParamInput, ,Cdate(Request("hdnUpdateDateTime")))		'Date
		
			cmdDeleteObj.Execute
			if objConn.Errors.Count <> 0 then
				DisplayError "BACK", "", objConn.Errors(0).NativeError, "CANNOT DELETE OBJECT", objConn.Errors(0).Description
				objConn.Errors.Clear
			end if
		end if	
 end select 			
			
if strContactID <> "" then
	strSQL = " SELECT ssr.staff_security_role_id" &_
			 " ,      r.security_role_name" &_
			 " ,      ssr.comments" &_
			 " ,      ssr.update_date_time" &_
			 " FROM   msaccess.staff_security_role ssr" &_
			 " ,      msaccess.security_role r" &_
			 " WHERE  r.security_role_id = ssr.security_role_id" &_
			 " AND    ssr.staff_id = " & strContactID &_
			 " ORDER BY r.security_role_name"

	set objRsServiceContact = objConn.Execute(StrSql)
end if

%>

<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<LINK href="stylesheets/styles.css" rel="stylesheet" type="text/css">
<STYLE>
.regularItem {
	cursor: hand;
}
.whiteItem {
	cursor: hand;
	background-color: white;
}
.Highlight {
	cursor: hand; 
	background-color: #00974f;
	color: white;
}
</STYLE>

<script type="text/javascript">

var oldHighlightedElement;
var oldHighlightedElementClassName;

function cell_onClick(dtUpdate, intServLocID)
{
	
	document.frmIFR.hdnStaffSecurityRoleID.value = intServLocID;
	document.frmIFR.hdnUpdateDateTime.value = dtUpdate; 
		
	//highlight current record
	if (oldHighlightedElement != null) 
	{
		oldHighlightedElement.className = oldHighlightedElementClassName
	}
	oldHighlightedElement = window.event.srcElement.parentElement;
	oldHighlightedElementClassName = oldHighlightedElement.className;
	oldHighlightedElement.className = "Highlight";
}

</script>

</HEAD>
<BODY>
<form name="frmIFR" action="<%=ASP_NAME%>" method="POST">
<input type="hidden" name="hdnStaffSecurityRoleID" value="">
<input type="hidden" name="hdnUpdateDateTime" value="">

<TABLE border=1 cellspacing=0 cellpadding=2 width="100%">
	<thead>
		<th nowrap>Role</th>
		<th nowrap>Comments</th>
	</thead>
	<tbody>
		<%
		dim k
		k = 0
		if strContactID <> "" then
			while not objRsServiceContact.EOF
				if Int(k/2) = k/2 then
					Response.Write "<tr class=""regularItem"">"
				else
					Response.Write "<tr class=""whiteItem"">"
				end if
				k = k+1
			%>
					<td width=50% nowrap onClick="cell_onClick('<%=routineJavaScriptString(objRsServiceContact("UPDATE_DATE_TIME"))%>',<%=objRsServiceContact("STAFF_SECURITY_ROLE_ID")%>);"><%=routineHTMLString(objRsServiceContact("SECURITY_ROLE_NAME"))%>&nbsp;</td>
					<td width=50% nowrap onClick="cell_onClick('<%=routineJavaScriptString(objRsServiceContact("UPDATE_DATE_TIME"))%>',<%=objRsServiceContact("STAFF_SECURITY_ROLE_ID")%>);"><%=routineHTMLString(objRsServiceContact("COMMENTS"))%>&nbsp;</td>
				</tr>
			<%
				objRsServiceContact.MoveNext
			wend
			objRsServiceContact.Close
			set objRsServiceContact = Nothing
		end if
		%>
	</tbody>
</table>
</FORM>
</BODY>
</HTML>
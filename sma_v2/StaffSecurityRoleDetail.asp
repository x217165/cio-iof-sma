<%@ Language=VBScript %>
<% Option Explicit %>
<% Response.Buffer = true %>
<% On Error Resume Next %>
<!--#include file="smaConstants.inc"-->
<!--#include file="databaseconnect.asp"-->
<!--#include file="smaProcs.inc"-->

<!-- This is the child detail screen for the StaffRole screen.
     It can have the following values passed into it:

     	Parameter				Details
     	-------------------------------------------------------------------------
    	StaffSecurityRoleID		the ID from the database of the service location contact
    							required for updates and deletes

    	ContactID				the ID of the Staff Memeber from the parent screen
    							required for creates

    	NewRole					must have a value of 'NEW'
	    						required for new records

    	hdnUpdateDateTime		the updateDateTime from the database
								required for updates and deletes
-->


<%
Const ASP_NAME = "StaffSecurityRoleDetail.asp"
Const NO_ID    = "null"

dim intAccessLevel

intAccessLevel = CInt(CheckLogon(strConst_Security))
if intAccessLevel < intConst_Access_ReadOnly then
	DisplayError "BACK", "", 0, "ACCESS DENIED", "You don't have access to Staff Security Roles. Please contact your system administrator."
end if


Dim strRealUserID
strRealUserID = Session("username")


Dim strStaffSecurityRoleID, strContactID, strCustomerName, strNewRole, strWinMessage, strUpdDate
Dim strSql
Dim objRS, objRsSecurityRole, objCmd
Dim strContactInfo

strStaffSecurityRoleID = Request("StaffSecurityRoleID")
if strStaffSecurityRoleID = "" then
	strStaffSecurityRoleID = NO_ID
end if

strContactID = Request("ContactID")

strNewRole = Request("NewRole")
strUpdDate = Request("hdnUpdateDateTime")


if strNewRole = "NEW" then
	strStaffSecurityRoleID = NO_ID
end if

select case Request("hdnFrmAction")
	case "SAVE"
		if (strStaffSecurityRoleID <> NO_ID) then

			if (intAccessLevel and intConst_Access_Update) <> intConst_Access_Update then
				DisplayError "BACK", "", 0, "UPDATE DENIED", "You don't have access to update Staff Security Roles. Please contact your system administrator."
			end if

			dim cmdUpdateObj
			set cmdUpdateObj = server.CreateObject("ADODB.Command")
			set cmdUpdateObj.ActiveConnection = objConn
			cmdUpdateObj.CommandType = adCmdStoredProc
			cmdUpdateObj.CommandText = "sma_sp_userid.spk_security_inter.sp_staff_security_role_update"

			'create params
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_userid", adVarChar, adParamInput, 20, strRealUserID)
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_staff_security_role_id", adNumeric, adParamInput, 20, strStaffSecurityRoleID)
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_staff_id", adNumeric, adParamInput, , Clng(strContactID))
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_security_role_id", adVarChar, adParamInput, 20, Clng(Request("selRole")))
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_last_update", adDBTimeStamp, adParamInput, ,Cdate(Request("hdnUpdateDateTime")))		'Date

			if Request("txtComments") <> "" then
				cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_comments", adVarChar, adParamInput, 255, Request("txtComments"))
			else
				cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_comments", adVarChar, adParamInput, 255, null)
			end if

			'call the stored proc
			cmdUpdateObj.Execute
			if objConn.Errors.Count <> 0 then
				DisplayError "BACK", "", objConn.Errors(0).NativeError, "CANNOT UPDATE OBJECT", objConn.Errors(0).Description
				objConn.Errors.Clear
			end if

			strStaffSecurityRoleID = Request("StaffSecurityRoleID")

			strWinMessage = "Record saved successfully. You can now see the changes you made."
		else
			'create a new record
			if (intAccessLevel and intConst_Access_Create) <> intConst_Access_Create then
				DisplayError "BACK", "", 0, "CREATE DENIED", "You don't have access to create Staff Security Roles. Please contact your system administrator."
			end if
			dim cmdInsertObj
			set cmdInsertObj = server.CreateObject("ADODB.Command")
			set cmdInsertObj.ActiveConnection = objConn
			cmdInsertObj.CommandType = adCmdStoredProc
			cmdInsertObj.CommandText = "sma_sp_userid.spk_Security_inter.sp_staff_security_role_insert"

			'create params
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_userid", adVarChar, adParamInput, 20, strRealUserID)
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_staff_security_role_id", adNumeric, adParamOutput, 20, null)
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_staff_id", adNumeric, adParamInput, , Clng(strContactID))
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_security_role_id", adVarChar, adParamInput, 20, Clng(Request("selRole")))

			if Request("txtComments") <> "" then
				cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_comments", adVarChar, adParamInput, 255, Request("txtComments"))
			else
				cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_comments", adVarChar, adParamInput, 255, null)
			end if

			'call the stored proc
			cmdInsertObj.Execute
			if objConn.Errors.Count <> 0 then
				DisplayError "BACK", "", objConn.Errors(0).NativeError, "CANNOT INSERT OBJECT", objConn.Errors(0).Description
				objConn.Errors.Clear
			else
				strStaffSecurityRoleID = cmdInsertObj.Parameters("p_staff_security_role_id").Value
			end if



			strWinMessage = "Record created successfully. You can now see the new record."
		end if

	case "DELETE"
		'delete record
		if (intAccessLevel and intConst_Access_Delete) <> intConst_Access_Delete then
			DisplayError "BACK", "", 0, "DELETE DENIED", "You don't have access to delete Staff Security Roles. Please contact your system administrator."
		end if

		dim cmdDeleteObj
		set cmdDeleteObj = server.CreateObject("ADODB.Command")
		set cmdDeleteObj.ActiveConnection = objConn
		cmdDeleteObj.CommandType = adCmdStoredProc
		cmdDeleteObj.CommandText = "sma_sp_userid.spk_security_inter.sp_staff_security_role_delete"
		cmdDeleteObj.Parameters.Append cmdDeleteObj.CreateParameter("p_staff_security_role_id", adNumeric, adParamInput, , clng(strStaffSecurityRoleID))					'number(9)
		cmdDeleteObj.Parameters.Append cmdDeleteObj.CreateParameter("p_last_update", adDBTimeStamp, adParamInput, ,Cdate(Request("hdnUpdateDateTime")))		'Date

		cmdDeleteObj.Execute
		if objConn.Errors.Count <> 0 then
			DisplayError "BACK", "", objConn.Errors(0).NativeError, "CANNOT DELETE OBJECT", objConn.Errors(0).Description
			objConn.Errors.Clear
		end if

		strStaffSecurityRoleID = NO_ID

		StrWinMessage = "Record deleted successfully."

end select

if strStaffSecurityRoleID <> NO_ID then
	'load the page data if required
	strSQL = " SELECT ssr.staff_security_role_id" &_
			 " ,      r.security_role_id" &_
			 " ,      ssr.record_status_ind" &_
			 " ,      ssr.comments" &_
			 " ,      ssr.staff_id" &_
			 " ,      to_char(ssr.create_date_time, 'MON-DD-YYYY HH24:MI:SS') create_date" &_
			 " ,      sma_sp_userid.spk_sma_library.sf_get_full_username(ssr.create_real_userid) create_real_userid" &_
			 " ,      to_char(ssr.update_date_time, 'MON-DD-YYYY HH24:MI:SS') update_date" &_
			 " ,      ssr.update_date_time" &_
			 " ,      sma_sp_userid.spk_sma_library.sf_get_full_username(ssr.update_real_userid) update_real_userid" &_
			 " FROM   msaccess.staff_security_role ssr" &_
			 " ,      msaccess.security_role r" &_
			 " WHERE  r.security_role_id = ssr.security_role_id" &_
			 " AND    ssr.staff_security_role_id = " & strStaffSecurityRoleID


	'Create Recordset object
	set objRs = objConn.execute(strSQL)
end if

'always load the Security Roles
strsql = " SELECT security_role_id" &_
		 " ,      security_role_name" &_
		 " ,      comments" &_
		 " FROM   msaccess.security_role" &_
		 " ORDER  BY security_role_name"

'Create the command object
set objCmd = Server.CreateObject("ADODB.command")
    objCmd.ActiveConnection = objconn
    objCmd.CommandText = strSql
    objCmd.CommandType = adCmdText

'Create Recordset object
set objRsSecurityRole = objCmd.Execute

%>

<HTML>
<HEAD>
<META HTTP-EQUIV="Pragma" CONTENT="no-cache">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<LINK REL="stylesheet" TYPE="text/css" HREF="stylesheets/styles.css">
<script type="text/javascript" SRC="GeneralJavaFunctions.js"></script>
<script type="text/javascript" SRC="AccessLevels.js"></script>
<TITLE>Staff Security Role Detail</TITLE>
<SCRIPT LANGUAGE=javascript>
<!--
var intAccessLevel = <%=intAccessLevel%>;
var boolNeedToSave = false;

function btnNew_click()
{
	if ((intAccessLevel & intConst_Access_Create) != intConst_Access_Create)
	{
		alert('Access Denied. Please contact your system administrator.');
		return;
	}

	self.document.location.href = "<%=ASP_NAME%>?NewRole=NEW&ContactID=<%=strContactID%>";
}

function fct_onDelete()
{
	if ((intAccessLevel & intConst_Access_Delete) != intConst_Access_Delete)
	{
		alert('Access Denied. Please contact your system administrator.');
		return;
	}

	if (confirm('Do you really want to delete this object?'))
	{
		boolNeedToSave = false;
		document.location = "<%=ASP_NAME%>?hdnFrmAction=DELETE&StaffSecurityRoleID=<%=strStaffSecurityRoleID%>&ContactID=<%=strContactID%>&hdnUpdateDateTime="+document.frmDetail.hdnUpdateDateTime.value;
	}
}

function btnClose_onclick()
{
	window.close();
	parent.opener.iFrame_display();
}

//-->
</SCRIPT>


<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--
function body_onBeforeUnload()
{
	//must set focus to save button because if user has changed only one field and has not
	//left it the on_change event will not have fired and the flag that determines whether
	//you need to save or not will be false
	document.frmDetail.btnSave.focus();

	//before you do the code below check that the 'need to save' flag is true and check
	//the user's access for either insert or update depending on which is
	//appropriate (i.e. for most of us this means whether the main id = 0 for a new record
	//or a value < or > 0 for an existing record)
	if ((intAccessLevel & intConst_Access_Update) == intConst_Access_Update || (intAccessLevel & intConst_Access_Create) == intConst_Access_Create)
	{
		if (boolNeedToSave == true)
		{
			event.returnValue = "There is unsaved data on the screen. To save changes, click CANCEL below then click SAVE on the main form.";
		}
	}
}

function frmDetail_onsubmit()
{
	if ((intAccessLevel & intConst_Access_Update) != intConst_Access_Update || (intAccessLevel & intConst_Access_Create) != intConst_Access_Create)
	{
		alert('Access Denied. Please contact your system administrator.');
		return false;
	}

	if (document.frmDetail.txtComments.value.length > 255)
	{
		alert("You are only allowed to enter a maximum of 255 characters for a comment. You have entered " + document.frmDetail.txtComments.value.length + ".");
		document.frmDetail.txtComments.focus();
		return false;
	}

	document.frmDetail.hdnFrmAction.value = "SAVE";
	boolNeedToSave = false;
	document.forms[0].submit();
	return true;
}

function btnContactLookup_onClick()
{

	if (document.frmDetail.txtWorkFor.value != "")
	{
		SetCookie("WorkFor", document.frmDetail.txtWorkFor.value);
	}
	if (document.frmDetail.txtLName.value != "")
	{
		 SetCookie("LName", document.frmDetail.txtLName.value);
	}
	if (document.frmDetail.txtFName.value != "")
	{
		 SetCookie("FName", document.frmDetail.txtFName.value);
	}

	SetCookie("WinName", 'Popup');
	window.open('SearchFrame.asp?fraSrc=Contact', 'Popup', 'top=50, left=100, height=600, width=800' ) ;
	on_change();
}

function on_change()
{
	boolNeedToSave = true;
}
//-->
</SCRIPT>
</HEAD>
<BODY onload="parent.opener.iFrame_display();"onbeforeunload="body_onBeforeUnload();">

<FORM name=frmDetail ACTION="<%=ASP_NAME%>" LANGUAGE=javascript>
<INPUT name="hdnFrmAction"        type="hidden" value="">
<INPUT name="hdnUpdateDateTime"   type="hidden" value="<%if strStaffSecurityRoleID <> NO_ID then  Response.Write objRS("UPDATE_DATE_TIME")       else Response.Write null  end if%>" >
<INPUT name="ContactID"           type="hidden" value="<%if strStaffSecurityRoleID <> NO_ID then  Response.Write objRS("STAFF_ID")               else Response.Write strContactID  end if%>" >
<INPUT name="StaffSecurityRoleID" type="hidden" value="<%if strStaffSecurityRoleID <> NO_ID then  Response.Write objRS("STAFF_SECURITY_ROLE_ID") else Response.Write null  end if%>" >

<TABLE border=0 width=100%>
	<thead>
		<TR ><TD colspan=2>Staff Security Role Detail</td></tr>
	</thead>
	<tbody>
	<TR>
		<TD ALIGN=RIGHT width=20% NOWRAP>Security Role<font color=red>*</font></TD>
		<TD width=80%>
			<SELECT id=selRole name=SelRole style="HEIGHT: 20px; WIDTH: 380px">
			<%
			Dim roleComment	'used to set the intial value of txtRoleComment
			Do while Not objRsSecurityRole.EOF
				Response.write "<OPTION "
				if roleComment = "" then
					roleComment = objRsSecurityRole("COMMENTS")
				end if
				if strStaffSecurityRoleID <> NO_ID then
					if cLng(objRsSecurityRole("SECURITY_ROLE_ID")) = clng(objRS("SECURITY_ROLE_ID")) then
						Response.Write " selected "
						roleComment = objRsSecurityRole("COMMENTS")
					end if
				end if
				Response.Write " VALUE =""" & routineHTMLString(objRsSecurityRole("SECURITY_ROLE_ID")) & """>" & routineHTMLString(objRsSecurityRole("SECURITY_ROLE_NAME")) & "</OPTION>"
				objRsSecurityRole.MoveNext
			Loop
			%>
			</SELECT>
		</TD>
	</TR>
	<TR>
		<TD ALIGN=RIGHT width=20% NOWRAP>Comments</TD>
		<TD width=80%>
			<TEXTAREA id=txtComments name=txtComments style="WIDTH: 380px; HEIGHT: 75px"><%if strStaffSecurityRoleID <> NO_ID then  Response.Write objRS("COMMENTS") else Response.Write null  end if%></TEXTAREA>
		</TD>
	</TR>
	</tbody>
</TABLE>

<TABLE>
	  <TR><TD align=right colspan=5>
			<INPUT id=btnClose  name=btnClose  type=button value=Close  style="WIDTH: 2cm" onclick="return btnClose_onclick();">&nbsp;&nbsp;
			<INPUT id=btnDelete name=btnDelete type=button value=Delete style="WIDTH: 2cm" onclick="return fct_onDelete();">&nbsp;&nbsp;
			<INPUT id=btnReset  name=btnReset  type=reset  value=Reset  style="WIDTH: 2cm" >&nbsp;&nbsp;
			<INPUT id=btnAddNew name=btnAddNew type=button value=New    style="WIDTH: 2cm" onclick="return btnNew_click();">&nbsp;&nbsp;
			<INPUT id=btnSave   name=btnSave   type=button value=Save   style="WIDTH: 2cm" onclick="return frmDetail_onsubmit();">&nbsp;&nbsp;
	  </TD></TR>
</table>

<FIELDSET >
	<LEGEND ALIGN=RIGHT><B>Audit Information</B></LEGEND>
	<Div SIZE=8pt ALIGN=RIGHT>
		Record Status Indicator:
		<INPUT align=left name=txtRecordStatusInd type=text style="HEIGHT: 20px; WIDTH: 18px"disabled value=<%if  strStaffSecurityRoleID <> NO_ID then  Response.Write """"&objRS("RECORD_STATUS_IND")&"""" else Response.Write """""" end if%>  >&nbsp;&nbsp;&nbsp;
		Create Date:&nbsp;&nbsp;
		<INPUT align=center name=txtcrdate type=text style="HEIGHT: 20px; WIDTH: 150px"disabled value=<%if  strStaffSecurityRoleID <> NO_ID then  Response.Write """"&objRS("CREATE_DATE")&"""" else Response.Write """""" end if%> >&nbsp;
		&nbsp;
		Created By:
		<INPUT align=right name=txtcrby type=text style="HEIGHT: 20px; WIDTH: 100px"disabled value=<%if  strStaffSecurityRoleID <> NO_ID then  Response.Write """"&objRS("CREATE_REAL_USERID")&"""" else Response.Write """""" end if%> ><BR>
		Update Date:
		<INPUT align= center name=txtupdate type=text style="HEIGHT: 20px; WIDTH: 150px"disabled value=<%if  strStaffSecurityRoleID <> NO_ID then  Response.Write """"&objRS("UPDATE_DATE")&"""" else Response.Write """""" end if%>  >&nbsp;&nbsp;
		Updated By:
		<INPUT align=right name=txtupby type=text style="HEIGHT: 20px; WIDTH: 100px"disabled value=<%if  strStaffSecurityRoleID <> NO_ID then  Response.Write """"&objRS("UPDATE_REAL_USERID")&"""" else Response.Write """""" end if%>  >
	</DIV>
</FIELDSET>

</FORM>
<%

if strStaffSecurityRoleID <> NO_ID then

	'Clean up our ADO objects if they were opened
	objRS.close
	set objRS = Nothing

	objRsSecurityRole.close
	set objRsSecurityRole = Nothing

	objConn.close
	set ObjConn = Nothing

end if

%>


</BODY>
</HTML>

<%@ Language=VBScript %>
<%  Option Explicit
 on error resume next
 Response.Buffer = true %>
<!--#include file = "smaConstants.inc"-->
<!--#include file = "smaProcs.inc"-->
<!--#include file = "databaseconnect.asp" -->
<!--
*********************************************************************************************
* Page name:	SAttUsageDetail.asp															*
* Purpose:		To display the service attribute and its associated values					*
* Created by:	Linda Chen																	*
* Date:			August 2009																	*
*********************************************************************************************
-->
<%
Dim strAttID, strAttvID, strWinMessage, strWinLocation
Dim strtxtAttID, strtxtAttvID
Dim	 strErrMessage
Dim intAccessLevel
Dim strSQL, objRsAtt, objRsAttv, objRsAttUsage
Dim struserid


	'Check user's rights
	intAccessLevel = CInt(CheckLogon(strConst_ServiceCatalogue))
	If (intAccessLevel And intConst_Access_ReadOnly) <> intConst_Access_ReadOnly Then
		DisplayError "BACK", "", 0, "ACCESS DENIED", "You don't have access to Line of Business. Please contact your system administrator"
	End If

	strWinMessage = ""
	strAttID = Request("hdnAttID")
	strAttvID = Request("hdnAttvID")

'	response.write("strAttID is " & strAttID & "strAttvID is " & strAttvID)
'	response.end
	struserid = Session("username")

	if (strAttID <> "" and strAttvID <> "") then
		strSQL	= 	"SELECT CREATE_DATE_TIME, " &_
					"CREATE_REAL_USERID, " &_
					"UPDATE_DATE_TIME, " &_
					"UPDATE_REAL_USERID " &_
					"FROM CRP.SRVC_TYPE_ATT_VAL_USAGE " &_
					"WHERE SRVC_TYPE_ATT_ID = " & strAttID &_
					" AND SRVC_TYPE_ATT_VAL_ID=" & strAttvID &_
					" AND RECORD_STATUS_IND ='A' "
	'response.write strSQL
	'response.end
	    set objRsAttUsage = objConn.Execute(strSQL)
	end if

	strSQL	= 	"SELECT SRVC_TYPE_ATT_NAME, " &_
				"SRVC_TYPE_ATT_ID " &_
				"FROM CRP.SRVC_TYPE_ATT " &_
				"WHERE RECORD_STATUS_IND = 'A' "&_
				" ORDER BY SRVC_TYPE_ATT_NAME "
	'response.write strSQL
	'response.end
    set objRsAtt = objConn.Execute(strSQL)

   	strSQL	= 	"SELECT SRVC_TYPE_ATT_VAL_NAME, " &_
				"SRVC_TYPE_ATT_VAL_ID " &_
				"FROM CRP.SRVC_TYPE_ATT_VAL " &_
				"WHERE RECORD_STATUS_IND = 'A' " &_
				" ORDER BY SRVC_TYPE_ATT_VAL_NAME "
	'response.write strSQL
	'response.end
    set objRsAttv = objConn.Execute(strSQL)


	Select Case UCase(Request("hdnFrmAction"))
		Case "SAVE"
		 	strtxtAttID = request("txtAttID")
			strtxtAttvID = request("txtAttvID")
			If Request("hdnAttID") <> 0 Then	'Update existing Service Type Attribute Usage
				If (intAccessLevel And intConst_Access_Update) <> intConst_Access_Update Then
					DisplayError "REFRESH", strWinLocation, 0, "UPDATE DENIED", "You don't have access to update Service Instance Attribute. Please contact your system administrator"
				End If

				dim cmdUpdateObj
				set cmdUpdateObj = server.CreateObject("ADODB.Command")
				set cmdUpdateObj.ActiveConnection = objConn
				cmdUpdateObj.CommandType = adCmdStoredProc
				cmdUpdateObj.CommandText = "SMA_SP_USERID.Sp_Srvtype_Att_Rule_Update"

				cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_userid", adVarChar , adParamInput, 20, struserid)
				cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_srvc_type_att_id",adNumeric , adParamInput,, Clng(strAttId))
				cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_srvc_type_att_val_id",adNumeric , adParamInput,, Clng(strAttvId))
				cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_newsrvc_type_att_id",adNumeric , adParamInput,, Clng(strtxtAttID))
				cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_newsrvc_type_att_val_id",adNumeric , adParamInput,, Clng(strtxtAttvID))


				'****************************
				'check parameter values
  				'****************************
   				'dim objparm
  				'for each objparm in cmdUpdateObj.Parameters
  				'	  Response.Write "<b>" & objparm.name & "</b>"
  				'	  Response.Write " has size:  " & objparm.Size & " "
  				'	  Response.Write " and value:  " & objparm.value & " "
  				'	  Response.Write " and datatype:  " & objparm.type & "<br> "
  				'next

  				'Response.Write "<b> count = " & cmdUpdateObj.Parameters.count & "<br>"
  				'dim nx
  				'for nx=0 to cmdUpdateObj.Parameters.count-1
  				'   Response.Write nx+1 & " parm value= " & cmdUpdateObj.Parameters.Item(nx).Value  & "<br>"
  				'next

  				'response.write (cmdUpdateObj.CommandText)
				'response.end

				response.write "Calling cmdUpdateObj.CommandText with p_srvc_type_att_id = " & strAttId
				response.write " <BR> and p_srvc_type_att_val_id= " & strAttvId
				response.write " <BR> and p_newsrvc_type_att_id = " & strtxtAttID
				response.write " <BR> and p_newsrvc_type_att_val_id = " & strtxtAttvID
				response.end

				cmdUpdateObj.Execute
				if objConn.Errors.Count <> 0 then
					DisplayError "BACK", "", objConn.Errors(0).NativeError, "CANNOT UPDATE RECORD", objConn.Errors(0).Description
					objConn.Errors.Clear
					response.redirect("SAttUsageDetail.asp")
				end if

			Else									'Create a new Service tye Attribute Usage
				'If (intAccessLevel And intConst_Access_Create) <> intConst_Access_Create Then
				'	DisplayError "REFRESH", strWinLocation, 0, "INSERT DENIED", "You don't have access to create Service Type Attribute Usage. Please contact your system administrator"
				'End If
				dim cmdInsertObj
				set cmdInsertObj = server.CreateObject("ADODB.Command")
				set cmdInsertObj.ActiveConnection = objConn
				cmdInsertObj.CommandType = adCmdStoredProc
				cmdInsertObj.CommandText = "SMA_SP_USERID.Sp_Srvtype_Att_Rule_Insert"

				cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_userid", adVarChar , adParamInput, 20, struserid)
 				cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_srvc_type_att_id",adNumeric , adParamInput,, Clng(strtxtAttID))
				cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_srvc_type_att_val_id",adNumeric , adParamInput,, Clng(strtxtAttvID))

				'****************************
				'check parameter values
  				'****************************

  				'dim objparm
  				'for each objparm in cmdInsertObj.Parameters
  				'	  Response.Write "<b>" & objparm.name & "</b>"
  				'	  Response.Write " has size:  " & objparm.Size & " "
  				'	  Response.Write " and value:  " & objparm.value & " "
  				'	  Response.Write " and datatype:  " & objparm.type & "<br> "
  				'next

  				'Response.Write "<b> count = " & cmdInsertObj.Parameters.count & "<br>"
  				'dim nx
  				'for nx=0 to cmdInsertObj.Parameters.count-1
  				'   Response.Write nx+1 & " parm value= " & cmdInsertObj.Parameters.Item(nx).Value  & "<br>"
  				'next

  				'response.write (cmdInsertObj.CommandText)
				'response.end
				response.write "Calling "
				response.write cmdInsertObj.CommandText
				response.write " with p_srvc_type_att_id = " & strtxtAttID + " and p_srvc_type_att_val_id= " & strtxtAttvID
				response.end
				cmdInsertObj.Execute
				if objConn.Errors.Count <> 0 then
					DisplayError "BACK", "", objConn.Errors(0).NativeError, "CANNOT ADD NEW RECORD", objConn.Errors(0).Description
					objConn.Errors.Clear
					response.redirect("SAttUsageDetail.asp")
				end if
			End If
		Case "DELETE"
				dim cmdDeleteObj
				set cmdDeleteObj = server.CreateObject("ADODB.Command")
				set cmdDeleteObj.ActiveConnection = objConn
				cmdDeleteObj.CommandType = adCmdStoredProc
				cmdDeleteObj.CommandText = "SMA_SP_USERID.Sp_Srvtype_Att_Rule_Delete"

				cmdDeleteObj.Parameters.Append cmdDeleteObj.CreateParameter("p_userid", adVarChar , adParamInput, 20, struserid)
 				cmdDeleteObj.Parameters.Append cmdDeleteObj.CreateParameter("p_srvc_type_att_id",adNumeric , adParamInput,, Clng(strAttID))
				cmdDeleteObj.Parameters.Append cmdDeleteObj.CreateParameter("p_srvc_type_att_val_id",adNumeric , adParamInput,, Clng(strAttvID))

				'****************************
				'check parameter values
  				'****************************

  				'dim objparm
  				'for each objparm in cmdDeleteObj.Parameters
  				'	  Response.Write "<b>" & objparm.name & "</b>"
  				'	  Response.Write " has size:  " & objparm.Size & " "
  				'	  Response.Write " and value:  " & objparm.value & " "
  				'	  Response.Write " and datatype:  " & objparm.type & "<br> "
  				'next

  				'Response.Write "<b> count = " & cmdDeleteObj.Parameters.count & "<br>"
  				'dim nx
  				'for nx=0 to cmdDeleteObj.Parameters.count-1
  				'   Response.Write nx+1 & " parm value= " & cmdDeleteObj.Parameters.Item(nx).Value  & "<br>"
  				'next

  				'response.write (cmdDeleteObj.CommandText)
				'response.end

				response.write "Calling "
				response.write cmdDeleteObj.CommandText
				response.write " with p_srvc_type_att_id = " & strAttID
				response.write " and p_srvc_type_att_val_id= " & strAttvID
				response.end


				cmdDeleteObj.Execute
				if objConn.Errors.Count <> 0 then
					DisplayError "BACK", "", objConn.Errors(0).NativeError, "CANNOT DELETE RECORD", objConn.Errors(0).Description
					objConn.Errors.Clear
					response.redirect("SAttUsageDetail.asp")
				else
					response.redirect("STypeAttUsage3.asp")
				end if
  		End Select
%>
<HTML>
<HEAD>
<META name="Generator" content="Microsoft Visual Studio 6.0">
<META http-equiv="Pragma" content="no-cache">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<LINK REL="stylesheet" TYPE="text/css" HREF="stylesheets/styles.css">
<SCRIPT type="text/javascript" language="javascript" src="AccessLevels.js"></SCRIPT>
<SCRIPT type="text/javascript" language="javascript" src="GeneralJavaFunctions.js"></SCRIPT>
<SCRIPT type="text/javascript" language="javascript" id="clientEventHandlersJS">
<!-- //Hide Client-Side SCRIPT
var strWinMessage = "<%=strWinMessage%>";
var intAccessLevel = <%=intAccessLevel%>;
var	bolSaveRequired = false;

setPageTitle("SMA - Service Attribute Instance");

function btnClose_onClick() {
	window.close();
	parent.opener.iSTAttuFrame_display();
}

function btnDelete_onClick() {
//**********************************************************************************************
// Function:	btnDelete_onClick
// Purpose:		To delete a line of Service Attribute/Value Usage
// Created By:	Linda Chen 07/21/2009
// Updated By:
//***********************************************************************************************
// Remove the comment in the 4 lines after test LC
//	if ((intAccessLevel & intConst_Access_Delete) != intConst_Access_Delete) {
//		alert('You do not have permission to DELETE a Service Instance Attribute.  Please contact your System Administrator.');
//		return false;
//	}

	if (document.frmAttUsageDetail.txtAttName.value == "") {
		alert('This Service Type Attribute does not exist in the database.');
		return false;
	}

	if (confirm('Do you really want to delete this object?')) {
		document.frmAttUsageDetail.hdnFrmAction.value = "DELETE";
		bolSaveRequired = false;
		document.frmAttUsageDetail.submit();
	}
//	document.location.href='STypeAttUsage.asp?';
}

//**********************************************************************************************
// Function:	btnSave_onClick
// Purpose:		To Save the added or updated Service Attribute/Value Usage
// Created By:	Linda Chen 07/21/2009
// Updated By:
//***********************************************************************************************

function btnSave_onClick() {
	if ((intAccessLevel & intConst_Access_Update) != intConst_Access_Update) {
		alert('You do not have permission to UPDATE a Service Type Attribute Usage.  Please contact your System Administrator.');
		return false;
	}

	if (document.frmAttUsageDetail.txtAttID.value == 0) {
		alert('Please enter the Service Type Attribute');
		document.frmAttUsageDetail.txtAttID.focus();
		return false;
	}

	if (document.frmAttUsageDetail.txtAttvID.value == 0) {
		alert('Please enter the Service Type Attribute Value');
		document.frmAttUsageDetail.txtAttvID.focus();
		return false;
	}

	document.frmAttUsageDetail.hdnFrmAction.value = "SAVE";
	bolSaveRequired = false;
	document.frmAttUsageDetail.submit();
//	window.close();
//	parent.opener.iSTAttuFrame_display();
}

function window_onBeforeUnload() {
	//Ensure that fct_onChange() fires for any changed data.
	document.frmAttUsageDetail.btnSave.focus();

	if (bolSaveRequired) {
		event.returnValue = "There is unsaved data on the screen. To save changes, click CANCEL below then click SAVE on the main FORM.";
	}
}

function window_onUnload() {
//
}

function ClearStatus() {
	window.status = "";
}

function DisplayStatus(strWinStatus) {
	window.status = strWinStatus;
	setTimeout('ClearStatus()', 5000);
}

function btnReset_onClick() {
	if(confirm('All change will be lost. Do you really want to reset the page?')){
		bolSaveRequired = false;
		document.location.href = "SAttUsageDetail.asp?hdnAttID=<%=strAttID%>";
	}
}
// Unhide Client-Side SCRIPT -->
</SCRIPT>
</HEAD>
<BODY language="javascript" onLoad="DisplayStatus(strWinMessage);" onBeforeUnload="window_onBeforeUnload();" onUnload="window_onUnload();">
<FORM id="frmAttUsageDetail" name="frmAttUsageDetail" action="SAttUsageDetail.asp" method="post">
	<input id="hdnFrmAction" name="hdnFrmAction" type=hidden>
	<INPUT id="hdnAttID" name="hdnAttID" type=hidden
	value="<%If strAttID <> "" Then Response.Write strAttID else response.write 0 end if %>">
	<INPUT id="hdnAttvID" name="hdnAttvID" type=hidden
	value="<%If strAttvID <> "" Then Response.Write strAttvID else response.write 0 end if %>">

<TABLE border="0" cols="4" width="100%">
<THEAD>
<TR>
	<TD align="left" colspan="3">Service Attribute Usage</TD>
	<TD align="right" width="2%">&nbsp;</TR>
</THEAD>
<TBODY>
<TR>
	<TD align="left" nowrap width="21%">&nbsp;Attribute Name<FONT color="red">*</FONT></TD>
	<TD ><SELECT id=txtAttID name=txtAttID>
				<OPTION value=0 ></OPTION>
				<% objRsAtt.movefirst
				Do while Not objRsAtt.EOF %>
		   		<option  <% if strAttID <> "" then
		   				if clng(strAttID) = clng(objRsAtt(1)) then
		              		response.write "selected "
		              	end if
		              end if %>
           		value = <% =objRsAtt(1) %>
		  		 > <% =objRsAtt(0)%> </option>
				<%  objRsAtt.MoveNext
				Loop %>
				</SELECT>
			</TD>

</TR size="28">
<TR>
	<TD align="left" nowrap width="21%">&nbsp;Attribute Value</TD>
	<TD align="left" colspan="2" nowrap>
	<SELECT id=txtAttvID name=txtAttvID>
				<OPTION value=0 ></OPTION>
				<% objRsAttv.movefirst
				Do while Not objRsAttv.EOF %>
		   		<option  <% if strAttvID <> "" then
		   				if clng(strAttvID) = clng(objRsAttv(1)) then
		              		response.write "selected "
		              	end if
		              end if %>
           		value = <% =objRsAttv(1) %>
		  		 > <% =objRsAttv(0)%> </option>
				<%  objRsAttv.MoveNext
				Loop %>
				</SELECT>
</TD>
</TR size="41">
</TR>
<TFOOT>
<TR>
	<TD colspan="4" align="right">
	<INPUT id="btnClose" name="btnClose" type="button" value="Close" style="width: 2cm" language="javascript" onClick="btnClose_onClick();">
	&nbsp;
	<INPUT id="btnDelete" name="btnDelete" type="hidden" value="Delete" style="width: 2cm" language="javascript" onClick="btnDelete_onClick();">&nbsp;

	<INPUT id="btnReset" name="btnReset" type="button" value="Reset" style="width: 2cm" language="javascript" onClick="btnReset_onClick();" >&nbsp;
	&nbsp;
	<INPUT id="btnSave" name="btnSave" type="button" value="Save" style="width: 2cm" onClick="return btnSave_onClick();">&nbsp;
</TR>
</TFOOT>
</TABLE>

<FIELDSET width="100%">
	<LEGEND align="right"><B>Audit Information</B></LEGEND>
	<DIV size="8pt" align="right">
	Record Status Indicator:<INPUT align="left" name="txtRecordStatusInd" type="text" style="width: 18px"
	disabled value="<%If strAttID <> 0 and  strAttvID <> 0 Then Response.Write "A" end if %>">&nbsp;&nbsp;&nbsp;
	Create Date:<INPUT align="center" name="txtRecordStatusInd" type="text" style="width: 150px"
	disabled value="<%If strAttID <> 0 and  strAttvID <> 0  Then Response.Write objRsAttUsage(0).value%>">&nbsp;
	Created By:<INPUT align="right" name="txtRecordStatusInd" type="text" style="width: 150px"
	disabled value="<%If strAttID <> 0 and  strAttvID <> 0 Then Response.Write objRsAttUsage(1).value%>"><BR>
	Update Date:<INPUT align="center" name="txtRecordStatusInd" type="text" style="width: 150px"
	disabled value="<%If strAttID <> 0 and  strAttvID <> 0 Then Response.Write objRsAttUsage(2).value%>">&nbsp;
	Updated By:<INPUT align="right" name="txtRecordStatusInd" type="text" style="width: 150px"
	disabled value="<%If strAttID <> 0 and  strAttvID <> 0 Then Response.Write objRsAttUsage(3).value%>">
	</DIV>
</FIELDSET>
</FORM>
<%
	'Clean up our ADO objects
	Set objRsAttUsage = Nothing
	set objRsAtt = Nothing
	set objRsAttv = Nothing

	objConn.Close
	Set ObjConn = Nothing
%>
</BODY>
</HTML>

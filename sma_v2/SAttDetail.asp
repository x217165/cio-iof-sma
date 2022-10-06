<%@ Language=VBScript %>
<% Option Explicit
 on error resume next
%>
<% Response.Buffer = true %>
<!--#include file = "smaConstants.inc"-->
<!--#include file = "smaProcs.inc"-->
<!--#include file = "databaseconnect.asp" -->
<!--
*********************************************************************************************
* Page name:	SAttDetail.asp																*
* Purpose:		To display the service attribute											*
* Created by:	Linda Chen																	*
* Date:			August 2009																	*
*********************************************************************************************
-->
<%
Dim strAttID, strWinMessage, strWinLocation
Dim	 strErrMessage
Dim intAccessLevel
Dim strSQL, objRsAtt
Dim struserid, strAttname, strAttdesc
Dim preSmaOnly, preSrtOnly, strSQLapp, ObjRsApp


	'Check user's rights
	intAccessLevel = CInt(CheckLogon(strConst_ServiceCatalogue))
	If (intAccessLevel And intConst_Access_ReadOnly) <> intConst_Access_ReadOnly Then
		DisplayError "BACK", "", 0, "ACCESS DENIED", "You don't have access to Line of Business. Please contact your system administrator"
	End If

	strWinMessage = ""
	strAttID = Request("hdnAttID")
	if (strAttID <>0) then
		strSQL = "SELECT COUNT(*) FROM " &_
			 " CRP.APPL_SRVC_TYP_ATT_RULE r, CRP.APPL_SRVC_TYP_ATT_RULE_STAT rs "  &_
			 " where r.APPL_SRVC_TYP_ATT_RULE_ID = rs.APPL_SRVC_TYP_ATT_RULE_ID " & _
			 " AND rs.APPL_SRVC_TYP_ATT_RULE_STAT_CD = 'A' "&_
			 " AND (rs.EFF_STOP_TS > sysdate or rs.EFF_STOP_TS = NULL) " &_
			 " AND R.SRVC_TYPE_ATT_ID = " & strAttID
		strSQLapp = strSQL + " AND APPLICATION_ID = 1 "
		set objRsApp = objConn.Execute(strSQLapp)
		preSmaOnly = objRSApp(0)

		set objRSApp = Nothing
		strSQLapp = strSQL + " AND APPLICATION_ID = 2 "
		set objRsApp = objConn.Execute(strSQLapp)
		preSrtOnly = objRSapp(0)

    end if
'	response.end
	struserid = Session("username")
	strSQL	= 	"SELECT SRVC_TYPE_ATT_NAME, " &_
				"SRVC_TYPE_ATT_ID, " &_
				"SRVC_TYPE_ATT_DESC, " &_
				"CREATE_DATE_TIME, " &_
				"CREATE_REAL_USERID, " &_
				"UPDATE_DATE_TIME, " &_
				"UPDATE_REAL_USERID "&_
				"FROM   CRP.SRVC_TYPE_ATT  " &_
				"WHERE  RECORD_STATUS_IND = 'A' "
	if (strAttID <> 0) then
		strSQL = strSQL + " AND SRVC_TYPE_ATT_ID = " & strAttID
	end if
	strSQL= strSQL + " ORDER BY SRVC_TYPE_ATT_NAME "
    set objRsAtt = objConn.Execute(strSQL)

	Select Case UCase(Request("hdnFrmAction"))
		Case "SAVE"
			strAttname = request("txtAttName")
			strAttDesc = request("txtAttDesc")
			preSrtOnly = request("SrtOnly")
			preSmaOnly = request("SmaOnly")
			If Request("hdnAttID") <> 0 Then	'Update existing Service Type Attribute
				If (intAccessLevel And intConst_Access_Update) <> intConst_Access_Update Then
					DisplayError "REFRESH", strWinLocation, 0, "UPDATE DENIED", "You don't have access to update Service Instance Attribute. Please contact your system administrator"
				End If
				dim cmdUpdateObj
				set cmdUpdateObj = server.CreateObject("ADODB.Command")
				set cmdUpdateObj.ActiveConnection = objConn
				cmdUpdateObj.CommandType = adCmdStoredProc
				cmdUpdateObj.CommandText = "SMA_SP_USERID.Sp_Srvtype_Att_Update"

				cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_userid", adVarChar , adParamInput, 20, struserid)
				cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_SRVC_TYPE_ATT_ID",adNumeric , adParamInput,, Clng(strAttID))
				cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_SRVC_TYPE_ATT_name",adVarChar , adParamInput,80, strAttname)
				cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_SRVC_TYPE_ATT_desc",adVarChar , adParamInput,255, strAttDesc)

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

				cmdUpdateObj.Execute
			Else									'Create a new Service Type Attribute
				If (intAccessLevel And intConst_Access_Create) <> intConst_Access_Create Then
					DisplayError "REFRESH", strWinLocation, 0, "INSERT DENIED", "You don't have access to create Service Instance Attribute. Please contact your system administrator"
				End If
				dim cmdInsertObj
				set cmdInsertObj = server.CreateObject("ADODB.Command")
				set cmdInsertObj.ActiveConnection = objConn
				cmdInsertObj.CommandType = adCmdStoredProc
				cmdInsertObj.CommandText = "SMA_SP_USERID.Sp_Srvtype_Att_Insert"
				cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_userid", adVarChar , adParamInput, 20, struserid)
 				cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_SRVC_TYPE_ATT_name",adVarChar , adParamInput,80, strAttname)
 				cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_SRVC_TYPE_ATT_desc",adVarChar , adParamInput,255, strAttDesc)

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

				cmdInsertObj.Execute
'				if objConn.Errors.Count <> 0 then
'					DisplayError "BACK", "", objConn.Errors(0).NativeError, "CANNOT ADD NEW RECORD", objConn.Errors(0).Description
'					objConn.Errors.Clear
'					response.redirect("SAttDetail.asp")
'				else
'					response.write("<script language=""javascript"">window.close();parent.opener.iSTAttFrame_display();</script>")
'				end if
			End If
			if objConn.Errors.Count <> 0 then
				DisplayError "BACK", "", objConn.Errors(0).NativeError, "CANNOT Update/ADD NEW RECORD", objConn.Errors(0).Description
				objConn.Errors.Clear
				response.redirect("SAttDetail.asp")
			else
				dim cmdAppUpdObj
				set cmdAppUpdObj = server.CreateObject("ADODB.Command")
				set cmdAppUpdObj.ActiveConnection = objConn
				cmdAppUpdObj.CommandType = adCmdStoredProc
				cmdAppUpdObj.CommandText = "SMA_SP_USERID.SRVC_ATT_SMA_DISPLAY_CHK"

				cmdAppUpdObj.Parameters.Append cmdAppUpdObj.CreateParameter("p_userid", adVarChar , adParamInput, 20, struserid)
 				cmdAppUpdObj.Parameters.Append cmdAppUpdObj.CreateParameter("p_srvc_type_att_name",adVarChar , adParamInput,80, strAttname)
 				cmdAppUpdObj.Parameters.Append cmdAppUpdObj.CreateParameter("p_appl_name",adVarChar , adParamInput,50, "SMA")
	 			if (len(preSmaOnly) >=3) then
	 				cmdAppUpdObj.Parameters.Append cmdAppUpdObj.CreateParameter("p_appl_disp_flag number",adNumeric, adParamInput,, 1)
	 			else
	 				cmdAppUpdObj.Parameters.Append cmdAppUpdObj.CreateParameter("p_appl_disp_flag number",adNumeric, adParamInput,, 0)
	 			end if
	 		'	response.write " preSmaOnly is "& preSmaOnly  & "<br>"
	 		'	Response.Write "<b> count = " & cmdUpdateObj.Parameters.count & "<br>"
  			'	dim nx
  			'	for nx=0 to cmdAppUpdObj.Parameters.count-1
  			'	   Response.Write nx+1 & " parm value= " & cmdAppUpdObj.Parameters.Item(nx).Value  & "<br>"
  			'	next

  			'	response.write (cmdAppUpdObj.CommandText)
			'	response.end
 				cmdAppUpdObj.Execute
 				if objConn.Errors.Count <> 0 then
					DisplayError "BACK", "", objConn.Errors(0).NativeError, "CANNOT ADD/UPDATE SMA APP RECORD", objConn.Errors(0).Description
					objConn.Errors.Clear
					response.redirect("SAttDetail.asp")
				end if

 				set cmdAppUpdObj = server.CreateObject("ADODB.Command")
			    set cmdAppUpdObj.ActiveConnection = objConn
			    cmdAppUpdObj.CommandType = adCmdStoredProc
			    cmdAppUpdObj.CommandText = "SMA_SP_USERID.SRVC_ATT_SMA_DISPLAY_CHK"

				cmdAppUpdObj.Parameters.Append cmdAppUpdObj.CreateParameter("p_userid", adVarChar , adParamInput, 20, struserid)
 				cmdAppUpdObj.Parameters.Append cmdAppUpdObj.CreateParameter("p_srvc_type_att_name",adVarChar , adParamInput,80, strAttname)
 				cmdAppUpdObj.Parameters.Append cmdAppUpdObj.CreateParameter("p_appl_name",adVarChar , adParamInput,50, "SRT")
 				if (len(preSrtOnly) >=3 ) then
	 				cmdAppUpdObj.Parameters.Append cmdAppUpdObj.CreateParameter("p_appl_disp_flag number",adNumeric, adParamInput,, 1)
				else
	 				cmdAppUpdObj.Parameters.Append cmdAppUpdObj.CreateParameter("p_appl_disp_flag number",adNumeric, adParamInput,, 0)
				end if
				'response.write "preSrtOnly is " & preSrtOnly & "<br>"
	 			'Response.Write "<b> count = " & cmdUpdateObj.Parameters.count & "<br>"
  				'dim nx
  				'for nx=0 to cmdAppUpdObj.Parameters.count-1
  				'   Response.Write nx+1 & " parm value= " & cmdAppUpdObj.Parameters.Item(nx).Value  & "<br>"
  				'next

  				'response.write (cmdAppUpdObj.CommandText)
				'response.end


	 			cmdAppUpdObj.Execute
	 			if objConn.Errors.Count <> 0 then
					DisplayError "BACK", "", objConn.Errors(0).NativeError, "CANNOT ADD/UPDATE SRT APP RECORD", objConn.Errors(0).Description
					objConn.Errors.Clear
					response.redirect("SAttDetail.asp")
				END IF
	 			'response.end
				response.write("<script language=""javascript"">window.close();parent.opener.iSTAttFrame_display();</script>")
			end if
		Case "DELETE"
			'Delete the attribute from application use first
 			set cmdAppUpdObj = server.CreateObject("ADODB.Command")
			set cmdAppUpdObj.ActiveConnection = objConn
			cmdAppUpdObj.CommandType = adCmdStoredProc
			cmdAppUpdObj.CommandText = "SMA_SP_USERID.SRVC_ATT_SMA_DISPLAY_CHK"

			cmdAppUpdObj.Parameters.Append cmdAppUpdObj.CreateParameter("p_userid", adVarChar , adParamInput, 20, struserid)
 		'	cmdAppUpdObj.Parameters.Append cmdAppUpdObj.CreateParameter("p_srvc_type_att_name",adVarChar , adParamInput,50, strAttname)
			cmdAppUpdObj.Parameters.Append cmdAppUpdObj.CreateParameter("p_srvc_type_att_name",adVarChar , adParamInput,80, objRsAtt(0))
 			cmdAppUpdObj.Parameters.Append cmdAppUpdObj.CreateParameter("p_appl_name",adVarChar , adParamInput,50, "SRT")
			cmdAppUpdObj.Parameters.Append cmdAppUpdObj.CreateParameter("p_appl_disp_flag number",adNumeric, adParamInput,, 0)

			Response.Write "<b> count = " & cmdUpdateObj.Parameters.count & "<br>"
			'dim nx
			'for nx=0 to cmdAppUpdObj.Parameters.count-1
 			'   Response.Write nx+1 & " parm value= " & cmdAppUpdObj.Parameters.Item(nx).Value  & "<br>"
 			'next

 			'response.write (cmdAppUpdObj.CommandText)
			'response.end
			cmdAppUpdObj.Execute

			set cmdAppUpdObj = server.CreateObject("ADODB.Command")
		    set cmdAppUpdObj.ActiveConnection = objConn
		    cmdAppUpdObj.CommandType = adCmdStoredProc
		    cmdAppUpdObj.CommandText = "SMA_SP_USERID.SRVC_ATT_SMA_DISPLAY_CHK"

			cmdAppUpdObj.Parameters.Append cmdAppUpdObj.CreateParameter("p_userid", adVarChar , adParamInput, 20, struserid)
 			cmdAppUpdObj.Parameters.Append cmdAppUpdObj.CreateParameter("p_srvc_type_att_name",adVarChar , adParamInput,80, objRsAtt(0))
 			cmdAppUpdObj.Parameters.Append cmdAppUpdObj.CreateParameter("p_appl_name",adVarChar , adParamInput,50, "SMA")
	 		cmdAppUpdObj.Parameters.Append cmdAppUpdObj.CreateParameter("p_appl_disp_flag number",adNumeric, adParamInput,, 0)

			'Response.Write "<b> count = " & cmdUpdateObj.Parameters.count & "<br>"
	  		''dim nx
	 		'for nx=0 to cmdAppUpdObj.Parameters.count-1
 	    	'   Response.Write nx+1 & " parm value= " & cmdAppUpdObj.Parameters.Item(nx).Value  & "<br>"
 	 		'next
 	 		'response.write (cmdAppUpdObj.CommandText)
 	 		'response.end
			cmdAppUpdObj.Execute

			if objConn.Errors.Count <> 0 then
				DisplayError "BACK", "", objConn.Errors(0).NativeError, "CANNOT ADD/UPDATE SMA APP RECORD", objConn.Errors(0).Description
				objConn.Errors.Clear
				response.redirect("StypeAttUsage1.asp")
			else
			' If delete the attribute from application is successful, delete it
		 		dim cmdDeleteObj
				set cmdDeleteObj = server.CreateObject("ADODB.Command")
				set cmdDeleteObj.ActiveConnection = objConn
				cmdDeleteObj.CommandType = adCmdStoredProc
				cmdDeleteObj.CommandText = "SMA_SP_USERID.Sp_Srvtype_Att_Delete"

				cmdDeleteObj.Parameters.Append cmdDeleteObj.CreateParameter("p_userid", adVarChar , adParamInput, 20, struserid)
				cmdDeleteObj.Parameters.Append cmdDeleteObj.CreateParameter("p_SRVC_TYPE_ATT_ID",adNumeric , adParamInput,, Clng(strAttID))

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
				cmdDeleteObj.Execute
				if objConn.Errors.Count <> 0 then
					DisplayError "BACK", "", objConn.Errors(0).NativeError, "CANNOT Delete RECORD", objConn.Errors(0).Description
					objConn.Errors.Clear

			    end if
   				response.redirect("StypeAttUsage1.asp")
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


function btnDelete_onClick() {
//**********************************************************************************************
// Function:	btnDelete_onClick
//
// Purpose:		To delete a line of business
//
// Created By:	Gilles Archer 09/27/2000
//
// Updated By:
//***********************************************************************************************
// Remove the comment in the 4 lines after test LC
//	if ((intAccessLevel & intConst_Access_Delete) != intConst_Access_Delete) {
//		alert('You do not have permission to DELETE a Service Instance Attribute.  Please contact your System Administrator.');
//		return false;
//	}

	if (document.frmAttDetail.txtAttName.value == "") {
		alert('This Service Type Attribute does not exist in the database.');
		return false;
	}

	if (confirm('Do you really want to delete this object?')) {
		document.frmAttDetail.hdnFrmAction.value = "DELETE";
		bolSaveRequired = false;
		document.frmAttDetail.submit();
	}
//	document.location.href='STypeAttUsage.asp?';
}


//function fct_onChange() {
//	bolSaveRequired = true;
//}

function btnSave_onClick() {
	if ((intAccessLevel & intConst_Access_Update) != intConst_Access_Update) {
		alert('You do not have permission to UPDATE a Service Type Attribute.  Please contact your System Administrator.');
		return false;
	}

	if (document.frmAttDetail.txtAttName.value == "") {
		alert('Please enter the Service Type Attribute');
		document.frmAttDetail.txtAttName.focus();
		return false;
	}

	if (document.frmAttDetail.txtAttDesc.value == "") {
		alert('Please enter the Service Type Attribute Description');
		document.frmAttDetail.txtAttDesc.focus();
		return false;
	}

	document.frmAttDetail.hdnFrmAction.value = "SAVE";
	bolSaveRequired = false;
	document.frmAttDetail.submit();
//	window.close();
//	parent.opener.iSTAttFrame_display();

//	document.location.href='SAttMRDetail.asp?';
//	self.document.location.href='SInstAttDetail.asp?hdnInstAttID='+document.frmAttDetail.hdnInstAttID.value;
}

function btnClose_onClick() {
	window.close();
	parent.opener.iSTAttFrame_display();
}



function window_onBeforeUnload() {
	//Ensure that fct_onChange() fires for any changed data.
	document.frmAttDetail.btnSave.focus();

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
		document.location.href = "SAttDetail.asp?hdnAttID=<%=strAttID%>";
	}
}



// Unhide Client-Side SCRIPT -->
</SCRIPT>
</HEAD>
<BODY language="javascript" onLoad="DisplayStatus(strWinMessage);" onBeforeUnload="window_onBeforeUnload();" onUnload="window_onUnload();">
<FORM id="frmAttDetail" name="frmAttDetail" action="SAttDetail.asp" method="post">
	<input id="hdnFrmAction" name="hdnFrmAction" type=hidden>
	<INPUT id="hdnAttID" name="hdnAttID" type=hidden
	value="<%If strAttID <>0 Then Response.Write strAttID else response.write 0 end if %>">
<TABLE border="0" cols="4" width="100%">
<THEAD>
<TR>
	<TD align="left" colspan="3">Service Type Attribute Detail</TD>
	<TD align="right" width="1%">&nbsp;</TR>
</THEAD>
<TBODY>
<TR>
	<TD align="left" nowrap width="13%">&nbsp;Attribute Name<FONT color="red">*</FONT></TD>
	<TD align="left" colspan="2" nowrap>
	<INPUT id="txtAttName" name="txtAttName"
	value="<% if strAttID <> 0 then response.write objRsAtt(0) else response.write "" end if %>" size="37" style="width: 500px" >
</TR size="28">
<TR>
	<TD align="left" nowrap width="13%">&nbsp;Attribute Description<FONT color="red">&nbsp;</FONT></TD>
	<TD align="left" colspan="2" nowrap>
	<INPUT id="txtAttDesc" name="txtAttDesc"
	value="<%if strAttID <> 0 then response.write objRsAtt(2) else response.write "" end if %>" size="37" style="width: 500px">
</TR size="41">

<tr>
<td></td>
<TD width="19%" >SMA Display
<INPUT id=SmaOnly name=SmaOnly tabindex=12 type=checkbox value=Yes
<% if Clng(preSmaOnly) = 1 then
	 response.write("CHECKED = yes ")
end if %>
 style="HEIGHT: 24px; WIDTH: 24px"></td>
<td width="64%">
SRT Display
<INPUT id=SrtOnly name=SrtOnly tabindex=12 type=checkbox value="YES"
<% if Clng(preSrtOnly) = 1 then
	 response.write("CHECKED = yes ")
end if %>
style="HEIGHT: 24px; WIDTH: 24px"></TD>

</tr>

</TR>
<TFOOT>
<TR>
	<TD colspan="4" align="right">
	<INPUT id="btnClose" name="btnClose" type="button" value="Close" style="width: 2cm" language="javascript" onClick="btnClose_onClick();">&nbsp;
	<INPUT id="btnDelete" name="btnDelete" type=hidden value="Delete" style="width: 76; height:22" language="javascript" onClick="btnDelete_onClick();">&nbsp;
	<INPUT id="btnReset" name="btnReset" type="button" value="Reset" style="width: 2cm" language="javascript" onClick="btnReset_onClick();" >&nbsp;
	<INPUT id="btnSave" name="btnSave" type="button" value="Save" style="width: 76; height:26" onClick="return btnSave_onClick();">&nbsp;
</TR>
</TFOOT>
</TABLE>

<FIELDSET width="100%">
	<LEGEND align="right"><B>Audit Information</B></LEGEND>
	<DIV size="8pt" align="right">
	Record Status Indicator:<INPUT align="left" name="txtRecordStatusInd" type="text" style="width: 18px"
	disabled value="<%If strAttID <> 0 Then Response.Write "A" end if %>">&nbsp;&nbsp;&nbsp;
	Create Date:<INPUT align="center" name="txtRecordStatusInd" type="text" style="width: 150px"
	disabled value="<%If strAttID <>0 Then Response.Write objRsAtt(3).value%>">&nbsp;
	Created By:<INPUT align="right" name="txtRecordStatusInd" type="text" style="width: 150px"
	disabled value="<%If strAttID <>0 Then Response.Write objRsAtt(4).value%>"><BR>
	Update Date:<INPUT align="center" name="txtRecordStatusInd" type="text" style="width: 150px"
	disabled value="<%If strAttID <>0 Then Response.Write objRsAtt(5).value%>">&nbsp;
	Updated By:<INPUT align="right" name="txtRecordStatusInd" type="text" style="width: 150px"
	disabled value="<%If strAttID <>0 Then Response.Write objRsAtt(6).value%>">
	</DIV>
</FIELDSET>
</FORM>
<%
	'Clean up our ADO objects
	Set objRsAtt = Nothing
	Set ObjRsApp = Nothing
	objConn.Close
	Set ObjConn = Nothing
%>
</BODY>
</HTML>

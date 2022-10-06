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
* Page name:	SInstAttvDetail.asp															*
* Purpose:		To display the service instance value detail								*
* Created by:	Linda Chen																	*
* Date:			August 2009																	*
*********************************************************************************************
-->
<%
Dim strWinMessage, strWinLocation, strErrMessage
Dim intAccessLevel
Dim strSQL, objRsInstAttv
Dim struserid, strInstAttvID,strAttvalue


	'Check user's rights
	intAccessLevel = CInt(CheckLogon(strConst_ServiceCatalogue))
	If (intAccessLevel And intConst_Access_ReadOnly) <> intConst_Access_ReadOnly Then
		DisplayError "BACK", "", 0, "ACCESS DENIED", "You don't have access to Service Instance Attribute Value. Please contact your system administrator"
	End If

	strWinMessage = ""
	strInstAttvID = Request("hdnInstAttvID")
	struserid = Session("username")
	strSQL	= 	"SELECT SRVC_INSTNC_ATT_VAL, " &_
				"SRVC_INSTNC_ATT_VAL_ID, " &_
				"CREATE_DATE_TIME, " &_
				"CREATE_REAL_USERID, " &_
				"UPDATE_DATE_TIME, " &_
				"UPDATE_REAL_USERID "&_
				"FROM   SO.SRVC_INSTNC_ATT_VAL  " &_
				"WHERE  RECORD_STATUS_IND = 'A' "
	if (strInstAttvID <> 0 and strInstAttvID <> "") then
		strSQL = strSQL + " AND SRVC_INSTNC_ATT_VAL_ID = " & strInstAttvID
	end if
	strSQL= strSQL + " ORDER BY SRVC_INSTNC_ATT_VAL "

	'response.write strSQL
	'response.end
    set objRsInstAttv = objConn.Execute(strSQL)

	Select Case UCase(Request("hdnFrmAction"))
		Case "SAVE"
			strAttvalue = request("txtInstAttvalue")
			If Request("hdnInstAttvID") <> 0 Then	'Update existing Service Instance Attribute
				If (intAccessLevel And intConst_Access_Update) <> intConst_Access_Update Then
					DisplayError "REFRESH", strWinLocation, 0, "UPDATE DENIED", "You don't have access to update Service Instance Attribute Value. Please contact your system administrator"
				End If

				dim cmdUpdateObj
				set cmdUpdateObj = server.CreateObject("ADODB.Command")
				set cmdUpdateObj.ActiveConnection = objConn
				cmdUpdateObj.CommandType = adCmdStoredProc
				cmdUpdateObj.CommandText = "SMA_SP_USERID.SP_SRVINST_ATTV_UPDATE"

				cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_userid", adVarChar , adParamInput, 20, struserid)
				cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_SRVC_INSTNC_ATT_ID",adNumeric , adParamInput,, Clng(strInstAttvID))
				cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_SRVC_INSTNC_ATT_val",adVarChar , adParamInput,80, strAttvalue)

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
				if objConn.Errors.Count <> 0 then
				  DisplayError "BACK", "", objConn.Errors(0).NativeError, "CANNOT ADD NEW RECORD", objConn.Errors(0).Description
				   objConn.Errors.Clear
				   response.redirect("SInstAttDetail.asp")
			    else
				    response.write("<script language=""javascript"">window.close();parent.opener.iSInstvFrame_display();</script>")
			    end if

'				strSQL	=	"Update so.srvc_instnc_att_val " &_
'							"set SRVC_INSTNC_ATT_VAL = '" & strAttvalue &_
'							"', UPDATE_DATE_TIME = SYSDATE,  " &_
'							"UPDATE_DB_USERID = 'JAGORA', " &_
'							"UPDATE_REAL_USERID = '" & struserid & "' " &_
'							"WHERE SRVC_INSTNC_ATT_VAL_ID = " & strInstAttvID
				'response.write(strSQL)
				'response.end
'				strErrMessage = "CANNOT UPDATE OBJECT"
'				On Error Resume Next
'				objConn.Execute(strSQL)
'				If objConn.Errors.Count <> 0 then
'					DisplayError "BACK", "", objConn.Errors(0).NativeError, strErrMessage, objConn.Errors(0).Description
'					objConn.Errors.Clear
'					response.redirect("SInstAttvDetail.asp")
''				else
'					response.redirect("STypeInstUsage.asp")
'				End If

			Else									'Create a new Service Instance Attribute
				If (intAccessLevel And intConst_Access_Create) <> intConst_Access_Create Then
					DisplayError "REFRESH", strWinLocation, 0, "INSERT DENIED", "You don't have access to create Service Instance Attribute Value. Please contact your system administrator"
				End If

				dim cmdInsertObj
				set cmdInsertObj = server.CreateObject("ADODB.Command")
				set cmdInsertObj.ActiveConnection = objConn
				cmdInsertObj.CommandType = adCmdStoredProc
				cmdInsertObj.CommandText = "SMA_SP_USERID.SP_SRVINST_ATTV_INSERT"
				cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_userid", adVarChar , adParamInput, 20, struserid)
 				cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_SRVC_INSTNC_ATT_VAL",adVarChar , adParamInput,80, strAttvalue)

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
			End If

			if objConn.Errors.Count <> 0 then
				DisplayError "BACK", "", objConn.Errors(0).NativeError, "CANNOT ADD NEW RECORD", objConn.Errors(0).Description
				objConn.Errors.Clear
				response.redirect("SInstAttDetail.asp")
			else
				response.write("<script language=""javascript"">window.close();parent.opener.iSInstvFrame_display();</script>")
			end if


		'	strSQL	=	"Insert into so.srvc_instnc_att_val " &_
		'					"(SRVC_INSTNC_ATT_VAL, "	&_
		'					"CREATE_DATE_TIME, CREATE_DB_USERID,  " &_
		'					"CREATE_REAL_USERID, UPDATE_DATE_TIME, " &_
		'					"UPDATE_DB_USERID,	UPDATE_REAL_USERID, RECORD_STATUS_IND) " &_
		'					"VALUES " &_
		'					"('" & strAttvalue & "',"  &_
		'					"SYSDATE, 'JAGORA', '" & struserid & "', SYSDATE, " &_
		'					"'JAGORA', '" & struserid & "', 'A')"
		'		'response.write(strSQL)
		'		'response.end
		'		strErrMessage = "CANNOT CREATE OBJECT"
		'		On Error Resume Next
		'		objConn.Execute(strSQL)
		'		If objConn.Errors.Count <> 0 then
		'			DisplayError "BACK", "", objConn.Errors(0).NativeError, strErrMessage, objConn.Errors(0).Description
		'			objConn.Errors.Clear
		'			response.redirect("SInstAttvDetail.asp")
'		'		else
'		'			response.redirect("STypeInstUsage.asp")
'		'		End If

'			End If
		Case "DELETE"
		        If (intAccessLevel And intConst_Access_Delete) <> intConst_Access_Delete Then
					DisplayError "REFRESH", strWinLocation, 0, "INSERT DENIED", "You don't have access to delete Service Instance Attribute value. Please contact your system administrator"
				End If

				dim cmdDeleteObj
				set cmdDeleteObj = server.CreateObject("ADODB.Command")
				set cmdDeleteObj.ActiveConnection = objConn
				cmdDeleteObj.CommandType = adCmdStoredProc
				cmdDeleteObj.CommandText = "SMA_SP_USERID.SP_SRVINST_ATTV_DELETE"

				cmdDeleteObj.Parameters.Append cmdDeleteObj.CreateParameter("p_userid", adVarChar , adParamInput, 20, struserid)
				cmdDeleteObj.Parameters.Append cmdDeleteObj.CreateParameter("p_SRVC_INST_ATT_VAL_ID",adNumeric , adParamInput,, Clng(strInstAttvID))

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
				    response.redirect("SInstAttvDetail.asp")
			    else
				    response.redirect("SInstAttUsage2.asp")
				end if

'			strSQL	=	"UPDATE so.srvc_instnc_att_val " &_
'						"SET RECORD_STATUS_IND ='D' " &_
'						"WHERE SRVC_INSTNC_ATT_VAL_ID = " & strInstAttvID
'				'response.write(strSQL)
'				'response.end
'				strErrMessage = "CANNOT DELETE OBJECT"
'				On Error Resume Next
'				objConn.Execute(strSQL)
'				If objConn.Errors.Count <> 0 then
'					DisplayError "BACK", "", objConn.Errors(0).NativeError, strErrMessage, objConn.Errors(0).Description
'					objConn.Errors.Clear
'					response.redirect("SInstAttvDetail.asp")
'				else
'					response.redirect("SInstAttUsage2.asp")
'				End If
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

	if (document.frmInstAttvDetail.txtInstAttvalue.value == "") {
		alert('This Service Instance Attribute does not exist in the database.');
		return false;
	}

	if (confirm('Do you really want to delete this object?')) {
		document.frmInstAttvDetail.hdnFrmAction.value = "DELETE";
		bolSaveRequired = false;
		document.frmInstAttvDetail.submit();
	}
//	document.location.href='STypeInstUsage.asp?';
}

function btnClose_onClick() {
	window.close();
	parent.opener.iSInstvFrame_display();
}
//function fct_onChange() {
//	bolSaveRequired = true;
//}

function btnSave_onClick() {
	if ((intAccessLevel & intConst_Access_Update) != intConst_Access_Update) {
		alert('You do not have permission to UPDATE a Service Instance Attribute Value.  Please contact your System Administrator.');
		return false;
	}

	if (document.frmInstAttvDetail.txtInstAttvalue == "") {
		alert('Please enter the Service Instance Attribute Value');
		document.frmInstAttvDetail.txtInstAttvalue.focus();
		return false;
	}

	document.frmInstAttvDetail.hdnFrmAction.value = "SAVE";
	bolSaveRequired = false;
	document.frmInstAttvDetail.submit();
//	window.close();
//	parent.opener.iSInstvFrame_display();
}

function window_onBeforeUnload() {
	//Ensure that fct_onChange() fires for any changed data.
	document.frmInstAttvDetail.btnSave.focus();

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
		document.location.href = "SInstAttvDetail.asp?hdnInstAttvID=<%=strInstAttvID%>";
	}
}
// Unhide Client-Side SCRIPT -->
</SCRIPT>
</HEAD>
<BODY language="javascript" onLoad="DisplayStatus(strWinMessage);" onBeforeUnload="window_onBeforeUnload();" onUnload="window_onUnload();">
<FORM id="frmInstAttvDetail" name="frmInstAttvDetail" action="SInstAttvDetail.asp" method="post">
	<input id="hdnFrmAction" name="hdnFrmAction" type=hidden>
	<INPUT id="hdnInstAttvID" name="hdnInstAttvID" type=hidden
	value="<%If strInstAttvID <>0 Then Response.Write strInstAttvID else response.write 0 end if %>">
<TABLE border="0" cols="4" width="100%">
<THEAD>
<TR>
	<TD align="left" colspan="3">Service Instance Attribute Value Detail</TD>
	<TD align="right" width="2%">&nbsp;</TR>
</THEAD>
<TBODY>
<TR>
	<TD align="left" nowrap width="21%">Instance Attribute Value<FONT color="red">*</FONT></TD>
	<TD align="left" colspan="2" nowrap>
	<INPUT id="txtInstAttvalue" name="txtInstAttvalue"
	value="<% if strInstAttvID <> 0 then response.write objRsInstAttv(0) else response.write "" end if %>" style="width: 500px" >
</TR size="28">
</TR>
<TFOOT>
<TR>
	<TD colspan="4" align="right">
	<INPUT id="btnClose" name="btnClose" type="button" value="Close" style="width: 2cm" language="javascript" onClick="btnClose_onClick();">&nbsp;
	&nbsp;
	<INPUT id="btnDelete" name="btnDelete" type="hidden" value=Delete style="width: 2cm" language="javascript" onClick="btnDelete_onClick();">&nbsp;
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
	disabled value="<%If strInstAttvID <> 0 Then Response.Write "A" end if %>">&nbsp;&nbsp;&nbsp;
	Create Date:<INPUT align="center" name="txtRecordStatusInd" type="text" style="width: 150px"
	disabled value="<%If strInstAttvID <>0 Then Response.Write objRsInstAttv(2).value%>">&nbsp;
	Created By:<INPUT align="right" name="txtRecordStatusInd" type="text" style="width: 150px"
	disabled value="<%If strInstAttvID <>0 Then Response.Write objRsInstAttv(3).value%>"><BR>
	Update Date:<INPUT align="center" name="txtRecordStatusInd" type="text" style="width: 150px"
	disabled value="<%If strInstAttvID <>0 Then Response.Write objRsInstAttv(4).value%>">&nbsp;
	Updated By:<INPUT align="right" name="txtRecordStatusInd" type="text" style="width: 150px"
	disabled value="<%If strInstAttvID <>0 Then Response.Write objRsInstAttv(5).value%>">
	</DIV>
</FIELDSET>
</FORM>
<%
	'Clean up our ADO objects
	Set objRsInstAttv = Nothing
	objConn.Close
	Set ObjConn = Nothing
%>
</BODY>
</HTML>

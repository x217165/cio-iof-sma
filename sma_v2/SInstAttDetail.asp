<%@ Language=VBScript %> <% Option Explicit
 on error resume next
%> <% Response.Buffer = true %>
<!--#include file = "smaConstants.inc"-->
<!--#include file = "smaProcs.inc"-->
<!--#include file = "databaseconnect.asp" -->
<!--
*********************************************************************************************
* Page name:	SInstAttDetail.asp															*
* Purpose:		To display the service instance attribute detail							*
* Created by:	Linda Chen																	*
* Date:			August 2009																	*
*********************************************************************************************
--><%
Dim strInstAttID, strWinMessage, strWinLocation
Dim	 strErrMessage
Dim intAccessLevel
Dim strSQL, objRsSelect
Dim struserid, strAttname, strAttdesc, strFmtId


dim objRsInstAtt


	'Check user's rights
	intAccessLevel = CInt(CheckLogon(strConst_ServiceCatalogue))
	If (intAccessLevel And intConst_Access_ReadOnly) <> intConst_Access_ReadOnly Then
		DisplayError "BACK", "", 0, "ACCESS DENIED", "You don't have access to Line of Business. Please contact your system administrator"
	End If

	strWinMessage = ""
	strInstAttID = Request("hdnInstAttID")
	struserid = Session("username")
	strSQL	= 	"SELECT att.SRVC_INSTNC_ATT_NAME, " &_
				"att.SRVC_INSTNC_ATT_ID, " &_
				"att.SRVC_INSTNC_ATT_DESC, " &_
				"att.CREATE_DATE_TIME, " &_
				"att.CREATE_REAL_USERID, " &_
				"att.UPDATE_DATE_TIME, " &_
				"att.UPDATE_REAL_USERID, "&_
                "decode(att.SRVC_INSTNC_ATT_VAL_FORMAT_ID, NULL,0,  SRVC_INSTNC_ATT_VAL_FORMAT_ID  ) "&_
                "as SRVC_INSTNC_ATT_VAL_FORMAT_ID  "&_
				"FROM   so.SRVC_INSTNC_ATT att " &_
				"WHERE  RECORD_STATUS_IND = 'A' "
	if (strInstAttID <> 0) then
		strSQL = strSQL + " AND SRVC_INSTNC_ATT_ID = " & strInstAttID
	end if
	strSQL= strSQL + " ORDER BY SRVC_INSTNC_ATT_NAME "
    set objRsInstAtt = objConn.Execute(strSQL)
  ' response.write strSQL
   'response.write objRsInstAtt(7)
  ' response.end






' Got Att Format

	Set objRSSelect = Server.CreateObject("ADODB.Recordset")
    strSQL = "SELECT SRVC_INSTNC_ATT_VAL_FORMAT_ID, FORMAT_NAME  " &_
		  "FROM   SO.SRVC_INSTNC_ATT_VAL_FORMAT" &_
 	  " ORDER BY SRVC_INSTNC_ATT_VAL_FORMAT_ID "

	On Error Resume Next
	objRsSelect.Open strSQL, objConn, adOpenForwardOnly, adLockReadOnly, adCmdText
	If objConn.Errors.Count <> 0 Then
		DisplayError "BACK", "", objConn.Errors(0).NativeError, "ERROR LOADING DATA (Line of Business)", objConn.Errors(0).Description
		objConn.Errors.Clear
	End If




	if (strInstAttID <> 0) then
    	strFmtId = objRsInstAtt(7)
	end if



	Select Case UCase(Request("hdnFrmAction"))
		Case "SAVE"
			strAttname = request("txtInstAttName")
			strAttDesc = request("txtInstAttDesc")
			strFmtId = request("selfmtAtt")
			If Request("hdnInstAttID") <> 0 Then	'Update existing Service Instance Attribute
				If (intAccessLevel And intConst_Access_Update) <> intConst_Access_Update Then
					DisplayError "REFRESH", strWinLocation, 0, "UPDATE DENIED", "You don't have access to update Service Instance Attribute. Please contact your system administrator"
				End If
				dim cmdUpdateObj
				set cmdUpdateObj = server.CreateObject("ADODB.Command")
				set cmdUpdateObj.ActiveConnection = objConn
				cmdUpdateObj.CommandType = adCmdStoredProc
				cmdUpdateObj.CommandText = "SMA_SP_USERID.Sp_Srvinst_Att_Update"

				cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_userid", adVarChar , adParamInput, 20, struserid)
				cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_SRVC_INSTNC_ATT_ID",adNumeric , adParamInput,, Clng(strInstAttID))
				cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_SRVC_INSTNC_ATT_name",adVarChar , adParamInput,80, strAttname)
				cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_SRVC_INSTNC_ATT_desc",adVarChar , adParamInput,255, strAttDesc)
		        cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_SRVC_INSTNC_ATT_fmtID",adNumeric , adParamInput,, Clng(strFmtId))					'****************************
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
				    response.write("<script language=""javascript"">window.close();parent.opener.iSInstFrame_display();parent.opener.iSInstuFrame_display();")
				   ' if Clng(strFmtId)>0 then
				    	response.write("parent.opener.iSInstvFrame_display();")
				   ' end if
				    response.write "</script>"
			    end if

			Else									'Create a new Service Instance Attribute
				If (intAccessLevel And intConst_Access_Create) <> intConst_Access_Create Then
					DisplayError "REFRESH", strWinLocation, 0, "INSERT DENIED", "You don't have access to create Service Instance Attribute. Please contact your system administrator"
				End If
				dim cmdInsertObj
				set cmdInsertObj = server.CreateObject("ADODB.Command")
				set cmdInsertObj.ActiveConnection = objConn
				cmdInsertObj.CommandType = adCmdStoredProc
				cmdInsertObj.CommandText = "SMA_SP_USERID.Sp_srvinst_Att_Insert"
				cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_userid", adVarChar , adParamInput, 20, struserid)
 				cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_INSTNC_ATT_name",adVarChar , adParamInput,80, strAttname)
 				cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_INSTNC_ATT_name_desc",adVarChar , adParamInput,255, strAttDesc)
		        cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_SRVC_INSTNC_ATT_fmtID",adNumeric , adParamInput,, Clng(strFmtId))				'****************************
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
				response.write("<script language=""javascript"">window.close();parent.opener.iSInstFrame_display();parent.opener.iSInstuFrame_display();")
				if Clng(strFmtId)>0 then
				  	response.write("parent.opener.iSInstFrame_display(); parent.opener.iSInstvFrame_display();")
				end if
				response.write "</script>"
			end if

		Case "DELETE"
		        If (intAccessLevel And intConst_Access_Delete) <> intConst_Access_Delete Then
					DisplayError "REFRESH", strWinLocation, 0, "INSERT DENIED", "You don't have access to delete Service Instance Attribute. Please contact your system administrator"
				End If

				dim cmdDeleteObj
				set cmdDeleteObj = server.CreateObject("ADODB.Command")
				set cmdDeleteObj.ActiveConnection = objConn
				cmdDeleteObj.CommandType = adCmdStoredProc
				cmdDeleteObj.CommandText = "SMA_SP_USERID.Sp_Srvinst_Att_Delete"

				cmdDeleteObj.Parameters.Append cmdDeleteObj.CreateParameter("p_userid", adVarChar , adParamInput, 20, struserid)
				cmdDeleteObj.Parameters.Append cmdDeleteObj.CreateParameter("p_SRVC_INST_ATT_ID",adNumeric , adParamInput,, Clng(strInstAttID))
		        cmdDeleteObj.Parameters.Append cmdDeleteObj.CreateParameter("p_SRVC_INSTNC_ATT_fmtID",adNumeric , adParamInput,, Clng(strFmtId))
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
		 		'response.write Clng(strFmtId)
  				'Response.Write "<b> count = " & cmdDeleteObj.Parameters.count & "<br>"
  				''dim nx
  				'for nx=0 to cmdDeleteObj.Parameters.count-1
  				'   Response.Write nx+1 & " parm value= " & cmdDeleteObj.Parameters.Item(nx).Value  & "<br>"
  				'next

  				'response.write (cmdDeleteObj.CommandText)
				'response.end
				cmdDeleteObj.Execute

				if objConn.Errors.Count <> 0 then
					DisplayError "BACK", "", objConn.Errors(0).NativeError, "CANNOT Delete RECORD", objConn.Errors(0).Description
					objConn.Errors.Clear
				    response.redirect("SInstAttDetail.asp")
			    else
				    'response.redirect("SInstAttUsage1.asp")
				    response.write("<script language=""javascript"">window.close();parent.DisplayStatus("" "");")
				    response.write "</script>"

				end if


  		End Select	    ' For service attribute dropdown list


%>
<html>

<head>
<meta name="Generator" content="Microsoft Visual Studio 6.0">
<meta http-equiv="Pragma" content="no-cache">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<link rel="stylesheet" type="text/css" href="stylesheets/styles.css">
<script type="text/javascript" language="javascript" src="AccessLevels.js"></script>
<script type="text/javascript" language="javascript" src="GeneralJavaFunctions.js"></script>
<script type="text/javascript" language="javascript" id="clientEventHandlersJS">

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

	if (document.frmInstAttDetail.txtInstAttName.value == "") {
		alert('This Service Instance Attribute does not exist in the database.');
		return false;
	}

	if (confirm('Do you really want to delete this object?')) {
		document.frmInstAttDetail.hdnFrmAction.value = "DELETE";
		bolSaveRequired = false;
		document.frmInstAttDetail.submit();
	}
//	document.location.href='STypeInstUsage.asp?';
}


//function fct_onChange() {
//	bolSaveRequired = true;
//}

function btnSave_onClick() {
	if ((intAccessLevel & intConst_Access_Update) != intConst_Access_Update) {
		alert('You do not have permission to UPDATE a Service Instance Attribute.  Please contact your System Administrator.');
		return false;
	}

	if (document.frmInstAttDetail.txtInstAttName == "") {
		alert('Please enter the Service Instance Attribute');
		document.frmInstAttDetail.txtInstAttName.focus();
		return false;
	}

	if (document.frmInstAttDetail.txtInstAttDesc.value == "") {
		alert('Please enter the Service Instance Attribute Description');
		document.frmInstAttDetail.txtInstAttDesc.focus();
		return false;
	}

	document.frmInstAttDetail.hdnFrmAction.value = "SAVE";
	bolSaveRequired = false;
	document.frmInstAttDetail.submit();
//	window.close();
//	parent.opener.iSInstFrame_display();
}

function btnClose_onClick(){
	window.close();
	parent.opener.iSInstFrame_display();
}




function window_onBeforeUnload() {
	//Ensure that fct_onChange() fires for any changed data.
	document.frmInstAttDetail.btnSave.focus();

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
		document.location.href = "SInstAttDetail.asp?hdnInstAttID=<%=strInstAttID%>";
	}
}
// Unhide Client-Side SCRIPT -->
</script>
</head>

<body language="javascript" onload="DisplayStatus(strWinMessage);" onbeforeunload="window_onBeforeUnload();" onunload="window_onUnload();">

<form id="frmInstAttDetail" name="frmInstAttDetail" action="SInstAttDetail.asp" method="post">
	<input id="hdnFrmAction" name="hdnFrmAction" type="hidden">
	<input id="hdnInstAttID" name="hdnInstAttID" type="hidden"  value="<%If strInstAttID <>0 Then Response.Write strInstAttID else response.write 0 end if %>">
	<table border="0" cols="4" width="100%">
		<thead>
			<tr>
				<td align="left" colspan="3">Service Instance Attribute Detail</td>
				<td align="right" width="2%">&nbsp; </td>
			</tr>
		</thead>
		<tr>
			<td align="left" nowrap width="21%">Instance Attribute Name<font color="red">*</font></td>
			<td align="left" colspan="2" nowrap>
			<input id="txtInstAttName" name="txtInstAttName" value='<% if strInstAttID <> 0 then response.write objRsInstAtt(0) else response.write "" end if %>' style="width: 500px">
			</td size="28">
		</tr>
		<tr>
			<td align="left" nowrap width="21%">Instance Attribute Description<font color="red">
			</font></td>
			<td align="left" colspan="2" nowrap>
			<input id="txtInstAttDesc" name="txtInstAttDesc" value='<%if strInstAttID <> 0 then response.write objRsInstAtt(2) else response.write "" end if %>' style="width: 500px">
			</td size="41">
		</tr>
		<tr>
			<td align="left" nowrap width="21%">Instance Attribute Format<font color="red">
			</font></td>
			<td align="left" colspan="2" nowrap>
			<SELECT id=selfmtAtt name=selfmtAtt style="HEIGHT: 22; WIDTH: 498px">
							  <%if strInstAttID <> 0  then
							       if Clng(objRsInstAtt("SRVC_INSTNC_ATT_VAL_FORMAT_ID")) > 0  then%>


									  <%Do while Not objRsSelect.EOF%>
							   			<option value = "<% response.write objRsSelect("SRVC_INSTNC_ATT_VAL_FORMAT_ID")%>"
							   			 <% if StrComp(objRsSelect("SRVC_INSTNC_ATT_VAL_FORMAT_ID"),objRsInstAtt("SRVC_INSTNC_ATT_VAL_FORMAT_ID"),0)= 0 then response.write " selected" end if%>> <% =objRsSelect("FORMAT_NAME")%> </option>
									    <%  objRsSelect.movenext
									   loop
									   %>
 							   		   <OPTION value=0><%response.write "Predefined Value" %></OPTION>
 							   		<%else %>
						       			 <OPTION value=0 ><%response.write "Predefined Value" %></OPTION>
                                   		 <%do while not objRsSelect.EOF%>
	   		                       			 <option  value = <% =objRsSelect("SRVC_INSTNC_ATT_VAL_FORMAT_ID") %>> <% =objRsSelect("FORMAT_NAME")%> </option>
		                           		 <%  objRsSelect.movenext
		                           		 loop
									      %>
		                     		<%end if %>

							  <%else %>
						       		<OPTION value=0 ><%response.write "Predefined Value" %></OPTION>
                                    <%do while not objRsSelect.EOF%>
	   		                       <option  value = <% =objRsSelect("SRVC_INSTNC_ATT_VAL_FORMAT_ID") %>> <% =objRsSelect("FORMAT_NAME")%> </option>
		                            <%  objRsSelect.movenext
		                           loop
								   %>
		                      <%end if %>

		</SELECT>
 </td size="41">
		</tr>
		<tfoot>
			<tr>
				<td colspan="4" align="right">
				<input id="btnClose" name="btnClose" type="button" value="Close" style="width: 2cm" language="javascript" onclick="btnClose_onClick();">&nbsp;&nbsp;&nbsp;
				<input id="btnDelete" name="btnDelete" type="hidden" value="Delete" style="width: 2cm" language="javascript" onclick="btnDelete_onClick();">&nbsp;
				<input id="btnReset" name="btnReset" type="button" value="Reset" style="width: 2cm" language="javascript" onclick="btnReset_onClick();">&nbsp;&nbsp;&nbsp;
				<input id="btnSave" name="btnSave" type="button" value="Save" style="width: 2cm" onclick="return btnSave_onClick();">&nbsp;
				</td>
			</tr>
		</tfoot>
	</table>
	<fieldset width="100%">
	<legend align="right"><b>Audit Information</b></legend>
	<div size="8pt" align="right">
		Record Status Indicator:<input align="left" name="txtRecordStatusInd" type="text" style="width: 18px" disabled value='<%If strInstAttID <> 0 Then Response.Write "A" end if %>'>&nbsp;&nbsp;&nbsp;
		Create Date:<input align="center" name="txtRecordStatusInd" type="text" style="width: 150px" disabled value="<%If strInstAttID <>0 Then Response.Write objRsInstAtt(3).value%>">&nbsp;
		Created By:<input align="right" name="txtRecordStatusInd" type="text" style="width: 150px" disabled value="<%If strInstAttID <>0 Then Response.Write objRsInstAtt(4).value%>"><br>
		Update Date:<input align="center" name="txtRecordStatusInd" type="text" style="width: 150px" disabled value="<%If strInstAttID <>0 Then Response.Write objRsInstAtt(5).value%>">&nbsp;
		Updated By:<input align="right" name="txtRecordStatusInd" type="text" style="width: 150px" disabled value="<%If strInstAttID <>0 Then Response.Write objRsInstAtt(6).value%>">
	</div>
	</fieldset>
</form>
<%
	'Clean up our ADO objects
	Set objRsSelect = Nothing
	Set  objRsInstAtt = Nothing
	objConn.Close
	Set ObjConn = Nothing
%>

</body>

</html>

<%@ Language=VBScript %>
<%Option Explicit%>
<%Response.Buffer = True%>
<!--#include file = "smaConstants.inc"-->
<!--#include file = "smaProcs.inc"-->
<!--#include file = "databaseconnect.asp" -->
<!--
*********************************************************************************************
* Page name:	SCategoryDetail.asp															*
* Purpose:		To display the Service Category												*
*				Chosen via SCategoryList.asp												*
*																							*
* Created by:	Gilles Archer 09/27/2000													*
*																							*
*********************************************************************************************
-->
<%
Dim strServiceCategoryID, datUpdateDateTime, strWinMessage, strWinName, strWinLocation
Dim	objCommand, objRS, objRSFrench, objRSSelect, strSQL, strErrMessage, strLANG
Dim p_service_french
Dim intAccessLevel

	'Check user's rights
	intAccessLevel = CInt(CheckLogon(strConst_ServiceCatalogue))
	If (intAccessLevel And intConst_Access_ReadOnly) <> intConst_Access_ReadOnly  Then
		DisplayError "BACK", "", 0, "ACCESS DENIED", "You don't have access to Service Category. Please contact your system administrator"
	End If

	strWinName = Request("WinName")
	strWinMessage = ""
	strServiceCategoryID = Request.QueryString("ServiceCategoryID")

	if len(trim(Request.Form("txtServiceFrench"))) <> 0 Then
		p_service_french = trim(Request.Form("txtServiceFrench"))
	Else
		p_service_french = "NULL"
	end if

	Select Case UCase(Request("hdnFrmAction"))
		Case "SAVE"
			Set objCommand = Server.CreateObject("ADODB.Command")
			Set objCommand.ActiveConnection = objConn
			objCommand.CommandType = adCmdStoredProc

			If Request("hdnServiceCategoryID") <> "" Then	'Save existing Service Category
				If (intAccessLevel And intConst_Access_Update) <> intConst_Access_Update Then
					DisplayError "REFRESH", strWinLocation, 0, "UPDATE DENIED", "You don't have access to update service categories. Please contact your system administrator"
				End If

				objCommand.CommandText = "sma_sp_userid.spk_sma_admin_inter.sp_servcat_update"
				objCommand.Parameters.Append objCommand.CreateParameter("p_userid", adVarChar, adParamInput, 20, Session("username"))
				objCommand.Parameters.Append objCommand.CreateParameter("p_service_id", adNumeric, adParamInput, , CLng(Request("hdnServiceCategoryID")))
				objCommand.Parameters.Append objCommand.CreateParameter("p_service_desc", adVarChar, adParamInput, 80, Request("txtSCategoryDescription"))
				objCommand.Parameters.Append objCommand.CreateParameter("p_service_french", adVarChar, adParamInput, 80, p_service_french)
				objCommand.Parameters.Append objCommand.CreateParameter("p_lob_id", adNumeric, adParamInput, , Request("selLOB"))
				objCommand.Parameters.Append objCommand.CreateParameter("p_last_update_dt", adDBTimeStamp, adParamInput, , CDate(Request("hdnUpdateDateTime")))

				strErrMessage = "CANNOT UPDATE OBJECT"
			Else									'Create a new Service Category
				If (intAccessLevel And intConst_Access_Create) <> intConst_Access_Create Then
					DisplayError "REFRESH", strWinLocation, 0, "INSERT DENIED", "You don't have access to create service categories. Please contact your system administrator"
				End If

				objCommand.CommandText = "sma_sp_userid.spk_sma_admin_inter.sp_servcat_insert"
				objCommand.Parameters.Append objCommand.CreateParameter("p_userid", adVarChar, adParamInput, 20, Session("username"))
				objCommand.Parameters.Append objCommand.CreateParameter("p_service_id", adNumeric, adParamOutput, , Null)
				objCommand.Parameters.Append objCommand.CreateParameter("p_service_desc", adVarChar, adParamInput, 80, Request("txtSCategoryDescription"))
				objCommand.Parameters.Append objCommand.CreateParameter("p_service_french", adVarChar, adParamInput, 80, p_service_french)
				objCommand.Parameters.Append objCommand.CreateParameter("p_lob_id", adNumeric, adParamInput, , Request("selLOB"))

				strErrMessage = "CANNOT CREATE OBJECT"
			End If

			On Error Resume Next
			objCommand.Execute
			If objConn.Errors.Count <> 0 then
				DisplayError "BACK", "", objConn.Errors(0).NativeError, strErrMessage, objConn.Errors(0).Description
				objConn.Errors.Clear
			End If
			strServiceCategoryID = CStr(objCommand.Parameters("p_service_id").Value)
			Set objCommand = Nothing
			strWinMessage = "Record saved successfully. You can now see the changes you made."

		Case "DELETE"
			If (intAccessLevel And intConst_Access_Delete) <> intConst_Access_Delete Then
				DisplayError "REFRESH", strWinLocation, 0, "DELETE DENIED", "You don't have access to delete service categories. Please contact your system administrator"
			End If

			Set objCommand = Server.CreateObject("ADODB.Command")
			Set objCommand.ActiveConnection = objConn
			objCommand.CommandType = adCmdStoredProc
			objCommand.CommandText = "sma_sp_userid.spk_sma_admin_inter.sp_servcat_delete"
			objCommand.Parameters.Append objCommand.CreateParameter("p_service_id", adNumeric, adParamInput, , CLng(Request("hdnServiceCategoryID")))					'number(9)
			objCommand.Parameters.Append objCommand.CreateParameter("p_last_update_dt", adDBTimeStamp, adParamInput, , CDate(Request("hdnUpdateDateTime")))		'Date

  			On Error Resume Next
			objCommand.Execute
			If objConn.Errors.Count <> 0 Then
				DisplayError "BACK", "", objConn.Errors(0).NativeError, "CANNOT DELETE OBJECT", objConn.Errors(0).Description
				objConn.Errors.Clear
			End If
			Set objCommand = Nothing
			strServiceCategoryID = "DEL"
			strWinMessage = "Record deleted successfully."
	End Select

	If IsNumeric(strServiceCategoryID) Then
		strSQL = "SELECT SERVICE_CATEGORY_ID, SERVICE_CATEGORY_DESC, LOB_ID, " &_
				"TO_CHAR(CREATE_DATE_TIME, 'MON-DD-YYYY HH24:MI:SS') AS CREATE_DATE_TIME, " &_
				"sma_sp_userid.spk_sma_library.sf_get_full_username(CREATE_REAL_USERID) AS CREATE_REAL_USERID, " &_
				"TO_CHAR(UPDATE_DATE_TIME, 'MON-DD-YYYY HH24:MI:SS') AS UPDATE_DATE_TIME, " &_
				"sma_sp_userid.spk_sma_library.sf_get_full_username(UPDATE_REAL_USERID) AS UPDATE_REAL_USERID, " &_
				"UPDATE_DATE_TIME AS LAST_UPDATE_DATE_TIME, " &_
				"RECORD_STATUS_IND " &_
				"FROM CRP.SERVICE_CATEGORY " &_
				"WHERE SERVICE_CATEGORY_ID = " & strServiceCategoryID &_
				" AND RECORD_STATUS_IND = 'A'"

		'Create Recordset object
		Set objRS = Server.CreateObject("ADODB.Recordset")
		On Error Resume Next
		objRS.Open strSQL, objConn, adOpenForwardOnly, adLockReadOnly, adCmdText
		If objConn.Errors.Count <> 0 Then
			DisplayError "BACK", "", objConn.Errors(0).NativeError, "ERROR LOADING DATA", objConn.Errors(0).Description
			objConn.Errors.Clear
		End If

		strSQL = " SELECT SERVICE_CATEGORY_ID, SERVICE_CATEGORY_LANG_DESC " &_
				 " FROM CRP.SERVICE_CATEGORY_LANG " &_
				 " WHERE SERVICE_CATEGORY_ID = " & strServiceCategoryID &_
				 " AND RECORD_STATUS_IND = 'A' "

		'Create Recordset object
		Set objRSFrench = Server.CreateObject("ADODB.Recordset")
		On Error Resume Next
		objRSFrench.Open strSQL, objConn, adOpenForwardOnly, adLockReadOnly, adCmdText
		If objConn.Errors.Count <> 0 Then
			DisplayError "BACK", "", objConn.Errors(0).NativeError, "ERROR LOADING DATA", objConn.Errors(0).Description
			objConn.Errors.Clear
		End If

	End If

	'TQ_INOSS
	strLANG = Request.Cookies("UserInformation")("language_preference")
	if (Len(strLANG) = 0) then strLANG = "EN"

	'Get the Line of Business : TQ_INOSS
	strSQL = "SELECT LOB_ID, LOB_CODE, LOB_DESC " &_
			"FROM CRP.V_LOB " &_
			"WHERE lob_id NOT IN" &_
		        	"(SELECT lob_id " &_
		        	"FROM crp.v_lob " &_
		        	"WHERE language_preference_lcode = '" & strLANG & "' ) " &_
			"AND LANGUAGE_PREFERENCE_LCODE = 'EN'" &_
			"UNION SELECT LOB_ID, LOB_CODE, LOB_DESC " &_
			"FROM crp.v_lob " &_
			"WHERE language_preference_lcode = '" & strLANG & "'" &_
			"AND RECORD_STATUS_IND = 'A' " &_
			"ORDER BY LOB_DESC ASC "

	'Create Recordset object
	Set objRSSelect = Server.CreateObject("ADODB.Recordset")
	On Error Resume Next
	objRSSelect.Open strSQL, objConn, adOpenForwardOnly, adLockReadOnly, adCmdText
	If objConn.Errors.Count <> 0 Then
		DisplayError "BACK", "", objConn.Errors(0).NativeError, "ERROR LOADING DATA", objConn.Errors(0).Description
		objConn.Errors.Clear
	End If

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
var strWinName = "<%=strWinName%>";
var strWinMessage = "<%=strWinMessage%>";
var intAccessLevel = <%=intAccessLevel%>;
var	bolSaveRequired = false;

setPageTitle("SMA - Service Category");

function fct_selNavigate() {
//***********************************************************************************************
// Function:	selNavigate_onChange															*
//																								*
// Purpose:		To display the page selected by the user from Quick Navigation drop-down box.	*
//																								*
// Created By:	Gilles Archer 09/27/2000														*
//																								*
// Updated By:																					*
//***********************************************************************************************
var strPageName = document.frmSCategoryDetail.selNavigate.item(document.frmSCategoryDetail.selNavigate.selectedIndex).value;
var strBusinessID = document.frmSCategoryDetail.hdnBusinessID.value;
var strServiceCategoryID = document.frmSCategoryDetail.hdnServiceCategoryID.value;

	switch (strPageName) {
		case "LOB":
			document.frmSCategoryDetail.selNavigate.selectedIndex = 0;
			self.location.href = "LOBDetail.asp?BusinessID=" + strBusinessID;
			break;

		case "STypes":
			document.frmSCategoryDetail.selNavigate.selectedIndex = 0;
			SetCookie("BusinessID", strBusinessID);
			SetCookie("ServiceCategoryID", strServiceCategoryID);
			self.location.href = "SearchFrame.asp?fraSrc=ServiceType";
			break ;

		default:
			// do nothing ;
	}
}

function btnDelete_onClick() {
//**********************************************************************************************
// Function:	btnDelete_onClick
//
// Purpose:		To delete a service category
//
// Created By:	Gilles Archer 09/27/2000
//
// Updated By:
//***********************************************************************************************
	if ((intAccessLevel & intConst_Access_Delete) != intConst_Access_Delete) {
		alert('You do not have permission to DELETE a Service Category.  Please contact your System Administrator.');
		return false;
	}

	if (document.frmSCategoryDetail.hdnServiceCategoryID.value == "") {
		alert('This Service Category does not exist in the database.');
		return false;
	}

	if (confirm('Do you really want to delete this object?')) {
		document.frmSCategoryDetail.hdnFrmAction.value = "DELETE";
		document.frmSCategoryDetail.submit();
	}
}

function btnNew_onClick() {
	if ((intAccessLevel & intConst_Access_Create) != intConst_Access_Create) {
		alert('You do not have permission to CREATE a Service Category.  Please contact your System Administrator.');
		return false;
	}
	document.location = "SCategoryDetail.asp?ServiceCategoryID=NEW";
}

function fct_onChange() {
	bolSaveRequired = true;
}

function btnSave_onClick() {
	if ((intAccessLevel & intConst_Access_Update) != intConst_Access_Update) {
		alert('You do not have permission to UPDATE a Service Category.  Please contact your System Administrator.');
		return false;
	}

	if (document.frmSCategoryDetail.selLOB.value == "") {
		alert('Please SELECT a Line of Business');
		document.frmSCategoryDetail.selLOB.focus();
		return false;
	}

	if (document.frmSCategoryDetail.txtSCategoryDescription.value == "") {
		alert('Please enter the Service Category Description');
		document.frmSCategoryDetail.txtSCategoryDescription.focus();
		return false;
	}

	document.frmSCategoryDetail.hdnFrmAction.value = "SAVE";
	bolSaveRequired = false;
	document.frmSCategoryDetail.submit();
	return true;
}

function btnReferences_onClick() {
var strOwner = 'CRP';			// owner name must be in Uppercase
var strTableName = 'SERVICE_CATEGORY';		// replace ADDRESS with your own table name and table name must be in Uppercase
var strRecordID = document.frmSCategoryDetail.hdnServiceCategoryID.value ;   // insert your record id
var strURL;

	if (strRecordID == "") {
		alert("No references. This is a new record.");
		return false;
	}

	strURL = "Dependency.asp?Owner=" + strOwner + "&TableName=" + strTableName + "&RecordID=" + strRecordID;
	window.open(strURL, 'Popup', 'top=100, left=100, width=500, height=300');
}

function window_onBeforeUnload() {
	//Ensure that fct_onChange() fires for any changed data.
	document.frmSCategoryDetail.btnSave.focus();

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
	if(confirm('All changes will be lost. Do you really want to reset the page?')){
		bolSaveRequired = false;
		document.location.href = "SCategoryDetail.asp?ServiceCategoryID=<%=strServiceCategoryID%>";
	}
}
// Unhide Client-Side SCRIPT -->
</SCRIPT>
</HEAD>
<BODY language="javascript" onLoad="DisplayStatus(strWinMessage);" onBeforeUnload="window_onBeforeUnload();" onUnload="window_onUnload();">
<FORM id="frmSCategoryDetail" name="frmSCategoryDetail" action="SCategoryDetail.asp" method="post">
	<INPUT type="hidden" id="hdnBusinessID" name="hdnBusinessID" value="<%If IsNumeric(strServiceCategoryID) Then Response.Write objRS.Fields("LOB_ID").Value%>">
	<INPUT type="hidden" id="hdnServiceCategoryID" name="hdnServiceCategoryID" value="<%If IsNumeric(strServiceCategoryID) Then Response.Write objRS.Fields("SERVICE_CATEGORY_ID").Value%>">
	<INPUT type="hidden" id="hdnFrmAction" name="hdnFrmAction" value="">
	<INPUT type="hidden" id="hdnUpdateDateTime" name="hdnUpdateDateTime" value="<%If IsNumeric(strServiceCategoryID) Then Response.Write objRS.Fields("LAST_UPDATE_DATE_TIME").Value%>">
<TABLE border="0" cols="4" width="100%">
<THEAD>
<TR>
	<TD align="left" colspan="3">Service Category Detail</TD>
	<TD align="right"><SELECT valign="top" id="selNavigate" name="selNavigate" onChange="fct_selNavigate();">
		<OPTION value="DEFAULT" selected>Quickly Goto ...</OPTION>
		<OPTION value="LOB">Line of Business</OPTION>
		<OPTION value="STypes">Service Types</OPTION></SELECT></TD>
</TR>
</THEAD>
<TBODY>
<TR>
	<TD align="right" nowrap>Line of Business<FONT color="red">*</FONT></TD>
	<TD align="left" colspan="2" nowrap>
	<SELECT id="selLOB" name="selLOB" style="width: 350px" onChange="fct_onChange();">
		<OPTION></OPTION>
		<%Do While Not objRSSelect.EOF
			If IsNumeric(strServiceCategoryID) Then
				If StrComp(CStr(objRSSelect.Fields("LOB_ID").Value), CStr(objRS.Fields("LOB_ID").Value), 0) = 0 Then
					Response.Write "<OPTION selected value='" & objRSSelect.Fields("LOB_ID").Value & "'>" & objRSSelect.Fields("LOB_CODE").Value & " - " & objRSSelect.Fields("LOB_DESC").Value & "</OPTION>"
				Else
					Response.Write "<OPTION value='" & objRSSelect.Fields("LOB_ID").Value & "'>" & objRSSelect.Fields("LOB_CODE").Value & " - " & objRSSelect.Fields("LOB_DESC").Value & "</OPTION>"
				End If
			Else
				Response.Write "<OPTION value='" & objRSSelect.Fields("LOB_ID").Value & "'>" & objRSSelect.Fields("LOB_CODE").Value & " - " & objRSSelect.Fields("LOB_DESC").Value & "</OPTION>"
			End If
			objRSSelect.MoveNext
		Loop
		objRSSelect.Close
		Set objRSSelect = Nothing%>
	</SELECT></TD>
</TR>
<TR>
	<TD align="right" nowrap>English Description<FONT color="red">*</FONT></TD>
	<TD align="left" colspan="2" nowrap><INPUT id="txtSCategoryDescription" name="txtSCategoryDescription" onChange="fct_onChange();" value="<%If IsNumeric(strServiceCategoryID) Then Response.Write objRS.Fields("SERVICE_CATEGORY_DESC").Value%>" maxlength="80" size="80"></TD>
</TR>
<TR>
	<TD align="right" nowrap>Description Française&nbsp;<br/>French Description<FONT color="red">&nbsp</FONT></TD>
	<TD align="left" colspan="2" nowrap><INPUT id="txtServiceFrench" name="txtServiceFrench" onChange="fct_onChange();" value="<%If IsNumeric(strServiceCategoryID) Then Response.Write objRSFrench.Fields("SERVICE_CATEGORY_LANG_DESC").Value%>" maxlength="80" size="80"></TD>
</TR>
</TBODY>
<TFOOT>
<TR>
	<TD colspan="4" align="right">
	<INPUT id="btnReferences" name="btnReferences" type="button" value="References" style="width: 2.2cm" language="javascript" onClick="return btnReferences_onClick();">&nbsp;
	<INPUT id="btnDelete" name="btnDelete" type="button" value="Delete" style="width: 2cm" language="javascript" onClick="btnDelete_onClick();">&nbsp;
	<INPUT id="btnReset" name="btnReset" type="button" value="Reset" style="width: 2cm" language="javascript" onClick="btnReset_onClick();">&nbsp;
	<INPUT id="btnNew" name="btnNew" type="button" value="New" style="width: 2cm" language="javascript" onClick="btnNew_onClick();">&nbsp;
	<INPUT id="btnSave" name="btnSave" type="button" value="Save" style="width: 2cm"  onClick="return btnSave_onClick();">&nbsp;</TD>
</TR>
</TFOOT>
</TABLE>
<FIELDSET width="100%">
	<LEGEND align="right"><B>Audit Information</B></LEGEND>
	<DIV size="8pt" align="right">
	Record Status Indicator:<INPUT align="left" name="txtRecordStatusInd" type="text" style="width: 18px" disabled value="<%If IsNumeric(strServiceCategoryID) Then Response.Write objRS.Fields("RECORD_STATUS_IND").Value%>">&nbsp;&nbsp;&nbsp;
	Create Date:<INPUT align="center" name="txtRecordStatusInd" type="text" style="width: 150px" disabled value="<%If IsNumeric(strServiceCategoryID) Then Response.Write objRS.Fields("CREATE_DATE_TIME").Value%>">&nbsp;
	Created By:<INPUT align="right" name="txtRecordStatusInd" type="text" style="width: 150px" disabled value="<%If IsNumeric(strServiceCategoryID) Then Response.Write objRS.Fields("CREATE_REAL_USERID").Value%>"><BR>
	Update Date:<INPUT align="center" name="txtRecordStatusInd" type="text" style="width: 150px" disabled value="<%If IsNumeric(strServiceCategoryID) Then Response.Write objRS.Fields("UPDATE_DATE_TIME").Value%>">&nbsp;
	Updated By:<INPUT align="right" name="txtRecordStatusInd" type="text" style="width: 150px" disabled value="<%If IsNumeric(strServiceCategoryID) Then Response.Write objRS.Fields("UPDATE_REAL_USERID").Value%>">
	</DIV>
</FIELDSET>
</FORM>
<%
	'Clean up our ADO objects
	Set objRS = Nothing
	objConn.Close
	Set ObjConn = Nothing
%>
</BODY>
</HTML>

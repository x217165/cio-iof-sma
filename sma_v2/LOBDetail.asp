<%@ Language=VBScript %>
<%Option Explicit%>
<%Response.Buffer = True%>
<!--#include file = "smaConstants.inc"-->
<!--#include file = "smaProcs.inc"-->
<!--#include file = "databaseconnect.asp" -->
<!--
*********************************************************************************************
* Page name:	LOBDetail.asp																*
* Purpose:		To display the Line of Business												*
*				Chosen via LOBList.asp														*
*																							*
* Created by:	Gilles Archer 09/27/2000													*
*																							*
*********************************************************************************************
-->
<%
Dim strBusinessID, datUpdateDateTime, strWinMessage, strWinLocation
Dim	objComm, objRs, objCommand, strSQL, strAdminFlag, strErrMessage
Dim p_userid, p_lob_id, p_lob_code, p_lob_desc, p_lob_account, p_admin_only, p_origin_source, p_last_update_dt, p_lob_french
Dim intAccessLevel

	'Check user's rights
	intAccessLevel = CInt(CheckLogon(strConst_ServiceCatalogue))
	If (intAccessLevel And intConst_Access_ReadOnly) <> intConst_Access_ReadOnly Then
		DisplayError "BACK", "", 0, "ACCESS DENIED", "You don't have access to Line of Business. Please contact your system administrator"
	End If

	strWinMessage = ""
	strBusinessID = Request("BusinessID")

	p_userid = Session("username")

	If IsNumeric(Request.Form("hdnBusinessID")) Then
		p_lob_id = CLng(Request.Form("hdnBusinessID"))
	Else
		p_lob_id = Null
	End If

	If Len(Request.Form("txtLOBCode")) <> 0 Then
		p_lob_code = UCase(Trim(Request.Form("txtLOBCode")))
	Else
		p_lob_code = Null
	End If

	If Len(Request.Form("txtLOBDescription")) <> 0 Then
		p_lob_desc = Trim(Request.Form("txtLOBDescription"))
	Else
		p_lob_desc = Null
	End If

	If Len(Trim(Request.Form("txtLOBFrench"))) <> 0 Then
		p_lob_french = Trim(Request.Form("txtLOBFrench"))
	Else
		p_lob_french = "NULL"
	End If

	If Len(Request.Form("txtLOBAccountCode")) <> 0 Then
		p_lob_account = UCase(Trim(Request.Form("txtLOBAccountCode")))
	Else
		p_lob_account = Null
	End If

	If Len(Request.Form("chkLOBAdminOnly")) <> 0 Then
		p_admin_only = "Y"
	Else
		p_admin_only = "N"
	End If

	If Len(Request("selLOBOriginatingSource")) <> 0 Then
		p_origin_source = Request("selLOBOriginatingSource")
	Else
		p_origin_source = Null
	End If

	If IsDate(Request("hdnUpdateDateTime")) Then
		p_last_update_dt = CDate(Request("hdnUpdateDateTime"))
	Else
		p_last_update_dt = Null
	End If

	'Response.Write "<BR>p_userid: " & p_userid
	'Response.Write "<BR>p_lob_id: " & p_lob_id
	'Response.Write "<BR>p_lob_code: " & p_lob_code
	'Response.Write "<BR>p_lob_desc: " & p_lob_desc
	'Response.Write "<BR>p_lob_account: " & p_lob_account
	'Response.Write "<BR>p_admin_only: " & p_admin_only
	'Response.Write "<BR>p_origin_source: " & p_origin_source
	'Response.Write "<BR>p_last_update_dt: " & p_last_update_dt
	'Response.End

	Select Case UCase(Request("hdnFrmAction"))
		Case "SAVE"
			Set objCommand = Server.CreateObject("ADODB.Command")
			Set objCommand.ActiveConnection = objConn
			objCommand.CommandType = adCmdStoredProc

			If IsNumeric(Request("hdnBusinessID")) Then	'Save existing Line of Business
				If (intAccessLevel And intConst_Access_Update) <> intConst_Access_Update Then
					DisplayError "REFRESH", strWinLocation, 0, "UPDATE DENIED", "You don't have access to update lines of business. Please contact your system administrator"
				End If

				objCommand.CommandText = "sma_sp_userid.spk_sma_admin_inter.sp_lob_update"
				objCommand.Parameters.Append objCommand.CreateParameter("p_userid", adVarChar, adParamInput, 20, p_userid)
				objCommand.Parameters.Append objCommand.CreateParameter("p_lob_id", adNumeric, adParamInput, , p_lob_id)
				objCommand.Parameters.Append objCommand.CreateParameter("p_lob_code", adVarChar, adParamInput, 6, p_lob_code)
				objCommand.Parameters.Append objCommand.CreateParameter("p_lob_desc", adVarChar, adParamInput, 80, p_lob_desc)
				objCommand.Parameters.Append objCommand.CreateParameter("p_lob_french", adVarChar, adParamInput, 80, p_lob_french)
				objCommand.Parameters.Append objCommand.CreateParameter("p_lob_account", adChar, adParamInput, 3, p_lob_account)
				objCommand.Parameters.Append objCommand.CreateParameter("p_admin_only", adChar, adParamInput, 1, p_admin_only)
				objCommand.Parameters.Append objCommand.CreateParameter("p_origin_source", adVarChar, adParamInput, 10, p_origin_source)
				objCommand.Parameters.Append objCommand.CreateParameter("p_last_update_dt", adDBTimeStamp, adParamInput, , p_last_update_dt)

				strErrMessage = "CANNOT UPDATE OBJECT"
			Else									'Create a new Line of Business
				If (intAccessLevel And intConst_Access_Create) <> intConst_Access_Create Then
					DisplayError "REFRESH", strWinLocation, 0, "INSERT DENIED", "You don't have access to create lines of business. Please contact your system administrator"
				End If

				objCommand.CommandText = "sma_sp_userid.spk_sma_admin_inter.sp_lob_insert"
				objCommand.Parameters.Append objCommand.CreateParameter("p_userid", adVarChar, adParamInput, 20, p_userid)
				objCommand.Parameters.Append objCommand.CreateParameter("p_lob_id", adNumeric, adParamOutput, , p_lob_id)
				objCommand.Parameters.Append objCommand.CreateParameter("p_lob_code", adVarChar, adParamInput, 6, p_lob_code)
				objCommand.Parameters.Append objCommand.CreateParameter("p_lob_desc", adVarChar, adParamInput, 80, p_lob_desc)
				objCommand.Parameters.Append objCommand.CreateParameter("p_lob_french", adVarChar, adParamInput, 80, p_lob_french)
				objCommand.Parameters.Append objCommand.CreateParameter("p_lob_account", adChar, adParamInput, 3, p_lob_account)
				objCommand.Parameters.Append objCommand.CreateParameter("p_admin_only", adChar, adParamInput, 1, p_admin_only)
				objCommand.Parameters.Append objCommand.CreateParameter("p_origin_source", adVarChar, adParamInput, 10, p_origin_source)

				strErrMessage = "CANNOT CREATE OBJECT"
			End If

			On Error Resume Next
			objCommand.Execute
			If objConn.Errors.Count <> 0 then
				DisplayError "BACK", "", objConn.Errors(0).NativeError, strErrMessage, objConn.Errors(0).Description
				objConn.Errors.Clear
			End If
			strBusinessID = CStr(objCommand.Parameters("p_lob_id").Value)
			strWinMessage = "Record saved successfully. You can now see the changes you made."

		Case "DELETE"
			If (intAccessLevel And intConst_Access_Delete) <> intConst_Access_Delete Then
				DisplayError "REFRESH", strWinLocation, 0, "DELETE DENIED", "You don't have access to delete lines of business. Please contact your system administrator"
			End If

			Set objCommand = Server.CreateObject("ADODB.Command")
			Set objCommand.ActiveConnection = objConn
			objCommand.CommandType = adCmdStoredProc
			objCommand.CommandText = "sma_sp_userid.spk_sma_admin_inter.sp_lob_delete"
			objCommand.Parameters.Append objCommand.CreateParameter("p_lob_id", adNumeric, adParamInput, , p_lob_id)					'number(9)
			objCommand.Parameters.Append objCommand.CreateParameter("p_last_update_dt", adDBTimeStamp, adParamInput, , p_last_update_dt)		'Date

  			On Error Resume Next
			objCommand.Execute
			If objConn.Errors.Count <> 0 Then
				DisplayError "BACK", "", objConn.Errors(0).NativeError, "CANNOT DELETE OBJECT", objConn.Errors(0).Description
				objConn.Errors.Clear
			End If
			strBusinessID = "DEL"
			strWinMessage = "Record deleted successfully."
	End Select

	If IsNumeric(strBusinessID) Then
		strSQL = "SELECT LOB_ID, LOB_CODE, LOB_DESC, LOB_ACCOUNT_CODE, " &_
				"ADMIN_ONLY_FLAG, ORIGINATING_SOURCE_LCODE, " &_
				"TO_CHAR(CREATE_DATE_TIME, 'MON-DD-YYYY HH24:MI:SS') AS CREATE_DATE_TIME, " &_
				"sma_sp_userid.spk_sma_library.sf_get_full_username(CREATE_REAL_USERID) AS CREATE_REAL_USERID, " &_
				"TO_CHAR(UPDATE_DATE_TIME, 'MON-DD-YYYY HH24:MI:SS') AS UPDATE_DATE_TIME, " &_
				"sma_sp_userid.spk_sma_library.sf_get_full_username(UPDATE_REAL_USERID) AS UPDATE_REAL_USERID, " &_
				"UPDATE_DATE_TIME AS LAST_UPDATE_DATE_TIME, " &_
				"RECORD_STATUS_IND " &_
				"FROM CRP.LOB " &_
				"WHERE LOB_ID = " & strBusinessID &_
				" AND RECORD_STATUS_IND = 'A'"

		'Create Recordset object
		Set objRS = Server.CreateObject("ADODB.Recordset")
		'Create the command object
		Set objComm = Server.CreateObject("ADODB.Command")
		Set objComm.ActiveConnection = objConn
		objComm.CommandText = strSQL
		objComm.CommandType = adCmdText

		On Error Resume Next
		Set objRS = objComm.Execute
		If objConn.Errors.Count <> 0 Then
			DisplayError "BACK", "", objConn.Errors(0).NativeError, "ERROR LOADING DATA", objConn.Errors(0).Description
			objConn.Errors.Clear
		End If
	End If

	Dim objRSSelect

	strSQL = "SELECT ORIGINATING_SOURCE_LCODE, ORIGINATING_SOURCE_DESC " &_
			"FROM CRP.LCODE_ORIGINATING_SOURCE " &_
			"WHERE RECORD_STATUS_IND = 'A' " &_
			"ORDER BY SORT_ORDER ASC"

	'Create Recordset object
	Set objRSSelect = Server.CreateObject("ADODB.Recordset")
	'Create the command object
	Set objComm = Server.CreateObject("ADODB.Command")
	Set objComm.ActiveConnection = objConn
	objComm.CommandText = strSQL
	objComm.CommandType = adCmdText

	On Error Resume Next
	Set objRSSelect = objComm.Execute
	If objConn.Errors.Count <> 0 Then
		DisplayError "BACK", "", objConn.Errors(0).NativeError, "ERROR LOADING DATA", objConn.Errors(0).Description
		objConn.Errors.Clear
	End If

	If IsNumeric(strBusinessID) Then
		Dim objFrench

		strSQL = " SELECT lob_lang_desc" &_
				 " FROM crp.lob_lang" &_
				 " WHERE language_preference_lcode = 'FR'" &_
				 " AND lob_id = " & strBusinessID &_
				 " AND record_status_ind = 'A' "

		Set objFrench = Server.CreateObject("ADODB.Recordset")
		Set objComm = Server.CreateObject("ADODB.Command")
		Set objComm.ActiveConnection = objConn
		objComm.CommandText = strSQL
		objComm.CommandType = adCmdText

		On Error Resume Next
		set objFrench = objComm.Execute
		if objConn.Errors.Count <> 0 Then
			DisplayError "BACK", "", objConn.Errors(0).NativeError, "ERROR LOADING DATA", objConn.Errors(0).Description
			objConn.Errors.Clear
		End If
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
var strWinMessage = "<%=strWinMessage%>";
var intAccessLevel = <%=intAccessLevel%>;
var	bolSaveRequired = false;

setPageTitle("SMA - Line of Business");

function fct_selNavigate(){
//***********************************************************************************************
// Function: selNavigate_onChange																*
//																								*
// Purpose:		To display the page selected by the user from Quick Navigation drop-down box.	*
//																								*
// Created By:	Gills Archer	Oct 01 2000														*
//																								*
// Updated By:																					*
//***********************************************************************************************
var strPageName = document.frmLOBDetail.selNavigate.item(document.frmLOBDetail.selNavigate.selectedIndex).value;

	switch (strPageName) {
		case "SCategories":
			document.frmLOBDetail.selNavigate.selectedIndex = 0;
			SetCookie("BusinessID", "<%=strBusinessID%>");
			document.location.href = "SearchFrame.asp?fraSrc=ServiceCategory";
			break;

		case "STypes":
			document.frmLOBDetail.selNavigate.selectedIndex = 0;
			SetCookie("BusinessID", "<%=strBusinessID%>");
			document.location.href = "SearchFrame.asp?fraSrc=ServiceType";
			break;

		case "DEFAULT":
			// do nothing ;
	}
}

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
	if ((intAccessLevel & intConst_Access_Delete) != intConst_Access_Delete) {
		alert('You do not have permission to DELETE a Line of Business.  Please contact your System Administrator.');
		return false;
	}

	if (document.frmLOBDetail.hdnBusinessID.value == "") {
		alert('This Line of Business does not exist in the database.');
		return false;
	}

	if (confirm('Do you really want to delete this object?')) {
		document.frmLOBDetail.hdnFrmAction.value = "DELETE";
		document.frmLOBDetail.submit();
	}
}

function btnNew_onClick() {
	if ((intAccessLevel & intConst_Access_Create) != intConst_Access_Create) {
		alert('You do not have permission to CREATE a Line of Business.  Please contact your System Administrator.');
		return false;
	}
	document.location ="LOBDetail.asp?BusinessID=NEW";
}

function fct_onChange() {
	bolSaveRequired = true;
}

function btnSave_onClick() {
	if ((intAccessLevel & intConst_Access_Update) != intConst_Access_Update) {
		alert('You do not have permission to UPDATE a Line of Business.  Please contact your System Administrator.');
		return false;
	}

	if (document.frmLOBDetail.txtLOBCode.value == "") {
		alert('Please enter the Line of Business Code');
		document.frmLOBDetail.txtLOBCode.focus();
		return false;
	}

	if (document.frmLOBDetail.txtLOBDescription.value == "") {
		alert('Please enter the Line of Business Description');
		document.frmLOBDetail.txtLOBDescription.focus();
		return false;
	}

	if (document.frmLOBDetail.txtLOBAccountCode.value == "") {
		alert("Please enter the Line of Business Account Code");
		document.frmLOBDetail.txtLOBAccountCode.focus();
		return false;
	}

	if (document.frmLOBDetail.selLOBOriginatingSource.value == "") {
		alert("Please enter the Line of Business Originating Source");
		document.frmLOBDetail.selLOBOriginatingSource.focus();
		return false;
	}

	document.frmLOBDetail.hdnFrmAction.value = "SAVE";
	bolSaveRequired = false;
	document.frmLOBDetail.submit();
	return true;
}

function btnReferences_onClick() {
var strOwner = 'CRP';			// owner name must be in Uppercase
var strTableName = 'LOB';		// replace ADDRESS with your own table name and table name must be in Uppercase
var strRecordID = document.frmLOBDetail.hdnBusinessID.value ;   // insert your record id
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
	document.frmLOBDetail.btnSave.focus();

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
		document.location.href = "LOBDetail.asp?BusinessID=<%=strBusinessID%>";
	}
}
// Unhide Client-Side SCRIPT -->
</SCRIPT>
</HEAD>
<BODY language="javascript" onLoad="DisplayStatus(strWinMessage);" onBeforeUnload="window_onBeforeUnload();" onUnload="window_onUnload();">
<FORM id="frmLOBDetail" name="frmLOBDetail" action="LOBDetail.asp" method="post">
	<INPUT type="hidden" id="hdnBusinessID" name="hdnBusinessID" value="<%If IsNumeric(strBusinessID) Then Response.Write objRS.Fields("LOB_ID").Value%>">
	<INPUT type="hidden" id="hdnFrmAction" name="hdnFrmAction" value="">
	<INPUT type="hidden" id="hdnUpdateDateTime" name="hdnUpdateDateTime" value="<%If IsNumeric(strBusinessID) Then Response.Write objRS.Fields("LAST_UPDATE_DATE_TIME").Value%>">
<TABLE border="0" cols="4" width="100%">
<THEAD>
<TR>
	<TD align="left" colspan="3">Line of Business Detail</TD>
	<TD align="right"><SELECT valign="top" id="selNavigate" name="selNavigate" onChange="fct_selNavigate();">
		<OPTION value="DEFAULT" selected>Quickly Goto ...</OPTION>
		<OPTION value="SCategories">Service Categories</OPTION>
		<OPTION value="STypes">Service Types</OPTION></SELECT>
</TR>
</THEAD>
<TBODY>
<TR>
	<TD align="right" nowrap>Code<FONT color="red">*</FONT></TD>
	<TD align="left" colspan="2" nowrap><INPUT id="txtLOBCode" name="txtLOBCode" onChange="fct_onChange();" value="<%If IsNumeric(strBusinessID) Then Response.Write objRS.Fields("LOB_CODE").Value%>" maxlength="6" size="6"></TD>
</TR>
<TR>
	<TD align="right" nowrap>English Description<FONT color="red">*</FONT></TD>
	<TD align="left" colspan="2" nowrap><INPUT id="txtLOBDescription" name="txtLOBDescription" onChange="fct_onChange();" value="<%If IsNumeric(strBusinessID) Then Response.Write objRS.Fields("LOB_DESC").Value%>" maxlength="80" size="80"></TD>
</TR>
<TR>
	<TD align="right" nowrap>Description Française&nbsp;<br/>French Description<FONT color="red">&nbsp;</FONT></TD>
	<TD align="left" colspan="2" nowrap><INPUT id="txtLOBFrench" name="txtLOBFrench" onChange="fct_onChange();" value="<%If IsNumeric(strBusinessID) Then Response.Write objFrench.Fields("LOB_LANG_DESC").Value%>" maxlength="80" size="80"></TD>
</TR>
<TR>
	<TD align="right" nowrap>Account Code<FONT color="red">*</FONT></TD>
	<TD align="left" colspan="2" nowrap><INPUT id="txtLOBAccountCode" name="txtLOBAccountCode" onChange="fct_onChange();" value="<%If IsNumeric(strBusinessID) Then Response.Write objRS.Fields("LOB_ACCOUNT_CODE").Value%>" maxlength="3" size="3"></TD>
</TR>
<TR>
	<TD align="right" nowrap>Admin Only?<FONT color="red">*</FONT></TD>
	<TD align="left" colspan="2" nowrap><INPUT id="chkLOBAdminOnly" name="chkLOBAdminOnly" onChange="fct_onChange();" type="checkbox" <%If IsNumeric(strBusinessID) Then If objRS.Fields("ADMIN_ONLY_FLAG").Value = "Y" Then Response.Write "checked" End If%>></TD>
</TR>
<TR>
	<TD align="right" nowrap>Originating Source<FONT color="red">*</FONT></TD>
	<TD align="left" colspan="2" nowrap>
	<SELECT id="selLOBOriginatingSource" name="selLOBOriginatingSource" onChange="fct_onChange();">
		<OPTION></OPTION>
		<%Do While Not objRSSelect.EOF
			If IsNumeric(strBusinessID) Then
				If objRSSelect.Fields("ORIGINATING_SOURCE_LCODE").Value = objRS.Fields("ORIGINATING_SOURCE_LCODE").Value Then
					Response.Write "<OPTION selected value='" & objRSSelect.Fields("ORIGINATING_SOURCE_LCODE").Value & "'>" & objRSSelect.Fields("ORIGINATING_SOURCE_DESC").Value & "</OPTION>"
				Else
					Response.Write "<OPTION value='" & objRSSelect.Fields("ORIGINATING_SOURCE_LCODE").Value & "'>" & objRSSelect.Fields("ORIGINATING_SOURCE_DESC").Value & "</OPTION>"
				End If
			Else
				Response.Write "<OPTION value='" & objRSSelect.Fields("ORIGINATING_SOURCE_LCODE").Value & "'>" & objRSSelect.Fields("ORIGINATING_SOURCE_DESC").Value & "</OPTION>"
			End If
			objRSSelect.MoveNext
		Loop%>
	</SELECT>
</TR>
<TFOOT>
<TR>
	<TD colspan="4" align="right">
	<INPUT id="btnReferences" name="btnReferences" type="button" value="References" style="width: 2.2cm" language="javascript" onClick="return btnReferences_onClick();">&nbsp;
	<INPUT id="btnDelete" name="btnDelete" type="button" value="Delete" style="width: 2cm" language="javascript" onClick="btnDelete_onClick();">&nbsp;
	<INPUT id="btnReset" name="btnReset" type="button" value="Reset" style="width: 2cm" language="javascript" onClick="btnReset_onClick();" >&nbsp;
	<INPUT id="btnNew" name="btnNew" type="button" value="New" style="width: 2cm" language="javascript" onClick="btnNew_onClick();">&nbsp;
	<INPUT id="btnSave" name="btnSave" type="button" value="Save" style="width: 2cm" onClick="return btnSave_onClick();">&nbsp;
</TR>
</TFOOT>
</TABLE>

<FIELDSET width="100%">
	<LEGEND align="right"><B>Audit Information</B></LEGEND>
	<DIV size="8pt" align="right">
	Record Status Indicator:<INPUT align="left" name="txtRecordStatusInd" type="text" style="width: 18px" disabled value="<%If IsNumeric(strBusinessID) Then Response.Write objRS.Fields("RECORD_STATUS_IND").Value%>">&nbsp;&nbsp;&nbsp;
	Create Date:<INPUT align="center" name="txtRecordStatusInd" type="text" style="width: 150px" disabled value="<%If IsNumeric(strBusinessID) Then Response.Write objRS.Fields("CREATE_DATE_TIME").Value%>">&nbsp;
	Created By:<INPUT align="right" name="txtRecordStatusInd" type="text" style="width: 150px" disabled value="<%If IsNumeric(strBusinessID) Then Response.Write objRS.Fields("CREATE_REAL_USERID").Value%>"><BR>
	Update Date:<INPUT align="center" name="txtRecordStatusInd" type="text" style="width: 150px" disabled value="<%If IsNumeric(strBusinessID) Then Response.Write objRS.Fields("UPDATE_DATE_TIME").Value%>">&nbsp;
	Updated By:<INPUT align="right" name="txtRecordStatusInd" type="text" style="width: 150px" disabled value="<%If IsNumeric(strBusinessID) Then Response.Write objRS.Fields("UPDATE_REAL_USERID").Value%>">
	</DIV>
</FIELDSET>
</FORM>
<%
	'Clean up our ADO objects
	Set objRS = Nothing
	Set objRSSelect = Nothing
	Set objComm = Nothing
	objConn.Close
	Set ObjConn = Nothing
%>
</BODY>
</HTML>

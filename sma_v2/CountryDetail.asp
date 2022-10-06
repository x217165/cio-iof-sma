<%@ Language=VBScript %>
<%Option Explicit%>
<%Response.Buffer = True%>
<!--#include file = "smaConstants.inc"-->
<!--#include file = "smaProcs.inc"-->
<!--#include file = "databaseconnect.asp" -->
<!--
*********************************************************************************************
* Page name:	CountryDetail.asp															*
* Purpose:		To display the Country Detail												*
*				Chosen via CityDetail.asp or CountryDetail.asp								*
*																							*
* Created by:	Gilles Archer Oct 06 2000													*
*																							*
*********************************************************************************************
-->
<%
Dim strCountryCode, strWinMessage, strWinLocation
Dim	objCommand, objRS, objCountries, strSQL, strErrMessage
Dim p_userid, p_pkey, p_country, p_country_desc, p_sort, p_last_update_dt
Dim intAccessLevel

	'Check user's rights
	intAccessLevel = CInt(CheckLogon(strConst_ServiceCatalogue))
	If (intAccessLevel And intConst_Access_ReadOnly) <> intConst_Access_ReadOnly Then
		DisplayError "BACK", "", 0, "ACCESS DENIED", "You don't have access to Countries. Please contact your system administrator"
	End If

	strWinMessage = ""
	strCountryCode = UCase(Trim(Request("CountryCode")))

	p_userid = Trim(Session("username"))

	If Len(Trim(Request.Form("hdnPrimaryKey"))) > 0 Then
		p_pkey = Trim(Request.Form("hdnPrimaryKey"))
	Else
		p_pkey = Null
	End If

	If Len(Trim(Request.Form("txtCountryCode"))) > 0 Then
		p_country = UCase(Trim(Request.Form("txtCountryCode")))
	Else
		p_country = Null
	End If

	If Len(Trim(Request.Form("txtCountryName"))) > 0 Then
		p_country_desc = Trim(Request.Form("txtCountryName"))
	Else
		p_country_desc = Null
	End If

	If IsNumeric(Request.Form("txtSortOrder")) Then
		p_sort = CInt(Request.Form("txtSortOrder"))
	Else
		p_sort = Null
	End If

	If IsDate(Request.Form("hdnUpdateDateTime")) Then
		p_last_update_dt = CDate(Request.Form("hdnUpdateDateTime"))
	Else
		p_last_update_dt = Null
	End If

	'Response.Write "<BR>p_userid: " & p_userid
	'Response.Write "<BR>p_pkey: " & p_pkey
	'Response.Write "<BR>p_country: " & p_country
	'Response.Write "<BR>p_country_desc: " & p_country_desc
	'Response.Write "<BR>p_sort: " & p_sort
	'Response.Write "<BR>p_last_update_dt: " & p_last_update_dt
	'Response.End

	Select Case UCase(Request("hdnFrmAction"))
		Case "SAVE"
			Set objCommand = Server.CreateObject("ADODB.Command")
			Set objCommand.ActiveConnection = objConn
			objCommand.CommandType = adCmdStoredProc

			If Len(Request("hdnPrimaryKey")) <> 0 Then	'Save existing Country
				If (intAccessLevel And intConst_Access_Update) <> intConst_Access_Update Then
					DisplayError "REFRESH", strWinLocation, 0, "UPDATE DENIED", "You don't have access to update provinces. Please contact your system administrator"
				End If

				objCommand.CommandText = "sma_sp_userid.spk_sma_admin_inter.sp_country_update"
				objCommand.Parameters.Append objCommand.CreateParameter("p_userid", adVarChar, adParamInput, 20, p_userid)
				objCommand.Parameters.Append objCommand.CreateParameter("p_pkey", adVarChar, adParamInput, 2, p_pkey)
				objCommand.Parameters.Append objCommand.CreateParameter("p_country", adChar, adParamInput, 2, p_country)
				objCommand.Parameters.Append objCommand.CreateParameter("p_country_desc", adVarChar, adParamInput, 30, p_country_desc)
				objCommand.Parameters.Append objCommand.CreateParameter("p_sort", adNumeric, adParamInput, , p_sort)
				objCommand.Parameters.Append objCommand.CreateParameter("p_last_update_dt", adDBTimeStamp, adParamInput, , p_last_update_dt)

				strErrMessage = "CANNOT UPDATE OBJECT"
			Else									'Create a new Country
				If (intAccessLevel And intConst_Access_Create) <> intConst_Access_Create Then
					DisplayError "REFRESH", strWinLocation, 0, "INSERT DENIED", "You don't have access to create provinces. Please contact your system administrator"
				End If

				objCommand.CommandText = "sma_sp_userid.spk_sma_admin_inter.sp_country_insert"
				objCommand.Parameters.Append objCommand.CreateParameter("p_userid", adVarChar, adParamInput, 20, p_userid)
				objCommand.Parameters.Append objCommand.CreateParameter("p_pkey", adVarChar, adParamOutput, 2, p_pkey)
				objCommand.Parameters.Append objCommand.CreateParameter("p_country", adChar, adParamInput, 2, p_country)
				objCommand.Parameters.Append objCommand.CreateParameter("p_country_desc", adVarChar, adParamInput, 30, p_country_desc)
				objCommand.Parameters.Append objCommand.CreateParameter("p_sort", adNumeric, adParamInput, , p_sort)

				strErrMessage = "CANNOT CREATE OBJECT"
			End If

			On Error Resume Next
			objCommand.Execute
			If objConn.Errors.Count <> 0 then
				DisplayError "BACK", "", objConn.Errors(0).NativeError, strErrMessage, objConn.Errors(0).Description
				objConn.Errors.Clear
			End If
			strCountryCode =objCommand.Parameters("p_country").Value

			strWinMessage = "Record saved successfully. You can now see the changes you made."

		Case "DELETE"
			If (intAccessLevel And intConst_Access_Delete) <> intConst_Access_Delete Then
				DisplayError "REFRESH", strWinLocation, 0, "DELETE DENIED", "You don't have access to delete provinces. Please contact your system administrator"
			End If

			Set objCommand = Server.CreateObject("ADODB.Command")
			Set objCommand.ActiveConnection = objConn
			objCommand.CommandType = adCmdStoredProc
			objCommand.CommandText = "sma_sp_userid.spk_sma_admin_inter.sp_country_delete"
			objCommand.Parameters.Append objCommand.CreateParameter("p_pkey", adVarChar, adParamInput, 2, p_pkey)
			objCommand.Parameters.Append objCommand.CreateParameter("p_last_update_dt", adDBTimeStamp, adParamInput, , p_last_update_dt)		'Date

  			On Error Resume Next
			objCommand.Execute
			If objConn.Errors.Count <> 0 Then
				DisplayError "BACK", "", objConn.Errors(0).NativeError, "CANNOT DELETE OBJECT", objConn.Errors(0).Description
				objConn.Errors.Clear
			End If
			strCountryCode = ""
			strWinMessage = "Record deleted successfully."
	End Select

	If StrComp(strCountryCode, "NEW", 0) = 0 Then strCountryCode = ""

	If Len(strCountryCode) <> 0 Then
		strSQL = "SELECT COUNTRY_LCODE, " &_
			"COUNTRY_DESC, " &_
			"SORT_ORDER, " &_
			"TO_CHAR(CREATE_DATE_TIME, 'MON-DD-YYYY HH24:MI:SS') AS CREATE_DATE_TIME, " &_
			"sma_sp_userid.spk_sma_library.sf_get_full_username(CREATE_REAL_USERID) AS CREATE_REAL_USERID, " &_
			"TO_CHAR(UPDATE_DATE_TIME, 'MON-DD-YYYY HH24:MI:SS') AS UPDATE_DATE_TIME, " &_
			"sma_sp_userid.spk_sma_library.sf_get_full_username(UPDATE_REAL_USERID) AS UPDATE_REAL_USERID, " &_
			"UPDATE_DATE_TIME AS LAST_UPDATE_DATE_TIME, " &_
			"RECORD_STATUS_IND " &_
			"FROM CRP.LCODE_COUNTRY " &_
			"WHERE RECORD_STATUS_IND = 'A' "

		If Len(strCountryCode) > 0 Then
			strSQL = strSQL & " AND COUNTRY_LCODE = '" & strCountryCode & "'"
		End If

'		Response.Write strSQL
'		Response.End

		'Create Recordset object
		Set objRS = Server.CreateObject("ADODB.Recordset")
		On Error Resume Next
		objRS.Open strSQL, objConn, adOpenForwardOnly, adLockReadOnly
		If objConn.Errors.Count <> 0 Then
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
var bolSaveRequired = false;

setPageTitle("SMA - Country");

function fct_selNavigate(){
//***********************************************************************************************
// Function: selNavigate_onChange																*
//																								*
// Purpose:		To display the page selected by the user from Quick Navigation drop-down box.	*
//																								*
// Created By:	Gills Archer	Oct 06 2000														*
//																								*
// Updated By:																					*
//***********************************************************************************************
var strPageName = document.frmCountryDetail.selNavigate.item(document.frmCountryDetail.selNavigate.selectedIndex).value;
var strCountryCode = document.frmCountryDetail.hdnCountryCode.value;

	switch (strPageName) {
		case "Municipalities":
			document.frmCountryDetail.selNavigate.selectedIndex = 0;
			SetCookie("CountryCode", strCountryCode);
			document.location.href = "SearchFrame.asp?fraSrc=Municipalities";
			break;
		case "Provinces":
			document.frmCountryDetail.selNavigate.selectedIndex = 0;
			SetCookie("CountryCode", strCountryCode);
			document.location.href = "SearchFrame.asp?fraSrc=Municipalities";
			break;
		case "DEFAULT":			//Do Nothing
			break;
		default:				//Do Nothing
			break;
	}
}

function btnDelete_onClick() {
//**********************************************************************************************
// Function:	btnDelete_onClick
//
// Purpose:		To delete a Country
//
// Created By:	Gilles Archer 09/27/2000
//
// Updated By:
//***********************************************************************************************
	if ((intAccessLevel & intConst_Access_Delete) != intConst_Access_Delete) {
		alert('You do not have permission to DELETE a Country.  Please contact your System Administrator.');
		return false;
	}

	if (document.frmCountryDetail.hdnCountryCode.value == "") {
		alert('This Country does not exist in the database.');
		return false;
	}

	if (confirm('Do you really want to delete this object?')) {
		document.frmCountryDetail.hdnFrmAction.value = "DELETE";
		document.frmCountryDetail.submit();
	}
}

function btnNew_onClick() {
	if ((intAccessLevel & intConst_Access_Create) != intConst_Access_Create) {
		alert('You do not have permission to CREATE a Country.  Please contact your System Administrator.');
		return false;
	}
	document.location ="CountryDetail.asp?CountryCode=NEW";
}

function fct_onChange() {
	bolSaveRequired = true;
}

function btnSave_onClick() {
	if ((intAccessLevel & intConst_Access_Update) != intConst_Access_Update) {
		alert('You do not have permission to UPDATE a Country.  Please contact your System Administrator.');
		return false;
	}

	if (document.frmCountryDetail.txtCountryCode.value == "") {
		alert('Please enter the Country code');
		document.frmCountryDetail.txtCountryCode.focus();
		return false;
	}

	if (document.frmCountryDetail.txtCountryName.selectedIndex == 0) {
		alert('Please enter the Country name');
		document.frmCountryDetail.txtCountryName.focus();
		return false;
	}

	if (document.frmCountryDetail.txtSortOrder.value == "") {
		alert('Please enter a Sort Order');
		document.frmCountryDetail.txtSortOrder.focus();
		return false;
	}

	if (isNaN(Number(document.frmCountryDetail.txtSortOrder.value))) {
		alert('Please enter a numerical value for Sort Order');
		document.frmCountryDetail.txtSortOrder.focus();
		return false;
	}

	document.frmCountryDetail.hdnFrmAction.value = "SAVE";
	bolSaveRequired = false;
	document.frmCountryDetail.submit();
	return true;
}

function window_onBeforeUnload() {
	//Ensure that fct_onChange() fires for any changed data.
	document.frmCountryDetail.btnSave.focus();

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
	document.location.href = "CountryDetail.asp?CountryCode=<%=strCountryCode%>";
}
// Unhide Client-Side SCRIPT -->
</SCRIPT>
</HEAD>
<BODY language="javascript" onLoad="DisplayStatus(strWinMessage);" onBeforeUnload="window_onBeforeUnload();" onUnload="window_onUnload();">
<FORM id="frmCountryDetail" name="frmCountryDetail" action="CountryDetail.asp" method="post">
	<INPUT type="hidden" id="hdnPrimaryKey" name="hdnPrimaryKey" value="<%If Len(strCountryCode) <> 0 Then Response.Write objRS.Fields("COUNTRY_LCODE").Value%>">
	<INPUT type="hidden" id="hdnCountryCode" name="hdnCountryCode" value="<%If Len(strCountryCode) <> 0 Then Response.Write objRS.Fields("COUNTRY_LCODE").Value%>">
	<INPUT type="hidden" id="hdnUpdateDateTime" name="hdnUpdateDateTime" value="<%If Len(strCountryCode) <> 0 Then Response.Write objRS.Fields("LAST_UPDATE_DATE_TIME").Value%>">
	<INPUT type="hidden" id="hdnFrmAction" name="hdnFrmAction" value="">
<TABLE border="0" cols="4" width="100%">
<THEAD>
<TR>
	<TD align="left" colspan="3">Country Detail</TD>
	<TD align="right"><SELECT valign="top" id="selNavigate" name="selNavigate" onChange="fct_selNavigate();">
		<OPTION value="DEFAULT" selected>Quickly Goto ...</OPTION>
		<OPTION value="Municipalities">Municipalities</OPTION>
		<OPTION value="Provinces">Provinces / States</OPTION></TD>
</TR>
</THEAD>
<TBODY>
<TR>
	<TD align="right">Country Code<FONT color="red">*</FONT></TD>
	<TD align="left"><INPUT type="text" id="txtCountryCode" name="txtCountryCode" maxlength="2" size="3" onChange="fct_onChange();" value="<%If Len(strCountryCode) <> 0 Then Response.Write objRS.Fields("COUNTRY_LCODE").Value%>"></TD>
</TR>
<TR>
	<TD align="right">Country Name<FONT color="red">*</FONT></TD>
	<TD align="left"><INPUT type="text" id="txtCountryName" name="txtCountryName" maxlength="30" size="30" onChange="fct_onChange();" value="<%If Len(strCountryCode) <> 0 Then Response.Write objRS.Fields("COUNTRY_DESC").Value%>"></TD>
</TR>
<TR>
	<TD align="right">Sort Order<FONT color="red">*</FONT></TD>
	<TD align="left"><INPUT type="text" id="txtSortOrder" name="txtSortOrder" maxlength="3" size="3" onChange="fct_onChange();" value="<%If Len(strCountryCode) <> 0 Then Response.Write objRS.Fields("SORT_ORDER").Value%>"></TD>
</TR>
</TBODY>
<TFOOT>
<TR>
	<TD colspan="4" align="right">
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
	Record Status Indicator:<INPUT align="left" name="txtRecordStatusInd" type="text" style="width: 18px" disabled value="<%If Len(strCountryCode) <> 0 Then Response.Write objRS.Fields("RECORD_STATUS_IND").Value%>">&nbsp;&nbsp;&nbsp;
	Create Date:<INPUT align="center" name="txtRecordStatusInd" type="text" style="width: 150px" disabled value="<%If Len(strCountryCode) <> 0 Then Response.Write objRS.Fields("CREATE_DATE_TIME").Value%>">&nbsp;
	Created By:<INPUT align="right" name="txtRecordStatusInd" type="text" style="width: 150px" disabled value="<%If Len(strCountryCode) <> 0 Then Response.Write objRS.Fields("CREATE_REAL_USERID").Value%>"><BR>
	Update Date:<INPUT align="center" name="txtRecordStatusInd" type="text" style="width: 150px" disabled value="<%If Len(strCountryCode) <> 0 Then Response.Write objRS.Fields("UPDATE_DATE_TIME").Value%>">&nbsp;
	Updated By:<INPUT align="right" name="txtRecordStatusInd" type="text" style="width: 150px" disabled value="<%If Len(strCountryCode) <> 0 Then Response.Write objRS.Fields("UPDATE_REAL_USERID").Value%>">
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

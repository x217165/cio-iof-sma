<%@ Language=VBScript %>
<%Option Explicit%>
<%Response.Buffer = True%>
<!--#include file = "smaConstants.inc"-->
<!--#include file = "smaProcs.inc"-->
<!--#include file = "databaseconnect.asp" -->
<!--
*********************************************************************************************
* Page name:	ProvinceDetail.asp															*
* Purpose:		To display the Province	Detail												*
*				Chosen via ProvinceDetail.asp													*
*																							*
* Created by:	Gilles Archer Oct 06 2000													*
*																							*
*********************************************************************************************
-->
<%
Dim strProvinceCode, strCountryCode, strWinMessage, strWinLocation
Dim	objCommand, objRS, objCountries, strSQL, strErrMessage
Dim p_userid, p_pkey, p_province, p_country, p_province_name, p_sort, p_last_update_dt
Dim intAccessLevel

	'Check user's rights
	intAccessLevel = CInt(CheckLogon(strConst_ServiceCatalogue))
	If (intAccessLevel And intConst_Access_ReadOnly) <> intConst_Access_ReadOnly Then
		DisplayError "BACK", "", 0, "ACCESS DENIED", "You don't have access to Provinces. Please contact your system administrator"
	End If

	strWinMessage = ""
	strProvinceCode = UCase(Trim(Request("ProvinceCode")))
	strCountryCode = UCase(Trim(Request("CountryCode")))

	p_userid = Trim(Session("username"))

	If Len(Trim(Request.Form("hdnPrimaryKey"))) > 0 Then
		p_pkey = Trim(Request.Form("hdnPrimaryKey"))
	Else
		p_pkey = Null
	End If

	If Len(Trim(Request.Form("txtProvinceCode"))) > 0 Then
		p_province = UCase(Trim(Request.Form("txtProvinceCode")))
	Else
		p_province = Null
	End If

	If Len(Trim(Request.Form("txtProvinceName"))) > 0 Then
		p_province_name = Trim(Request.Form("txtProvinceName"))
	Else
		p_province_name = Null
	End If

	If Len(Trim(Request.Form("selCountry"))) > 0 Then
		p_country = Trim(Request.Form("selCountry"))
	Else
		p_country = Null
	End If

	If Len(Request.Form("txtSortOrder")) > 0 And IsNumeric(Request.Form("txtSortOrder")) Then
		p_sort = CLng(Request.Form("txtSortOrder"))
	Else
		p_sort = Null
	End If

	If IsDate(Request.Form("hdnUpdateDateTime")) Then
		p_last_update_dt = CDate(Request.Form("hdnUpdateDateTime"))
	Else
		p_last_update_dt = Null
	End If

	Select Case UCase(Request.Form("hdnFrmAction"))
		Case "SAVE"
			Set objCommand = Server.CreateObject("ADODB.Command")
			Set objCommand.ActiveConnection = objConn
			objCommand.CommandType = adCmdStoredProc

			If Not IsNull(p_pkey) Then	'Save existing Province
				If (intAccessLevel And intConst_Access_Update) <> intConst_Access_Update Then
					DisplayError "REFRESH", strWinLocation, 0, "UPDATE DENIED", "You don't have access to update provinces. Please contact your system administrator"
				End If

				objCommand.CommandText = "sma_sp_userid.spk_sma_admin_inter.sp_province_update"
				objCommand.Parameters.Append objCommand.CreateParameter("p_userid", adVarChar, adParamInput, 20, p_userid)
				objCommand.Parameters.Append objCommand.CreateParameter("p_pkey", adVarChar, adParamInput, 56, p_pkey)
				objCommand.Parameters.Append objCommand.CreateParameter("p_province", adChar, adParamInput, 2, p_province)
				objCommand.Parameters.Append objCommand.CreateParameter("p_country", adChar, adParamInput, 2, p_country)
				objCommand.Parameters.Append objCommand.CreateParameter("p_province_name", adVarChar, adParamInput, 50, p_province_name)
				objCommand.Parameters.Append objCommand.CreateParameter("p_sort", adNumeric, adParamInput, , p_sort)
				objCommand.Parameters.Append objCommand.CreateParameter("p_last_update_dt", adDBTimeStamp, adParamInput, , p_last_update_dt)

				strErrMessage = "CANNOT UPDATE OBJECT"
			Else									'Create a new Province
				If (intAccessLevel And intConst_Access_Create) <> intConst_Access_Create Then
					DisplayError "REFRESH", strWinLocation, 0, "INSERT DENIED", "You don't have access to create provinces. Please contact your system administrator"
				End If

				objCommand.CommandText = "sma_sp_userid.spk_sma_admin_inter.sp_province_insert"
				objCommand.Parameters.Append objCommand.CreateParameter("p_userid", adVarChar, adParamInput, 20, p_userid)
				objCommand.Parameters.Append objCommand.CreateParameter("p_pkey", adVarChar, adParamOutput, 56, Null)
				objCommand.Parameters.Append objCommand.CreateParameter("p_province", adChar, adParamInput, 2, p_province)
				objCommand.Parameters.Append objCommand.CreateParameter("p_country", adChar, adParamInput, 2, p_country)
				objCommand.Parameters.Append objCommand.CreateParameter("p_province_name", adVarChar, adParamInput, 50, p_province_name)
				objCommand.Parameters.Append objCommand.CreateParameter("p_sort", adNumeric, adParamInput, , p_sort)

				strErrMessage = "CANNOT CREATE OBJECT"
			End If

			On Error Resume Next
			objCommand.Execute
			If objConn.Errors.Count <> 0 then
				DisplayError "BACK", "", objConn.Errors(0).NativeError, strErrMessage, objConn.Errors(0).Description
				objConn.Errors.Clear
			End If
			strProvinceCode = objCommand.Parameters("p_province").Value
			strCountryCode =objCommand.Parameters("p_country").Value

			strWinMessage = "Record saved successfully. You can now see the changes you made."

		Case "DELETE"
			If (intAccessLevel And intConst_Access_Delete) <> intConst_Access_Delete Then
				DisplayError "REFRESH", strWinLocation, 0, "DELETE DENIED", "You don't have access to delete provinces. Please contact your system administrator"
			End If

			Set objCommand = Server.CreateObject("ADODB.Command")
			Set objCommand.ActiveConnection = objConn
			objCommand.CommandType = adCmdStoredProc
			objCommand.CommandText = "sma_sp_userid.spk_sma_admin_inter.sp_province_delete"
			objCommand.Parameters.Append objCommand.CreateParameter("p_pkey", adVarChar, adParamInput, 56, p_pkey)
			objCommand.Parameters.Append objCommand.CreateParameter("p_last_update_dt", adDBTimeStamp, adParamInput, , p_last_update_dt)		'Date

  			On Error Resume Next
			objCommand.Execute
			If objConn.Errors.Count <> 0 Then
				DisplayError "BACK", "", objConn.Errors(0).NativeError, "CANNOT DELETE OBJECT", objConn.Errors(0).Description
				objConn.Errors.Clear
			End If
			strProvinceCode = ""
			strWinMessage = "Record deleted successfully."
	End Select

	If StrComp(strProvinceCode, "NEW", 0) = 0 Then strProvinceCode = ""

	If Len(strProvinceCode) <> 0 Then
		strSQL = "SELECT UPPER(P.PROVINCE_STATE_LCODE) AS PROVINCE_STATE_LCODE, " &_
			"P.PROVINCE_STATE_NAME, " &_
			"UPPER(P.COUNTRY_LCODE) AS COUNTRY_LCODE, " &_
			"P.SORT_ORDER, " &_
			"TO_CHAR(P.CREATE_DATE_TIME, 'MON-DD-YYYY HH24:MI:SS') AS CREATE_DATE_TIME, " &_
			"sma_sp_userid.spk_sma_library.sf_get_full_username(P.CREATE_REAL_USERID) AS CREATE_REAL_USERID, " &_
			"TO_CHAR(P.UPDATE_DATE_TIME, 'MON-DD-YYYY HH24:MI:SS') AS UPDATE_DATE_TIME, " &_
			"sma_sp_userid.spk_sma_library.sf_get_full_username(P.UPDATE_REAL_USERID) AS UPDATE_REAL_USERID, " &_
			"P.UPDATE_DATE_TIME AS LAST_UPDATE_DATE_TIME, " &_
			"P.RECORD_STATUS_IND " &_
			"FROM CRP.LCODE_PROVINCE_STATE P, " &_
			"CRP.LCODE_COUNTRY C " &_
			"WHERE P.COUNTRY_LCODE = C.COUNTRY_LCODE " &_
				"AND P.RECORD_STATUS_IND = 'A'"

		If Len(strProvinceCode) > 0 Then
			strSQL = strSQL & " AND UPPER(P.PROVINCE_STATE_LCODE) = '" & strProvinceCode & "'"
		End If

		If Len(strCountryCode) > 0 Then
			strSQL = strSQL & " AND UPPER(P.COUNTRY_LCODE) = '" & strCountryCode & "'"
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

	strSQL = "SELECT UPPER(C.COUNTRY_LCODE) AS COUNTRY_LCODE, " &_
			"C.COUNTRY_DESC " &_
			"FROM CRP.LCODE_COUNTRY C " &_
			"WHERE C.RECORD_STATUS_IND = 'A' " &_
			"ORDER BY C.COUNTRY_LCODE ASC"

	'Create Recordset object
	Set objCountries = Server.CreateObject("ADODB.Recordset")
	On Error Resume Next
	objCountries.Open strSQL, objConn, adOpenForwardOnly, adLockReadOnly
	If objConn.Errors.Count <> 0 Then
		DisplayError "BACK", "", objConn.Errors(0).NativeError, "ERROR LOADING DATA (Countries)", objConn.Errors(0).Description
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
var strWinMessage = "<%=strWinMessage%>";
var intAccessLevel = <%=intAccessLevel%>;
var bolSaveRequired = false;

setPageTitle("SMA - Province");

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
var strPageName = document.frmProvinceDetail.selNavigate.item(document.frmProvinceDetail.selNavigate.selectedIndex).value;
var strProvinceCode = document.frmProvinceDetail.hdnProvinceCode.value;
var strCountryCode = document.frmProvinceDetail.hdnCountryCode.value;

	switch (strPageName) {
		case "Municipalities":
			document.frmProvinceDetail.selNavigate.selectedIndex = 0;
			SetCookie("ProvinceCode", strProvinceCode);
			SetCookie("CountryCode", strCountryCode);
			document.location.href = "SearchFrame.asp?fraSrc=Municipalities";
			break;

		case "Country":
			document.frmProvinceDetail.selNavigate.selectedIndex = 0;
			self.location.href = "CountryDetail.asp?CountryCode=" + strCountryCode;
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
// Purpose:		To delete a Province
//
// Created By:	Gilles Archer 09/27/2000
//
// Updated By:
//***********************************************************************************************
	if ((intAccessLevel & intConst_Access_Delete) != intConst_Access_Delete) {
		alert('You do not have permission to DELETE a Province.  Please contact your System Administrator.');
		return false;
	}

	if (document.frmProvinceDetail.hdnProvinceCode.value == "") {
		alert('This Province does not exist in the database.');
		return false;
	}

	if (confirm('Do you really want to delete this object?')) {
		document.frmProvinceDetail.hdnFrmAction.value = "DELETE";
		document.frmProvinceDetail.submit();
	}
}

function btnNew_onClick() {
	if ((intAccessLevel & intConst_Access_Create) != intConst_Access_Create) {
		alert('You do not have permission to CREATE a Province.  Please contact your System Administrator.');
		return false;
	}
	document.location ="ProvinceDetail.asp?ProvinceCode=NEW";
}

function fct_onChange() {
	bolSaveRequired = true;
}

function btnSave_onClick() {
	if ((intAccessLevel & intConst_Access_Update) != intConst_Access_Update) {
		alert('You do not have permission to UPDATE a Province.  Please contact your System Administrator.');
		return false;
	}

	if (document.frmProvinceDetail.txtProvinceCode.value == "") {
		alert('Please enter the Province code');
		document.frmProvinceDetail.txtProvinceCode.focus();
		return false;
	}

	if (document.frmProvinceDetail.txtProvinceName.value == "") {
		alert('Please enter the Province name');
		document.frmProvinceDetail.txtProvinceName.focus();
		return false;
	}

	if (document.frmProvinceDetail.selCountry.selectedIndex == 0) {
		alert('Please select a Country');
		document.frmProvinceDetail.selCountry.focus();
		return false;
	}

	if (document.frmProvinceDetail.txtSortOrder.value == "") {
		alert('Please enter a Sort Order');
		document.frmProvinceDetail.txtSortOrder.focus();
		return false;
	}

	if (isNaN(Number(document.frmProvinceDetail.txtSortOrder.value))) {
		alert('Please enter a numerical value for Sort Order');
		document.frmProvinceDetail.txtSortOrder.focus();
		return false;
	}

	document.frmProvinceDetail.hdnFrmAction.value = "SAVE";
	bolSaveRequired = false;
	document.frmProvinceDetail.submit();
	return true;
}

function window_onBeforeUnload() {
	//Ensure that fct_onChange() fires for any changed data.
	document.frmProvinceDetail.btnSave.focus();

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
	document.location.href = "ProvinceDetail.asp?ProvinceCode=<%=strProvinceCode%>&CountryCode=<%=strCountryCode%>";
}
// Unhide Client-Side SCRIPT -->
</SCRIPT>
</HEAD>
<BODY language="javascript" onLoad="DisplayStatus(strWinMessage);" onBeforeUnload="window_onBeforeUnload();" onUnload="window_onUnload();">
<FORM id="frmProvinceDetail" name="frmProvinceDetail" action="ProvinceDetail.asp" method="post">
	<INPUT type="hidden" id="hdnPrimaryKey" name="hdnPrimaryKey" value="<%If Len(strProvinceCode) <> 0 Then Response.Write objRS.Fields("PROVINCE_STATE_LCODE").Value & "~" & objRS.Fields("COUNTRY_LCODE").Value%>">
	<INPUT type="hidden" id="hdnProvinceCode" name="hdnProvinceCode" value="<%If Len(strProvinceCode) <> 0 Then Response.Write objRS.Fields("PROVINCE_STATE_LCODE").Value%>">
	<INPUT type="hidden" id="hdnCountryCode" name="hdnCountryCode" value="<%If Len(strProvinceCode) <> 0 Then Response.Write objRS.Fields("COUNTRY_LCODE").Value%>">
	<INPUT type="hidden" id="hdnUpdateDateTime" name="hdnUpdateDateTime" value="<%If Len(strProvinceCode) <> 0 Then Response.Write objRS.Fields("LAST_UPDATE_DATE_TIME").Value%>">
	<INPUT type="hidden" id="hdnFrmAction" name="hdnFrmAction" value="">
<TABLE border="0" cols="4" width="100%">
<THEAD>
<TR>
	<TD align="left" colspan="3">Province Detail</TD>
	<TD align="right"><SELECT valign="top" id="selNavigate" name="selNavigate" onChange="fct_selNavigate();">
		<OPTION value="DEFAULT" selected>Quickly Goto ...</OPTION>
		<OPTION value="Municipalities">Municipalities</OPTION>
		<OPTION value="Country">Country</OPTION></TD>
</TR>
</THEAD>
<TBODY>
<TR>
	<TD align="right">Province Code<FONT color="red">*</FONT></TD>
	<TD align="left"><INPUT type="text" id="txtProvinceCode" name="txtProvinceCode" maxlength="2" size="3" onChange="fct_onChange();" value="<%If Len(strProvinceCode) <> 0 Then Response.Write objRS.Fields("PROVINCE_STATE_LCODE").Value%>"></TD>
</TR>
<TR>
	<TD align="right">Province Name<FONT color="red">*</FONT></TD>
	<TD align="left"><INPUT type="text" id="txtProvinceName" name="txtProvinceName" maxlength="50" size="50" onChange="fct_onChange();" value="<%If Len(strProvinceCode) <> 0 Then Response.Write objRS.Fields("PROVINCE_STATE_NAME").Value%>"></TD>
</TR>
<TR>
	<TD align="right">Country<FONT color="red">*</FONT></TD>
	<TD align="left">
	<SELECT id="selCountry" name="selCountry" style="width: 200" onChange="fct_onChange();">
		<OPTION></OPTION>
		<%Do While Not objCountries.EOF
			If Len(strProvinceCode) <> 0 Then
				If StrComp(objCountries.Fields("COUNTRY_LCODE").Value, objRS.Fields("COUNTRY_LCODE").Value, 0) = 0 Then
					Response.Write "<OPTION selected value='" & objCountries.Fields("COUNTRY_LCODE").Value & "'>" & objCountries.Fields("COUNTRY_LCODE").Value & " - " & objCountries.Fields("COUNTRY_DESC").Value & "</OPTION>"
				Else
					Response.Write "<OPTION value='" & objCountries.Fields("COUNTRY_LCODE").Value & "'>" & objCountries.Fields("COUNTRY_LCODE").Value & " - " & objCountries.Fields("COUNTRY_DESC").Value & "</OPTION>"
				End If
			Else
				Response.Write "<OPTION value='" & objCountries.Fields("COUNTRY_LCODE").Value & "'>" & objCountries.Fields("COUNTRY_LCODE").Value & " - " & objCountries.Fields("COUNTRY_DESC").Value & "</OPTION>"
			End If
			objCountries.MoveNext
		Loop
		objCountries.Close
		Set objCountries = Nothing%>
	</SELECT></TD>
</TR>
<TR>
	<TD align="right">Sort Order<FONT color="red">*</FONT></TD>
	<TD align="left"><INPUT type="text" id="txtSortOrder" name="txtSortOrder" maxlength="3" size="3" onChange="fct_onChange();" value="<%If Len(strProvinceCode) <> 0 Then Response.Write objRS.Fields("SORT_ORDER").Value%>"></TD>
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
	Record Status Indicator:<INPUT align="left" name="txtRecordStatusInd" type="text" style="width: 18px" disabled value="<%If Len(strProvinceCode) <> 0 Then Response.Write objRS.Fields("RECORD_STATUS_IND").Value%>">&nbsp;&nbsp;&nbsp;
	Create Date:<INPUT align="center" name="txtRecordStatusInd" type="text" style="width: 150px" disabled value="<%If Len(strProvinceCode) <> 0 Then Response.Write objRS.Fields("CREATE_DATE_TIME").Value%>">&nbsp;
	Created By:<INPUT align="right" name="txtRecordStatusInd" type="text" style="width: 150px" disabled value="<%If Len(strProvinceCode) <> 0 Then Response.Write objRS.Fields("CREATE_REAL_USERID").Value%>"><BR>
	Update Date:<INPUT align="center" name="txtRecordStatusInd" type="text" style="width: 150px" disabled value="<%If Len(strProvinceCode) <> 0 Then Response.Write objRS.Fields("UPDATE_DATE_TIME").Value%>">&nbsp;
	Updated By:<INPUT align="right" name="txtRecordStatusInd" type="text" style="width: 150px" disabled value="<%If Len(strProvinceCode) <> 0 Then Response.Write objRS.Fields("UPDATE_REAL_USERID").Value%>">
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

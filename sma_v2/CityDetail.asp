<%@ Language=VBScript %>
<%Option Explicit%>
<%Response.Buffer = True%>
<!--#include file = "smaConstants.inc"-->
<!--#include file = "smaProcs.inc"-->
<!--#include file = "databaseconnect.asp" -->
<!--
*********************************************************************************************
* Page name:	CityDetail.asp																*
* Purpose:		To display the City	Detail													*
*				Chosen via CityList.asp														*
*																							*
* Created by:	Gilles Archer Oct 06 2000													*
*																							*
*********************************************************************************************
        Date		Author			Changes/enhancements made
        -----		------		------------------------------------------------------
       20-Jul-01	 DTy		When displaying province/state, check country code as
								well.  Otherwise, a wrong province/state from another country
								is selected.
       18-Jan-06	ACheung		Add customer managed service region as an edited field.
*********************************************************************************************
-->
<%
Dim strCityName, strProvinceCode, strCountryCode, strWinMessage, strWinLocation	, strServiceRegion
Dim	objCommand, objRS, objProvinces, objCountries, strSQL, strErrMessage, rsSRrp, sql
Dim p_userid, p_pkey, p_name, p_province, p_country, p_last_update_dt, p_clli_code, p_service_region
Dim intAccessLevel

	'Check user's rights
	intAccessLevel = CInt(CheckLogon(strConst_ServiceCatalogue))
	If (intAccessLevel And intConst_Access_ReadOnly) <> intConst_Access_ReadOnly Then
		DisplayError "BACK", "", 0, "ACCESS DENIED", "You don't have access to Municipalities. Please contact your system administrator"
	End If

	strWinMessage = ""
	strCityName = Trim(Request("CityName"))
	strProvinceCode = Trim(Request("ProvinceCode"))
	strCountryCode = Trim(Request("CountryCode"))
	strServiceRegion = Trim(Request("ServiceRegCode"))
'UB:
'	strServiceRegion = Trim(Request("selServiceRegion"))

	p_userid = Trim(Session("username"))

	If Len(Request.Form("hdnPrimaryKey")) <> 0 Then
		p_pkey = Request.Form("hdnPrimaryKey")
	Else
		p_pkey = Null
	End If

	If Len(Trim(Request.Form("txtCityName"))) <> 0 Then
		p_name = Trim(Request.Form("txtCityName"))
	Else
		p_name = Null
	End If

	If Len(Trim(Request.Form("selProvince"))) <> 0 Then
		p_province = Trim(Request.Form("selProvince"))
	Else
		p_province = Null
	End If

	If Len(Trim(Request.Form("selCountry"))) <> 0 Then
		p_country = Trim(Request.Form("selCountry"))
	Else
		p_country = Null
	End If

	If Len(CLng(Request("selServiceRegion"))) <> 0 Then
		p_service_region = CLng(Request("selServiceRegion"))
	Else
		p_service_region = Null
	End If

	If IsDate(Request.Form("hdnUpdateDateTime")) Then
		p_last_update_dt = CDate(Request.Form("hdnUpdateDateTime"))
	Else
		p_last_update_dt = Null
	End If

	If Len(Trim(Request.Form("txtCLLICode"))) <> 0 Then
		p_clli_code = Replace(UCase(Trim(Request.Form("txtCLLICode"))), "'", "''")
	Else
		p_clli_code = Null
	End If

'Response.Write selServiceRegion
'Response.End

	Select Case UCase(Request("hdnFrmAction"))
		Case "SAVE"
			Set objCommand = Server.CreateObject("ADODB.Command")
			Set objCommand.ActiveConnection = objConn
			objCommand.CommandType = adCmdStoredProc

			If Len(Request("hdnPrimaryKey")) <> 0 Then	'Save existing Municipality
				If (intAccessLevel And intConst_Access_Update) <> intConst_Access_Update Then
					DisplayError "REFRESH", strWinLocation, 0, "UPDATE DENIED", "You don't have access to update lines of business. Please contact your system administrator"
				End If

				objCommand.CommandText = "sma_sp_userid.spk_sma_admin_inter.sp_munic_update"
				objCommand.Parameters.Append objCommand.CreateParameter("p_userid", adVarChar, adParamInput, 20, p_userid)
				objCommand.Parameters.Append objCommand.CreateParameter("p_pkey", adVarChar, adParamInput, 60, p_pkey)
				objCommand.Parameters.Append objCommand.CreateParameter("p_name", adVarChar, adParamInput, 50, p_name)
				objCommand.Parameters.Append objCommand.CreateParameter("p_province", adChar, adParamInput, 2, p_province)
				objCommand.Parameters.Append objCommand.CreateParameter("p_country", adChar, adParamInput, 2, p_country)
				objCommand.Parameters.Append objCommand.CreateParameter("p_last_update_dt", adDBTimeStamp, adParamInput, , p_last_update_dt)
				objCommand.Parameters.Append objCommand.CreateParameter("p_clli_code", adChar, adParamInput, 4, p_clli_code)
				objCommand.Parameters.Append objCommand.CreateParameter("p_service_region", adNumeric, adParamInput, 5, p_service_region)  'UB Service region

				strErrMessage = "CANNOT UPDATE OBJECT"
			Else									'Create a new Municipality
				If (intAccessLevel And intConst_Access_Create) <> intConst_Access_Create Then
					DisplayError "REFRESH", strWinLocation, 0, "INSERT DENIED", "You don't have access to create lines of business. Please contact your system administrator"
				End If

				objCommand.CommandText = "sma_sp_userid.spk_sma_admin_inter.sp_munic_insert"
				objCommand.Parameters.Append objCommand.CreateParameter("p_userid", adVarChar, adParamInput, 20, p_userid)
				objCommand.Parameters.Append objCommand.CreateParameter("p_pkey", adVarChar, adParamOutput, 60, Null)
				objCommand.Parameters.Append objCommand.CreateParameter("p_name", adVarChar, adParamInput, 50, p_name)
				objCommand.Parameters.Append objCommand.CreateParameter("p_province", adChar, adParamInput, 2, p_province)
				objCommand.Parameters.Append objCommand.CreateParameter("p_country", adChar, adParamInput, 2, p_country)
				objCommand.Parameters.Append objCommand.CreateParameter("p_clli_code", adChar, adParamInput, 4, p_clli_code)
				objCommand.Parameters.Append objCommand.CreateParameter("p_service_region", adNumeric, adParamInput, 5, p_service_region)  'UB Service region

				strErrMessage = "CANNOT CREATE OBJECT"
			End If

			On Error Resume Next
			objCommand.Execute
			If objConn.Errors.Count <> 0 then
				DisplayError "BACK", "", objConn.Errors(0).NativeError, strErrMessage, objConn.Errors(0).Description
				objConn.Errors.Clear
			End If
			strCityName = objCommand.Parameters("p_name").Value
			strProvinceCode = objCommand.Parameters("p_province").Value
			strCountryCode = objCommand.Parameters("p_country").Value
			strServiceRegion = objCommand.Parameters("p_service_region").Value

			strWinMessage = "Record saved successfully. You can now see the changes you made."

		Case "DELETE"
			If (intAccessLevel And intConst_Access_Delete) <> intConst_Access_Delete Then
				DisplayError "REFRESH", strWinLocation, 0, "DELETE DENIED", "You don't have access to delete lines of business. Please contact your system administrator"
			End If

			Set objCommand = Server.CreateObject("ADODB.Command")
			Set objCommand.ActiveConnection = objConn
			objCommand.CommandType = adCmdStoredProc

			objCommand.CommandText = "sma_sp_userid.spk_sma_admin_inter.sp_munic_delete"
			objCommand.Parameters.Append objCommand.CreateParameter("p_pkey", adVarChar, adParamInput, 60, p_pkey)
			objCommand.Parameters.Append objCommand.CreateParameter("p_last_update_dt", adDBTimeStamp, adParamInput, , p_last_update_dt)		'Date

  			On Error Resume Next
			objCommand.Execute
			If objConn.Errors.Count <> 0 Then
				DisplayError "BACK", "", objConn.Errors(0).NativeError, "CANNOT DELETE OBJECT", objConn.Errors(0).Description
				objConn.Errors.Clear
			End If
			strCityName = ""
			strWinMessage = "Record deleted successfully."
	End Select

	If StrComp(strCityName, "NEW", 0) = 0 Then strCityName = ""

	If Len(strCityName) <> 0 Then
		strSQL = "SELECT M.MUNICIPALITY_NAME, " &_
				"M.CLLI_CODE, " &_
				"UPPER(M.PROVINCE_STATE_LCODE) AS PROVINCE_STATE_LCODE, " &_
				"UPPER(M.COUNTRY_LCODE) AS COUNTRY_LCODE, " &_
				"TO_CHAR(M.CREATE_DATE_TIME, 'MON-DD-YYYY HH24:MI:SS') AS CREATE_DATE_TIME, " &_
				"sma_sp_userid.spk_sma_library.sf_get_full_username(M.CREATE_REAL_USERID) AS CREATE_REAL_USERID, " &_
				"TO_CHAR(M.UPDATE_DATE_TIME, 'MON-DD-YYYY HH24:MI:SS') AS UPDATE_DATE_TIME, " &_
				"sma_sp_userid.spk_sma_library.sf_get_full_username(M.UPDATE_REAL_USERID) AS UPDATE_REAL_USERID, " &_
				"M.UPDATE_DATE_TIME AS LAST_UPDATE_DATE_TIME, " &_
				"M.RECORD_STATUS_IND, " &_
				"SR.CUST_MGD_SRVC_RGN_NAME AS CUST_MGD_SRVC_RGN_NAME, " &_
				"M.CUST_MGD_SRVC_RGN_LCODE AS CUST_MGD_SRVC_RGN_LCODE " &_
				"FROM CRP.MUNICIPALITY_LOOKUP M, " &_
				"CRP.LCODE_PROVINCE_STATE P, " &_
				"CRP.LCODE_COUNTRY C, " &_
				"CRP.LCODE_CUST_MGD_SRVC_RGN SR " &_
				"WHERE M.PROVINCE_STATE_LCODE = P.PROVINCE_STATE_LCODE " &_
				"AND M.COUNTRY_LCODE = P.COUNTRY_LCODE " &_
				"AND P.COUNTRY_LCODE = C.COUNTRY_LCODE " &_
				"AND M.CUST_MGD_SRVC_RGN_LCODE = SR.CUST_MGD_SRVC_RGN_LCODE " &_
				"AND M.RECORD_STATUS_IND = 'A'"

		If Len(strCityName) > 0 Then
			strSQL = strSQL & " AND M.MUNICIPALITY_NAME = '" & Replace(strCityName, "'", "''") & "'"
		End If

		If Len(strProvinceCode) > 0 Then
			strSQL = strSQL & " AND M.PROVINCE_STATE_LCODE = '" & strProvinceCode & "'"
		End If

		If Len(strCountryCode) > 0 Then
			strSQL = strSQL & " AND M.COUNTRY_LCODE = '" & strCountryCode & "'"
		End If

		If Len(strServiceRegion) > 0 Then
			strSQL = strSQL & " AND M.CUST_MGD_SRVC_RGN_LCODE = '" & strServiceRegion & "'"
		End If


		'Response.Write strSQL
		'Response.End

		'Create Recordset object
		Set objRS = Server.CreateObject("ADODB.Recordset")
		On Error Resume Next
		objRS.Open strSQL, objConn, adOpenForwardOnly, adLockReadOnly
		If objConn.Errors.Count <> 0 Then
			DisplayError "BACK", "", objConn.Errors(0).NativeError, "ERROR LOADING DATA", objConn.Errors(0).Description
			objConn.Errors.Clear
		End If
	End If

	strSQL = "SELECT UPPER(PS.PROVINCE_STATE_LCODE) AS PROVINCE_STATE_LCODE, " &_
			"UPPER(PS.COUNTRY_LCODE) AS COUNTRY_LCODE, " &_
			"PS.PROVINCE_STATE_NAME " &_
			"FROM CRP.LCODE_PROVINCE_STATE PS " &_
			"WHERE PS.RECORD_STATUS_IND = 'A' " &_
			"ORDER BY PS.COUNTRY_LCODE ASC, PS.PROVINCE_STATE_LCODE ASC"

	'Create Recordset object
	On Error Resume Next
	Set objProvinces = Server.CreateObject("ADODB.Recordset")
	objProvinces.Open strSQL, objConn, adOpenForwardOnly, adLockReadOnly
	If objConn.Errors.Count <> 0 Then
		DisplayError "BACK", "", objConn.Errors(0).NativeError, "ERROR LOADING DATA (Provinces)", objConn.Errors(0).Description
		objConn.Errors.Clear
	End If

	strSQL = "SELECT UPPER(C.COUNTRY_LCODE) AS COUNTRY_LCODE, " &_
			"C.COUNTRY_DESC " &_
			"FROM CRP.LCODE_COUNTRY C " &_
			"WHERE C.RECORD_STATUS_IND = 'A' " &_
			"ORDER BY C.COUNTRY_LCODE ASC"

	'Create Recordset object
	On Error Resume Next
	Set objCountries = Server.CreateObject("ADODB.Recordset")
	objCountries.Open strSQL, objConn, adOpenForwardOnly, adLockReadOnly
	If objConn.Errors.Count <> 0 Then
		DisplayError "BACK", "", objConn.Errors(0).NativeError, "ERROR LOADING DATA (Countries)", objConn.Errors(0).Description
		objConn.Errors.Clear
	End If
'UB: get the Customer Managed Service Region
	sql = "select CUST_MGD_SRVC_RGN_NAME, CUST_MGD_SRVC_RGN_LCODE from CRP.LCODE_CUST_MGD_SRVC_RGN where RECORD_STATUS_IND='A' ORDER BY CUST_MGD_SRVC_RGN_LCODE"
	set rsSRrp = Server.CreateObject("ADODB.Recordset")
	rsSRrp.CursorLocation = adUseClient
	rsSRrp.Open sql, objConn
	if err then
		DisplayError "BACK", "", err.Number, "CANNOT CREATE RECORDSET", err.Description
	end if

	if rsSRrp.EOF then
		DisplayError "BACK", "", 999, "CANNOT CREATE CONTACT ROLE LIST", "EOF condition occured in rsSRrp recordset."
	end if

	'release the active connection, keep the recordset open
	set rsSRrp.ActiveConnection = nothing
'UB: end
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

setPageTitle("SMA - Municipality");

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
var strPageName = document.frmCityDetail.selNavigate.item(document.frmCityDetail.selNavigate.selectedIndex).value;
var strProvinceCode = document.frmCityDetail.hdnProvinceCode.value;
var strCountryCode = document.frmCityDetail.hdnCountryCode.value;

	switch (strPageName) {
		case "Province":
			document.frmCityDetail.selNavigate.selectedIndex = 0;
			self.location.href = "ProvinceDetail.asp?ProvinceCode=" + strProvinceCode + "&CountryCode=" + strCountryCode;
			break;

		case "Country":
			document.frmCityDetail.selNavigate.selectedIndex = 0;
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
// Purpose:		To delete a line of business
//
// Created By:	Gilles Archer 09/27/2000
//
// Updated By:
//***********************************************************************************************
	if ((intAccessLevel & intConst_Access_Delete) != intConst_Access_Delete) {
		alert('You do not have permission to DELETE a Municipality.  Please contact your System Administrator.');
		return false;
	}

	if (document.frmCityDetail.hdnCityName.value == "") {
		alert('This Municipality does not exist in the database.');
		return false;
	}

	if (confirm('Do you really want to delete this object?')) {
		document.frmCityDetail.hdnFrmAction.value = "DELETE";
		document.frmCityDetail.submit();
	}
}

function btnNew_onClick() {
	if ((intAccessLevel & intConst_Access_Create) != intConst_Access_Create) {
		alert('You do not have permission to CREATE a Municipality.  Please contact your System Administrator.');
		return false;
	}
	document.location ="CityDetail.asp?CityName=NEW";
}

function fct_onChange() {
	bolSaveRequired = true;
}

function btnSave_onClick() {
	if ((intAccessLevel & intConst_Access_Update) != intConst_Access_Update) {
		alert('You do not have permission to UPDATE a Municipality.  Please contact your System Administrator.');
		return false;
	}

	if (document.frmCityDetail.txtCityName.value == "") {
		alert('Please enter the City / Municipality name');
		document.frmCityDetail.txtCityName.focus();
		return false;
	}

/*	if (document.frmCityDetail.txtCLLICode.value == "") {
		alert('Please enter the CLLI Code');
		document.frmCityDetail.txtCLLICode.focus();
		return false;
	}*/

	if (document.frmCityDetail.selProvince.selectedIndex == 0) {
		alert('Please select a Province / State');
		document.frmCityDetail.selProvince.focus();
		return false;
	}

	if (document.frmCityDetail.selCountry.selectedIndex == 0) {
		alert('Please select a Country');
		document.frmCityDetail.selCountry.focus();
		return false;
	}

	document.frmCityDetail.hdnFrmAction.value = "SAVE";
	bolSaveRequired = false;
	document.frmCityDetail.submit();
	return true;
}

function window_onBeforeUnload() {
	//Ensure that fct_onChange() fires for any changed data.
	document.frmCityDetail.btnSave.focus();

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
		document.location.href = "CityDetail.asp?CityName=<%=strCityName%>&Province=<%=strProvinceCode%>&Country=<%=strCountryCode%>&ServiceRegCode=<%=strServiceRegion%>";
	}
}
// Unhide Client-Side SCRIPT -->
</SCRIPT>
</HEAD>
<BODY language="javascript" onLoad="DisplayStatus(strWinMessage);" onBeforeUnload="window_onBeforeUnload();" onUnload="window_onUnload();">
<FORM id="frmCityDetail" name="frmCityDetail" action="CityDetail.asp" method="post">
	<INPUT type="hidden" id="hdnCityName" name="hdnCityName" value="<%If Len(strCityName) <> 0 Then Response.Write objRS.Fields("MUNICIPALITY_NAME").Value%>">
	<INPUT type="hidden" id="hdnPrimaryKey" name="hdnPrimaryKey" value="<%If Len(strCityName) <> 0 Then Response.Write objRS.Fields("MUNICIPALITY_NAME").Value & "~" & objRS.Fields("PROVINCE_STATE_LCODE").Value & "~" & objRS.Fields("COUNTRY_LCODE").Value%>">
	<INPUT type="hidden" id="hdnProvinceCode" name="hdnProvinceCode" value="<%If Len(strCityName) <> 0 Then Response.Write objRS.Fields("PROVINCE_STATE_LCODE").Value%>">
	<INPUT type="hidden" id="hdnCountryCode" name="hdnCountryCode" value="<%If Len(strCityName) <> 0 Then Response.Write objRS.Fields("COUNTRY_LCODE").Value%>">
	<INPUT type="hidden" id="hdnServiceRegCode" name="hdnServiceRegCode" value="<%If Len(strCityName) <> 0 Then Response.Write objRS.Fields("CUST_MGD_SRVC_RGN_LCODE").Value%>">
	<INPUT type="hidden" id="hdnUpdateDateTime" name="hdnUpdateDateTime" value="<%If Len(strCityName) <> 0 Then Response.Write objRS.Fields("LAST_UPDATE_DATE_TIME").Value%>">
	<INPUT type="hidden" id="hdnFrmAction" name="hdnFrmAction" value="">
<TABLE border="0" cols="4" width="100%">
<THEAD>
<TR>
	<TD align="left" colspan="3">Municipality Detail</TD>
	<TD align="right"><SELECT valign="top" id="selNavigate" name="selNavigate" onChange="fct_selNavigate();">
		<OPTION value="DEFAULT" selected>Quickly Goto ...</OPTION>
		<OPTION value="Province">Province/State</OPTION>
		<OPTION value="Country">Country</OPTION></TD>
</TR>
</THEAD>
<TBODY>
<TR>
	<TD align="right">City / Municipality Name<FONT color="red">*</FONT></TD>
	<TD align="left"><INPUT type="text" id="txtCityName" name="txtCityName" maxlength="50" size="50" onChange="fct_onChange();" value="<%If Len(strCityName) <> 0 Then Response.Write objRS.Fields("MUNICIPALITY_NAME").Value%>"></TD>
</TR>
<TR>
	<TD align="right">CLLI Code</TD>
	<TD align="left"><INPUT type="text" id="txtCLLICode" name="txtCLLICode" maxlength="4" size="5" onChange="fct_onChange();" value="<%If Len(strCityName) <> 0 Then Response.Write objRS.Fields("CLLI_CODE").Value%>"></TD>
</TR>
<TR>
	<TD align="right">Province / State<FONT color="red">*</FONT></TD>
	<TD align="left">
	<SELECT id="selProvince" name="selProvince" style="width: 200" onChange="fct_onChange();">
		<OPTION></OPTION>
		<%Do While Not objProvinces.EOF
			If Len(strCityName) <> 0 Then
				If StrComp(objProvinces.Fields("COUNTRY_LCODE").Value, objRS.Fields("COUNTRY_LCODE").Value, 0) = 0 then
				   if StrComp(objProvinces.Fields("PROVINCE_STATE_LCODE").Value, objRS.Fields("PROVINCE_STATE_LCODE").Value, 0) = 0 Then
					  Response.Write "<OPTION selected value='" & objProvinces.Fields("PROVINCE_STATE_LCODE").Value & "'>" & objProvinces.Fields("PROVINCE_STATE_LCODE").Value & " - " & objProvinces.Fields("PROVINCE_STATE_NAME").Value & "</OPTION>"
				   Else
					  Response.Write "<OPTION value='" & objProvinces.Fields("PROVINCE_STATE_LCODE").Value & "'>" & objProvinces.Fields("PROVINCE_STATE_LCODE").Value & " - " & objProvinces.Fields("PROVINCE_STATE_NAME").Value & "</OPTION>"
				   end if
				Else
					Response.Write "<OPTION value='" & objProvinces.Fields("PROVINCE_STATE_LCODE").Value & "'>" & objProvinces.Fields("PROVINCE_STATE_LCODE").Value & " - " & objProvinces.Fields("PROVINCE_STATE_NAME").Value & "</OPTION>"
				End If
			Else
				Response.Write "<OPTION value='" & objProvinces.Fields("PROVINCE_STATE_LCODE").Value & "'>" & objProvinces.Fields("PROVINCE_STATE_LCODE").Value & " - " & objProvinces.Fields("PROVINCE_STATE_NAME").Value & "</OPTION>"
			End If
			objProvinces.MoveNext
		Loop
		objProvinces.Close
		Set objProvinces = Nothing%>
	</SELECT></TD>
</TR>
<TR>
	<TD align="right">Country<FONT color="red">*</FONT></TD>
	<TD align="left">
	<SELECT id="selCountry" name="selCountry" style="width: 200" onChange="fct_onChange();">
		<OPTION></OPTION>
		<%Do While Not objCountries.EOF
			If Len(strCityName) <> 0 Then
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
	<TD align="right">Customer Managed Service Region</TD>
        <td align="left">
	<SELECT id=selServiceRegion name=selServiceRegion style="width: 200" onChange="fct_onChange();">
	<%
		while not rsSRrp.EOF
			Response.Write "<OPTION"
			if strServiceRegion <> "" then if CLng(objRS("CUST_MGD_SRVC_RGN_LCODE")) = CLng(rsSRrp("CUST_MGD_SRVC_RGN_LCODE")) then Response.write " selected"
				Response.Write " VALUE="& rsSRrp("CUST_MGD_SRVC_RGN_LCODE") &">" & routineHtmlString(rsSRrp("CUST_MGD_SRVC_RGN_NAME")) & "</OPTION>" &vbCrLf
			rsSRrp.MoveNext
		wend
		rsSRrp.Close
		%>
	</SELECT>
	</td>
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
	Record Status Indicator:<INPUT align="left" name="txtRecordStatusInd" type="text" style="width: 18px" disabled value="<%If Len(strCityName) <> 0 Then Response.Write objRS.Fields("RECORD_STATUS_IND").Value%>">&nbsp;&nbsp;&nbsp;
	Create Date:<INPUT align="center" name="txtRecordStatusInd" type="text" style="width: 150px" disabled value="<%If Len(strCityName) <> 0 Then Response.Write objRS.Fields("CREATE_DATE_TIME").Value%>">&nbsp;
	Created By:<INPUT align="right" name="txtRecordStatusInd" type="text" style="width: 150px" disabled value="<%If Len(strCityName) <> 0 Then Response.Write objRS.Fields("CREATE_REAL_USERID").Value%>"><BR>
	Update Date:<INPUT align="center" name="txtRecordStatusInd" type="text" style="width: 150px" disabled value="<%If Len(strCityName) <> 0 Then Response.Write objRS.Fields("UPDATE_DATE_TIME").Value%>">&nbsp;
	Updated By:<INPUT align="right" name="txtRecordStatusInd" type="text" style="width: 150px" disabled value="<%If Len(strCityName) <> 0 Then Response.Write objRS.Fields("UPDATE_REAL_USERID").Value%>">
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

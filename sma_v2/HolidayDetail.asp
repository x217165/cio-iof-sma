<%@ Language=VBScript %>
<%Option Explicit%>
<%Response.Buffer = True%>
<!--#include file = "smaConstants.inc"-->
<!--#include file = "smaProcs.inc"-->
<!--#include file = "databaseconnect.asp" -->
<!--
*************************************************************************************************
* Page name:	SLADetail.asp																	*
* Purpose:		To display the Service Level Agreement											*
*				Chosen via SLAList.asp															*
*																								*
* Created by:	Gilles Archer 10/02/2000														*
*																								*
*************************************************************************************************
-->
<%
Dim strHolidayID, datUpdateDateTime, strWinMessage, strWinLocation
Dim	objRS, objProvinces, objCountries, objCommand, strSQL, strWhere, strOrderBy, strErrMessage, lIndex
Dim p_userid, p_holiday_id, p_holiday_desc, p_last_update_dt
Dim intAccessLevel

	'Check user's rights
	intAccessLevel = CInt(CheckLogon(strConst_ServiceCatalogue))
	If (intAccessLevel And intConst_Access_ReadOnly) <> intConst_Access_ReadOnly  Then
		DisplayError "BACK", "", 0, "ACCESS DENIED", "You don't have access to Holiday.  Please contact your system administrator"
	End If

	strWinMessage = ""
	strHolidayID = Request("HolidayID")

	p_userid = Session("username")

	If Len(Request.Form("hdnHolidayID")) <> 0 Then
		p_holiday_id = CLng(Request.Form("hdnHolidayID"))
	Else
		p_holiday_id = Null
	End If

	If Len(Request.Form("txtHolidayDescription")) <> 0 Then
		p_holiday_desc = Replace(Trim(Request.Form("txtHolidayDescription")), "'", "''")
	Else
		p_holiday_desc = Null
	End If

	If IsDate(Request.Form("hdnUpdateDateTime")) Then
		p_last_update_dt = CDate(Request.Form("hdnUpdateDateTime"))
	Else
		p_last_update_dt = Null
	End If

	Select Case UCase(Request("hdnFrmAction"))
		Case "SAVE"
			Set objCommand = Server.CreateObject("ADODB.Command")
			Set objCommand.ActiveConnection = objConn
			objCommand.CommandType = adCmdStoredProc

			If IsNumeric(Request("hdnHolidayID")) Then	'Save existing Service Type
				If (intAccessLevel And intConst_Access_Update) <> intConst_Access_Update Then
					DisplayError "REFRESH", strWinLocation, 0, "UPDATE DENIED", "You don't have access to update Holiday.  Please contact your system administrator"
				End If

				objCommand.CommandText = "sma_sp_userid.spk_sma_admin_inter.sp_holiday_update"
				objCommand.Parameters.Append objCommand.CreateParameter("p_userid", adVarChar, adParamInput, 20, p_userid)
				objCommand.Parameters.Append objCommand.CreateParameter("p_holiday_id", adNumeric, adParamInput, , p_holiday_id)
				objCommand.Parameters.Append objCommand.CreateParameter("p_holiday_desc", adVarChar, adParamInput, 80, p_holiday_desc)
				objCommand.Parameters.Append objCommand.CreateParameter("p_last_update_dt", adDBTimeStamp, adParamInput, , p_last_update_dt)		'Date

				strErrMessage = "CANNOT UPDATE OBJECT"
			Else										'Create a new Service Type
				If (intAccessLevel And intConst_Access_Create) <> intConst_Access_Create Then
					DisplayError "REFRESH", strWinLocation, 0, "INSERT DENIED", "You don't have access to create Holiday.  Please contact your system administrator"
				End If

				objCommand.CommandText = "sma_sp_userid.spk_sma_admin_inter.sp_holiday_insert"
				objCommand.Parameters.Append objCommand.CreateParameter("p_userid", adVarChar, adParamInput, 20, p_userid)
				objCommand.Parameters.Append objCommand.CreateParameter("p_holiday_id", adNumeric, adParamOutput, , p_holiday_id)
				objCommand.Parameters.Append objCommand.CreateParameter("p_holiday_desc", adVarChar, adParamInput, 80, p_holiday_desc)

				strErrMessage = "CANNOT CREATE OBJECT"
			End If

			On Error Resume Next
			objCommand.Execute
			If objConn.Errors.Count <> 0 then
				DisplayError "BACK", "", objConn.Errors(0).NativeError, strErrMessage, objConn.Errors(0).Description
				objConn.Errors.Clear
			End If
			strHolidayID = CStr(objCommand.Parameters("p_holiday_id").Value)
			Set objCommand = Nothing
			strWinMessage = "Record saved successfully. You can now see the changes you made."

		Case "DELETE"
			If (intAccessLevel And intConst_Access_Delete) <> intConst_Access_Delete Then
				DisplayError "REFRESH", strWinLocation, 0, "DELETE DENIED", "You don't have access to delete Holiday.  Please contact your system administrator"
			End If

			Set objCommand = Server.CreateObject("ADODB.Command")
			Set objCommand.ActiveConnection = objConn
			objCommand.CommandType = adCmdStoredProc
			objCommand.CommandText = "sma_sp_userid.spk_sma_admin_inter.sp_holiday_delete"
			objCommand.Parameters.Append objCommand.CreateParameter("p_holiday_id", adNumeric, adParamInput, , CLng(strHolidayID))					'number(9)
			objCommand.Parameters.Append objCommand.CreateParameter("p_last_update_dt", adDBTimeStamp, adParamInput, ,CDate(datUpdateDateTime))		'Date

  			On Error Resume Next
			objCommand.Execute
			If objConn.Errors.Count <> 0 Then
				DisplayError "BACK", "", objConn.Errors(0).NativeError, "CANNOT DELETE OBJECT", objConn.Errors(0).Description
				objConn.Errors.Clear
			End If
			Set objCommand = Nothing
			strHolidayID = "DEL"
			strWinMessage = "Record deleted successfully."
	End Select

	If IsNumeric(strHolidayID) Then
		strSQL = "SELECT HOLIDAY_ID, " &_
				"HOLIDAY_NAME, " &_
				"HOLIDAY_MONTH, " &_
				"HOLIDAY_DAY, " &_
				"TO_CHAR(HOLIDAY_DATE, 'MON-DD-YYYY') AS HOLIDAY_DATE, " &_
				"PROVINCE_STATE_LCODE, " &_
				"COUNTRY_LCODE, " &_
				"TO_CHAR(CREATE_DATE_TIME, 'MON-DD-YYYY HH24:MI:SS') AS CREATE_DATE_TIME, " &_
				"sma_sp_userid.spk_sma_library.sf_get_full_username(CREATE_REAL_USERID) AS CREATE_REAL_USERID, " &_
				"TO_CHAR(UPDATE_DATE_TIME, 'MON-DD-YYYY HH24:MI:SS') AS UPDATE_DATE_TIME, " &_
				"sma_sp_userid.spk_sma_library.sf_get_full_username(UPDATE_REAL_USERID) AS UPDATE_REAL_USERID, " &_
				"UPDATE_DATE_TIME AS LAST_UPDATE_DATE_TIME, " &_
				"RECORD_STATUS_IND " &_
				"FROM CRP.HOLIDAY " &_
				"WHERE HOLIDAY_ID = " & strHolidayID

		'Create Recordset object
		Set objRS = Server.CreateObject("ADODB.Recordset")
		On Error Resume Next
		objRS.Open strSQL, objConn, adOpenForwardOnly, adLockReadOnly, adCmdText
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
	Set objProvinces = Server.CreateObject("ADODB.Recordset")
	On Error Resume Next
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
<LINK rel="stylesheet" type="text/css" href="stylesheets/styles.css">
<SCRIPT type="text/javascript" language="javascript" src="AccessLevels.js"></SCRIPT>
<SCRIPT type="text/javascript" language="javascript" src="GeneralJavaFunctions.js"></SCRIPT>
<SCRIPT type="text/javascript" language="javascript" id="clientEventHandlersJS">
<!-- //Hide Client-Side SCRIPT
var strWinMessage = "<%=strWinMessage%>";
var intAccessLevel = <%=intAccessLevel%>;
var bolSaveRequired = false;

setPageTitle("SMA - Holiday Definition");

function fct_setDays(iIndex) {
var intDays = 31;
var strMonth = document.frmHolidayDetail.item("selmonth", iIndex).options[document.frmHolidayDetail.item("selmonth", iIndex).selectedIndex].value;
var strYear = document.frmHolidayDetail.item("selyear", iIndex).options[document.frmHolidayDetail.item("selyear", iIndex).selectedIndex].value;
var intCurrentDay = document.frmHolidayDetail.item("selday", iIndex).options[document.frmHolidayDetail.item("selday", iIndex).selectedIndex].value;
var intCounter = document.frmHolidayDetail.item("selday", iIndex).options.length;

	switch (strMonth) {
		case "02":
			if (strYear % 4 != 0) { intDays = 28; }
			else if (strYear % 400 == 0) { intDays = 29; }
			else if (strYear % 100 == 0) { intDays = 28; }
			else { intDays = 29; }
			break;
		case "04": intDays = 30; break;
		case "06": intDays = 30; break;
		case "09": intDays = 30; break;
		case "11": intDays = 30; break;
		default: intDays=31; break;
	}
	if (intCounter <= intDays) {
		while (intCounter <= intDays) {
			var oOption = new Option(intCounter, intCounter);
			document.frmHolidayDetail.item("selday", iIndex).options[intCounter++] = oOption;
		}
	}
	else {
		while (intCounter > intDays) {
			document.frmHolidayDetail.item("selday", iIndex).options[intCounter--] = null;
		}
	}
	if (intCurrentDay > intDays) {
		document.frmHolidayDetail.item("selday", iIndex).selectedIndex = intDays;
	}
	bolSaveRequired = true;
}

function btnCalendar_onClick(iIndex) {
	var NewWin;
	SetCookie("Field", iIndex);
	NewWin=window.open("TheCalendar.asp","NewWin","toolbar=no,status=no,width=260,height=225,menubar=no resize=no");
	//NewWin.creator=self;
	NewWin.focus();
	bolSaveRequired = true;
}

function selNavigate_onChange(){
//***********************************************************************************************
// Function:	selNavigate_onChange															*
//																								*
// Purpose:		To display the page selected by the user from Quick Navigation drop-down box.	*
//																								*
// Created By:	Gilles Archer 09/27/2000														*
//																								*
// Updated By:																					*
//***********************************************************************************************
var strPageName = document.frmHolidayDetail.selNavigate.item(document.frmHolidayDetail.selNavigate.selectedIndex).value ;

	switch (strPageName) {

		case "DEFAULT":
			// do nothing ;
	}
}

function btnDelete_onClick() {
//**********************************************************************************************
// Function:	btnDelete_onClick
//
// Purpose:		To delete a holiday
//
// Created By:	Gilles Archer 09/27/2000
//
// Updated By:
//***********************************************************************************************
	if ((intAccessLevel & intConst_Access_Delete) != intConst_Access_Delete) {
		alert('You do not have permission to DELETE a Holiday.  Please contact your System Administrator.');
		return false;
	}

	if (document.frmHolidayDetail.hdnHolidayID.value == "") {
		alert('This Holiday does not exist in the database.');
		return false;
	}

	if (confirm('Do you really want to delete this object?')){
		document.frmHolidayDetail.hdnFrmAction.value = "DELETE";
		document.frmHolidayDetail.submit();
	}
}

function btnNew_onClick() {
	if ((intAccessLevel & intConst_Access_Create) != intConst_Access_Create) {
		alert('You do not have permission to CREATE a Holiday.  Please contact your System Administrator.');
		return false;
	}
	document.location = "HolidayDetail.asp?HolidayID=NEW";
}

function fct_onChange() {
	bolSaveRequired = true;
}

function btnSave_onClick() {
	if ((intAccessLevel & intConst_Access_Update) != intConst_Access_Update) {
		alert('You do not have permission to UPDATE a Holiday.  Please contact your System Administrator.');
		return false;
	}

	if (document.frmHolidayDetail.txtHolidayDescription.value == "") {
		alert('Please enter the Holiday Description');
		document.frmHolidayDetail.txtHolidayDescription.focus();
		return false;
	}

	document.frmHolidayDetail.hdnFrmAction.value = "SAVE";
	bolSaveRequired = false;
	document.frmHolidayDetail.submit();
	return true;
}

function btnReferences_onClick() {
var strOwner = 'CRP';			// owner name must be in Uppercase
var strTableName = 'HOLIDAY';		// replace ADDRESS with your own table name and table name must be in Uppercase
var strRecordID = document.frmHolidayDetail.hdnHolidayID.value ;   // insert your record id
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
	document.frmHolidayDetail.btnSave.focus();

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
	document.location.href = "HolidayDetail.asp?HolidayID=<%=strHolidayID%>";
}
// Unhide Client-Side SCRIPT -->
</SCRIPT>
</HEAD>
<!--<BODY language="javascript" onLoad="DisplayStatus(strWinMessage);" onBeforeUnload="window_onBeforeUnload();" onUnload="window_onUnload();">-->
<BODY>
<FORM id="frmHolidayDetail" name="frmHolidayDetail" action="HolidayDetail.asp" method="post">
	<INPUT type="hidden" id="hdnHolidayID" name="hdnHolidayID" value="<%If IsNumeric(strHolidayID) Then Response.Write objRS.Fields("HOLIDAY_ID").Value%>">
	<INPUT type="hidden" id="hdnUpdateDateTime" name="hdnUpdateDateTime" value="<%If IsNumeric(strHolidayID) Then Response.Write objRS.Fields("LAST_UPDATE_DATE_TIME").Value%>">
	<INPUT type="hidden" id="hdnFrmAction" name="hdnFrmAction" value="">
<TABLE border="0" cols="4" width="100%">
<THEAD>
<TR>
	<TD align="left" colspan="3">Holiday Detail</TD>
	<TD align="right"><SELECT valign="top" id="selNavigate" name="selNavigate" onChange="fct_selNavigate();">
		<OPTION value="DEFAULT" selected>Quickly Goto ...</OPTION>
</TR>
</THEAD>
<TBODY>
<TR>
	<TD align="right">Holiday<FONT color="red">*</FONT></TD>
	<TD align="left" colspan="3" nowrap><INPUT type="text" id="txtHolidayName" name="txtHolidayName" disabled onChange="fct_onChange();" maxlength="30" size="30" value="<%If IsNumeric(strHolidayID) Then Response.Write objRS.Fields("HOLIDAY_NAME").Value%>"></TD>
</TR>
<TR>
	<TD align="right">Yearly Date</TD>
	<TD align="left">
	    <SELECT id="selmonth" name="selmonth" disabled onChange="fct_setDays(0);return fct_onChange();">
		<OPTION></OPTION>
		<%For lIndex = 1 to 12
			Response.Write "<OPTION "
			If IsNumeric(strHolidayID) Then
				If Not IsNull(objRS.Fields("HOLIDAY_MONTH").Value) Then
					If lIndex = CLng(objRS.Fields("HOLIDAY_MONTH").Value) Then Response.Write "selected "
				End If
			End If
			If lIndex < 10 Then	lIndex = "0" & lIndex
			Response.Write "value='" & lIndex & "'>" & monthName(lIndex, False) & "</OPTION>"
		Next%>
		</SELECT>
		<SELECT id="selday" name="selday" disabled onChange="fct_setDays(0);return fct_onChange();">
		<OPTION></OPTION>
		<%For lIndex = 1 to 31
			Response.Write "<OPTION "
			If IsNumeric(strHolidayID) Then
				If Not IsNull(objRS.Fields("HOLIDAY_DAY").Value) Then
					If lIndex = CLng(objRS.Fields("HOLIDAY_DAY").Value) Then Response.Write "selected "
				End If
			End If
			If lIndex < 10 Then	lIndex = "0" & lIndex
			Response.Write "value='" & lIndex & "'>" & lIndex & "</OPTION>"
		Next%>
		</SELECT>
		<SELECT id="selyear" name="selyear" disabled onChange="fct_setDays(0);return fct_onChange();">
		<OPTION></OPTION>
		<%For lIndex = intBaseYear To Year(Now) + 7
			Response.Write "<OPTION value='" & lIndex & "'>" & lIndex & "</OPTION>"
		Next%>
		</SELECT>
		<INPUT id="btnCalendar" name="btnCalendar" type="button" disabled value="..." language="javascript" onClick="return btnCalendar_onClick(0);return fct_onChange();"></TD>
	<TD align="right">Province / State</TD>
	<TD align="left">
	<SELECT id="selProvince" name="selProvince" style="width: 200" disabled onChange="fct_onChange();">
		<OPTION></OPTION>
		<%Do While Not objProvinces.EOF
			If Len(strCityName) <> 0 Then
				If StrComp(objProvinces.Fields("PROVINCE_STATE_LCODE").Value, objRS.Fields("PROVINCE_STATE_LCODE").Value, 0) = 0 Then
					Response.Write "<OPTION selected value='" & objProvinces.Fields("PROVINCE_STATE_LCODE").Value & "'>" & objProvinces.Fields("PROVINCE_STATE_LCODE").Value & " - " & objProvinces.Fields("PROVINCE_STATE_NAME").Value & "</OPTION>"
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
	<TD align="right">Year specific Date</TD>
	<TD align="left">
	    <SELECT id="selmonth" name="selmonth" disabled onChange="fct_setDays(1);return fct_onChange();">
		<OPTION></OPTION>
		<%For lIndex = 1 to 12
			Response.Write "<OPTION "
			If IsNumeric(strHolidayID) Then
				If lIndex = Month(objRS.Fields("HOLIDAY_DATE").Value) Then Response.Write "selected "
			End If
			If lIndex < 10 Then	lIndex = "0" & lIndex
			Response.Write "value='" & lIndex & "'>" & monthName(lIndex, False) & "</OPTION>"
		Next%>
		</SELECT>
		<SELECT id="selday" name="selday" disabled onChange="fct_setDays(1);return fct_onChange();">
		<OPTION></OPTION>
		<%For lIndex = 1 to 31
			Response.Write "<OPTION "
			If IsNumeric(strHolidayID) Then
				If lIndex = Day(objRS.Fields("HOLIDAY_DATE").Value) Then Response.Write "selected "
			End If
			If lIndex < 10 Then	lIndex = "0" & lIndex
			Response.Write "value='" & lIndex & "'>" & lIndex & "</OPTION>"
		Next%>
		</SELECT>
		<SELECT id="selyear" name="selyear" disabled onChange="fct_setDays(1);return fct_onChange();">
		<OPTION></OPTION>
		<%For lIndex = intBaseYear To Year(Now) + 7
			Response.Write "<OPTION "
			If IsNumeric(strHolidayID) Then
				If lIndex = Year(objRS.Fields("HOLIDAY_DATE").Value) Then Response.Write "selected "
			End If
			Response.Write "value='" & lIndex & "'>" & lIndex & "</OPTION>"
		Next%>
		</SELECT>
		<INPUT id="btnCalendar" name="btnCalendar" type="button" disabled value="..." language="javascript" onClick="return btnCalendar_onClick(1);return fct_onChange();"></TD>
	<TD align="right">Country</TD>
	<TD align="left">
	<SELECT id="selCountry" name="selCountry" style="width: 200" disabled onChange="fct_onChange();">
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
</TBODY>
<TFOOT>
<TR>
	<TD colspan="4" align="right">
	<INPUT id="btnReferences" name="btnReferences" type="button" value="References" style="width: 2.2cm" language="javascript" onClick="return btnReferences_onClick();">&nbsp;
	<INPUT disabled id="btnDelete" name="btnDelete" type="button" value="Delete" style="width: 2cm" language="javascript" onClick="return btnDelete_onClick();">&nbsp;
	<INPUT disabled id="btnReset" name="btnReset" type="button" value="Reset" style="width: 2cm" language="javascript" onClick="return btnReset_onClick();">&nbsp;
	<INPUT disabled id="btnNew" name="btnNew" type="button" value="New" style="width: 2cm" language="javascript" onClick="return btnNew_onClick();">&nbsp;
	<INPUT disabled id="btnSave" name="btnSave" type="button" value="Save" style="width: 2cm" onClick="return btnSave_onClick();">&nbsp;</TD>
</TR>
</TFOOT>
</TABLE>
<FIELDSET width="100%">
	<LEGEND align="right"><b>Audit Information</b></LEGEND>
	<DIV size="8pt" align="right">
	Record Status Indicator:<INPUT align="left" name="txtRecordStatusInd" type="text" style="width: 18px" disabled value="<%If IsNumeric(strHolidayID) Then Response.Write objRS.Fields("RECORD_STATUS_IND").Value%>">&nbsp;&nbsp;&nbsp;
	Create Date:<INPUT align="center" name="txtRecordStatusInd" type="text" style="width: 150px" disabled value="<%If IsNumeric(strHolidayID) Then Response.Write objRS.Fields("CREATE_DATE_TIME").Value%>">&nbsp;
	Created By:<INPUT align="right" name="txtRecordStatusInd" type="text" style="width: 150px" disabled value="<%If IsNumeric(strHolidayID) Then Response.Write objRS.Fields("CREATE_REAL_USERID").Value%>"><br>
	Update Date:<INPUT align="center" name="txtRecordStatusInd" type="text" style="width: 150px" disabled value="<%If IsNumeric(strHolidayID) Then Response.Write objRS.Fields("UPDATE_DATE_TIME").Value%>">&nbsp;
	Updated By:<INPUT align="right" name="txtRecordStatusInd" type="text" style="width: 150px" disabled value="<%If IsNumeric(strHolidayID) Then Response.Write objRS.Fields("UPDATE_REAL_USERID").Value%>">
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

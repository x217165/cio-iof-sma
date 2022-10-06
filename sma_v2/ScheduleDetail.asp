<%@ Language=VBScript %>
<%Option Explicit%>
<%Response.Buffer = True%>
<!--#include file = "smaConstants.inc"-->
<!--#include file = "smaProcs.inc"-->
<!--#include file = "databaseconnect.asp" -->
<!--
*************************************************************************************************
* Page name:	SLADetail.asp																	*
* Purpose:		To display the Schedule Definition												*
*				Chosen via ScheduleList.asp														*
*																								*
* Created by:	Gilles Archer 10/02/2000														*
*																								*
*************************************************************************************************
-->
<%
Dim strScheduleID, datUpdateDateTime, strWinMessage, strWinLocation, strErrMessage, lIndex
Dim	objCommand, objMaster, objDetail, objProvinces, objCountries, strSQL, strWhere
Dim p_userid, p_schedule_id, p_schedule_name, p_include_holiday_flag, p_last_update_dt
Dim p_detail_schedule_detail_id, p_detail_day_of_weeks, p_detail_start_hours, p_detail_end_hours
Dim p_province_state_lcode, p_country_lcode
Dim intAccessLevel

	'Check user's rights
	intAccessLevel = CInt(CheckLogon(strConst_ServiceCatalogue))
	If (intAccessLevel And intConst_Access_ReadOnly) <> intConst_Access_ReadOnly  Then
		DisplayError "BACK", "", 0, "ACCESS DENIED", "You don't have access to Schedule.  Please contact your system administrator"
	End If

	strWinMessage = ""
	strScheduleID = Request("ScheduleID")

	'Retrieve Form Information
	p_userid = Session("username")

	If IsNumeric(Request.Form("hdnScheduleID")) Then
		p_schedule_id = CLng(Request.Form("hdnScheduleID"))
	Else
		p_schedule_id = Null
	End If

	If Len(Request.Form("txtScheduleName")) <> 0 Then
		p_schedule_name = Trim(Request.Form("txtScheduleName"))
	Else
		p_schedule_name = Null
	End If

	If Len(Request.Form("chkHoliday")) <> 0 Then
		p_include_holiday_flag = "Y"
	Else
		p_include_holiday_flag = "N"
	End If

	If IsDate(Request.Form("hdnUpdateDateTime")) Then
		p_last_update_dt = CDate(Request.Form("hdnUpdateDateTime"))
	Else
		p_last_update_dt = Null
	End If

	For lIndex = 1 To Request.Form("hdnScheduleDetailID").Count
		p_detail_schedule_detail_id = p_detail_schedule_detail_id & Request.Form("hdnScheduleDetailID").Item(lIndex) & "~"
	Next

	For lIndex = 1 To Request.Form("txtWeekDay").Count
		p_detail_day_of_weeks = p_detail_day_of_weeks & UCase(Left(Request.Form("txtWeekDay").Item(lIndex), 3)) & "~"
	Next

	For lIndex = 1 To Request.Form("txtStartHour").Count
		p_detail_start_hours = p_detail_start_hours & "0" & Request.Form("txtStartHour").Item(lIndex) & "~"
	Next

	For lIndex = 1 To Request.Form("txtEndHour").Count
		p_detail_end_hours = p_detail_end_hours & "0" & Request.Form("txtEndHour").Item(lIndex) & "~"
	Next

	If Len(Request.Form("selProvince")) <> 0 Then
		p_province_state_lcode = UCase(Trim(Request.Form("selProvince")))
	Else
		p_province_state_lcode = Null
	End If

	If Len(Request.Form("selCountry")) <> 0 Then
		p_country_lcode = UCase(Trim(Request.Form("selCountry")))
	Else
		p_country_lcode = Null
	End If
	'Response.Write "<BR>p_userid: " & Request.Form("txtWeekDay").Count
	'Response.Write "<BR>p_userid: " & p_userid
	'Response.Write "<BR>p_schedule_id: " & p_schedule_id
	'Response.Write "<BR>p_schedule_name: " & p_schedule_name
	'Response.Write "<BR>p_include_holiday_flag: " & p_include_holiday_flag
	'Response.Write "<BR>p_last_update_dt: " & p_last_update_dt
	'Response.Write "<BR>p_detail_schedule_detail_id: " & p_detail_schedule_detail_id
	'Response.Write "<BR>p_detail_day_of_weeks: " & p_detail_day_of_weeks
	'Response.Write "<BR>p_detail_start_hours: " & p_detail_start_hours
	'Response.Write "<BR>p_detail_end_hours: " & p_detail_end_hours
	'Response.Write "<BR>p_province_state_lcode: " & p_province_state_lcode
	'Response.Write "<BR>p_country_lcode: " & p_country_lcode
	'Response.End

	Select Case UCase(Request("hdnFrmAction"))
		Case "SAVE"
			Set objCommand = Server.CreateObject("ADODB.Command")
			Set objCommand.ActiveConnection = objConn
			objCommand.CommandType = adCmdStoredProc

			If IsNumeric(Request("hdnScheduleID")) Then	'Save existing Service Type
				If (intAccessLevel And intConst_Access_Update) <> intConst_Access_Update Then
					DisplayError "REFRESH", strWinLocation, 0, "UPDATE DENIED", "You don't have access to update Schedule.  Please contact your system administrator"
				End If

				objCommand.CommandText = "sma_sp_userid.spk_sma_admin_inter.sp_schedule_update"
				objCommand.Parameters.Append objCommand.CreateParameter("p_userid", adVarChar, adParamInput, 20, p_userid)
				objCommand.Parameters.Append objCommand.CreateParameter("p_schedule_id", adNumeric, adParamInput, , p_schedule_id)
				objCommand.Parameters.Append objCommand.CreateParameter("p_schedule_name", adVarChar, adParamInput, 30, p_schedule_name)
				objCommand.Parameters.Append objCommand.CreateParameter("p_include_holiday_flag", adChar, adParamInput, 1, p_include_holiday_flag)
				objCommand.Parameters.Append objCommand.CreateParameter("p_last_update_dt", adDBTimeStamp, adParamInput, , p_last_update_dt)
				objCommand.Parameters.Append objCommand.CreateParameter("p_detail_schedule_detail_id", adVarChar, adParamInput, 100, p_detail_schedule_detail_id)
				objCommand.Parameters.Append objCommand.CreateParameter("p_detail_day_of_weeks", adVarChar, adParamInput, 40, p_detail_day_of_weeks)
				objCommand.Parameters.Append objCommand.CreateParameter("p_detail_start_hours", adVarChar, adParamInput, 100, p_detail_start_hours)
				objCommand.Parameters.Append objCommand.CreateParameter("p_detail_end_hours", adVarChar, adParamInput, 100, p_detail_end_hours)
				objCommand.Parameters.Append objCommand.CreateParameter("p_province_state_lcode", adChar, adParamInput, 2, p_province_state_lcode)
				objCommand.Parameters.Append objCommand.CreateParameter("p_country_lcode", adChar, adParamInput, 2, p_country_lcode)

				strErrMessage = "CANNOT UPDATE OBJECT"
			Else										'Create a new Service Type
				If (intAccessLevel And intConst_Access_Create) <> intConst_Access_Create Then
					DisplayError "REFRESH", strWinLocation, 0, "INSERT DENIED", "You don't have access to create Schedule.  Please contact your system administrator"
				End If

				objCommand.CommandText = "sma_sp_userid.spk_sma_admin_inter.sp_schedule_insert"
				objCommand.Parameters.Append objCommand.CreateParameter("p_userid", adVarChar, adParamInput, 20, p_userid)
				objCommand.Parameters.Append objCommand.CreateParameter("p_schedule_id", adNumeric, adParamOutput, , Null)
				objCommand.Parameters.Append objCommand.CreateParameter("p_schedule_name", adVarChar, adParamInput, 30, p_schedule_name)
				objCommand.Parameters.Append objCommand.CreateParameter("p_include_holiday_flag", adChar, adParamInput, 1, p_include_holiday_flag)
				objCommand.Parameters.Append objCommand.CreateParameter("p_detail_schedule_detail_id", adVarChar, adParamOutput, 100, Null)
				objCommand.Parameters.Append objCommand.CreateParameter("p_detail_day_of_weeks", adVarChar, adParamInput, 40, p_detail_day_of_weeks)
				objCommand.Parameters.Append objCommand.CreateParameter("p_detail_start_hours", adVarChar, adParamInput, 100, p_detail_start_hours)
				objCommand.Parameters.Append objCommand.CreateParameter("p_detail_end_hours", adVarChar, adParamInput, 100, p_detail_end_hours)
				objCommand.Parameters.Append objCommand.CreateParameter("p_province_state_lcode", adChar, adParamInput, 2, p_province_state_lcode)
				objCommand.Parameters.Append objCommand.CreateParameter("p_country_lcode", adChar, adParamInput, 2, p_country_lcode)

				strErrMessage = "CANNOT CREATE OBJECT"
			End If

			On Error Resume Next
			objCommand.Execute
			If objConn.Errors.Count <> 0 then
				DisplayError "BACK", "", objConn.Errors(0).NativeError, strErrMessage, objConn.Errors(0).Description
				objConn.Errors.Clear
			End If
			strScheduleID = CStr(objCommand.Parameters("p_schedule_id").Value)
			strWinMessage = "Record saved successfully. You can now see the changes you made."

		Case "DELETE"
			If (intAccessLevel And intConst_Access_Delete) <> intConst_Access_Delete Then
				DisplayError "REFRESH", strWinLocation, 0, "DELETE DENIED", "You don't have access to delete Schedule.  Please contact your system administrator"
			End If

			Set objCommand = Server.CreateObject("ADODB.Command")
			Set objCommand.ActiveConnection = objConn
			objCommand.CommandType = adCmdStoredProc
			objCommand.CommandText = "sma_sp_userid.spk_sma_admin_inter.sp_schedule_delete"
			objCommand.Parameters.Append objCommand.CreateParameter("p_schedule_id", adNumeric, adParamInput, , p_schedule_id)					'number(9)
			objCommand.Parameters.Append objCommand.CreateParameter("p_last_update_dt", adDBTimeStamp, adParamInput, , p_last_update_dt)		'Date

  			On Error Resume Next
			objCommand.Execute
			If objConn.Errors.Count <> 0 Then
				DisplayError "BACK", "", objConn.Errors(0).NativeError, "CANNOT DELETE OBJECT", objConn.Errors(0).Description
				objConn.Errors.Clear
			End If
			strScheduleID = "DEL"
			strWinMessage = "Record deleted successfully."
	End Select

	If IsNumeric(strScheduleID) Then
		strSQL = "SELECT SCHEDULE_ID, SCHEDULE_NAME, " &_
				"INCLUDE_HOLIDAY_FLAG, PROVINCE_STATE_LCODE, COUNTRY_LCODE, " &_
				"TO_CHAR(CREATE_DATE_TIME, 'MON-DD-YYYY HH24:MI:SS') AS CREATE_DATE_TIME, " &_
				"sma_sp_userid.spk_sma_library.sf_get_full_username(CREATE_REAL_USERID) AS CREATE_REAL_USERID, " &_
				"TO_CHAR(UPDATE_DATE_TIME, 'MON-DD-YYYY HH24:MI:SS') AS UPDATE_DATE_TIME, " &_
				"sma_sp_userid.spk_sma_library.sf_get_full_username(UPDATE_REAL_USERID) AS UPDATE_REAL_USERID, " &_
				"RECORD_STATUS_IND, " &_
				"UPDATE_DATE_TIME AS LAST_UPDATE_DATE_TIME " &_
				"FROM CRP.SCHEDULE " &_
				"WHERE SCHEDULE_ID = " & strScheduleID

		'Create Recordset object
		Set objMaster = Server.CreateObject("ADODB.Recordset")
		On Error Resume Next
		Set objMaster = objCommand.Execute
		objMaster.Open strSQL, objConn, adOpenForwardOnly, adLockReadOnly, adCmdText
		If objConn.Errors.Count <> 0 Then
			DisplayError "BACK", "", objConn.Errors(0).NativeError, "ERROR LOADING DATA (Master)", objConn.Errors(0).Description
			objConn.Errors.Clear
		End If

		strSQL = "SELECT SCHEDULE_DETAIL_ID, " &_
				"DECODE(DAY_OF_WEEK, 'SUN', 'Sunday', 'MON', 'Monday', 'TUE', 'Tuesday', 'WED', 'Wednesday', 'THU', 'Thursday', 'FRI', 'Friday', 'SAT', 'Saturday') AS DAY_NAME, " &_
				"START_HOUR, " &_
				"END_HOUR, " &_
				"DECODE(DAY_OF_WEEK, 'SUN', 1, 'MON', 2, 'TUE', 3, 'WED', 4, 'THU', 5, 'FRI', 6, 'SAT', 7) AS WEEKDAY " &_
				"FROM CRP.SCHEDULE_DETAIL " &_
				"WHERE SCHEDULE_ID = " & strScheduleID &_
				" ORDER BY WEEKDAY"

		'Create Recordset object
		Set objDetail = Server.CreateObject("ADODB.Recordset")
		On Error Resume Next
		Set objDetail = objCommand.Execute
		objDetail.Open strSQL, objConn, adOpenDynamic, adLockReadOnly, adCmdText
		If objConn.Errors.Count <> 0 Then
			DisplayError "BACK", "", objConn.Errors(0).NativeError, "ERROR LOADING DATA (Master)", objConn.Errors(0).Description
			objConn.Errors.Clear
		End If
	End If

	'Selection
	strSQL = "SELECT UPPER(PS.PROVINCE_STATE_LCODE) AS PROVINCE_STATE_LCODE, " &_
			"UPPER(PS.COUNTRY_LCODE) AS COUNTRY_LCODE, " &_
			"PS.PROVINCE_STATE_NAME " &_
			"FROM CRP.LCODE_PROVINCE_STATE PS " &_
			"WHERE PS.RECORD_STATUS_IND = 'A' " &_
			"ORDER BY PS.COUNTRY_LCODE ASC, PS.PROVINCE_STATE_LCODE ASC"

	'Create Recordset object
	Set objProvinces = Server.CreateObject("ADODB.Recordset")
	On Error Resume Next
	objProvinces.Open strSQL, objConn, adOpenForwardOnly, adLockReadOnly, adCmdText
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
	objCountries.Open strSQL, objConn, adOpenForwardOnly, adLockReadOnly, adCmdText
	If objConn.Errors.Count <> 0 Then
		DisplayError "BACK", "", objConn.Errors(0).NativeError, "ERROR LOADING DATA (Countries)", objConn.Errors(0).Description
		objConn.Errors.Clear
	End If

Dim arrWeekDay(6)
	arrWeekDay(0) = "Sunday"
	arrWeekDay(1) = "Monday"
	arrWeekDay(2) = "Tuesday"
	arrWeekDay(3) = "Wednesday"
	arrWeekDay(4) = "Thursday"
	arrWeekDay(5) = "Friday"
	arrWeekDay(6) = "Saturday"
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

setPageTitle("SMA - Schedule Definition");

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
var strPageName = document.frmScheduleDetail.selNavigate.item(document.frmScheduleDetail.selNavigate.selectedIndex).value ;

	switch (strPageName) {
		case "DEFAULT": break;
			// do nothing ;
	}
}

function btnDelete_onClick() {
//**********************************************************************************************
// Function:	btnDelete_onClick
//
// Purpose:		To delete a Schedule
//
// Created By:	Gilles Archer 09/27/2000
//
// Updated By:
//***********************************************************************************************
	if ((intAccessLevel & intConst_Access_Delete) != intConst_Access_Delete) {
		alert('You do not have permission to DELETE a Schedule.  Please contact your System Administrator.');
		return false;
	}

	if (document.frmScheduleDetail.hdnScheduleID.value == "") {
		alert('This Schedule does not exist in the database.');
		return false;
	}

	if (confirm('Do you really want to delete this object?')) {
		document.frmScheduleDetail.hdnFrmAction.value = "DELETE";
		document.frmScheduleDetail.submit();
	}
}

function btnNew_onClick() {
	if ((intAccessLevel & intConst_Access_Create) != intConst_Access_Create) {
		alert('You do not have permission to CREATE a Schedule.  Please contact your System Administrator.');
		return false;
	}
	document.location = "ScheduleDetail.asp?ScheduleID=NEW";
}

function btnReferences_onClick() {
var strOwner = 'CRP';			// owner name must be in Uppercase
var strTableName = 'SCHEDULE';		// replace ADDRESS with your own table name and table name must be in Uppercase
var strRecordID = document.frmScheduleDetail.hdnScheduleID.value ;   // insert your record id
var strURL;

	if (strRecordID == "") {
		alert("No references. This is a new record.");
		return false;
	}

	strURL = "Dependency.asp?Owner=" + strOwner + "&TableName=" + strTableName + "&RecordID=" + strRecordID;
	window.open(strURL, 'Popup', 'top=100, left=100, width=500, height=300');
}

function fct_onChange() {
	bolSaveRequired = true;
}

function btnSave_onClick() {
	if ((intAccessLevel & intConst_Access_Update) != intConst_Access_Update) {
		alert('You do not have permission to UPDATE a Schedule.  Please contact your System Administrator.');
		return false;
	}

	if (document.frmScheduleDetail.txtScheduleName.value == "") {
		alert('Please enter the Schedule Description');
		document.frmScheduleDetail.txtScheduleName.focus();
		return false;
	}

	document.frmScheduleDetail.hdnFrmAction.value = "SAVE";
	bolSaveRequired = false;
	document.frmScheduleDetail.submit();
	return true;
}

function window_onBeforeUnload() {
	//Ensure that fct_onChange() fires for any changed data.
	document.frmScheduleDetail.btnSave.focus();

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
	document.location.href = "ScheduleDetail.asp?ScheduleID=<%=strScheduleID%>";
}
// Unhide Client-Side SCRIPT -->
</SCRIPT>
</HEAD>
<BODY language="javascript" onLoad="DisplayStatus(strWinMessage);" onBeforeUnload="window_onBeforeUnload();" onUnload="window_onUnload();">
<FORM id="frmScheduleDetail" name="frmScheduleDetail" action="ScheduleDetail.asp" method="post">
	<INPUT type="hidden" id="hdnServiceLevelID" name="hdnScheduleID" value="<%If IsNumeric(strScheduleID) Then Response.Write objMaster.Fields("SCHEDULE_ID").Value%>">
	<INPUT type="hidden" id="hdnUpdateDateTime" name="hdnUpdateDateTime" value="<%If IsNumeric(strScheduleID) Then Response.Write objMaster.Fields("LAST_UPDATE_DATE_TIME").Value%>">
	<INPUT type="hidden" id="hdnFrmAction" name="hdnFrmAction" value="">
<TABLE border="0" cellpadding="1" cellspacing="0" rules=all width="100%">
<THEAD>
<TR>
	<TD align="left" colspan="2">Schedule Detail</TD>
	<TD align="right"><SELECT id="selNavigate" name="selNavigate" onChange="fct_selNavigate();">
		<OPTION value="DEFAULT" selected>Quickly Goto ...</OPTION>
	</TD>
</TR>
</THEAD>
<TBODY>
<TR valign="middle">
	<TD align="right">Schedule Name<FONT color="red">*</FONT></TD>
	<TD align="left"><INPUT id="txtScheduleName" name="txtScheduleName" type="text" maxlength="30" size="30" onChange="fct_onChange();" value="<%If IsNumeric(strScheduleID) Then Response.Write objMaster.Fields("SCHEDULE_NAME").Value%>"></TD>
	<TD rowspan="8">
		<TABLE border="1" cellpadding="0" cellspacing="0" width="100%">
		<THEAD>
		<TR align="center" valign="middle">
			<TD>Weekday</TD>
			<TD>Start Time</TD>
			<TD>End Time</TD>
		</TR>
		</THEAD>
		<%If IsNumeric(strScheduleID) Then
			Do Until objDetail.EOF%>
			<TR align="center" valign="middle">
				<TD><INPUT id="txtWeekDay" name="txtWeekDay" type="text" style="width=100%" readonly value="<%=objDetail.Fields("DAY_NAME").Value%>"><INPUT id="hdnScheduleDetailID" name="hdnScheduleDetailID" type="hidden" value="<%=objDetail.Fields("SCHEDULE_DETAIL_ID").Value%>"></TD>
				<TD><INPUT id="txtStartHour" name="txtStartHour" type="text" maxlength="4" size="4" onChange="fct_onChange();" value="<%=objDetail.Fields("START_HOUR").Value%>"></TD>
				<TD><INPUT id="txtEndHour" name="txtEndHour" type="text" maxlength="4" size="4" onChange="fct_onChange();" value="<%=objDetail.Fields("END_HOUR").Value%>"></TD>
			</TR>
			<%objDetail.MoveNext
			Loop
			objDetail.Close
			Set objDetail = Nothing
		Else
			For lIndex = LBound(arrWeekDay) To UBound(arrWeekDay)%>
			<TR align="center" valign="middle">
				<TD><INPUT id="txtWeekDay" name="txtWeekDay" type="text" style="width=100%" readonly value="<%=arrWeekDay(lIndex)%>"><INPUT id="hdnScheduleDetailID" name="hdnScheduleDetailID" type="hidden" value=""></TD>
				<TD><INPUT id="txtStartHour" name="txtStartHour" type="text" maxlength="4" size="4" onChange="fct_onChange();" value=""></TD>
				<TD><INPUT id="txtEndHour" name="txtEndHour" type="text" maxlength="4" size="4" onChange="fct_onChange();" value=""></TD>
			</TR>
			<%Next
		End If%>
		</TABLE>
	</TD>
</TR>
<TR valign="middle">
	<TD align="right">Holiday<FONT color="red">*</FONT></TD>
	<TD align="left"><INPUT id="chkHoliday" name="chkHoliday" type="checkbox" onChange="fct_onChange();" <%If IsNumeric(strScheduleID) Then If objMaster.Fields("INCLUDE_HOLIDAY_FLAG").Value = "Y" Then Response.Write "checked"%>></TD>
</TR>
<TR valign="middle">
	<TD align="right">Province / State</TD>
	<TD align="left">
	<SELECT id="selProvince" name="selProvince" style="width: 200" onChange="fct_onChange();">
		<OPTION></OPTION>
		<%Do While Not objProvinces.EOF
			If IsNumeric(strScheduleID) Then
				If StrComp(objProvinces.Fields("PROVINCE_STATE_LCODE").Value, objMaster.Fields("PROVINCE_STATE_LCODE").Value, 0) = 0 Then
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
<TR valign="middle">
	<TD align="right">Country</TD>
	<TD align="left">
	<SELECT id="selCountry" name="selCountry" style="width: 200" onChange="fct_onChange();">
		<OPTION></OPTION>
		<%Do While Not objCountries.EOF
			If IsNumeric(strScheduleID) Then
				If StrComp(objCountries.Fields("COUNTRY_LCODE").Value, objMaster.Fields("COUNTRY_LCODE").Value, 0) = 0 Then
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
<TR><TD>&nbsp;</TD><TD>&nbsp;</TD></TR>
<TR><TD>&nbsp;</TD><TD>&nbsp;</TD></TR>
<TR><TD>&nbsp;</TD><TD>&nbsp;</TD></TR>
<TR><TD>&nbsp;</TD><TD>&nbsp;</TD></TR>
</TBODY>
<TFOOT>
<TR>
	<TD colspan="4" align="right">
	<INPUT id="btnReferences" name="btnReferences" type="button" value="References" style="width: 2.2cm" language="javascript" onClick="return btnReferences_onClick();">&nbsp;
	<INPUT id="btnDelete" name="btnDelete" type="button" value="Delete" style="width: 2cm" language="javascript" onClick="return btnDelete_onClick();">&nbsp;
	<INPUT id="btnReset" name="btnReset" type="button" value="Reset" style="width: 2cm" language="javascript" onClick="return btnReset_onClick();">&nbsp;
	<INPUT id="btnNew" name="btnNew" type="button" value="New" style="width: 2cm" language="javascript" onClick="return btnNew_onClick();">&nbsp;
	<INPUT id="btnSave" name="btnSave" type="button" value="Save" style="width: 2cm" onClick="return btnSave_onClick();">&nbsp;</TD>
</TR>
</TFOOT>
</TABLE>
<!--<BR>
<TABLE align=right border=1 cellPadding=0 cellSpacing=0 width=100% rules=groups>
<THEAD>
<TR align=right bgcolor=white nowrap>
	<TD colspan=6><FONT color=black>Audit Information</FONT>&nbsp;</TD>
</TR>
</THEAD>
<TR align=right bgcolor=white nowrap>
	<TD>Record Status Indicator:</TD>
	<TD style="width: 18px"><INPUT name="txtRecordStatusInd" type="text" style="width: 18px" disabled value="<%If IsNumeric(strScheduleID) Then Response.Write objMaster.Fields("RECORD_STATUS_IND").Value%>"></TD>
	<TD style="width: 80px">Create Date:</TD>
	<TD style="width: 150px"><INPUT name="txtRecordStatusInd" type="text" style="width: 150px" disabled value="<%If IsNumeric(strScheduleID) Then Response.Write objMaster.Fields("CREATE_DATE_TIME").Value%>"></TD>
	<TD style="width: 80px">Created By:</TD>
	<TD style="width: 150px"><INPUT name="txtRecordStatusInd" type="text" style="width: 150px" disabled value="<%If IsNumeric(strScheduleID) Then Response.Write objMaster.Fields("CREATE_REAL_USERID").Value%>"></TD>
</TR>
<TR align=right bgcolor=white nowrap>
	<TD>&nbsp;</TD>
	<TD>&nbsp;</TD>
	<TD style="width: 80px">Update Date:</TD>
	<TD style="width: 150px"><INPUT name="txtRecordStatusInd" type="text" style="width: 150px" disabled value="<%If IsNumeric(strScheduleID) Then Response.Write objMaster.Fields("UPDATE_DATE_TIME").Value%>"></TD>
	<TD style="width: 80px">Updated By:</TD>
	<TD style="width: 150px"><INPUT name="txtRecordStatusInd" type="text" style="width: 150px" disabled value="<%If IsNumeric(strScheduleID) Then Response.Write objMaster.Fields("UPDATE_REAL_USERID").Value%>"></TD>
</TR>
</TABLE>-->
<FIELDSET width="100%">
	<LEGEND align="right"><b>Audit Information</b></LEGEND>
	<DIV size="8pt" align="right">
	Record Status Indicator:<INPUT name="txtRecordStatusInd" type="text" style="width: 18px" disabled value="<%If IsNumeric(strScheduleID) Then Response.Write objMaster.Fields("RECORD_STATUS_IND").Value%>">&nbsp;&nbsp;&nbsp;
	Create Date:<INPUT name="txtRecordStatusInd" type="text" style="width: 150px" disabled value="<%If IsNumeric(strScheduleID) Then Response.Write objMaster.Fields("CREATE_DATE_TIME").Value%>">&nbsp;
	Created By:<INPUT name="txtRecordStatusInd" type="text" style="width: 150px" disabled value="<%If IsNumeric(strScheduleID) Then Response.Write objMaster.Fields("CREATE_REAL_USERID").Value%>"><br>
	Update Date:<INPUT name="txtRecordStatusInd" type="text" style="width: 150px" disabled value="<%If IsNumeric(strScheduleID) Then Response.Write objMaster.Fields("UPDATE_DATE_TIME").Value%>">&nbsp;
	Updated By:<INPUT name="txtRecordStatusInd" type="text" style="width: 150px" disabled value="<%If IsNumeric(strScheduleID) Then Response.Write objMaster.Fields("UPDATE_REAL_USERID").Value%>">
	</DIV>
</FIELDSET>-
</FORM>
</BODY>
<%
	'Clean up our ADO objects
	Set objMaster = Nothing
	objConn.Close
	Set ObjConn = Nothing
%>
</HTML>

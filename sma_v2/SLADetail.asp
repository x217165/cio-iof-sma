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
Dim strServiceLevelID, datUpdateDateTime, strWinMessage, strWinLocation
Dim	objCommand, objRS, objSchedule, strSQL, strFrom, strWhere, strOrderBy, strErrMessage, lIndex
Dim p_userid, p_sla_id, p_sla_desc, p_last_update_dt
Dim p_av_sched_id, p_av_percent, p_av_cons_days, p_rep_thresh_mins, p_rep_percent, p_rep_cons_days
Dim p_through_pps, p_through_percent, p_through_cons_days, p_lat_thresh_mins, p_lat_percentage
Dim p_lat_cons_days, p_resp_thresh_min, p_mon_sched_id, p_hdesk_sched_id, p_maint_sched_id, p_comments
Dim intAccessLevel

	'Check user's rights
	intAccessLevel = CInt(CheckLogon(strConst_ServiceCatalogue))
	If (intAccessLevel And intConst_Access_ReadOnly) <> intConst_Access_ReadOnly  Then
		DisplayError "BACK", "", 0, "ACCESS DENIED", "You don't have access to Service Level Agreement.  Please contact your system administrator"
	End If

	strWinMessage = ""
	strServiceLevelID = Request("ServiceLevelID")

	'Mandatory Parameters

	p_userid = Session("username")

	If IsNumeric(Request("hdnServiceLevelID")) Then
		p_sla_id = CLng(Request("hdnServiceLevelID"))
	Else
		p_sla_id = Null
	End If

	If Len(Request("txtSLADescription")) <> 0 Then
		p_sla_desc = Trim(Request("txtSLADescription"))
	Else
		p_sla_desc = Null
	End If

	If IsDate(Request("hdnUpdateDateTime")) Then
		p_last_update_dt = CDate(Request("hdnUpdateDateTime"))
	Else
		p_last_update_dt = Null
	End If

	'Optional Parameters
	If Len(Request.Form("selAvailableScheduleID")) > 0 Then
		p_av_sched_id = CLng(Request.Form("selAvailableScheduleID"))
	Else
		p_av_sched_id = Null
	End If

	If Len(Request.Form("txtAvailablePercentage")) > 0 Then
		p_av_percent = CCur(Request.Form("txtAvailablePercentage"))
	Else
		p_av_percent = Null
	End If

	If Len(Request.Form("txtAvailableDays")) > 0 Then
		p_av_cons_days = CCur(Request.Form("txtAvailableDays"))
	Else
		p_av_cons_days = Null
	End If

	If Len(Request.Form("txtRepairThreshold")) > 0 Then
		p_rep_thresh_mins = CCur(Request.Form("txtRepairThreshold"))
	Else
		p_rep_thresh_mins = Null
	End If

	If Len(Request.Form("txtRepairPercentage")) > 0 Then
		p_rep_percent = CCur(Request.Form("txtRepairPercentage"))
	Else
		p_rep_percent = Null
	End If

	If Len(Request.Form("txtRepairDays")) > 0 Then
		p_rep_cons_days = CCur(Request.Form("txtRepairDays"))
	Else
		p_rep_cons_days = Null
	End If

	If Len(Request.Form("txtThroughputThreshold")) > 0 Then
		p_through_pps = CCur(Request.Form("txtThroughputThreshold"))
	Else
		p_through_pps = Null
	End If

	If Len(Request.Form("txtThroughputPercentage")) > 0 Then
		p_through_percent = CCur(Request.Form("txtThroughputPercentage"))
	Else
		p_through_percent = Null
	End If

	If Len(Request.Form("txtThroughputDays")) > 0 Then
		p_through_cons_days = CCur(Request.Form("txtThroughputDays"))
	Else
		p_through_cons_days = Null
	End If

	If Len(Request.Form("txtLatencyThreshold")) > 0 Then
		p_lat_thresh_mins = CCur(Request.Form("txtLatencyThreshold"))
	Else
		p_lat_thresh_mins = Null
	End If

	If Len(Request.Form("txtLatencyPercentage")) > 0 Then
		p_lat_percentage = CCur(Request.Form("txtLatencyPercentage"))
	Else
		p_lat_percentage = Null
	End If

	If Len(Request.Form("txtLatencyDays")) > 0 Then
		p_lat_cons_days = CCur(Request.Form("txtLatencyDays"))
	Else
		p_lat_cons_days = Null
	End If

	If Len(Request.Form("txtResponseThreshold")) > 0 Then
		p_resp_thresh_min = CCur(Request.Form("txtResponseThreshold"))
	Else
		p_resp_thresh_min = Null
	End If

	If Len(Request.Form("selMonitorScheduleID")) > 0 Then
		p_mon_sched_id = CLng(Request.Form("selMonitorScheduleID"))
	Else
		p_mon_sched_id = Null
	End If

	If Len(Request.Form("selHelpDeskScheduleID")) > 0 Then
		p_hdesk_sched_id = CLng(Request.Form("selHelpDeskScheduleID"))
	Else
		p_hdesk_sched_id = Null
	End If

	If Len(Request.Form("selMaintenanceScheduleID")) > 0 Then
		p_maint_sched_id = CLng(Request.Form("selMaintenanceScheduleID"))
	Else
		p_maint_sched_id = Null
	End If

	If Len(Request.Form("txtSLAComments")) > 0 Then
		If Len(Request.Form("txtSLAComments")) > 2000 Then
			p_comments = Left(Trim(Request.Form("txtSLAComments")), 2000)
		Else
			p_comments = Trim(Request.Form("txtSLAComments"))
		End If
	Else
		p_comments = Null
	End If

	Select Case UCase(Request("hdnFrmAction"))
		Case "SAVE"
			Set objCommand = Server.CreateObject("ADODB.Command")
			Set objCommand.ActiveConnection = objConn
			objCommand.CommandType = adCmdStoredProc

			If IsNumeric(Request("hdnServiceLevelID")) Then	'Save existing Service Level
				If (intAccessLevel And intConst_Access_Update) <> intConst_Access_Update Then
					DisplayError "REFRESH", strWinLocation, 0, "UPDATE DENIED", "You don't have access to update Service Level Agreement.  Please contact your system administrator"
				End If

				objCommand.CommandText = "sma_sp_userid.spk_sma_admin_inter.sp_sla_update"
				objCommand.Parameters.Append objCommand.CreateParameter("p_userid", adVarChar, adParamInput, 20, Session("username"))
				objCommand.Parameters.Append objCommand.CreateParameter("p_sla_id", adNumeric, adParamInput, , CLng(Request("hdnServiceLevelID")))
				objCommand.Parameters.Append objCommand.CreateParameter("p_sla_desc", adVarChar, adParamInput, 80, Trim(Request("txtSLADescription")))
				objCommand.Parameters.Append objCommand.CreateParameter("p_last_update_dt", adDBTimeStamp, adParamInput, , CDate(Request("hdnUpdateDateTime")))

				strErrMessage = "CANNOT UPDATE OBJECT"
			Else										'Create a new Service Level
				If (intAccessLevel And intConst_Access_Create) <> intConst_Access_Create Then
					DisplayError "REFRESH", strWinLocation, 0, "INSERT DENIED", "You don't have access to create Service Level Agreement.  Please contact your system administrator"
				End If

				objCommand.CommandText = "sma_sp_userid.spk_sma_admin_inter.sp_sla_insert"
				objCommand.Parameters.Append objCommand.CreateParameter("p_userid", adVarChar, adParamInput, 20, Session("username"))
				objCommand.Parameters.Append objCommand.CreateParameter("p_sla_id", adNumeric, adParamOutput, , Null)
				objCommand.Parameters.Append objCommand.CreateParameter("p_sla_desc", adVarChar, adParamInput, 80, Trim(Request.Form("txtSLADescription")))

				strErrMessage = "CANNOT CREATE OBJECT"
			End If

			objCommand.Parameters.Append objCommand.CreateParameter("p_av_sched_id", adNumeric, adParamInput, , p_av_sched_id)
			objCommand.Parameters.Append objCommand.CreateParameter("p_av_percent", adDecimal, adParamInput, , p_av_percent)
			objCommand.Parameters.Append objCommand.CreateParameter("p_av_cons_days", adNumeric, adParamInput, , p_av_cons_days)
			objCommand.Parameters.Append objCommand.CreateParameter("p_rep_thresh_mins", adNumeric, adParamInput, , p_rep_thresh_mins)
			objCommand.Parameters.Append objCommand.CreateParameter("p_rep_percent", adDecimal, adParamInput, , p_rep_percent)
			objCommand.Parameters.Append objCommand.CreateParameter("p_rep_cons_days", adNumeric, adParamInput, , p_rep_cons_days)
			objCommand.Parameters.Append objCommand.CreateParameter("p_through_pps", adNumeric, adParamInput, , p_through_pps)
			objCommand.Parameters.Append objCommand.CreateParameter("p_through_percent", adDecimal, adParamInput, , p_through_percent)
			objCommand.Parameters.Append objCommand.CreateParameter("p_through_cons_days", adNumeric, adParamInput, , p_through_cons_days)
			objCommand.Parameters.Append objCommand.CreateParameter("p_lat_thresh_mins", adNumeric, adParamInput, , p_lat_thresh_mins)
			objCommand.Parameters.Append objCommand.CreateParameter("p_lat_percentage", adDecimal, adParamInput, , p_lat_percentage)
			objCommand.Parameters.Append objCommand.CreateParameter("p_lat_cons_days", adNumeric, adParamInput, , p_lat_cons_days)
			objCommand.Parameters.Append objCommand.CreateParameter("p_resp_thresh_min", adNumeric, adParamInput, , p_resp_thresh_min)
			objCommand.Parameters.Append objCommand.CreateParameter("p_mon_sched_id", adNumeric, adParamInput, , p_mon_sched_id)
			objCommand.Parameters.Append objCommand.CreateParameter("p_hdesk_sched_id", adNumeric, adParamInput, , p_hdesk_sched_id)
			objCommand.Parameters.Append objCommand.CreateParameter("p_maint_sched_id", adNumeric, adParamInput, , p_maint_sched_id)
			objCommand.Parameters.Append objCommand.CreateParameter("p_comments", adVarChar, adParamInput, 2000, p_comments)

			On Error Resume Next
			objCommand.Execute
			If objConn.Errors.Count <> 0 then
				DisplayError "BACK", "", objConn.Errors(0).NativeError, strErrMessage, objConn.Errors(0).Description
				objConn.Errors.Clear
			End If
			strServiceLevelID = CStr(objCommand.Parameters("p_sla_id").Value)
			strWinMessage = "Record saved successfully. You can now see the changes you made."

		Case "DELETE"
			If (intAccessLevel And intConst_Access_Delete) <> intConst_Access_Delete Then
				DisplayError "REFRESH", strWinLocation, 0, "DELETE DENIED", "You don't have access to delete Service Level Agreement.  Please contact your system administrator"
			End If

			Set objCommand = Server.CreateObject("ADODB.Command")
			Set objCommand.ActiveConnection = objConn
			objCommand.CommandType = adCmdStoredProc
			objCommand.CommandText = "sma_sp_userid.spk_sma_admin_inter.sp_sla_delete"

			objCommand.Parameters.Append objCommand.CreateParameter("p_sla_id", adNumeric, adParamInput, , p_sla_id)					'number(9)
			objCommand.Parameters.Append objCommand.CreateParameter("p_last_update_dt", adDBTimeStamp, adParamInput, , p_last_update_dt)		'Date

  			On Error Resume Next
			objCommand.Execute
			If objConn.Errors.Count <> 0 Then
				DisplayError "BACK", "", objConn.Errors(0).NativeError, "CANNOT DELETE OBJECT", objConn.Errors(0).Description
				objConn.Errors.Clear
			End If
			strServiceLevelID = "DEL"
			strWinMessage = "Record deleted successfully."
	End Select

	If IsNumeric(strServiceLevelID) Then
		strSQL = "SELECT " &_
				"SLA.SERVICE_LEVEL_AGREEMENT_ID, " &_
				"SLA.SERVICE_LEVEL_AGREEMENT_DESC, " &_
				"SLA.AVAILABLE_SCHEDULE_ID, " &_
				"SCH_SLA.SCHEDULE_NAME AS AVAILABLE_SCHEDULE_NAME, " &_
				"SLA.AVAILABLE_PERCENTAGE, " &_
				"SLA.AVAILABLE_CONSECUTIVE_DAYS, " &_
				"SLA.REPAIR_THRESHOLD_MINS, " &_
				"SLA.REPAIR_PERCENTAGE, " &_
				"SLA.REPAIR_CONSECUTIVE_DAYS, " &_
				"SLA.THROUGHPUT_THRESHOLD_PPS, " &_
				"SLA.THROUGHPUT_PERCENTAGE, " &_
				"SLA.THROUGHPUT_CONSECUTIVE_DAYS, " &_
				"SLA.LATENCY_THRESHOLD_MS, " &_
				"SLA.LATENCY_PERCENTAGE, " &_
				"SLA.LATENCY_CONSECUTIVE_DAYS, " &_
				"SLA.RESPONSE_THRESHOLD_MINS, " &_
				"SLA.MONITOR_SCHEDULE_ID, " &_
				"SCH_MON.SCHEDULE_NAME AS MONITOR_SCHEDULE_NAME, " &_
				"SLA.HELP_DESK_SCHEDULE_ID, " &_
				"SCH_HLP.SCHEDULE_NAME AS HELPDESK_SCHEDULE_NAME, " &_
				"SLA.MAINTENANCE_SCHEDULE_ID, " &_
				"SCH_MAI.SCHEDULE_NAME AS MAINTENANCE_SCHEDULE_NAME, " &_
				"SLA.COMMENTS, " &_
				"TO_CHAR(SLA.CREATE_DATE_TIME, 'MON-DD-YYYY HH24:MI:SS') AS CREATE_DATE_TIME, " &_
				"sma_sp_userid.spk_sma_library.sf_get_full_username(SLA.CREATE_REAL_USERID) AS CREATE_REAL_USERID, " &_
				"TO_CHAR(SLA.UPDATE_DATE_TIME, 'MON-DD-YYYY HH24:MI:SS') AS UPDATE_DATE_TIME, " &_
				"sma_sp_userid.spk_sma_library.sf_get_full_username(SLA.UPDATE_REAL_USERID) AS UPDATE_REAL_USERID, " &_
				"SLA.RECORD_STATUS_IND, " &_
				"SLA.UPDATE_DATE_TIME AS LAST_UPDATE_DATE_TIME " &_
				"FROM " &_
				"CRP.SERVICE_LEVEL_AGREEMENT SLA, " &_
				"CRP.SCHEDULE SCH_SLA, " &_
				"CRP.SCHEDULE SCH_MON, " &_
				"CRP.SCHEDULE SCH_HLP, " &_
				"CRP.SCHEDULE SCH_MAI " &_
				"WHERE " &_
				"SLA.AVAILABLE_SCHEDULE_ID = SCH_SLA.SCHEDULE_ID (+)	AND " &_
				"SLA.MONITOR_SCHEDULE_ID = SCH_MON.SCHEDULE_ID (+) 		AND " &_
				"SLA.HELP_DESK_SCHEDULE_ID = SCH_HLP.SCHEDULE_ID (+)	AND " &_
				"SLA.MAINTENANCE_SCHEDULE_ID = SCH_MAI.SCHEDULE_ID (+)	AND " &_
				"SLA.RECORD_STATUS_IND = 'A'							AND " &_
				"SLA.SERVICE_LEVEL_AGREEMENT_ID = " & strServiceLevelID

		'Create Recordset object
		Set objRS = Server.CreateObject("ADODB.Recordset")
		'Create the command object
		Set objCommand = Server.CreateObject("ADODB.Command")
		Set objCommand.ActiveConnection = objConn
		objCommand.CommandText = strSQL
		objCommand.CommandType = adCmdText

		On Error Resume Next
		Set objRS = objCommand.Execute
		If objConn.Errors.Count <> 0 Then
			DisplayError "BACK", "", objConn.Errors(0).NativeError, "ERROR LOADING DATA (Service Level)", objConn.Errors(0).Description
			objConn.Errors.Clear
		End If
	End If

	strSQL = "SELECT S.SCHEDULE_ID, S.SCHEDULE_NAME " &_
			"FROM CRP.SCHEDULE S " &_
			"WHERE S.RECORD_STATUS_IND = 'A' " &_
			"ORDER BY S.SCHEDULE_NAME"

	Set objSchedule = Server.CreateObject("ADODB.Recordset")
	objSchedule.Open strSQL, objConn, adOpenStatic, adLockReadOnly, adCmdText
	If objConn.Errors.Count <> 0 Then
		DisplayError "BACK", "", objConn.Errors(0).NativeError, "ERROR LOADING DATA (Schedules)", objConn.Errors(0).Description
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

setPageTitle("SMA - Service Level Agreement");

function fct_selNavigate(){
//***********************************************************************************************
// Function:	selNavigate_onChange															*
//																								*
// Purpose:		To display the page selected by the user from Quick Navigation drop-down box.	*
//																								*
// Created By:	Gilles Archer 09/27/2000														*
//																								*
// Updated By:																					*
//***********************************************************************************************
var strPageName = document.frmSLADetail.selNavigate.item(document.frmSLADetail.selNavigate.selectedIndex).value;

	switch (strPageName) {
		case "ServiceTypes":
			document.frmSLADetail.selNavigate.selectedIndex = 0;
			var strServiceLevelID = document.frmSLADetail.hdnServiceLevelID.value;
			SetCookie("ServiceLevelID", strServiceLevelID);
			self.location.href = "SearchFrame.asp?fraSrc=ServiceType";
			break;
		case "Available":
			document.frmSLADetail.selNavigate.selectedIndex = 0;
			var strScheduleID = document.frmSLADetail.selAvailableScheduleID.value;
			self.location.href = "ScheduleDetail.asp?ScheduleID=" + strScheduleID;
			break;
		case "Monitor":
			document.frmSLADetail.selNavigate.selectedIndex = 0;
			var strScheduleID = document.frmSLADetail.selMonitorScheduleID.value;
			self.location.href = "ScheduleDetail.asp?ScheduleID=" + strScheduleID;
			break;
		case "HelpDesk":
			document.frmSLADetail.selNavigate.selectedIndex = 0;
			var strScheduleID = document.frmSLADetail.selHelpDeskScheduleID.value;
			self.location.href = "ScheduleDetail.asp?ScheduleID=" + strScheduleID;
			break;
		case "Maintenance":
			document.frmSLADetail.selNavigate.selectedIndex = 0;
			var strScheduleID = document.frmSLADetail.selMaintenanceScheduleID.value;
			self.location.href = "ScheduleDetail.asp?ScheduleID=" + strScheduleID;
			break;
		case "DEFAULT":							//Do nothing
			break;
		default:								//Do nothing
			break;
	}
}

function btnDelete_onClick() {
//**********************************************************************************************
// Function:	btnDelete_onClick
//
// Purpose:		To delete a Service Level Agreement
//
// Created By:	Gilles Archer 09/27/2000
//
// Updated By:
//***********************************************************************************************
	if ((intAccessLevel & intConst_Access_Delete) != intConst_Access_Delete) {
		alert('You do not have permission to DELETE a Service Level Agreement.  Please contact your System Administrator.');
		return false;
	}

	if (document.frmSLADetail.hdnServiceLevelID.value == "") {
		alert('This Service Level Agreement does not exist in the database.');
		return false;
	}

	if (confirm('Do you really want to delete this object?')){
		document.frmSLADetail.hdnFrmAction.value = "DELETE";
		document.frmSLADetail.submit();
	}
}

function btnNew_onClick() {
	if ((intAccessLevel & intConst_Access_Create) != intConst_Access_Create) {
		alert('You do not have permission to CREATE a Service Level Agreement.  Please contact your System Administrator.');
		return false;
	}
	document.location = "SLADetail.asp?ServiceLevelID=NEW";
}

function fct_onChange() {
	bolSaveRequired = true;
}

function btnSave_onClick() {
	if ((intAccessLevel & intConst_Access_Update) != intConst_Access_Update) {
		alert('You do not have permission to UPDATE a Service Level Agreement.  Please contact your System Administrator.');
		return false;
	}

	if (document.frmSLADetail.txtSLADescription.value == "") {
		alert('Please enter a description for this SLA.');
		document.frmSLADetail.txtSLADescription.focus();
		return false;
	}

	if (document.frmSLADetail.txtRepairThreshold.value != "") {
		if (isNaN(Number(document.frmSLADetail.txtRepairThreshold.value))) {
			alert('Please enter a numerical value for Repair Threshold');
			document.frmSLADetail.txtRepairThreshold.focus();
			return false;
		}
	}

	if (document.frmSLADetail.txtAvailablePercentage.value != "") {
		if (isNaN(Number(document.frmSLADetail.txtAvailablePercentage.value))) {
			alert('Please enter a numerical value for Available Percentage');
			document.frmSLADetail.txtAvailablePercentage.focus();
			return false;
		}
	}

	if (document.frmSLADetail.txtRepairPercentage.value != "") {
		if (isNaN(Number(document.frmSLADetail.txtRepairPercentage.value))) {
			alert('Please enter a numerical value for Repair Percentage');
			document.frmSLADetail.txtRepairPercentage.focus();
			return false;
		}
	}

	if (document.frmSLADetail.txtAvailableDays.value != "") {
		if (isNaN(Number(document.frmSLADetail.txtAvailableDays.value))) {
			alert('Please enter a numerical value for Available Days');
			document.frmSLADetail.txtAvailableDays.focus();
			return false;
		}
	}

	if (document.frmSLADetail.txtRepairDays.value != "") {
		if (isNaN(Number(document.frmSLADetail.txtRepairDays.value))) {
			alert('Please enter a numerical value for Repair Days');
			document.frmSLADetail.txtRepairDays.focus();
			return false;
		}
	}

	if (document.frmSLADetail.txtThroughputThreshold.value != "") {
		if (isNaN(Number(document.frmSLADetail.txtThroughputThreshold.value))) {
			alert('Please enter a numerical value for Throughput Threshold');
			document.frmSLADetail.txtThroughputThreshold.focus();
			return false;
		}
	}

	if (document.frmSLADetail.txtLatencyThreshold.value != "") {
		if (isNaN(Number(document.frmSLADetail.txtLatencyThreshold.value))) {
			alert('Please enter a numerical value for Latency Threshold');
			document.frmSLADetail.txtLatencyThreshold.focus();
			return false;
		}
	}

	if (document.frmSLADetail.txtThroughputPercentage.value != "") {
		if (isNaN(Number(document.frmSLADetail.txtThroughputPercentage.value))) {
			alert('Please enter a numerical value for Throughput Percentage');
			document.frmSLADetail.txtThroughputPercentage.focus();
			return false;
		}
	}

	if (document.frmSLADetail.txtLatencyPercentage.value != "") {
		if (isNaN(Number(document.frmSLADetail.txtLatencyPercentage.value))) {
			alert('Please enter a numerical value for Latency Percentage');
			document.frmSLADetail.txtLatencyPercentage.focus();
			return false;
		}
	}

	if (document.frmSLADetail.txtThroughputDays.value != "") {
		if (isNaN(Number(document.frmSLADetail.txtThroughputDays.value))) {
			alert('Please enter a numerical value for Throughput Days');
			document.frmSLADetail.txtThroughputDays.focus();
			return false;
		}
	}

	if (document.frmSLADetail.txtLatencyDays.value != "") {
		if (isNaN(Number(document.frmSLADetail.txtLatencyDays.value))) {
			alert('Please enter a numerical value for Latency Days');
			document.frmSLADetail.txtLatencyDays.focus();
			return false;
		}
	}

	if (document.frmSLADetail.txtResponseThreshold.value != "") {
		if (isNaN(Number(document.frmSLADetail.txtResponseThreshold.value))) {
			alert('Please enter a numerical value for Response Threshold');
			document.frmSLADetail.txtResponseThreshold.focus();
			return false;
		}
	}

	if (document.frmSLADetail.selAvailableScheduleID.selectedIndex == 0) {
	}
	if (document.frmSLADetail.selMonitorScheduleID.selectedIndex == 0) {
	}
	if (document.frmSLADetail.selHelpDeskScheduleID.selectedIndex == 0) {
	}
	if (document.frmSLADetail.selMaintenanceScheduleID.selectedIndex == 0) {
	}

	var strComments = document.frmSLADetail.txtSLAComments.value;
	if (strComments.length > 2000) {
		alert('The SLA Comment can be at most 2000 characters.\n\nYou entered ' + strComments.length + ' character(s).');
		document.frmSLADetail.txtSLAComments.focus();
		return false;
	}

	document.frmSLADetail.hdnFrmAction.value = "SAVE";
	bolSaveRequired = false;
	document.frmSLADetail.submit();
	return true;
}

function btnReferences_onClick() {
var strOwner = 'CRP';			// owner name must be in Uppercase
var strTableName = 'SERVICE_LEVEL_AGREEMENT';		// replace ADDRESS with your own table name and table name must be in Uppercase
var strRecordID = document.frmSLADetail.hdnServiceLevelID.value ;   // insert your record id
var strURL;

	if (strRecordID == "") {
		alert("No references. This is a new record.");
		return false;
	}

	strURL = "Dependency.asp?Owner=" + strOwner + "&TableName=" + strTableName + "&RecordID=" + strRecordID;
	window.open(strURL, 'Popup', 'top=100, left=100, width=500, height=300');
}

function window_onLoad() {
//
}

function window_onBeforeUnload() {
	//Ensure that fct_onChange() fires for any changed data.
	document.frmSLADetail.btnSave.focus();

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
	document.location.href = "SLADetail.asp?ServiceLevelID=<%=strServiceLevelID%>";
}
// Unhide Client-Side SCRIPT -->
</SCRIPT>
</HEAD>
<BODY language="javascript" onLoad="window_onLoad();DisplayStatus(strWinMessage);" onBeforeUnload="window_onBeforeUnload();" onUnload="window_onUnload();">
<FORM id="frmSLADetail" name="frmSLADetail" action="SLADetail.asp" method="post">
	<INPUT type="hidden" id="hdnServiceLevelID" name="hdnServiceLevelID" value="<%If IsNumeric(strServiceLevelID) Then Response.Write objRS.Fields("SERVICE_LEVEL_AGREEMENT_ID").Value%>">
	<INPUT type="hidden" id="hdnUpdateDateTime" name="hdnUpdateDateTime" value="<%If IsNumeric(strServiceLevelID) Then Response.Write objRS.Fields("LAST_UPDATE_DATE_TIME").Value%>">
	<INPUT type="hidden" id="hdnFrmAction" name="hdnFrmAction" value="">
<TABLE border="0" cellpadding="2" cellspacing="0" cols="4" width="100%">
<THEAD>
<TR valign="top">
	<TD align="left" colspan="3">Service Level Agreement Detail</TD>
	<TD align="right">
	<SELECT id="selNavigate" name="selNavigate" onChange="fct_selNavigate();">
		<OPTION value="DEFAULT" selected>Quickly Goto ...</OPTION>
		<OPTION value="ServiceTypes">Service Types</OPTION>
		<OPTION value="Available">Available Schedule</OPTION>
		<OPTION value="Monitor">Monitor Schedule</OPTION>
		<OPTION value="HelpDesk">Help Desk Schedule</OPTION>
		<OPTION value="Maintenance">Maintenance Schedule</OPTION>
	</SELECT></TD>
</TR>
</THEAD>
<TBODY>
<TR valign="top">
	<TD align="right" width="20%">SLA Description<FONT color="red">*</FONT></TD>
	<TD align="left" colspan="3"><INPUT id="txtSLADescription" name="txtSLADescription" type="text"maxlength="80" size="80" style="width: 100%" onChange="fct_onChange();" value="<%If IsNumeric(strServiceLevelID) Then Response.Write objRS.Fields("SERVICE_LEVEL_AGREEMENT_DESC").Value%>"></TD>
</TR>
<TR valign="top">
	<TD align="right" width="20%">Available Schedule</TD>
	<TD align="left" width="30%"><SELECT id="selAvailableScheduleID" name="selAvailableScheduleID" style="width: 100%" onChange="fct_onChange();">
		<OPTION></OPTION>
		<%objSchedule.MoveFirst
		Do While Not objSchedule.EOF
			If IsNumeric(strServiceLevelID) Then
				If Not IsNull(objRS.Fields("AVAILABLE_SCHEDULE_ID").Value) Then
					If StrComp(CStr(objSchedule.Fields("SCHEDULE_ID").Value), CStr(objRS.Fields("AVAILABLE_SCHEDULE_ID").Value), 0) = 0 Then
						Response.Write "<OPTION selected value='" & objSchedule.Fields("SCHEDULE_ID").Value & "'>" & objSchedule.Fields("SCHEDULE_NAME").Value & "</OPTION>"
					Else
						Response.Write "<OPTION value='" & objSchedule.Fields("SCHEDULE_ID").Value & "'>" & objSchedule.Fields("SCHEDULE_NAME").Value & "</OPTION>"
					End If
				Else
					Response.Write "<OPTION value='" & objSchedule.Fields("SCHEDULE_ID").Value & "'>" & objSchedule.Fields("SCHEDULE_NAME").Value & "</OPTION>"
				End If
			Else
				Response.Write "<OPTION value='" & objSchedule.Fields("SCHEDULE_ID").Value & "'>" & objSchedule.Fields("SCHEDULE_NAME").Value & "</OPTION>"
			End If
			objSchedule.MoveNext
		Loop%>
	</SELECT></TD>
	<TD align="right" width="20%">Repair Threshold (Minutes)</TD>
	<TD align="left" width="30%"><INPUT id="txtRepairThreshold" name="txtRepairThreshold" type="text"style="width: 100%" onChange="fct_onChange();" value="<%If IsNumeric(strServiceLevelID) Then Response.Write ObjRS.Fields("REPAIR_THRESHOLD_MINS").Value%>"></TD>
</TR>
<TR valign="top">
	<TD align="right" width="20%">Available Percentage</TD>
	<TD align="left" width="30%"><INPUT id="txtAvailablePercentage" name="txtAvailablePercentage" type="text"style="width: 100%" onChange="fct_onChange();" value="<%If IsNumeric(strServiceLevelID) Then Response.Write ObjRS.Fields("AVAILABLE_PERCENTAGE").Value%>"></TD>
	<TD align="right" width="20%">Repair Percentage</TD>
	<TD align="left" width="30%"><INPUT id="txtRepairPercentage" name="txtRepairPercentage" type="text"style="width: 100%" onChange="fct_onChange();" value="<%If IsNumeric(strServiceLevelID) Then Response.Write ObjRS.Fields("REPAIR_PERCENTAGE").Value%>"></TD>
</TR>
<TR valign="top">
	<TD align="right" width="20%">Available Consecutive Days</TD>
	<TD align="left" width="30%"><INPUT id="txtAvailableDays" name="txtAvailableDays" type="text"style="width: 100%" onChange="fct_onChange();" value="<%If IsNumeric(strServiceLevelID) Then Response.Write ObjRS.Fields("AVAILABLE_CONSECUTIVE_DAYS").Value%>"></TD>
	<TD align="right" width="20%">Repair Consecutive Days</TD>
	<TD align="left" width="30%"><INPUT id="txtRepairDays" name="txtRepairDays" type="text"style="width: 100%" onChange="fct_onChange();" value="<%If IsNumeric(strServiceLevelID) Then Response.Write ObjRS.Fields("REPAIR_CONSECUTIVE_DAYS").Value%>"></TD>
</TR>
<TR valign="top">
	<TD align="right" width="20%">Throughput Threshold (PPS)</TD>
	<TD align="left" width="30%"><INPUT id="txtThroughputThreshold" name="txtThroughputThreshold" type="text"style="width: 100%" onChange="fct_onChange();" value="<%If IsNumeric(strServiceLevelID) Then Response.Write ObjRS.Fields("THROUGHPUT_THRESHOLD_PPS").Value%>"></TD>
	<TD align="right" width="20%">Latency Threshold (ms)</TD>
	<TD align="left" width="30%"><INPUT id="txtLatencyThreshold" name="txtLatencyThreshold" type="text"style="width: 100%" onChange="fct_onChange();" value="<%If IsNumeric(strServiceLevelID) Then Response.Write ObjRS.Fields("LATENCY_THRESHOLD_MS").Value%>"></TD>
</TR>
<TR valign="top">
	<TD align="right" width="20%">Throughput Percentage</TD>
	<TD align="left" width="30%"><INPUT id="txtThroughputPercentage" name="txtThroughputPercentage" type="text"style="width: 100%" onChange="fct_onChange();" value="<%If IsNumeric(strServiceLevelID) Then Response.Write ObjRS.Fields("THROUGHPUT_PERCENTAGE").Value%>"></TD>
	<TD align="right" width="20%">Latency Percentage</TD>
	<TD align="left" width="30%"><INPUT id="txtLatencyPercentage" name="txtLatencyPercentage" type="text"style="width: 100%" onChange="fct_onChange();" value="<%If IsNumeric(strServiceLevelID) Then Response.Write ObjRS.Fields("LATENCY_PERCENTAGE").Value%>"></TD>
</TR>
<TR valign="top">
	<TD align="right" width="20%">Throughput Consecutive Days</TD>
	<TD align="left" width="30%"><INPUT id="txtThroughputDays" name="txtThroughputDays" type="text"style="width: 100%" onChange="fct_onChange();" value="<%If IsNumeric(strServiceLevelID) Then Response.Write ObjRS.Fields("THROUGHPUT_CONSECUTIVE_DAYS").Value%>"></TD>
	<TD align="right" width="20%">Latency Consecutive Days</TD>
	<TD align="left" width="30%"><INPUT id="txtLatencyDays" name="txtLatencyDays" type="text"style="width: 100%" onChange="fct_onChange();" value="<%If IsNumeric(strServiceLevelID) Then Response.Write ObjRS.Fields("LATENCY_CONSECUTIVE_DAYS").Value%>"></TD>
</TR>
<TR valign="top">
	<TD align="right" width="20%">Response Threshold (Minutes)</TD>
	<TD align="left" width="30%"><INPUT id="txtResponseThreshold" name="txtResponseThreshold" type="text"style="width: 100%" onChange="fct_onChange();" value="<%If IsNumeric(strServiceLevelID) Then Response.Write ObjRS.Fields("RESPONSE_THRESHOLD_MINS").Value%>"></TD>
	<TD align="right" width="20%">Monitor Schedule</TD>
	<TD align="left" width="30%"><SELECT id="selMonitorScheduleID" name="selMonitorScheduleID" style="width: 100%" onChange="fct_onChange();">
		<OPTION></OPTION>
		<%objSchedule.MoveFirst
		Do While Not objSchedule.EOF
			If IsNumeric(strServiceLevelID) Then
				If Not IsNull(objRS.Fields("MONITOR_SCHEDULE_ID").Value) Then
					If StrComp(CStr(objSchedule.Fields("SCHEDULE_ID").Value), CStr(objRS.Fields("MONITOR_SCHEDULE_ID").Value), 0) = 0 Then
						Response.Write "<OPTION selected value='" & objSchedule.Fields("SCHEDULE_ID").Value & "'>" & objSchedule.Fields("SCHEDULE_NAME").Value & "</OPTION>"
					Else
						Response.Write "<OPTION value='" & objSchedule.Fields("SCHEDULE_ID").Value & "'>" & objSchedule.Fields("SCHEDULE_NAME").Value & "</OPTION>"
					End If
				Else
					Response.Write "<OPTION value='" & objSchedule.Fields("SCHEDULE_ID").Value & "'>" & objSchedule.Fields("SCHEDULE_NAME").Value & "</OPTION>"
				End If
			Else
				Response.Write "<OPTION value='" & objSchedule.Fields("SCHEDULE_ID").Value & "'>" & objSchedule.Fields("SCHEDULE_NAME").Value & "</OPTION>"
			End If
			objSchedule.MoveNext
		Loop%>
	</SELECT></TD>
</TR>
<TR valign="top">
	<TD align="right" width="20%">Help Desk Schedule</TD>
	<TD align="left" width="30%"><SELECT id="selHelpDeskScheduleID" name="selHelpDeskScheduleID" style="width: 100%" onChange="fct_onChange();">
		<OPTION></OPTION>
		<%objSchedule.MoveFirst
		Do While Not objSchedule.EOF
			If IsNumeric(strServiceLevelID) Then
				If Not IsNull(objRS.Fields("HELP_DESK_SCHEDULE_ID").Value) Then
					If StrComp(CStr(objSchedule.Fields("SCHEDULE_ID").Value), CStr(objRS.Fields("HELP_DESK_SCHEDULE_ID").Value), 0) = 0 Then
						Response.Write "<OPTION selected value='" & objSchedule.Fields("SCHEDULE_ID").Value & "'>" & objSchedule.Fields("SCHEDULE_NAME").Value & "</OPTION>"
					Else
						Response.Write "<OPTION value='" & objSchedule.Fields("SCHEDULE_ID").Value & "'>" & objSchedule.Fields("SCHEDULE_NAME").Value & "</OPTION>"
					End If
				Else
					Response.Write "<OPTION value='" & objSchedule.Fields("SCHEDULE_ID").Value & "'>" & objSchedule.Fields("SCHEDULE_NAME").Value & "</OPTION>"
				End If
			Else
				Response.Write "<OPTION value='" & objSchedule.Fields("SCHEDULE_ID").Value & "'>" & objSchedule.Fields("SCHEDULE_NAME").Value & "</OPTION>"
			End If
			objSchedule.MoveNext
		Loop%>
	</SELECT></TD>
	<TD align="right" width="20%">Maintenance Schedule</TD>
	<TD align="left" width="30%"><SELECT id="selMaintenanceScheduleID" name="selMaintenanceScheduleID" style="width: 100%" onChange="fct_onChange();">
		<OPTION></OPTION>
		<%objSchedule.MoveFirst
		Do While Not objSchedule.EOF
			If IsNumeric(strServiceLevelID) Then
				If Not IsNull(objRS.Fields("MAINTENANCE_SCHEDULE_ID").Value) Then
					If StrComp(CStr(objSchedule.Fields("SCHEDULE_ID").Value), CStr(objRS.Fields("MAINTENANCE_SCHEDULE_ID").Value), 0) = 0 Then
						Response.Write "<OPTION selected value='" & objSchedule.Fields("SCHEDULE_ID").Value & "'>" & objSchedule.Fields("SCHEDULE_NAME").Value & "</OPTION>"
					Else
						Response.Write "<OPTION value='" & objSchedule.Fields("SCHEDULE_ID").Value & "'>" & objSchedule.Fields("SCHEDULE_NAME").Value & "</OPTION>"
					End If
				Else
					Response.Write "<OPTION value='" & objSchedule.Fields("SCHEDULE_ID").Value & "'>" & objSchedule.Fields("SCHEDULE_NAME").Value & "</OPTION>"
				End If
			Else
				Response.Write "<OPTION value='" & objSchedule.Fields("SCHEDULE_ID").Value & "'>" & objSchedule.Fields("SCHEDULE_NAME").Value & "</OPTION>"
			End If
			objSchedule.MoveNext
		Loop%>
	</SELECT></TD>
</TR>
<TR valign="top">
	<TD align="right" width="20%">Comments</TD>
	<TD align="left" colspan="3">
	<TEXTAREA id="txtSLAComments" name="txtSLAComments" cols="100" rows="10" maxlength="2000" style="width: 100%" onChange="fct_onChange();"><%If IsNumeric(strServiceLevelID) Then Response.Write ObjRS.Fields("COMMENTS").Value%></TEXTAREA></TD>
</TR>
</TBODY>
<TFOOT>
<TR valign="top">
	<TD colspan="4" align="right">
	<INPUT id="btnReferences" name="btnReferences" type="button" value="References" style="width: 2.2cm" language="javascript" onClick="return btnReferences_onClick();">&nbsp;
	<INPUT id="btnDelete" name="btnDelete" type="button" value="Delete" style="width: 2cm" language="javascript" onClick="return btnDelete_onClick();">&nbsp;
	<INPUT id="btnReset" name="btnReset" type="button" value="Reset" style="width: 2cm" language="javascript" onClick="return btnReset_onClick();">&nbsp;
	<INPUT id="btnNew" name="btnNew" type="button" value="New" style="width: 2cm" language="javascript" onClick="return btnNew_onClick();">&nbsp;
	<INPUT id="btnSave" name="btnSave" type="button" value="Save" style="width: 2cm" onClick="return btnSave_onClick();"></TD>
</TR>
</TFOOT>
</TABLE>
<FIELDSET width="100%">
	<LEGEND align="right"><b>Audit Information</b></LEGEND>
	<DIV size="8pt" align="right">
	Record Status Indicator:<INPUT align="left" name="txtRecordStatusInd" type="text" style="width: 18px" disabled value="<%If IsNumeric(strServiceLevelID) Then Response.Write objRS.Fields("RECORD_STATUS_IND").Value%>">&nbsp;&nbsp;&nbsp;
	Create Date:<INPUT align="center" name="txtRecordStatusInd" type="text" style="width: 150px" disabled value="<%If IsNumeric(strServiceLevelID) Then Response.Write objRS.Fields("CREATE_DATE_TIME").Value%>">&nbsp;
	Created By:<INPUT align="right" name="txtRecordStatusInd" type="text" style="width: 150px" disabled value="<%If IsNumeric(strServiceLevelID) Then Response.Write objRS.Fields("CREATE_REAL_USERID").Value%>"><br>
	Update Date:<INPUT align="center" name="txtRecordStatusInd" type="text" style="width: 150px" disabled value="<%If IsNumeric(strServiceLevelID) Then Response.Write objRS.Fields("UPDATE_DATE_TIME").Value%>">&nbsp;
	Updated By:<INPUT align="right" name="txtRecordStatusInd" type="text" style="width: 150px" disabled value="<%If IsNumeric(strServiceLevelID) Then Response.Write objRS.Fields("UPDATE_REAL_USERID").Value%>">
	</DIV>
</FIELDSET>
</FORM>
<%
	'Clean up our ADO objects
	Set objRS = Nothing
	Set objCommand = Nothing
	objConn.Close
	Set ObjConn = Nothing
%>
</BODY>
</HTML>

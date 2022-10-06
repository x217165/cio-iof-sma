<%@ Language=VBScript %>
<%Option Explicit%>
<%Response.Buffer = True%>
<!--#include file = "smaConstants.inc"-->
<!--#include file = "smaProcs.inc"-->
<!--#include file = "databaseconnect.asp" -->
<!--
******************************************************************************
*
*
* In Param:		This pages reads following cookies
*
*
*
*******************************************************************************
-->
<%
Dim strWinName, strScheduleDescription, lIndex
Dim objCommand, objRS, objProvinces, objCountries, strSQL, strWhere, strOrderBy
Dim intAccessLevel

	intAccessLevel = CInt(CheckLogon(strConst_ServiceCatalogue))
	If (intAccessLevel And intConst_Access_ReadOnly) <> intConst_Access_ReadOnly Then
		DisplayError "BACK", "", 0, "ACCESS DENIED", "You don't have access to Schedule.  Please contact your system administrator"
	End If

	strWinName = Request("WinName")
	strScheduleDescription = Request("txtScheduleDescription")

	'Optional Criteria selection
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
<!--
var intAccessLevel = <%=intAccessLevel%>;
var strWinName = "<%=strWinName%>";
var strScheduleDescription = "<%=strScheduleDescription%>";

//set section title
setPageTitle("SMA - Schedule Definition");

function window_onLoad() {
//****************************************************************************************
//
//
//
//
//***************************************************************************************
	if (strScheduleDescription != "") {
		DeleteCookie("ScheduleDescription");
		DeleteCookie("WinName");
		document.frmScheduleSearch.submit();
	}
}

function btnNew_onClick() {
//************************************************************************************************
// Function:	btnAddNew_onClick()
//
// Purpose:		To bring up a blank Schedule Detail page so that user can enter a new Schedule.
//
// Created By:	Gilles Archer Oct 02 2000
//
// Updated By:
//************************************************************************************************
	if ((intAccessLevel & intConst_Access_Create) != intConst_Access_Create) {
		alert('You do not have permission to CREATE a Schedule.  Please contact your System Administrator.');
		return false;
	}
	parent.document.location.href ="ScheduleDetail.asp?ScheduleID=NEW";
}

function btnClear_onClick() {
	with (document.frmScheduleSearch) {
		txtScheduleDescription.value = "";
		selProvince.selectedIndex = 0;
		selCountry.selectedIndex = 0;
		chkHolidayOnly.checked = false;
		chkActiveOnly.checked = true;
	}
}
//-->
</SCRIPT>
</HEAD>
<BODY language="javascript" onLoad="window_onLoad()">
<FORM id="frmScheduleSearch" name="frmScheduleSearch" method="post" action="ScheduleList.asp" target="fraResult">
	<INPUT type="hidden" id="hdnWinName" name="hdnWinName" value="<%=strWinName%>">
	<INPUT type="hidden" id="hdnScheduleDescription" name="hdnScheduleDescription" value="<%=strScheduleDescription%>">
<TABLE cols="4" width="100%">
<THEAD>
	<TR><TD colspan="4" align="left">Schedule Search</TD></TR>
</THEAD>
<TBODY>
<TR>
	<TD align="right">Schedule Description</TD>
	<TD align="left"><INPUT id="txtScheduleDescription" name="txtScheduleDescription" value="<%=strScheduleDescription%>" size="30" maxlength="30"></TD>
</TR>
<TR>
	<TD align="right">Province / State</TD>
	<TD align="left">
	<SELECT id="selProvince" name="selProvince" style="width: 200">
		<OPTION></OPTION>
		<%Do While Not objProvinces.EOF
			Response.Write "<OPTION value='" & objProvinces.Fields("PROVINCE_STATE_LCODE").Value & "'>" & objProvinces.Fields("PROVINCE_STATE_LCODE").Value & " - " & objProvinces.Fields("PROVINCE_STATE_NAME").Value & "</OPTION>"
			objProvinces.MoveNext
		Loop
		objProvinces.Close
		Set objProvinces = Nothing%>
	</SELECT></TD>
</TR>
<TR>
	<TD align="right">Country</TD>
	<TD align="left">
	<SELECT id="selCountry" name="selCountry" style="width: 200">
		<OPTION></OPTION>
		<%Do While Not objCountries.EOF
			Response.Write "<OPTION value='" & objCountries.Fields("COUNTRY_LCODE").Value & "'>" & objCountries.Fields("COUNTRY_LCODE").Value & " - " & objCountries.Fields("COUNTRY_DESC").Value & "</OPTION>"
			objCountries.MoveNext
		Loop
		objCountries.Close
		Set objCountries = Nothing%>
	</SELECT></TD>
</TR>
<TR>
	<TD></TD>
	<TD>
		<INPUT id="chkHolidayOnly" name="chkHolidayOnly" type="checkbox">Holiday Only&nbsp;
		<INPUT id="chkActiveOnly" name="chkActiveOnly" type="checkbox" checked>Active Only
	</TD>

</TR>
</TBODY>
<TFOOT>
<TR>
	<TD colspan="4" align="right">
	<%If UCase(strWinName) <> UCase("Popup") Then%>
	<INPUT id="btnNew" name="btnNew" type="button" value="New" style="width: 2cm" language="javascript" onClick="btnNew_onClick()">&nbsp;
	<%End If%>
	<INPUT id="btnClear" name="btnClear" type="button" value="Clear" style="width: 2cm" language="javascript" onClick="btnClear_onClick();">&nbsp;
	<INPUT id="btnSearch" name="btnSearch" type="submit" value="Search" style="width: 2cm" language="javascript">&nbsp;</TD>
</TR>
</TFOOT>
</TABLE>
</FORM>
</BODY>
</HTML>

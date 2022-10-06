<%@ Language=VBScript %>
<% Option Explicit %>
<% Response.Buffer = True %>
<!--#include file = "smaConstants.inc" -->
<!--#include file = "databaseconnect.asp"-->
<!--#include file = "smaProcs.inc" -->
<!--****************************************************************************************
* Page name:	StaffCriteria.asp
* Purpose:		To dynamically set the criteria to search for a Staff.
*				Results are displayed via StaffList.asp
*
* Created by:	Gilles Archer Oct 23 2000
*****************************************************************************************
        Date		Author			Changes/enhancements made
        -----		------		------------------------------------------------------
       19-Feb-02	 DTy		Increase email address size from 50 t0 60.
*************************************************************************************************
-->
<%
Dim strLName, strFName, strWinName
Dim objDepartment, objStaffStatus, strSQL
Dim intAccessLevel

	intAccessLevel = CInt(CheckLogon(strConst_Security))
	If (intAccessLevel And intConst_Access_ReadOnly) <> intConst_Access_ReadOnly Then
		DisplayError "BACK", "", 0, "ACCESS DENIED", "You don't have access to Staff Maintenance. Please contact your system administrator"
	End If

	'navigation: read cookies
	strLName = Trim(Request.Cookies("LName"))
	strFName = Trim(Request.Cookies("FName"))
	strWinName = Trim(Request.Cookies("WinName"))

	strSQL = "SELECT DEPT.DEPARTMENT_ID, DEPT.DEPARTMENT_DESC " &_
			"FROM CRP.DEPARTMENT_LOOKUP DEPT " &_
			"WHERE DEPT.RECORD_STATUS_IND = 'A' " &_
			"ORDER BY DEPT.DEPARTMENT_DESC ASC"

	On Error Resume Next
	Set objDepartment = Server.CreateObject("ADODB.Recordset")
	objDepartment.Open strSQL, objConn, adOpenForwardOnly, adLockReadOnly, adCmdText
	If objConn.Errors.Count <> 0 Then
		DisplayError "BACK", "", objConn.Errors(0).NativeError, "ERROR LOADING DATA (Departments)", objConn.Errors(0).Description
		objConn.Errors.Clear
	End If

	'Staff Status
	strSQL = "SELECT STAFF_STATUS_LCODE, STAFF_STATUS_DESC " &_
			"FROM CRP.LCODE_STAFF_STATUS " &_
			"WHERE RECORD_STATUS_IND = 'A' " &_
			"ORDER BY STAFF_STATUS_DESC"

	Set objStaffStatus = Server.CreateObject("ADODB.Recordset")
	objStaffStatus.Open strSQL, objConn, adOpenForwardOnly, adLockReadOnly, adCmdText
	If objConn.Errors.Count <> 0 Then
		DisplayError "BACK", "", objConn.Errors(0).NativeError, "ERROR LOADING DATA (Staff Status)", objConn.Errors(0).Description
		objConn.Errors.Clear
	End If
%>
<HTML>
<HEAD>
<META name="GENERATOR" content="Microsoft Visual Studio 6.0">
<META http-equiv="Pragma" content="no-cache">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<LINK rel="stylesheet" type="text/css" href="stylesheets/styles.css">
<SCRIPT type="text/javascript" language="javascript" src="AccessLevels.js"></SCRIPT>
<SCRIPT type="text/javascript" language="javascript" src="GeneralJavaFunctions.js"></SCRIPT>
<SCRIPT type="text/javascript" language="javascript" id="clientEventHandlersJS">
<!--
var	intAccessLevel = <%=intAccessLevel%>;

//set page heading
setPageTitle("SMA - Staff Maintenance");

function window_onLoad() {
/**********************************************************************************************
*  Function:	window_onload																  *
*  Purpose:		Submit the form automatically when called via lookup or Quick Navigation box; *
*				EXCEPT do not submit if lookup calling field is blank.						  *
*  Created By:	Nancy Mooney 08/31/2000														  *
***********************************************************************************************/
var strNameLast = document.frmStaffCriteria.txtNameLast.value;
var strNameFirst = document.frmStaffCriteria.txtNameFirst.value;
var strWinName = document.frmStaffCriteria.hdnWinName.value;

	if (strNameLast !=  "" || strNameFirst !=  "") {
 		DeleteCookie("LName");
 		DeleteCookie("FName");
		DeleteCookie("WinName");
		document.frmStaffCriteria.submit();
	}
}

function validate() {
//**********************************************************************************************
// Function:	validate()																	   *
// Purpose:		To alert user that criteria should be entered to avoid a full database search  *
// Created By:	Nancy Mooney		09/25/2000												   *
// Updated By:																				   *
//**********************************************************************************************
	if (document.frmStaffCriteria.txtNameLast.value == "" &&
		document.frmStaffCriteria.txtNameFirst.value == "" &&
		document.frmStaffCriteria.txtEmail.value == "" &&
		document.frmStaffCriteria.txtEmpNo.value == "" &&
		document.frmStaffCriteria.txtWPhoneArea.value == "" &&
		document.frmStaffCriteria.txtWPhoneMid.value == "" &&
		document.frmStaffCriteria.txtWPhoneEnd.value == "" &&
		document.frmStaffCriteria.selDepartment.selectedIndex == 0 &&
		document.frmStaffCriteria.txtUserID.value == "" &&
		document.frmStaffCriteria.selStatus.selectedIndex == 0) {
		var bolConfirm = window.confirm("No search criteria have been entered. This search may take a long time...Continue?")
		return (bolConfirm);
	}
	if (isNaN(Number(document.frmStaffCriteria.txtWPhoneArea.value))) {
		alert('Please enter a numerical value for Work Phone');
		document.frmStaffCriteria.txtWPhoneArea.focus();
		document.frmStaffCriteria.txtWPhoneArea.select();
		return false;
	}
	if (isNaN(Number(document.frmStaffCriteria.txtWPhoneMid.value))) {
		alert('Please enter a numerical value for Work Phone');
		document.frmStaffCriteria.txtWPhoneMid.focus();
		document.frmStaffCriteria.txtWPhoneMid.select();
		return false;
	}
	if (isNaN(Number(document.frmStaffCriteria.txtWPhoneEnd.value))) {
		alert('Please enter a numerical value for Work Phone');
		document.frmStaffCriteria.txtWPhoneEnd.focus();
		document.frmStaffCriteria.txtWPhoneEnd.select();
		return false;
	}
	if (isNaN(Number(document.frmStaffCriteria.txtEmpNo.value))) {
		alert('Please enter a numerical value for Employee Number');
		document.frmStaffCriteria.txtEmpNo.focus();
		document.frmStaffCriteria.txtEmpNo.select();
		return false;
	}

	// search critiera have been entered so continue search
	return true ;
}

function btnNew_onClick(){
	if ((intAccessLevel & intConst_Access_Create) != intConst_Access_Create) {
		alert('Access denied. Please see your system administrator.');
		return false;
	}
	parent.document.location.href ="StaffDetail.asp?ContactID=NEW";
}

function btnClear_onClick(){
	with (document.frmStaffCriteria) {
		txtNameLast.value = "";
		txtNameFirst.value = "";
		txtEmail.value = "";
		txtEmpNo.value = "";
		selDepartment.selectedIndex = 0;
		selStatus.selectedIndex = 0;
		txtWPhoneArea.value = "";
		txtWPhoneMid.value = "";
		txtWPhoneEnd.value = "";
		txtUserID.value="";
		chkActiveOnly.checked = true;
	}
}

//End of SCRIPT hiding-->
</SCRIPT>
</HEAD>
<BODY language="javascript" onLoad="window_onLoad();">
<FORM id="frmStaffCriteria" name="frmStaffCriteria" method="post" action="StaffList.asp" target="fraResult" onSubmit="return validate();">
	<INPUT type="hidden" id="hdnWinName" name="hdnWinName" value="<%=strWinName%>">
<TABLE border="0" cols="4" width="100%">
<THEAD>
	<TR valign="middle"><TH align="left" colspan="4">Staff Search</TH></TR>
</THEAD>
<TBODY>
	<TR valign="middle">
		<TD align="right">Last Name</TD>
		<TD align="left"><INPUT id="txtNameLast" name="txtNameLast" type="text" maxlength="20" size="20" value=""></TD>
		<TD align="right">Employee Number</TD>
		<TD align="left"><INPUT id="txtEmpNo" name="txtEmpNo" type="text" maxlength="8" size="8" value=""></TD>
	</TR>
	<TR valign="middle">
		<TD align="right">First Name</TD>
		<TD align="left"><INPUT id="txtNameFirst" name="txtNameFirst" type="text" maxlength="20" size="20" value=""></TD>
		<TD align="right">Department</TD>
		<TD align="left"><SELECT id="selDepartment" name="selDepartment" style="width: 200px">
		<OPTION></OPTION>
		<%Do Until objDepartment.EOF%>
		<OPTION value="<%=objDepartment.Fields("DEPARTMENT_ID").Value%>"><%=objDepartment.Fields("DEPARTMENT_DESC").Value%></OPTION>
		<%objDepartment.MoveNext
		Loop
		objDepartment.Close
		Set objDepartment = Nothing%>
		</SELECT></TD>
	</TR>
	<TR>
		<TD align="right">UserID</TD>
		<TD align="left"><INPUT id="txtUserID" name="txtUserID" type="text" maxlength="8" size="8" value=""></TD>
		<TD align="right">Status</TD>
		<TD align="left">
		<SELECT id="selStatus" name="selStatus" style="width: 200px">
				<OPTION></OPTION>
				<%Do Until objStaffStatus.EOF%>
					<OPTION value="<%=objStaffStatus.Fields("STAFF_STATUS_LCODE").Value%>"><%=objStaffStatus.Fields("STAFF_STATUS_DESC").Value%></OPTION>
					<%objStaffStatus.MoveNext
				Loop
				objStaffStatus.Close
				Set objStaffStatus = Nothing%>
			</SELECT>
		</TD>
	</TR>
	<TR valign="middle">
		<TD align="right">Email Address</TD>
		<TD align="left"><INPUT id="txtEmail" name="txtEmail" type="text" maxlength="60" size="60" style="width=12cm" value=""></TD>
		<TD align="right">Active Only</TD>
		<TD align="left"><INPUT id="chkActiveOnly" name="chkActiveOnly" type="checkbox" checked></TD>
	</TR>
	<TR valign="middle">
		<TD align="right">Work Phone</TD>
		<TD align="left">
			(<INPUT id="txtWPhoneArea" name="txtWPhoneArea" type="text" size="3" maxlength="3">)
			<INPUT id="txtWPhoneMid" name="txtWPhoneMid" type="text" size="3" maxlength="3">-
			<INPUT id="txtWPhoneEnd" name="txtWPhoneEnd" type="text" size="4" maxlength="4"></TD>
	</TR>
</TBODY>
<TFOOT>
	<TR valign="middle">
		<TD align="right" colspan="4">
		<%If UCase(strWinName) <> UCase("Popup") Then%>
			<INPUT id="btnNew" name="btnNew" type="button" value="New" style="width: 2cm" language="javascript" onClick="btnNew_onClick();">&nbsp;
		<%End If%>
		<INPUT id="btnClear" name="btnClear" type="button" value="Clear" style="width: 2cm" language="javascript" onClick="btnClear_onClick();">&nbsp;
		<INPUT id="btnSearch" name="btnSearch" type="submit" value="Search" style="width: 2cm" language="javascript">&nbsp;</TD>
	</TR>
</TFOOT>
</TABLE>
</FORM>
</BODY>
<%
	Set objConn = Nothing
%>
</HTML>

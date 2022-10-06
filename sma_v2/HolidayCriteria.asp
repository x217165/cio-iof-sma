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
*					SLADesc
*
*
*******************************************************************************
-->
<%
Dim strWinName, strHolidayDescription, lIndex
Dim objCommand, objRS, strSQL, strWhereClause
Dim intAccessLevel

	intAccessLevel = CInt(CheckLogon(strConst_ServiceCatalogue))
	If (intAccessLevel And intConst_Access_ReadOnly) <> intConst_Access_ReadOnly Then
		DisplayError "BACK", "", 0, "ACCESS DENIED", "You don't have access to Holiday.  Please contact your system administrator"
	End If

	strWinName = Request("WinName")
	strHolidayDescription = Request("txtHolidayDescription")
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
var strHolidayDescription = "<%=strHolidayDescription%>";

//set section title
setPageTitle("SMA - Holiday");

function window_onLoad() {
//****************************************************************************************
//
//
//
//
//***************************************************************************************
	if (strHolidayDescription != "") {
		DeleteCookie("HolidayDescription");
		DeleteCookie("WinName");
		document.frmHolidaySearch.submit();
	}
}

/*function btnNew_onClick() {
//************************************************************************************************
// Function:	btnAddNew_onClick()
//
// Purpose:		To bring up a blank Service Level Detail page so that user can enter a new SLA.
//
// Created By:	Gilles Archer Oct 02 2000
//
// Updated By:
//************************************************************************************************
	if ((intAccessLevel & intConst_Access_Create) != intConst_Access_Create) {
		alert('You do not have permission to CREATE a Holiday.  Please contact your System Administrator.');
		return false;
	}
	parent.document.location.href ="HolidayDetail.asp?HolidayID=NEW";
}*/

function btnClear_onClick() {
	document.frmHolidaySearch.txtHolidayDescription.value = "";
	document.frmHolidaySearch.chkActiveOnly.checked = true;
}
//-->
</SCRIPT>
</HEAD>
<BODY language="javascript" onLoad="window_onLoad()">
<FORM id="frmHolidaySearch" name="frmHolidaySearch" method="post" action="HolidayList.asp" target="fraResult">
	<INPUT type="hidden" id="hdnWinName" name="hdnWinName" value="<%=strWinName%>">
	<INPUT type="hidden" id="hdnHolidayDescription" name="hdnHolidayDescription" value="<%=strHolidayDescription%>">
<TABLE cols="2" width="100%">
<THEAD>
	<TR><TD colspan="2" align="left">Holiday Search</TD></TR>
</THEAD>
<TBODY>
	<TR>
		<TD align="right">Holiday Description</TD>
		<TD align="left"><INPUT id="txtHolidayDescription" name="txtHolidayDescription" value="<%=strHolidayDescription%>" size="80" maxlength="80"></TD>
	</TR>
	<TR>
		<TD align="right">Active Only</TD>
		<TD align="left"><INPUT id="chkActiveOnly" name="chkActiveOnly" type="checkbox" checked></TD>
	</TR>
</TBODY>
<TFOOT>
	<TR>
		<TD colspan="2" align="right">
<!--
		<%If UCase(strWinName) <> UCase("Popup") Then%>
		<INPUT id="btnNew" name="btnNew" type="button" value="New" style="width: 2cm" language="javascript" onClick="btnNew_onClick()">&nbsp;
		<%End If%>
-->
		<INPUT id="btnClear" name="btnClear" type="button" value="Clear" style="width: 2cm" language="javascript" onClick="btnClear_onClick();">&nbsp;
		<INPUT id="btnSearch" name="btnSearch" type="submit" value="Search" style="width: 2cm" language="javascript">&nbsp;</TD>
	</TR>
</TFOOT>
</TABLE>
</FORM>
</BODY>
</HTML>

<%@ Language=VBScript %>
<%Option Explicit%>
<%Response.Buffer = True%>
<!--#include file = "smaConstants.inc"-->
<!--#include file = "smaProcs.inc"-->
<!--#include file = "databaseconnect.asp" -->
<!--
*********************************************************************************************
* Page name:	CountryCriteria.asp														*
* Purpose:		To display the Country Search												*
*				Chosen via searchFrame.asp													*
*																							*
* Created by:	Gilles Archer Oct 06 2000													*
*																							*
*********************************************************************************************
-->
<%
Dim strWinName, strCountryName, strCountryCode
Dim objRS, objCommand, strSQL
Dim intAccessLevel

	'Check user's rights
	intAccessLevel = CInt(CheckLogon(strConst_ServiceCatalogue))
	If (intAccessLevel And intConst_Access_ReadOnly) <> intConst_Access_ReadOnly Then
		DisplayError "BACK", "", 0, "ACCESS DENIED", "You don't have access to Countries. Please contact your system administrator"
	End If

	strWinName = Request("WinName")

	strCountryName = Request("CountryName")
	strCountryCode = Request("CountryCode")
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

//set section title
setPageTitle("SMA - Countries");

function window_onLoad() {
//****************************************************************************************
//
//
//
//
//***************************************************************************************
var strCountryCode = document.frmCountrySearch.txtCountryCode.value;
var strCountryName = document.frmCountrySearch.txtCountryName.value;

	if (strCountryName != "" || strCountryCode != "") {
		DeleteCookie("CountryName");
		DeleteCookie("CountryCode");
		DeleteCookie("WinName");
		document.frmCountrySearch.submit();
	}
}

function btnNew_onClick() {
	if ((intAccessLevel & intConst_Access_Create) != intConst_Access_Create) {
		alert('You do not have permission to CREATE a Country.  Please contact your System Administrator.');
		return false;
	}
	parent.document.location.href = "CountryDetail.asp?CountryID=NEW";
}

function btnClear_onClick() {
	document.frmCountrySearch.txtCountryCode.value = "";
	document.frmCountrySearch.txtCountryName.value = "";
}
//-->
</SCRIPT>
</HEAD>
<BODY language="javascript" onload="return window_onLoad()">
<FORM id="frmCountrySearch" name="frmCountrySearch" method="post" action="CountryList.asp" target="fraResult">
	<INPUT id="hdnWinName" name="hdnWinName" type="hidden" value="<%=strWinName%>">
	<INPUT id="hdnCountryCode" name="hdnCountryCode" type="hidden" value="<%=strCountryCode%>">
	<INPUT id="hdnCountryName" name="hdnCountryName" type="hidden" value="<%=strCountryName%>">
<TABLE border="0" cols="4" width="100%">
<THEAD><TR><TD colspan="4" align="left">Country Search</TD></TR></THEAD>
<TBODY>
	<TR>
		<TD align="right">Country Code</TD>
		<TD align="left"><INPUT id="txtCountryCode" name="txtCountryCode" type="text" value="<%=strCountryCode%>"></TD>
	</TR>
	<TR>
		<TD align="right">Country</TD>
		<TD align="left"><INPUT id="txtCountryName" name="txtCountryName" type="text" value="<%=strCountryName%>"></TD>
	</TR>
</TBODY>
<TFOOT>
	<TR>
		<TD colspan="4" align="right">
		<%If UCase(strWinName) <> UCase("Popup") Then%>
		<INPUT id="btnNew" name="btnNew" type="button" style="width: 2cm" value="New" language="javascript" onClick="btnNew_onClick();">&nbsp;
		<%End If%>
		<INPUT id="btnClear" name="btnClear" type="button" style="width: 2cm" value="Clear" language="javascript" onclick="return btnClear_onClick();">&nbsp;
		<INPUT id="btnSearch" name="btnSearch" type="submit" style="width: 2cm" value="Search">&nbsp;</TD>
	</TR>
</TFOOT>
</TABLE>
</FORM>
</BODY>
</HTML>

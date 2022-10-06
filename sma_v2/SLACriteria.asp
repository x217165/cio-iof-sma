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
Dim strWinName, strSLADescription, lIndex
Dim objCommand, objRS, strSQL, strWhereClause
Dim intAccessLevel

	intAccessLevel = CInt(CheckLogon(strConst_ServiceCatalogue))
	If (intAccessLevel And intConst_Access_ReadOnly) <> intConst_Access_ReadOnly Then
		DisplayError "BACK", "", 0, "ACCESS DENIED", "You don't have access to Service Level Agreement.  Please contact your system administrator"
	End If
 
	strWinName = Request("WinName")
	strSLADescription = Request("SLADesc")
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
var strSLADesc = "<%=strSLADescription%>";

//set section title
setPageTitle("SMA - Service Level Agreement");

function window_onLoad() {
//****************************************************************************************
//
//
//
//
//***************************************************************************************
	if (strSLADesc != "") {
		DeleteCookie("SLADesc");
		DeleteCookie("WinName");
		document.frmSLASearch.submit();
	}
}

function btnNew_onClick() {
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
		alert('You do not have permission to CREATE a Service Level Agreement.  Please contact your System Administrator.');
		return false;
	}		
	parent.document.location.href ="SLADetail.asp?ServiceLevelID=NEW";
}

function btnClear_onClick() {
	document.frmSLASearch.txtSLADescription.value = "";
	document.frmSLASearch.chkActiveOnly.checked = true;
}

//-->
</SCRIPT>
</HEAD>
<BODY language="javascript" onLoad="window_onLoad()">
<FORM id="frmSLASearch" name="frmSLASearch" method="post" action="SLAList.asp" target="fraResult">
	<INPUT type="hidden" id="hdnWinName" name="hdnWinName" value="<%=strWinName%>">
	<INPUT type="hidden" id="hdnSLADesc" name="hdnSLADesc" value="<%=strSLADescription%>">
<TABLE cols="4" width="100%">
<THEAD>
	<TR><TD colspan="4" align="left">Service Level Agreement Search</TD></TR>
</THEAD>
<TBODY>
	<TR>
		<TD align="right">SLA Description</TD>
		<TD align="left"><INPUT id="txtSLADescription" name="txtSLADescription" value="<%=strSLADescription%>" size="80" maxlength="80"></TD>    
	</TR> 
	<TR>
		<TD align="right">Active Only</TD>
		<TD align="left"><INPUT id="chkActiveOnly" name="chkActiveOnly" type="checkbox" checked></TD>
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

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
*					LOBDescription
*
* Updated By:	Al Hunt Sep 23 2004 - TQINOSS - Add strLANG (language preference) processing
*
*******************************************************************************
-->
<%
Dim strWinName, strLOBCode, strLOBDescription, strLANG
Dim objRS, strSQL, strWhereClause, strOrderBy
Dim intAccessLevel

	intAccessLevel = CInt(CheckLogon(strConst_ServiceCatalogue))
	If (intAccessLevel And intConst_Access_ReadOnly) <> intConst_Access_ReadOnly Then
		DisplayError "BACK", "", 0, "ACCESS DENIED", "You don't have access to Line of Business. Please contact your system administrator"
	End If

	strLOBCode = Request.Cookies("LOBCode")
	strLOBDescription = Request.Cookies("LOBDescription")
	strWinName	= Request.Cookies("WinName")

'TQ_INOSS
	strLANG = Request.Cookies("UserInformation")("language_preference")
	if (Len(strLANG) = 0) then strLANG = "EN"

	'Get the Line of Business : TQ_INOSS - shows English or French names only depending on the user perference
	strSQL = "SELECT LOB_ID, LOB_CODE, LOB_DESC " &_
			"FROM CRP.V_LOB " &_
			"WHERE lob_id NOT IN" &_
		        	"(SELECT lob_id " &_
		        	"FROM crp.v_lob " &_
		        	"WHERE language_preference_lcode = '" & strLANG & "' ) " &_
			"AND LANGUAGE_PREFERENCE_LCODE = 'EN'" &_
			"AND RECORD_STATUS_IND = 'A'" &_
			"UNION SELECT LOB_ID, LOB_CODE, LOB_DESC " &_
			"FROM crp.v_lob "

	strWhereClause = "WHERE language_preference_lcode = '" & strLANG & "'" &_
			 "AND RECORD_STATUS_IND = 'A' "

	strOrderBy = " ORDER BY LOB_DESC"
	strSQL = strSQL & strWhereClause & strOrderBy

	On Error Resume Next
	Set objRS = Server.CreateObject("ADODB.Recordset")
	Set objRS = objCommand.Execute
	objRS.Open strSQL, objConn, adOpenForwardOnly, adLockReadOnly, adCmdText
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

//set section title
setPageTitle("SMA - Line of Business");

function window_onLoad() {
//****************************************************************************************
//
//
//
//
//***************************************************************************************
var strLOBCode = document.frmLOBSearch.hdnLOBCode.value;

	if (strLOBCode != "") {
		DeleteCookie("LOBCode");
		DeleteCookie("WinName");
		document.frmLOBSearch.submit();
	}
}

function btnNew_onClick() {
//************************************************************************************************
// Function:	btnAddNew_onClick()
//
// Purpose:		To bring up a blank Line of Business Detail page so that user can enter a new LOB.
//
// Created By:	Gilles Archer Oct 02 2000
//
// Updated By:
//************************************************************************************************
	if ((intAccessLevel & intConst_Access_Create) != intConst_Access_Create) {
		alert('You do not have permission to CREATE a Line of Business.  Please contact your System Administrator.');
		return false;
	}
	parent.document.location.href ="LOBDetail.asp?BusinessID=NEW";
}

function btnClear_onClick() {
	document.frmLOBSearch.selLOB.selectedIndex = 0;
	document.frmLOBSearch.chkActiveOnly.checked = true;
}
//-->
</SCRIPT>
</HEAD>
<BODY language="javascript" onLoad="window_onLoad();">
<FORM id="frmLOBSearch" name="frmLOBSearch" method="post" action="LOBList.asp" target="fraResult">
	<INPUT type="hidden" id="hdnLOBCode" name="hdnLOBCode" value="<%=strLOBCode%>">
<TABLE cols="4" width=100%>
<THEAD>
	<TR><TD colspan="4" align="left">Line of Business Search</TD></TR>
</THEAD>
<TBODY>
	<TR>
		<TD align="right">Line of Business</TD>
		<TD align="left"><SELECT id="selLOB" name="selLOB" style="width: 350px">
			<OPTION></OPTION>
			<%Do While Not objRS.EOF
				If StrComp(strLOBCode, objRS.Fields("LOB_CODE").Value, 0) = 0 Then %>
				<OPTION value="<%=objRS.Fields("LOB_ID").Value%>" selected><%=objRS.Fields("LOB_CODE").Value & " - " & objRS.Fields("LOB_DESC").Value%></OPTION>
				<%Else%>
				<OPTION value="<%=objRS.Fields("LOB_ID").Value%>"><%=objRS.Fields("LOB_CODE").Value & " - " & objRS.Fields("LOB_DESC").Value%></OPTION>
				<%End If
				objRS.MoveNext
			Loop
			objRS.Close
			Set objRS = Nothing%>
			</SELECT>
		</TD>
	</TR>
</TBODY>
<TFOOT>
	<TR>
		<TD colspan="4" align="right">
		<%If UCase(strWinName) <> UCase("Popup") Then%>
		<INPUT id="btnNew" name="btnNew" type="button" value="New" style="width: 2cm" language="javascript" onClick="btnNew_onClick()">&nbsp;
		<%End If%>
		<INPUT id="btnClear" name="btnClear" type="button" value="Clear" style="width: 2cm" language="javascript" onClick="btnClear_onClick()">&nbsp;
		<INPUT id="btnSearch" name="btnSearch" type="submit" value="Search" style="width: 2cm" language="javascript">&nbsp;</TD>
	</TR>
</TFOOT>
</TABLE>
</FORM>
</BODY>
</HTML>

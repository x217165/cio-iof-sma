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
*					SCategoryDescription
*
*
*******************************************************************************
-->
<%
Dim strWinName, strBusinessID
Dim objCommand, objRS, strSQL, strWhereClause, strLANG
Dim intAccessLevel

	intAccessLevel = CInt(CheckLogon(strConst_ServiceCatalogue))
	If (intAccessLevel And intConst_Access_ReadOnly) <> intConst_Access_ReadOnly Then
		DisplayError "BACK", "", 0, "ACCESS DENIED", "You don't have access to Service Category. Please contact your system administrator"
	End If

	strWinName	= Request("WinName")
	strBusinessID = Request("BusinessID")

	'TQ_INOSS
	strLANG = Request.Cookies("UserInformation")("language_preference")
	if (Len(strLANG) = 0) then strLANG = "EN"

	'Get the Line of Business : TQ_INOSS
	strSQL = "SELECT LOB_ID, LOB_CODE, LOB_DESC " &_
			"FROM CRP.V_LOB " &_
			"WHERE lob_id NOT IN" &_
		        	"(SELECT lob_id " &_
		        	"FROM crp.v_lob " &_
		        	"WHERE language_preference_lcode = '" & strLANG & "' ) " &_
			"AND LANGUAGE_PREFERENCE_LCODE = 'EN'" &_
			"AND RECORD_STATUS_IND = 'A'" &_
			"UNION SELECT LOB_ID, LOB_CODE, LOB_DESC " &_
			"FROM crp.v_lob " &_
			"WHERE language_preference_lcode = '" & strLANG & "'" &_
			"AND RECORD_STATUS_IND = 'A' "

	If IsNumeric(strBusinessID) And Len(strBusinessID) > 0 Then
		strWhereClause = " AND LOB_ID = " & strBusinessID
	End If

	strSQL = strSQL & strWhereClause & " ORDER BY LOB_DESC ASC"

	'Response.Write strSQL
	'Response.End

	Set objRS = Server.CreateObject("ADODB.Recordset")
	Set objCommand = Server.CreateObject("ADODB.Command")
	objCommand.ActiveConnection = objConn
	objCommand.CommandText = strSQL
	objCommand.CommandType = adCmdText

	'Create Recordset object
	Set objRS = objCommand.Execute
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
var intWinName = "<%=strWinName%>";
var strBusinessID = "<%=strBusinessID%>";

//set section title
setPageTitle("SMA - Service Category");

function window_onLoad() {
//****************************************************************************************
//
//
//
//
//***************************************************************************************
	if (strBusinessID != "") {
		DeleteCookie("BusinessID");
		DeleteCookie("ServiceCategoryID");
		DeleteCookie("ServiceType");
		DeleteCookie("STypeDesc");
		DeleteCookie("WinName");
		document.frmSCategorySearch.submit();
	}
}

function btnNew_onClick() {
//************************************************************************************************
// Function:	btnAddNew_onClick()
//
// Purpose:		To bring up a blank Service Category Detail page so that user can enter a new SC.
//
// Created By:	Gilles Archer Oct 02 2000
//
// Updated By:
//************************************************************************************************
	if ((intAccessLevel & intConst_Access_Create) != intConst_Access_Create) {
		alert('You do not have permission to CREATE a Service Category.  Please contact your System Administrator.');
		return false;
	}
	parent.document.location.href ="SCategoryDetail.asp?ServiceCategoryID=NEW";
}

function btnClear_onClick() {
	document.frmSCategorySearch.selLOB.selectedIndex = 0;
	document.frmSCategorySearch.txtSCategoryDescription.value = "";
	document.frmSCategorySearch.chkActiveOnly.checked = true;
}

//-->
</SCRIPT>
</HEAD>
<BODY language="javascript" onLoad="window_onLoad()">
<FORM id="frmSCategorySearch" name="frmSCategorySearch" method="post" action="SCategoryList.asp" target="fraResult">
	<INPUT type="hidden" id="hdnBusinessID" name="hdnBusinessID" value="<%=strBusinessID%>">
	<INPUT type="hidden" id="hdnWinName" name="hdnWinName" value="<%=strWinName%>">
<TABLE cols="4" width=100%>
<THEAD>
	<TR><TD colspan="4" align="left">Service Category Search</TD></TR>
</THEAD>
<TBODY>
	<TR>
		<TD align="right">Line of Business</TD>
		<TD align="left"><SELECT id="selLOB" name="selLOB" style="width: 350px">
			<OPTION></OPTION>
			<%Do While Not objRS.EOF
				If StrComp(strBusinessID, CStr(objRS.Fields("LOB_ID").Value), 0) = 0 Then %>
				<OPTION value="<%=objRS.Fields("LOB_ID").Value%>" selected><%=objRS.Fields("LOB_CODE").Value & " - " & objRS.Fields("LOB_DESC").Value%></OPTION>
				<%Else%>
				<OPTION value="<%=objRS.Fields("LOB_ID").Value%>"><%=objRS.Fields("LOB_CODE").Value & " - " & objRS.Fields("LOB_DESC").Value%></OPTION>
				<%End If
				objRS.MoveNext
			Loop%>
			</SELECT>
		</TD>
	</TR>
	<TR>
		<TD align="right">Service Category</TD>
		<TD align="left"><INPUT id="txtSCategoryDescription" name="txtSCategoryDescription" value="" size="80" maxlength="80"></TD>
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
		<INPUT id="btnNew" name="btnNew" type="button" value="New" style="width: 2cm" language="javascript" onClick="btnNew_onClick()"> &nbsp;
		<%End If%>
		<INPUT id="btnClear" name="btnClear" type="button" value="Clear" style="width: 2cm" language="javascript" onClick="btnClear_onClick()">&nbsp;
		<INPUT id="btnSearch" name="btnSearch" type="submit" value="Search" style="width: 2cm">&nbsp;</TD>
	</TR>
</TFOOT>
</TABLE>
</FORM>
<%
	'Clean ADO Objects
	objRS.Close
	Set objRS = Nothing
	Set objCommand = Nothing
	objConn.Close
	Set objConn = Nothing
%>
</BODY>
</HTML>

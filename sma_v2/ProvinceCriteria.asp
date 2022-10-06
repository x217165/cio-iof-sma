<%@ Language=VBScript %>
<%Option Explicit%>
<%Response.Buffer = True%>
<!--#include file = "smaConstants.inc"-->
<!--#include file = "smaProcs.inc"-->
<!--#include file = "databaseconnect.asp" -->
<!--
*********************************************************************************************
* Page name:	ProvinceCriteria.asp														*
* Purpose:		To display the Province Search												*
*				Chosen via searchFrame.asp													*
*																							*
* Created by:	Gilles Archer Oct 06 2000													*
*																							*
*********************************************************************************************
-->
<%
Dim strWinName, strProvinceName, strProvinceCode, strCountryName, strCountryCode
Dim objCountries, strSQL
Dim intAccessLevel

	'Check user's rights
	intAccessLevel = CInt(CheckLogon(strConst_ServiceCatalogue))
	If (intAccessLevel And intConst_Access_ReadOnly) <> intConst_Access_ReadOnly Then
		DisplayError "BACK", "", 0, "ACCESS DENIED", "You don't have access to Provinces. Please contact your system administrator"
	End If

	strWinName = Request("WinName")

	strProvinceName = Request("ProvinceName")
	strProvinceCode = Request("ProvinceCode")
	strCountryCode = Request("CountryCode")

	strSQL = "SELECT UPPER(C.COUNTRY_LCODE) AS COUNTRY_LCODE, " &_
			"C.COUNTRY_DESC " &_
			"FROM CRP.LCODE_COUNTRY C " &_
			"WHERE C.RECORD_STATUS_IND = 'A' " &_
			"ORDER BY C.COUNTRY_LCODE ASC"

	'Create Recordset object
	On Error Resume Next
	Set objCountries = Server.CreateObject("ADODB.Recordset")
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

//set section title
setPageTitle("SMA - Provinces");

function window_onLoad() {
//****************************************************************************************
//
//
//
//
//***************************************************************************************
var strProvince = document.frmProvinceSearch.txtProvince.value;
var strProvinceCode = document.frmProvinceSearch.txtProvinceCode.value;
var strCountryCode = document.frmProvinceSearch.selCountry.value;

	if (strProvince != "" || strProvinceCode !="" || strCountryCode != "") {
		DeleteCookie("ProvinceName");
		DeleteCookie("ProvinceCode");
		DeleteCookie("CountryCode");
		DeleteCookie("WinName");
		document.frmProvinceSearch.submit();
	}
}

function btnNew_onClick() {
	if ((intAccessLevel & intConst_Access_Create) != intConst_Access_Create) {
		alert('You do not have permission to CREATE a Province.  Please contact your System Administrator.');
		return false;
	}
	parent.document.location.href = "ProvinceDetail.asp?ProvinceID=NEW";
}

function btnClear_onClick() {
	document.frmProvinceSearch.txtProvince.value = "";
	document.frmProvinceSearch.txtProvinceCode.value = "";
	document.frmProvinceSearch.selCountry.selectedIndex = 0;
}
//-->
</SCRIPT>
</HEAD>
<BODY language="javascript" onload="return window_onLoad()">
<FORM id="frmProvinceSearch" name="frmProvinceSearch" method="post" action="ProvinceList.asp" target="fraResult">
	<INPUT id="hdnWinName" name="hdnWinName" type="hidden" value="<%=strWinName%>">
	<INPUT id="hdnProvinceName" name="hdnProvinceName" type="hidden" value="<%=strProvinceName%>">
	<INPUT id="hdnProvinceCode" name="hdnProvinceCode" type="hidden" value="<%=strProvinceCode%>">
<TABLE border="0" cols="4" width="100%">
<THEAD><TR><TD colspan="4" align="left">Province Search</TD></TR></THEAD>
<TBODY>
	<TR>
		<TD align="right">Province</TD>
		<TD align="left"><INPUT id="txtProvince" name="txtProvince" type="text"value="<%=strProvinceName%>"></TD>
		<TD align="right">Country</TD>
		<TD align="left">
		<SELECT id="selCountry" name="selCountry" style="width: 200">
			<OPTION></OPTION>
			<%Do While Not objCountries.EOF
				If Len(strProvinceCode) <> 0 Then
					If StrComp(objCountries.Fields("COUNTRY_LCODE").Value, strProvinceCode, 0) = 0 Then
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
	<TR>
		<TD align="right">Province Code</TD>
		<TD align="left"><INPUT id="txtProvinceCode" name="txtProvinceCode" type="text"value="<%=strProvinceCode%>"></TD>
	</TR>
</TBODY>
<TFOOT>
	<TR>
		<TD colspan="4" align="right">
		<%If UCase(strWinName) <> UCase("Popup") Then%>
		<INPUT id="btnNew" name="btnNew" type="button" value="New" style="width: 2cm" language="javascript" onClick="btnNew_onClick();">&nbsp;
		<%End If%>
		<INPUT id="btnClear" name="btnClear" type="button" style="width: 2cm" value="Clear" language="javascript" onclick="return btnClear_onClick();">&nbsp;
		<INPUT id="btnSearch" name="btnSearch" type="submit" style="width: 2cm" value="Search">&nbsp;</TD>
	</TR>
</TFOOT>
</TABLE>
</FORM>
</BODY>
</HTML>

<%@ Language=VBScript %>
<%Option Explicit%>
<%Response.Buffer = True%>
<!--#include file = "smaConstants.inc"-->
<!--#include file = "smaProcs.inc"-->
<!--#include file = "databaseconnect.asp" -->
<!--
*********************************************************************************************
* Page name:	CityCriteria.asp															*
* Purpose:		To display the City Search													*
*				Chosen via searchFrame.asp													*
*																							*
* Created by:	Gilles Archer Oct 06 2000													*
*																							*
*********************************************************************************************
-->
<%
Dim strWinName, strCityName, strProvinceCode, strCountryCode, strServiceRegion
Dim objProvinces, objCountries, strSQL
Dim intAccessLevel

'UB create lists
dim sql
dim rsSRrp
'UB: end

	'Check user's rights
	intAccessLevel = CInt(CheckLogon(strConst_ServiceCatalogue))
	If (intAccessLevel And intConst_Access_ReadOnly) <> intConst_Access_ReadOnly Then
		DisplayError "BACK", "", 0, "ACCESS DENIED", "You don't have access to Municipalities. Please contact your system administrator"
	End If
 
	strWinName = Request("WinName")
	strCityName = Request("CityName")
	strProvinceCode = Request("ProvinceCode")
	strCountryCode = Request("CountryCode")
	strServiceRegion = Request("ServiceRegCode")

	strSQL = "SELECT UPPER(PS.PROVINCE_STATE_LCODE) AS PROVINCE_STATE_LCODE, " &_
			"UPPER(PS.COUNTRY_LCODE) AS COUNTRY_LCODE, " &_
			"PS.PROVINCE_STATE_NAME " &_
			"FROM CRP.LCODE_PROVINCE_STATE PS " &_
			"WHERE PS.RECORD_STATUS_IND = 'A' " &_
			"ORDER BY PS.COUNTRY_LCODE ASC, PS.PROVINCE_STATE_LCODE ASC"

	'Create Recordset object  
	On Error Resume Next
	Set objProvinces = Server.CreateObject("ADODB.Recordset")
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
	On Error Resume Next
	Set objCountries = Server.CreateObject("ADODB.Recordset")
	objCountries.Open strSQL, objConn, adOpenForwardOnly, adLockReadOnly, adCmdText
	If objConn.Errors.Count <> 0 Then
		DisplayError "BACK", "", objConn.Errors(0).NativeError, "ERROR LOADING DATA (Countries)", objConn.Errors(0).Description
		objConn.Errors.Clear
	End If

'UB: get the Customer Managed Service Region
	sql = "select CUST_MGD_SRVC_RGN_NAME from CRP.LCODE_CUST_MGD_SRVC_RGN where RECORD_STATUS_IND='A' ORDER BY CUST_MGD_SRVC_RGN_LCODE"
	set rsSRrp = Server.CreateObject("ADODB.Recordset")
	rsSRrp.CursorLocation = adUseClient
	rsSRrp.Open sql, objConn
	if err then
		DisplayError "BACK", "", err.Number, "CANNOT CREATE RECORDSET", err.Description
	end if
	'release the active connection, keep the recordset open
	set rsSRrp.ActiveConnection = nothing 
'UB: end
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
setPageTitle("SMA - Municipalities");

function window_onLoad() {
//****************************************************************************************
//
//
//
//
//***************************************************************************************
var strCityName = document.frmCitySearch.txtCity.value;
var strProvinceCode = document.frmCitySearch.selProvince.value;
var strCountryCode = document.frmCitySearch.selCountry.value; 
var strServiceRegion = document.frmCitySearch.selServiceRegion.value;
	
	if (strCityName != "" || strProvinceCode !="" || strCountryCode != "" || strServiceRegion != "") {	
		DeleteCookie("CityName");
		DeleteCookie("ProvinceCode");
		DeleteCookie("CountryCode");
		DeleteCookie("ServiceRegCode");
		DeleteCookie("WinName");
		document.frmCitySearch.submit();
	} 
}

function btnNew_onClick() {
	if ((intAccessLevel & intConst_Access_Create) != intConst_Access_Create) {
		alert('You do not have permission to CREATE a Municipality.  Please contact your System Administrator.');
		return false;
	}
	parent.document.location.href = "CityDetail.asp?CityID=NEW";
}

function btnClear_onClick() {
	with (document.frmCitySearch) {
		txtCity.value = "";
		txtClliCode.value = "";  
		selProvince.selectedIndex = 0;
		selCountry.selectedIndex = 0;
		selServiceRegion.selectedIndex = 0;	
	}
}

function fct_onSubmit() {
  if (document.frmCitySearch.txtCity.value == "" && 
		document.frmCitySearch.txtClliCode.value == "" && 
		document.frmCitySearch.selProvince.selectedIndex == 0 &&
		document.frmCitySearch.selCountry.selectedIndex == 0 &&
		document.frmCitySearch.selServiceRegion.selectedIndex == 0 ) {
		
		bolConfirm = window.confirm("No Search Criteria have been entered.\n\nThis search may take a long time... Continue?");
		return (bolConfirm);			 
  }
   // search critiera has been entered so continue search
   return true;
}
//-->
</SCRIPT>
</HEAD>
<BODY language="javascript" onload="return window_onLoad();">
<FORM id="frmCitySearch" name="frmCitySearch" method="post" action="CityList.asp" target="fraResult" onsubmit="return fct_onSubmit();">
	<INPUT id="hdnWinName" name="hdnWinName" type="hidden" value="<%=strWinName%>">
	<INPUT id="hdnCityName" name="hdnCityName" type="hidden" value="<%=strCityName%>">
	<INPUT id="hdnProvinceCode" name="hdnProvinceCode" type="hidden" value="<%=strProvinceCode%>">
	<INPUT id="hdnCountryCode" name="hdnCountryCode" type="hidden" value="<%=strCountryCode%>">
	<INPUT id="hdnServiceRegion" name="hdnServiceRegion" type="hidden" value="<%=strServiceRegion%>">
<TABLE border="0" cols="4" width="100%">
<THEAD><TR><TD colspan="4" align="left">Municipality Search</TD></TR></THEAD>
<TBODY>  
<TR>
	<TD align="right">City / Municipality Name</TD>
	<TD align="left"><INPUT type="text" id="txtCity" name="txtCity" maxlength="50" size="50" value="<%=strCityName%>"></TD>
	<TD align="right">Province / State</TD>
	<TD align="left">
	<SELECT id="selProvince" name="selProvince" style="width: 200">
		<OPTION></OPTION>
		<%Do While Not objProvinces.EOF
			If Len(strCityName) <> 0 Then
				If StrComp(objProvinces.Fields("PROVINCE_STATE_LCODE").Value, strProvinceCode, 0) = 0 Then
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
<TR>
	<TD align="right">City Short Name</TD>
	<TD align="left"><INPUT type="text" id="txtClliCode" name="txtClliCode" maxlength="4" size="5"></TD>
	<TD align="right">Country</TD>
	<TD align="left">
	<SELECT id="selCountry" name="selCountry" style="width: 200">
		<OPTION></OPTION>
		<%Do While Not objCountries.EOF
			If Len(strCityName) <> 0 Then
				If StrComp(objCountries.Fields("COUNTRY_LCODE").Value, strCountryCode, 0) = 0 Then
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
	<TD align="right">Customer Managed Service Region</TD>
        <td align="left">
		<SELECT id=selServiceRegion name=selServiceRegion tabindex=12>
			<option selected value="ALL"> </OPTION>
			<%
			while not rsSRrp.EOF		 
				Response.Write "<option>" & routineHtmlString(rsSRrp(0)) & "</option>" & vbCrLf
				rsSRrp.MoveNext
			wend
			rsSRrp.Close
			%>
	</SELECT></td>
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

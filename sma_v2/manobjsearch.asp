﻿<%@ LANGUAGE=VBSCRIPT %>
<%
option explicit
on error resume next
%>
<!-- #include file=smaConstants.inc -->
<!-- #include file=smaProcs.inc -->
<!-- #include file=databaseconnect.asp -->
<!--
*************************************************************************************
* File Name:	manobjsearch.asp
*
* Purpose:
*
* In Param:		This page reads following cookies
*				CustomerName
*				AssetID
*				WinName
*				ServLocName
*
* Out Param:
*
* Created By:
**************************************************************************************
		 Date		Author			Changes/enhancements made

       25-Jan-02   Adam Haydey  Added Customer Service City, Customer Service Address, TAC Assset Code and Non-Correlated Only search fields.
				                  TAC Asset Code was added to the search results.
       14-Mar-02	 DTy		Add Port Name and LAN IP as search fields.
       10-Aug-04   ACheung		Add LYNX repair priority
       29-Jul-15   PSmith  Set Cookies in validation so the back key works
       05-Oct-15   PSmith  Only sumbit() for pop-up windows
       03-Feb-16   PSmith  Don't pre-populate search criteria
                           and moved managed object above customer.
**************************************************************************************
-->
<%
'check users access rights
dim intAccessLevel
intAccessLevel = CInt(CheckLogon(strConst_ManagedObjects))
if intAccessLevel < intConst_Access_ReadOnly then
	DisplayError "BACK", "", 0, "ACCESS DENIED", "You don't have access to managed object. Please contact your system administrator"
end if

'get cookies
dim strCustomerName,strManObjName,strAssetId, strWinName, strServLocName, strServLocAdd , strServLocCity
strCustomerName = Request.Cookies("CustomerName")
strManObjName = Request.Cookies("ManObjName")
strAssetId = Request.Cookies("AssetID")
strWinName	= Request.Cookies("WinName")
strServLocName = Request.Cookies("ServLocName")

'create lists
dim sql
dim rsNetworkElementType, rsRegion, rsLYNXrp

'get the list of network element types
sql = "select NETWORK_ELEMENT_TYPE_CODE from CRP.NETWORK_ELEMENT_TYPE where RECORD_STATUS_IND='A' ORDER BY NETWORK_ELEMENT_TYPE_CODE"
set rsNetworkElementType = Server.CreateObject("ADODB.Recordset")
rsNetworkElementType.CursorLocation = adUseClient
rsNetworkElementType.Open sql, objConn
if err then
	DisplayError "BACK", "", err.Number, "CANNOT CREATE RECORDSET", err.Description
end if
'release the active connection, keep the recordset open
set rsNetworkElementType.ActiveConnection = nothing

'get the regions
sql = "select NOC_REGION_LCODE, NOC_REGION_DESC from CRP.LCODE_NOC_REGION where RECORD_STATUS_IND='A'"
set rsRegion = Server.CreateObject("ADODB.Recordset")
rsRegion.CursorLocation = adUseClient
rsRegion.Open sql, objConn
if err then
	DisplayError "BACK", "", err.Number, "CANNOT CREATE RECORDSET", err.Description
end if
'release the active connection, keep the recordset open
set rsRegion.ActiveConnection = nothing

'get the LYNX repair priority
sql = "select LYNX_DEF_SEV_DESC from CRP.LCODE_LYNX_DEF_SEV where RECORD_STATUS_IND='A' ORDER BY LYNX_DEF_SEV_LCODE"
set rsLYNXrp = Server.CreateObject("ADODB.Recordset")
rsLYNXrp.CursorLocation = adUseClient
rsLYNXrp.Open sql, objConn
if err then
	DisplayError "BACK", "", err.Number, "CANNOT CREATE RECORDSET", err.Description
end if
'release the active connection, keep the recordset open
set rsLYNXrp.ActiveConnection = nothing


'get the support group recordset
dim rsSG
sql = "SELECT REMEDY_SUPPORT_GROUP_ID, GROUP_NAME FROM CRP.V_REMEDY_SUPPORT_GROUP ORDER BY GROUP_NAME"
set rsSG=server.CreateObject("ADODB.Recordset")
rsSG.CursorLocation = adUseClient
rsSG.Open sql, objConn
if err then
	DisplayError "BACK", "", err.Number, "CANNOT CREATE RECORDSET", err.Description
end if

if rsSG.EOF then
	DisplayError "BACK", "", 999, "CANNOT CREATE SUPPORT GROUP LIST", "EOF condition occured in rsSG recordset."
end if

set rsSG.ActiveConnection = nothing

objConn.close
set objConn = nothing
%>

<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft FrontPage 12.0">
<META HTTP-EQUIV="Pragma" CONTENT="no-cache">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<TITLE>SMA - Managed Objects Search</TITLE>
<LINK href="stylesheets/styles.css" rel="stylesheet" type="text/css">

<script type="text/javascript" SRC="GeneralJavaFunctions.js"></script>
<script type="text/javascript" SRC="AccessLevels.js"></script>
<script>
var intAccessLevel=<%=intAccessLevel%>;

//set section title
setPageTitle("SMA - Managed Objects");

function window_onload() {

	/************************************************************************************************
	*  Function:	window_onload																	*
	*																								*
	*  Purpose:		To submit the form automatically when values have been received from a cookie   *
	*				and have been stored in hidden form controls.									*
	*																								*
	*  Created By:	Nancy Mooney 09/03/2000															*
	*																								*
	*************************************************************************************************/

	 	var strWinName;
	 	strWinName = document.frmManObjSearch.hdnWinName.value ;
	 	if (strWinName !=  "" ){
 			DeleteCookie("WinName") ;
		}

		var strCustomerName,strManObjName,strAssetId;

		strCustomerName = document.frmManObjSearch.txtCustomer.value;
		strManObjName = document.frmManObjSearch.txtManObjName.value;
		strAssetId = document.frmManObjSearch.txtAssetId.value ;
		strServLocName = document.frmManObjSearch.txtServLoc.value ;

	 	DeleteCookie("CustomerName");
	 	DeleteCookie("ManObjName");
	 	DeleteCookie("AssetID");
	 	DeleteCookie("ServLocName");

	 	if ( strWinName == "Popup" && ((strCustomerName != "") || (strManObjName != "") || (strAssetId != "") || strServLocName != "" )){
		  SetCookie("CustomerName", document.frmManObjSearch.txtCustomer.value);
		  SetCookie("ManObjName", document.frmManObjSearch.txtManObjName.value);
		  SetCookie("AssetID", document.frmManObjSearch.txtAssetId.value);
		  SetCookie("ServLocName", document.frmManObjSearch.txtServLoc.value);
		  thinking(parent.fraResult);
 			document.frmManObjSearch.submit();
 		}
	}

function confirm_search(theForm)
{
var bolConfirm;

	if ((isWhitespace(theForm.txtAssetId.value) &&isWhitespace(theForm.txtCustomer.value)&& (theForm.selRegion.selectedIndex ==0) &&
		isWhitespace(theForm.txtManObjName.value) && (theForm.selSupportGroup.selectedIndex ==0) && isWhitespace(theForm.txtIPAddress.value) &&
		(theForm.selManObjType.selectedIndex ==0) && isWhitespace(theForm.txtOBDialup.value) && isWhitespace(theForm.txtServLoc.value) &&
		isWhitespace(theForm.txtServLocCity.value) && isWhitespace(theForm.txtServLocAdd.value) && isWhitespace(theForm.txtBarCode.value) &&
        	isWhitespace(theForm.txtManObjPort.value) && isWhitespace(theForm.txtManObjLANIP.value) && (theForm.selRepairPriority.selectedIndex ==0) &&
		(theForm.chkNullIP.checked == false) && (theForm.chkNullOBDialup.checked == false) && (theForm.chkNonCorrOnly.checked == false)))
	{

		bolConfirm = window.confirm("No Search Criteria have been entered. This search may take a long time..Continue?");
		if (!bolConfirm){
			return false;
		}
	}

   thinking(parent.fraResult);

   return true;
}

function btnClear_onclick() {
	with(document.frmManObjSearch){
		txtAssetId.value = "";
		txtCustomer.value = "";
		selRegion.selectedIndex = 0;
		txtManObjName.value = "";
		selSupportGroup.selectedIndex = 0;
		txtIPAddress.value = "";
		selManObjType.selectedIndex = 0;
		txtOBDialup.value = "";
		chkActiveOnly.checked = true;
		txtServLoc.value = "";
		chkNullIP.checked = false;
		chkNullOBDialup.checked = false;
		txtServLocCity.value ="";
		txtManObjPort.value = "";
		txtManObjLANIP.value = "";
		txtServLocAdd.value ="";
		txtBarCode.value="";
		chkNonCorrOnly.checked=false;
		selRepairPriority.selectedIndex = 0 ;
	}


}

function btnNew_onclick(){
	if ((intAccessLevel & intConst_Access_Create) != intConst_Access_Create){
		alert('Access denied. Please contact your system administrator.');
		return;
	}
	parent.document.location='manobjdet.asp?ne_id=';
}

</script>

</HEAD>
<BODY LANGUAGE=javascript onload="return window_onload();">
<form name=frmManObjSearch action="manobjlist.asp" method=POST target="fraResult" onSubmit="return confirm_search(this)">
<input type="hidden" name="hdnWinName" value="<%=strWinName%>">
<input type="hidden" name="txtAssetId" value="<%=strAssetId%>">
<table border="0" width="100%">
<thead align=left>
  <tr>
    <td align=Left colSpan=4>Managed Objects Search</td>
  </tr>
</thead>
<tbody>
  <tr>
    <td align=right>Managed Object Name/Alias</td>
    <td><INPUT name=txtManObjName tabindex=2 size=25 maxlength=30 value="<%=strManObjName%>"></td>
    <td align=right>Support Group</td>
	<td><SELECT name=selSupportGroup tabindex=11>
		<OPTION></OPTION>
		<%
		while not rsSG.EOF
			Response.Write "<OPTION"
			Response.Write " VALUE="& rsSG("REMEDY_SUPPORT_GROUP_ID") &">" & routineHtmlString(rsSG("GROUP_NAME")) & "</OPTION>" &vbCrLf
			rsSG.MoveNext
		wend
		rsSG.Close
		%>
		</SELECT>
	</td>
  </tr>
  <tr>
    <td align=right width="20%">Customer</td>
    <td width="30%"><INPUT name=txtCustomer tabindex=1 size=25 maxlength=50 value="<%=strCustomerName%>"></td>
    <td align=right width="20%">Region</td>
    <td width="30%"><SELECT name=selRegion tabindex=10>
		<option selected value="ALL"> </OPTION>
		<%
		while not rsRegion.EOF
			Response.Write "<option value='" & rsRegion(0) & "'>" & routineHtmlString(rsRegion(1)) &"</option>" & vbCrLf
			rsRegion.MoveNext
		wend
		rsRegion.Close
		%>
0	</SELECT></td>
  </tr>
  <tr>
	<td align=right>IP Address</td>
	<td><INPUT name=txtIPAddress tabindex=3 size=16 maxlength=30 value=""> / null? <input name=chkNullIP tabindex=4 type=checkbox></td>
    <td align=right>Managed Object Type</td>
    <td >
		<SELECT name=selManObjType tabindex=12>
			<option selected value="ALL"> </OPTION>
			<%
			while not rsNetworkElementType.EOF
				Response.Write "<option>" & routineHtmlString(rsNetworkElementType(0)) & "</option>" & vbCrLf
				rsNetworkElementType.MoveNext
			wend
			rsNetworkElementType.Close
			%>
		</SELECT></td>
  </tr>
  <tr>
	<td align=right>Out of Band Dialup</td>
	<td><INPUT name=txtOBDialup tabindex=5 size=16 maxlength=30 value=""> / null? <input type=checkbox name=chkNullOBDialup tabindex=6></td>
      <td align=right>Repair Priority</td>
    <td >
		<SELECT id=selRepairPriority name=selRepairPriority tabindex=12>
			<option selected value="ALL"> </OPTION>
			<%
			while not rsLYNXrp.EOF
				Response.Write "<option>" & routineHtmlString(rsLYNXrp(0)) & "</option>" & vbCrLf
				rsLYNXrp.MoveNext
			wend
			rsLYNXrp.Close
			%>
		</SELECT></td>
  </TR>
  <TR>
    <td align=right>Service Location</td>
    <td><INPUT name=txtServLoc tabindex=7 size=25 maxlength=50 value="<%=strServLocName%>"></td>
    <TD align=right nowrap width=15%>TAC Asset Code (Barcode) </TD>
	<TD ALIGN=LEFT ><INPUT id=txtBarCode name=txtBarCode tabindex=13 style="HEIGHT: 23px; WIDTH: 200px" value="<%=strBarcode%>" ></TD>
  </TR>
  <TR>
	<TD align=right nowrap width=15%>Service Location Address</TD>
	<TD ALIGN=LEFT width=20%><INPUT id=txtServLocAdd name=txtServLocAdd tabindex=8 style="HEIGHT: 23px; WIDTH: 200px" value="<%=strServLocAdd%>" ></TD>
    <td align=right>Port Name</td>
    <td><INPUT name=txtManObjPort tabindex=15 size=25 maxlength=30></td>
  </TR>
  <TR>
    <TD align=right nowrap width=15%>Service Location City</TD>
	<TD ALIGN=LEFT><INPUT id=txtServLocCity name=txtServLocCity tabindex=9 style="HEIGHT: 23px; WIDTH: 200px" value="<%=strServLocCity%>" ></TD>
    <td align=right>LAN IP</td>
    <td><INPUT name=txtManObjLANIP tabindex=16 size=50 maxlength=50></td>
  </tr>
  <tr>
	<td></td>
	<td></td>
	<td align=right>Active Only</td>
	<td align=left ><INPUT name=chkActiveOnly tabindex=17 type=checkbox value=yes checked>   <nowrap width=15%>    Non-correlated Only <INPUT name=chkNonCorrOnly tabindex=14 type=checkbox> <TD>
  </tr>

  <tr>
	<td></td>
	<td></td>
	<td align=right colspan=2>
	<% if strWinName <> "Popup" then %>
		<INPUT name=btnNew type=button value=New tabindex=18 style= "width: 2cm" onclick="btnNew_onclick();">&nbsp;&nbsp;
	<% end if %>
		<INPUT name=btnClear type=button value=Clear tabindex=19 style= "width: 2cm" onclick="return btnClear_onclick();">&nbsp;&nbsp;
		<INPUT name=btnSubmit type=submit tabindex=20 value=Search style= "width: 2cm" >&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
	</td>
  </tr>
<tbody>
</table>
</form>
</BODY>
</HTML>

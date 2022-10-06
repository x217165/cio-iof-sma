<%@ LANGUAGE=VBSCRIPT %>
<% 
option explicit
on error resume next 
%>
<!-- #include file=smaConstants.inc -->
<!-- #include file=smaProcs.inc -->
<!-- #include file=databaseconnect.asp -->
<!--
*************************************************************************************
* File Name:	RSAST3search.asp
*
* Purpose:	Input criteria for searching gateway circuits
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
* Edited by:    Adam Haydey Jan 25, 2001
*               Added Customer Service City, Customer Service Address, TAC Assset Code and Non-Correlated Only search fields.
*				The service Location search defaults to the Asset Address chosen (Street, City and Province fields)
*				TAC Asset Code was added to the search results.
**************************************************************************************

        Date		Author			Changes/enhancements made
        -----		------		------------------------------------------------------
       10-04-01	     DTy		Change field names and variables.

                                  Gateway IP Address to GW Router T1 Serial IP Address
                                  txtIPGate          to txtGWRT1SIIPAddr
                                  strIPGateway       to strGWRT1SIIPAddr
                                  
                                  Gateway DLCI POS   to Gateway DLCI (X25)
                                  txtGWPOS           to txtDLCIX25

                                  Gateway DLCI IP    to Gateway DLCI (IP)
                                  txtGWDLCI          to txtDLCIIP

                                  WAN IP Address     to WAN IP Port Address
                                  txtWANIP           to txtWANIPAddr

                                  PNG IP Address     to LAN IP Port Address
                                  txtPNGIP           to txtLANIPAddr

                                  Site Name/Address  to Site Address
                                  txtSiteAdd         to txtSiteAddr 
                                  strSiteAdd         to strSiteAddr
                                  
                                  strServLocAdd      to strServLocAddr

                                  Move Gateway DLCI (X25) above Gateway DLCI (IP)
                                  
                                  Remove WAN IP DLCI & txtWANDLCI)
                                  Remove POS IP DLCI & txtPOSDLCI
                                  
                                  Increase Tail Circuit Number from 15 to 19 chars
                                  Increase Order Number from 15 to 16 chars
                                  
                                  Re-arrange field positions.

                                  Define missing variables.
                                  Correct menu sub-title to SMA - POS PLUS
**************************************************************************************
-->
<%
'check users access rights
dim intAccessLevel
intAccessLevel = CInt(CheckLogon(strConst_RSAS)) 
if intAccessLevel < intConst_Access_ReadOnly then
	DisplayError "BACK", "", 0, "ACCESS DENIED", "You don't have access to POS PLUS maintenance screens. Please contact your system administrator"
end if

'get cookies
dim strCustomerName, strWinName, strGWRT1SIIPAddr, strDLCIX25, strDLCIIP
dim strWANIPAddr, strLANIPAddr, strNodeName, strSiteAddr, strLocalX25DNA

'strCustomerName = Request.Cookies("CustomerName")
'strWinName	= Request.Cookies("WinName") 

'create lists
dim sql
dim rsNetworkElementType, rsRegion

'get the list of network element types
sql = "select NETWORK_ELEMENT_TYPE_CODE from CRP.NETWORK_ELEMENT_TYPE where RECORD_STATUS_IND='A'"
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
<META NAME="GENERATOR" Content="Microsoft FrontPage 4.0">
<META HTTP-EQUIV="Pragma" CONTENT="no-cache">
<TITLE>SMA - RSAS POS TIER3 Search</TITLE>
<LINK href="stylesheets/styles.css" rel="stylesheet" type="text/css">

<script type="text/javascript" SRC="GeneralJavaFunctions.js"></script>
<script type="text/javascript" SRC="AccessLevels.js"></script>
<script>
var intAccessLevel=<%=intAccessLevel%>;

//set section title
setPageTitle("SMA - POS PLUS Tier3");

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
	 	strWinName = document.frmRSAST3Search.hdnWinName.value ; 
	 	if (strWinName !=  "" ){
 			DeleteCookie("WinName") ;
		}
		
		var strCustomerName;
		strCustomerName = document.frmRSAST3Search.txtCustomer.value;
	 	DeleteCookie("CustomerName");
	 	if (strCustomerName != "")
	 	{
 			document.frmRSAST3Search.submit();
 		}
	}

function confirm_search(theForm)
{
var bolConfirm


	if (isWhitespace(theForm.txtCustomer.value) && isWhitespace(theForm.txtGWRT1SIIPAddr.value) &&
        isWhitespace(theForm.txtDLCIX25.value) && isWhitespace(theForm.txtDLCIIP.value) && 
		isWhitespace(theForm.txtNodeName.value) && isWhitespace(theForm.txtSiteAddr.value) &&
		isWhitespace(theForm.txtWANIPAddr.value) && isWhitespace(theForm.txtLANIPAddr.value) &&
		isWhitespace(theForm.txtTCNumber.value)  && isWhitespace(theForm.txtLocalX25DNA.value)&&
		isWhitespace(theForm.txtOrderNo.value))
		{
		bolConfirm = window.confirm("No Search Criteria have been entered. This search may take a long time..Continue?");
		if (bolConfirm){
			return true;
		}
		else
		{
			return false;
		}
	}
   return true;
}
	
function btnClear_onclick() {
	with(document.frmRSAST3Search){
		
		txtCustomer.value = "";
		txtGWRT1SIIPAddr.value = "";
		txtNodeName.value = "";
		txtLANIPAddr.value = "";
		txtDLCIIP.value = "";
		txtDLCIX25.value = "";
		chkActiveOnly.checked = true;
		txtTCNumber.value="";
		txtSiteAddr.value="";
		txtOrderNo.value="";
		txtWANIPAddr.value="";
		txtLocalX25DNA.value="";
	} 
}	

function btnNewGW_onclick(){
	if ((intAccessLevel & intConst_Access_Create) != intConst_Access_Create){
		alert('Access denied. Please contact your system administrator.');
		return;
	}
	parent.document.location='RSAST3GWDetail.asp?action=new&GWID=';
}
/*function btnNewTC_onclick(){
	if ((intAccessLevel & intConst_Access_Create) != intConst_Access_Create){
		alert('Access denied. Please contact your system administrator.');
		return;
	}
	parent.document.location='RSAST3Detail.asp?hdnTailCircuitID=';
}*/
</script>

</HEAD>
<BODY LANGUAGE=javascript onload="return window_onload();">
<form name=frmRSAST3Search action="RSAST3list.asp" method=POST target="fraResult" onSubmit="return confirm_search(this)">
<input type="hidden" name="hdnWinName" value="<%=strWinName%>">
<table border="0" width="100%">
<thead align=left>
  <tr>
    <td align=Left colSpan=4>POS PLUS Search</td>
  </tr>
</thead>
<tbody>
  <tr>
    <td align=right width="20%">Customer</td>
    <td width="30%"><INPUT name=txtCustomer tabindex=1 size=25 maxlength=50 value="<%=strCustomerName%>"></td>
		<!--option selected value="ALL"> </OPTION>	
		<%
		while not rsRegion.EOF
			Response.Write "<option value='" & rsRegion(0) & "'>" & routineHtmlString(rsRegion(1)) &"</option>" & vbCrLf
			rsRegion.MoveNext
		wend
		rsRegion.Close
		%>
	</SELECT></td-->

    <TD align=right nowrap width=15%>WAN IP Port Address</TD>
	<TD ALIGN=LEFT><INPUT id=txtWANIPAddr name=txtWANIPAddr tabindex=7 size=20 maxlength=20 value="<%=strWANIPAddr%>" ></TD>
  </tr>
  <tr>
    <td align=right>GW Router T1 Serial Interface IP Address </td>
    <td><INPUT name=txtGWRT1SIIPAddr tabindex=2 size=20 maxlength=20 value="<%=strGWRT1SIIPAddr%>"></td>
    <td align=right>LAN IP Port Address</td>
	<td ALIGN=LEFT><INPUT id=txtLANIPAddr name=txtLANIPAddr tabindex=8 size=20 maxlength=20 value="<%=strLANIPAddr%>" ></TD>
		<!--OPTION></OPTION>
		<%
		while not rsSG.EOF 
			Response.Write "<OPTION"
			Response.Write " VALUE="& rsSG("REMEDY_SUPPORT_GROUP_ID") &">" & routineHtmlString(rsSG("GROUP_NAME")) & "</OPTION>" &vbCrLf
			rsSG.MoveNext
		wend
		rsSG.Close
		%>
		</SELECT-->
  </tr>

  <tr>
	<td align=right>Gateway DLCI (X25)</td>
	<td><INPUT name=txtDLCIX25 tabindex=3 size=10 maxlength=10 value="<%=strDLCIX25%>"></td>
	<!--/ null? <input type=checkbox name=chkNullOBDialup></td-->
	<TD align=right nowrap width=15%>Tail Circuit Number </TD>
	<TD ALIGN=LEFT ><INPUT id=txtTCNumber name=txtTCNumber tabindex=9 size=19 maxlength=19 value="<%=strTCNumber%>"> </TD>
	<!--value="<%=strBarcode%>" ></TD--> 
  </TR>

  <tr>
	<td align=right>Gateway DLCI (IP)</td>
	<td><INPUT name=txtDLCIIP tabindex=4 size=10 maxlength=10 value="<%=strDLCIIP%>"></td>
	<!--/ null? <input name=chkNullIP type=checkbox></td-->
	<TD align=right nowrap width=15%>Local X25 Additional DNA</TD>
	<TD ALIGN=LEFT ><INPUT id=txtLocalX25DNA name=txtLocalX25DNA tabindex=10 size=10 maxlength=10 value="<%strLocalX25DNA%>"> </TD>
	<!--value="<%=strBarcode%>" ></TD--> 
  </tr>

  <TR>
    <td align=right>Node Name</td>
    <td><INPUT name=txtNodeName tabindex=5 size=10 maxlength=10 value="<%strNodeName%>"></td>
	<!--td align=right nowrap width=15%>Poll Code</td>
	<td align=left colspan=3 nowrap><select id=selPoll name=selPoll tabindex=12 style="width: 100">
	<option selected value="ALL"> </OPTION>
			<%
			while not rsNetworkElementType.EOF
				Response.Write "<option>" & routineHtmlString(rsNetworkElementType(0)) & "</option>" & vbCrLf
				rsNetworkElementType.MoveNext
			wend
			rsNetworkElementType.Close
			%>
		</SELECT-->

	<td align=right>Order Number</td>
	<td align=left ><INPUT name=txtOrderNo tabindex=11 size=16 maxlength=16 value=""> </td>
  </TR>
  <!--TR>
    <td align=right>Loop Back IP Address</td>
    <td><INPUT name=txtServLoc tabindex=5 size=25 maxlength=50 value="<%=strServLocName%>"></td>
	<td align=right nowrap width=15%>Poll Code</TD>
	<td align=left><INPUT name=chkNonCorrOnly tabindex=12 size=8 maxlength=30 value=""> </td>
	<td align=right>Active Only</td>
	<td align=left ><INPUT name=chkActiveOnly tabindex=14 type=checkbox value=yes checked><TD>	

  </TR-->

  <TR>
	<TD align=right nowrap width=15%>Site Address</TD>
	<TD ALIGN=LEFT width=20%><INPUT id=txtSiteAddr name=txtSiteAddr tabindex=6 size=30 maxlength=55 value="<%=strSiteAddr%>" ></TD> 
	
	<td align=right>Active Only</td>
	<td align=left ><INPUT name=chkActiveOnly tabindex=12 type=checkbox value=yes checked> </td>
  </TR>

  <TR>
	<td align=right colspan=2>
	<td align=right colspan=2>
	<% if strWinName <> "Popup" then %>
		<INPUT name=btnNew type=button value="New GW" tabindex=13 style= "width: 2cm" onclick="btnNewGW_onclick();">&nbsp;&nbsp;
		<!--INPUT name=btnNewTC type=button value="New TC" tabindex=14 style= "width: 2cm" onclick="btnNewTC_onclick();">&nbsp;&nbsp;-->
	<% end if %>
		<INPUT name=btnClear type=button value=Clear tabindex=14 style= "width: 2cm" onclick="return btnClear_onclick();">&nbsp;&nbsp;
		<INPUT name=btnSubmit type=submit tabindex=15 value=Search style= "width: 2cm" >&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
	</td>
  </tr>
<tbody>
</table>
</form>
</BODY>
</HTML>

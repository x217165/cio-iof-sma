<%@ Language=VBScript %>
<% option explicit %>
<% Response.Buffer = true %>

<!--#include file="smaConstants.inc"-->
<!--#include file="smaProcs.inc"-->
<!--#include file = "databaseconnect.asp"-->
<!--
********************************************************************************************
* Page name:	AddressCriteria.asp
* Purpose:		To dynamically set the criteria to search for an address.
*				Results are displayed via AddressList.asp
*
* In Param:		This page reads following cookies
*				CustomerName
*				WinName
*
* Created by:	Sara Sangha	Aug. 14th, 2000
*
********************************************************************************************
        Date		Author			Changes/enhancements made
        -----		------		------------------------------------------------------
       14-Nov-01	  DTy		Turn buffer on.
       24-Feb-02	  DTy		Simplified Search Screen.
       29-Jul-15   PSmith  Set Cookies in validation so the back key works
       05-Oct-15   PSmith  Only sumbit() for pop-up windows
       03-Feb-16   PSmith  Don't pre-populate search criteria
**************************************************************************************************
-->
<%

dim intAccessLevel
intAccessLevel = CInt(CheckLogon(strConst_Address))
if (intAccessLevel and intConst_Access_ReadOnly) <> intConst_Access_ReadOnly then
	DisplayError "BACK", "", 0, "ACCESS DENIED", "You don't have access to Address. Please contact your system administrator"
end if

dim objRsProvince, strSQL,  rsRegion
dim strCustomerName, strWinName

strCustomerName = Request.Cookies("CustomerName")
strWinName	= Request.Cookies("WinName")

strSQL = "select s.province_state_lcode, s.province_state_name " &_
		 "from crp.lcode_province_state s, " &_
				"crp.lcode_country c " &_
		 "where s.record_status_ind = 'A' " &_
		 "and	  s.country_lcode = c.country_lcode " &_
		 "order by s.country_lcode, s.province_state_name "

set objRsProvince = objConn.Execute(StrSql)

	'get Region List
	strSQL = "select noc_region_lcode, noc_region_desc" & _
			 " from crp.lcode_noc_region" & _
			 " where record_status_ind = 'A'" & _
			 " order by noc_region_desc"
	set rsRegion = Server.CreateObject("ADODB.Recordset")
	rsRegion.CursorLocation = adUseClient
	rsRegion.Open strSQL, objConn
	if err then
	end if
	'release the active connection, keep the recordset open
	set rsRegion.ActiveConnection = nothing

%>

<HTML>
<HEAD>
	<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
	<META HTTP-EQUIV="Pragma" CONTENT="no-cache">
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
	<LINK rel="stylesheet" type="text/css" href="stylesheets/styles.css">
	<script type="text/javascript" SRC="GeneralJavaFunctions.js"></script>
	<script type="text/javascript" SRC="AccessLevels.js"></script>
	<SCRIPT type = "text/javascript">

//*************************************Java Functions*******************************************

var intAccessLevel = "<%=intAccessLevel%>" ;

//set section title
setPageTitle("SMA - Address");

function validate(theForm){
//***********************************************************************************************
// Function:	validate()																		*
//																								*
// Purpose:		To validate that minimum search critiera has been entered so that full database *
//				search can be avoided.															*
//																								*
// Created By:	Sara Sangha		Aug. 28th, 2000													*
//																								*
// Updated By:																					*
//***********************************************************************************************

var bolConfirm;

if (isWhitespace(theForm.txtCustomerName.value) &&
    isWhitespace(theForm.txtStreet.value) &&
    isWhitespace(theForm.txtCity.value) &&
    isWhitespace(theForm.txtPostal.value) &&
    (theForm.selProvince.selectedIndex == 0) &&
    (theForm.selRegion.selectedIndex == 0 ))
 {
   bolConfirm = window.confirm("No search criteria have been entered. This search may take a long time...Continue?");
    if (!bolConfirm){
	 // abort search
     return false;
    }
  }
   // search critiera has been entered so continue search
   thinking(parent.fraResult);
   return true;
 }

function btnNew_onclick() {
//************************************************************************************************
// Function:	btnAddNew_onclick()
//
// Purpose:		To bring up a blank Address Detail page so that user can enter a new address.
//
// Created By:	Sara Sangha Aug. 28th, 2000
//
// Updated By:
//************************************************************************************************


	if ((intAccessLevel & intConst_Access_Create) != intConst_Access_Create) {
		alert('Access denied. Please contact your system administrator.');
		return;
	}

	parent.document.location.href ="AddressDetail.asp?AddressID=0" ;

}

function window_onload() {
/************************************************************************************************
*  Function:	window_onload																	*
*																								*
*  Purpose:		To submit the form automatically when txtCustomerName has a value.				*
*				This will happen when this page is called by a lookup function or by the Quick	*
*    		    Navigation drop-down box, which had already saved some values in the cookies	*
* 				and this form's VBScript has read those values and put in the form controls.	*
*																								*
*  Created By:	Sara Sangha Aug 25, 2000														*
*																								*
*  Updated By:																					*
*************************************************************************************************/

		var strCustomerName ;
	 	var strWinName;
	 	var hdnCustomerID;

	 	strCustomerName = document.frmAddSearch.txtCustomerName.value ;
	 	strWinName = document.frmAddSearch.hdnWinName.value ;
	 	hdnCustomerID = GetCookie("CustomerID");

	 	if (strWinName == "Simple"){
	 	   SetCookie("strSimple", "Simple");
	 	   SetCookie("txtCustomerName", strCustomerName);
	 	   SetCookie("CustomerID", hdnCustomerID);}
	 	else
	 	{DeleteCookie("strSimple")}

	 	DeleteCookie("CustomerName") ;
 		DeleteCookie("WinName") ;
	 	if ( strWinName == "Popup" && strCustomerName !=  "" ){
		  SetCookie("CustomerName", document.frmAddSearch.txtCustomerName.value);
		  thinking(parent.fraResult);
 			document.frmAddSearch.submit() ;
 		}
}
function btnClear_onclick() {
	document.frmAddSearch.txtCustomerName.value = ""
	document.frmAddSearch.txtStreet.value = ""
	document.frmAddSearch.txtCity.value = ""
	document.frmAddSearch.selProvince.selectedIndex = 0 ;
	document.frmAddSearch.txtPostal.value = ""
	document.frmAddSearch.chkActiveOnly.checked=true;
	document.frmAddSearch.selRegion.selectedIndex = 0 ;
}

//********************************* end of java functions******************************************
//-->
</SCRIPT>
</HEAD>
<BODY LANGUAGE=javascript onload="return window_onload()">
<FORM name=frmAddSearch method=post action="AddressList.asp" target="fraResult" onSubmit="return validate(this)">
	<!-- hidden fields -->
	<INPUT name=hdnWinName type=hidden value="<%=strWinName%>">

<TABLE width=100%>
	<thead>
    <% if strWinName = "Simple" then %>
  	   <TR><TD align=left colspan=3>Simplified Address Search</td></TR>
    <% else %>
  	   <TR><TD align=left colspan=4>Address Search</td></TR>
    <% end if %>
	</thead>
    <TR>
        <TD align=right width=15%>Customer Name</TD>
        <TD align=left width=25%><INPUT id=txtCustomerName name=txtCustomerName tabindex=1 size=30 value="<%=strCustomerName%>"></TD>

        <% if strWinName <> "Simple" then %>
           <TD align=right width=20%>Region</TD>
           <TD><select id=selRegion name=selRegion tabindex=6>
				   <option></option>
				   <%while not rsRegion.EOF
				        Response.write "<option value='" & rsRegion(0) & "'>" & routineHtmlString(rsRegion(1)) & "</option>" & vbCrLf
				        rsRegion.movenext
				     wend
				     rsRegion.Close
				   %>
			</select></td>
        <% end if %>

    <TR>
        <TD align=right width=15%>Street</TD>
        <TD align=left width=25%><INPUT id=txtStreet name=txtStreet tabindex=2 size=30 ></TD>

        <% if strWinName <> "Simple" then %>
              <TD align=right width=20%>Address Type:</TD>
              <TD><INPUT id=radAddressType name=radAddressType type=radio value=billing tabindex=7>Billing&nbsp;&nbsp;
			  <INPUT id=radAddressType name=radAddressType type=radio value=mailing tabindex=7>Mailing
		      </TD>
        <% end if %>
		</TR>

    <TR>
        <TD align=right width=15%>City / Municipality</TD>
        <TD width=25%><INPUT id=txtCity name=txtCity size=30 tabindex=3></td>
        <td width=20%>&nbsp;</td>

        <% if strWinName <> "Simple" then %>
		<td><INPUT id=radAddressType name=radAddressType type=radio value=primary tabindex=7>
			Primary
			<INPUT id=radAddressType name=radAddressType type=radio checked value=all tabindex=7>
			All
		</TD>
        <% end if %>

	</TR>

    <% if strWinName <> "Simple" then %>
       <TR>
       <TD align=right>Postal / Zip Code</TD>
       <TD><INPUT id=txtPostal name=txtPostal size=30 tabindex=4></TD>
       <TD align=right width=20%>Active Only</TD>
       <TD><INPUT id=chkActiveOnly name=chkActiveOnly type=checkbox VALUE=yes checked tabindex=8></TD>
       </TR>
    <% end if %>

    <TR>
        <TD align=right width=15%>Province</TD>
        <TD align=left width=25%><SELECT name=selProvince tabindex=5>
			<OPTION></OPTION>
			<%Do while Not objRsProvince.EOF
				Response.write "<OPTION VALUE ="& objRsProvince(0) & ">" & objRsProvince(0) & "&nbsp;&nbsp;&nbsp;" & objRsProvince(1) & "</OPTION>"
				objRsProvince.MoveNext
				Loop
			%>
			</SELECT>
        </TD>

    <% if strWinName = "Simple" then %>
          <TR>
          <TD align=right width=20%>Active Only</TD>
          <TD><INPUT id=chkActiveOnly name=chkActiveOnly type=checkbox VALUE=yes checked tabindex=8></TD>
    <% end if %>

       <TD colSpan=2 align=right>
    <% if strWinName <> "Popup" then %>
       <INPUT id=btnNew name=btnNew type=button  style="width: 2cm" value=New LANGUAGE=javascript onclick="return btnNew_onclick()" tabindex=9> &nbsp;&nbsp;
    <% end if %>

      <INPUT id=btnClear name=btnClear type=button style="width: 2cm" value=Clear LANGUAGE=javascript onclick="return btnClear_onclick()" tabindex=10>  &nbsp;&nbsp;
      <INPUT id=btnSearch name=btnSearch type=submit style="width: 2cm" value=Search tabindex=11> &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
      </TD>
    </TR>
</TABLE>
</FORM>
</BODY>
</HTML>

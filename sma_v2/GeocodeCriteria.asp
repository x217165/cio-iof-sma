<%@ Language=VBScript %>
<% option explicit %>
<% Response.Buffer = true %>

<!--#include file="smaConstants.inc"-->
<!--#include file="smaProcs.inc"-->
<!--#include file = "databaseconnect.asp"-->
<!--
********************************************************************************************
* Page name:	GeocodeCriteria.asp
* Purpose:		To dynamically set the criteria to search for an address.
*				Results are displayed via GeocodeList.asp
*
* In Param:		This page reads following cookies
*				GeoStreet
*				GeoCity
*				WinName
*
* Created by:	Sara Sangha	Aug. 14th, 2000
*
********************************************************************************************
        Date		Author			Changes/enhancements made
        -----		------			------------------------------------------------------
       12-May-08	ACheung, LChen		NGSM CLLI impelementation
***************************************************************************************************
-->
<%

dim intAccessLevel
intAccessLevel = CInt(CheckLogon(strConst_Address))
if (intAccessLevel and intConst_Access_ReadOnly) <> intConst_Access_ReadOnly then
	DisplayError "BACK", "", 0, "ACCESS DENIED", "You don't have access to Address. Please contact your system administrator"
end if

dim objRsProvince, strSQL,  rsRegion
'dim strGeocode
dim straddress, strCity
dim strWinName

'strGeocode = Request.Cookies("Geocode")
strWinName	= Request.Cookies("WinName")
straddress      = Request.Cookies("GeoStreet")
strCity         = Request.Cookies("GeoCity")

strSQL = "select distinct province as province from crp.lcode_geocodeid "

set objRsProvince = objConn.Execute(StrSql)

'Response.Write "Win=" & strWinName
'Response.Write "Street=" & straddress
'Response.Write "City=" & strCity
'Response.End
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
setPageTitle("SMA - CLLI Code");

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

var bolConfirm
if (isWhitespace(theForm.txtGeoclli.value) &&
    isWhitespace(theForm.txtGeocllicodeid.value) &&
    isWhitespace(theForm.txtAddress.value) &&
    isWhitespace(theForm.txtCity.value) &&
    isWhitespace(theForm.txtPostal.value) &&
    isWhitespace(theForm.txtDescription.value) &&
    (theForm.selProvince.selectedIndex == 0))
 {
   bolConfirm = window.confirm("No search criteria have been entered. This search may take a long time...Continue?");
    if (bolConfirm){
	  // continue search even no creitiera has been entered
      return true;
    }
    else
    {
	 // abort search
     return false;
    }
  }
   // search critiera has been entered so continue search
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

	parent.document.location.href ="GeocodeCriteria.asp?GeocodeID=0" ;

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

//		var strGeocllicode ;
	 	var strWinName;
	 	var hdnGeocodeID;

//	 	strGeocllicode = document.frmGeocodeSearch.txtGeoclli.value ;
	 	strWinName = document.frmGeocodeSearch.hdnWinName.value ;
//	 	hdnGeocodeID = GetCookie("GeocodeID");

	 	if (strWinName == "Simple"){
	 	   SetCookie("strSimple", "Simple");
	 	  // SetCookie("txtAddress", straddress);
	 	   SetCookie("GeocodeID", hdnGeocodeID);}
	 	else
	 	{DeleteCookie("strSimple")}

	 	DeleteCookie("GeocodeID") ;
 		DeleteCookie("WinName") ;
//	 	if ( strGeocllicode !=  "" ){
// 			document.frmGeocodeSearch.submit() ;
// 		}
}
function btnClear_onclick() {
	document.frmGeocodeSearch.txtGeocllicodeid.value = ""
	document.frmGeocodeSearch.txtAddress.value = ""
	document.frmGeocodeSearch.txtDescription.value = ""
	document.frmGeocodeSearch.txtCity.value = ""
	document.frmGeocodeSearch.txtPostal.value = ""
	document.frmGeocodeSearch.selProvince.selectedIndex = 0 ;
	document.frmGeocodeSearch.txtGeoclli.value = ""

//	document.frmGeocodeSearch.chkActiveOnly.checked=true;
}

//********************************* end of java functions******************************************
//-->
</SCRIPT>
</HEAD>
<BODY LANGUAGE=javascript onload="return window_onload()">
<FORM name=frmGeocodeSearch method=post action="GeocodeList.asp" target="fraResult" onSubmit="return validate(this)">
	<!-- hidden fields -->
	<INPUT name=hdnWinName type=hidden value="<%=strWinName%>">

<TABLE width=100%>
	<thead>
    <% if strWinName = "Simple" then %>
  	   <TR><TD align=left colspan=3>Simplified CLLI CODE Search</td></TR>
    <% else %>
  	   <TR><TD align=left colspan=4>CLLI CODE Search</td></TR>
    <% end if %>
	</thead>
    <TR>
        <TD align=right width=15%>GeoCode ID </TD>
        <TD align=left width=25%><INPUT id=txtGeocllicodeid name=txtGeocllicodeid tabindex=1 size=30></TD>

        <% if strWinName <> "Simple" then %>
           <TD align=right width=20%>CLLI CODE </TD>
           <TD><INPUT id=txtGeoclli name=txtGeoclli size=30 tabindex=2></td>
        <% end if %>

    <TR>
        <TD align=right width=15%>Address</TD>
        <TD align=left width=25%><INPUT id=txtAddress name=txtAddress tabindex=3 size=30 value="<%=straddress%>" ></TD>

        <% if strWinName <> "Simple" then %>
              <TD align=right width=20%>Description </TD>
              <TD><INPUT id=txtDescription name=txtDescription size=30 tabindex=4>
		      </TD>
        <% end if %>
		</TR>

    <TR>
        <TD align=right width=15%>City</TD>
        <TD width=25%><INPUT id=txtCity name=txtCity size=30 tabindex=5 value="<%=strCity%>"></td>
        <td width=20%>&nbsp;</td>

        <% if strWinName <> "Simple" then %>
		<td>&nbsp;</TD>
        <% end if %>

	</TR>

    <% if strWinName <> "Simple" then %>
       <TR>
       <TD align=right>Postal Code</TD>
       <TD><INPUT id=txtPostal name=txtPostal size=30 tabindex=6></TD>
       <TD align=right width=20%>&nbsp;</TD>
       <TD>&nbsp;</TD>
       </TR>
    <% end if %>

    <TR>
        <TD align=right width=15%>Province</TD>
        <TD align=left width=25%><SELECT name=selProvince tabindex=7>
			<OPTION></OPTION>
			<%  objRsProvince.MoveFirst
			    Do while Not objRsProvince.EOF
'				response.write "<option value = " & objRsProvince(0) & ">" & objRsProvince(0) & "</OPTION>"%>
   				<option> <% response.write(objRsProvince(0)) %></OPTION>
			<%	    objRsProvince.MoveNext
			    Loop
			%>
			</SELECT>
        </TD>

    <% if strWinName <> "Simple" then %>
          <TR>
          <TD align=right width=20%>Active Only</TD>
          <TD><INPUT id=chkActiveOnly name=chkActiveOnly type=checkbox VALUE=yes checked tabindex=8></TD>
    <% end if %>

       <TD colSpan=2 align=right>
    <% if strWinName <> "Simple" then %><INPUT id=btnClear name=btnClear type=button style="width: 2cm" value=Clear LANGUAGE=javascript onclick="return btnClear_onclick()" tabindex=9>&nbsp; &nbsp;&nbsp;
    <% end if %>&nbsp;  &nbsp;&nbsp;
      <INPUT id=btnSearch name=btnSearch type=submit style="width: 2cm" value=Search tabindex=10> &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
      </TD>
    </TR>
</TABLE>
</FORM>
</BODY>
</HTML>

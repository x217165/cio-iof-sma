<%@ Language=VBScript %>
<% option explicit %>
<% Response.Buffer = true %>

<!--#include file="smaConstants.inc"-->
<!--#include file="smaProcs.inc"-->
<!--#include file = "databaseconnect.asp"-->
<!--
********************************************************************************************
* Page name:	RSAST3AddrCriteria.asp
* Purpose:		To dynamically set the criteria to search for a simplified address.
*				Results are displayed via RSAST3AddrList.asp
*               Contains a few search fields than the original version.
*	
* In Param:		This page reads following cookies
*				CustomerName
*				WinName
*
* Created by:	DTy		Dec 31, 2001 based on AddressCriteria.asp modified for RSAS POS PLUS.
*  
********************************************************************************************
        Date		Author			Changes/enhancements made
        -----		------		------------------------------------------------------
***************************************************************************************************
-->
<%

dim intAccessLevel
intAccessLevel = CInt(CheckLogon(strConst_Address))
if (intAccessLevel and intConst_Access_ReadOnly) <> intConst_Access_ReadOnly then
	DisplayError "BACK", "", 0, "ACCESS DENIED", "You don't have access to Address. Please contact your system administrator"
end if

dim strSQL				
dim strCustomerName, strWinName

dim rsLocation					
dim strStreet, strMunicipality, strProvince, strCountry, intSiteAddressID, strSiteAddress

'bolConfirm = window.confirm("Building city/province/country list.  Please wait ...");

strWinName	     = Request.Cookies("WinName") 

strCustomerName  = Request.Cookies("CustomerName")
strStreet        = Request.Cookies("Street")
strMunicipality  = Request.Cookies("Municipality")
strProvince      = Request.Cookies("Province")
strCountry       = Request.Cookies("Country")
intSiteAddressID = Request.Cookies("SiteAddressD")
strSiteAddress   = Request.Cookies("SiteAddress")


'GetLocation: municipality + province + country
strSQL = "select m.municipality_name, s.province_state_name, c.country_desc, " &_
		 "       m.province_state_lcode, m.country_lcode " &_
		 "from crp.municipality_lookup m, " &_
				"crp.lcode_province_state s, " &_
				"crp.lcode_country c " &_
		 "where m.record_status_ind = 'A' " &_
		 "and s.record_status_ind = 'A' " &_
		 "and c.record_status_ind = 'A' " &_
		 "and m.province_state_lcode = s.province_state_lcode " &_
		 "and m.country_lcode = c.country_lcode " &_
		 "order by m.municipality_name, s.province_state_name, c.country_desc "

set rsLocation = objConn.Execute(StrSql)

%>

<HTML>
<HEAD>
	<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
	<META HTTP-EQUIV="Pragma" CONTENT="no-cache">
	<LINK rel="stylesheet" type="text/css" href="stylesheets/styles.css">
	<script type="text/javascript" SRC="GeneralJavaFunctions.js"></script>
	<script type="text/javascript" SRC="AccessLevels.js"></script> 
	<SCRIPT type = "text/javascript">

//*************************************Java Functions*******************************************

var intAccessLevel = "<%=intAccessLevel%>" ;

//set section title
setPageTitle("RSAS POS - Address");
	
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
if (isWhitespace(theForm.txtCustomerName.value) && 
    isWhitespace(theForm.txtStreet.value) && 
    (theForm.selLocation.selectedIndex == 0))
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
	parent.document.location.href ="RSAST3AddrDetail.asp?AddressID=0" ;
//	window.open('SearchFrame.asp?fraSrc=RSAST3AddrDetail', 'Popup', 'top=50, left=100, WIDTH=800, HEIGHT=600' ) ;
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
	 	
	 	strCustomerName = document.frmAddrSearch.txtCustomerName.value ;
	 	strWinName = document.frmAddrSearch.hdnWinName.value ; 
			 	
	 	DeleteCookie("CustomerName") ;
 		DeleteCookie("WinName") ;

	 	if ( strCustomerName !=  "" ){
	 	
 			document.frmAddrSearch.submit() ;  	
 		
 		}	
}
function btnClear_onclick() {
	document.frmAddrSearch.txtCustomerName.value = ""
	document.frmAddrSearch.txtStreet.value = "" 
	document.frmAddrSearch.selLocation.selectedIndex = 0 ;   
	document.frmAddrSearch.chkActiveOnly.checked=true;   
}

//********************************* end of java functions******************************************
//-->
</SCRIPT>
</HEAD>
<BODY LANGUAGE=javascript onload="return window_onload()">
<FORM name=frmAddrSearch method=post action="RSAST3AddrList.asp" target="fraResult" onSubmit="return validate(this)">
	<!-- hidden fields -->
	<INPUT name=hdnWinName type=hidden value="<%=strWinName%>">
	<INPUT name=hdnCustName type=hidden value="<%=strCustomerName%>">
	
<TABLE width=100%>
	<thead>
	<TR><TD align=left colspan=4>Address Search</td></TR>
	</thead>
    <TR>
        <TD align=right colspan=2>Customer Name</TD>
        <TD align=left colspan=2><INPUT id=txtCustomerName name=txtCustomerName disabled size=50 value="<%=strCustomerName%>"></TD>
    <TR>
        <TD align=right colspan=2>Street</TD>
        <TD align=left colspan=2><INPUT id=txtStreet name=txtStreet tabindex=1 size=75 ></TD>
	</TR>

    <TR>
        <TD align=right colspan=2>Municipality/Province/Country</TD>
        <TD width=25%>
			<SELECT name=selLocation tabindex=2 size=1>
			<OPTION></OPTION>
			<%Do while Not RSLocation.EOF
				 if rsLocation(0)& rsLocation(3) & rsLocation(4) = strMunicipality & strProvince & strCountry then
					Response.write "<OPTION VALUE ="& rsLocation(0) & "/" & rsLocation(1) & "/" & rsLocation(2) & " SELECTED>" & rsLocation(0) & ", &nbsp;" & rsLocation(1) & ", &nbsp;" & rsLocation(2)& "</OPTION>"
				 else
					Response.write "<OPTION VALUE ="& rsLocation(0) & "/" & rsLocation(1) & "/" & rsLocation(2) & ">" & rsLocation(0) & ", &nbsp;" & rsLocation(1) & ", &nbsp;" & rsLocation(2)& "</OPTION>"
				 end if
				rsLocation.MoveNext
				Loop
				rsLocation.Close
			%>	
			</SELECT>
		</TD>
	</TR>
    <TR>
        <TD align=right colspan=2>Active Only</TD>
        <TD><INPUT id=chkActiveOnly name=chkActiveOnly tabindex=3 type=checkbox VALUE=yes checked tabindex=4></TD>
	</TR>
	
	<TR>
		<TD></TD>
        <TD colSpan=2 align=right>
            <INPUT id=btnNew name=btnNew type=button  style="width: 2cm" value=Save LANGUAGE=javascript onclick="return btnNew_onclick()" tabindex=4> &nbsp;&nbsp;
            <INPUT id=btnClear name=btnClear type=button style="width: 2cm" value=Clear LANGUAGE=javascript onclick="return btnClear_onclick()" tabindex=5>  &nbsp;&nbsp;
            <INPUT id=btnSearch name=btnSearch type=submit style="width: 2cm" value=Search tabindex=6> &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
        </TD>
    </TR>
</TABLE>
</FORM>
</BODY>
</HTML>

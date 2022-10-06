<%@ Language=VBScript %>
<%
Response.buffer=true
Response.Expires = -1
Response.ExpiresAbsolute = Now() -1 
Response.AddHeader "pragma", "no-store"
Response.AddHeader "cache-control","no-store, no-cache, must-revalidate"
%>
<!--#include file="smaConstants.inc"-->
<!--#include file="smaProcs.inc"-->
<!--#include file = "databaseconnect.asp"-->
<!--

********************************************************************************************
* Page name:	CustServCPCriteria.asp
* Purpose:		To dynamically set the criteria to search for a Customer Service with Customer Profile information available.
*				Results are displayed via CustServCPList.asp
*
* In Param:		This form reads following cookies :
*				 - WinName
*				 - CustomerName
*				 - CustomerService
*				 - ServiceEnd
*
* Created by:	Sara Sangha	Aug. 14th, 2000
* Edited by:    Adam Haydey Jan. 25th, 2000
*               Added Customer Service City and  Customer Service Address search fields.
********************************************************************************************
        Date		Author			Changes/enhancements made
        -----		------		------------------------------------------------------
       28-Feb-02	 DTy		Add 'Alias' to 'Customer Service Name' field name.
       09-Sep-04     MW         Add  Repair Priority select list
   	   10-Aug-12     ACheung	Add Customer ID and Customer Shortname
   	   27-May-13	 ACheung	Add Customer Profile (adapted from CustServCriteria.asp)
       29-Jul-15   PSmith  Set Cookies in validation so the back key works
       05-Oct-15   PSmith  Only sumbit() for pop-up windows
       03-Feb-16   PSmith  Don't pre-populate search criteria
********************************************************************************************
-->
<%
dim intAccessLevel
dim objRsRegion ,objRsSupportGroup, objRsStatus, objRsRepairPriority, objRsCustomerName, strSQL
dim strCustomerService, strWinName,strServiceEnd, strServLocName, StrCustomerName, strLANG
Dim strServiceTypeID, strServiceTypeName, strCustomerID, strCustomerShortName
dim strCustomerProfileName, strCustomerProfileID, strVpnName

intAccessLevel = CInt(CheckLogon(strConst_CustomerService))
if (intAccessLevel and intConst_Access_ReadOnly) <> intConst_Access_ReadOnly then
	DisplayError "BACK", "", 0, "ACCESS DENIED", "You don't have access to view customer service. Please contact your system administrator."
end if

	strLANG = Request.Cookies("UserInformation")("language_preference")
	if (Len(strLANG) = 0) then strLANG = "EN"

	strCustomerService = Request.Cookies("CustomerService")
	strWinName	= Request.Cookies("WinName")
    strServiceEnd = Request.Cookies("ServiceEnd")
    strServLocName = Request.Cookies("ServLocName")
    'Response.Write "end=" & strServiceEnd
	strCustomerName = Request.Cookies("CustomerName")
	strCustomerServiceId = Request.Cookies("CustomerServiceID")
	strServiceTypeID = Request.Cookies("ServiceTypeID")
	strServiceTypeName = Request.Cookies("ServiceTypeName")
	strCustomerID = Request.Cookies("hdnCustomerID")
	strCustomerShortName = Request.Cookies("CustomerShortName")

	'get a list of region codes
	strSQL = "select noc_region_lcode, " &_
					"noc_region_desc " &_
			 "from crp.lcode_noc_region " &_
			 "where record_status_ind = 'A' " &_
			 "order by noc_region_desc"

	set objRsRegion = objConn.Execute(strSQL)

	'get a list of service status codes
	strSQL = "SELECT service_status_code, " &_
					"service_status_name " &_
			 "FROM crp.service_status " &_
			 "WHERE record_status_ind = 'A' " &_
			 "order by service_status_name "

	set objRsStatus = objConn.Execute(strSQL)

	'get a list of support groups
	strSQL = "SELECT remedy_support_group_id, " &_
					"group_name " &_
			  "FROM crp.v_remedy_support_group " &_
			  "order by group_name"

	set objRsSupportGroup = objConn.Execute(strSQL)


	'get the LYNX repair priority list

	strSQL = "SELECT lynx_def_sev_lcode " &_
	         ",      lynx_def_sev_desc " &_
			 "FROM crp.lcode_lynx_def_sev " &_
			 "WHERE record_status_ind='A' "  &_
			 "ORDER BY lynx_def_sev_lcode"

	set objRsRepairPriority = objConn.Execute(strSQL)

	' get the exact customer name

	'strSQL = "SELECT distinct customer_name " &_
	'         ",      customer_id " &_
	'	     "FROM crp.customer " &_
	'	     "WHERE record_status_ind='A' "  &_
	'		 "AND CUSTOMER_ID = " &  strCustomerID

	'set objRsCustomerName = objConn.Execute(strSQL)

	'strCustomerName = objRsCustomerName("customer_name")
%>

<HTML>
<HEAD>
<LINK rel="stylesheet" type="text/css" href="stylesheets/styles.css">
<TITLE>Customer Service Search</TITLE>
<META NAME="GENERATOR" Content="Microsoft FrontPage 12.0">
<META HTTP-EQUIV="Pragma" CONTENT="no-cache">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<script type="text/javascript" SRC="GeneralJavaFunctions.js"></script>
<script type="text/javascript" SRC="AccessLevels.js"></script>
<script type="text/javascript">
//****************************************** Java Functions *****************************
//set section title
var intAccessLevel = "<%=intAccessLevel%>" ;
setPageTitle("SMA - Customer Service");
if ( window.history.replaceState ) {
  window.history.replaceState( null, null, window.location.href );
}
function validate(theForm){

	var bolConfirm ;
	if (isWhitespace(theForm.txtCustomerServiceDesc.value) && isWhitespace(theForm.selSupportGroup.value)
		&& isWhitespace(theForm.txtCustomerProfileName.value)
		&& isWhitespace(theForm.txtCustomerProfileID.value)
	    && isWhitespace(theForm.selRepairPriority.value)
		&& isWhitespace(theForm.txtCustomerName.value)&& isWhitespace(theForm.selStatus.value)
		&& isWhitespace(theForm.txtCustomerServiceID.value) && isWhitespace(theForm.selRegion.value)
		&& isWhitespace(theForm.txtServiceLocationName.value) && isWhitespace(theForm.txtOrderNo.value)
		&& isWhitespace(theForm.txtServiceType.value)  && isWhitespace(theForm.txtServiceCity.value)
		&& isWhitespace(theForm.txtServiceAddress.value)&& isWhitespace(theForm.txtCustomerShortName.value)
		&& isWhitespace(theForm.txtCustomerID.value))
	{
		bolConfirm = window.confirm("No Search Criteria have been entered. This search may take a long time..Continue?");
		if (!bolConfirm){
			return false;
		}
	}

	if ( theForm.txtCustomerServiceID.value != "" && isNaN(Number(theForm.txtCustomerServiceID.value))) {
		alert("Customer Service ID must be a number");
		document.frmCSCPSearch.txtCustomerServiceID.focus();
		document.frmCSCPSearch.txtCustomerServiceID.select();
		return(false) ;
	}
	else if ( theForm.txtCustomerID.value != "" && isNaN(Number(theForm.txtCustomerID.value))) {
		alert("Customer ID must be a number");
		document.frmCSCPSearch.txtCustomerID.focus();
		document.frmCSCPSearch.txtCustomerID.select();
		return(false) ;
	}

  // Start thinking
  thinking(parent.fraResult);

   return true;

}

function fct_lookupCustomerProfile(Cust){

//    SetCookie("ServiceEnd", Cust);
	//var strCustomerName = document.frmCSSearch.txtCustomerName.value;
	var strCustomerProfileName = document.frmCSCPSearch.txtCustomerProfileName.value;

	//alert (Cust);

 	if (Cust != ""){
 		SetCookie("ServiceEnd",Cust);
		//alert (Cust);
 		}

	if (strCustomerProfileName != "" ) {
		SetCookie("CustomerProfileName", strCustomerProfileName) ;
		//alert (strCustomerProfileName);
		}


	SetCookie("WinName", 'Popup');
	window.open('SearchFrame.asp?fraSrc=Cust_Profile', 'Popup', 'top=50, left=100, height=600, width=800' ) ;
	//document.frmCSSearch.txtCustomerName.value = txtCustomerName ;
}


function btnAdd_onclick(){

	if ((intAccessLevel & intConst_Access_Create) != intConst_Access_Create) {
		alert('Access denied. Please contact your system administrator.');
		return;
	}

//	parent.location = 'CustServDetail.asp?CustServID=0'
	parent.location = 'CustServDetail.asp?CustServID=0&ServiceTypeID=0'
}

function fct_lookupCustomer(CustService){

    //SetCookie("ServiceEnd", CustService);
	var strCustomerName = document.frmCSCPSearch.txtCustomerName.value;

	//alert (CustService);

 	if (CustService != ""){
 		SetCookie("ServiceEnd",CustService);
		//alert (CustService);
 		}

	if (strCustomerName != "" ) {
		SetCookie("CustomerName", strCustomerName ) ;
		//alert (strCustomerName);
		}

	/*if (document.frmCSCPSearch.txtCustomerName.value != "")
		{SetCookie("CustomerName", document.frmCSCPSearch.txtCustomerName.value);
		}*/

	SetCookie("WinName", 'Popup');
	window.open('SearchFrame.asp?fraSrc=Cust', 'Popup', 'top=50, left=100, height=600, width=800' ) ;
	//document.frmCSCPSearch.txtCustomerName.value = txtCustomerName ;
}

function window_onload() {
//***********************************************************************************************
// Function: window_onload																		*
//																								*
// Purpose:		To submit the form automatically when txtCustomerName has a value.				*
//				This will happen when this page is called by a lookup function or by the Quick	*
//			    Navigation drop-down box, which had already saved some values in the cookies	*
//				and this form's VBScript has read those values and put in the form controls.	*
//																								*
// Created By:	Sara Sangha Aug. 25th 2000														*
//																								*
// Updated By:																					*
//***********************************************************************************************
	var strCustomerService = document.frmCSCPSearch.txtCustomerServiceDesc.value;
	var strCustomerName = window.frmCSCPSearch.txtCustomerName.value ;
	var strServiceLocationName = window.frmCSCPSearch.txtServiceLocationName.value;
	var intCustomerServiceID = document.frmCSCPSearch.txtCustomerServiceID.value;
	var strServiceTypeName = document.frmCSCPSearch.txtServiceType.value;
  var strWinName = document.frmCSCPSearch.hdnWinName.value;
  
 	DeleteCookie("CustomerProfileName");
 	DeleteCookie("strCustomerProfileID");

	DeleteCookie("WinName");
	DeleteCookie("CustomerName");
	DeleteCookie("CustomerService");
	DeleteCookie("ServiceEnd");
	DeleteCookie("ServLocName");
	DeleteCookie("CustomerServiceID");
	DeleteCookie("ServiceTypeID");
	DeleteCookie("ServiceTypeName");
	DeleteCookie("hdnCustomerID");
	DeleteCookie("CustomerShortName");


	if (strWinName == "Popup" && (( intCustomerServiceID != "" ) ||( strCustomerName != "" ) || ( strCustomerService != "" ) || (strServiceTypeName != "" ) || (strServiceLocationName != "" )))
	{
		SetCookie("CustomerName",document.frmCSCPSearch.txtCustomerName.value);
		SetCookie("CustomerServiceID",document.frmCSCPSearch.txtCustomerServiceID.value);
		SetCookie("CustomerService",document.frmCSCPSearch.txtCustomerServiceDesc.value);
		SetCookie("ServiceTypeName",document.frmCSCPSearch.txtServiceType.value);
		SetCookie("ServLocName",document.frmCSCPSearch.txtServiceLocationName.value);
    thinking(parent.fraResult);
		document.frmCSCPSearch.submit() ;
	}
}

function btnClear_onclick() {
	document.frmCSCPSearch.txtCustomerName.value = "" ;
	document.frmCSCPSearch.selStatus.selectedIndex = 0 ;
	document.frmCSCPSearch.txtCustomerServiceDesc.value = "" ;
	document.frmCSCPSearch.selRegion.selectedIndex = 0 ;
	document.frmCSCPSearch.txtServiceLocationName.value = "" ;
	document.frmCSCPSearch.txtCustomerServiceID.value = "" ;
	document.frmCSCPSearch.txtServiceCity.value = "" ;
	document.frmCSCPSearch.txtServiceAddress.value = "" ;
	document.frmCSCPSearch.selSupportGroup.selectedIndex = 0 ;
	document.frmCSCPSearch.txtOrderNO.value = "" ;
	document.frmCSCPSearch.txtServiceType.value = "" ;
	document.frmCSCPSearch.selRepairPriority.selectedIndex = 0 ;
	document.frmCSCPSearch.chkActiveOnly.checked=true;
	document.frmCSCPSearch.chkPrefLangOnly.checked=true;
	document.frmCSCPSearch.txtCustomerShortName.value = "" ;
	document.frmCSCPSearch.txtCustomerID.value = "" ;
	document.frmCSCPSearch.txtCustomerProfileName.value = "" ;
	document.frmCSCPSearch.txtCustomerProfileID.value = "" ;
	document.frmCSCPSearch.txtVpnName.value = "" ;
}



//-->
</SCRIPT>
<style type="text/css">
.style1 {
	margin-left: 0px;
}
</style>
</HEAD>
<BODY LANGUAGE=javascript onload="return window_onload()">

<FORM name=frmCSCPSearch method=post action="CustServCPList.asp" target="fraResult" onSubmit="return validate(this)">
	<INPUT id=hdnServiceEnd name=hdnServiceEnd type=hidden value="<%=strServiceEnd%>"><INPUT id=hdnWinName name=hdnWinName type=hidden value="<%=strWinName%>">
	<INPUT type="hidden" id="hdnServiceTypeID" name="hdnServiceTypeID" value="<%=strServiceTypeID%>">
<TABLE>
	<thead><tr><td colspan=4 align=left>Customer Service Search</tr></td></thead>
    <TR>
        <TD width=15% align=right nowrap>Customer Service ID</TD>
        <TD><INPUT id=txtCustomerServiceID name=txtCustomerServiceID tabindex=1 style="HEIGHT: 22px; WIDTH: 200px" value="<%=strCustomerServiceId%>" ></TD>

		<TD width= 15% align=right nowrap>Customer Name</TD>
		<TD nowrap align=left>
			<INPUT  name=txtCustomerName tabindex=2 type=text style="WIDTH: 250px; height: 26px; color:silver;" readonly maxlength=50
				value="<%=unescape(strCustomerName)%>" align=left><%if strCustomerName <> null then Response.write routineHTMLString(strCustomerName) else Response.Write null end if%>
		    <INPUT  name=btnCustomerLookup type=button value=... LANGUAGE=javascript onclick="fct_lookupCustomer('E')"></TD>
   </TR>
	<TR>
        <TD width=15% align=right nowrap>Customer Service Name/Alias</TD>
        <TD width=20%><INPUT id=txtCustomerServiceDesc name=txtCustomerServiceDesc tabindex=2 style="HEIGHT: 22px; WIDTH: 270px" value="<%=unescape(strCustomerService)%>">
		<TD width=15% align=right nowrap>Customer Short Name</TD>
        <TD width=20% nowrap align=left>
			<INPUT id=txtCustomerShortName name=txtCustomerShortName tabindex=3 style="HEIGHT: 22px; WIDTH: 270px" value="<%=strCustomerShortName%>" ></TD>
	</TR>
	<TR>
        <TD width=15% align=right nowrap>Service Type</TD>
		<TD width=20%><INPUT id=txtServiceType name=txtServiceType tabindex=6 style="HEIGHT: 22px; WIDTH: 270px" value="<%=strServiceTypeName%>"></TD>
        <TD width=15% align=right nowrap>Customer ID</TD>
        <TD><INPUT id=txtCustomerID name=txtCustomerID tabindex=4 style="HEIGHT: 22px; WIDTH: 200px" value="<%=strCustomerID%>" ></TD>
	</TR>
    <TR>
        <TD width=15% align=right nowrap>Service Location Name</TD>
        <TD width=20% ><INPUT id=txtServiceLocationName name=txtServiceLocationName tabindex=3 style="HEIGHT: 22px; WIDTH: 270px" value="<%=strServLocName%>" ></TD>
		<TD align=right nowrap width=15%>Customer Profile Name</TD>
		<TD align=left nowrap><INPUT name=txtCustomerProfileName tabindex=2 type=text size=40 style="WIDTH: 250px; height: 26px; color:silver;" readonly maxlength=50 value="<%=routineHTMLString(strCustomerProfileName)%>" align=left><INPUT  name=btnCustomerProfileLookup type=button value=... LANGUAGE=javascript onclick="fct_lookupCustomerProfile('B')"></TD>
	</TR>
    <TR>
		<TD width=15% align=right nowrap>Service Address </TD>
        <TD width=20% ><INPUT id=txtServiceAddress name=txtServiceAddress tabindex=4 style="HEIGHT: 22px; WIDTH: 270px" value="<%=strServAddress%>" ></TD>
			<TD align=right nowrap width=20%>Customer Profile ID (CPID)</TD>
			<TD align=left width=30%><INPUT name=txtCustomerProfileID size=40 maxlength=50 tabindex=1 value="<%=routineHTMLString(strCustomerProfileID)%>"></TD>
    </TR>
	<TR>
	    <TD width=15% align=right nowrap>Support Group</TD>
        <TD width=20% >
		<SELECT id=selSupportGroup name=selSupportGroup style="HEIGHT: 22px; WIDTH: 271px" tabindex=5>
			<OPTION value = " "selected>
			<% Do while not objRsSupportGroup.EOF %>
				<OPTION VALUE = "<%=objRsSupportGroup(0)%>" > <%=objRsSupportGroup(1)%>
				<%objRsSupportGroup.MoveNext%>
			<%Loop%>
			</SELECT></TD>
        <TD width= 15% align=right nowrap>Status</TD>
        <TD><SELECT id=selStatus name=selStatus tabindex=8 style="HEIGHT: 22px; WIDTH: 200px">
			<OPTION value = " ">
			<OPTION VALUE="AllExceptTerm" <% if ( strCustomerName = "" and strServLocName = "" and strCustomerService = "" ) then Response.write " selected " end if %>>All Except Terminated</option>
			<% Do while not objRsStatus.EOF %>
				<OPTION VALUE = "<%=objRsStatus(0)%>" > <%=objRsStatus(1)%>
				<%objRsStatus.MoveNext%>
			<%Loop%>
			</SELECT></TD>
	</TR>
	<TR>
		<TD width=15% align=right nowrap>Repair Priority</TD>
		<TD width=20% ><SELECT id=selRepairPriority name=selRepairPriority style="HEIGHT: 22px; WIDTH: 160px" tabindex=6>
			<option selected value="">&nbsp;</OPTION>
			<%
			while not objRsRepairPriority.EOF
				Response.Write "<option value=" & routineHtmlString(objRsRepairPriority(0)) & ">"
				Response.Write routineHtmlString(objRsRepairPriority(1))
				Response.Write "</option>"
				objRsRepairPriority.MoveNext
			wend
			objRsRepairPriority.Close
			%>
			</SELECT></TD>
		<TD colspan=1 align=right nowrap> Region </TD>
        <TD><SELECT id=selRegion name=selRegion tabindex=7 style="HEIGHT: 22px; WIDTH: 200px" >
				<OPTION value = " " selected>
				<% Do while not objRsRegion.EOF %>
				<OPTION VALUE = "<%=objRsRegion(0)%>" > <%=objRsRegion(1)%>
				<%objRsRegion.MoveNext%>
				<%Loop%>
			</SELECT></TD></TR>
   	<TR>
        <TD width=15% align=right nowrap>VPN Name</TD>
		<TD width=20%><INPUT id=txtVpnName name=txtVpnName tabindex=6 style="HEIGHT: 22px; WIDTH: 270px" value="<%=strVpnName%>"></TD>
        <TD width=15% align=right nowrap>Service Location City</TD>
        <TD><INPUT id=txtServiceCity name=txtServiceCity tabindex=10 style="HEIGHT: 22px; WIDTH: 200px" value="<%=strServCity%>" ></TD>
	</TR>
	<TR>
	    <TD width=15% align=right nowrap></TD>
    	<TD> </TD>
	    <TD width=15% align=right nowrap>Order No.</TD>
        <TD><INPUT id=txtOrderNO name=txtOrderNo tabindex=11 style="HEIGHT: 22px; WIDTH: 200px" ></TD></TR>
    <TR>
    	<TD>&nbsp;</TD>
	    <TD>&nbsp;</TD>
	    <TD width=15% align=right nowrap></TD>
		<TD align=left nowrap>Active&nbsp;Only&nbsp;<INPUT id=chkActiveOnly name=chkActiveOnly tabindex=12 type=checkbox value=YES CHECKED style="HEIGHT: 24px; WIDTH: 24px">&nbsp;&nbsp;Pref'd Lang Only&nbsp;<INPUT id=chkPrefLangOnly name=chkPrefLangOnly tabindex=12 type=checkbox value=YES CHECKED style="HEIGHT: 24px; WIDTH: 24px"></TD>
    </TR>
    <TR>
	    <TD width=15% align=right nowrap></TD>
    	<TD> </TD>
    	<TD> </TD>
		<TD><nobr>
		<% if strWinName <> "Popup" then %>
			<INPUT id=btnAdd name=btnAdd type=button value=New style="width: 2cm"   LANGUAGE=javascript onclick="return btnAdd_onclick()">
		<% end if %>
			<INPUT id=btnClear name=btnClear type=button value=Clear style="width: 2cm"  LANGUAGE=javascript onclick="return btnClear_onclick()">
			<INPUT id=btnSearch name=btnSearch type=submit value=Search style="width: 2cm" > </nobr>
        </TD>
    </TR>
     <TR>
    	<TD>&nbsp;</TD>
	    <TD>&nbsp;</TD>
    </TR>
    </TABLE>
</FORM>
</BODY>
</HTML>
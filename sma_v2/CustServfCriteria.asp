<%@ Language=VBScript %>
<%Option Explicit
  on error resume next %>
<!--#include file="smaConstants.inc"-->
<!--#include file="smaProcs.inc"-->
<!--#include file = "databaseconnect.asp"-->
<!--

********************************************************************************************
* Page name:	CustServCriteria.asp
* Purpose:		To dynamically set the criteria to search for a Customer Service.
*				Results are displayed via CustServList.asp
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
       29-Jul-15   PSmith  Set Cookies in validation so the back key works
       03-Feb-16   PSmith  Don't pre-populate search criteria
********************************************************************************************
-->
<%
dim intAccessLevel
'dim objRsRegion ,objRsSupportGroup, objRsRepairPriority
'dim strWinName,strServiceEnd, strLANG, strServiceTypeID
dim objRsStatus, strSQL
dim strCustomerServiceId, strCustomerService, strServLocName, StrCustomerName
Dim  strServiceTypeName, strCustomerID, strCustomerShortName


intAccessLevel = CInt(CheckLogon(strConst_CustomerService))
if (intAccessLevel and intConst_Access_ReadOnly) <> intConst_Access_ReadOnly then
	DisplayError "BACK", "", 0, "ACCESS DENIED", "You don't have access to view customer service. Please contact your system administrator."
end if

	'strLANG = Request.Cookies("UserInformation")("language_preference")
	'if (Len(strLANG) = 0) then strLANG = "EN"

	strCustomerServiceId = Request("txtCustomerServiceID")
	strCustomerService = Request("txtCustomerServiceDesc")
	strServiceTypeName = Request("txtServiceType")
    strServLocName = Request("txtServiceLocationName")


	'strServiceTypeID = Request.Cookies("ServiceTypeID")


	'The following 3 values are returned from CustList.asp
	strCustomerID = Request("hdnCustomerID")
	strCustomerName = Request("txtCustomerName")
	strCustomerShortName = Request.Cookies("CustomerShortName")

	'get a list of region codes
	'strSQL = "select noc_region_lcode, " &_
	'				"noc_region_desc " &_
	'		 "from crp.lcode_noc_region " &_
	'		 "where record_status_ind = 'A' " &_
	'		 "order by noc_region_desc"

	'set objRsRegion = objConn.Execute(strSQL)

	'get a list of service status codes
	strSQL = "SELECT service_status_code, " &_
					"service_status_name " &_
			 "FROM crp.service_status " &_
			 "WHERE record_status_ind = 'A' " &_
			 "order by service_status_name "

	set objRsStatus = objConn.Execute(strSQL)

	'get a list of support groups
	'strSQL = "SELECT remedy_support_group_id, " &_
	'				"group_name " &_
	'		  "FROM crp.v_remedy_support_group " &_
	'		  "order by group_name"

	'set objRsSupportGroup = objConn.Execute(strSQL)


	'get the LYNX repair priority list

	'strSQL = "SELECT lynx_def_sev_lcode " &_
	'         ",      lynx_def_sev_desc " &_
	'		 "FROM crp.lcode_lynx_def_sev " &_
	'		 "WHERE record_status_ind='A' "  &_
	'		 "ORDER BY lynx_def_sev_lcode"

	'set objRsRepairPriority = objConn.Execute(strSQL)


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

function validate(theForm){

	var bolConfirm ;
	if (isWhitespace(theForm.txtCustomerServiceDesc.value)
//	    && isWhitespace(theForm.selSupportGroup.value)
//	   && isWhitespace(theForm.selRepairPriority.value)
		&& isWhitespace(theForm.txtCustomerName.value)&& isWhitespace(theForm.selStatus.value)
		&& isWhitespace(theForm.txtCustomerServiceID.value) && isWhitespace(theForm.selRegion.value)
 		&& isWhitespace(theForm.txtServiceLocationName.value)
// 		&& isWhitespace(theForm.txtOrderNo.value)
 		&& isWhitespace(theForm.txtServiceType.value)
//      && isWhitespace(theForm.txtServiceCity.value)
//		&& isWhitespace(theForm.txtServiceAddress.value)&& isWhitespace(theForm.txtCustomerShortName.value)
		&& isWhitespace(theForm.txtCustomerID.value))
	{
		bolConfirm = window.confirm("No Search Criteria have been entered. This search may take a long time..Continue?");
		if (!bolConfirm){
			return false;
		}
	}

	if ( isNaN(Number(theForm.txtCustomerServiceID.value))) {
		alert("Customer Service ID must be a number");
		document.frmFCSSearch.txtCustomerServiceID.focus();
		document.frmFCSSearch.txtCustomerServiceID.select();
		return(false) ;
	}
	else if ( isNaN(Number(theForm.txtCustomerID.value))) {
		alert("Customer ID must be a number");
		document.frmFCSSearch.txtCustomerID.focus();
		document.frmFCSSearch.txtCustomerID.select();
		return(false) ;
	}

  // Start thinking
  thinking(parent.fraResult);

   return true;

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
	var strCustomerName = document.frmFCSSearch.txtCustomerName.value;

	//alert (CustService);

 	if (CustService != ""){
 		SetCookie("ServiceEnd",CustService);
		//alert (CustService);
 		}

	if (strCustomerName != "" ) {
		SetCookie("CustomerName", strCustomerName ) ;
		//alert (strCustomerName);
		}

	/*if (document.frmFCSSearch.txtCustomerName.value != "")
		{SetCookie("CustomerName", document.frmFCSSearch.txtCustomerName.value);
		}*/

	SetCookie("WinName", 'Popup');
	window.open('SearchFrame.asp?fraSrc=Cust', 'Popup', 'top=50, left=100, height=600, width=800' ) ;
	//document.frmFCSSearch.txtCustomerName.value = txtCustomerName ;
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
	var strCustomerService = document.frmFCSSearch.txtCustomerServiceDesc.value;
	var strCustomerName = window.frmFCSSearch.txtCustomerName.value ;
	var strServiceLocationName = window.frmFCSSearch.txtServiceLocationName.value;
	var intCustomerServiceID = document.frmFCSSearch.txtCustomerServiceID.value;
	var strServiceTypeName = document.frmFCSSearch.txtServiceType.value;
  var strWinName = document.frmFCSSearch.hdnWinName.value;

	DeleteCookie("WinName");
	DeleteCookie("CustomerName");
	DeleteCookie("CustomerService");
	DeleteCookie("ServiceEnd");
	DeleteCookie("ServLocName");
	DeleteCookie("CustomerServiceID");
	//DeleteCookie("ServiceTypeID");
	DeleteCookie("ServiceTypeName");
	DeleteCookie("hdnCustomerID");
	DeleteCookie("CustomerShortName");


	if (strWinName == "Popup" && (( strCustomerName != "" ) || ( strCustomerService != "" ) || (strServiceTypeName != "" )))

	{
		SetCookie("CustomerName",document.frmFCSSearch.txtCustomerName.value);
		SetCookie("CustomerService",document.frmFCSSearch.txtCustomerServiceDesc.value);
		SetCookie("ServiceTypeName",document.frmFCSSearch.txtServiceType.value);
    thinking(parent.fraResult);
		document.frmFCSSearch.submit() ;
	}
}

function btnClear_onclick() {
  with (document.frmFCSSearch)
  {
		txtCustomerName.value = "" ;
		document.frmFCSSearch.selStatus.selectedIndex = 0 ;
		txtCustomerServiceDesc.value = "" ;
		txtServiceLocationName.value = "" ;
		txtCustomerServiceID.value = "" ;
		document.frmFCSSearch.txtServiceType.value = "" ;
		document.frmFCSSearch.txtCustomerShortName.value = "" ;
		document.frmFCSSearch.txtCustomerID.value = "" ;
		document.frmFCSSearch.hdnCustomerID.value = "" ;
  }

}


//-->
</SCRIPT>
</HEAD>
<BODY LANGUAGE=javascript onload="return window_onload()">

<FORM name=frmFCSSearch method=post action="CustServfList.asp" target="fraResult" onSubmit="return validate(this)">
<INPUT id=hdnCustomerID name=hdnCustomerID type=hidden value="<%=strCustomerID%>">

<TABLE>
	<thead><tr><td colspan=4 align=left>Customer Service (with features) Search</tr></td></thead>
    <TR>
        <TD width=15% align=right nowrap>Customer Service ID</TD>
        <TD><INPUT id=txtCustomerServiceID name=txtCustomerServiceID tabindex=1 style="HEIGHT: 22px; WIDTH: 200px" value="<%if strCustomerServiceId<>"" then response.write strCustomerServiceId else response.write "" end if%>" ></TD>

		<TD width= 15% align=right nowrap>Customer Name<font color=red></font></TD>
		<TD colspan=2 nowrap align=left>
			<INPUT  name=txtCustomerName tabindex=2 type=text style="WIDTH: 250px; height: 26px; color:silver;" readonly maxlength=50
				value="<%=unescape(strCustomerName)%>" align=left><%if strCustomerName <> null then Response.write routineHTMLString(strCustomerName) else Response.Write null end if%>
		    <INPUT  name=btnCustomerLookup type=button value=... LANGUAGE=javascript onclick="fct_lookupCustomer('E')"></TD>
   </TR>
	<TR>
        <TD width=15% align=right nowrap>Customer Service Name/Alias</TD>
        <TD width=20%><INPUT id=txtCustomerServiceDesc name=txtCustomerServiceDesc tabindex=2 style="HEIGHT: 22px; WIDTH: 270px" value="<%=unescape(strCustomerService)%>">
		<TD width=15% align=right nowrap>Customer Short Name</TD>
        <TD width=20% colspan=2 nowrap align=left>
			<INPUT id=txtCustomerShortName name=txtCustomerShortName tabindex=3 style="HEIGHT: 22px; WIDTH: 270px" value="<%=strCustomerShortName%>" ></TD>
	</TR>
	<TR>
        <TD width=15% align=right nowrap>Service Type</TD>
		<TD width=20%><INPUT id=txtServiceType name=txtServiceType tabindex=6 style="HEIGHT: 22px; WIDTH: 270px"
		    value="<% if strServiceTypeName <> "" then response.write strServiceTypeName else response.write "" end if%>"></TD>
        <TD width=15% align=right nowrap>Customer ID</TD>
        <TD><INPUT id=txtCustomerID name=txtCustomerID tabindex=4 style="HEIGHT: 22px; WIDTH: 200px" value="<%=strCustomerID%>" ></TD>
	</TR>
    <TR>
        <TD width=15% align=right nowrap>Service Location Name</TD>
        <TD width=20% ><INPUT id=txtServiceLocationName name=txtServiceLocationName tabindex=3 style="HEIGHT: 22px; WIDTH: 270px" value="<% if strServLocName<> "" then response.write strServLocName else response.write "" end if%>" ></TD>
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
	    <TD width=15% align=right nowrap></TD>
    	<TD> </TD>
    	<TD> </TD>
	</TR>
	<TR>
	    <TD width=15% align=right nowrap></TD>
    	<TD> </TD>
    	<TD> </TD>
		<TD><nobr>
			<INPUT id=btnClear name=btnClear type=button value=Clear style="width: 2cm"  LANGUAGE=javascript onclick="return btnClear_onclick()">
			<INPUT id=btnSearch name=btnSearch type=submit value=Search style="width: 2cm" > </nobr>
        </TD></TR>

    </TABLE>
</FORM>
</BODY>
</HTML>
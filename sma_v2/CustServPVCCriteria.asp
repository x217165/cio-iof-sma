<%@ Language=VBScript %>
<%Option Explicit
  on error resume next %>
<!--#include file="smaConstants.inc"-->
<!--#include file="smaProcs.inc"-->
<!--#include file = "databaseconnect.asp"-->
<!--

********************************************************************************************
* Page name:	CustServPVCCriteria.asp
* Purpose:		To dynamically set the criteria to search for a Customer Service.
*				Results are displayed via CustServPVCList.asp
*
* In Param:		This form reads following cookies :
*				 - WinName
*				 - CustomerName
*				 - CustomerService
*				 - ServiceEnd
*
* Created by:	(Original CustServCriteria) Sara Sangha	Aug. 14th, 2000
* Edited by:    Adam Haydey Jan. 25th, 2000
*               Added Customer Service City and  Customer Service Address search fields.
********************************************************************************************
***************************************************************************************************

		 Date		Author			Changes/enhancements made
		12-Mar-01	A Haydey		Created the CustServPVCCriteria page as a separate page to allow for
									managed object name to be part of the search criteria and as well
									displaying project code when called from the PVC Detail screen.

***************************************************************************************************

-->
<%
dim intAccessLevel
dim objRsRegion ,objRsSupportGroup, objRsStatus, strSQL
dim strCustomerService, strWinName,strServiceEnd, strServLocName, StrCustomerName
Dim strServiceTypeID, strServiceTypeName, strManObjName

intAccessLevel = CInt(CheckLogon(strConst_CustomerService))
if (intAccessLevel and intConst_Access_ReadOnly) <> intConst_Access_ReadOnly then
	DisplayError "BACK", "", 0, "ACCESS DENIED", "You don't have access to view customer service. Please contact your system administrator."
end if



	strCustomerService = Request.Cookies("CustomerService")
	strWinName	= Request.Cookies("WinName")
    strServiceEnd = Request.Cookies("ServiceEnd")
    strServLocName = Request.Cookies("ServLocName")
    'Response.Write "end=" & strServiceEnd
	strCustomerName = Request.Cookies("CustomerName")
	strCustomerServiceId = Request.Cookies("CustomerServiceID")
	strServiceTypeID = Request.Cookies("ServiceTypeID")
	strServiceTypeName = Request.Cookies("ServiceTypeName")


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


%>

<HTML>
<HEAD>
<LINK rel="stylesheet" type="text/css" href="stylesheets/styles.css">
<TITLE>Customer Service Search for PVC</TITLE>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<META HTTP-EQUIV="Pragma" CONTENT="no-cache">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<script type="text/javascript" SRC="GeneralJavaFunctions.js"></script>
<script type="text/javascript" SRC="AccessLevels.js"></script>
<script type="text/javascript">
//****************************************** Java Functions *****************************
//set section title
var intAccessLevel = "<%=intAccessLevel%>" ;
setPageTitle("SMA - Customer Service (for PVC)");

function validate(theForm){

	var bolConfirm ;
	if (isWhitespace(theForm.txtCustomerServiceDesc.value) && isWhitespace(theForm.selSupportGroup.value)
		&& isWhitespace(theForm.txtCustomerName.value)&& isWhitespace(theForm.selStatus.value)
		&& isWhitespace(theForm.txtCustomerServiceID.value) && isWhitespace(theForm.selRegion.value)
		&& isWhitespace(theForm.txtServiceLocationName.value) && isWhitespace(theForm.txtOrderNo.value)
		&& isWhitespace(theForm.txtServiceType.value)  && isWhitespace(theForm.txtServiceCity.value)
		&& isWhitespace(theForm.txtServiceAddress.value) && isWhitespace(theForm.txtManObjName.value))
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

	if ( isNaN(Number(theForm.txtCustomerServiceID.value))) {
		alert("Customer Service ID must be a number");
		document.frmCSSearch.txtCustomerServiceID.focus();
		document.frmCSSearch.txtCustomerServiceID.select();
		return(false) ;
	}
	else
		{return(true);
	}

   return true;

}

function btnAdd_onclick(){

	if ((intAccessLevel & intConst_Access_Create) != intConst_Access_Create) {
		alert('Access denied. Please contact your system administrator.');
		return;
	}

	parent.location = 'CustServDetail.asp?CustServID=0'
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
	var strCustomerService = document.frmCSSearch.txtCustomerServiceDesc.value;
	var strCustomerName = window.frmCSSearch.txtCustomerName.value ;
	var strServiceLocationName = window.frmCSSearch.txtServiceLocationName.value;
	var intCustomerServiceID = document.frmCSSearch.txtCustomerServiceID.value;
	var strServiceTypeName = document.frmCSSearch.txtServiceType.value;
	var  strManObjName = document.frmCSSearch.txtManObjName.value;


	DeleteCookie("WinName");
	DeleteCookie("CustomerName");
	DeleteCookie("CustomerService");
	DeleteCookie("ServiceEnd");
	DeleteCookie("ServLocName");
	DeleteCookie("CustomerServiceID");
	DeleteCookie("ServiceTypeID");
	DeleteCookie("ServiceTypeName");

	if (( intCustomerServiceID != "" ) ||( strCustomerName != "" ) || ( strCustomerService != "" ) || (strServiceTypeName != "" ) || (strServiceLocationName != "" ))
	{
		document.frmCSSearch.submit() ;
	}
}

function btnClear_onclick() {
	document.frmCSSearch.txtCustomerName.value = "" ;
	document.frmCSSearch.selStatus.selectedIndex = 0 ;
	document.frmCSSearch.txtCustomerServiceDesc.value = "" ;
	document.frmCSSearch.selRegion.selectedIndex = 0 ;
	document.frmCSSearch.txtServiceLocationName.value = "" ;
	document.frmCSSearch.txtCustomerServiceID.value = "" ;
	document.frmCSSearch.txtServiceCity.value = "" ;
	document.frmCSSearch.txtServiceAddress.value = "" ;
	document.frmCSSearch.selSupportGroup.selectedIndex = 0 ;
	document.frmCSSearch.txtOrderNO.value = "" ;
	document.frmCSSearch.txtServiceType.value = "" ;
	document.frmCSSearch.txtManObjName.value = "" ;
	document.frmCSSearch.chkActiveOnly.checked=true;

}


//-->
</SCRIPT>
</HEAD>
<BODY LANGUAGE=javascript onload="return window_onload()">

<FORM name=frmCSSearch method=post action="CustServPVCList.asp" target="fraResult" onSubmit="return validate(this)">
	<INPUT type="hidden" id="hdnServiceTypeID" name="hdnServiceTypeID" value="<%=strServiceTypeID%>">
<TABLE>
	<thead><tr><td colspan=4 align=left>Customer Service Search</tr></td></thead>
    <TR>
		<TD width= 15% align=right nowrap>Customer Name</TD>
        <TD width=20% ><INPUT id=txtCustomerName name=txtCustomerName tabindex=1 style="HEIGHT: 22px; WIDTH: 270px" value="<%=strCustomerName%>" ></TD>
		<TD width=15% align=right nowrap > Region </TD>
        <TD><SELECT id=selRegion name=selRegion tabindex=7 style="HEIGHT: 22px; WIDTH: 200px" >
				<OPTION value = " " selected>
				<% Do while not objRsRegion.EOF %>
				<OPTION VALUE = "<%=objRsRegion(0)%>" > <%=objRsRegion(1)%>
				<%objRsRegion.MoveNext%>
				<%Loop%>
			</SELECT></TD>
   </TR>
    <TR>
        <TD width=15% align=right nowrap>Customer Service Name</TD>
        <TD width=20%><INPUT id=txtCustomerServiceDesc name=txtCustomerServiceDesc tabindex=2 style="HEIGHT: 22px; WIDTH: 270px" value="<%=unescape(strCustomerService)%>">
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
        <TD width=15% align=right nowrap>Service Location Name</TD>
        <TD width=20% ><INPUT id=txtServiceLocationName name=txtServiceLocationName tabindex=3 style="HEIGHT: 22px; WIDTH: 270px" value="<%=strServLocName%>" ></TD>
        <TD width=15% align=right nowrap>Customer Service ID</TD>
        <TD><INPUT id=txtCustomerServiceID name=txtCustomerServiceID tabindex=9 style="HEIGHT: 22px; WIDTH: 200px" value="<%=strCustomerServiceId%>" ></TD>
    </TR>
	<TR>
		<TD width=15% align=right nowrap>Service Address </TD>
        <TD width=20% ><INPUT id=txtServiceAddress name=txtServiceAddress tabindex=4 style="HEIGHT: 22px; WIDTH: 270px" value="<%=strServAddress%>" ></TD>
        <TD width=15% align=right nowrap>Service Location City</TD>
        <TD><INPUT id=txtServiceCity name=txtServiceCity tabindex=10 style="HEIGHT: 22px; WIDTH: 200px" value="<%=strServCity%>" ></TD>
	</TR>
	<TR>
	    <TD width=15% align=right nowrap>Support Group</TD>
        <TD width=20% ><SELECT id=selSupportGroup name=selSupportGroup style="HEIGHT: 22px; WIDTH: 270px" tabindex=5>
			<OPTION value = " "selected>
			<% Do while not objRsSupportGroup.EOF %>
				<OPTION VALUE = "<%=objRsSupportGroup(0)%>" > <%=objRsSupportGroup(1)%>
				<%objRsSupportGroup.MoveNext%>
			<%Loop%>
			</SELECT></TD>
	    <TD width=15% align=right nowrap>Order No.</TD>
        <TD><INPUT id=txtOrderNO name=txtOrderNo tabindex=11 style="HEIGHT: 22px; WIDTH: 200px" ></TD></TR>
    <TR>
		<TD width=15% align=right nowrap>Service Type</TD>
		<TD width=20%><INPUT id=txtServiceType name=txtServiceType tabindex=6 style="HEIGHT: 22px; WIDTH: 270px" value="<%=strServiceTypeName%>"></TD>
		<TD width=15% align=right nowrap>Active Only</TD>
		<TD align=left><INPUT id=chkActiveOnly name=chkActiveOnly tabindex=12 type=checkbox value=YES CHECKED style="HEIGHT: 24px; WIDTH: 24px"></TD>
	</TR>
	<TR>
		<TD width=15% align=right nowrap>Managed Object Name</TD>
		<TD width=20%><INPUT id=txtManObjName name=txtManObjName tabindex=6 style="HEIGHT: 22px; WIDTH: 270px" value="<%=strManObjName%>"></TD>


	  <INPUT id=hdnWinName name=hdnWinName type=hidden value="<%=strWinName%>">
	  <INPUT id=hdnServiceEnd name=hdnServiceEnd type=hidden value="<%=strServiceEnd%>">

		<TD colSpan=2 align=right >
		<% if strWinName <> "Popup" then %>
			<INPUT id=btnAdd name=btnAdd type=button value=New style="width: 2cm"  tabindex=15 LANGUAGE=javascript onclick="return btnAdd_onclick()">&nbsp;&nbsp;
		<% end if %>
			<INPUT id=btnClear name=btnClear type=button value=Clear style="width: 2cm"  tabindex=13 LANGUAGE=javascript onclick="return btnClear_onclick()">&nbsp;&nbsp;
			<INPUT id=btnSearch name=btnSearch type=submit value=Search style="width: 2cm" tabindex=14> &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
        </TD></TR>
</TABLE>
</FORM>
</BODY>
</HTML>
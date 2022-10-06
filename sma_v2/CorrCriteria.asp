<%@ Language=VBScript %>
<% 
option explicit 
on error resume next

%>
<!--#include file="smaConstants.inc"-->
<!--#include file="smaProcs.inc"-->
<!--#include file="databaseconnect.asp" -->
<!--
*******************************************************************************************
 Page name:	CorrCriteria.asp
 Purpose:		To dynamically set the criteria to search for correlation records.
				Results are displayed via CorrList.asp

 Created by:	Sara Sangha	08/14/2000
 Last updated by: Daniel Nica 10/03/2000 
  
********************************************************************************************
        Date		Author			Changes/enhancements made
        -----		------		------------------------------------------------------
       28-Feb-02	 DTy		Add 'Alias' to 'Customer Service Name' field name.
       29-Jul-15   PSmith  Set Cookies in validation so the back key works
       03-Feb-16   PSmith  Don't pre-populate search criteria
********************************************************************************************
-->
<%
'check users access rights


dim intAccessLevel
intAccessLevel = CInt(CheckLogon(strConst_CorrelationCustomer))
if intAccessLevel and intConst_Access_ReadOnly <> intConst_Access_ReadOnly then
	DisplayError "BACK", "", 0, "ACCESS DENIED", "You don't have access to correlation management. Please contact your system administrator"
end if

dim sql,strObjectName,StrType, strCustomerName

strObjectName = Request.Cookies("ObjectName")
StrType = Request.Cookies("Type")
strCustomerName = Request.Cookies("CustomerName") 

'get the regions
sql = "select NOC_REGION_LCODE, NOC_REGION_DESC from CRP.LCODE_NOC_REGION where RECORD_STATUS_IND='A'"
dim rsRegion
set rsRegion = Server.CreateObject("ADODB.Recordset")
rsRegion.CursorLocation = adUseClient
rsRegion.Open sql, objConn
if err then
	DisplayError "BACK", "", err.Number, "CANNOT CREATE RECORDSET", err.Description
end if
'release the active connection, keep the recordset open
set rsRegion.ActiveConnection = nothing 

'get status list
sql = "select SERVICE_STATUS_CODE, SERVICE_STATUS_NAME " &_
		"from CRP.SERVICE_STATUS " &_
		"where RECORD_STATUS_IND = 'A' " &_
		"order by SERVICE_STATUS_NAME "
dim rsStatus
set rsStatus = Server.CreateObject("ADODB.Recordset")
rsStatus.CursorLocation = adUseClient
rsStatus.Open sql, objConn
if err then
	DisplayError "BACK", "", err.Number, "CANNOT CREATE RECORDSET", err.Description
end if
rsStatus.MoveFirst
'release the active connection, keep the recordset open
set rsStatus.ActiveConnection = nothing	

'get support group list
sql = "select REMEDY_SUPPORT_GROUP_ID, GROUP_NAME " &_
		"from CRP.V_REMEDY_SUPPORT_GROUP " &_
		"order by GROUP_NAME "
dim rsSupportGroup
set rsSupportGroup = Server.CreateObject("ADODB.Recordset")
rsSupportGroup.CursorLocation = adUseClient
rsSupportGroup.Open sql, objConn
if err then
	DisplayError "BACK", "", err.Number, "CANNOT CREATE RECORDSET", err.Description
end if
'release the active connection, keep the recordset open
set rsSupportGroup.ActiveConnection = nothing	


objConn.close
set objConn = nothing

%>

<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<META HTTP-EQUIV="Pragma" CONTENT="no-cache">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<LINK rel="stylesheet" type="text/css" href="stylesheets/styles.css">
<script type="text/javascript" SRC="GeneralJavaFunctions.js"></script>
<SCRIPT type = "text/javascript">
//set section title
setPageTitle("SMA - Correlation Management");

//other functions
function btnClear_onClick(){
	with (document.frmCorrSearch){
		txtCustomerName.value = "";
		txtCustomerServiceDesc.value = "";
		txtObjectName.value = "";
		txtCustServID.value = "";
		selRegion.selectedIndex = 0;
		selSupportGroup.selectedIndex = 0;
		selStatus.selectedIndex = 0;
		ckhActive.checked=true; 
		chkMO.checked=true;
		chkPVC.checked=false;
		chkRoot.checked=false;
	}
}

function frm_submit(that){
	with (document.frmCorrSearch) {
		if (isNaN(txtCustServID.value)) {
			alert("Customer Service ID must be a number"); 
			txtCustServID.focus();
			txtCustServID.select();
			return(false)
		}
		if ((txtObjectName.value != '')&&(!chkMO.checked)&&(!chkPVC.checked)&&(!chkRoot.checked)){
			alert("If you wish to search by a object name please select at least one object type.");
			return(false)
		}
	}
	
  thinking(parent.fraResult);
	
	return(true);
}

function window_onload() {

 var strObjectName = document.frmCorrSearch.txtObjectName.value ;
 var strCustomerName = document.frmCorrSearch.txtCustomerName.value ;
	
	DeleteCookie("ObjectName");
	DeleteCookie("Type") ;
	DeleteCookie("CustomerName");
	
//	if ( (strObjectName != "" || strCustomerName != "")) {
	if (strObjectName != "") {
//		  SetCookie("CustomerName", document.frmCorrSearch.txtCustomerName.value);
//		  SetCookie("ObjectName", document.frmCorrSearch.txtObjectName.value);
		  thinking(parent.fraResult);
		document.frmCorrSearch.submit();
	}  

}

function btnNew_onclick(){
	parent.document.location = 'corrdetail.asp?CustomerServiceID=';
}

//-->
</SCRIPT>
</HEAD>

<BODY LANGUAGE=javascript onload="return window_onload()">
<FORM name="frmCorrSearch" method="POST" action="CorrList.asp" target="fraResult" onSubmit="return frm_submit(this)">
<table border="0" width="100%">
<thead align=left>
  <tr>
    <td align=Left colSpan=4>Correlation Search</td>
  </tr>
</thead>
<tbody>
	<TR>
		<TD align=right>Customer Name</TD>
		<TD><INPUT name=txtCustomerName tabindex=1 size=40 maxlength=30  value="<%=strCustomerName%>"></TD>
		<TD align=right>Region</TD>
		<TD>
			<SELECT name=selRegion tabindex=8 tabindex=8 style="HEIGHT: 22px; WIDTH:225px">
				<option selected value="ALL"> </option>	
				<%
				while not rsRegion.EOF
					Response.Write "<option value='" & rsRegion(0) & "'>" & routineHtmlString(rsRegion(1)) &"</option>" & vbCrLf
					rsRegion.MoveNext
				wend
				rsRegion.Close
				%>		
			</SELECT>
		</TD>
	</TR>
	<TR>
		
		<TD align=right>Customer Service Name/Alias</TD>
		<TD><INPUT name="txtCustomerServiceDesc" tabindex=2 size=40 maxlength=80 value=""></TD>
		<TD align=right>Support Group</TD>
		<TD>
			<SELECT name=selSupportGroup style="HEIGHT: 22px; WIDTH:225px" tabindex=9>
				<option value="ALL"> </option>
				<%while not rsSupportGroup.EOF
					If rsSupportGroup(1) = "Current Customer" Then
						Response.Write "<option selected value='" & rsSupportGroup(0) & "'>" & routineHtmlString(rsSupportGroup(1)) & "</option>" & vbCrLf
					Else  
						Response.write "<option value='" & rsSupportGroup(0)& "'>" & routineHtmlString(rsSupportGroup(1)) & "</option>" & vbCrLf
					End If
					rsSupportGroup.MoveNext
				wend
				rsSupportGroup.Close
				%>
			</SELECT>
		</TD>
	</TR>
	<TR>
		<TD align=right>Customer Service ID</TD>
		<TD><INPUT name=txtCustServID size=10 maxlength=9  tabindex=3 value=""></TD>
		<TD align=right>Status</TD>
		<TD>
			<SELECT id=selStatus name=selStatus style="HEIGHT: 22px; WIDTH:225px" tabindex=10>
				<option value="ALL"> </option>
				<%while not rsStatus.EOF
					If rsStatus(1) = "Current Customer" Then
						Response.Write "<option selected value='" & rsStatus(0) & "'>" & routineHtmlString(rsStatus(1)) & "</option>" & vbCrLf
					Else  
						Response.write "<option value='" & rsStatus(0)& "'>" & routineHtmlString(rsStatus(1)) & "</option>" & vbCrLf
					End If
					rsStatus.MoveNext
				wend
				rsStatus.Close
				%>
			</SELECT>
		</TD>
	</TR>
	<TR>
		<TD align=right>Object Name or Root CSID</TD>
		<TD><INPUT name=txtObjectName size=40 maxlength=30 tabindex=4 value="<%=strObjectName%>"></TD>
		<TD align=right>Active Only</TD>
		<TD><INPUT type="checkbox" name=ckhActive checked tabindex=11></TD>
	</TR>
	<TR>
		<TD></TD>
		<TD>
			<INPUT type="checkbox" name=chkMO tabindex=5 <% if (StrType <> "CustServ") and (StrType <> "Facility") then Response.write  " checked  " end if %>>Managed Object&nbsp;
			<INPUT type="checkbox" name=chkPVC tabindex=6 <% if (StrType = "Facility") then Response.write  " checked  " end if %>>Facility/PVC&nbsp;
			<INPUT type="checkbox" name=chkRoot tabindex=7 <% if (StrType = "CustServ") then Response.write  " checked  " end if %>>Root Service&nbsp;
		</TD>
		<TD colSpan=2 align=right>
<!--			<INPUT name=btnAddNew type=button value=New style="width: 2cm" onclick="btnNew_onclick()">&nbsp;&nbsp; -->
			<INPUT name=btnClear tabindex=12 type=button value=Clear style="width: 2cm" onClick="btnClear_onClick();">&nbsp;&nbsp;
			<INPUT name=btnSearch tabindex=13 type=submit style="width: 2cm" value=Search>&nbsp;&nbsp;
		</TD>
	</TR>
</tbody>
</TABLE>
</FORM>
</BODY>
</HTML>

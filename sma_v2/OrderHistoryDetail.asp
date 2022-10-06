<%@ Language=VBScript %>
<% option explicit %>
<!--#include file="SmaConstants.inc"-->
<!--#include file="SmaProcs.inc"-->
<!--#include file="databaseconnect.asp"-->
<!--
***************************************************************************************************
* Name:		CustServList.asp i.e. Customer Service List
*
* Purpose:	This page reads users's search critiera and bring back a list of matching Customer 
*			Service records. 
*
* Created By:	Sara Sangha 08/01/00
***************************************************************************************************
-->

<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<LINK rel="stylesheet" type="text/css" href="stylesheets/styles.css">
<script type="text/javascript" SRC="GeneralJavaFunctions.js"></script>
<script type="text/javascript">

//**************************************** Java Functions *********************************
//set section title
setPageTitle("SMA - Order History");

function btnSearch_onclick(){
	document.location = 'SearchFrame.asp?fraSrc=OrderHistory' 

}

function selNavigate_onchange(){
//***********************************************************************************************
// Function:	selNavigate_onchange															*
//																								*
// Purpose:		To display the page selected by the user from Quick Navigation drop-down box.	*
//																								*	
// Created By:	Sara Sangha	Aug. 25th, 2000														*
//																								*	
// Updated By:																					*	
//***********************************************************************************************

 var strPageName = document.frmOrderHistoryDetail.selNavigate.item(document.frmOrderHistoryDetail.selNavigate.selectedIndex).value ;   
 var strCustomerID =  document.frmOrderHistoryDetail.hdnCustomerID.value ;  
 var strCustomerServiceID = document.frmOrderHistoryDetail.hdnCustomerServiceID.value ; 
 
 	   
	switch ( strPageName ) {
	
	case 'Cust' :
		document.frmOrderHistoryDetail.selNavigate.selectedIndex=0;
		self.location.href = 'CustDetail.asp?CustomerID=' + strCustomerID ; 
		break ; 
	
	case 'CustServ' :
		document.frmOrderHistoryDetail.selNavigate.selectedIndex=0;
		self.location.href = 'CustServDetail.asp?CustServID=' + strCustomerServiceID ; 
		break ;
		
	case 'DEFAULT' :
		// do nothing ;
	}
	
}
//***************************************** End of Java Functions *************************
</SCRIPT>
</HEAD>
<BODY>
<FORM name=frmOrderHistoryDetail>

<%
  
Dim  objRsOrderHistory, strSoDetailID, strCustServID, strSQL, strWhereClause, strServLocAddress
  
strSoDetailID = unescape(Request.QueryString("SoDetailID"))
strCustServID = Request.QueryString("CustServID")

'Response.Write (strSODetailID & "***" & strCustServID)

 if strSoDetailID <> 0 then
	StrSql = "select  c.customer_service_desc, " &_
				"c.customer_service_id, " &_
				"c.customer_id, " &_
				"c.service_location_id, " &_
				"v.so_detail_id, " &_
				"v.sales_status as order_detail_status, " &_
				"v.service_order_id, " &_
				"v.sequence_no, " &_
				"v.service_type, " &_
				"v.order_type, " &_
				"v.customer_name, " &_
				"v.site, " &_
				"to_char(v.customer_requested_due_date, 'MON-DD-YYYY') as customer_requested_due_date, " &_
				"v.region_queue, " &_
				"v.order_prime, " &_
				"to_char(v.order_accept_date, 'MON-DD-YYYY') as order_accept_date, " &_
				"to_char(v.scheduled_due_date, 'MON-DD-YYYY') as scheduled_due_date, " &_
				"to_char(v.order_complete_date, 'MON-DD-YYYY') as order_complete_date, " &_
				"to_char(v.sales_complete_date, 'MON-DD-YYYY') as order_detail_complete_date, " &_
				"v.order_status, " &_
				"v.billing_prime, " &_
				"to_char(v.billing_complete_date, 'MON-DD-YYYY') as billing_complete_date, " &_
				"v.fac_prime, " &_
				"to_char(v.fac_complete_date, 'MON-DD-YYYY') as fac_complete_date, " &_
				"v.design_prime, " &_
				"to_char(v.design_complete_date,'MON-DD-YYYY') as design_complete_date, " &_
				"v.custtest_prime, " &_
				"to_char(v.custtest_complete_date, 'MON-DD-YYYY') as custtest_complete_date, " &_
				"to_char(C.CREATE_DATE_TIME, 'MON-DD-YYYY') as distribution_date, " &_
				"sma_sp_userid.spk_sma_library.sf_get_full_username(C.CREATE_REAL_USERID) as distribution_prime " &_			
			"from so.v_ecops_order v, " &_
    			
			" crp.customer_service c, " &_
				"so.so_detail d " &_
			"where c.customer_service_id = d.customer_service_id " &_
			"and d.so_detail_id = v.so_detail_id " &_
			"and v.so_detail_id = " & strSoDetailID
			
' do not need csid as list passes detail id			"and c.customer_service_id = " & strCustServID 				
   
  'Response.Write strsql
 ' Response.End 
  	
   set objRsOrderHistory = objConn.Execute(strSQL)
   if not objRsOrderHistory.eof then

%>

<!-- Hidden variables -->
	<INPUT name=hdnCustomerServiceID type=hidden value="<%=objRsOrderHistory("customer_service_id")%>">
	<INPUT name=hdnCustomerID type=hidden value="<%=objRsOrderHistory("customer_id")%>">

	
<TABLE>
 <THEAD>

		<TD align=left>Order History</TD>
		<TD></TD>
		<TD></TD>
		<TD align=right> 
		<SELECT id=selNavigate name=selNavigate LANGUAGE=javascript onchange="return selNavigate_onchange()">
			<OPTION value='DEFAULT'>Quickly Goto...</OPTION>
			<OPTION value=Cust>Customer</OPTION>
			<OPTION value=CustServ>Customer Service</OPTION>
		</SELECT> &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
		</TD>
	
 </THEAD>
 <TBODY>	
  <TR>
	<TD align=right>Customer Service Name</TD>
	<TD align=left><INPUT disabled name=txtCustomerServName style="WIDTH: 400px" value="<%=objRsOrderHistory("customer_service_desc")%>">
	<TD align=right>Order Number</TD>
    <TD align=left><INPUT disabled id=txtOrderNumber name=txtOrderNumber style="WIDTH: 100px" value="<%=objRsOrderHistory("service_order_id")%>" ></TD>
    <TD></TD></TR>
  <TR>	
	<TD align=right>Customer Name</TD>
	<TD align=left><INPUT disabled name=txtCustomerName style="WIDTH: 400px" value="<%=objRsOrderHistory("customer_name")%>">
	<TD align=right>Order Type</TD>
    <TD align=left><INPUT disabled id=txtOrderType name=txtOrderType style="WIDTH: 100px" value="<%=objRsOrderHistory("order_type")%>"></TD>
    <TD></TD></TR>
  <TR>
	<TD align=right>Service Type</TD>
	<TD align=left><INPUT disabled id=txtServiceType name=txtServiceType style="WIDTH: 400px" value="<%=objRsOrderHistory("service_type")%>"></TD>
    <TD align=right>Sequence No</TD>
    <TD align=left><INPUT disabled id=txtSeqNo name=txtSeqNo style="WIDTH: 100px" value="<%=objRsOrderHistory("sequence_no")%>"></TD></TR>
    <TD></TD></TR>
		
  <TR>
    <TD align=right>Site</TD>
    <TD align=left><INPUT disabled id=txtSite name=txtSite style="WIDTH: 400px" value="<%=objRsOrderHistory("site")%>"></TD>
	<TD align=right>Customer Requested Due Date</TD>
	<TD align=left><INPUT disabled id=txtCustDueDate name=txtCustDueDate style="WIDTH: 100px" value="<%=objRsOrderHistory("customer_requested_due_date")%>"></TD>
	<TD></TD></TR>
  <TR>
    <TD align=right>Region Queue</TD>
	<TD align=left><INPUT disabled id=txtReqiongQueue name=txtReqiongQueue style="WIDTH: 400px" value="<%=objRsOrderHistory("region_queue")%>">
    <TD align=right>Scheduled Due Date</TD>
    <TD align=left><INPUT disabled id=txtScheduledDueDate name=txtScheduledDueDate style="WIDTH: 100px" value="<%=objRsOrderHistory("scheduled_due_date")%>"></TD></TR>
    <TD></TD></TR>
  <TR>
    <TD align=right>Order Prime</TD>
    <TD align=left><INPUT disabled id=txtOderPrime name=txtOderPrime  style="WIDTH: 400px" value="<%=objRsOrderHistory("order_prime")%>" ></TD>
    <TD align=right>Order Accepted Date</TD>
    <TD align=left><INPUT disabled id=txtOrderDate name=txtOrderDate style="WIDTH: 100px" value="<%=objRsOrderHistory("order_accept_date")%>"></TD></TR>
	<TD></TD></TR>
  <TR>
    <TD align=right>Billing Pime</TD>
    <TD align=left><INPUT disabled id=txtBillingPrime name=txtBillingPrime style="WIDTH: 400px" value="<%=objRsOrderHistory("billing_prime")%>"></TD>
    <TD align=right>Billing Complete Date</TD>
    <TD align=left><INPUT disabled id=txtBillingCompleteDate name=txtBillingCompleteDate style="WIDTH: 100px" value="<%=objRsOrderHistory("billing_complete_date")%>"></TD></TR>
	<TD></TD></TR>
  <TR>
    <TD align=right>Facility Prime</TD>
    <TD align=left><INPUT disabled id=txtFacPrime name=txtFacPrime style="WIDTH: 400px" value="<%=objRsOrderHistory("fac_prime")%>"></TD>
    <TD align=right>Facility Complete Date</TD>
    <TD align=left><INPUT disabled id=txtFacCompleteDate name=txtFacCompleteDate style="WIDTH: 100px" value="<%=objRsOrderHistory("fac_complete_date")%>"></TD></TR>
	<TD></TD></TR>
  <TR>
	<TD align=right >Design Prime</TD>
	<TD align=left><INPUT disabled id=txtDesignPrime name=txtDesignPrime style="WIDTH: 400px" value="<%=objRsOrderHistory("design_prime")%>"></TD>
	<TD align=right>Design Complete Date</TD>
	<TD align=left><INPUT disabled id=txtDesingCompleteDate name=txtDesingCompleteDate style="WIDTH: 100px" value="<%=objRsOrderHistory("design_complete_date")%>"></TD>  	</TR>
	<TD></TD></TR>
  <TR>
	<TD align=right>Customer Test Prime</TD>
	<TD align=left><INPUT disabled id=txtCustTestPrime name=txtDesignPrime style="WIDTH: 400px" value="<%=objRsOrderHistory("custtest_prime")%>"></TD>
	<TD align=right>Customer Test Complete Date</TD>
	<TD align=left><INPUT disabled id=txtCustTestCompleteDate name=txtDesingCompleteDate style="WIDTH: 100px" value="<%=objRsOrderHistory("custtest_complete_date")%>"></TD>  	</TR>
	<TD></TD></TR>
  <TR>
	<TD align=right>Order Status</TD>
	<TD align=left><INPUT disabled id=txtOrderStatus name=txtOrderStatus style="WIDTH: 400px" value="<%=objRsOrderHistory("order_status")%>"></TD>
	<TD align=right> Order Complete Date</TD>
	<TD align=left><INPUT disabled id=txtOrderCompleteDate name=txtOrderCompleteDate style="WIDTH: 100px" value="<%=objRsOrderHistory("order_complete_date")%>"></TD>
	<TD></TD></TR>
  <TR>
	<TD align=right>Order Detail Status</TD>
	<TD align=left><INPUT disabled id=txtOrderDetStatus name=txtOrderDetStatus style="WIDTH: 400px" value="<%=objRsOrderHistory("order_detail_status")%>"></TD>
	<TD align=right> Order Detail Complete Date</TD>
	<TD align=left><INPUT disabled id=txtOrderDetCompleteDate name=txtOrderDetCompleteDate style="WIDTH: 100px" value="<%=objRsOrderHistory("order_detail_complete_date")%>"></TD>
	<TD></TD></TR>
  <TR>
	<TD align=right>Distribution Prime</TD>
	<TD align=left><INPUT disabled id=txtDistributionPrime name=txtDistributionPrime style="WIDTH: 400px" value="<%=objRsOrderHistory("distribution_prime")%>"></TD>
	<TD align=right> Distribution Date</TD>
	<TD align=left><INPUT disabled id=txtDistributionDate name=txtDistributionDate style="WIDTH: 100px" value="<%=objRsOrderHistory("distribution_date")%>"></TD>
	<TD></TD></TR>
  <TR>
	<TD align=right colspan=5>
		<INPUT id=btnSearch name=btnSearch type=button value=Search LANGUAGE=javascript onclick="return btnSearch_onclick()"> &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
		</TD></TR>
 <TBODY>		
</TABLE>
	
	<% END IF
end if
		
	 'clean up ADO objects
	
		set objRsOrderHistory = nothing
		objConn.close
		set objConn = nothing	
	%>
	
</FORM>
</BODY>
</HTML>

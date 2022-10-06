<%@ Language=VBScript %>
<% option explicit %>
<!--% on error resume next%-->
<!--
********************************************************************************************
* Page name:	XLSSOList.asp                                                              *
* Purpose:		To generate the Service Order Validation Spreadsheet                       *
*				Parameters entered via XLSEntry.asp                                        *
*                                                                                          *
* Created by:	Dan S. Ty	04/01/2002                                                     *
*                                                                                          *
********************************************************************************************
*       Date		Author			Changes/enhancements made                              *
*       -----		------		------------------------------------------------------     *
*                                                                                          *
********************************************************************************************
-->
<!--#include file="smaConstants.inc"-->
<!--#include file="smaProcs.inc"-->
<!--#include file="databaseconnect.asp"-->
<%

'Check Access rights - check other locations in this page.
dim intAccessLevel
intAccessLevel = CInt(CheckLogon(strConst_ESDCleanup))
If (intAccessLevel And intConst_Access_ReadOnly) <> intConst_Access_ReadOnly  Then
	DisplayError "BACK", "", 0, "ACCESS DENIED", "You don't have access to ESD Cleanup functions. Please contact your system administrator"
End If

dim lngCustomerID, strCustomerName, bolActiveOnly, strXLSFile
dim rsXLSSOList, aList
dim strSQL
dim intPageNumber, intPageCount
dim strMyWinName, strBgColor

'get parameters
lngCustomerID   = Request("hdnCustomerID")
strCustomerName = Request("hdnCustomerName")
bolActiveOnly   = Request("chkActiveOnly")
strXLSFile      = Request("hdnExport")

'build query
strSQL = " SELECT " & _
  "     a.so_id" & _
  "   , a.order_type_lcode" & _
  "   , d.order_status_desc" & _
  "   , e.order_source_desc" & _
  "   , a.completion_date" & _
  "   , a.record_status_ind" & _
  "   , a.requesting_customer_id" & _
  "   , f1.customer_name" & _
  "   , f1.record_status_ind" & _
  "   , a.billing_customer_id" & _
  "   , f2.customer_name" & _
  "   , f2.record_status_ind" & _
  "   , a.mainstream_customer_id" & _
  "   , a.arbor_account_no" & _
  "   , a.prime_contact_id" & _
  "   , c1.contact_name" & _
  "   , c1.record_status_ind" & _
  "   , b.contact_id" & _
  "   , c2.contact_name" & _
  "   , c2.record_status_ind" & _
  "   , b.contact_type_lcode" & _
  "   , b.record_status_ind" & _
  "   , g.billing_customer_id" & _
  "   , f3.customer_name" & _
  "   , f3.record_status_ind" & _
  " FROM"  & _
  "   so.service_order a" & _
  " , so.so_header_contact b" & _
  " , crp.contact c1" & _
  " , crp.contact c2" & _
  " , so.lcode_order_status d" & _
  " , so.lcode_order_source e" & _
  " , crp.customer f1" & _
  " , crp.customer f2" & _
  " , crp.customer f3" & _
  " , so.so_detail g" & _
  " WHERE" & _
  "       (a.billing_customer_id = " & lngCustomerID & " OR a.requesting_customer_id = " & lngCustomerID & ")" & _
  "   AND a.so_id                  = b.so_id" & _
  "   AND a.so_id                  = g.so_id (+)" & _
  "   AND a.order_status_lcode     = d.order_status_lcode" & _
  "   AND a.order_source_lcode     = e.order_source_lcode" & _
  "   AND a.prime_contact_id       = c1.contact_id  (+)" & _
  "   AND b.contact_id             = c2.contact_id  (+)" & _
  "   AND a.requesting_customer_id = f1.customer_id (+)" & _
  "   AND a.billing_customer_id    = f2.customer_id (+)" & _
  "   AND g.billing_customer_id    = f3.customer_id (+)"

If bolActiveOnly = "yes" Then
   strSQL = strSQL & " AND a.record_status_ind = 'A'"
End If

strSQL = strSQL & " ORDER by 1"

'get the recordset
set rsXLSSOList=server.CreateObject("ADODB.Recordset")
rsXLSSOList.Open strSQL, objConn

If err then
	DisplayError "BACK", "", err.Number, "XLSSOList.asp - Cannot open database" , err.Description
End if

'put recordset into array
if not rsXLSSOList.EOF then
	aList = rsXLSSOList.GetRows
else
	Response.Write "0 Record Found"
	Response.End
end if

'release and kill the recordset and the connection objects
rsXLSSOList.Close
set rsXLSSOList = nothing
objConn.Close
set objConn = nothing

'Create the validation spreadsheet
if strXLSFile = "" and request("action") = "" then
	'get real userid
	dim strRealUserID
	strRealUserID = Session("username")
	'determine export path
	dim strExportPath, liLength
	strExportPath = Request.ServerVariables("PATH_TRANSLATED")
	While (Right(strExportPath, 1) <> "\" And Len(strExportPath) <> 0)
		liLength = Len(strExportPath) - 1
		strExportPath = Left(strExportPath, liLength)
	Wend
	strExportPath = strExportPath & "export\save\"

	'create scripting object
	dim objFSO, objTxtStream
	set objFSO = server.CreateObject("Scripting.FileSystemObject")

	'create export file and save for future use.
	strXLSFile =  "CustID" & request("hdnCustomerID") & "-SO-" & year(now())  & "-" & month(now()) & "-" & day(now()) & "-" & hour(now()) & "-" & minute(now()) & "-" & second(now()) & "-" & strRealUserID & ".xls"
	set objTxtStream = objFSO.CreateTextFile(strExportPath & strXLSFile, false, false)
	if err then
		DisplayError "CLOSE", "", err.Number, "XLSSOList.asp - Cannot create Excel spreadsheet file due to the following reasons. Please contact your website administrator.", err.Description
	end if

	with objTxtStream
		.WriteLine "<table border=1>"

		'export the header
		.WriteLine "<THEAD>"
		.WriteLine "<TH>Service Order ID</TH>"
		.WriteLine "<TH>Order Type</TH>"
		.WriteLine "<TH>Order Status</TH>"
		.WriteLine "<TH>Order Source</TH>"
		.WriteLine "<TH>Order Completion Date</TH>"
		.WriteLine "<TH>Order Record Status</TH>"
		.WriteLine "<TH>REQ Customer ID</TH>"
		.WriteLine "<TH>REQ Customer Name</TH>"
		.WriteLine "<TH>REQ Record Status</TH>"
		.WriteLine "<TH>BILL Customer ID</TH>"
		.WriteLine "<TH>BILL Customer Name</TH>"
		.WriteLine "<TH>BILL Record Status</TH>"
		.WriteLine "<TH>MAINSTREAD Customer ID</TH>"
		.WriteLine "<TH>ARBOR Accolunt No.</TH>"
		.WriteLine "<TH>PRIME Contact ID</TH>"
		.WriteLine "<TH>PRIME Contact Name</TH>"
		.WriteLine "<TH>PRIME Record Status</TH>"
		.WriteLine "<TH>CONTACT Contact ID</TH>"
		.WriteLine "<TH>CONTACT Contact Name</TH>"
		.WriteLine "<TH>CONTACT Record Status</TH>"
		.WriteLine "<TH>CONTACT Contact Type</TH>"
		.WriteLine "<TH>CONTACT Contact Type Record Status</TH>"
		.WriteLine "<TH>DETAIL Billing Customer ID</TH>"
		.WriteLine "<TH>DETAIL Billing Customer Name</TH>"
		.WriteLine "<TH>DETAIL Billing Customer Record Status</TH>"
		.WriteLine "</THEAD>"

		'export the body
		for k = 0 to UBound(aList, 2)
			.WriteLine "<TR>"
			.WriteLine "<TD NOWRAP>" & routineHtmlString(aList(0,k))  & "&nbsp;</TD>" & vbCrLf
			.WriteLine "<TD NOWRAP>" & routineHtmlString(aList(1,k))  & "&nbsp;</TD>" & vbCrLf
			.WriteLine "<TD NOWRAP>" & routineHtmlString(aList(2,k))  & "&nbsp;</TD>" & vbCrLf
			.WriteLine "<TD NOWRAP>" & routineHtmlString(aList(3,k))  & "&nbsp;</TD>" & vbCrLf
			.WriteLine "<TD NOWRAP>" & routineHtmlString(aList(4,k))  & "&nbsp;</TD>" & vbCrLf
			.WriteLine "<TD NOWRAP>" & routineHtmlString(aList(5,k))  & "&nbsp;</TD>" & vbCrLf
			.WriteLine "<TD NOWRAP>" & routineHtmlString(aList(6,k))  & "&nbsp;</TD>" & vbCrLf
			.WriteLine "<TD NOWRAP>" & routineHtmlString(aList(7,k))  & "&nbsp;</TD>" & vbCrLf
			.WriteLine "<TD NOWRAP>" & routineHtmlString(aList(8,k))  & "&nbsp;</TD>" & vbCrLf
			.WriteLine "<TD NOWRAP>" & routineHtmlString(aList(9,k))  & "&nbsp;</TD>" & vbCrLf
			.WriteLine "<TD NOWRAP>" & routineHtmlString(aList(10,k)) & "&nbsp;</TD>" & vbCrLf
			.WriteLine "<TD NOWRAP>" & routineHtmlString(aList(11,k)) & "&nbsp;</TD>" & vbCrLf
			.WriteLine "<TD NOWRAP>" & routineHtmlString(aList(12,k)) & "&nbsp;</TD>" & vbCrLf
			.WriteLine "<TD NOWRAP>" & routineHtmlString(aList(13,k)) & "&nbsp;</TD>" & vbCrLf
			.WriteLine "<TD NOWRAP>" & routineHtmlString(aList(14,k)) & "&nbsp;</TD>" & vbCrLf
			.WriteLine "<TD NOWRAP>" & routineHtmlString(aList(15,k)) & "&nbsp;</TD>" & vbCrLf
			.WriteLine "<TD NOWRAP>" & routineHtmlString(aList(16,k)) & "&nbsp;</TD>" & vbCrLf
			.WriteLine "<TD NOWRAP>" & routineHtmlString(aList(17,k)) & "&nbsp;</TD>" & vbCrLf
			.WriteLine "<TD NOWRAP>" & routineHtmlString(aList(18,k)) & "&nbsp;</TD>" & vbCrLf
			.WriteLine "<TD NOWRAP>" & routineHtmlString(aList(19,k)) & "&nbsp;</TD>" & vbCrLf
			.WriteLine "<TD NOWRAP>" & routineHtmlString(aList(20,k)) & "&nbsp;</TD>" & vbCrLf
			.WriteLine "<TD NOWRAP>" & routineHtmlString(aList(21,k)) & "&nbsp;</TD>" & vbCrLf
			.WriteLine "<TD NOWRAP>" & routineHtmlString(aList(22,k)) & "&nbsp;</TD>" & vbCrLf
			.WriteLine "<TD NOWRAP>" & routineHtmlString(aList(23,k)) & "&nbsp;</TD>" & vbCrLf
			.WriteLine "<TD NOWRAP>" & routineHtmlString(aList(24,k)) & "&nbsp;</TD></TR>" & vbCrLf
		next
			.WriteLine "</table>"
		end with
		objTxtStream.Close
		Set objTxtStream = Nothing
		Set objFSO = Nothing
end if

'calculate page number
intPageCount = Int(UBound(aList,2) / intConstDisplayPageSize) + 1
select case Request("Action")
	case "<<"	intPageNumber = 1
	case "<"	intPageNumber = Request("txtPageNumber")-1
				if intPageNumber < 1 then intPageNumber = 1
	case ">"	intPageNumber = Request("txtPageNumber") + 1
				if intPageNumber > intPageCount then intPageNumber = intPageCount
	case ">>"	intPageNumber=intPageCount
	case else	if Request("hdnExport") <> "" then
						strsql = "<script type=""text/javascript"">document.location=""export/save/" & strXLSFile & """;</script>"
						Response.Write strsql
						Response.End
				'Response.redirect "export/"&strRealUserID&"-customer.xls"
				elseif Request("txtGoToPageNo") <> "" then
					intPageNumber = CInt(Request("txtGoToPageNo"))
				else
					intPageNumber = 1
				end if
end select

	if intPageNumber < 1 then intPageNumber = 1
	if intPageNumber > intPageCount then intPageNumber = intPageCount

	dim k,m,n
	m = (intPageNumber - 1) * intConstDisplayPageSize
	n = (intPageNumber) * intConstDisplayPageSize - 1
	if n > UBound(aList,2) then
		n=UBound(aList,2)
	end if

	'check if the client is still connected just before sending any html to the browser
	if Response.IsClientConnected = false then
		Response.End
	end if

	'catch any unexpected error
	if err then
		DisplayError "BACK", "", err.Number, "Unexpected error.", err.Description
	end if
%>

<HTML>
<HEAD>
	<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
	<LINK rel="stylesheet" type="text/css" href="stylesheets/styles.css" type="text/css">
	<title>Service Management Application</title>
	<script type="text/javascript" SRC="GeneralJavaFunctions.js"></script>

	<script ID=clientEventHandlersJS type="text/javascript">
	<!--
setPageTitle("SMA - Validation Spreadsheets");

	//-->
	</SCRIPT>

</head>
<body>
<form name=frmXLSSOList action="XLSSOList.asp" method=post>
	<input type=hidden name=hdnCustomerID   value="<%=lngCustomerID%>">
	<input type=hidden name=hdnCustomerName value="<%=strCustomerName%>">
	<input type=hidden name="hdnExport"     value="<%=strXLSFile%>">
	<input type=hidden name=chkActiveOnly   value="<%=bolActiveOnly%>">
<TABLE border=1 cellPadding=2 cellSpacing=0 width="100%">
	<THEAD>
	<TR>
		<TH align=left nowrap>Service Order ID</TH>
		<TH align=left nowrap>Order Type</TH>
		<TH align=left nowrap>Order Status</TH>
		<TH align=left nowrap>Order Source</TH>
		<TH align=left nowrap>Order Completion Date</TH>
		<TH align=left nowrap>Order Record Status</TH>
		<TH align=left nowrap>REQ Customer ID</TH>
		<TH align=left nowrap>REQ Customer Name</TH>
		<TH align=left nowrap>REQ Record Status</TH>
		<TH align=left nowrap>BILL Customer ID</TH>
		<TH align=left nowrap>BILL Customer Name</TH>
		<TH align=left nowrap>BILL Record Status</TH>
		<TH align=left nowrap>MAINSTREAM Customer ID</TH>
		<TH align=left nowrap>ARBOR Account No.</TH>
		<TH align=left nowrap>PRIME Contact ID</TH>
		<TH align=left nowrap>PRIME Contact Name</TH>
		<TH align=left nowrap>PRIME Record Status</TH>
		<TH align=left nowrap>CONTACT Contact ID</TH>
		<TH align=left nowrap>CONTACT Contact Name</TH>
		<TH align=left nowrap>CONTACT Record Status</TH>
		<TH align=left nowrap>CONTACT Contact Type</TH>
		<TH align=left nowrap>CONTACT Contact Type Record Status</TH>
		<TH align=left nowrap>DETAIL BILLING Customer ID</TH>
		<TH align=left nowrap>DETAIL BILLING Customer Name</TH>
		<TH align=left nowrap>DETAIL BILLING Customer Record Status</TH>
    </TR>
	</THEAD>
	<TBODY>
<%
'display the table
	for k = m to n

		'alternate row background color
		if Int(k/2) = k/2 then
			Response.Write "<tr bgcolor=White>"
		else
			Response.Write "<tr>"
		end if

		Response.Write "<td nowrap>" & aList(0,k)  & "&nbsp;</td>" & vbCrLf
		Response.Write "<td nowrap>" & aList(1,k)  & "&nbsp;</td>" & vbCrLf
		Response.Write "<td nowrap>" & aList(2,k)  & "&nbsp;</td>" & vbCrLf
		Response.Write "<td nowrap>" & aList(3,k)  & "&nbsp;</td>" & vbCrLf
		Response.Write "<td nowrap>" & aList(4,k)  & "&nbsp;</td>" & vbCrLf
		Response.Write "<td nowrap>" & aList(5,k)  & "&nbsp;</td>" & vbCrLf
		Response.Write "<td nowrap>" & aList(6,k)  & "&nbsp;</td>" & vbCrLf
		Response.Write "<td nowrap>" & aList(7,k)  & "&nbsp;</td>" & vbCrLf
		Response.Write "<td nowrap>" & aList(8,k)  & "&nbsp;</td>" & vbCrLf
		Response.Write "<td nowrap>" & aList(9,k)  & "&nbsp;</td>" & vbCrLf
		Response.Write "<td nowrap>" & aList(10,k) & "&nbsp;</td>" & vbCrLf
		Response.Write "<td nowrap>" & aList(11,k) & "&nbsp;</td>" & vbCrLf
		Response.Write "<td nowrap>" & aList(12,k) & "&nbsp;</td>" & vbCrLf
		Response.Write "<td nowrap>" & aList(13,k) & "&nbsp;</td>" & vbCrLf
		Response.Write "<td nowrap>" & aList(14,k) & "&nbsp;</td>" & vbCrLf
		Response.Write "<td nowrap>" & aList(15,k) & "&nbsp;</td>" & vbCrLf
		Response.Write "<td nowrap>" & aList(16,k) & "&nbsp;</td>" & vbCrLf
		Response.Write "<td nowrap>" & aList(17,k) & "&nbsp;</td>" & vbCrLf
		Response.Write "<td nowrap>" & aList(18,k) & "&nbsp;</td>" & vbCrLf
		Response.Write "<td nowrap>" & aList(19,k) & "&nbsp;</td>" & vbCrLf
		Response.Write "<td nowrap>" & aList(20,k) & "&nbsp;</td>" & vbCrLf
		Response.Write "<td nowrap>" & aList(21,k) & "&nbsp;</td>" & vbCrLf
		Response.Write "<td nowrap>" & aList(22,k) & "&nbsp;</td>" & vbCrLf
		Response.Write "<td nowrap>" & aList(23,k) & "&nbsp;</td>" & vbCrLf
		Response.Write "<td nowrap>" & aList(24,k) & "&nbsp;</td></tr>" & vbCrLf

   next
	%>
</TBODY>
<TFOOT>
<TR>
<TD align=left colSpan=25>
	<input type=hidden   name=hdnWinName    value="<%=strMyWinName%>">
	<input type=hidden   name=txtPageNumber value=<%=intPageNumber%>>
	<input type="submit" name="action" value="&lt;&lt;">
	<input type="submit" name="action" value="&lt;">
	<input type="text"   name="txtGoToPageNo" onClick="document.frmXLSSOList.txtGoToPageNo.value=''" title="You can jump to a specific page by typing the page number in this box" value="page <%=intPageNumber%> of <%=intPageCount%>" style="HEIGHT: 22px; WIDTH: 150px">
	<input type="submit" name="action" value="&gt;">
	<input type="submit" name="action" value="&gt;&gt;">&nbsp;&nbsp;
	<img SRC="images/excel.gif" onclick="document.frmXLSSOList.target='new';document.frmXLSSOList.hdnExport.value='<%=strXLSFile%>';document.frmXLSSOList.submit();document.frmXLSSOList.target='_self';" WIDTH="32" HEIGHT="32">
</TD>
</TR>
</TFOOT>
<CAPTION>Records <%=m+1%> to <%=n+1%> of <%=UBound(aList, 2)+1 & " records"%></CAPTION>
</table>
</form>
</body>
</html>









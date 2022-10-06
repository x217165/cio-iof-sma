<%@ Language=VBScript %>
<% option explicit
 on error resume next %>
<!--#include file="SmaConstants.inc"-->
<!--#include file="databaseconnect.asp"-->
<!--#include file="smaProcs.inc"-->

<!--
********************************************************************************************************
* Name:			OrderHistoryList.asp
*
* Purpose:		This page reads users's search critiera and bring back a list of matching SO_DETAIL records.
*
* Created By:	Sara Sangha Sept. 2nd, 2000
********************************************************************************************************
-->
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<LINK rel="stylesheet" type="text/css" href="stylesheets/styles.css">
</HEAD>
 <%

 dim aList, intPageNumber, intPageCount
 dim strCustomerServiceDesc, intCustomerServiceID
 dim strSQL, strWhereClause, strOrderBy

	strCustomerServiceDesc = UCase(trim(Request.Form("txtCustomerServiceName")))
	intCustomerServiceID = trim(Request.Form("txtCustomerServiceID"))

	strSQL = "select s.customer_service_id, " &_
					"s.customer_service_desc, " &_
					"v.service_order_id,  " &_
					"v.so_detail_id,  " &_
					"v.sequence_no, " &_
					"v.order_status, " &_
					"v.order_type, " &_
					"v.customer_name, " &_
					"v.site, " &_
					"to_char(s.CREATE_DATE_TIME, 'MON-DD-YYYY') as distribution_date, " &_
					"sma_sp_userid.spk_sma_library.sf_get_full_username(s.CREATE_REAL_USERID) as distribution_prime, " &_
					"to_char(v.order_complete_date, 'MON-DD-YYYY') as order_complete_date, " &_
					"v.sales_status as order_detail_status, " &_

 				     "to_char(v.sales_complete_date, 'MON-DD-YYYY') as order_detail_complete_date " &_
			"from   crp.customer_service s, " &_
					"so.so_detail d, " &_
					"so.v_ecops_order v	"

	strWhereClause = "where s.customer_service_id = d.customer_service_id " &_
 					 "and d.so_detail_id = v.so_detail_id "

	'add other search parameters to the where clause
	IF  LEN(strCustomerServiceDesc) > 0 THEN
      strWhereClause = strWhereClause & " AND UPPER(s.customer_service_desc) LIKE '" & routineOraString(strCustomerServiceDesc) &"%'"
	END IF

	IF  LEN(intCustomerServiceID) > 0 THEN
      strWhereClause = strWhereClause & " AND s.customer_service_id =" & intCustomerServiceID
	END IF


	'strOrderBy = " order by upper(s.customer_service_desc)"
	strOrderBy = " order by v.service_order_id, v.sequence_no "

	'join all pieces to make a complete query
	strsql = strSQL & strWhereClause & strOrderBy

	'Response.Write( strsql )       'display SQL for debugging
	'Response.End

	Dim objRsResult,Recordcnt,strbgcolor

	set objRsResult = objConn.Execute(StrSql)
	if not objRsResult.EOF then
		aList = objRsResult.GetRows
	else
		Response.Write "0 records found"
		Response.end
	end if

   'release and kill the recordset and the connection objects
	objRsResult.Close
	set objRsResult = nothing

	objConn.close
	set objConn = nothing

   'calculate page number
	intPageCount = Int(UBound(aList, 2) / intConstDisplayPageSize) + 1
	select case Request("Action")
		case "<<"		intPageNumber = 1
		case "<"		intPageNumber = Request("txtPageNumber") - 1
					if intPageNumber < 1 then intPageNumber = 1
		case ">"		intPageNumber = Request("txtPageNumber") + 1
					if intPageNumber > intPageCount then intPageNumber = intPageCount
		case ">>"		intPageNumber = intPageCount
		case else		if Request("hdnExport") <> "" then
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
						strExportPath = strExportPath & "export\"

						'create scripting object
						dim objFSO, objTxtStream
						set objFSO = server.CreateObject("Scripting.FileSystemObject")
						'create export file (overwrite if exists)
						set objTxtStream = objFSO.CreateTextFile(strExportPath&strRealUserID&".xls", true, false)

						if err then
							DisplayError "CLOSE", "", err.Number, ASP_NAME & " - Cannot create Excel spreadsheet file due to the following reasons.  Please contact your website administrator.", err.Description
						end if

						with objTxtStream
							.WriteLine "<table border=1>"

							'export the header
							.WriteLine "<TR>"
							.WriteLine "<TH>Customer Service Name</TD>"
							.WriteLine "<TH>Service ID</TD>"
							.WriteLine "<TH>Order No.</TD>"
							.WriteLine "<TH>Seq No.</TD>"
							.WriteLine "<TH>Order Status</TD>"
							.WriteLine "<TH>Order Type</TD>"
							.WriteLine "<TH>Customer Name</TD>"
							.WriteLine "<TH>Site</TD>"
							.WriteLine "<TH>Distribution Date</TD>"
							.WriteLine "<TH>Distribution Prime</TD>"
							.WriteLine "<TH>Compeletion Date</TD>"
							.WriteLine "<TH>Order Detail Status</TD>"
							.WriteLine "<TH>Ord. Det. Comp. Date</TD>"
							.WriteLine "<TD NOWRAP>&nbsp;&nbsp;&nbsp;</TD>"
							.WriteLine "</TR>"

							'export the body
							for k = 0 to UBound(aList, 2)
								.WriteLine "<TR>"
								.WriteLine "<TD NOWRAP>"&routineHtmlString(aList(1,k))&" &nbsp; </TD>"
								.WriteLine "<TD NOWRAP>"&routineHtmlString(aList(0,k))&" &nbsp; </TD>"
								.WriteLine "<TD NOWRAP>"&routineHtmlString(aList(2,k))&" &nbsp; </TD>"
								.WriteLine "<TD NOWRAP>"&routineHtmlString(aList(4,k))&" &nbsp;</TD>"
								.WriteLine "<TD NOWRAP>"&routineHtmlString(aList(5,k))&" &nbsp; </TD>"
								.WriteLine "<TD NOWRAP>"&routineHtmlString(aList(6,k))&" &nbsp; </TD>"
								.WriteLine "<TD NOWRAP>"&routineHtmlString(aList(7,k))&" &nbsp; </TD>"
								.WriteLine "<TD NOWRAP>"&routineHtmlString(aList(8,k))&" &nbsp; </TD>"
								.WriteLine "<TD NOWRAP>"&routineHtmlString(aList(9,k))&" &nbsp; </TD>"
								.WriteLine "<TD NOWRAP>"&routineHtmlString(aList(10,k))&" &nbsp; </TD>"
								.WriteLine "<TD NOWRAP>"&routineHtmlString(aList(11,k))&" &nbsp; </TD>"
								.WriteLine "<TD NOWRAP>"&routineHtmlString(aList(12,k))&" &nbsp; </TD>"
								.WriteLine "<TD NOWRAP>"&routineHtmlString(aList(13,k))&" &nbsp; </TD>"
								.WriteLine "<TD NOWRAP>&nbsp;&nbsp;&nbsp;</TD>"
								.WriteLine "</TR>"
							next
							.WriteLine "</table>"
						end with
						objTxtStream.Close
						strsql = "<script type=""text/javascript"">document.location=""export/"&strRealUserID&".xls"";</script>"
						Response.Write strsql
						Response.End
						Response.redirect "export/"&strRealUserID&".xls"



					elseif Request("txtGoToPageNo") <> "" then
						intPageNumber = CInt(Request("txtGoToPageNo"))
					else
						intPageNumber = 1
					end if
	end select

	if intPageNumber < 1 then
		intPageNumber = 1
	end if
	if intPageNumber > intPageCount then
		intPageNumber = intPageCount
	end if

	dim k, m, n
	m = (intPageNumber - 1 ) * intConstDisplayPageSize
	n = (intPageNumber) * intConstDisplayPageSize - 1
	if n > UBound(aList, 2) then
		n = UBound(aList, 2)
	end if

	'check if the client is still connected just before sending any html to the browser
	if response.isclientconnected = false then
		Response.End
	end if

	'catch any unexpected error
	if err then
		DisplayError "BACK", "", err.Number, "Unexpected error", err.Description
	end if

%>
<BODY>
<FORM method=post name=frmOrderHistory action="OrderHistoryList.asp">

    <input type=hidden name=txtCustomerServiceDesc value="<%=strCustomerServiceDesc%>">
    <input type=hidden name=txtCustomerServiceID value="<%=intCustomerServiceID%>">
    <input type=hidden name="hdnExport" value>

<TABLE  border=1 cellPadding=2 cellSpacing=0 width="100%">
  <THEAD>
    <TR>
        <TH>Customer Service Name</TH>
        <TH>Service ID</TH>
        <TH>Order No.</TH>
        <TH>Seq No.</TH>
        <TH>Order Status</TH>
        <TH>Order Type</TH>
        <TH>Customer Name</TH>
        <TH>Site</TH>
        <TH>Distribution Date</TH>
        <TH>Distribution Prime</TH>
        <TH>Completion Date</TH>
        <TH>Order Detail Status</TH>
        <TH>Ord. Det. Comp. Date</TH>
     </TR>
 </THEAD>
 <TBODY>
<%for k = m to n
	'Alternate row background colour
	if Int(k/2) = k/2 then
		Response.write "<TR>"
	else
		Response.write "<TR bgcolor=White>"
	end if

	Response.write "<TD nowrap><a TARGET=""_parent"" href=""OrderHistoryDetail.asp?SoDetailID="&aList(3,k)&"&CustServID="&aList(0,k)&""">"&routineHtmlString(aList(1,k))&"</a>&nbsp;</TD>" &vbCrLf
	Response.write "<TD nowrap><a TARGET=""_parent"" href=""OrderHistoryDetail.asp?SoDetailID="&aList(3,k)&"&CustServID="&aList(0,k)&""">"&routineHtmlString(aList(0,k))&"</a>&nbsp;</TD>" &vbCrLf
	Response.write "<TD nowrap><a TARGET=""_parent"" href=""OrderHistoryDetail.asp?SoDetailID="&aList(3,k)&"&CustServID="&aList(0,k)&""">"&routineHtmlString(aList(2,k))&"</a>&nbsp;</TD>" &vbCrLf
	Response.write "<TD nowrap><a TARGET=""_parent"" href=""OrderHistoryDetail.asp?SoDetailID="&aList(3,k)&"&CustServID="&aList(0,k)&""">"&routineHtmlString(aList(4,k))&"</a>&nbsp;</TD>" &vbCrLf
	Response.write "<TD nowrap><a TARGET=""_parent"" href=""OrderHistoryDetail.asp?SoDetailID="&aList(3,k)&"&CustServID="&aList(0,k)&""">"&routineHtmlString(aList(5,k))&"</a>&nbsp;</TD>" &vbCrLf
	Response.write "<TD nowrap><a TARGET=""_parent"" href=""OrderHistoryDetail.asp?SoDetailID="&aList(3,k)&"&CustServID="&aList(0,k)&""">"&routineHtmlString(aList(6,k))&"</a>&nbsp;</TD>" &vbCrLf
	Response.write "<TD nowrap><a TARGET=""_parent"" href=""OrderHistoryDetail.asp?SoDetailID="&aList(3,k)&"&CustServID="&aList(0,k)&""">"&routineHtmlString(aList(7,k))&"</a>&nbsp;</TD>" &vbCrLf
	Response.write "<TD nowrap><a TARGET=""_parent"" href=""OrderHistoryDetail.asp?SoDetailID="&aList(3,k)&"&CustServID="&aList(0,k)&""">"&routineHtmlString(aList(8,k))&"</a>&nbsp;</TD>" &vbCrLf
	Response.write "<TD nowrap><a TARGET=""_parent"" href=""OrderHistoryDetail.asp?SoDetailID="&aList(3,k)&"&CustServID="&aList(0,k)&""">"&routineHtmlString(aList(9,k))&"</a>&nbsp;</TD>" &vbCrLf
	Response.write "<TD nowrap><a TARGET=""_parent"" href=""OrderHistoryDetail.asp?SoDetailID="&aList(3,k)&"&CustServID="&aList(0,k)&""">"&routineHtmlString(aList(10,k))&"</a>&nbsp;</TD>" &vbCrLf
	Response.write "<TD nowrap><a TARGET=""_parent"" href=""OrderHistoryDetail.asp?SoDetailID="&aList(3,k)&"&CustServID="&aList(0,k)&""">"&routineHtmlString(aList(11,k))&"</a>&nbsp;</TD>" &vbCrLf
	Response.write "<TD nowrap><a TARGET=""_parent"" href=""OrderHistoryDetail.asp?SoDetailID="&aList(3,k)&"&CustServID="&aList(0,k)&""">"&routineHtmlString(aList(12,k))&"</a>&nbsp;</TD>" &vbCrLf
	Response.write "<TD nowrap><a TARGET=""_parent"" href=""OrderHistoryDetail.asp?SoDetailID="&aList(3,k)&"&CustServID="&aList(0,k)&""">"&routineHtmlString(aList(13,k))&"</a>&nbsp;</TD>" &vbCrLf
	Response.Write "</TR>"



next
%>
</TBODY>
<TFOOT>
<TR>
<TD align=left colSpan=13>
	<input type=hidden name=txtPageNumber value=<%=intPageNumber%>>
	<input type="submit" name="action" value="&lt;&lt;">
	<input type="submit" name="action" value="&lt;">
	<input type="text" name="txtGoToPageNo" title="You can jump to a specific page by typing the page number in this box" onclick="document.frmOrderHistory.txtGoToPageNo.value=''" value="page <%=intPageNumber%> of <%=intPageCount%>" style="HEIGHT: 22px; WIDTH: 150px">
	<input type="submit" name="action" value="&gt;">
	<input type="submit" name="action" value="&gt;&gt;">
	<img SRC="images/excel.gif" onclick="document.frmOrderHistory.target='new'; document.frmOrderHistory.hdnExport.value='xls';document.frmOrderHistory.submit();document.frmOrderHistory.hdnExport.value='';document.frmOrderHistory.target='_self';" WIDTH="32" HEIGHT="32">
</TD>
</TR>
</TFOOT>
<CAPTION>Records <%=m+1%> to <%=n+1%> of <%=UBound(aList, 2)+1 & " records"%></CAPTION>
</TABLE>
</FORM>
</BODY>
</HTML>

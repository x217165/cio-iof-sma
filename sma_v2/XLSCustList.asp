<%@ Language=VBScript %>
<% option explicit %>
<!--% on error resume next%-->
<!--
********************************************************************************************
* Page name:	XLSCustList.asp                                                            *
* Purpose:		To generate the Customer Validation Spreadsheet                            *
*				Parameters entered via XLSEntry.asp                                        *
*                                                                                          *
* Created by:	Dan S. Ty	03/31/2002                                                     *
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
dim rsXLSCustList, aList
dim strSQL
dim intPageNumber, intPageCount
dim strMyWinName, strBgColor

'get parameters
lngCustomerID   = Request("hdnCustomerID")
strCustomerName = Request("hdnCustomerName")
bolActiveOnly   = Request("chkActiveOnly")

dim strRealUserID
strRealUserID  =Session("username")
strXLSFile     = Request("hdnExport")

'build query
'List Customer that has Customer Care (CSM) assigned.
strSQL = "SELECT" & _
  "     TRIM(UPPER(REPLACE(TRANSLATE(" &_
  "     (REPLACE(     REPLACE(    REPLACE(   REPLACE(         REPLACE(         REPLACE( REPLACE(" &_
  "      REPLACE(     REPLACE(    REPLACE(   REPLACE(         REPLACE(         REPLACE( REPLACE(" &_
  "      REPLACE(     REPLACE(    REPLACE(   REPLACE(         REPLACE(         REPLACE( REPLACE(" &_
  "      UPPER(a.customer_name), " &_
  "      'THE '   ), ' THE '   ), 'CENTER' ), 'CENTRE'), 'CTR'         ), 'CORPORATION' ), 'XXX' )," &_
  "      'CORP'   ), '(CANADA)'), '(AB)'   ), '(BC)'  ), ' AB '        ), ' BC'         ), ' OF ')," &_
  "      'COMPANY'), 'CO.'     ), 'LIMITED'), 'LTD'   ), 'TECHNOLOGIES'), 'TECHNOLOGY'  ), 'WWW' )), '~`!@#$%^&*()_-+={}[]:,<>,.?/|\', '                              '), ' ')))" &_
  "   , a.customer_id, DECODE(a.record_status_ind, 'A', ' ', a.record_status_ind)" &_
  "   , a.customer_status_lcode, a.customer_name, a.customer_short_name, b.customer_name_alias_id" &_
  "   , b.customer_name_alias_upper, a.noc_region_lcode, d.contact_id" &_
  "   , DECODE(d.record_status_ind, 'A', '', d.record_status_ind)" &_
  "   , d.contact_name || ' ' || d.middle_name, DECODE(d.staff_flag, 'Y', d.staff_flag, ' ')" &_
  "   , DECODE(d.work_number_ext, null, d.work_number, d.work_number || ' Ext ' || d.work_number_ext)" &_
  "   , d.home_number, d.cell_number, d.pager_number, d.fax_number, d.email_address, a.comments" &_
  " FROM" &_
  "   crp.customer a, crp.customer_name_alias b, crp.customer_contact c, crp.contact d" &_
  " WHERE" & _
  "       a.customer_id = b.customer_id" & _
  "   AND b.customer_id = c.customer_id" & _
  "   AND c.contact_id = d.contact_id" & _
  "   AND (c.customer_contact_type_lcode is not null and upper(c.customer_contact_type_lcode) = 'CUSTCARE')" & _
  "   AND a.customer_id = " & lngCustomerID

If bolActiveOnly = "yes" then
   strSQL = strSQL & " AND a.record_status_ind = 'A'"
end if

'List Customer that has no Customer Care (CSM) assigned.
strSQL = strSQL & " UNION SELECT " & _
  "     TRIM(UPPER(REPLACE(TRANSLATE(" & _
  "       (REPLACE(     REPLACE(    REPLACE(   REPLACE(         REPLACE(         REPLACE( REPLACE(" & _
  "       REPLACE(     REPLACE(    REPLACE(   REPLACE(         REPLACE(         REPLACE( REPLACE(" & _
  "       REPLACE(     REPLACE(    REPLACE(   REPLACE(         REPLACE(         REPLACE( REPLACE(" & _
  "       UPPER(a.customer_name), " & _
  "       'THE '   ), ' THE '   ), 'CENTER' ), 'CENTRE'), 'CTR'         ), 'CORPORATION' ), 'XXX' )," & _
  "       'CORP'   ), '(CANADA)'), '(AB)'   ), '(BC)'  ), ' AB '        ), ' BC'         ), ' OF ')," & _
  "       'COMPANY'), 'CO.'     ), 'LIMITED'), 'LTD'   ), 'TECHNOLOGIES'), 'TECHNOLOGY'  ), 'WWW' )), '~`!@#$%^&*()_-+={}[]:,<>,.?/|\', '                              '), ' ')))" & _
  "     , a.customer_id, DECODE(a.record_status_ind, 'A', '', a.record_status_ind)" & _
  "     , a.customer_status_lcode, a.customer_name, a.customer_short_name, b.customer_name_alias_id" & _
  "     , b.customer_name_alias_upper, a.noc_region_lcode" & _
  "     , 0, null, null, null, null, null, null, null, null, null, a.comments" & _
  " FROM crp.customer a, crp.customer_name_alias b" & _
  " WHERE" & _
  "       a.customer_id = b.customer_id AND a.customer_id NOT IN" & _
  "       (SELECT DISTINCT a.customer_id" & _
  "          FROM crp.customer a, crp.customer_name_alias b, crp.customer_contact c, crp.contact d" & _
  "          WHERE a.customer_id = b.customer_id" & _
  "            AND b.customer_id = c.customer_id" & _
  "            AND c.contact_id = d.contact_id" & _
  "            AND (c.customer_contact_type_lcode IS NOT NULL AND upper(c.customer_contact_type_lcode) = 'CUSTCARE'))" & _
  "   AND a.customer_id = " & lngCustomerID

If bolActiveOnly = "yes" then
   strSQL = strSQL & " AND a.record_status_ind = 'A'"
end if

strSQL = strSQL & "  ORDER BY 1, 5, 8"

'get the recordset
set rsXLSCustList=server.CreateObject("ADODB.Recordset")
rsXLSCustList.Open strSQL, objConn

If err then
   DisplayError "BACK", "", err.Number, "XLSCustList.asp - Cannot open database" , err.Description
End if

'put recordset into array
if not rsXLSCustList.EOF then
   aList = rsXLSCustList.GetRows
else
   Response.Write "0 Record Found"
   rssponse.End
end if

'release and kill the recordset and the connection objects
rsXLSCustList.Close
set rsXLSCustList = nothing
objConn.Close
set objConn = nothing

'Create the validation spreadsheet
if strXLSFile = "" and request("action") = "" then
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
	strXLSFile =  "CustID" & request("hdnCustomerID") & "-Cust-" & year(now())  & "-" & month(now()) & "-" & day(now()) & "-" & hour(now()) & "-" & minute(now()) & "-" & second(now()) & "-" & strRealUserID & ".xls"

	set objTxtStream = objFSO.CreateTextFile(strExportPath & strXLSFile, false, false)
	if err then
		DisplayError "CLOSE", "", err.Number, "XLSCustList.asp - Cannot create Excel spreadsheet file due to the following reasons. Please contact your website administrator.", err.Description
	end if

	with objTxtStream
		.WriteLine "<table border=1>"

		'export the header
		.WriteLine "<THEAD>"
		.WriteLine "<TH>Customer ID</TH>"
		.WriteLine "<TH>Customer Record Status</TH>"
		.WriteLine "<TH>Status</TH>"
		.WriteLine "<TH>Customer Name</TH>"
		.WriteLine "<TH>Short Name</TH>"
		.WriteLine "<TH>Alias ID</TH>"
		.WriteLine "<TH>Alias</TH>"
		.WriteLine "<TH>Region</TH>"
		.WriteLine "<TH>Contact ID</TH>"
		.WriteLine "<TH>Contact Record Status</TH>"
		.WriteLine "<TH>Customer Service Manager</TH>"
		.WriteLine "<TH>TELUS Staff</TH>"
		.WriteLine "<TH>Work Phone</TH>"
		.WriteLine "<TH>Home Phone</TH>"
		.WriteLine "<TH>Cell Phone</TH>"
		.WriteLine "<TH>Pager</TH>"
		.WriteLine "<TH>Fax Number</TH>"
		.WriteLine "<TH>Email Address</TH>"
		.WriteLine "</THEAD>"

		'export the body
		for k = 0 to UBound(aList, 2)
			.WriteLine "<TR>"
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
			.WriteLine "<TD NOWRAP>" & routineHtmlString(aList(17,k)) & "&nbsp;</TD></TR>" & vbCrLf
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
						strSQL = "<script type=""text/javascript"">document.location=""export/save/" & strXLSFile & """;</script>"
						Response.Write strSql
						Response.End
				'Response.redirect "export/"&strRealUserID&"-customer.xls"
				elseif Request("txtGoToPageNo") <> "" and IsNumeric(Request("txtGoToPageNo"))then
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
<form name=frmXLSCustList action="XLSCustList.asp" method=post>
	<input type=hidden name=hdnCustomerID   value="<%=lngCustomerID%>">
	<input type=hidden name=hdnCustomerName value="<%=strCustomerName%>">
	<input type=hidden name=chkActiveOnly   value="<%=bolActiveOnly%>">
	<input type=hidden name="hdnExport"     value="<%=strXLSFile%>">
<TABLE border=1 cellPadding=2 cellSpacing=0 width="100%">
	<THEAD>
	<TR>
		<TH align=left nowrap>Customer ID</TH>
		<TH align=left nowrap>Customer Record Status</TH>
		<TH align=left nowrap>Status</TH>
		<TH align=left nowrap>Customer Name</TH>
		<TH align=left nowrap>Short Name</TH>
		<TH align=left nowrap>Alias ID</TH>
		<TH align=left nowrap>Alias</TH>
		<TH align=left nowrap>Region</TH>
		<TH align=left nowrap>Contact ID</TH>
		<TH align=left nowrap>Contact Record Status</TH>
		<TH align=left nowrap>Customer Service Manager</TH>
		<TH align=left nowrap>TELUS Staff</TH>
		<TH align=left nowrap>Work Phone</TH>
		<TH align=left nowrap>Home Phone</TH>
		<TH align=left nowrap>Cell Phone</TH>
		<TH align=left nowrap>Pager</TH>
		<TH align=left nowrap>Fax Number</TH>
		<TH align=left nowrap>Email Address</TH>
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
		Response.Write "<td nowrap>" & aList(18,k) & "&nbsp;</td></tr>" & vbCrLf

   next
	%>
</TBODY>
<TFOOT>
<TR>
<TD align=left colSpan=18>
	<input type=hidden   name=hdnWinName    value="<%=strMyWinName%>">
	<input type=hidden   name=txtPageNumber value=<%=intPageNumber%>>
	<input type="submit" name="action" value="&lt;&lt;">
	<input type="submit" name="action" value="&lt;">
	<input type="text"   name="txtGoToPageNo" onClick="document.frmXLSCustList.txtGoToPageNo.value=''" title="You can jump to a specific page by typing the page number in this box" value="page <%=intPageNumber%> of <%=intPageCount%>" style="HEIGHT: 22px; WIDTH: 150px">
	<input type="submit" name="action" value="&gt;">
	<input type="submit" name="action" value="&gt;&gt;">&nbsp;&nbsp;
	<img SRC="images/excel.gif" onclick="document.frmXLSCustList.target='new';document.frmXLSCustList.hdnExport.value='<%=strXLSFile%>';document.frmXLSCustList.submit();document.frmXLSCustList.target='_self';" WIDTH="32" HEIGHT="32">
</TD>
</TR>
</TFOOT>
<CAPTION>Records <%=m+1%> to <%=n+1%> of <%=UBound(aList, 2)+1 & " records"%></CAPTION>
</table>
</form>
</body>
</html>

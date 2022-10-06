<%@ Language=VBScript %>
<% option explicit %>
<!--% on error resume next%-->
<!--
********************************************************************************************
* Page name:	XLSCustServList.asp                                                        *
* Purpose:		To generate the Customer Service Validation Spreadsheet                    *
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
dim rsXLSCustServList, aList
dim strSQL
dim intPageNumber, intPageCount
dim strMyWinName, strBgColor

'get parameters
lngCustomerID   = Request("hdnCustomerID")
strCustomerName = Request("hdnCustomerName")
bolActiveOnly   = Request("chkActiveOnly")

strXLSFile     = Request("hdnExport")

'build query

'Extract Correlated/Uncorrelated Managed Objects, Customer Services and Service Locations.
strSQL = " SELECT " & _
  "     cu.customer_id, upper(cu.customer_name), sl.service_location_id, upper(sl.service_location_name)" & _
  "   , ne.network_element_id, upper(ne.network_element_name), ne.network_element_type_code, ne.network_element_desc" & _
  "   , cs.customer_service_id, cs.customer_service_desc, st.service_type_desc" & _
  "   , ad.address_id, ad.house_number, ad.building_name, ad.street_name, ad.municipality_name, ad.province_state_lcode" & _
  " FROM" & _
  "   crp.customer cu," & _
  "   crp.customer_name_alias ca," & _
  "   crp.network_element ne," & _
  "   crp.managed_correlation mc," & _
  "   crp.customer_service cs," & _
  "   crp.service_location sl," & _
  "   crp.service_type st," & _
  "   crp.address ad" & _
  " WHERE" & _
  "       cu.customer_id = ca.customer_id" & _
  "   AND ca.customer_id = ne.customer_id (+)" & _
  "   AND ne.network_element_id = mc.network_element_id (+)" & _
  "   AND mc.customer_service_id = cs.customer_service_id (+)" & _
  "   AND cs.service_location_id = sl.service_location_id (+)" & _
  "   AND cs.service_type_id = st.service_type_id (+)" & _
  "   AND sl.address_id = ad.address_id (+)" & _
  "   AND cu.customer_id = " & lngCustomerID

If bolActiveOnly = "yes" Then
   strSQL = strSQL & _
     " AND cs.record_status_ind = 'A'" & _
     " AND sl.record_status_ind = 'A'"
End If

'Extract Uncorrelated Customer Services with the corresponding Service Locations.
strSQL = strSQL & " UNION SELECT " & _
  "     cu.customer_id, upper(cu.customer_name), sl.service_location_id, upper(sl.service_location_name)" & _
  "   , 0 as network_element_id, null as network_element_name, null as network_element_type_code, null as network_element_desc" & _
  "   , cs.customer_service_id, cs.customer_service_desc, st.service_type_desc" & _
  "   , ad.address_id, ad.house_number, ad.building_name, ad.street_name, ad.municipality_name, ad.province_state_lcode" & _
  " FROM" & _
  "   crp.customer cu," & _
  "   crp.customer_name_alias ca," & _
  "   crp.customer_service cs," & _
  "   crp.service_location sl," & _
  "   crp.service_type st," & _
  "   crp.address ad" & _
  " WHERE" & _
  "       cu.customer_id = ca.customer_id" & _
  "   AND ca.customer_id = cs.customer_id (+)" & _
  "   AND cs.customer_service_id not in (SELECT DISTINCT customer_service_id FROM crp.managed_correlation)" & _
  "   AND cs.service_location_id = sl.service_location_id (+)" & _
  "   AND cs.service_type_id = st.service_type_id (+)" & _
  "   AND sl.address_id = ad.address_id (+)" & _
  "   AND cu.customer_id = " & lngCustomerID

If bolActiveOnly = "yes" Then
   strSQL = strSQL & _
     " AND cs.record_status_ind = 'A'" & _
     " AND sl.record_status_ind = 'A'"
End If

'Extract Uncorrelated Service Locations with the corresponding Customer Services.
strSQL = strSQL & " UNION SELECT " & _
  "     cu.customer_id, upper(cu.customer_name), sl.service_location_id, upper(sl.service_location_name)" & _
  "   , 0 as network_element_id, null as network_element_name, null as network_element_type_code, null as network_element_desc" & _
  "   , cs.customer_service_id, cs.customer_service_desc, st.service_type_desc" & _
  "   , ad.address_id, ad.house_number, ad.building_name, ad.street_name, ad.municipality_name, ad.province_state_lcode" & _
  " FROM" & _
  "   crp.customer cu," & _
  "   crp.customer_name_alias ca," & _
  "   crp.customer_service cs," & _
  "   crp.service_location sl," & _
  "	  crp.service_type st," & _
  "	  crp.address ad" & _
  " WHERE" & _
  "       cu.customer_id = ca.customer_id" & _
  "   AND ca.customer_id = cs.customer_id (+)" & _
  "   AND cs.customer_service_id in (SELECT DISTINCT customer_service_id FROM crp.managed_correlation)" & _
  "   AND cs.service_location_id not in (SELECT DISTINCT service_location_id FROM crp.network_element)" & _
  "   AND cs.service_location_id = sl.service_location_id (+)" & _
  "   AND cs.service_type_id = st.service_type_id (+)" & _
  "   AND sl.address_id = ad.address_id (+)" & _
  "   AND cu.customer_id = " & lngCustomerID

If bolActiveOnly = "yes" Then
   strSQL = strSQL & _
     " AND cs.record_status_ind = 'A'" & _
     " AND sl.record_status_ind = 'A'"
End If

'Extract Customer Services with no Service Locations.
strSQL = strSQL & " UNION SELECT " & _
  "     cu.customer_id, upper(cu.customer_name), 0 as service_location_id, null as service_location_name" & _
  "   , 0 as network_element_id, null as network_element_name, null as network_element_type_code, null as network_element_desc" & _
  "   , cs.customer_service_id, cs.customer_service_desc, st.service_type_desc" & _
  "   , 0 as address_id, 0 as house_number, null as building_name, null as street_name, null as municipality_name, null as province_state_lcode" & _
  " FROM" & _
  "   crp.customer cu," & _
  "   crp.customer_name_alias ca," & _
  "   crp.customer_service cs," & _
  "   crp.service_type st" & _
  " WHERE" & _
  "       cu.customer_id = ca.customer_id" & _
  "   AND ca.customer_id = cs.customer_id (+)" & _
  "   AND (cs.service_location_id is null OR cs.service_location_id is not null AND cs.service_location_id not in (SELECT DISTINCT service_location_id FROM crp.service_location))" & _
  "   AND cs.service_type_id = st.service_type_id (+)" & _
  "   AND cu.customer_id = " & lngCustomerID

If bolActiveOnly = "yes" Then
   strSQL = strSQL & _
     " AND cs.record_status_ind = 'A'"
End If

'Extract Service Locations with no Customer Services.
strSQL = strSQL & " UNION SELECT " & _
  "     cu.customer_id, upper(cu.customer_name), sl.service_location_id, upper(sl.service_location_name)" & _
  "   , 0 as network_element_id, null as network_element_name, null as network_element_type_code, null as network_element_desc" & _
  "   , 0 as customer_service_id, null as customer_service_desc, null as service_type_desc" & _
  "   , ad.address_id, house_number, ad.building_name, ad.street_name, ad.municipality_name, ad.province_state_lcode" & _
  " FROM" & _
  "   crp.customer cu," & _
  "   crp.customer_name_alias ca," & _
  "   crp.service_location sl," & _
  "   crp.address ad" & _
  " WHERE cu.customer_id = ca.customer_id" & _
  "   AND ca.customer_id = sl.customer_id (+)" & _
  "   AND sl.service_location_id not in (SELECT DISTINCT service_location_id FROM crp.customer_service WHERE service_location_id is not null)" & _
  "   AND sl.address_id = ad.address_id (+)" & _
  "   AND cu.customer_id = " & lngCustomerID

If bolActiveOnly = "yes" Then
   strSQL = strSQL & _
     " AND sl.record_status_ind = 'A'"
End If

strSQL = strSQL & " ORDER BY 2, 1, 4, 6, 10"

'get the recordset
set rsXLSCustServList=server.CreateObject("ADODB.Recordset")
rsXLSCustServList.Open strSQL, objConn

If err then
	DisplayError "BACK", "", err.Number, "XLSCustServList.asp - Cannot open database" , err.Description
End if

'put recordset into array
if not rsXLSCustServList.EOF then
	aList = rsXLSCustServList.GetRows
else
	Response.Write "0 Record Found"
	Response.End
end if

'release and kill the recordset and the connection objects
rsXLSCustServList.Close
set rsXLSCustServList = nothing
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
	strXLSFile =  "CustID" & request("hdnCustomerID") & "-CustServ-" & year(now())  & "-" & month(now()) & "-" & day(now()) & "-" & hour(now()) & "-" & minute(now()) & "-" & second(now()) & "-" & strRealUserID & ".xls"
	set objTxtStream = objFSO.CreateTextFile(strExportPath & strXLSFile, false, false)
	if err then
		DisplayError "CLOSE", "", err.Number, "XLSCustServList.asp - Cannot create Excel spreadsheet file due to the following reasons. Please contact your website administrator.", err.Description
	end if

	with objTxtStream
		.WriteLine "<table border=1>"

		'export the header
		.WriteLine "<THEAD>"
		.WriteLine "<TH>Customer ID</TH>"
		.WriteLine "<TH>Customer Name</TH>"
		.WriteLine "<TH>SL ID</TH>"
		.WriteLine "<TH>SL Name</TH>"
		.WriteLine "<TH>NE ID</TH>"
		.WriteLine "<TH>NE Name</TH>"
		.WriteLine "<TH>NE Type</TH>"
		.WriteLine "<TH>NE Description</TH>"
		.WriteLine "<TH>CS ID</TH>"
		.WriteLine "<TH>CS Description</TH>"
		.WriteLine "<TH>Service Type</TH>"
		.WriteLine "<TH>Address ID</TH>"
		.WriteLine "<TH>House Number</TH>"
		.WriteLine "<TH>Building</TH>"
		.WriteLine "<TH>Street</TH>"
		.WriteLine "<TH>Municipality</TH>"
		.WriteLine "<TH>Province</TH>"
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
			.WriteLine "<TD NOWRAP>" & routineHtmlString(aList(16,k)) & "&nbsp;</TD></TR>" & vbCrLf
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
					Response.Write strSQL
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
<form name=frmXLSCustServList action="XLSCustServList.asp" method=post>
	<input type=hidden name=hdnCustomerID   value="<%=lngCustomerID%>">
	<input type=hidden name=hdnCustomerName value="<%=strCustomerName%>">
	<input type=hidden name=chkActiveOnly   value="<%=bolActiveOnly%>">
	<input type=hidden name="hdnExport"     value="<%=strXLSFile%>">
<TABLE border=1 cellPadding=2 cellSpacing=0 width="100%">
	<THEAD>
	<TR>
		<TH align=left nowrap>Customer ID</TH>
		<TH align=left nowrap>Customer Name</TH>
		<TH align=left nowrap>SL ID</TH>
		<TH align=left nowrap>SL Name</TH>
		<TH align=left nowrap>NE ID</TH>
		<TH align=left nowrap>NE Name</TH>
		<TH align=left nowrap>NE Type</TH>
		<TH align=left nowrap>NE Description</TH>
		<TH align=left nowrap>CS ID</TH>
		<TH align=left nowrap>CS Description</TH>
		<TH align=left nowrap>Service Type</TH>
		<TH align=left nowrap>Address ID</TH>
		<TH align=left nowrap>House Number</TH>
		<TH align=left nowrap>Building</TH>
		<TH align=left nowrap>Street</TH>
		<TH align=left nowrap>Municipality</TH>
		<TH align=left nowrap>Province</TH>
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
		Response.Write "<td nowrap>" & aList(16,k) & "&nbsp;</td></tr>" & vbCrLf

   next
	%>
</TBODY>
<TFOOT>
<TR>
<TD align=left colSpan=17>
	<input type=hidden   name=hdnWinName    value="<%=strMyWinName%>">
	<input type=hidden   name=txtPageNumber value=<%=intPageNumber%>>
	<input type="submit" name="action" value="&lt;&lt;">
	<input type="submit" name="action" value="&lt;">
	<input type="text"   name="txtGoToPageNo" onClick="document.frmXLSCustServList.txtGoToPageNo.value=''" title="You can jump to a specific page by typing the page number in this box" value="page <%=intPageNumber%> of <%=intPageCount%>" style="HEIGHT: 22px; WIDTH: 150px">
	<input type="submit" name="action" value="&gt;">
	<input type="submit" name="action" value="&gt;&gt;">&nbsp;&nbsp;
	<img SRC="images/excel.gif" onclick="document.frmXLSCustServList.target='new';document.frmXLSCustServList.hdnExport.value='<%=strXLSFile%>';document.frmXLSCustServList.submit();document.frmXLSCustServList.target='_self';" WIDTH="32" HEIGHT="32">
</TD>
</TR>
</TFOOT>
<CAPTION>Records <%=m+1%> to <%=n+1%> of <%=UBound(aList, 2)+1 & " records"%></CAPTION>
</table>
</form>
</body>
</html>

<%@ Language=VBScript %>
<% option explicit %>
<!--% on error resume next%-->
<!--
********************************************************************************************
* Page name:	XLSContList.asp                                                            *
* Purpose:		To generate the Contact Validation Spreadsheet                             *
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
dim rsXLSContList, aList
dim strSQL
dim intPageNumber, intPageCount
dim strMyWinName, strBgColor

'get parameters
lngCustomerID   = Request("hdnCustomerID")
strCustomerName = Request("hdnCustomerName")
bolActiveOnly   = Request("chkActiveOnly")

strXLSFile      = Request("hdnExport")

'build query

'Employee Role.
strSQL = " SELECT " & _
  "     TRIM(UPPER(REPLACE(TRANSLATE(ct.last_name || ct.first_name || ct.middle_name, '~`!@#$%^&*()_-+={}[]:,<>,.?/|\', '                              '), ' ', '')))" & _
  "   , ct.contact_id, ct.contact_name, ct.last_name, ct.first_name, ct.middle_name, cu.customer_id, cu.customer_name" & _
  "   , null, 'Employee', null, null, null, DECODE(ct.staff_flag, 'Y', ct.staff_flag, ' ')" & _
  "   , cm.contact_method_desc, DECODE(ct.work_number_ext, null, ct.work_number, ct.work_number || ' Ext ' || ct.work_number_ext)" & _
  "   , ct.home_number, ct.cell_number, ct.pager_number, ct.fax_number, ct.email_address" & _
  " FROM" & _
  "   crp.contact ct" & _
  " , crp.customer cu" & _
  " , crp.contact_method cm" & _
  " WHERE " & _
  "       ct.work_for_customer_id = cu.customer_id" & _
  "   AND ct.prefercontactmethodcode = cm.contact_method_code (+)" & _
  "   AND cu.customer_id = " & lngCustomerID

If bolActiveOnly = "yes" Then
   strSQL = strSQL & _
     " AND (ct.staff_status_lcode is null or ct.staff_status_lcode is not null and ct.staff_status_lcode <> 'Departed')" & _
     " AND ct.record_status_ind = 'A'" & _
     " AND cu.record_status_ind = 'A'"
End If

'Customer Contact Role.
strSQL = strSQL & " UNION SELECT " & _
  "     TRIM(UPPER(REPLACE(TRANSLATE(ct.last_name || ct.first_name || ct.middle_name, '~`!@#$%^&*()_-+={}[]:,<>,.?/|\', '                              '), ' ', '')))" & _
  "   , ct.contact_id, ct.contact_name, ct.last_name, ct.first_name, ct.middle_name, cu.customer_id, cu.customer_name" & _
  "   , cc.customer_contact_type_lcode, lc.customer_contact_type_desc, TO_CHAR(cc.contact_priority, '00'), null, null, DECODE(ct.staff_flag, 'Y', ct.staff_flag, ' ')" & _
  "   , cm.contact_method_desc, DECODE(ct.work_number_ext, null, ct.work_number, ct.work_number || ' Ext ' || ct.work_number_ext)" & _
  "   , ct.home_number, ct.cell_number, ct.pager_number, ct.fax_number, ct.email_address" & _
  " FROM" & _
  "   crp.customer_contact cc" & _
  " , crp.contact ct" & _
  " , crp.customer cu" & _
  " , crp.lcode_customer_contact_type lc" & _
  " , crp.contact_method cm" & _
  " WHERE" & _
  "        cc.contact_id  = ct.contact_id" & _
  "   AND  cc.customer_id = cu.customer_id" & _
  "   AND  cc.customer_contact_type_lcode = lc.customer_contact_type_lcode" & _
  "   AND  ct.prefercontactmethodcode     = cm.contact_method_code (+)" & _
  "   AND  cu.customer_id = " & lngCustomerID

If bolActiveOnly = "yes" Then
   strSQL = strSQL & _
     " AND (ct.staff_status_lcode is null or ct.staff_status_lcode is not null and ct.staff_status_lcode <> 'Departed')" & _
     " AND cc.record_status_ind = 'A'" & _
     " AND ct.record_status_ind = 'A'" & _
     " AND cu.record_status_ind = 'A'"
End If

'Customer Service Contact Role.
strSQL = strSQL & " UNION SELECT " & _
  "     TRIM(UPPER(REPLACE(TRANSLATE(ct.last_name || ct.first_name || ct.middle_name, '~`!@#$%^&*()_-+={}[]:,<>,.?/|\', '                              '), ' ', '')))" & _
  "   , ct.contact_id, ct.contact_name, ct.last_name, ct.first_name, ct.middle_name, cu.customer_id, cu.customer_name" & _
  "   , csc.cust_serv_contact_type_lcode, lc.cust_serv_contact_type_desc, TO_CHAR(csc.contact_priority, '00'), '(' || TO_CHAR(cs.customer_service_id, '000000000') || '): ' || cs.customer_service_desc, null, DECODE(ct.staff_flag, 'Y', ct.staff_flag, ' ')" & _
  "   , cm.contact_method_desc, DECODE(ct.work_number_ext, null, ct.work_number, ct.work_number || ' Ext ' || ct.work_number_ext)" & _
  "   , ct.home_number, ct.cell_number, ct.pager_number, ct.fax_number, ct.email_address " & _
  " FROM" & _
  "   crp.customer_service_contact csc" & _
  " , crp.customer_service cs" & _
  " , crp.contact ct" & _
  " , crp.customer cu" & _
  " , crp.lcode_cust_serv_contact_type lc" & _
  " , crp.contact_method cm " & _
  " WHERE" & _
  "       csc.contact_id = ct.contact_id" & _
  "   AND csc.customer_service_id = cs.customer_service_id" & _
  "   AND cs.customer_id = cu.customer_id" & _
  "   AND csc.cust_serv_contact_type_lcode = lc.cust_serv_contact_type_lcode" & _
  "   AND ct.prefercontactmethodcode = cm.contact_method_code (+)" & _
  "   AND cu.customer_id = " & lngCustomerID

If bolActiveOnly = "yes" Then
   strSQL = strSQL & _
     " AND (ct.staff_status_lcode is null or ct.staff_status_lcode is not null and ct.staff_status_lcode <> 'Departed')" & _
     " AND csc.record_status_ind = 'A'" & _
     " AND cs.record_status_ind  = 'A'" & _
     " AND ct.record_status_ind  = 'A'" & _
     " AND cu.record_status_ind  = 'A'"
End If

'Service Location Contact Role.
strSQL = strSQL & " UNION SELECT" & _
  "      TRIM(UPPER(REPLACE(TRANSLATE(ct.last_name || ct.first_name || ct.middle_name, '~`!@#$%^&*()_-+={}[]:,<>,.?/|\', '                              '), ' ', '')))" & _
  "    , ct.contact_id, ct.contact_name, ct.last_name, ct.first_name, ct.middle_name, cu.customer_id, cu.customer_name" & _
  " , slc.serv_loc_contact_type_lcode, lc.serv_loc_contact_type_desc, TO_CHAR(slc.contact_priority, '00'), null, sl.service_location_name, DECODE(ct.staff_flag, 'Y', ct.staff_flag, ' ')" & _
  " , cm.contact_method_desc, DECODE(ct.work_number_ext, null, ct.work_number, ct.work_number || ' Ext ' || ct.work_number_ext)" & _
  " , ct.home_number, ct.cell_number, ct.pager_number, ct.fax_number, ct.email_address" & _
  " FROM" & _
  "   crp.service_location_contact slc" & _
  " , crp.service_location sl" & _
  " , crp.contact ct" & _
  " , crp.customer cu" & _
  " , crp.lcode_serv_loc_contact_type lc" & _
  " , crp.contact_method cm" & _
  " WHERE" & _
  "       slc.contact_id = ct.contact_id" & _
  "   AND slc.service_location_id = sl.service_location_id" & _
  "   AND sl.customer_id = cu.customer_id" & _
  "   AND slc.serv_loc_contact_type_lcode = lc.serv_loc_contact_type_lcode" & _
  "   AND ct.prefercontactmethodcode = cm.contact_method_code (+)" & _
  "   AND cu.customer_id = " & lngCustomerID

If bolActiveOnly = "yes" Then
   strSQL = strSQL & _
   " AND (ct.staff_status_lcode is null or ct.staff_status_lcode is not null and ct.staff_status_lcode <> 'Departed')" & _
   " AND slc.record_status_ind = 'A'" & _
   " AND sl.record_status_ind  = 'A'" & _
   " AND ct.record_status_ind  = 'A'" & _
   " AND cu.record_status_ind  = 'A'"
End If

'Design Staff.
strSQL = strSQL & " UNION SELECT " & _
  "     TRIM(UPPER(REPLACE(TRANSLATE(ct.last_name || ct.first_name || ct.middle_name, '~`!@#$%^&*()_-+={}[]:,<>,.?/|\', '                              '), ' ', '')))" & _
  "   , ct.contact_id, ct.contact_name, ct.last_name, ct.first_name, ct.middle_name, cu.customer_id, cu.customer_name" & _
  "   , null, 'Design Staff', null, null, null, DECODE(ct.staff_flag, 'Y', ct.staff_flag, ' ')" & _
  "   , cm.contact_method_desc, DECODE(ct.work_number_ext, null, ct.work_number, ct.work_number || ' Ext ' || ct.work_number_ext)" & _
  "   , ct.home_number, ct.cell_number, ct.pager_number, ct.fax_number, ct.email_address" & _
  " FROM" & _
  "   crp.customer_service cs" & _
  " , crp.contact ct" & _
  " , crp.customer cu" & _
  " , crp.contact_method cm" & _
  " WHERE" & _
  "       cs.design_staff_id is not null" & _
  "   AND cs.design_staff_id = ct.contact_id" & _
  "   AND cs.customer_id = cu.customer_id" & _
  "   AND ct.prefercontactmethodcode = cm.contact_method_code (+)" & _
  "   AND cu.customer_id = " & lngCustomerID

If bolActiveOnly = "yes" Then
   strSQL = strSQL & _
     " AND (ct.staff_status_lcode is null or ct.staff_status_lcode is not null and ct.staff_status_lcode <> 'Departed')" & _
     " AND cs.record_status_ind  = 'A'" & _
     " AND ct.record_status_ind  = 'A'" & _
     " AND cu.record_status_ind  = 'A'"
End If

'Installation Staff.
strSQL = strSQL & " UNION SELECT " & _
  "     TRIM(UPPER(REPLACE(TRANSLATE(ct.last_name || ct.first_name || ct.middle_name, '~`!@#$%^&*()_-+={}[]:,<>,.?/|\', '                              '), ' ', '')))" & _
  "   , ct.contact_id, ct.contact_name, ct.last_name, ct.first_name, ct.middle_name, cu.customer_id, cu.customer_name" & _
  "   , null, 'Installation Staff', null, null, null, DECODE(ct.staff_flag, 'Y', ct.staff_flag, ' ')" & _
  "   , cm.contact_method_desc, DECODE(ct.work_number_ext, null, ct.work_number, ct.work_number || ' Ext ' || ct.work_number_ext)" & _
  "   , ct.home_number, ct.cell_number, ct.pager_number, ct.fax_number, ct.email_address" & _
  " FROM" & _
  "   crp.customer_service cs" & _
  " , crp.contact ct" & _
  " , crp.customer cu" & _
  " , crp.contact_method cm" & _
  " WHERE" & _
  "       cs.installation_staff_id is not null" & _
  "   AND cs.installation_staff_id = ct.contact_id" & _
  "   AND cs.customer_id = cu.customer_id" & _
  "   AND ct.prefercontactmethodcode = cm.contact_method_code (+)" & _
  "   AND cu.customer_id = " & lngCustomerID

If bolActiveOnly = "yes" Then
   strSQL = strSQL & _
  "  AND (ct.staff_status_lcode is null or ct.staff_status_lcode is not null and ct.staff_status_lcode <> 'Departed')" & _
  "  AND cs.record_status_ind  = 'A'" & _
  "  AND ct.record_status_ind  = 'A'" & _
  "  AND cu.record_status_ind  = 'A'"
End If

'Operations Staff.
strSQL = strSQL & " UNION SELECT " & _
  "     TRIM(UPPER(REPLACE(TRANSLATE(ct.last_name || ct.first_name || ct.middle_name, '~`!@#$%^&*()_-+={}[]:,<>,.?/|\', '                              '), ' ', '')))" & _
  "   , ct.contact_id, ct.contact_name, ct.last_name, ct.first_name, ct.middle_name, cu.customer_id, cu.customer_name" & _
  "   , null, 'Operations Staff', null, null, null, DECODE(ct.staff_flag, 'Y', ct.staff_flag, ' ')" & _
  "   , cm.contact_method_desc, DECODE(ct.work_number_ext, null, ct.work_number, ct.work_number || ' Ext ' || ct.work_number_ext)" & _
  "   , ct.home_number, ct.cell_number, ct.pager_number, ct.fax_number, ct.email_address" & _
  " FROM" & _
  "   crp.customer_service cs" & _
  " , crp.contact ct" & _
  " , crp.customer cu" & _
  " , crp.contact_method cm" & _
  " WHERE" & _
  "       cs.operations_staff_id is not null" & _
  "   AND cs.operations_staff_id = ct.contact_id" & _
  "   AND cs.customer_id = cu.customer_id" & _
  "   AND ct.prefercontactmethodcode = cm.contact_method_code (+)" & _
  "   AND cu.customer_id = " & lngCustomerID

If bolActiveOnly = "yes" Then
   strSQL = strSQL & _
     " AND (ct.staff_status_lcode is null or ct.staff_status_lcode is not null and ct.staff_status_lcode <> 'Departed')" & _
     " AND cs.record_status_ind  = 'A'" & _
     " AND ct.record_status_ind  = 'A'" & _
     " AND cu.record_status_ind  = 'A'"
End If

'Implementation Staff.
strSQL = strSQL & " UNION SELECT " & _
  "      TRIM(UPPER(REPLACE(TRANSLATE(ct.last_name || ct.first_name || ct.middle_name, '~`!@#$%^&*()_-+={}[]:,<>,.?/|\', '                              '), ' ', '')))" & _
  "   , ct.contact_id, ct.contact_name, ct.last_name, ct.first_name, ct.middle_name, cu.customer_id, cu.customer_name" & _
  "   , null, 'Implementation Staff', null, null, null, DECODE(ct.staff_flag, 'Y', ct.staff_flag, ' ')" & _
  "   , cm.contact_method_desc, DECODE(ct.work_number_ext, null, ct.work_number, ct.work_number || ' Ext ' || ct.work_number_ext)" & _
  "   , ct.home_number, ct.cell_number , ct.pager_number, ct.fax_number, ct.email_address" & _
  " FROM" & _
  "   crp.customer_service cs" & _
  " , crp.contact ct" & _
  " , crp.customer cu" & _
  " , crp.contact_method cm" & _
  " WHERE" & _
  "       (cs.implementation_staff_id is not null and cs.implementation_staff_id = ct.contact_id)" & _
  "   AND (cs.customer_id is not null and cs.customer_id = cu.customer_id)" & _
  "   AND ct.prefercontactmethodcode = cm.contact_method_code (+)" & _
  "   AND cu.customer_id = " & lngCustomerID

If bolActiveOnly = "yes" Then
   strSQL = strSQL & _
     " AND (ct.staff_status_lcode is null or ct.staff_status_lcode is not null and ct.staff_status_lcode <> 'Departed')" & _
     " AND cs.record_status_ind  = 'A'" & _
     " AND ct.record_status_ind  = 'A'" & _
     " AND cu.record_status_ind  = 'A'"
End If

strSQL = strSQL & " ORDER by 1, 8, 9, 11"

'get the recordset
set rsXLSContList=server.CreateObject("ADODB.Recordset")
rsXLSContList.Open strSQL, objConn

If err then
	DisplayError "BACK", "", err.Number, "XLSContList.asp - Cannot open database" , err.Description
End if

'put recordset into array
if not rsXLSContList.EOF then
	aList = rsXLSContList.GetRows
else
	Response.Write "0 Record Found"
	Response.End
end if

'release and kill the recordset and the connection objects
rsXLSContList.Close
set rsXLSContList = nothing
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
	strXLSFile =  "CustID" & request("hdnCustomerID") & "-Cont-" & year(now())  & "-" & month(now()) & "-" & day(now()) & "-" & hour(now()) & "-" & minute(now()) & "-" & second(now()) & "-" & strRealUserID & ".xls"
    set objTxtStream = objFSO.CreateTextFile(strExportPath & strXLSFile, false, false)
	if err then
		DisplayError "CLOSE", "", err.Number, "XLSContList.asp - Cannot create Excel spreadsheet file due to the following reasons. Please contact your website administrator.", err.Description
	end if

	with objTxtStream
		.WriteLine "<table border=1>"

		'export the header
		.WriteLine "<THEAD>"
		.WriteLine "<TH>Contact ID</TH>"
		.WriteLine "<TH>Contact Name</TH>"
		.WriteLine "<TH>Last Name</TH>"
		.WriteLine "<TH>First Name</TH>"
		.WriteLine "<TH>Middle Name</TH>"
		.WriteLine "<TH>Customer ID</TH>"
		.WriteLine "<TH>Customer Name</TH>"
		.WriteLine "<TH>Role Code</TH>"
		.WriteLine "<TH>Role Description</TH>"
		.WriteLine "<TH>Priority</TH>"
		.WriteLine "<TH>Customer Service</TH>"
		.WriteLine "<TH>Service Location</TH>"
		.WriteLine "<TH>TELUS Staff</TH>"
		.WriteLine "<TH>Contact Method</TH>"
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
			.WriteLine "<TD NOWRAP>" & routineHtmlString(aList(17,k)) & "&nbsp;</TD>" & vbCrLf
			.WriteLine "<TD NOWRAP>" & routineHtmlString(aList(18,k)) & "&nbsp;</TD>" & vbCrLf
			.WriteLine "<TD NOWRAP>" & routineHtmlString(aList(19,k)) & "&nbsp;</TD>" & vbCrLf
			.WriteLine "<TD NOWRAP>" & routineHtmlString(aList(20,k)) & "&nbsp;</TD></TR>" & vbCrLf
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
<form name=frmXLSContList action="XLSContList.asp" method=post>
	<input type=hidden name=hdnCustomerID   value="<%=lngCustomerID%>">
	<input type=hidden name=hdnCustomerName value="<%=strCustomerName%>">
	<input type=hidden name="hdnExport"     value="<%=strXLSFile%>">
    <TABLE border=1 cellPadding=2 cellSpacing=0 width="100%">
	<THEAD>
	<TR>
		<TH align=left nowrap>Contact ID</TH>
		<TH align=left nowrap>Contact Name</TH>
		<TH align=left nowrap>Last Name</TH>
		<TH align=left nowrap>First Name</TH>
		<TH align=left nowrap>Middle Name</TH>
		<TH align=left nowrap>Customer ID</TH>
		<TH align=left nowrap>Customer Name</TH>
		<TH align=left nowrap>Role Code</TH>
		<TH align=left nowrap>Role Description</TH>
		<TH align=left nowrap>Priority</TH>
		<TH align=left nowrap>Customer Service</TH>
		<TH align=left nowrap>Service Location</TH>
		<TH align=left nowrap>TELUS Staff</TH>
		<TH align=left nowrap>Contact Method</TH>
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
		Response.Write "<td nowrap>" & aList(18,k) & "&nbsp;</td>" & vbCrLf
		Response.Write "<td nowrap>" & aList(19,k) & "&nbsp;</td>" & vbCrLf
		Response.Write "<td nowrap>" & aList(20,k) & "&nbsp;</td></tr>" & vbCrLf

   next
	%>
</TBODY>
<TFOOT>
<TR>
<TD align=left colSpan=20>
	<input type=hidden   name=hdnWinName    value="<%=strMyWinName%>">
	<input type=hidden   name=txtPageNumber value=<%=intPageNumber%>>
	<input type=hidden   name=chkActiveOnly value="<%=bolActiveOnly%>">
	<input type="submit" name="action" value="&lt;&lt;">
	<input type="submit" name="action" value="&lt;">
	<input type="text"   name="txtGoToPageNo" onClick="document.frmXLSContList.txtGoToPageNo.value=''" title="You can jump to a specific page by typing the page number in this box" value="page <%=intPageNumber%> of <%=intPageCount%>" style="HEIGHT: 22px; WIDTH: 150px">
	<input type="submit" name="action" value="&gt;">
	<input type="submit" name="action" value="&gt;&gt;">&nbsp;&nbsp;
	<img SRC="images/excel.gif" onclick="document.frmXLSContList.target='new';document.frmXLSContList.hdnExport.value='<%=strXLSFile%>';document.frmXLSContList.submit();document.frmXLSContList.target='_self';" WIDTH="32" HEIGHT="32">
</TD>
</TR>
</TFOOT>
<CAPTION>Records <%=m+1%> to <%=n+1%> of <%=UBound(aList, 2)+1 & " records"%></CAPTION>
</table>
</form>
</body>
</html>









<%@ Language=VBScript %>
<% option explicit%>
<!--%on error resume next%-->

<!--#include file="smaConstants.inc"-->
<!--#include file="databaseconnect.asp"-->
<!--#include file="smaProcs.inc"-->

<!--
*************************************************************************************
* File Name:	RSAST3List.asp
*
* Purpose:	    List Gateway circuits based on search criteria entered.
*
* In Param:		This page reads following cookies
*				CustomerName
*				AssetID
*				WinName
*				ServLocName
*
* Out Param:
*
* Created By:	Shawn Meyers
**************************************************************************************

        Date		Author			Changes/enhancements made
        -----		------		------------------------------------------------------
       10-04-01	     DTy		Change field names and variables.
                                Gateway IP         to GW Router T1 Serial IP Address
                                strGWIPAdd         to strGWRT1SIIPAddr
                                strWANIPAdd        to strWANIPAddr
                                strPNGIPAdd        to strLANIPAddr
                                txtPNGIP           to txtLANIPAddr
                                txtIPGate          to txtGWRT1SIIPAddr
                                txtLX25DNA         to txtLocalX25DNA
                                strLX25DNA         to strLocalX25DNA
                                Gateway DLCI IP    to Gateway DLCI (IP)
                                txtGWDLCI          to txtDLCIIP
                                strGWDLCIIP		   to strDLCIIP
                                Gateway DLCI POS   to Gateway DLCI (X25)
                                txtGWPOS           to txtDLCIX25
                                strGWDLCIPOS	   to strDLCIX25

                                WAN IP             to WAN IP Port Address
                                PNG IP             to LAN IP Port Address
                                PNG_IP_ID          to LAN_IP_ID
                                Site Name/Address  to Site Address

                                GATEWAY_DLCI_POS   to GATEWAY_DLCI_X25

                                Move Gateway DLCI (X25) before Gateway DLCI (IP)
                                Delete WAN IP DLCI (txtWANDLCI & strWANIPDLCI)
                                Delete POS IP DLCI (txtPOSIPDLCI & strPOSIPDLCI)
								Correct index pointers.

								Align column headings.
								Replace <TH></TD> with <TD></TD> to fix out of column
								  alignment in Excel 2000.

								Add new field 'Gateway Circuit Number'.
								Add Address ID in the extract and HREF.
								Expand list sort sequence
       18-Feb-02	 DTy		Active customers are those whose status is either
                                  'Prospect', 'OnHold' or 'Current'.
**************************************************************************************
-->

<%
'check the present user's rights

dim intAccessLevel
intAccessLevel = CInt(CheckLogon(strConst_RSAS))+400
if intAccessLevel < intConst_Access_ReadOnly then
	DisplayError "BACK", "", 0, "ACCESS DENIED", "You don't have access to RSAS maintainance screens. Please contact your system administrator."
end if

'declare the connection variables
dim sql
dim rsRSAST3List

'declare variable to be used in array for rows
dim aList
'declare the caller variable to be used for gobacks, etc.
dim strMyWinName

'get the caller variable value from the previous page
strMyWinName = Request("hdnWinName")

'declare results variables for fields in bottom navigation
dim intPageNumber
dim intPageCount

'these results variables are used to fill the displayed LIST fields
'use these ones for now - rest are dimmed below
dim strRSASCustID
dim strRSASCustomer
dim strRSASNodeName
dim strRSASSiteNameAddress
dim strGWRT1SIIPAddr
dim strDLCIIP
dim strDLCIX25
dim strWANIPAddr
dim strLANIPAddr
dim strTCNumber
dim strLocalX25DNA
dim strOrderNo
dim bolActiveOnly
dim strRecordStatus

'fill the results variables with data (from both textfields/hiddenfields on previous page)
'these are used to fill the displayed fields

strRSASCustomer = UCase(routineOraString(Trim(Request("txtCustomer"))))
strRSASNodeName = UCase(routineOraString(Trim(Request("txtNodeName"))))
strRSASSiteNameAddress = UCase(routineOraString(Trim(Request("txtSiteAddr"))))
strGWRT1SIIPAddr = UCase(routineOraString(Trim(Request("txtGWRT1SIIPAddr"))))
strDLCIIP = UCase(routineOraString(Trim(Request("txtDLCIIP"))))
strDLCIX25 = UCase(routineOraString(Trim(Request("txtDLCIX25"))))
strWANIPAddr = UCase(routineOraString(Trim(Request("txtWANIPAddr"))))
strLANIPAddr = UCase(routineOraString(Trim(Request("txtLANIPAddr"))))
strTCNumber = UCase(routineOraString(Trim(Request("txtTCNumber"))))
strLocalX25DNA = UCase(routineOraString(Trim(Request("txtLocalX25DNA"))))
strOrderNo = UCase(routineOraString(Trim(Request("txtOrderNo"))))

bolActiveOnly = Request("chkActiveOnly")

'connect to the database using the include file
'CONNECT using databaseconnect.asp

'extract the necessary data using sql query - tables have not been built yet -
' just a guess and will have to be modified

sql = "SELECT " &_
		"DISTINCT GA.GATEWAY_ID, "&_
		"GW_IP.IP_ADDRESS, " &_
		"GA.GATEWAY_DLCI_IP, " &_
		"GA.GATEWAY_DLCI_X25, " &_
		"GA.CUSTOMER_ID, " &_
		"CU.CUSTOMER_NAME, " &_
		"TC.TAIL_CIRCUIT_ID, " &_
		"WAN_IP.IP_ADDRESS, " &_
		"LAN_IP.IP_ADDRESS, " &_
		"TC.NODE_NAME, "&_
		"TC.TAIL_CIRCUIT_NUMBER, "&_
		"NVL(AD.BUILDING_NAME,'<NO BUILDING SPECIFIED>') ||CHR(13)||CHR(10)|| " &_
		"decode(AD.APARTMENT_NUMBER, null, null, rtrim(AD.APARTMENT_NUMBER) || ' ') || " &_
		"decode(to_char(AD.HOUSE_NUMBER) || AD.HOUSE_NUMBER_SUFFIX, null, null, rtrim(to_char(AD.house_number) || AD.house_number_suffix)  || ' ') || " &_
		"decode(AD.STREET_VECTOR, null, null, rtrim(AD.STREET_VECTOR) || ' ') || " &_
		"NVL(AD.LONG_STREET_NAME,'NO STREET ADDRESS')||CHR(13)||CHR(10)|| " &_
		"NVL(AD.MUNICIPALITY_NAME,'NO MUNICIPALITY')||' '|| " &_
		"NVL(AD.PROVINCE_STATE_LCODE,'NO PROVINCE')||CHR(13)||CHR(10)|| " &_
		"NVL(AD.POSTAL_CODE_ZIP,'NO POSTAL CODE') ADDRESS, " &_
		"TC.ORDER_NUMBER, " &_
		"GA.GATEWAY_CIRCUIT_NUMBER, " &_
		"AD.ADDRESS_ID " &_
	"FROM " &_
		"CRP.RSAS_GATEWAY          GA, "&_
		"CRP.RSAS_IP_ADDRESS    GW_IP, "&_
		"CRP.RSAS_IP_ADDRESS   WAN_IP, "&_
		"CRP.RSAS_IP_ADDRESS   LAN_IP, "&_
		"CRP.RSAS_TAIL_CIRCUIT     TC, "&_
		"CRP.CUSTOMER              CU, "&_
		"CRP.ADDRESS               AD "
	if strRSASCustomer <> "" then
		sql = sql & ", CRP.CUSTOMER_NAME_ALIAS CA "
	end if

    if strLocalX25DNA <> "" then
	   sql = sql & ", CRP.RSAS_DEVICE DV "
    end if

	sql = sql & "WHERE " &_
		"GA.GATEWAY_ID = TC.GATEWAY_ID (+) "&_
		"AND GA.GATEWAY_IP_ID = GW_IP.IP_ADDRESS_ID "&_
		"AND TC.WAN_IP_ID = WAN_IP.IP_ADDRESS_ID (+) "&_
		"AND TC.LAN_IP_ID = LAN_IP.IP_ADDRESS_ID (+) "&_
		"AND GA.CUSTOMER_ID = CU.CUSTOMER_ID (+) "&_
		"AND TC.SITE_ADDRESS_ID = AD.ADDRESS_ID (+) "

if strRSASCustomer <> "" then
	sql = sql &	" AND CU.CUSTOMER_ID = CA.CUSTOMER_ID " &_
		" AND UPPER(CA.CUSTOMER_NAME_ALIAS_UPPER) LIKE '" &routineOraString(strRSASCustomer)& "%' "
end if


if strRSASNodeName <> "" then
	sql = sql &	" AND UPPER(TC.NODE_NAME) LIKE '" &routineOraString(strRSASNodeName)& "%' "
end if

if strRSASSiteNameAddress <> "" then 'Needs to properly use the Address search ...
	'sql = sql &	" AND UPPER(AD.LONG_STREET_NAME) LIKE '" &routineOraString(strRSASSiteNameAddress)& "%' "
	sql = sql & " AND ((Upper(NVL(AD.BUILDING_NAME,'<NO BUILDING SPECIFIED>') ||CHR(13)||CHR(10)|| " &_
					"decode(AD.APARTMENT_NUMBER, null, null, rtrim(AD.APARTMENT_NUMBER) || ' ') || " &_
					"decode(to_char(AD.HOUSE_NUMBER) || AD.HOUSE_NUMBER_SUFFIX, null, null, rtrim(to_char(AD.house_number) || AD.house_number_suffix)  || ' ') || " &_
					"decode(AD.STREET_VECTOR, null, null, rtrim(AD.STREET_VECTOR) || ' ') || " &_
					"NVL(AD.LONG_STREET_NAME,'NO STREET ADDRESS')||CHR(13)||CHR(10)|| " &_
					"NVL(AD.MUNICIPALITY_NAME,'NO MUNICIPALITY')||' '|| " &_
					"NVL(AD.PROVINCE_STATE_LCODE,'NO PROVINCE')||CHR(13)||CHR(10)|| " &_
					"NVL(AD.POSTAL_CODE_ZIP,'NO POSTAL CODE')))  LIKE '" & routineOraString(strRSASSiteNameAddress) & "%' )"
end if

if strGWRT1SIIPAddr <> "" then
	sql = sql &	" AND UPPER(GW_IP.IP_ADDRESS) LIKE '" &routineOraString(strGWRT1SIIPAddr)& "%' "
end if

if strDLCIIP <> "" then
	sql = sql &	" AND UPPER(GA.GATEWAY_DLCI_IP) LIKE '" &routineOraString(strDLCIIP)& "%' "
end if

if strDLCIX25 <> "" then
	sql = sql &	" AND UPPER(GA.GATEWAY_DLCI_X25) LIKE '" &routineOraString(strDLCIX25)& "%' "
end if

if strWANIPAddr <> "" then
	sql = sql &	" AND UPPER(WAN_IP.IP_ADDRESS) LIKE '" &routineOraString(strWANIPAddr)& "%' "
end if

if strLANIPAddr <> "" then
	sql = sql &	" AND UPPER(LAN_IP.IP_ADDRESS) LIKE '" &routineOraString(strLANIPAddr)& "%' "
end if

if strTCNumber <> "" then
	sql = sql &	" AND UPPER(TC.TAIL_CIRCUIT_NUMBER) LIKE '" &routineOraString(strTCNumber)& "%' "
end if

'DEVICE if strPollCode <> "" then
	'sql = sql &	" AND UPPER(DEV.POLL_CODE) LIKE '" &routineOraString(strPollCode)& "%' "
'end if

if strLocalX25DNA <> "" then
	sql =sql & " AND UPPER(TC.TAIL_CIRCUIT_ID) = DV.TAIL_CIRCUIT_ID AND UPPER(DV.LOCAL_X25_DNA) LIKE  '"  &routineOraString(strLocalX25DNA)& "%' "
end if

if strOrderNo <> "" then
	sql = sql &	" AND UPPER(TC.ORDER_NUMBER) LIKE '" &routineOraString(strOrderNo)& "%' "
end if

	If bolActiveOnly = "yes" then
		strRecordStatus = " and GA.record_status_ind = 'A' " &_
		                  " and CU.customer_status_lcode IN ('Prospect', 'Current', 'OnHold')"
	Else 'no
		strRecordStatus = " "
	End If

sql = sql & strRecordStatus
sql = sql & "ORDER BY GA.CUSTOMER_ID, GATEWAY_CIRCUIT_NUMBER, GW_IP.IP_ADDRESS, " &_
      "GA.GATEWAY_DLCI_IP, GA.GATEWAY_DLCI_X25, WAN_IP.IP_ADDRESS, LAN_IP.IP_ADDRESS, " &_
      " TC.NODE_NAME, TC.TAIL_CIRCUIT_NUMBER"

'Response.Write (sql & "<p>")
'Response.end

'set the recordset and parse through the data
set rsRSAST3List=server.CreateObject("ADODB.Recordset")
rsRSAST3List.Open sql, objConn
if err then
	DisplayError "BACK", "", err.Number, "RSAST3List.asp - Cannot open database", err.Description
end if

'search through the recordset and get the data
if not rsRSAST3List.EOF then
	aList = rsRSAST3List.GetRows
else
	Response.Write "0 record found"
	Response.end
end if



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
                              strRealUserID =Session("username")

                              'determine export path

                              dim strExportPath, liLength
                              strExportPath =Request.ServerVariables("PATH_TRANSLATED")
                              While (Right(strExportPath, 1) <> "\" And Len(strExportPath) <> 0)
                                   liLength = Len(strExportPath) - 1
                                   strExportPath = Left(strExportPath, liLength)
                              Wend
                              strExportPath = strExportPath & "export\"

                              'create the scripting object

                              dim objFSO, objTxtStream
                              set objFSO = server.CreateObject("Scripting.FileSystemObject")

                              'create the export text file (overwrite if it already exists)

                              set objTxtStream = objFSO.CreateTextFile(strExportPath&strRealUserID&"-RSAST3List.xls", true,false)

								if err then
									DisplayError "CLOSE", "", err.Number, "RSAST3List.asp - Cannot create Excel spreadsheet file due to the following reasons.  Please contact your website administrator.", err.Description
								end if

							  with objTxtStream

                                   .WriteLine "<table border=1>"

                                   'export the table header
                                   .WriteLine "<TR>"

                                   .WriteLine "<TD align = 'center'>Customer</TD>"
                                   .WriteLine "<TD align = 'center'>GW Circuit Number</TD>"
                                   .WriteLine "<TD align = 'center'>GW Router T1 Serial Interface IP</TD>"
                                   .WriteLine "<TD align = 'center'>Gateway DLCI (X25)</TD>"
                                   .WriteLine "<TD align = 'center'>Gateway DLCI (IP)</TD>"
                                   .WriteLine "<TD align = 'center'>WAN IP Port Address</TD>"
                                   .WriteLine "<TD align = 'center'>LAN IP Port Address</TD>"
                                   .WriteLine "<TD align = 'center'>Node Name</TD>"
                                   .WriteLine "<TD align = 'center'>Tail Circuit Number</TD>"
                                   .WriteLine "<TD align = 'center'>Site Address</TD>"
                                   .WriteLine "<TD align = 'center'>Order Number</TD>"								   'end the table header
                                   .WriteLine "</TR>"

                                   'export the body
                                   for k = 0 to UBound(aList, 2)
                                         'Alternate row background colour
                                         if Int(k/2) = k/2 then
'                                             .WriteLine "<TR bgcolor=#ffffcc>"
                                              .WriteLine "<TR>"
                                         else
'                                             .WriteLine "<TR bgcolor=#ffffff>"
                                              .WriteLine "<TR>"
                                         end if


                                         'fill the table with data

                                         .WriteLine "<TD NOWRAP>"&routineHtmlString(aList(5,k))&"</TD>"
                                         .WriteLine "<TD NOWRAP>"&routineHtmlString(aList(13,k))&"</TD>"
                                         .WriteLine "<TD NOWRAP>"&routineHtmlString(aList(1,k))&"</TD>"
                                         .WriteLine "<TD NOWRAP>"&routineHtmlString(aList(3,k))&"</TD>"
                                         .WriteLine "<TD NOWRAP>"&routineHtmlString(aList(2,k))&"</TD>"
                                         .WriteLine "<TD NOWRAP>"&routineHtmlString(aList(7,k))&"</TD>"
                                         .WriteLine "<TD NOWRAP>"&routineHtmlString(aList(8,k))&"</TD>"
                                         .WriteLine "<TD NOWRAP>"&routineHtmlString(aList(9,k))&"</TD>"
                                         .WriteLine "<TD NOWRAP>"&routineHtmlString(aList(10,k))&"</TD>"
                                         .WriteLine "<TD NOWRAP>"&routineHtmlString(aList(11,k))&"</TD>"
                                         .WriteLine "<TD NOWRAP>"&routineHtmlString(aList(12,k))&"</TD>"

                                         .WriteLine "</TR>"
                                   next
                                   .WriteLine "</table>"

                              end with

                              objTxtStream.Close
								sql = "<script type=""text/javascript"">document.location=""export/"&strRealUserID&"-RSAST3List.xls"";</script>"
								Response.Write sql
								Response.End

                              'Response.redirect "export/"&strRealUserID&"-RSAST3List.xls"

						elseif Request("txtGoToPageNo") <> "" then
						intPageNumber = CInt(Request("txtGoToPageNo"))
					else
						intPageNumber = 1
					end if
end select

if intPageNumber < 1 then intPageNumber = 1
if intPageNumber > intPageCount then intPageNumber = intPageCount

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
<html>
<head>
<title>RSAS Results</title>
<link href="stylesheets/styles.css" rel="stylesheet" type="text/css">
</head>
<script TYPE="TEXT/JAVASCRIPT">


function go_back(strTCNumber, strGWRT1SIIPAddr, strRSASCustID, strRSASCustomer)
{
	parent.opener.document.forms[0].txtTCNumber.value = strTCNumber;
	parent.opener.document.forms[0].txtGWRT1SIIPAddr.value = strGWRT1SIIPAddr;
	parent.opener.document.forms[0].txtRSASCustID.value = strRSASCustID;
	parent.opener.document.forms[0].txtRSASCustomer.value = strRSASCustomer;
	parent.window.close ();
}
</script>
<body>

<!--hidden fields are filled with values from previous page as well-->
<form name="frmRSAST3List" action="RSAST3List.asp" method="POST">
    <input type="hidden" name="txtRSASCustID" value="<%=strRSASCustID%>">
    <input type="hidden" name="txtRSASCustomer" value="<%=strRSASCustomer%>">
    <input type="hidden" name="txtRSASNodeName" value="<%=strRSASNodeName%>">
    <input type="hidden" name="txtRSASSiteNameAddress" value="<%=strRSASSiteNameAddress%>">
    <input type="hidden" name="txtGWRT1SIIPAddr" value="<%=strGWRT1SIIPAddr%>">
    <input type="hidden" name="txtDLCIIP" value="<%=strDLCIIP%>">
    <input type="hidden" name="txtGWPOS" value="<%=strDLCIX25%>">
    <input type="hidden" name="txtWANIPAddr" value="<%=strWANIPAddr%>">
    <input type="hidden" name="txtLANIPAddr" value="<%=strLANIPAddr%>">
    <input type="hidden" name="txtTCNumber" value="<%=strTCNumber%>">
    <input type="hidden" name="txtLocalX25DNA" value="<%=strLocalX25DNA%>">
    <input type="hidden" name="txtOrderNo" value="<%=strOrderNo%>">
    <input type=hidden   name=chkActiveOnly value="<%=bolActiveOnly%>">
    <input type="hidden" name="hdnWinName" value="<%=strMyWinName%>">
    <input type="hidden" name="hdnExport" value>

<table border="1" cellPadding="2" cellSpacing="0" width="100%">
<thead>
	<tr>
	   <!-- <TH align=left>Catalogue ID</TH> -->
		<th align="center">Customer</th>
		<th align="center">GW Circuit Number</th>
		<th align="center">GW Router T1 Serial Interface IP</th>
		<th align="center">Gateway DLCI (X25)</th>
		<th align="center">Gateway DLCI (IP)</th>
		<th align="center">WAN IP Port Address</th>
		<th align="center">LAN IP Port Address</th>
		<th align="center">Node Name</th>
		<th align="center">Tail Circuit Number</th>
		<th align="center">Site Address</th>
		<th align="center">Order Number</th>
	</tr>
</thead>
<tbody>
<%
'display the table
for k = m to n
	'Alternate row background colour
	if Int(k/2) = k/2 then
		Response.write "<TR>"
	else
		Response.write "<TR bgcolor=White>"
	end if

	'this first condition is the list that appears in the popup window
	'if the lookup button is pressed.


	'this has to be modified
	if strMyWinName = "Popup" then
		Response.Write "<TD NOWRAP><a href=# onclick=""go_back('"&aList(6,k)&"','"&routineJavascriptString(aList(0,k))&"','"&routineJavascriptString(aList(4,k))&"','"&routineJavascriptString(aList(5,k))&"')"">"&routineHtmlString(aList(5,k))&"&nbsp;</a></TD>"&vbCrLf
		Response.Write "<TD NOWRAP><a href=# onclick=""go_back('"&aList(6,k)&"','"&routineJavascriptString(aList(0,k))&"','"&routineJavascriptString(aList(4,k))&"','"&routineJavascriptString(aList(5,k))&"')"">"&routineHtmlString(aList(13,k))&"&nbsp;</a></TD>"&vbCrLf
		Response.Write "<TD NOWRAP><a href=# onclick=""go_back('"&aList(6,k)&"','"&routineJavascriptString(aList(0,k))&"','"&routineJavascriptString(aList(4,k))&"','"&routineJavascriptString(aList(5,k))&"')"">"&routineHtmlString(aList(1,k))&"&nbsp;</a></TD>"&vbCrLf
		Response.Write "<TD NOWRAP><a href=# onclick=""go_back('"&aList(6,k)&"','"&routineJavascriptString(aList(0,k))&"','"&routineJavascriptString(aList(4,k))&"','"&routineJavascriptString(aList(5,k))&"')"">"&routineHtmlString(aList(3,k))&"&nbsp;</a></TD>"&vbCrLf
		Response.Write "<TD NOWRAP><a href=# onclick=""go_back('"&aList(6,k)&"','"&routineJavascriptString(aList(0,k))&"','"&routineJavascriptString(aList(4,k))&"','"&routineJavascriptString(aList(5,k))&"')"">"&routineHtmlString(aList(2,k))&"&nbsp;</a></TD>"&vbCrLf
		Response.Write "<TD NOWRAP><a href=# onclick=""go_back('"&aList(6,k)&"','"&routineJavascriptString(aList(0,k))&"','"&routineJavascriptString(aList(4,k))&"','"&routineJavascriptString(aList(5,k))&"')"">"&routineHtmlString(aList(7,k))&"&nbsp;</a></TD>"&vbCrLf
		Response.Write "<TD NOWRAP><a href=# onclick=""go_back('"&aList(6,k)&"','"&routineJavascriptString(aList(0,k))&"','"&routineJavascriptString(aList(4,k))&"','"&routineJavascriptString(aList(5,k))&"')"">"&routineHtmlString(aList(8,k))&"&nbsp;</a></TD>"&vbCrLf
		Response.Write "<TD NOWRAP><a href=# onclick=""go_back('"&aList(6,k)&"','"&routineJavascriptString(aList(0,k))&"','"&routineJavascriptString(aList(4,k))&"','"&routineJavascriptString(aList(5,k))&"')"">"&routineHtmlString(aList(9,k))&"&nbsp;</a></TD>"&vbCrLf
		Response.Write "<TD NOWRAP><a href=# onclick=""go_back('"&aList(6,k)&"','"&routineJavascriptString(aList(0,k))&"','"&routineJavascriptString(aList(4,k))&"','"&routineJavascriptString(aList(5,k))&"')"">"&routineHtmlString(aList(10,k))&"&nbsp;</a></TD>"&vbCrLf
		Response.Write "<TD NOWRAP><a href=# onclick=""go_back('"&aList(6,k)&"','"&routineJavascriptString(aList(0,k))&"','"&routineJavascriptString(aList(4,k))&"','"&routineJavascriptString(aList(5,k))&"')"">"&routineHtmlString(aList(11,k))&"&nbsp;</a></TD>"&vbCrLf
		Response.Write "<TD NOWRAP><a href=# onclick=""go_back('"&aList(6,k)&"','"&routineJavascriptString(aList(0,k))&"','"&routineJavascriptString(aList(4,k))&"','"&routineJavascriptString(aList(5,k))&"')"">"&routineHtmlString(aList(12,k))&"&nbsp;</a></TD>"&vbCrLf

	'this second condition is the list that appears normally.
	else
		Response.Write "<TD NOWRAP><a target=""_parent"" href=""RSAST3GWDetail.asp?action=edit&GWID="&aList(0,k)&"&CustID="&aList(4,k)&"&AddrID="&aList(14,k)&""">"&routineHtmlString(aList(5,k))&"&nbsp;</a></TD>"&vbCrLf
		Response.Write "<TD NOWRAP><a target=""_parent"" href=""RSAST3GWDetail.asp?action=edit&GWID="&aList(0,k)&"&CustID="&aList(4,k)&"&AddrID="&aList(14,k)&""">"&routineHtmlString(aList(13,k))&"&nbsp;</a></TD>"&vbCrLf
		Response.Write "<TD NOWRAP><a target=""_parent"" href=""RSAST3GWDetail.asp?action=edit&GWID="&aList(0,k)&"&CustID="&aList(4,k)&"&AddrID="&aList(14,k)&""">"&routineHtmlString(aList(1,k))&"&nbsp;</a></TD>"&vbCrLf
		Response.Write "<TD NOWRAP><a target=""_parent"" href=""RSAST3GWDetail.asp?action=edit&GWID="&aList(0,k)&"&CustID="&aList(4,k)&"&AddrID="&aList(14,k)&""">"&routineHtmlString(aList(3,k))&"&nbsp;</a></TD>"&vbCrLf
		Response.Write "<TD NOWRAP><a target=""_parent"" href=""RSAST3GWDetail.asp?action=edit&GWID="&aList(0,k)&"&CustID="&aList(4,k)&"&AddrID="&aList(14,k)&""">"&routineHtmlString(aList(2,k))&"&nbsp;</a></TD>"&vbCrLf
		Response.Write "<TD NOWRAP><a target=""_parent"" href=""RSAST3GWDetail.asp?action=edit&GWID="&aList(0,k)&"&CustID="&aList(4,k)&"&AddrID="&aList(14,k)&""">"&routineHtmlString(aList(7,k))&"&nbsp;</a></TD>"&vbCrLf
		Response.Write "<TD NOWRAP><a target=""_parent"" href=""RSAST3GWDetail.asp?action=edit&GWID="&aList(0,k)&"&CustID="&aList(4,k)&"&AddrID="&aList(14,k)&""">"&routineHtmlString(aList(8,k))&"&nbsp;</a></TD>"&vbCrLf
		Response.Write "<TD NOWRAP><a target=""_parent"" href=""RSAST3GWDetail.asp?action=edit&GWID="&aList(0,k)&"&CustID="&aList(4,k)&"&AddrID="&aList(14,k)&""">"&routineHtmlString(aList(9,k))&"&nbsp;</a></TD>"&vbCrLf
		Response.Write "<TD NOWRAP><a target=""_parent"" href=""RSAST3GWDetail.asp?action=edit&GWID="&aList(0,k)&"&CustID="&aList(4,k)&"&AddrID="&aList(14,k)&""">"&routineHtmlString(aList(10,k))&"&nbsp;</a></TD>"&vbCrLf
		Response.Write "<TD NOWRAP><a target=""_parent"" href=""RSAST3GWDetail.asp?action=edit&GWID="&aList(0,k)&"&CustID="&aList(4,k)&"&AddrID="&aList(14,k)&""">"&routineHtmlString(aList(11,k))&"&nbsp;</a></TD>"&vbCrLf
		Response.Write "<TD NOWRAP><a target=""_parent"" href=""RSAST3GWDetail.asp?action=edit&GWID="&aList(0,k)&"&CustID="&aList(4,k)&"&AddrID="&aList(14,k)&""">"&routineHtmlString(aList(12,k))&"&nbsp;</a></TD>"&vbCrLf
	end if
	Response.Write "</TR>"
next
%>

</tbody>
<tfoot>
<tr>
<td align="left" colSpan="12">
	<input type="hidden" name="txtPageNumber" value="<%=intPageNumber%>">
	<input type="submit" name="action" value="&lt;&lt;">
	<input type="submit" name="action" value="&lt;">
	<input type="text" name="txtGoToPageNo" onClick="document.frmRSAST3List.txtGoToPageNo.value=''" title="You can jump to a specific page by typing the page number in this box" value="page <%=intPageNumber%> of <%=intPageCount%>" style="HEIGHT: 22px; WIDTH: 150px">
	<input type="submit" name="action" value="&gt;">
	<input type="submit" name="action" value="&gt;&gt;">
    <img SRC="images/excel.gif" onclick="document.frmRSAST3List.target='new';document.frmRSAST3List.hdnExport.value='xls';document.frmRSAST3List.submit();document.frmRSAST3List.hdnExport.value='';document.frmRSAST3List.target='_self';" </TD WIDTH="32" HEIGHT="32">
</tr>
</tfoot>
<caption>Records <%=m+1%> to <%=n+1%> of <%=UBound(aList, 2)+1 & " records"%></caption>
</table>
</form>

<%
'close the recordset and the connection objects
rsRSAST3List.Close
set rsRSAST3List = nothing

objConn.close
set objConn = nothing


%>
</body>
</html>


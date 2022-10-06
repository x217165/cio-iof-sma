<%@ Language=VBScript %>
<%
option explicit
on error resume next
%>
<!--#include file="smaConstants.inc"-->
<!--#include file="databaseconnect.asp"-->
<!--#include file="smaProcs.inc"-->
<!--
*************************************************************************************
* File Name:	manobjlist.asp
*
* Purpose:
*
* In Param:
*
*
* Out Param:
*
* Created By:
* Edited by:    Adam Haydey Jan 25, 2001
*               Added Customer Service City, Customer Service Address, TAC Assset Code and Non-Correlated Only search fields.
*				TAC Asset Code was added to the search results.
**************************************************************************************
		 Date		Author			Changes/enhancements made
	  20-Jul-01	     DTy		When 'Active Only' is selected:
		                          Exclude Customers that are marked as soft deleted.
		                          Exclude Name Alias that are marked as soft deleted.
		                          Exclude Service Locations that are marked as soft deleted.
		                          Exclude Managed Correlation that are marked as soft deleted.
		                          Exclude Network Element that are marked as soft deleted.
       11-Feb-02	 DTy		Remove special characters on managed objects and customer service names
                                  when extracting records.
       18-Feb-02	 DTy		Active customers are those whose status is either
                                  'Prospect', 'OnHold' or 'Current'.
       14-Mar-02	 DTy		Add Port Name and LAN IP as search fields, in similar way
                                  as the Managed Object Alias Name.
       11-Aug-04	 ACheung  	Add Lynx default severity as search fields.
       24-Jul-15   PSmith  Rewrote the SQL query to be more efficient.
       04-Feb-16   PSmith  Simplified the NE name matching like condition.
**************************************************************************************
-->
<%
'check user's rights
dim intAccessLevel
intAccessLevel = CInt(CheckLogon(strConst_ManagedObjects))
if intAccessLevel < intConst_Access_ReadOnly then
	DisplayError "BACK", "", 0, "ACCESS DENIED", "You don't have access to managed object. Please contact your system administrator"
end if

dim sql
dim rsMOList

dim aList, intPageNumber, intPageCount, strWinName
dim strRegion, strManObjName, strManObjType, strCustomer, strServLoc, strSupportGroup, strAssetId, chkNullOBDialup, CHKnULLip
dim strIPAddress, strOBDialup, chkActiveOnly, strServLocCity, strServLocAdd, strBarCode, chkNonCorrOnly
dim strManObjPort, strManObjLANIP
dim strLynxSeverity

'get search criteria
strRegion = Request("selRegion")
strManObjName = UCase(Trim(Request("txtManObjName")))
strManObjType = Request("selManObjType")
strCustomer = UCase(Trim(Request("txtCustomer")))
strServLoc = UCase(Trim(Request("txtServLoc")))
strSupportGroup = UCase(Trim(Request("selSupportGroup")))
strAssetId = Request("txtAssetId")
strOBDialup = UCase(Trim(Request("txtOBDialup")))
chkNullOBDialup = Request("chkNullOBDialup")
strIPAddress = UCase(Trim(Request("txtIPAddress")))
chkNullIP = Request("chkNullIP")
strWinName = Request("hdnWinName")
chkActiveOnly = Request("chkActiveOnly")
strServLocCity = UCase(Trim(Request("txtServLocCity")))
strServLocAdd = UCase(Trim(Request("txtServLocAdd")))
strBarCode = UCase(Trim(Request("txtBarCode")))
chkNonCorrOnly = Request("chkNonCorrOnly")

strManObjPort  = UCase(Trim(Request("txtManObjPort")))
strManObjLANIP = UCase(Trim(Request("txtManObjLANIP")))

strLynxSeverity = Request("selRepairPriority")

'build query
sql = "SELECT " &_
		"DISTINCT MOL.NETWORK_ELEMENT_ID, "&_
		"MOL.NETWORK_ELEMENT_NAME, " &_
		"MOL.CUSTOMER_NAME, " &_
		"MOL.SERVICE_LOCATION_NAME, "&_
		"MOL.NETWORK_ELEMENT_TYPE_CODE, "&_
		"MOL.MANAGED_IP_ADDRESS, "&_
		"MOL.OUT_OF_BAND_DIALUP, "&_
		"MOL.NOC_REGION_LCODE, " &_
		"MOL.GROUP_NAME, " &_
		"MOL.REMEDY_SUPPORT_GROUP_ID, " &_
		"MOL.BARCODE, " &_
		"MOL.FULL_CLLI_CODE "&_
	"FROM "
		if strCustomer <> "" then
			sql = sql &	"CRP.CUSTOMER_NAME_ALIAS		CNA, "
		end if
		if strServLocCity <> "" OR strServLocAdd <> "" then
			sql = sql & "CRP.ADDRESS A, "
		end if
		if strLynxSeverity <> "ALL" then
			sql = sql &_
			"CRP.LCODE_LYNX_DEF_SEV LDS, "
		end if
		sql = sql &_
		"CRP.MAN_OBJECT_LIST MOL WHERE 1=1 "
		if strCustomer <> "" then
			sql = sql &	"AND MOL.CUSTOMER_ID = CNA.CUSTOMER_ID "
		end if
		if chkNonCorrOnly <> "" then
			sql = sql & "AND MOL.MANAGED_CORRELATION_ID is NULL "
		end if
		if strLynxSeverity <> "ALL" then
			sql = sql & "AND MOL.LYNX_DEF_SEV_LCODE = LDS.LYNX_DEF_SEV_LCODE (+) "
		end if
		if strServLocCity <> "" OR strServLocAdd <> "" then
			sql = sql & "AND MOL.ADDRESS_ID = A.ADDRESS_ID (+) "
		end if

if strManObjName <> "" then
 	 sql = sql & "AND (UPPER(MOL.NETWORK_ELEMENT_NAME) like '%" & strManObjName &_
	           "%' OR UPPER(MOL.NETWORK_ELEMENT_NAME_ALIAS) like '%" & strManObjName & "%'"
   if strManObjPort <> "" and strManObjLANIP <> "" then
	   sql = sql & "OR MOL.NETWORK_ELEMENT_ID IN (SELECT DISTINCT NETWORK_ELEMENT_ID FROM CRP.NETWORK_ELEMENT_PORT WHERE " & rtRmvSpChr("NETWORK_ELEMENT_PORT_NAME", "Y") & " LIKE '%" & rtRmvSpChr(strManObjPort, "N")& "%' and " &_
             rtRmvSpChr("NETWORK_ELEMENT_PORT_IP", "Y") & "LIKE '%" & rtRmvSpChr(strManObjLANIP, "N")& "%')) "
   else
      if strManObjPort <> "" then
	     sql = sql & " OR MOL.NETWORK_ELEMENT_ID IN (SELECT DISTINCT NETWORK_ELEMENT_ID FROM CRP.NETWORK_ELEMENT_PORT WHERE " & rtRmvSpChr("NETWORK_ELEMENT_PORT_NAME", "Y") & " LIKE '%" & rtRmvSpChr(strManObjPort, "N")& "%')) "
      else
	     if strManObjLANIP <> "" then
	        sql = sql & " OR MOL.NETWORK_ELEMENT_ID IN (SELECT DISTINCT NETWORK_ELEMENT_ID FROM CRP.NETWORK_ELEMENT_PORT WHERE " & rtRmvSpChr("NETWORK_ELEMENT_PORT_IP", "Y") & " LIKE '%" & rtRmvSpChr(strManObjLANIP, "N")& "%')) "
	     else
	        sql =sql & ") "
	     end if
	  end if
   end if
else
   'MO name/alias is empty and both Port Name & LAN IP are filled-up.
   if strManObjPort <> "" and strManObjLANIP <> "" then
	   sql = sql & "AND MOL.NETWORK_ELEMENT_ID IN (SELECT DISTINCT NETWORK_ELEMENT_ID FROM CRP.NETWORK_ELEMENT_PORT WHERE " & rtRmvSpChr("NETWORK_ELEMENT_PORT_NAME", "Y") & "LIKE '%" & rtRmvSpChr(strManObjPort, "N")& "%' and " &_
             rtRmvSpChr("NETWORK_ELEMENT_PORT_IP", "Y") & "LIKE '%" & rtRmvSpChr(strManObjLANIP, "N") & "%') "
   else
      'MO name/alias is empty and only Port Name is filled-up.
      if strManObjPort <> "" then
	     sql = sql & "AND MOL.NETWORK_ELEMENT_ID IN (SELECT DISTINCT NETWORK_ELEMENT_ID FROM CRP.NETWORK_ELEMENT_PORT WHERE " & rtRmvSpChr("NETWORK_ELEMENT_PORT_NAME", "Y") & "LIKE '%" & rtRmvSpChr(strManObjPort, "N")& "%') "
      end if

      'MO name/alias is empty and only LAN IP is filled-up.
      if strManObjLANIP <> "" then
	     sql = sql & "AND MOL.NETWORK_ELEMENT_ID IN (SELECT DISTINCT NETWORK_ELEMENT_ID FROM CRP.NETWORK_ELEMENT_PORT WHERE " & rtRmvSpChr("NETWORK_ELEMENT_PORT_IP", "Y") & "LIKE '%" & rtRmvSpChr(strManObjLANIP, "N")& "%') "
      end if
	end if
end if


if strManObjType <> "ALL" then
	sql = sql &	" AND MOL.NETWORK_ELEMENT_TYPE_CODE = '" &routineOraString(strManObjType)& "' "
end if

if strCustomer <> "" then
	sql = sql &	" AND UPPER(CNA.CUSTOMER_NAME_ALIAS_UPPER) LIKE '" &routineOraString(strCustomer)& "%' "
end if

if strServLoc <> "" then
	sql = sql &	" AND UPPER(MOL.SERVICE_LOCATION_NAME) LIKE '" &routineOraString(strServLoc)& "%' "
end if

if strServLocCity <> "" then
	sql = sql & " AND UPPER(A.MUNICIPALITY_NAME) LIKE '" &routineOraString(strServLocCity)& "%' "
end if

if strServLocAdd <> "" then
	sql = sql &_
		" AND Upper(NVL(A.BUILDING_NAME,'NO BUILDING NAME') ||CHR(13)||CHR(10)|| " &_
		"decode(A.APARTMENT_NUMBER, null, null, rtrim(A.APARTMENT_NUMBER) || ' ') || " &_
		"decode(A.HOUSE_NUMBER || A.HOUSE_NUMBER_SUFFIX, null, null, rtrim(A.house_number) || A.house_number_suffix)  || ' ') || " &_
		"decode(A.STREET_VECTOR, null, null, rtrim(A.STREET_VECTOR) || ' ') || " &_
		"NVL(A.LONG_STREET_NAME,'NO STREET ADDRESS')||CHR(13)||CHR(10)|| " &_
		"NVL(A.MUNICIPALITY_NAME,'NO MUNICIPALITY')||' '|| " &_
		"NVL(A.PROVINCE_STATE_LCODE,'NO PROVINCE')||CHR(13)||CHR(10)|| " &_
		"NVL(A.POSTAL_CODE_ZIP,'NO POSTAL CODE')  LIKE '%" & routineOraString(strServLocAdd) & "%' "
end if
if strSupportGroup <> "" then
	sql = sql &	" AND UPPER(MOL.REMEDY_SUPPORT_GROUP_ID) = '" &routineOraString(strSupportGroup)& "' "
end if

if strRegion <> "ALL" then
	sql = sql &	" AND MOL.NOC_REGION_LCODE = '" &routineOraString(strRegion)& "' "
end if

if strAssetId <> "" then
	sql = sql &	" AND MOL.ASSET_ID = " & strAssetId
end if

if chkNullIP <> "" then
	sql = sql & " AND (UPPER(MOL.MANAGED_IP_ADDRESS) = 'N/A' OR MOL.MANAGED_IP_ADDRESS is null) "
elseif strIPAddress <> "" then
	sql = sql & " AND MOL.MANAGED_IP_ADDRESS LIKE '" & routineOraString(strIPAddress) & "%' "
end if

if chkNullOBDialup <> "" then
	sql = sql & " AND (UPPER(MOL.OUT_OF_BAND_DIALUP) = 'N/A' OR MOL.OUT_OF_BAND_DIALUP is null) "
elseif strOBDialup <> "" then
	sql = sql & " AND MOL.OUT_OF_BAND_DIALUP LIKE '" & routineOraString(strOBDialup) & "%' "
end if

if strBarCode <> "" then
	sql = sql & " AND UPPER(MOL.BARCODE) LIKE '" & routineOraString(strBarCode) & "%' "
end if
if chkActiveOnly <> "" then
	sql = sql & "  and MOL.customer_status_lcode IN ('Prospect', 'Current', 'OnHold')" &_
          " AND MOL.NE_RECORD_STATUS_IND = 'A' AND MOL.CUS_RECORD_STATUS_IND = 'A' " &_
	      " AND MOL.SL_RECORD_STATUS_IND (+) = 'A'"

	if strCustomer <> "" then
		sql = sql &	" AND CNA.RECORD_STATUS_IND = 'A'"
	end if

	if strServLocCity <> "" OR strServLocAdd <> "" then
		sql = sql & "AND A.RECORD_STATUS_IND (+) = 'A'"
	end if
end if

if strLynxSeverity <> "ALL" then
	sql = sql & "AND LDS.LYNX_DEF_SEV_DESC = '"  &routineOraString(strLynxSeverity)& "' "
end if

'order by object's name
sql = sql & " ORDER BY UPPER(NETWORK_ELEMENT_NAME)"
'Response.Write sql
'Response.End

'get the recordset
set rsMOList=server.CreateObject("ADODB.Recordset")
rsMOList.Open sql, objConn
if err then
	DisplayError "BACK", "", err.Number, "CANNOT CREATE RECORDSET", err.Description
end if


if not rsMOList.EOF then
	aList = rsMOList.GetRows
else
	Response.Write "0 records found"
	'check for cookie
	'if Request.cookies("MoTacname") <> "" then
		'Response.Write "<script>parent.document.location='manobjdet.asp?ne_id='</script>"
	'end if
	Response.end
end if

'release and kill the recordset and the connection objects
rsMOList.Close
set rsMOList = nothing

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
	case else	if Request("hdnExport") <> "" then
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
						set objTxtStream = objFSO.CreateTextFile(strExportPath&strRealUserID&"-ManagedObjects.xls", true, false)

						with objTxtStream
							.WriteLine "<table border=1>"

							'export the header
							.WriteLine "<TR>"
'							.WriteLine "<TR bgcolor=#ffcc99>"
							.WriteLine "<TH>Object Name</TH>"
							.WriteLine "<TH>Customer Name</TH>"
							.WriteLine "<TH>Service Location</TH>"
							.WriteLine "<TH>Type</TH>"
							.WriteLine "<TH>IP Address</TH>"
							.WriteLine "<TH>O/B Dialup</TH>"
							.WriteLine "<TH>Region</TH>"
							.WriteLine "<TH>Support Group</TH>"
							.WriteLine "<TH>TAC Asset Code</TH>"
							.WriteLine "<TH>CLLI Code</TH>"
							.WriteLine "<TD NOWRAP>&nbsp;&nbsp;</TH>"
							.WriteLine "</TR>"

							'export the body
							for k = 0 to UBound(aList, 2)
								'Alternate row background colour
								if Int(k/2) = k/2 then
'									.WriteLine "<TR bgcolor=#ffffcc>"
									.WriteLine "<TR>"
								else
'									.WriteLine "<TR bgcolor=#ffffff>"
									.WriteLine "<TR>"
								end if

								.WriteLine "<TD NOWRAP>"&routineHtmlString(aList(1,k))&" &nbsp; </TD>"
								.WriteLine "<TD NOWRAP>"&routineHtmlString(aList(2,k))&" &nbsp; </TD>"
								.WriteLine "<TD NOWRAP>"&routineHtmlString(aList(3,k))&" &nbsp; </TD>"
								.WriteLine "<TD NOWRAP>"&routineHtmlString(aList(4,k))&" &nbsp;</TD>"
								.WriteLine "<TD NOWRAP>"&routineHtmlString(aList(5,k))&" &nbsp; </TD>"
								.WriteLine "<TD NOWRAP>"&routineHtmlString(aList(6,k))&" &nbsp; </TD>"
								.WriteLine "<TD NOWRAP>"&routineHtmlString(aList(7,k))&" &nbsp; </TD>"
								.WriteLine "<TD NOWRAP>"&routineHtmlString(aList(8,k))&" &nbsp; </TD>"
								.WriteLine "<TD NOWRAP>"&routineHtmlString(aList(10,k))&" &nbsp; </TD>"
								.WriteLine "<TD NOWRAP>"&routineHtmlString(aList(11,k))&" &nbsp; </TD>"
								.WriteLine "<TD NOWRAP>&nbsp;&nbsp;</TD>"
								.WriteLine "</TR>"
							next
							.WriteLine "</table>"
						end with
						objTxtStream.Close
						sql = "<script type=""text/javascript"">document.location=""export/"&strRealUserID&"-ManagedObjects.xls"";</script>"
						Response.Write sql
						Response.End
'						Response.redirect "export/"&strRealUserID&"-ManagedObjects.xls"

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
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft FrontPage 12.0">
<META HTTP-EQUIV="Pragma" CONTENT="no-cache">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<TITLE>Managed Objects Results</TITLE>
<LINK href="stylesheets/styles.css" rel="stylesheet" type="text/css">
<script type="text/javascript" SRC="GeneralJavaFunctions.js"></script>
<script type="text/javascript">
function go_back(strMO_ID, strMO_NAME, strMO_TYPE) {
	parent.opener.document.forms[0].hdnNewElementID.value = strMO_ID;
	parent.opener.document.forms[0].hdnNewElementName.value = strMO_NAME;
	parent.opener.document.forms[0].hdnNewElementType.value = "MO";
	parent.opener.btn_iFrmAddNewElement();
	DeleteCookie("WinName");
	parent.window.close ();
}
</script>
</HEAD>
<BODY>

<FORM method=post name=frmMOList action="manobjlist.asp">
    <input type=hidden name=selRegion value="<%=strRegion%>">
    <input type=hidden name=txtManObjName value="<%=strManObjName%>">
    <input type=hidden name=selManObjType value="<%=strManObjType%>">
    <input type=hidden name=txtCustomer value="<%=strCustomer%>">
    <input type=hidden name=txtServLoc value="<%=strServLoc%>">
    <input type=hidden name=selSupportGroup value="<%=strSupportGroup%>">
    <input type=hidden name=txtOBDialup value="<%=strOBDialup%>">
    <input type=hidden name=chkNullOBDialup value="<%=chkNullOBDialup%>">
    <input type=hidden name=txtIPAddress value="<%=strIPAddress%>">
    <input type=hidden name=chkNullIP value="<%=chkNullIP%>">
    <input type=hidden name=hdnWinName value="<%=strWinName%>">
    <input type=hidden name=chkActiveOnly value="<%=chkActiveOnly%>">

    <input type=hidden name=txtServLocCity value="<%=strServLocCity%>">
    <input type=hidden name=txtServLocAdd value="<%=strServLocAdd%>">
    <input type=hidden name=txtBarCode value="<%=strBarCode%>">
    <input type=hidden name=chkNonCorrOnly value="<%=chkNonCorrOnly%>">
    <input type=hidden name=selRepairPriority value="<%=strLynxSeverity%>">
    <input type=hidden name="hdnExport" value>



<TABLE border=1 cellspacing=0 cellpadding=2 width="100%">
<THEAD>
	<TR>
		<TH>Object Name</TH>
		<TH>Customer</TH>
		<TH>Service Location</TH>
		<TH>Type</TH>
		<TH>IP Address</TH>
		<TH>O/B Dialup</TH>
		<TH>Region</TH>
		<TH>Support Group</TH>
		<TH>TAC Asset Code</TH>
		<TH>CLLI Code</TH>
	</TR>
</THEAD>
<TBODY>
<%
'display the table
for k = m to n
	'Alternate row background colour
	if Int(k/2) = k/2 then
		Response.write "<TR>"
	else
		Response.write "<TR bgcolor=White>"
	end if
	if strWinName = "Popup" then
		Response.Write "<td nowrap><a href=""#"" onClick=""go_back('"&aList(0,k)& "', '" &routineJavascriptString(aList(1,k))& "','" &routineJavascriptString(aList(4,k))& "')"">" &aList(1,k)& "</a>&nbsp;</td>"&vbCrLf
		Response.Write "<td nowrap><a href=""#"" onClick=""go_back('"&aList(0,k)& "', '" &routineJavascriptString(aList(1,k))& "','" &routineJavascriptString(aList(4,k))& "')"">" &aList(2,k)& "</a>&nbsp;</td>"&vbCrLf
		Response.Write "<td nowrap><a href=""#"" onClick=""go_back('"&aList(0,k)& "', '" &routineJavascriptString(aList(1,k))& "','" &routineJavascriptString(aList(4,k))& "')"">" &aList(3,k)& "</a>&nbsp;</td>"&vbCrLf
		Response.Write "<td nowrap><a href=""#"" onClick=""go_back('"&aList(0,k)& "', '" &routineJavascriptString(aList(1,k))& "','" &routineJavascriptString(aList(4,k))& "')"">" &aList(4,k)& "</a>&nbsp;</td>"&vbCrLf
		Response.Write "<td nowrap><a href=""#"" onClick=""go_back('"&aList(0,k)& "', '" &routineJavascriptString(aList(1,k))& "','" &routineJavascriptString(aList(4,k))& "')"">" &aList(5,k)& "</a>&nbsp;</td>"&vbCrLf
		Response.Write "<td nowrap><a href=""#"" onClick=""go_back('"&aList(0,k)& "', '" &routineJavascriptString(aList(1,k))& "','" &routineJavascriptString(aList(4,k))& "')"">" &aList(6,k)& "&nbsp;</a></td>"&vbCrLf
		Response.Write "<td nowrap><a href=""#"" onClick=""go_back('"&aList(0,k)& "', '" &routineJavascriptString(aList(1,k))& "','" &routineJavascriptString(aList(4,k))& "')"">" &aList(7,k)& "&nbsp;</a></td>"&vbCrLf
		Response.Write "<td nowrap><a href=""#"" onClick=""go_back('"&aList(0,k)& "', '" &routineJavascriptString(aList(1,k))& "','" &routineJavascriptString(aList(4,k))& "')"">" &aList(8,k)& "</a>&nbsp;</td>"&vbCrLf
		Response.Write "<td nowrap><a href=""#"" onClick=""go_back('"&aList(0,k)& "', '" &routineJavascriptString(aList(1,k))& "','" &routineJavascriptString(aList(4,k))& "')"">" &aList(10,k)& "</a>&nbsp;</td>"&vbCrLf
		Response.Write "<td nowrap><a href=""#"" onClick=""go_back('"&aList(0,k)& "', '" &routineJavascriptString(aList(1,k))& "','" &routineJavascriptString(aList(4,k))& "')"">" &aList(11,k)& "</a>&nbsp;</td>"&vbCrLf
		Response.Write "</TR>"
	else
		Response.Write "<TD nowrap><a target=""_parent"" href=""manobjdet.asp?ne_id="&aList(0,k)&""">"&routineHtmlString(aList(1,k))&"</a>&nbsp;</TD>"&vbCrLf
		Response.Write "<TD nowrap><a target=""_parent"" href=""manobjdet.asp?ne_id="&aList(0,k)&""">"&routineHtmlString(aList(2,k))&"</a>&nbsp;</TD>"&vbCrLf
		Response.Write "<TD nowrap><a target=""_parent"" href=""manobjdet.asp?ne_id="&aList(0,k)&""">"&routineHtmlString(aList(3,k))&"</a>&nbsp;</TD>"&vbCrLf
		Response.Write "<TD nowrap><a target=""_parent"" href=""manobjdet.asp?ne_id="&aList(0,k)&""">"&routineHtmlString(aList(4,k))&"</a>&nbsp;</TD>"&vbCrLf
		Response.Write "<TD nowrap><a target=""_parent"" href=""manobjdet.asp?ne_id="&aList(0,k)&""">"&routineHtmlString(aList(5,k))&"</a>&nbsp;</TD>"&vbCrLf
		Response.Write "<TD nowrap><a target=""_parent"" href=""manobjdet.asp?ne_id="&aList(0,k)&""">"&routineHtmlString(aList(6,k))&"&nbsp;</a></TD>"&vbCrLf
		Response.Write "<TD nowrap><a target=""_parent"" href=""manobjdet.asp?ne_id="&aList(0,k)&""">"&routineHtmlString(aList(7,k))&"&nbsp;</a></TD>"&vbCrLf
		Response.Write "<TD nowrap><a target=""_parent"" href=""manobjdet.asp?ne_id="&aList(0,k)&""">"&routineHtmlString(aList(8,k))&"</a>&nbsp;</TD>"&vbCrLf
		Response.Write "<TD nowrap><a target=""_parent"" href=""manobjdet.asp?ne_id="&aList(0,k)&""">"&routineHtmlString(aList(10,k))&"</a>&nbsp;</TD>"&vbCrLf
		Response.Write "<TD nowrap><a target=""_parent"" href=""manobjdet.asp?ne_id="&aList(0,k)&""">"&routineHtmlString(aList(11,k))&"</a>&nbsp;</TD>"&vbCrLf
		Response.Write "</TR>"
	end if
next
%>
</TBODY>
<TFOOT>
<TR>
<TD align=left colSpan=10>
	<input type=hidden name=txtPageNumber value=<%=intPageNumber%>>
	<input type="submit" name="action" value="&lt;&lt;">
	<input type="submit" name="action" value="&lt;">
	<input type="text" name="txtGoToPageNo" title="You can jump to a specific page by typing the page number in this box" onclick="document.frmMOList.txtGoToPageNo.value=''" value="page <%=intPageNumber%> of <%=intPageCount%>" style="HEIGHT: 22px; WIDTH: 150px">
	<input type="submit" name="action" value="&gt;">
	<input type="submit" name="action" value="&gt;&gt;">
	<img SRC="images/excel.gif" onclick="document.frmMOList.target='new'; document.frmMOList.hdnExport.value='xls';document.frmMOList.submit();document.frmMOList.hdnExport.value='';document.frmMOList.target='_self';" WIDTH="32" HEIGHT="32">
</TD>
</TR>
</TFOOT>
<CAPTION>Records <%=m+1%> to <%=n+1%> of <%=UBound(aList, 2)+1 & " records"%></CAPTION>
</TABLE>
</FORM>
</BODY>
</HTML>

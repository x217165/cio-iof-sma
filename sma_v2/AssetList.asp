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
* File Name:	AssetList.asp
*
* Purpose:
*
* In Param:		This page reads following cookies
*
* Out Param:
*
* Created By:
* Edited by:    Adam Haydey Jan 25, 2001
*               CR 1549 Asset now populates hdnProvinceCode.value, hdnCity.value and hdnStreetName.value
*				when it is used as a popup search screen.  This allows the Managed Object
*				detail screen to default the search for the Service Locations to the asset Address.
* Edited by:    Adam Haydey Mar 2, 2001
*               CR 1550 ... Tac Asset code (Barcode) is also used in the query and displayed in the result.
**************************************************************************************
		 Date		Author			Changes/enhancements made
		20-Jul-01	 DTy		When 'Active Only' is selected:
		                          Exclude Customers that are marked as soft deleted.
		                          Exclude Addresses that are marked as soft deleted.
		                          Exclude Assets that are marked as soft deleted.
		                        Do not outer join crp.CUSTOMER.
       18-Feb-02	 DTy		Active customers are those whose status is either
                                  'Prospect', 'OnHold' or 'Current'.
**************************************************************************************
-->
<%
'check user's rights
dim intAccessLevel
intAccessLevel = CInt(CheckLogon(strConst_Asset))
if intAccessLevel < intConst_Access_ReadOnly then
	DisplayError "BACK", "", 0, "ACCESS DENIED", "You don't have access to asset catalogue. Please contact your system administrator."
end if

dim sql, strMyWinName
dim rsACList

dim aList, intPageNumber, intPageCount
dim strAssetId,strTacName,strSerial,strWhereClause,strOrderBy,strCustomer
dim StrAssetType,StrSql,StrAssetMake,StrAssetModel,StrFromClause,StrRegion,strClliCode,StrActive, StrAssetCode

'get the caller
strMyWinName = Request("hdnWinName")

'get search criteria
strAssetId = routineOraString(UCase(Trim(Request("txtassetid"))))
strTacName = routineOraString(UCase(Trim(Request("txttacname"))))
strSerial = routineOraString(UCase(Trim(Request("txtserial"))))
StrAssetType = routineOraString(UCase(Trim(Request("selassettype"))))
StrAssetMake = routineOraString(UCase(Trim(Request("txtassetmake"))))
StrAssetModel = routineOraString(UCase(Trim(Request("txtassetmodel"))))
StrRegion  = routineOraString(UCase(TRIM(Request("selRegion"))))
strCustomer  = routineOraString(UCase(TRIM(Request("txtcustomerName"))))
strClliCode = routineOraString(UCase(TRIM(Request("txtcllicode"))))
StrActive = routineOraString(UCase(TRIM(Request("chkactive"))))
StrAssetCode = routineOraString(UCase(TRIM(Request("txtassetcode"))))
'Response.Write (strCustomer & "<BR>")
'Response.Write (Request("txtassetcode") & "<BR>")
'Response.Write (StrAssetCode)
'Response.End

'build query
StrSql = "SELECT A.ASSET_ID,B.ASSET_TYPE_DESC,A.PURCHASE_ORDER_NUMBER,TO_CHAR(A.DATE_RECEIVED,'MON-DD-YYYY')," &_
         "A.TAC_NAME,A.SERIAL_NUMBER,D.MAKE_DESC,E.MODEL_DESC,F.CUSTOMER_NAME,"&_
         "LTRIM(NVL(G.BUILDING_NAME,'')) ||' '||NVL(G.STREET,'')||' '||"&_
         "NVL(G.MUNICIPALITY_NAME,'')||' '||" &_
         "NVL(G.PROVINCE_STATE_LCODE,'')||' '||" &_
         "NVL(G.POSTAL_CODE_ZIP,'') ADDRESS_A ,F.NOC_REGION_LCODE,H.PART_NUMBER_DESC, "&_
         "F.CUSTOMER_ID, F.CUSTOMER_SHORT_NAME, C.ASSET_CATALOGUE_ID, G.STREET, G.MUNICIPALITY_NAME, G.PROVINCE_STATE_LCODE, A.ASSET_BARCODE"

StrFromClause = " FROM CRP.ASSET A," &_
                "CRP.ASSET_TYPE B," &_
                "CRP.ASSET_CATALOGUE C," &_
                "CRP.MAKE D," &_
                "CRP.MODEL E," &_
                "CRP.CUSTOMER F," &_
                "CRP.V_ADDRESS_CONSOLIDATED_STREET G, " &_
                "CRP.PART_NUMBER H "

strWhereClause = " WHERE A.ASSET_TYPE_ID = B.ASSET_TYPE_ID(+) AND " &_
                 "A.ASSET_CATALOGUE_ID = C.ASSET_CATALOGUE_ID(+) AND " &_
                 "C.MAKE_ID = D.MAKE_ID(+) AND " &_
                 "C.MODEL_ID =E.MODEL_ID(+) AND " &_
                 "A.CUSTOMER_ID = F.CUSTOMER_ID AND " &_
                 "A.ADDRESS_ID = G.ADDRESS_ID AND " &_
                 "C.PART_NUMBER_ID = H.PART_NUMBER_ID(+) "

strOrderBy = " ORDER BY A.ASSET_ID"

StrSql = StrSql & StrFromClause


IF  LEN(strAssetId) > 0 THEN
      strWhereClause = strWhereClause & " AND A.ASSET_ID = " & strAssetId
END IF


IF  LEN(strSerial) > 0 THEN
      strWhereClause = strWhereClause & " AND A.SERIAL_NUMBER LIKE '" & strSerial &"%'"
END IF


IF  LEN(strTacName ) > 0 THEN
      strWhereClause = strWhereClause & " AND A.TAC_NAME LIKE '" & strTacName  &"%'"
END IF


IF  LEN(StrAssetMake) > 0 THEN
      strWhereClause = strWhereClause & " AND UPPER(D.MAKE_DESC) LIKE '" & StrAssetMake  &"%'"

END IF

IF  LEN(StrAssetModel) > 0 THEN
      strWhereClause = strWhereClause & " AND UPPER(E.MODEL_DESC) LIKE '" & StrAssetModel  &"%'"

END IF

IF  LEN(StrAssetType) > 0 THEN
      strWhereClause = strWhereClause & " AND B.ASSET_TYPE_ID = " & StrAssetType

END IF

IF  LEN(strClliCode) > 0 THEN
      strWhereClause = strWhereClause & " AND UPPER(A.CLLI_CODE) LIKE '" & strClliCode  &"%'"

END IF


IF  LEN(strCustomer) > 0 THEN
      strWhereClause = strWhereClause & " AND A.CUSTOMER_ID IN " &_
                      "(SELECT CUSTOMER_ID FROM CRP.CUSTOMER_NAME_ALIAS " &_
                      " WHERE CUSTOMER_NAME_ALIAS_UPPER LIKE '" & strCustomer & "%'"

	IF  (StrActive="YES")  THEN
		strWhereClause = strWhereClause & " AND RECORD_STATUS_IND = 'A' "
	END IF
	strWhereClause = strWhereClause & ")"
END IF


IF  LEN(StrRegion) > 0 THEN
      strWhereClause = strWhereClause & " AND UPPER(F.NOC_REGION_LCODE) LIKE '" & StrRegion  &"%'"

END IF


'Response.Write (StrAssetCode)
'Response.Write (LEN(StrAssetCode))
'Response.End

IF LEN(StrAssetCode) >0 THEN
	  strWhereClause =strWhereClause & " AND UPPER(A.ASSET_BARCODE) LIKE '" & StrAssetCode &"%'"
END IF

IF  (StrActive="YES")  THEN
	strWhereClause = strWhereClause & " AND f.customer_status_lcode IN ('Prospect', 'Current', 'OnHold')" &_
	                 " AND A.RECORD_STATUS_IND = 'A' AND " & _
	                 "F.RECORD_STATUS_IND (+) = 'A' AND G.RECORD_STATUS_IND (+) = 'A' "
END IF

StrSql = StrSql  & strWhereClause & strOrderBy

'Response.Write "SQL STATEMENT WIH WHERE=" & StrSql & "<p>"


'Response.Write StrSql
'Response.End

'get the recordset

set rsACList=server.CreateObject("ADODB.Recordset")
rsACList.Open StrSql, objConn
if err then
	DisplayError "BACK", "", err.Number, "CANNOT CREATE RECORDSET", err.Description
end if

if not rsACList.EOF then
	aList = rsACList.GetRows
else
	Response.Write "0 records found"
	Response.end
end if

'release and kill the recordset and the connection objects
rsACList.Close
set rsACList = nothing

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
						set objTxtStream = objFSO.CreateTextFile(strExportPath&strRealUserID&"-asset.xls", true, false)

							if err then
								DisplayError "CLOSE", "", err.Number, "AssetList.asp - Cannot create Excel spreadsheet file due to the following reasons.  Please contact your website administrator.", err.Description
							end if


						with objTxtStream
							.WriteLine "<table border=1>"

							'export the header
							.WriteLine "<TR>"
'							.WriteLine "<TR bgcolor=#ffcc99>"

							.WriteLine "<TH>Asset ID</TH>"
							.WriteLine "<TH>Asset Type</TH>"
							.WriteLine "<TH>PO#</TH>"
							.WriteLine "<TH>Received Date</TH>"
							.WriteLine "<TH>Tac Name</TH>"
							.WriteLine "<TH>Serial#</TH>"
							.WriteLine "<TH>Make</TH>"
							.WriteLine "<TH>Model</TH>"
							.WriteLine "<TH>Part Number</TH>"
							.WriteLine "<TH>Customer</TH>"
							.WriteLine "<TH>Address</TH>"
							.WriteLine "<TH>Region</TH>"
							.WriteLine "<TH>TAC Asset Code</TH>"
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

								.WriteLine "<TD NOWRAP>"&routineHtmlString(aList(0,k))&"</TD>"
								.WriteLine "<TD NOWRAP>"&routineHtmlString(aList(1,k))&"</TD>"
								.WriteLine "<TD NOWRAP>"&routineHtmlString(aList(2,k))&"</TD>"
								.WriteLine "<TD NOWRAP>"&routineHtmlString(aList(3,k))&"</TD>"
								.WriteLine "<TD NOWRAP>"&routineHtmlString(aList(4,k))&"</TD>"
								.WriteLine "<TD NOWRAP>"&routineHtmlString(aList(5,k))&"</TD>"
								.WriteLine "<TD NOWRAP>"&routineHtmlString(aList(6,k))&"</TD>"
								.WriteLine "<TD NOWRAP>"&routineHtmlString(aList(7,k))&"</TD>"
								.WriteLine "<TD NOWRAP>"&routineHtmlString(aList(11,k))&"</TD>"
								.WriteLine "<TD NOWRAP>"&routineHtmlString(aList(8,k))&"</TD>"
								.WriteLine "<TD NOWRAP>"&routineHtmlString(aList(9,k))&"</TD>"
								.WriteLine "<TD NOWRAP>"&routineHtmlString(aList(10,k))&"</TD>"
								.WriteLine "<TD NOWRAP>"&routineHtmlString(aList(18,k))&"</TD>"
								.WriteLine "</TR>"
							next
							.WriteLine "</table>"
						end with
						objTxtStream.Close
						strsql = "<script type=""text/javascript"">document.location=""export/"&strRealUserID&"-asset.xls"";</script>"
						Response.Write strsql
						Response.End
'						Response.redirect "export/"&strRealUserID&"-asset.xls"
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
<TITLE>Asset Catalog Results</TITLE>
<META HTTP-EQUIV="Pragma" CONTENT="no-cache">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<LINK href="stylesheets/styles.css" rel="stylesheet" type="text/css">
</HEAD>
<SCRIPT TYPE="TEXT/JAVASCRIPT">
function go_back(lngAssetID, strAssetName,strAssetMake,strAssetModel,strAssetPartNo,strCustomerID,strCustomerShortName,strCustomerName,strAssetCatalogueID, strStreet, strCity, strProvince)
{
	// Replaces <none> with a blank for Managed Object Detail CR 1549
	if (strAssetMake=="<none>")
	{
		strAssetMake= "";
	}
	if (strAssetModel=="<none>")
	{
		strAssetModel= "";
	}
	if (strAssetPartNo=="<none>")
	{
		strAssetPartNo= "";
	}
	parent.opener.document.forms[0].hdnAssetID.value = lngAssetID;
	parent.opener.document.forms[0].txtAssetName.value = strAssetName;
	parent.opener.document.forms[0].txtAssetMake.value = strAssetMake;
	parent.opener.document.forms[0].txtAssetModel.value = strAssetModel;
	parent.opener.document.forms[0].txtAssetPartNo.value = strAssetPartNo;
	try{
		parent.opener.document.forms[0].hdnAssetCatalogueID.value = strAssetCatalogueID;
		parent.opener.document.forms[0].hdnCustomerID.value = strCustomerID;
		parent.opener.document.forms[0].txtCustomerShortName.value = strCustomerShortName;
		parent.opener.document.forms[0].txtCustomerName.value = strCustomerName;
		parent.opener.document.forms[0].hdnProvinceCode.value=strProvince;
		parent.opener.document.forms[0].hdnCity.value= strCity;
		parent.opener.document.forms[0].hdnStreetName.value = strStreet;
	}
	catch(e){}//do nothing, the caller don't need that info
	parent.window.close ();
}

</SCRIPT>
<BODY>

<FORM name=frmACList action="Assetlist.asp">
    <input type=hidden name=txtassetid value="<%=strAssetId%>">
    <input type=hidden name=txttacname value="<%=strTacName %>">
    <input type=hidden name=txtserial value="<%=strSerial%>">
    <input type=hidden name=selassettype value="<%=StrAssetType %>">
    <input type=hidden name=txtassetmake value="<%=StrAssetMake %>">
    <input type=hidden name=txtassetmodel value="<%=StrAssetModel %>">
    <input type=hidden name=selregion value="<%=StrRegion %>">
    <input type=hidden name=txtcustomerName value="<%=strCustomer %>">
    <input type=hidden name=hdnWinName value="<%=strMyWinName%>">
    <input type=hidden name=txtassetcode value="<%=StrAssetCode%>">
    <input type=hidden name="chkactive" value="<%=StrActive%>">
    <input type=hidden name="txtcllicode" value="<%=strClliCode %>">

    <input type="hidden" name="hdnExport" value>

<TABLE border=1 cellPadding=2 cellSpacing=0 width="100%">
<THEAD>
	<TR>
		<TH align=left>Asset ID</TH>
		<TH align=left>Asset Type</TH>
		<TH align=left>PO#</TH>
		<TH align=left>Received Date</TH>
		<TH align=left>Tac Name</TH>
		<TH align=left>Serial#</TH>
		<TH align=left>Make</TH>
		<TH align=left>Model</TH>
		<TH align=left>Part Number</TH>
		<TH align=left>Customer</TH>
		<TH align=left>Address</TH>
		<TH align=left>Region</TH>
		<TH align=left>TAC Asset</TH>
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
	if strMyWinName = "Popup" then
		Response.Write "<TD NOWRAP><a href=# onclick=""go_back("&aList(0,k)&",'"&routineJavascriptString(aList(4,k))&"','"&routineJavascriptString(aList(6,k))&"','"&routineJavascriptString(aList(7,k))&"','"&routineJavascriptString(aList(11,k))&"','"&aList(12,k)&"','"&aList(13,k)&"','"&aList(8,k)&"','"&aList(14,k)&"','"&aList(15,k)&"','"&aList(16,k)&"','"&aList(17,k)&"')"">"&routineHtmlString(aList(0,k))&"&nbsp;</a></TD>"&vbCrLf
		Response.Write "<TD NOWRAP><a href=# onclick=""go_back("&aList(0,k)&",'"&routineJavascriptString(aList(4,k))&"','"&routineJavascriptString(aList(6,k))&"','"&routineJavascriptString(aList(7,k))&"','"&routineJavascriptString(aList(11,k))&"','"&aList(12,k)&"','"&aList(13,k)&"','"&aList(8,k)&"','"&aList(14,k)&"','"&aList(15,k)&"','"&aList(16,k)&"','"&aList(17,k)&"')"">"&routineHtmlString(aList(1,k))&"&nbsp;</a></TD>"&vbCrLf
		Response.Write "<TD NOWRAP><a href=# onclick=""go_back("&aList(0,k)&",'"&routineJavascriptString(aList(4,k))&"','"&routineJavascriptString(aList(6,k))&"','"&routineJavascriptString(aList(7,k))&"','"&routineJavascriptString(aList(11,k))&"','"&aList(12,k)&"','"&aList(13,k)&"','"&aList(8,k)&"','"&aList(14,k)&"','"&aList(15,k)&"','"&aList(16,k)&"','"&aList(17,k)&"')"">"&routineHtmlString(aList(2,k))&"&nbsp;</a></TD>"&vbCrLf
		Response.Write "<TD NOWRAP><a href=# onclick=""go_back("&aList(0,k)&",'"&routineJavascriptString(aList(4,k))&"','"&routineJavascriptString(aList(6,k))&"','"&routineJavascriptString(aList(7,k))&"','"&routineJavascriptString(aList(11,k))&"','"&aList(12,k)&"','"&aList(13,k)&"','"&aList(8,k)&"','"&aList(14,k)&"','"&aList(15,k)&"','"&aList(16,k)&"','"&aList(17,k)&"')"">"&routineHtmlString(aList(3,k))&"&nbsp;</a></TD>"&vbCrLf
		Response.Write "<TD NOWRAP><a href=# onclick=""go_back("&aList(0,k)&",'"&routineJavascriptString(aList(4,k))&"','"&routineJavascriptString(aList(6,k))&"','"&routineJavascriptString(aList(7,k))&"','"&routineJavascriptString(aList(11,k))&"','"&aList(12,k)&"','"&aList(13,k)&"','"&aList(8,k)&"','"&aList(14,k)&"','"&aList(15,k)&"','"&aList(16,k)&"','"&aList(17,k)&"')"">"&routineHtmlString(aList(4,k))&"&nbsp;</a></TD>"&vbCrLf
		Response.Write "<TD NOWRAP><a href=# onclick=""go_back("&aList(0,k)&",'"&routineJavascriptString(aList(4,k))&"','"&routineJavascriptString(aList(6,k))&"','"&routineJavascriptString(aList(7,k))&"','"&routineJavascriptString(aList(11,k))&"','"&aList(12,k)&"','"&aList(13,k)&"','"&aList(8,k)&"','"&aList(14,k)&"','"&aList(15,k)&"','"&aList(16,k)&"','"&aList(17,k)&"')"">"&routineHtmlString(aList(5,k))&"&nbsp;</a></TD>"&vbCrLf
		Response.Write "<TD NOWRAP><a href=# onclick=""go_back("&aList(0,k)&",'"&routineJavascriptString(aList(4,k))&"','"&routineJavascriptString(aList(6,k))&"','"&routineJavascriptString(aList(7,k))&"','"&routineJavascriptString(aList(11,k))&"','"&aList(12,k)&"','"&aList(13,k)&"','"&aList(8,k)&"','"&aList(14,k)&"','"&aList(15,k)&"','"&aList(16,k)&"','"&aList(17,k)&"')"">"&routineHtmlString(aList(6,k))&"&nbsp;</a></TD>"&vbCrLf
		Response.Write "<TD NOWRAP><a href=# onclick=""go_back("&aList(0,k)&",'"&routineJavascriptString(aList(4,k))&"','"&routineJavascriptString(aList(6,k))&"','"&routineJavascriptString(aList(7,k))&"','"&routineJavascriptString(aList(11,k))&"','"&aList(12,k)&"','"&aList(13,k)&"','"&aList(8,k)&"','"&aList(14,k)&"','"&aList(15,k)&"','"&aList(16,k)&"','"&aList(17,k)&"')"">"&routineHtmlString(aList(7,k))&"&nbsp;</a></TD>"&vbCrLf
		Response.Write "<TD NOWRAP><a href=# onclick=""go_back("&aList(0,k)&",'"&routineJavascriptString(aList(4,k))&"','"&routineJavascriptString(aList(6,k))&"','"&routineJavascriptString(aList(7,k))&"','"&routineJavascriptString(aList(11,k))&"','"&aList(12,k)&"','"&aList(13,k)&"','"&aList(8,k)&"','"&aList(14,k)&"','"&aList(15,k)&"','"&aList(16,k)&"','"&aList(17,k)&"')"">"&routineHtmlString(aList(11,k))&"&nbsp;</a></TD>"&vbCrLf
		Response.Write "<TD NOWRAP><a href=# onclick=""go_back("&aList(0,k)&",'"&routineJavascriptString(aList(4,k))&"','"&routineJavascriptString(aList(6,k))&"','"&routineJavascriptString(aList(7,k))&"','"&routineJavascriptString(aList(11,k))&"','"&aList(12,k)&"','"&aList(13,k)&"','"&aList(8,k)&"','"&aList(14,k)&"','"&aList(15,k)&"','"&aList(16,k)&"','"&aList(17,k)&"')"">"&routineHtmlString(aList(8,k))&"&nbsp;</a></TD>"&vbCrLf
		Response.Write "<TD NOWRAP><a href=# onclick=""go_back("&aList(0,k)&",'"&routineJavascriptString(aList(4,k))&"','"&routineJavascriptString(aList(6,k))&"','"&routineJavascriptString(aList(7,k))&"','"&routineJavascriptString(aList(11,k))&"','"&aList(12,k)&"','"&aList(13,k)&"','"&aList(8,k)&"','"&aList(14,k)&"','"&aList(15,k)&"','"&aList(16,k)&"','"&aList(17,k)&"')"">"&routineHtmlString(aList(9,k))&"&nbsp;</a></TD>"&vbCrLf
		Response.Write "<TD NOWRAP><a href=# onclick=""go_back("&aList(0,k)&",'"&routineJavascriptString(aList(4,k))&"','"&routineJavascriptString(aList(6,k))&"','"&routineJavascriptString(aList(7,k))&"','"&routineJavascriptString(aList(11,k))&"','"&aList(12,k)&"','"&aList(13,k)&"','"&aList(8,k)&"','"&aList(14,k)&"','"&aList(15,k)&"','"&aList(16,k)&"','"&aList(17,k)&"')"">"&routineHtmlString(aList(10,k))&"&nbsp;</a></TD>"&vbCrLf
		Response.Write "<TD NOWRAP><a href=# onclick=""go_back("&aList(0,k)&",'"&routineJavascriptString(aList(4,k))&"','"&routineJavascriptString(aList(6,k))&"','"&routineJavascriptString(aList(7,k))&"','"&routineJavascriptString(aList(11,k))&"','"&aList(12,k)&"','"&aList(13,k)&"','"&aList(8,k)&"','"&aList(14,k)&"','"&aList(15,k)&"','"&aList(16,k)&"','"&aList(17,k)&"')"">"&routineHtmlString(aList(18,k))&"&nbsp;</a></TD>"&vbCrLf

	else
		Response.Write "<TD NOWRAP><a target=""_parent"" href=""AssetDetail.asp?asset_id="&aList(0,k)&""">"&routineHtmlString(aList(0,k))&"&nbsp;</a></TD>"&vbCrLf
		Response.Write "<TD NOWRAP><a target=""_parent"" href=""AssetDetail.asp?asset_id="&aList(0,k)&""">"&routineHtmlString(aList(1,k))&"&nbsp;</a></TD>"&vbCrLf
		Response.Write "<TD NOWRAP><a target=""_parent"" href=""AssetDetail.asp?asset_id="&aList(0,k)&""">"&routineHtmlString(aList(2,k))&"&nbsp;</a></TD>"&vbCrLf
		Response.Write "<TD NOWRAP><a target=""_parent"" href=""AssetDetail.asp?asset_id="&aList(0,k)&""">"&routineHtmlString(aList(3,k))&"&nbsp;</a></TD>"&vbCrLf
		Response.Write "<TD NOWRAP><a target=""_parent"" href=""AssetDetail.asp?asset_id="&aList(0,k)&""">"&routineHtmlString(aList(4,k))&"&nbsp;</a></TD>"&vbCrLf
		Response.Write "<TD NOWRAP><a target=""_parent"" href=""AssetDetail.asp?asset_id="&aList(0,k)&""">"&routineHtmlString(aList(5,k))&"&nbsp;</a></TD>"&vbCrLf
		Response.Write "<TD NOWRAP><a target=""_parent"" href=""AssetDetail.asp?asset_id="&aList(0,k)&""">"&routineHtmlString(aList(6,k))&"&nbsp;</a></TD>"&vbCrLf
		Response.Write "<TD NOWRAP><a target=""_parent"" href=""AssetDetail.asp?asset_id="&aList(0,k)&""">"&routineHtmlString(aList(7,k))&"&nbsp;</a></TD>"&vbCrLf
		Response.Write "<TD NOWRAP><a target=""_parent"" href=""AssetDetail.asp?asset_id="&aList(0,k)&""">"&routineHtmlString(aList(11,k))&"&nbsp;</a></TD>"&vbCrLf
		Response.Write "<TD NOWRAP><a target=""_parent"" href=""AssetDetail.asp?asset_id="&aList(0,k)&""">"&routineHtmlString(aList(8,k))&"&nbsp;</a></TD>"&vbCrLf
		Response.Write "<TD NOWRAP><a target=""_parent"" href=""AssetDetail.asp?asset_id="&aList(0,k)&""">"&routineHtmlString(aList(9,k))&"&nbsp;</a></TD>"&vbCrLf
		Response.Write "<TD NOWRAP><a target=""_parent"" href=""AssetDetail.asp?asset_id="&aList(0,k)&""">"&routineHtmlString(aList(10,k))&"&nbsp;</a></TD>"&vbCrLf
		Response.Write "<TD NOWRAP><a target=""_parent"" href=""AssetDetail.asp?asset_id="&aList(0,k)&""">"&routineHtmlString(aList(18,k))&"&nbsp;</a></TD>"&vbCrLf
	end if
	Response.Write "</TR>"
next
%>

</TBODY>
<TFOOT>
<TR>
<TD align=left colSpan=12>
	<input type=hidden name=txtPageNumber value=<%=intPageNumber%>>
	<input type="submit" name="action" value="&lt;&lt;">
	<input type="submit" name="action" value="&lt;">
	<input type="text" name="txtGoToPageNo" onClick="document.frmACList.txtGoToPageNo.value=''" title="You can jump to a specific page by typing the page number in this box" value="page <%=intPageNumber%> of <%=intPageCount%>" style="HEIGHT: 22px; WIDTH: 150px">
	<input type="submit" name="action" value="&gt;">
	<input type="submit" name="action" value="&gt;&gt;">
	<img SRC="images/excel.gif" onclick="document.frmACList.target='new';document.frmACList.hdnExport.value='xls';document.frmACList.submit();document.frmACList.hdnExport.value='';document.frmACList.target='_self';" WIDTH="32" HEIGHT="32">
</TD>
</TR>
</TFOOT>
<CAPTION>Records <%=m+1%> to <%=n+1%> of <%=UBound(aList, 2)+1 & " records"%></CAPTION>
</TABLE>
</FORM>
</BODY>
</HTML>

<%@ Language=VBScript %>
<% option explicit%>
<%on error resume next%>
<!--#include file="smaConstants.inc"-->
<!--#include file="databaseconnect.asp"-->
<!--#include file="smaProcs.inc"-->


<!--Shawn Myers-->



<%
'check the present user's rights

dim intAccessLevel
intAccessLevel = CInt(CheckLogon(strConst_AssetCatalogue))+400
if intAccessLevel < intConst_Access_ReadOnly then
	DisplayError "BACK", "", 0, "ACCESS DENIED", "You don't have access to asset catalogue. Please contact your system administrator."
end if

'declare the connection variables
dim sql
dim rsACList

'declare variable to be used in array for rows
dim aList
'declare the caller variable to be used for gobacks, etc.
dim strMyWinName

'get the caller variable value from the previous page
strMyWinName = Request("hdnWinName")

'declare results variables for fields in bottom navigation
dim intPageNumber
dim intPageCount

'these results variables are used to fill the displayed fields
dim strAssetCatalogueMake
dim strAssetCatalogueModel
dim strAssetCataloguePartNumber


'fill the results variables with data (from both textfields/hiddenfields on previous page)
'these are used to fill the displayed fields
strAssetCatalogueMake = UCase(Trim(Request("txtAssetCatalogueMake")))
strAssetCatalogueModel = UCase(Trim(Request("txtAssetCatalogueModel")))
strAssetCataloguePartNumber = UCase(Trim(Request("txtAssetCataloguePartNumber")))


'connect to the database using the include file
'CONNECT using databaseconnect.asp

'extract the necessary data using sql query

sql = "SELECT " &_
		"DISTINCT AC.ASSET_CATALOGUE_ID, "&_
		"AC.MAKE_ID, " &_
		"AC.MODEL_ID, " &_
		"AC.PART_NUMBER_ID, "&_
		"MA.MAKE_DESC, "&_
		"MO.MODEL_DESC, " &_
		"PN.PART_NUMBER_DESC " &_
	"FROM " &_
		"CRP.ASSET_CATALOGUE			AC, "&_
		"CRP.MAKE						MA, "&_
		"CRP.MODEL						MO, "&_
		"CRP.PART_NUMBER				PN "&_
	"WHERE " &_
		"AC.MAKE_ID = MA.MAKE_ID "&_
		"AND AC.MODEL_ID = MO.MODEL_ID "&_
		"AND AC.PART_NUMBER_ID= PN.PART_NUMBER_ID "&_
		"AND AC.RECORD_STATUS_IND = 'A'"

if strAssetCatalogueMake <> "" then
	sql = sql &	" AND UPPER(MA.MAKE_DESC) LIKE '" &routineOraString(strAssetCatalogueMake)& "%' "
end if

if strAssetCatalogueModel <> "ALL" then
	sql = sql &	" AND UPPER(MO.MODEL_DESC) LIKE '" &routineOraString(strAssetCatalogueModel)& "%' "
end if

if strAssetCataloguePartNumber <> "" then
	sql = sql &	" AND UPPER(PN.PART_NUMBER_DESC) LIKE '" &routineOraString(strAssetCataloguePartNumber)& "%' "
end if


sql = sql & "ORDER BY decode(upper(make_desc),'<NONE>',chr(0),upper(make_desc)), " &_
				     "decode(upper(model_desc),'<NONE>',chr(0),upper(model_desc)), " &_
				     "decode(upper(part_number_desc),'<NONE>',chr(0),upper(part_number_desc))"



'Response.Write (sql & "<p>")
'Response.end

'set the recordset and parse through the data
set rsACList=server.CreateObject("ADODB.Recordset")
rsACList.Open sql, objConn
if err then
	DisplayError "BACK", "", err.Number, "assetcatlist.asp - Cannot open database", err.Description
end if

'search through the recordset and get the data
if not rsACList.EOF then
	aList = rsACList.GetRows
else
	Response.Write "0 records found"
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

                              set objTxtStream = objFSO.CreateTextFile(strExportPath&strRealUserID&"-assetcat.xls", true,false)

								if err then
									DisplayError "CLOSE", "", err.Number, "AssetCatList.asp - Cannot create Excel spreadsheet file due to the following reasons.  Please contact your website administrator.", err.Description
								end if

							  with objTxtStream

                                   .WriteLine "<table border=1>"

                                   'export the table header
                                   .WriteLine "<TR>"

                                   .WriteLine "<TH>Make</TD>"
                                   .WriteLine "<TH>Model</TD>"
                                   .WriteLine "<TH>Part Number</TD>"


								   'end the table header
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

                                         .WriteLine "<TD NOWRAP>"&routineHtmlString(aList(4,k))&"</TD>"
                                         .WriteLine "<TD NOWRAP>"&routineHtmlString(aList(5,k))&"</TD>"
                                         .WriteLine "<TD NOWRAP>"&routineHtmlString(aList(6,k))&"</TD>"

                                         .WriteLine "</TR>"
                                   next
                                   .WriteLine "</table>"

                              end with

                              objTxtStream.Close
								sql = "<script type=""text/javascript"">document.location=""export/"&strRealUserID&"-assetcat.xls"";</script>"
								Response.Write sql
								Response.End

                              'Response.redirect "export/"&strRealUserID&"-assetcat.xls"



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
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Asset Catalog Results</title>
<link href="stylesheets/styles.css" rel="stylesheet" type="text/css">
</head>
<script TYPE="TEXT/JAVASCRIPT">


//NEW

function go_back(lngAssetCatalogID, strMakeDesc, strModelDesc, strPartNumberDesc)
{
	// CR 1550 and 1549 ... leave Make/Model/Part No. blank instead of with "<none>"

	if (strMakeDesc=="<none>")
	{
		strMakeDesc= "";
	}
	if (strModelDesc=="<none>")
	{
		strModelDesc= "";
	}
	if (strPartNumberDesc=="<none>")
	{
		strPartNumberDesc= "";
	}
	parent.opener.document.forms[0].hdnAssetCatalogueID.value = lngAssetCatalogID;
	parent.opener.document.forms[0].txtAssetMake.value = strMakeDesc;
	parent.opener.document.forms[0].txtAssetModel.value = strModelDesc;
	parent.opener.document.forms[0].txtAssetPartNo.value = strPartNumberDesc;

	parent.window.close ();
}

</script>
<body>

<!--hidden fields are filled with values from previous page as well-->
<form name="frmACList" action="assetcatlist.asp" method="POST">
    <input type="hidden" name="txtAssetCatalogueMake" value="<%=strAssetCatalogueMake%>">
    <input type="hidden" name="txtAssetCatalogueModel" value="<%=strAssetCatalogueModel%>">
    <input type="hidden" name="txtAssetCataloguePartNumber" value="<%=strAssetCataloguePartNumber%>">
    <input type="hidden" name="hdntxtAssetCatalogueID" value>
    <input type="hidden" name="hdnWinName" value="<%=strMyWinName%>">
    <input type="hidden" name="hdnExport" value>

<table border="1" cellPadding="2" cellSpacing="0" width="100%">
<thead>
	<tr>
	   <!-- <TH align=left>Catalogue ID</TH> -->
		<th align="left">Make</th>
		<th align="left">Model</th>
		<th align="left">Part Number</th>
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
	if strMyWinName = "Popup" then
		Response.Write "<TD NOWRAP><a href=# onclick=""go_back("&aList(0,k)&",'"&routineJavascriptString(aList(4,k))&"','"&routineJavascriptString(aList(5,k))&"','"&routineJavascriptString(aList(6,k))&"')"">"&routineHtmlString(aList(4,k))&"&nbsp;</a></TD>"&vbCrLf
		Response.Write "<TD NOWRAP><a href=# onclick=""go_back("&aList(0,k)&",'"&routineJavascriptString(aList(4,k))&"','"&routineJavascriptString(aList(5,k))&"','"&routineJavascriptString(aList(6,k))&"')"">"&routineHtmlString(aList(5,k))&"&nbsp;</a></TD>"&vbCrLf
		Response.Write "<TD NOWRAP><a href=# onclick=""go_back("&aList(0,k)&",'"&routineJavascriptString(aList(4,k))&"','"&routineJavascriptString(aList(5,k))&"','"&routineJavascriptString(aList(6,k))&"')"">"&routineHtmlString(aList(6,k))&"&nbsp;</a></TD>"&vbCrLf

	'this second condition is the list that appears normally.
	else
		Response.Write "<TD NOWRAP><a target=""_parent"" href=""assetcatdet.asp?hdntxtAssetCatalogueID="&aList(0,k)&""">"&routineHtmlString(aList(4,k))&"&nbsp;</a></TD>"&vbCrLf
		Response.Write "<TD NOWRAP><a target=""_parent"" href=""assetcatdet.asp?hdntxtAssetCatalogueID="&aList(0,k)&""">"&routineHtmlString(aList(5,k))&"&nbsp;</a></TD>"&vbCrLf
		Response.Write "<TD NOWRAP><a target=""_parent"" href=""assetcatdet.asp?hdntxtAssetCatalogueID="&aList(0,k)&""">"&routineHtmlString(aList(6,k))&"&nbsp;</a></TD>"&vbCrLf
	end if
	Response.Write "</TR>"
next
%>

</tbody>
<tfoot>
<tr>
<td align="left" colSpan="6">
	<input type="hidden" name="txtPageNumber" value="<%=intPageNumber%>">
	<input type="submit" name="action" value="&lt;&lt;">
	<input type="submit" name="action" value="&lt;">
	<input type="text" name="txtGoToPageNo" onClick="document.frmACList.txtGoToPageNo.value=''" title="You can jump to a specific page by typing the page number in this box" value="page <%=intPageNumber%> of <%=intPageCount%>" style="HEIGHT: 22px; WIDTH: 150px">
	<input type="submit" name="action" value="&gt;">
	<input type="submit" name="action" value="&gt;&gt;">
    <img SRC="images/excel.gif" onclick="document.frmACList.target='new';document.frmACList.hdnExport.value='xls';document.frmACList.submit();document.frmACList.hdnExport.value='';document.frmACList.target='_self';" </TD WIDTH="32" HEIGHT="32">
</tr>
</tfoot>
<caption>Records <%=m+1%> to <%=n+1%> of <%=UBound(aList, 2)+1 & " records"%></caption>
</table>
</form>

<%
'close the recordset and the connection objects
rsACList.Close
set rsACList = nothing

objConn.close
set objConn = nothing


%>
</body>
</html>

<%@ Language=VBScript %>
<% option explicit%>
<%on error resume next%>




<!--#include file="smaConstants.inc"-->
<!--#include file="databaseconnect.asp"-->
<!--#include file="smaProcs.inc"-->




<%
'check the present user's rights

dim intAccessLevel

intAccessLevel = CInt(CheckLogon(strConst_AssetTypeClassification))

if intAccessLevel < intConst_Access_ReadOnly then
	DisplayError "BACK", "", 0, "ACCESS DENIED", "You don't have access to asset catalogue. Please contact your system administrator."
end if

'****VARIABLES****

'declare the connection variables
dim sqlSelect, sqlFrom, sqlWhere, sqlRecordStatus, sqlOrderBy, sqlString, rsAssetClassList

'declare variable to be used in array for rows
dim aList

'declare the variables to get info from previous page
dim strMyWinName, strAssetClassDesc, strAssetClassType, bolActiveOnly

'declare results variables for fields in bottom navigation
dim intPageNumber, intPageCount


'request search values from previous page

	strMyWinName = Request("hdnWinName")
	strAssetClassType = Trim(Request("selAssetClassType"))
	strAssetClassDesc = UCase(Trim(Request("txtAssetClassDesc")))
	bolActiveOnly = trim(Request("chkActiveOnly"))

	'extract the necessary data using sql query
	sqlSelect = "SELECT " &_
				"ACT.ASSET_CLASS_TYPE_ID, " &_
				"ACT.ASSET_CLASS_TYPE_DESC, " &_
				"AC.ASSET_CLASS_ID, " &_
				"AC.ASSET_CLASS_DESC "

	sqlFrom = "FROM " &_
			  "CRP.ASSET_CLASS_TYPE ACT, " &_
			  "CRP.ASSET_CLASS AC "


	sqlWhere = "WHERE " &_
			   "ACT.ASSET_CLASS_TYPE_ID = AC.ASSET_CLASS_TYPE_ID "


	If strAssetClassType <> "" then
	      sqlWhere = sqlWhere & " AND ACT.ASSET_CLASS_TYPE_ID = " & strAssetClassType
	End If


	If strAssetClassDesc <> "" then
	      sqlWhere = sqlWhere & " AND Upper(AC.ASSET_CLASS_DESC) LIKE '" & strAssetClassDesc &"%'"
	End If




	'see all records?
	If bolActiveOnly = "yes" then
		sqlRecordStatus = " and AC.RECORD_STATUS_IND = 'A' "
	Else 'nope
		sqlRecordStatus = " "
	End If


sqlString = sqlSelect & sqlFrom & sqlWhere & sqlRecordStatus


'Response.Write sqlstring
'Response.end

'set the recordset and parse through the data
set rsAssetClassList=server.CreateObject("ADODB.Recordset")
rsAssetClassList.Open sqlString, objConn

	if err then
		DisplayError "BACK", "", err.Number, "AssetClassList.asp - Cannot open database", err.Description
	end if

'search through the recordset and get the data

if not rsAssetClassList.EOF then
	aList = rsAssetClassList.GetRows
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

                              set objTxtStream = objFSO.CreateTextFile(strExportPath&strRealUserID&"-assetclass.xls", true,false)

								if err then
										DisplayError "CLOSE", "", err.Number, "AssetClassList.asp - Cannot create Excel spreadsheet file due to the following reasons.  Please contact your website administrator.", err.Description
								end if


							  with objTxtStream

                                   .WriteLine "<table border=1>"

                                   'export the table header
                                   .WriteLine "<TR>"

                                   .WriteLine "<TH>Asset Class</TD>"
                                   .WriteLine "<TH>Asset Class Type</TD>"


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

                                         .WriteLine "<TD NOWRAP>"&routineHtmlString(aList(3,k))&"</TD>"
                                         .WriteLine "<TD NOWRAP>"&routineHtmlString(aList(1,k))&"</TD>"

                                         .WriteLine "</TR>"
                                   next
                                   .WriteLine "</table>"

                              end with

                              objTxtStream.Close
							strstring = "<script type=""text/javascript"">document.location=""export/"&strRealUserID&"-assetclass.xls"";</script>"
							Response.Write strstring
							Response.End
                              'Response.redirect "export/"&strRealUserID&"-assetclass.xls"

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
<title>Asset Class List</title>
<link href="stylesheets/styles.css" rel="stylesheet" type="text/css">
<meta HTTP-EQUIV="Pragma" CONTENT="no-cache">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
</head>
<script TYPE="TEXT/JAVASCRIPT">



//need to complete this function if this screen is used as a lookup
function go_back( lngTypeID, strTypeDesc, lngClassID, strClassDesc)
{
	parent.opener.document.forms[0].txtAssetClassTypeDesc.value = strTypeDesc;
	parent.opener.document.forms[0].hdnAssetClassID.value       = lngClassID;
	parent.opener.document.forms[0].txtAssetClassDesc.value     = strClassDesc;
	parent.window.close ();
}

</script>

<body>


<form name="frmssetClassList" action="AssetClassList.asp">

    <input type="hidden" name="hdnWinName" value="<%=strMyWinName%>">
    <input type="hidden" name="txtAssetClassDesc" value="<%=strAssetClassDesc%>">
    <input type="hidden" name="selAssetClassType" value="<%=strAssetClassType%>">
    <input type="hidden" name="hdnExport" value>

<table border="1" cellPadding="2" cellSpacing="0" width="100%">
	<thead>
		<tr>
			<!-- <TH align=left>Catalogue ID</TH> -->
			<th align="left">Asset Class</th>
			<th align="left">Asset Class Type</th>
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

	if strMyWinName = "Popup" then
	   Response.Write "<td><a href=""#"" onClick=""return go_back(" & routineJavaScriptString(aList(0, k)) & ", '" & routineJavaScriptString(aList(1, k)) & "', " & routineJavaScriptString(aList(2, k)) & ", '" & routineJavaScriptString(aList(3, k)) & "')"">" & aList(3,k) & "</a></td>" & vbCrLf
	   Response.Write "<td><a href=""#"" onClick=""return go_back(" & routineJavaScriptString(aList(0, k)) & ", '" & routineJavaScriptString(aList(1, k)) & "', " & routineJavaScriptString(aList(2, k)) & ", '" & routineJavaScriptString(aList(3, k)) & "')"">" & aList(1,k) & "</a></td>" & vbCrLf

	else
		Response.Write "<TD NOWRAP><a target=""_parent"" href=""assetclassdetail.asp?hdnAssetClassID="&aList(2,k)&""">"&routineHtmlString(aList(3,k))&"&nbsp;</a></TD>"&vbCrLf
		Response.Write "<TD NOWRAP><a target=""_parent"" href=""assetclassdetail.asp?hdnAssetClassID="&aList(2,k)&""">"&routineHtmlString(aList(1,k))&"&nbsp;</a></TD>"&vbCrLf

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
	<input type="text" name="txtGoToPageNo" title="You can jump to a specific page by typing the page number in this box" value="page <%=intPageNumber%> of <%=intPageCount%>" style="HEIGHT: 22px; WIDTH: 150px">
	<input type="submit" name="action" value="&gt;">
	<input type="submit" name="action" value="&gt;&gt;">
	<img SRC="images/excel.gif" onclick="document.frmssetClassList.target='new';document.frmssetClassList.hdnExport.value='xls';document.frmssetClassList.submit();document.frmssetClassList.hdnExport.value='';document.frmssetClassList.target='_self';" </TD WIDTH="32" HEIGHT="32">
</tr>
</tfoot>
<caption>Records <%=m+1%> to <%=n+1%> of <%=UBound(aList, 2)+1 & " records"%></caption>
</table>
</form>

<%

'close the recordset and the connection objects
rsAssetClassList.Close
set rsAssetClassList = nothing

objConn.close
set objConn = nothing


%>
</body>
</html>



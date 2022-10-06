<%@ Language=VBScript   %>
<% option explicit      %>
<% on error resume next %>
<% Response.Buffer = true %>
<!--#include file="smaConstants.inc"-->
<!--#include file="databaseconnect.asp"-->
<!--#include file="smaProcs.inc"-->
<!--********************************************************************************************* Page name:	AssetSubClassList.asp* Purpose:		To dynamically display the results of a search for an asset subclass.*	* In Param:		This page reads following parameters*               selType - the asset class type in the form of id¿Desc*               selClass - the asset class in the form of id¿Desc*				txtDesc - this is the make that is to be searched for (this was named criteris to make cloning for model and part number easier)** Out Param:    The following fields get set in the first form of the calling detail screen:*               hdnID*				txtDesc*               hdnTypeID*               txtTypeDesc*               hdnClassID*               txtClassDesc** Created by:	Chris Roe Oct. 17, 2000*  ********************************************************************************************-->
<%
'check user's rights
Const ASP_NAME    = "AssetSubClassList.asp"
Const DETAIL_PAGE = "AssetSubClassDetail.asp"

dim intAccessLevel
intAccessLevel = CInt(CheckLogon(strConst_AssetTypeClassification))
if (intAccessLevel and intConst_Access_ReadOnly) <> intConst_Access_ReadOnly then
	DisplayError "BACK", "", 0, "ACCESS DENIED", "You don't have access to asset subclasses. Please contact your system administrator."
end if

dim sql, strMyWinName
dim rsList

dim aList, intPageNumber, intPageCount
dim strSubclass, strType, strClass

'get the caller
strMyWinName = Request("hdnWinName")

'get search criteria
strSubclass = UCase(Trim(Request("txtDesc")))
if Trim(Request("selClass")) <> "" then
	strClass = split(Trim(Request("selClass")), strDelimiter)(0)
end if
if Trim(Request("selType")) <> "" then
	strType = Trim(Request("selType"))
end if

'build query
sql = " SELECT sc.asset_sub_class_id" &_
      " ,      sc.asset_sub_class_desc" &_
      " ,      c.asset_class_id" &_
      " ,      c.asset_class_desc" &_
      " ,      t.asset_class_type_id" &_
      " ,      t.asset_class_type_desc" &_
	  " FROM   crp.asset_sub_class  sc" &_
	  " ,      crp.asset_class      c" &_
	  " ,      crp.asset_class_type t" &_
	  " WHERE  t.asset_class_type_id = c.asset_class_type_id" &_
	  " AND    c.asset_class_id = sc.asset_class_id"

if strType <> "" then
	sql = sql & " AND t.asset_class_type_id = " & strType
end if

if strClass <> "" then
	sql = sql & " AND c.asset_class_id = " & strClass
end if

if strSubclass <> "" then
	sql = sql & " AND upper(sc.asset_sub_class_desc) LIKE '" & routineOraString(strSubclass) & "%'"
end if

if Trim(Request("chkActiveOnly")) <> "" then
	sql = sql & " AND sc.record_status_ind = 'A'"
end if

'order by make description
sql = sql & " ORDER BY asset_class_type_desc, asset_class_desc, asset_sub_class_desc"

'get the recordset
set rsList = server.CreateObject("ADODB.Recordset")
rsList.Open sql, objConn
if err then
	DisplayError "BACK", "", err.Number, ASP_NAME & " - Cannot open database", err.Description
end if

if not rsList.EOF then
	aList = rsList.GetRows
else
	Response.Write "0 records found"
	Response.end
end if

'release and kill the recordset and the connection objects
rsList.Close
set rsList = nothing

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
	'Case "Export"
	case else
			if Request("hdnExport") <> "" then
			Response.Write "exporting"
				Dim strRealUserID
				Dim strExportPath
				Dim liLength
				Dim objFSO
				Dim objTxtStream
				strRealUserID = Session("username")
				'determine export path
				strExportPath = Request.ServerVariables("PATH_TRANSLATED")
				Do While (Right(strExportPath, 1) <> "\" And Len(strExportPath) <> 0)
					liLength = Len(strExportPath) - 1
					strExportPath = Left(strExportPath, liLength)
				Loop
				strExportPath = strExportPath & "export\"

				'create scripting object
				Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
				'create export file (overwrite if exists)
				Set objTxtStream = objFSO.CreateTextFile(strExportPath & strRealUserID & "-assetsubclass.xls", True, False)
				if err then
					DisplayError "CLOSE", "", err.Number, ASP_NAME & " - Cannot create Excel spreadsheet file due to the following reasons.  Please contact your website administrator.", err.Description
				end if

				With objTxtStream
					.WriteLine "<TABLE border=1>"
					.WriteLine "<THEAD>"
					.WriteLine "<TH>Asset Class</TH>"
					.WriteLine "<TH>Asset Subclass</TH>"
					.WriteLine "<TH>Asset Type</TH>"
					.WriteLine "</THEAD>"

					'export the body
					For k = 0 To UBound(aList, 2)
						.WriteLine "<TR>"
						.WriteLine "<TD NOWRAP>" & routineHtmlString(aList(3, k)) & "&nbsp;</TD>"
						.WriteLine "<TD NOWRAP>" & routineHtmlString(aList(1, k)) & "&nbsp;</TD>"
						.WriteLine "<TD NOWRAP>" & routineHtmlString(aList(5, k)) & "&nbsp;</TD>"
						.WriteLine "</TR>"
					Next
					.WriteLine "</TABLE>"
				End With
				objTxtStream.Close
				Set objTxtStream = Nothing
				Set objFSO = Nothing
				'Response.Write "<SCRIPT type='text/javascript' language='javascript'>"
				'Response.Write "window.open('" & "export/" & strRealUserID & ".xls" & "');"
				'Response.Write "</SCRIPT>"
'				Response.Write "Redirecting"
				sql = "<script type=""text/javascript"">document.location=""export/"&strRealUserID&"-assetsubclass.xls"";</script>"
				Response.Write sql
				Response.End
'				Response.redirect "export/"&strRealUserID&".xls"
	'case else
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
<meta HTTP-EQUIV="Pragma" CONTENT="no-cache">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Asset Make Results</title>
<link href="stylesheets/styles.css" rel="stylesheet" type="text/css">
</head>
<script TYPE="TEXT/JAVASCRIPT">
function go_back(subclassID, subclassDesc, classID, classDesc, typeID, typeDesc)
{
	parent.opener.document.forms[0].hdnAssetClassTypeID.value = typeID;
	parent.opener.document.forms[0].txtAssetClassTypeDesc.value = typeDesc;
	parent.opener.document.forms[0].hdnAssetClassID.value = classID;
	parent.opener.document.forms[0].txtAssetClassDesc.value = classDesc;
	parent.opener.document.forms[0].hdnAssetSubClassID.value = subclassID;
	parent.opener.document.forms[0].txtAssetSubClassDesc.value = subclassDesc;
	parent.window.close();
}

function btnExcel_onClick()
{
	document.forms[0].target='new';
	document.forms[0].hdnExport.value='xls';
	document.forms[0].submit();
	document.forms[0].hdnExport.value='';
	document.forms[0].target='_self';

}

</script>
<body>

<form name="frmACList" action="<%=ASP_NAME%>" method="POST">
    <input type="hidden" name="hdnWinName" value="<%=strMyWinName%>">
	<input type="hidden" name="selClass" value="<%=Request("selClass")%>">
	<input type="hidden" name="selType" value="<%=Request("selType")%>">
	<input type="hidden" name="txtDesc" value="<%=Request("txtDesc")%>">
<table border="1" cellPadding="2" cellSpacing="0" width="100%">
<thead>
	<tr>
	   <!-- <TH align=left>Catalogue ID</TH> -->
		<th align="left">Asset Class</th>
		<th align="left">Asset Subclass</th>
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
		Response.Write "<TD NOWRAP><a href=# onclick=""go_back(" & aList(0,k) & ",'" & routineJavascriptString(aList(1,k)) & "'," & aList(2, k) & ",'" & routineJavascriptString(aList(3,k)) & "'," & aList(4, k) & ",'" & routineJavascriptString(aList(5,k)) & "')"">"&routineHtmlString(aList(3,k))&"&nbsp;</a></TD>"&vbCrLf
		Response.Write "<TD NOWRAP><a href=# onclick=""go_back(" & aList(0,k) & ",'" & routineJavascriptString(aList(1,k)) & "'," & aList(2, k) & ",'" & routineJavascriptString(aList(3,k)) & "'," & aList(4, k) & ",'" & routineJavascriptString(aList(5,k)) & "')"">"&routineHtmlString(aList(1,k))&"&nbsp;</a></TD>"&vbCrLf
		Response.Write "<TD NOWRAP><a href=# onclick=""go_back(" & aList(0,k) & ",'" & routineJavascriptString(aList(1,k)) & "'," & aList(2, k) & ",'" & routineJavascriptString(aList(3,k)) & "'," & aList(4, k) & ",'" & routineJavascriptString(aList(5,k)) & "')"">"&routineHtmlString(aList(5,k))&"&nbsp;</a></TD>"&vbCrLf
	else
		Response.Write "<TD NOWRAP><a target=""_parent"" href=""" & DETAIL_PAGE & "?hdnAssetSubclassID=" & aList(0,k) & """>"&routineHtmlString(aList(3,k))&"&nbsp;</a></TD>"&vbCrLf
		Response.Write "<TD NOWRAP><a target=""_parent"" href=""" & DETAIL_PAGE & "?hdnAssetSubclassID=" & aList(0,k) & """>"&routineHtmlString(aList(1,k))&"&nbsp;</a></TD>"&vbCrLf
		Response.Write "<TD NOWRAP><a target=""_parent"" href=""" & DETAIL_PAGE & "?hdnAssetSubclassID=" & aList(0,k) & """>"&routineHtmlString(aList(5,k))&"&nbsp;</a></TD>"&vbCrLf
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
	<input type="submit" name="action" value="Export" title="Export this list to Excel">
	<!--<INPUT type="submit" name="action" value="Export" title="Export this list to Excel"> -->
	<img src="images/excel.gif" onclick="btnExcel_onClick();" WIDTH="32" HEIGHT="32">
    <input type="hidden" name="hdnExport" value>

</td>
</tr>
</tfoot>
<caption>Records <%=m+1%> to <%=n+1%> of <%=UBound(aList, 2)+1 & " records"%></caption>
</table>
</form>
</body>
</html>

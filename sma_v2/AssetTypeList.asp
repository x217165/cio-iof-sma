<%@ Language=VBScript %>
<%Option Explicit%>
<%on error resume next%>
<!--#include file = "smaConstants.inc"-->
<!--#include file = "smaProcs.inc"-->
<!--#include file = "databaseconnect.asp" -->
<!--*****************************************************************************************************************************************************************************************-->
<html>
<head>
<META HTTP-EQUIV="Pragma" CONTENT="no-cache">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<link rel="stylesheet" type="text/css" href="stylesheets/styles.css">
<script type="text/javascript" language="javascript" src="AccessLevels.js"></script>
<script type="text/javascript" language="javascript" src="GeneralJavaFunctions.js"></script>
<script type="text/javascript" language="javascript">
<!--

//-->
</script>
<%
Dim strWinName, aList, intPageNumber, intPageCount
Dim strAssetClassTypeID, strAssetClassID, strAssetSubclassID, strAssetTypeID, strAssetTypeDesc
Dim chkActiveOnly
Dim objRs, strSQL, strWhereClause, strOrderBy


	strWinName = Trim(Request("hdnWinName"))
	strAssetClassTypeID = Trim(Request("hdnAssetClassTypeID"))
	strAssetClassID = Trim(Request("hdnAssetClassID"))
	strAssetSubclassID = Trim(Request("hdnAssetSubclassID"))
	strAssetTypeID = Trim(Request("hdnAssetTypeID"))
	strAssetTypeDesc = Ucase(Trim(routineOraString(Request("txtAssetTypeDesc"))))
	chkActiveOnly = Trim(Request("chkActiveOnly"))

	strSQL ="SELECT D.ASSET_TYPE_ID, " &_
			"B.ASSET_CLASS_DESC, " &_
			"C.ASSET_SUB_CLASS_DESC, " &_
			"D.ASSET_TYPE_DESC, " &_
			"A.ASSET_CLASS_TYPE_DESC " &_
			"FROM CRP.ASSET_CLASS_TYPE A, " &_
			"CRP.ASSET_CLASS B, " &_
			"CRP.ASSET_SUB_CLASS C, " &_
			"CRP.ASSET_TYPE D"

	strWhereClause = " WHERE D.ASSET_SUB_CLASS_ID (+) = C.ASSET_SUB_CLASS_ID " &_
					"AND C.ASSET_CLASS_ID (+) = B.ASSET_CLASS_ID " &_
					"AND B.ASSET_CLASS_TYPE_ID (+) = A.ASSET_CLASS_TYPE_ID"

	If Len(strAssetClassTypeID) > 0 Then
		strWhereClause = strWhereClause & " AND A.ASSET_CLASS_TYPE_ID = " & strAssetClassTypeID
	End If

	If Len(strAssetClassID) > 0 Then
		strWhereClause = strWhereClause & " AND B.ASSET_CLASS_ID = " & strAssetClassID
	End If

	If Len(strAssetSubclassID) > 0 Then
		strWhereClause = strWhereClause & " AND C.ASSET_SUB_CLASS_ID = " & strAssetSubclassID
	End If

	If Len(strAssetTypeDesc) > 0 Then
		strWhereClause = strWhereClause & " AND Upper(D.ASSET_TYPE_DESC) LIKE '" &(strAssetTypeDesc) & "%'"
	End If

	If Len(chkActiveOnly) > 0 Then
		strWhereClause = strWhereClause & " AND D.RECORD_STATUS_IND = 'A'"
	End If

	strOrderBy = " ORDER BY A.ASSET_CLASS_TYPE_DESC ASC, B.ASSET_CLASS_DESC ASC, C.ASSET_SUB_CLASS_DESC ASC"
	strSQL = strSQL & strWhereClause & strOrderBy

	'Response.Write (strSQL & "<BR>")
	'Response.End

	Set objRS = objConn.Execute(strSQL)
	If Not objRS.EOF Then
		aList = objRS.GetRows
	Else
		Response.Write "0 records found"
		Response.End
	End If

   'release and kill the recordset and the connection objects
	objRS.Close
	Set objRS = Nothing

	objConn.close
	Set objConn = Nothing

   'calculate page number
	intPageCount = Int(UBound(aList, 2) / intConstDisplayPageSize) + 1
	Select Case Request("Action")
		Case "<<"	intPageNumber = 1
		Case "<"	intPageNumber = Request("txtPageNumber") - 1
					If intPageNumber < 1 Then intPageNumber = 1
		Case ">"	intPageNumber = Request("txtPageNumber") + 1
					If intPageNumber > intPageCount Then intPageNumber = intPageCount
		Case ">>"	intPageNumber = intPageCount
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
						set objTxtStream = objFSO.CreateTextFile(strExportPath&strRealUserID&"-assetType.xls", true, false)

						if err then
							DisplayError "CLOSE", "", err.Number, "AssetTypeList.asp - Cannot create Excel spreadsheet file due to the following reasons. Please contact your website administrator.", err.Description
						end if

						with objTxtStream
							.WriteLine "<table border=1>"

							'export the header
							.WriteLine "<TR>"
							.WriteLine "<TH>Asset Class</TD>"
							.WriteLine "<TH>Asset Subclass</TD>"
							.WriteLine "<TH>Asset Type</TD>"
							.WriteLine "<TH>Asset Class Type</TD>"
							.WriteLine "</TR>"

							'export the body
							for k = 0 to UBound(aList, 2)
								.WriteLine "<TR>"
								.WriteLine "<TD NOWRAP>"&routineHtmlString(aList(1,k))&"&nbsp;</TD>"
								.WriteLine "<TD NOWRAP>"&routineHtmlString(aList(2,k))&"&nbsp;</TD>"
								.WriteLine "<TD NOWRAP>"&routineHtmlString(aList(3,k))&"&nbsp;</TD>"
								.WriteLine "<TD NOWRAP>"&routineHtmlString(aList(4,k))&"&nbsp;</TD>"
								.WriteLine "</TR>"
							next
							.WriteLine "</table>"
						end with
						objTxtStream.Close
						strsql = "<script type=""text/javascript"">document.location=""export/"&strRealUserID&"-assetType.xls"";</script>"
						Response.Write strsql
						Response.End
						Response.redirect "export/"&strRealUserID&"-assetType.xls"
					elseif Request("txtGoToPageNo") <> "" then
						intPageNumber = CInt(Request("txtGoToPageNo"))
					else
						intPageNumber = 1
					end if
	end select

	If intPageNumber < 1 Then
		intPageNumber = 1
	End If
	If intPageNumber > intPageCount Then
		intPageNumber = intPageCount
	End If

	Dim k, m, n
	m = (intPageNumber - 1 ) * intConstDisplayPageSize
	n = (intPageNumber * intConstDisplayPageSize) - 1
	If n > UBound(aList, 2) Then
		n = UBound(aList, 2)
	End If

	'check If the client is still connected just before sending any HTML To the browser
	If Not Response.IsClientConnected Then Response.End

	'catch any unexpected error
	If err Then	DisplayError "BACK", "", err.Number, "Unexpected error", err.Description
%>
</head>
<body language="javascript">
<form name="frmAssetTypeList" method="post" action="AssetTypeList.asp">

	<input type="hidden" name="hdnAssetClassTypeID" value="<%=strAssetClassTypeID%>">
	<input type="hidden" name="hdnAssetClassID"		value="<%=strAssetClassID%>">
	<input type="hidden" name="hdnAssetSubclassID"	value="<%=strAssetSubclassID%>">
	<input type="hidden" name="txtAssetTypeDesc"	value="<%=strAssetTypeDesc%>">
	<input type="hidden" name="chkActiveOnly"		value="<%=chkActiveOnly%>">
	<INPUT type="hidden" name="hdnExport"			value>

<table border="1" cellpadding="2" cellspacing="0" cols="4" width="100%">
 <thead>
	<td align="left">Asset Class</td>
	<td align="left">Asset Subclass</td>
	<td align="left">Asset Type</td>
	<td align="left">Asset Class Type</td>
 </thead>
 <tbody>
<%
'Response.Write("WinName:" & strWinName)
For k = m To n
	If Int(k/2) = k/2 Then
		Response.Write "<TR>"
	Else
		Response.Write "<TR bgcolor='white'>"
	End If

	If UCase(strWinName) <> UCase("Popup") Then
		Response.Write "<td nowrap><a target=""_parent"" href=""AssetTypeDetail.asp?AssetTypeID="&aList(0,k)&""">"&aList(1,k)&"</a>&nbsp;</td>"&vbCrLf
		Response.Write "<td nowrap><a target=""_parent"" href=""AssetTypeDetail.asp?AssetTypeID="&aList(0,k)&""">"&aList(2,k)&"</a>&nbsp;</td>"&vbCrLf
		Response.Write "<td nowrap><a target=""_parent"" href=""AssetTypeDetail.asp?AssetTypeID="&aList(0,k)&""">"&aList(3,k)&"</a>&nbsp;</td>"&vbCrLf
		Response.Write "<td nowrap><a target=""_parent"" href=""AssetTypeDetail.asp?AssetTypeID="&aList(0,k)&""">"&aList(4,k)&"</a>&nbsp;</td>"&vbCrLf
		Response.Write "</tr>"
	Else
		'currently this page is never called as a popup
	End If
Next
%>
</tbody>
<tfoot>
<tr>
<td align="left" colspan="4">
	<input type="hidden" name="txtPageNumber" value="<%=intPageNumber%>">
	<input type="submit" name="action" value="&lt;&lt;">
	<input type="submit" name="action" value="&lt;">
	<input type="text" name="txtGoToPageNo" title="You can jump To a specific page by typing the page number in this box" value="page <%=intPageNumber%> of <%=intPageCount%>" style="HEIGHT: 22px; WIDTH: 150px">
	<input type="submit" name="action" value="&gt;">
	<input type="submit" name="action" value="&gt;&gt;">&nbsp;&nbsp;
	<img SRC="images/excel.gif" onclick="document.frmAssetTypeList.target='new';document.frmAssetTypeList.hdnExport.value='xls';document.frmAssetTypeList.submit();document.frmAssetTypeList.hdnExport.value='';document.frmAssetTypeList.target='_self';" WIDTH="32" HEIGHT="32">
</td>
</tr>
</tfoot>
<caption>Records <%=m+1%> to <%=n+1%> of <%=UBound(aList, 2) + 1 & " records"%></caption>
</table>
</form>
</body>
</html>

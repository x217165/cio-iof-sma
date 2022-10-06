<%@ Language=VBScript %>
<%Option Explicit%>
<%Response.Buffer = True%>
<!--#include file = "smaConstants.inc"-->
<!--#include file = "smaProcs.inc"-->
<!--#include file = "databaseconnect.asp" -->
<!--
*****************************************************************************************
*
*
*
*
*
*
******************************************************************************************
-->
<HTML>
<HEAD>
<META name="Generator" content="Microsoft Visual Studio 6.0">
<META http-equiv="Pragma" content="no-cache">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<LINK rel="stylesheet" type="text/css" href="stylesheets/styles.css">
<SCRIPT type="text/javascript" language="javascript" src="AccessLevels.js"></SCRIPT>
<SCRIPT type="text/javascript" language="javascript" src="GeneralJavaFunctions.js"></SCRIPT>
<SCRIPT type="text/javascript" language="javascript" id="clientEventHandlersJS">
<!--
function go_back(strScheduleID, strScheduleDescription ){
//************************************************************************************************
// Function:	return go_back
//
// Purpose:		To write the values of selected row into the base window that called the lookup
//				function. In addition, this function closes the Popup window.
//
// Created By:	Gilles Archer Oct 05 2000
//
// Updated By:
//************************************************************************************************
	with (parent.opener.document.forms[0]) {
		hdnScheduleID.value = strScheduleID;
		txtScheduleDescription.value = strScheduleDescription;
	}
	parent.window.close();
}
//-->
</SCRIPT>
<%
Dim objFSO, objTxtStream, strExportPath, liLength
Dim strRealUserID, aList, intPageNumber, intPageCount
Dim strWinName, strScheduleID, strScheduleDesc, strProvinceCOde, strCountryCode, chkHolidayOnly, chkActiveOnly
Dim	objRs, objCommand, strSQL, strWhere, strOrderBy, strErrMessage, lIndex

	strWinName = Request.Form("hdnWinName")
	strScheduleID = Request("ScheduleID")
	strScheduleDesc = UCase(Trim(Request.Form("txtScheduleDescription")))
	strProvinceCode = Request.Form("selProvince")
	strCountryCode = Request.Form("selCountry")
	chkHolidayOnly = Request.Form("chkHolidayOnly")
	chkActiveOnly = Request.Form("chkActiveOnly")

	strSQL = "SELECT S.SCHEDULE_ID,	S.SCHEDULE_NAME, " &_
			"S.INCLUDE_HOLIDAY_FLAG, P.PROVINCE_STATE_NAME, C.COUNTRY_DESC " &_
			"FROM CRP.SCHEDULE S, CRP.LCODE_PROVINCE_STATE P, CRP.LCODE_COUNTRY C " &_
			"WHERE S.PROVINCE_STATE_LCODE = P.PROVINCE_STATE_LCODE (+)" &_
			"AND S.COUNTRY_LCODE = C.COUNTRY_LCODE (+)"

	If Len(strScheduleID) <> 0 Then
		strWhere = strWhere & " AND S.SCHEDULE_ID = " & strScheduleID
	End If

	If Len(strScheduleDesc) <> 0 Then
		strWhere = strWhere & " AND UPPER(S.SCHEDULE_NAME) LIKE '" & Replace(strScheduleDesc, "'", "''") & "%'"
	End If

	If Len(strProvinceCode) <> 0 Then
		strWhere = strWhere & " AND S.PROVINCE_STATE_LCODE = '" & strProvinceCode & "'"
	End If

	If Len(strCountryCode) <> 0 Then
		strWhere = strWhere & " AND S.COUNTRY_LCODE = '" & strCountryCode & "'"
	End If

	If Len(chkHolidayOnly) <> 0 Then
		strWhere = strWhere & " AND S.INCLUDE_HOLIDAY_FLAG = 'Y'"
	End If

	If Len(chkActiveOnly) <> 0 Then
		strWhere = strWhere & " AND S.RECORD_STATUS_IND = 'A'"
	End If

	strOrderBy = " ORDER BY C.COUNTRY_DESC ASC, P.PROVINCE_STATE_NAME ASC, S.SCHEDULE_NAME ASC"

	strSQL = strSQL & strWhere & strOrderBy

	Set objRS = objConn.Execute(StrSql)
	If Not objRS.EOF Then
		aList = objRS.GetRows
		'release and kill the recordset and the connection objects
		objRS.Close
		Set objRS = Nothing
		objConn.Close
		Set objConn = Nothing
	Else
		Response.Write "0 records found"
		Response.End
	End If

   'calculate page number
	intPageCount = Int(UBound(aList, 2) / intConstDisplayPageSize) + 1
	Select Case Request("Action")
		Case "<<"	intPageNumber = 1
		Case "<"	intPageNumber = Request("txtPageNumber") - 1
					If intPageNumber < 1 Then intPageNumber = 1
		Case ">"	intPageNumber = Request("txtPageNumber") + 1
					If intPageNumber > intPageCount Then intPageNumber = intPageCount
		Case ">>"	intPageNumber = intPageCount
		Case Else	If Request("hdnExport") <> "" then
		'Case "Export"
					strRealUserID = Session("username")
					'determine export path
					strExportPath = Request.ServerVariables("PATH_TRANSLATED")
					Do While (Right(strExportPath, 1) <> "\" And Len(strExportPath) <> 0)
						liLength = Len(strExportPath) - 1
						strExportPath = Left(strExportPath, liLength)
					Loop
					strExportPath = strExportPath & "export\"

					'create scripting object
					On Error Resume Next
					Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
					'create export file (overwrite if exists)
					Set objTxtStream = objFSO.CreateTextFile(strExportPath & strRealUserID & "-schedule.xls", True, False)
					If err Then
						DisplayError "CLOSE", err.Number, "ScheduleList.asp - Cannot create Excel spreadsheet file due to the following errors.  Please contact your system administrator.", err.Description
					End If

					With objTxtStream
						.WriteLine "<TABLE border=1>"
						.WriteLine "<THEAD>"
						.WriteLine "<TH>Schedule</TH>"
						.WriteLine "<TH>Holiday</TH>"
						.WriteLine "<TH>Province / State</TH>"
						.WriteLine "<TH>Country</TH>"
						.WriteLine "</THEAD>"

						'export the body
						For k = 0 To UBound(aList, 2)
							.WriteLine "<TR>"
							.WriteLine "<TD NOWRAP>" & routineHtmlString(aList(1, k)) & "&nbsp;</TD>"
							.WriteLine "<TD NOWRAP>" & routineHtmlString(aList(2, k)) & "&nbsp;</TD>"
							.WriteLine "<TD NOWRAP>" & routineHtmlString(aList(3, k)) & "&nbsp;</TD>"
							.WriteLine "<TD NOWRAP>" & routineHtmlString(aList(4, k)) & "&nbsp;</TD>"
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
						strsql = "<script type=""text/javascript"">document.location=""export/"&strRealUserID&"-schedule.xls"";</script>"
						Response.Write strsql
						Response.End
'					Response.Redirect "export/" & strRealUserID & "-schedule.xls"

		'Case Else
					ElseIf Request("txtGoToPageNo") <> "" Then
						intPageNumber = CInt(Request("txtGoToPageNo"))
					Else
						intPageNumber = 1
					End If
	End Select

	If intPageNumber < 1 Then intPageNumber = 1
	If intPageNumber > intPageCount Then intPageNumber = intPageCount

	Dim k, m, n
	m = (intPageNumber - 1 ) * intConstDisplayPageSize
	n = (intPageNumber * intConstDisplayPageSize) - 1
	If n > UBound(aList, 2) Then n = UBound(aList, 2)

	'check If the client is still connected just before sending any HTML to the browser
	If Not Response.IsClientConnected Then Response.End

	'catch any unexpected error
	If err Then DisplayError "BACK", "", err.Number, "Unexpected error", err.Description
%>
</HEAD>
<BODY language="javascript">
<FORM id="frmScheduleList" name="frmScheduleList" method="post" action="ScheduleList.asp">
	<INPUT type="hidden" id="hdnWinName" name="hdnWinName" value="<%=strWinName%>">
	<INPUT type="hidden" id="txtScheduleID" name="txtScheduleID" value="<%=strScheduleID%>">
	<INPUT type="hidden" id="txtScheduleDescription" name="txtScheduleDescription" value="<%=strScheduleDesc%>">
	<INPUT type="hidden" id="selProvince" name="selProvince" value="<%=strProvinceCode%>">
	<INPUT type="hidden" id="selCountry" name="selCountry" value="<%=strCountryCode%>">
	<INPUT type="hidden" id="chkHolidayOnly" name="chkHolidayOnly" value="<%=chkHolidayOnly%>">
	<INPUT type="hidden" id="chkActiveOnly" name="chkActiveOnly" value="<%=chkActiveOnly%>">
	<INPUT type="hidden" id="hdnExport" name="hdnExport" value="">
<TABLE border="1" cellpadding="2" cellspacing="0" cols="4" width="100%">
<THEAD>
	<TR>
		<TH align="left" nowrap>Schedule</TH>
		<TH align="left" nowrap>Holiday</TH>
		<TH align="left" nowrap>Province / State</TH>
		<TH align="left" nowrap>Country</TH>
	</TR>
</THEAD>
<%
For k = m to n
	If Int(k/2) = k/2 Then
		Response.Write "<TR>"
	Else
		Response.Write "<TR bgcolor='white'>"
	End If

	If UCase(strWinName) <> UCase("Popup") Then%>
		<TD align="left" nowrap><A href="ScheduleDetail.asp?ScheduleID=<%=routineHtmlString(aList(0, k))%>" target="_parent"><%=routineHtmlString(aList(1, k))%>&nbsp;</A></TD>
		<TD align="left" nowrap><A href="ScheduleDetail.asp?ScheduleID=<%=routineHtmlString(aList(0, k))%>" target="_parent"><%=routineHtmlString(aList(2, k))%>&nbsp;</A></TD>
		<TD align="left" nowrap><A href="ScheduleDetail.asp?ScheduleID=<%=routineHtmlString(aList(0, k))%>" target="_parent"><%=routineHtmlString(aList(3, k))%>&nbsp;</A></TD>
		<TD align="left" nowrap><A href="ScheduleDetail.asp?ScheduleID=<%=routineHtmlString(aList(0, k))%>" target="_parent"><%=routineHtmlString(aList(4, k))%>&nbsp;</A></TD></TR>
	<%Else%>
		<TD align="left" nowrap><A href="#" onClick="return go_back('<%=routineHtmlString(List(0, k))%>', '<%=routineHtmlString(aList(1, k))%>');"><%=routineHtmlString(aList(1, k))%>&nbsp;</A></TD>
		<TD align="left" nowrap><A href="#" onClick="return go_back('<%=routineHtmlString(List(0, k))%>', '<%=routineHtmlString(aList(1, k))%>');"><%=routineHtmlString(aList(2, k))%>&nbsp;</A></TD>
		<TD align="left" nowrap><A href="#" onClick="return go_back('<%=routineHtmlString(List(0, k))%>', '<%=routineHtmlString(aList(1, k))%>');"><%=routineHtmlString(aList(3, k))%>&nbsp;</A></TD>
		<TD align="left" nowrap><A href="#" onClick="return go_back('<%=routineHtmlString(List(0, k))%>', '<%=routineHtmlString(aList(1, k))%>');"><%=routineHtmlString(aList(4, k))%>&nbsp;</A></TD></TR>
	<%End If
Next
%>
</TBODY>
<TFOOT>
<TR>
<TD align="left" colspan="4">
	<INPUT type="hidden" name="txtPageNumber" value="<%=intPageNumber%>">
	<INPUT type="submit" name="action" value="&lt;&lt;">
	<INPUT type="submit" name="action" value="&lt;">
	<INPUT type="text" name="txtGoToPageNo" onClick="document.forms[0].txtGoToPageNo.value=''" title="You can jump to a specific page by typing the page number in this box" value="page <%=intPageNumber%> of <%=intPageCount%>" style="HEIGHT: 22px; WIDTH: 150px">
	<INPUT type="submit" name="action" value="&gt;">
	<INPUT type="submit" name="action" value="&gt;&gt;">
<!--	<INPUT type="submit" name="action" value="Export" title="Export this list to Excel"> -->
	<IMG src="images/excel.gif" onClick="document.forms[0].target='new';document.forms[0].hdnExport.value='xls';document.forms[0].submit();document.forms[0].hdnExport.value='';document.forms[0].target='_self';" width="32" height="32">
</TD>
</TR>
</TFOOT>
<CAPTION>Records <%=m+1%> to <%=n+1%> of <%=UBound(aList, 2) + 1 & " records"%></CAPTION>
</TABLE>
</FORM>
</BODY>
</HTML>

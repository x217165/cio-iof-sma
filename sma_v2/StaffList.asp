<%@ Language=VBSCRIPT %>
<%Option Explicit%>
<%Response.Buffer = True%>
<!--#include file = "smaConstants.inc"-->
<!--#include file = "smaProcs.inc"-->
<!--#include file = "databaseconnect.asp" -->
<!--
***************************************************************************************************
* Name:			StaffList.asp
*
* Purpose:		To display the results of a Staff search.
*				Search criteria are chosen via StaffCriteria.asp
* Created By:	Gilles Archer Oct 23 2000
***************************************************************************************************
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
//-->
</SCRIPT>
<%
Dim objFSO, objTxtStream, strExportPath, liLength
Dim strRealUserID, aList, intPageNumber, intPageCount
Dim strWinName, strNameLast, strNameFirst, strEmail, strWPhoneArea, strWPhoneMid, strWPhoneEnd, strWPhone, strEmpNo, strDepartment, strStatus, chkActiveOnly, strUserID
Dim objRS, strSQL, strWhere, strOrderBy

	'get submitted values (search criteria; window name)
	strWinName = Trim(Request.Form("hdnWinName"))

	strNameLast = UCase(Trim(Request.Form("txtNameLast")))
	strNameFirst = UCase(Trim(Request.Form("txtNameFirst")))
	strEmail = UCase(Trim(Request.Form("txtEmail")))
	strWPhoneArea = Trim(Request.Form("txtWPhoneArea"))
	strWPhoneMid = Trim(Request.Form("txtWPhoneMid"))
	strWPhoneEnd = Trim(Request.Form("txtWPhoneEnd"))
	strWPhone = strWPhoneArea & strWPhoneMid & strWPhoneEnd
	strEmpNo = Trim(Request.Form("txtEmpNo"))
	If Len(strEmpNo) <> 0 Then
		Do While Len(strEmpNo) < 8
			strEmpNo = "0" & strEmpNo
		Loop
	End If
	strDepartment = Trim(Request.Form("selDepartment"))
	strStatus = Trim(Request.Form("selStatus"))
	chkActiveOnly = Request.Form("chkActiveOnly")
	strUserID = UCase(Trim(Request("txtUserID")))

	'build query
	strSQL = "SELECT " &_
			"CO.CONTACT_ID, " &_
			"CO.LAST_NAME, " &_
			"CO.FIRST_NAME, " &_
			"CO.EMAIL_ADDRESS, " &_
			"DECODE(CO.WORK_NUMBER, null, '&nbsp;', '(' || SUBSTR(CO.WORK_NUMBER, 1, 3) || ') ' || SUBSTR(CO.WORK_NUMBER, 4, 3) || '-' || SUBSTR(CO.WORK_NUMBER, 7, 4)) AS WORK_NUMBER, " &_
			"CO.WORK_NUMBER_EXT, " &_
			"CO.POSITION_TITLE, " &_
			"CO.EMPLOYEE_NUMBER, " &_
			"AD.BUILDING_NAME, " &_
			"AD.LONG_STREET_NAME, " &_
			"AD.MUNICIPALITY_NAME, " &_
			"PS.PROVINCE_STATE_NAME " &_
			"FROM CRP.CONTACT CO, " &_
			"CRP.ADDRESS AD, " &_
			"CRP.LCODE_PROVINCE_STATE PS"

	strWhere = " WHERE CO.ADDRESS_ID = AD.ADDRESS_ID (+) " &_
			"AND AD.PROVINCE_STATE_LCODE = PS.PROVINCE_STATE_LCODE (+) " &_
			"AND CO.STAFF_FLAG = 'Y'"

	'add other search parameters to the where clause
	If Len(strNameLast) <> 0 Then
		strWhere = strWhere & " AND UPPER(CO.LAST_NAME) LIKE '" & Replace(strNameLast, "'", "''") & "%'"
	End If

	If Len(strNameFirst) <> 0 Then
		strWhere = strWhere & " AND UPPER(CO.FIRST_NAME) LIKE '" & Replace(strNameFirst, "'", "''") & "%'"
	End If

	If Len(strEmail) <> 0 Then
		strWhere = strWhere &  " AND UPPER(CO.EMAIL_ADDRESS) LIKE '" & Replace(strEmail, "'", "''") & "%'"
	End If

	If Len(strWPhone) <> 0 Then
		If Len(strWPhone) = 10 Then
			strWhere = strWhere & " AND CO.WORK_NUMBER = '" & strWPhone & "'"
		Else
			strWhere = strWhere & " AND CO.WORK_NUMBER LIKE '" & strWPhone & "%'"
		End If
    End If

	If Len(strEmpNo) <> 0 Then
		strWhere = strWhere & " AND CO.EMPLOYEE_NUMBER = '" & strEmpNo & "'"
	End If

	If IsNumeric(strDepartment) Then
		strWhere = strWhere & " AND CO.DEPARTMENT_ID = " & strDepartment
	End If

	If Len(strStatus) <> 0 Then
		strWhere = strWhere & " AND CO.STAFF_STATUS_LCODE = '" & strStatus & "'"
	End If

	If Len(strUserID) <> 0 Then
		strWhere = strWhere & " AND Upper(co.userid) LIKE '" & routineOraString(strUserid) & "%'"
	End If

	If Len(chkActiveOnly) <> 0 Then
		strWhere = strWhere & " AND CO.RECORD_STATUS_IND = 'A'"
	End If

	strOrderBy = " ORDER BY CO.CONTACT_NAME ASC"

	'join all pieces to make a complete query
	strSQL = strSQL & strWhere & strOrderBy

	'Response.Write strSQL
	'Response.End

    'get the recordset
    Set objRS = Server.CreateObject("ADODB.Recordset")
    objRS.Open strSQL, objConn, adOpenForwardOnly, adLockReadOnly, adCmdText
    If err Then DisplayError "BACK", "", err.Number, "StaffList.asp - Cannot open database" , err.Description

	'put recordset into array
	If Not objRS.EOF Then
		aList = objRS.GetRows
		'release and kill the recordset and the connection objects
		objRS.Close
		Set objRS = Nothing
		objConn.Close
		Set objConn = Nothing
	Else
		Response.Write "0 Records Found"
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
		Case ">>"	intPageNumber=intPageCount
		Case Else	If Request("hdnExport") <> "" Then
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
					Set objTxtStream = objFSO.CreateTextFile(strExportPath & strRealUserID & "-staff.xls", True, False)
					If err Then
						DisplayError "CLOSE", err.Number, "StaffList.asp - Cannot create Excel spreadsheet file due to the following errors.  Please contact your system administrator.", err.Description
					End If

					With objTxtStream
						.WriteLine "<TABLE border=1>"
						.WriteLine "<THEAD>"
						.WriteLine "<TH>Last Name</TH>"
						.WriteLine "<TH>First Name</TH>"
						.WriteLine "<TH>Email Address</TH>"
						.WriteLine "<TH>Work Phone</TH>"
						.WriteLine "<TH>Ext</TH>"
						.WriteLine "<TH>Position</TH>"
						.WriteLine "<TH>Emp No</TH>"
						.WriteLine "<TH>Building</TH>"
						.WriteLine "<TH>Address</TH>"
						.WriteLine "<TH>City</TH>"
						.WriteLine "<TH>Prov/State</TH>"
						.WriteLine "</THEAD>"

						'export the body
						For k = 0 To UBound(aList, 2)
							.WriteLine "<TR>"
							.WriteLine "<TD NOWRAP>" & routineHtmlString(aList(1, k)) & "&nbsp;</TD>"
							.WriteLine "<TD NOWRAP>" & routineHtmlString(aList(2, k)) & "&nbsp;</TD>"
							.WriteLine "<TD NOWRAP>" & routineHtmlString(aList(3, k)) & "&nbsp;</TD>"
							.WriteLine "<TD NOWRAP>" & routineHtmlString(aList(4, k)) & "&nbsp;</TD>"
							.WriteLine "<TD NOWRAP>" & routineHtmlString(aList(5, k)) & "&nbsp;</TD>"
							.WriteLine "<TD NOWRAP>" & routineHtmlString(aList(6, k)) & "&nbsp;</TD>"
							.WriteLine "<TD NOWRAP>" & routineHtmlString(aList(7, k)) & "&nbsp;</TD>"
							.WriteLine "<TD NOWRAP>" & routineHtmlString(aList(8, k)) & "&nbsp;</TD>"
							.WriteLine "<TD NOWRAP>" & routineHtmlString(aList(9, k)) & "&nbsp;</TD>"
							.WriteLine "<TD NOWRAP>" & routineHtmlString(aList(10, k)) & "&nbsp;</TD>"
							.WriteLine "<TD NOWRAP>" & routineHtmlString(aList(11, k)) & "&nbsp;</TD>"
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
						strsql = "<script type=""text/javascript"">document.location=""export/"&strRealUserID&"-staff.xls"";</script>"
						Response.Write strsql
						Response.End
'					Response.Redirect "export/" & strRealUserID & ".xls"

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

	'check if the client is still connected just before sending any html to the browser
	If Not Response.IsClientConnected Then Response.End

	'catch any unexpected error
	If err Then DisplayError "BACK", "", err.Number, "Unexpected error", err.Description
%>
</HEAD>
<BODY>
<FORM id="frmStaffList" name="frmStaffList" method="post" action="StaffList.asp">
	<INPUT type="hidden" id="hdnWinName" name="hdnWinName" value="<%=strWinName%>">
	<INPUT type="hidden" id="txtNameLast" name="txtNameLast" value="<%=strNameLast%>">
	<INPUT type="hidden" id="txtNameFirst" name="txtNameFirst" value="<%=strNameFirst%>">
	<INPUT type="hidden" id="txtWPhoneArea" name="txtWPhoneArea" value="<%=strWPhoneArea%>">
	<INPUT type="hidden" id="txtWPhoneMid" name="txtWPhoneMid" value="<%=strWPhoneMid%>">
	<INPUT type="hidden" id="txtWPhoneEnd" name="txtWPhoneEnd" value="<%=strWPhoneEnd%>">
	<INPUT type="hidden" id="txtEmail" name="txtEmail" value="<%=strEmail%>">
	<INPUT type="hidden" id="txtEmpNo" name="txtEmpNo" value="<%=strEmpNo%>">
	<INPUT type="hidden" id="selDepartment" name="selDepartment" value="<%=strDepartment%>">
	<INPUT type="hidden" id="selStatus" name="selStatus" value="<%=strStatus%>">
	<INPUT type="hidden" id="chkActiveOnly" name="chkActiveOnly" value="<%=chkActiveOnly%>">
	<INPUT type="hidden" id="hdnExport" name="hdnExport" value="">
	<INPUT type="hidden" id="txtUserID" name="txtUserID" value="<%=strUserID%>">
<TABLE border="1" cellpadding="2" cellspacing="0" cols="11" width="100%">
<THEAD>
	<TR>
		<TH align="left" nowrap>Last Name</TH>
		<TH align="left" nowrap>First Name</TH>
		<TH align="left" nowrap>Email Address</TH>
		<TH align="left" nowrap>Work Phone</TH>
		<TH align="left" nowrap>Ext</TH>
		<TH align="left" nowrap>Position</TH>
		<TH align="left" nowrap>Emp No</TH>
		<TH align="left" nowrap>Building</TH>
		<TH align="left" nowrap>Address</TH>
		<TH align="left" nowrap>City</TH>
		<TH align="left" nowrap>Prov/State</TH>
	</TR>
</THEAD>
<TBODY>
<%
For k = m To n
	If Int(k/2) = k/2 Then
		Response.Write "<TR>"
	Else
		Response.Write "<TR bgcolor='white'>"
	End If%>
		<TD nowrap><A href="StaffDetail.asp?ContactID=<%=routineHtmlString(aList(0, k))%>" target="_parent"><%=routineHtmlString(aList(1, k))%>&nbsp;</A></TD>
		<TD nowrap><A href="StaffDetail.asp?ContactID=<%=routineHtmlString(aList(0, k))%>" target="_parent"><%=routineHtmlString(aList(2, k))%>&nbsp;</A></TD>
		<TD nowrap><A href="StaffDetail.asp?ContactID=<%=routineHtmlString(aList(0, k))%>" target="_parent"><%=routineHtmlString(aList(3, k))%>&nbsp;</A></TD>
		<TD nowrap><A href="StaffDetail.asp?ContactID=<%=routineHtmlString(aList(0, k))%>" target="_parent"><%=routineHtmlString(aList(4, k))%>&nbsp;</A></TD>
		<TD nowrap><A href="StaffDetail.asp?ContactID=<%=routineHtmlString(aList(0, k))%>" target="_parent"><%=routineHtmlString(aList(5, k))%>&nbsp;</A></TD>
		<TD nowrap><A href="StaffDetail.asp?ContactID=<%=routineHtmlString(aList(0, k))%>" target="_parent"><%=routineHtmlString(aList(6, k))%>&nbsp;</A></TD>
		<TD nowrap><A href="StaffDetail.asp?ContactID=<%=routineHtmlString(aList(0, k))%>" target="_parent"><%=routineHtmlString(aList(7, k))%>&nbsp;</A></TD>
		<TD nowrap><A href="StaffDetail.asp?ContactID=<%=routineHtmlString(aList(0, k))%>" target="_parent"><%=routineHtmlString(aList(8, k))%>&nbsp;</A></TD>
		<TD nowrap><A href="StaffDetail.asp?ContactID=<%=routineHtmlString(aList(0, k))%>" target="_parent"><%=routineHtmlString(aList(9, k))%>&nbsp;</A></TD>
		<TD nowrap><A href="StaffDetail.asp?ContactID=<%=routineHtmlString(aList(0, k))%>" target="_parent"><%=routineHtmlString(aList(10, k))%>&nbsp;</A></TD>
		<TD nowrap><A href="StaffDetail.asp?ContactID=<%=routineHtmlString(aList(0, k))%>" target="_parent"><%=routineHtmlString(aList(11, k))%>&nbsp;</A></TD>
	</TR>
<%Next%>
</TBODY>
<TFOOT>
<TR>
<TH align="left" colspan="11">
	<INPUT type="hidden" name="txtPageNumber" value="<%=intPageNumber%>">
	<INPUT type="submit" name="action" value="&lt;&lt;">
	<INPUT type="submit" name="action" value="&lt;">
	<INPUT type="text" name="txtGoToPageNo" onClick="document.forms[0].txtGoToPageNo.value=''" title="You can jump to a specific page by typing the page number in this box" value="page <%=intPageNumber%> of <%=intPageCount%>" style="HEIGHT: 22px; WIDTH: 150px">
	<INPUT type="submit" name="action" value="&gt;">
	<INPUT type="submit" name="action" value="&gt;&gt;">
<!--	<INPUT type="submit" name="action" value="Export" title="Export this list to Excel"> -->
	<IMG src="images/excel.gif" onClick="document.forms[0].target='new';document.forms[0].hdnExport.value='xls';document.forms[0].submit();document.forms[0].hdnExport.value='';document.forms[0].target='_self';" width="32" height="32">
</TH>
</TR>
</TFOOT>
<CAPTION>Records <%=m+1%> to <%=n+1%> of <%=UBound(aList, 2) + 1 & " records"%></CAPTION>
</TABLE>
</FORM>
</BODY>
</HTML>

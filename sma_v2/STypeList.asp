<%@ Language=VBScript %>
<%Option Explicit%>
<%Response.Buffer = True%>
<!--#include file = "smaConstants.inc"-->
<!--#include file = "smaProcs.inc"-->
<!--#include file = "databaseconnect.asp" -->
<!--
*****************************************************************************************
* File Name:	STypeList.asp
*
* Author:		Gillies Archer
*
* Purpoase:		Display a list of Service Type record matching the search criteria entered
*				in the STypeCriteria.asp file.
*
* Modified:		Aug 2009 LC Dispaly Service Type ID, French Description of service type if
*				any, and if owned by Netcracker
******************************************************************************************
-->
<%
Dim objFSO, objTxtStream, strExportPath, liLength
Dim strRealUserID, aList, intPageNumber, intPageCount, aListEN
Dim strWinName, strLOBID, strSCategoryID, strSLevelID, strSTypeID, strSTypeDescription, strSStatus, strRevenue
Dim strMonth, strDay, strYear, strDate, chkActiveOnly, chkPrefLangOnly, strLANG
Dim objRs, strSQL, strWhereClause, strOrderBy, iCount
Dim objRsFrenchDesc, intSTLConstDisplayPageSize
dim temp

intSTLConstDisplayPageSize = 8
'TQ_INOSS
	strLANG = Request.Cookies("UserInformation")("language_preference")
	if (Len(strLANG) = 0) then strLANG = "EN"

	strWinName = Trim(Request.Form("hdnWinName"))
	strLOBID = Trim(Request.Form("selLOB"))
	strSCategoryID = Trim(Request.Form("selSCategory"))
	strSLevelID = Trim(Request.Form("selSLevel"))
	strSStatus = Trim(Request.Form("selSStatus"))
	strSTypeID = Trim(Request.Form("hdnSTypeID"))
	strSTypeDescription = UCase(Trim(Request.Form("txtSTypeDescription")))
	strMonth = Trim(Request.Form("selmonth"))
	strDay = Trim(Request.Form("selday"))
	strYear = Trim(Request.Form("selyear"))
	strDate = strMonth & "/" & strDay & "/" & strYear
	strRevenue = trim(Request.Form("selRevenue"))
	chkActiveOnly = Trim(Request.Form("chkActiveOnly"))
	chkPrefLangOnly = trim(Request.Form("chkPrefLangOnly"))

	strSQL ="SELECT " &_
				"A.LOB_CODE,  " &_
				"A.LOB_DESC,  " &_
				"B.SERVICE_CATEGORY_DESC,  " &_
				"C.SERVICE_TYPE_ID,  " &_
				"C.SERVICE_TYPE_DESC,  " &_
				"TO_CHAR(C.SERVICE_TYPE_START_DATE, 'MON-DD-YYYY'),  " &_
				"TO_CHAR(C.SERVICE_TYPE_END_DATE,  'MON-DD-YYYY'),  " &_
				"D.SERVICE_LEVEL_AGREEMENT_ID,  " &_
				"R.REVENUE_REGION_DESC, " &_
				"D.SERVICE_LEVEL_AGREEMENT_DESC,  " &_
				"C.SEND_TO_NC_LCODE, " &_
				"L.service_type_lang_desc " &_
			"FROM  " &_
				"CRP.LOB A, " &_
				"CRP.SERVICE_CATEGORY B,  " &_
				"CRP.V_SERVICE_TYPE C,  " &_
				"CRP.SERVICE_LEVEL_AGREEMENT D,  " &_
				"CRP.SERVICETYPE_REGION_XREF X,  " &_
				"SO.LCODE_REVENUE_REGION R, "    &_
				"crp.service_type_lang L "

	strWhereClause = " WHERE " &_
						"A.LOB_ID = B.LOB_ID AND " &_
						"B.SERVICE_CATEGORY_ID = C.SERVICE_CATEGORY_ID AND " &_
						"C.SERVICE_TYPE_ID = X.SERVICE_TYPE_ID(+) AND " &_
						"X.REGION_LCODE = R.REVENUE_REGION_LCODE(+) AND " &_
						"X.SERVICE_LEVEL_AGREEMENT_ID = D.SERVICE_LEVEL_AGREEMENT_ID(+) AND " &_
						"c.service_type_id =l.service_type_id(+) "

	If Len(strLOBID) <> 0 Then
		strWhereClause = strWhereClause & " AND A.LOB_ID = " & strLOBID
	End If

	If Len(strSCategoryID) <> 0 Then
		strWhereClause = strWhereClause & " AND B.SERVICE_CATEGORY_ID = " & strSCategoryID
	End If

	If Len(strSLevelID) <> 0 Then
		strWhereClause = strWhereClause & " AND X.SERVICE_LEVEL_AGREEMENT_ID = " & strSLevelID
	End If

	If Len(strSTypeID) <> 0 Then
		strWhereClause = strWhereClause & " AND C.SERVICE_TYPE_ID = " & strSTypeID
	End If

	If Len(strSTypeDescription) <> 0 Then
		strWhereClause = strWhereClause & " AND UPPER(C.SERVICE_TYPE_DESC) LIKE '" & Replace(strSTypeDescription, "'", "''") & "%'"
	End If

	If IsDate(StrDate) Then
	    strWhereClause = strWhereClause & " AND (TO_DATE('" & StrDate & "', 'MM/DD/YYYY') BETWEEN C.SERVICE_TYPE_START_DATE AND NVL(C.SERVICE_TYPE_END_DATE, TO_DATE('01/01/3000', 'MM/DD/YYYY')))"
	End If

	If Len(strRevenue) <> 0 Then
		strWhereClause = strWhereClause & " AND  ( R.REVENUE_REGION_LCODE = " & strRevenue & " )"
	End If

	If Len(chkActiveOnly) <> 0 Then
		strWhereClause = strWhereClause & " AND C.RECORD_STATUS_IND = 'A'"
	End If

	If Len(chkPrefLangOnly) <> 0 Then
		strWhereClause = strWhereClause & " AND C.language_preference_lcode like '" & strLANG & "' "
	End If

	strOrderBy = " ORDER BY A.LOB_DESC ASC, B.SERVICE_CATEGORY_DESC ASC, C.SERVICE_TYPE_DESC ASC, R.REVENUE_REGION_DESC ASC"
	strSQL = strSQL & strWhereClause & strOrderBy

	'Response.Write (strSQL)
	'Response.End

	Set objRS = objConn.Execute(strSQL)
	If Not objRS.EOF Then
		aList = objRS.GetRows
		'release and kill the recordset and the connection objects
'		objRS.Close
'		Set objRS = Nothing
'		objConn.Close
'		Set objConn = Nothing
	Else
		Response.Write "0 records found"
		Response.End
	End If

   'calculate page number
	intPageCount = Int(UBound(aList, 2) / intSTLConstDisplayPageSize) + 1
	Select Case Request("Action")
		Case "<<"	intPageNumber = 1
		Case "<"	intPageNumber = Request("txtPageNumber") - 1
					If intPageNumber < 1 Then intPageNumber = 1
		Case ">"	intPageNumber = Request("txtPageNumber") + 1
					If intPageNumber > intPageCount Then intPageNumber = intPageCount
		Case ">>"	intPageNumber = intPageCount
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
					Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
					'create export file (overwrite if exists)
					Set objTxtStream = objFSO.CreateTextFile(strExportPath & strRealUserID & "-SType.xls", True, False)

					With objTxtStream
						.WriteLine "<TABLE border=1>"
						.WriteLine "<THEAD>"
						.WriteLine "<TH>LOB</TH>"
						.WriteLine "<TH>Line of Business</TH>"
						.WriteLine "<TH>Service Category</TH>"
						.WriteLine "<TH>Service Type</TH>"
						.WriteLine "<TH>Service Type (French)</TH>"
						.WriteLine "<TH>Start Date</TH>"
						.WriteLine "<TH>End Date</TH>"
						.WriteLine "<TH>Revenue Region</TH>"
						.WriteLine "<TH>Service Level Agreement</TH>"
						.WriteLine "</THEAD>"

						'export the body
						For k = 0 To UBound(aList, 2)
							.WriteLine "<TR>"
							.WriteLine "<TD NOWRAP>" & routineHtmlString(aList(0, k)) & "&nbsp;</TD>"
							.WriteLine "<TD NOWRAP>" & routineHtmlString(aList(1, k)) & "&nbsp;</TD>"
							.WriteLine "<TD NOWRAP>" & routineHtmlString(aList(2, k)) & "&nbsp;</TD>"
							.WriteLine "<TD NOWRAP>" & routineHtmlString(aList(4, k)) & "&nbsp;</TD>"
							.WriteLine "<TD NOWRAP>" & routineHtmlString(aList(11, k)) & "&nbsp;</TD>"

							.WriteLine "<TD NOWRAP>" & routineHtmlString(aList(5, k)) & "&nbsp;</TD>"
							.WriteLine "<TD NOWRAP>" & routineHtmlString(aList(6, k)) & "&nbsp;</TD>"
							.WriteLine "<TD NOWRAP>" & routineHtmlString(aList(8, k)) & "&nbsp;</TD>"
							.WriteLine "<TD NOWRAP>" & routineHtmlString(aList(9, k)) & "&nbsp;</TD>"
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
						strsql = "<script type=""text/javascript"">document.location=""export/"&strRealUserID&"-SType.xls"";</script>"
						Response.Write strsql
						Response.End
'					Response.Redirect "export/" & strRealUserID & "-SType.xls"

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
	m = (intPageNumber - 1 ) * intSTLConstDisplayPageSize
	n = (intPageNumber * intSTLConstDisplayPageSize) - 1
	If n > UBound(aList, 2) Then n = UBound(aList, 2)

	'check If the client is still connected just before sending any HTML To the browser
	If Not Response.IsClientConnected Then Response.End

	'catch any unexpected error
	If err Then	DisplayError "BACK", "", err.Number, "Unexpected error", err.Description
%>
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

function body_onLoad() {
	DeleteCookie("ServiceType");
	DeleteCookie("STypeDesc");
}

function go_back(strSTypeID, strSTypeDescription, strServiceLevelID, strServiceLevelDesc, strSTypeEN){
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

	DeleteCookie("ServiceType");
	DeleteCookie("STypeDesc");
	DeleteCookie("WinName");

with (parent.opener.document.forms[0]) {
		hdnServiceTypeID.value = strSTypeID;
		txtServiceType.value = strSTypeDescription;
		hdnSLAID.value = strServiceLevelID;
		txtServiceLevelAgreement.value = strServiceLevelDesc;
		hdnSTypeEN.value = strSTypeEN;
	}
	parent.window.close();
}
//-->
</SCRIPT>
</HEAD>
<BODY language="javascript" onLoad="body_onLoad();">
<FORM id="frmSTypeList" name="frmSTypeList" method="post" action="STypeList.asp">
	<INPUT  id="hdnWinName" name="hdnWinName" value="<%=strWinName%>" type=hidden>
	<INPUT type="hidden" id="selLOB" name="selLOB" value="<%=strLOBID%>">
	<INPUT type="hidden" id="selSCategory" name="selSCategory" value="<%=strSCategoryID%>">
	<INPUT type="hidden" id="selSLevel" name="selSLevel" value="<%=strSLevelID%>">
	<INPUT type="hidden" id="selSStatus" name="selSStatus" value="<%=strSStatus%>">
	<INPUT type="hidden" id="selRevenue" name="selRevenue" value="<%=strRevenue%>">
	<INPUT type="hidden" id="txtSTypeDescription" name="txtSTypeDescription" value="<%=strSTypeDescription%>">
	<INPUT type="hidden" id="selmonth" name="selmonth" value="<%=strMonth%>">
	<INPUT type="hidden" id="selday" name="selday" value="<%=strDay%>">
	<INPUT type="hidden" id="selyear" name="selyear" value="<%=strYear%>">
    <INPUT type="hidden" id="chkActiveOnly" name="chkActiveOnly" value="<%=chkActiveOnly%>">
	<INPUT  id="chkPrefLangOnly" type=hidden name="chkPrefLangOnly" value="<%=chkPrefLangOnly%>">
	<INPUT type="hidden" id="hdnExport" name="hdnExport" value="">
<TABLE border="1" cellpadding="2" cellspacing="0" cols="8" width="100%">
<THEAD>
	<TH align="left" nowrap>STID</TH>
	<TH align="left" nowrap>Service Type</TH>
	<TH align="left" nowrap>French Description</TH>
	<TH align="left" nowrap>Start Date</TH>
	<TH align="left" nowrap>End Date</TH>
	<TH align="left" nowrap>Revenue Region</TH>
	<TH align="left" nowrap>Service Level Agreement</TH>
	<TH align="left" nowrap>LOB</TH>
	<TH align="left" nowrap>Service Category</TH>
	<TH align="left" nowrap>Owned By  </TH>

</THEAD>

<%
For k = m To n
	If Int(k/2) = k/2 Then
		Response.Write "<TR>"
	Else
		Response.Write "<TR bgcolor='white'>"
	End If

	dim rsSTypeEN, txtSTypeEN

	strSQL = "select service_type_desc from crp.service_type where service_type_id = '" & aList(3,k) & "'"

'	response.write strSQL

	Set objRS = objConn.Execute(strSQL)
	If Not objRS.EOF Then
		txtSTypeEN = objRS("service_type_desc")
		'release and kill the recordset and the connection objects
'		objRS.Close
'		Set objRS = Nothing
'		objConn.Close
'		Set objConn = Nothing
	End If

'	response.write txtSTypeEN
'	response.end

	If UCase(strWinName) <> UCase("Popup") Then%>
		<TD align="left" nowrap><A href="STypeDetail.asp?ServiceTypeID=<%=routineHtmlString(aList(3, k))%>&ncflag=<%=routineHtmlString(aList(10, k))%>"
								   target="_parent"><%=routineHtmlString(aList(3, k))%>&nbsp;</A></TD>
		<TD align="left" nowrap><A href="STypeDetail.asp?ServiceTypeID=<%=routineHtmlString(aList(3, k))%>&ncflag=<%=routineHtmlString(aList(10, k))%>"
								   target="_parent"><%=routineHtmlString(aList(4, k))%> </a></TD>
		<TD align="left" nowrap><A href="STypeDetail.asp?ServiceTypeID=<%=routineHtmlString(aList(3, k))%>&ncflag=<%=routineHtmlString(aList(10, k))%>"
								   target="_parent"><%=routineHtmlString(aList(11, k))%>
								   <% 'strSQL="select service_type_lang_desc from crp.service_type_lang " &_
								   		'  "where service_type_id = " & aList(3, k)
								  ' set objRsFrenchDesc = objConn.Execute(strSQL)
								   'If Not objRsFrenchDesc.EOF then
								   'response.write "<br>"
								  ' response.write objRsFrenchDesc(0)
								  ' end if
								  ' objRsFrenchDesc.close
								   'set objRsFrenchDesc = Nothing %>&nbsp;</A></TD>
		<TD align="left" nowrap><A href="STypeDetail.asp?ServiceTypeID=<%=routineHtmlString(aList(3, k))%>&ncflag=<%=routineHtmlString(aList(10, k))%>" target="_parent"><%=routineHtmlString(aList(5, k))%>&nbsp;</A></TD>
		<TD align="left" nowrap><A href="STypeDetail.asp?ServiceTypeID=<%=routineHtmlString(aList(3, k))%>&ncflag=<%=routineHtmlString(aList(10, k))%>"
								   target="_parent"><%=routineHtmlString(aList(6, k))%>&nbsp;</A></TD>
		<TD align="left" nowrap><A href="STypeDetail.asp?ServiceTypeID=<%=routineHtmlString(aList(3, k))%>&ncflag=<%=routineHtmlString(aList(10, k))%>"
								   target="_parent"><%=routineHtmlString(aList(8, k))%>&nbsp;</A></TD>
		<TD align="left" nowrap><A href="STypeDetail.asp?ServiceTypeID=<%=routineHtmlString(aList(3, k))%>&ncflag=<%=routineHtmlString(aList(10, k))%>"
								   target="_parent"><%=routineHtmlString(aList(9, k))%>&nbsp;</A></TD>
		<TD align="left" nowrap><A href="STypeDetail.asp?ServiceTypeID=<%=routineHtmlString(aList(3, k))%>&ncflag=<%=routineHtmlString(aList(10, k))%>"
								   target="_parent"><%=routineHtmlString(aList(0, k))%>&nbsp;</A></TD>
		<TD align="left" nowrap><A href="STypeDetail.asp?ServiceTypeID=<%=routineHtmlString(aList(3, k))%>&ncflag=<%=routineHtmlString(aList(10, k))%>"
								   target="_parent"><%=routineHtmlString(aList(2, k))%>&nbsp;</A></TD>
		<TD align="left" nowrap><A href="STypeDetail.asp?ServiceTypeID=<%=routineHtmlString(aList(3, k))%>&ncflag=<%=routineHtmlString(aList(10, k))%>"
								   target="_parent"><% if Clng(aList(10,k)) >=1 then response.write "NetCracker" else response.write "" end if%>&nbsp;</A></TD></TR>

	<%Else%>
		<TD align="left" nowrap height="23"><A href="#" onClick="return go_back('<%=aList(3, k)%>', '<%=routineHtmlString(aList(4, k))%>', '<%=routineHtmlString(aList(7, k))%>', '<%=routineHtmlString(aList(9, k))%>', '<%=routineHtmlString(txtSTypeEN)%>');"><%=routineHtmlString(aList(3, k))%>&nbsp;</A></TD>
		<TD align="left" nowrap height="23"><A href="#" onClick="return go_back('<%=aList(3, k)%>', '<%=routineHtmlString(aList(4, k))%>', '<%=routineHtmlString(aList(7, k))%>', '<%=routineHtmlString(aList(9, k))%>', '<%=routineHtmlString(txtSTypeEN)%>');">
		<%=routineHtmlString(aList(4, k))%>
		<% strSQL="select service_type_lang_desc from crp.service_type_lang " &_
								   		  "where service_type_id = " & aList(3, k)
								   set objRsFrenchDesc = objConn.Execute(strSQL)
								   If Not objRsFrenchDesc.EOF then
								   response.write "<br>"
								   response.write objRsFrenchDesc(0)
								   end if
								   objRsFrenchDesc.close
								   set objRsFrenchDesc = Nothing %>&nbsp;</A></TD>
		<TD align="left" nowrap height="23"><A href="#" onClick="return go_back('<%=aList(3, k)%>', '<%=routineHtmlString(aList(4, k))%>', '<%=routineHtmlString(aList(7, k))%>', '<%=routineHtmlString(aList(9, k))%>', '<%=routineHtmlString(txtSTypeEN)%>');"><%=routineHtmlString(aList(5, k))%>&nbsp;</A></TD>
		<TD align="left" nowrap height="23"><A href="#" onClick="return go_back('<%=aList(3, k)%>', '<%=routineHtmlString(aList(4, k))%>', '<%=routineHtmlString(aList(7, k))%>', '<%=routineHtmlString(aList(9, k))%>', '<%=routineHtmlString(txtSTypeEN)%>');"><%=routineHtmlString(aList(6, k))%>&nbsp;</A></TD>
		<TD align="left" nowrap height="23"><A href="#" onClick="return go_back('<%=aList(3, k)%>', '<%=routineHtmlString(aList(4, k))%>', '<%=routineHtmlString(aList(7, k))%>', '<%=routineHtmlString(aList(9, k))%>', '<%=routineHtmlString(txtSTypeEN)%>');"><%=routineHtmlString(aList(8, k))%>&nbsp;</A></TD>
		<TD align="left" nowrap height="23"><A href="#" onClick="return go_back('<%=aList(3, k)%>', '<%=routineHtmlString(aList(4, k))%>', '<%=routineHtmlString(aList(7, k))%>', '<%=routineHtmlString(aList(9, k))%>', '<%=routineHtmlString(txtSTypeEN)%>');"><%=routineHtmlString(aList(9, k))%>&nbsp;</A></TD>
		<TD align="left" nowrap height="23"><A href="#" onClick="return go_back('<%=aList(3, k)%>', '<%=routineHtmlString(aList(4, k))%>', '<%=routineHtmlString(aList(7, k))%>', '<%=routineHtmlString(aList(9, k))%>', '<%=routineHtmlString(txtSTypeEN)%>');"><%=routineHtmlString(aList(0, k))%>&nbsp;</A></TD>
		<TD align="left" nowrap height="23"><A href="#" onClick="return go_back('<%=aList(3, k)%>', '<%=routineHtmlString(aList(4, k))%>', '<%=routineHtmlString(aList(7, k))%>', '<%=routineHtmlString(aList(9, k))%>', '<%=routineHtmlString(txtSTypeEN)%>');"><%=routineHtmlString(aList(2, k))%>&nbsp;</A></TD>
		<TD align="left" nowrap height="23"><A href="#" onClick="return go_back('<%=aList(3, k)%>', '<%=routineHtmlString(aList(4, k))%>', '<%=routineHtmlString(aList(7, k))%>', '<%=routineHtmlString(aList(9, k))%>', '<%=routineHtmlString(txtSTypeEN)%>');"><%response.write "NetCracker"%>&nbsp;</A></TD></tr>
	<%End If
Next
'release and kill the recordset and the connection objects
	objRS.Close
	Set objRS = Nothing
	objConn.Close
	Set objConn = Nothing
%>
</TBODY>
<TFOOT>
<TR>
<TD align="left" colspan="10">
	<INPUT type="hidden" name="txtPageNumber" value="<%=intPageNumber%>">
	<INPUT type="submit" name="action" value="&lt;&lt;">
	<INPUT type="submit" name="action" value="&lt;">
	<INPUT type="text" name="txtGoToPageNo" onClick="document.forms[0].txtGoToPageNo.value=''" title="You can jump To a specific page by typing the page number in this box" value="page <%=intPageNumber%> of <%=intPageCount%>" style="HEIGHT: 22px; WIDTH: 150px">
	<INPUT type="submit" name="action" value="&gt;">
	<INPUT type="submit" name="action" value="&gt;&gt;">
	<IMG src="images/excel.gif" onClick="document.forms[0].target='new';document.forms[0].hdnExport.value='xls';document.forms[0].submit();document.forms[0].hdnExport.value='';document.forms[0].target='_self';" width="32" height="32">
</TD>
</TR>
</TFOOT>
<CAPTION>Records <%=m+1%> To <%=n+1%> of <%=UBound(aList, 2) + 1 & " records"%></CAPTION>
</TABLE>
</FORM>
</BODY>
</HTML>

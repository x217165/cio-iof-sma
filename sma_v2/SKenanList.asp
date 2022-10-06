<%@ Language=VBScript %>
<%Option Explicit%>
<%Response.Buffer = True%>
<!--#include file="smaConstants.inc"-->
<!--#include file="databaseconnect.asp"-->
<!-- #include file="kenanconnect.asp" -->
<!--#include file="smaProcs.inc"-->
<!--
*****************************************************************************************
* File Name:	STKenanList.asp
* Author:		Linda chen
* Purpoase:		Display a list of Service Type Kenan Attribute records matching the search criteria entered
*				in the SKenanACriteria.asp file.
* Date:			August 2009
******************************************************************************************
-->
<%
Dim objFSO, objTxtStream, strExportPath, liLength
Dim strRealUserID, aList, intPageNumber, intPageCount, aListEN
Dim strWinName, strLOBID, strSCategoryID, strSTypeID, strSTypeDescription
Dim  chkActiveOnly, chkPrefLangOnly, strLANG
Dim objRs, strSQL, strWhereClause, strOrderBy, iCount
dim temp
dim strKenanCP, strKenanPK, objRsSTKenanCP, objRsSTKenanPK, strken, strkenCP, strkenCPID, strkenPK, strkenPKID




'TQ_INOSS
	'strLANG = Request.Cookies("UserInformation")("language_preference")
	'if (Len(strLANG) = 0) then strLANG = "EN"
	strLANG = "EN"

	strWinName = Trim(Request.Form("hdnWinName"))
	strLOBID = Trim(Request.Form("selLOB"))

	strSCategoryID = Trim(Request.Form("selSCategory"))
	strSTypeID = Trim(Request.Form("hdnSTypeID"))
	strSTypeDescription = UCase(Trim(Request.Form("txtSTypeDescription")))

	strKenanCP = Trim(Request.Form("selSTKenanCP"))
	strKenanPK = UCase(Trim(Request.Form("selSTKenanPK")))

	chkActiveOnly = Trim(Request.Form("chkActiveOnly"))
	' strSTTID = Request.Form("selSTTID")

	'response.write("LOB is " & strLOBID)
	'response.write("CategoryID is " & strSCategoryID)
	'response.write("STypeDescription is " & strSTypeDescription)
	'response.write("strKenanCP is " & strKenanCP)
	'response.write("strKenanPK is " & strKenanPK)
	'response.end

	

	strSQL ="SELECT " &_
				"A.LOB_CODE,  " &_
				"A.LOB_DESC,  " &_
				"B.SERVICE_CATEGORY_DESC,  " &_
				"C.SERVICE_TYPE_ID,  " &_
				"C.SERVICE_TYPE_DESC,  " &_
				"x.component_id, " &_
				"x.package_id " &_
			"FROM  " &_
				"CRP.LOB A, " &_
				"CRP.SERVICE_CATEGORY B,  " &_
				"CRP.V_SERVICE_TYPE C, " &_
				"CRP.service_type_Kenan_xref X  "


	strWhereClause = " WHERE " &_
						"A.LOB_ID = B.LOB_ID AND " &_
						"B.SERVICE_CATEGORY_ID = C.SERVICE_CATEGORY_ID AND " &_
						"C.SERVICE_TYPE_ID = X.SERVICE_TYPE_ID(+) and " &_
						"X.RECORD_STATUS_IND  = 'A' "&_
						"AND C.language_preference_lcode like '" & strLANG & "' "

	If Len(strLOBID) <> 0 Then
		strWhereClause = strWhereClause & " AND A.LOB_ID = " & strLOBID
	End If

	If Len(strSCategoryID) <> 0 Then
		strWhereClause = strWhereClause & " AND B.SERVICE_CATEGORY_ID = " & strSCategoryID
	End If

'	If Len(strSTTID) <> 0 Then
'		strWhereClause = strWhereClause & " AND au.srvc_type_att_id = " & strSTTID
'	End If

	If Len(strSTypeID) <> 0 Then
		strWhereClause = strWhereClause & " AND C.SERVICE_TYPE_ID = " & strSTypeID
	End If

	If Len(strSTypeDescription) <> 0 Then
		strWhereClause = strWhereClause & " AND UPPER(C.SERVICE_TYPE_DESC) LIKE '" & Replace(strSTypeDescription, "'", "''") & "%'"
	End If

	if Len(strKenanCP) <> 0 Then
		strWhereClause = strWhereClause & " AND X.COMPONENT_ID = " & strkenanCP
	End If

	if Len(strKenanPK) <> 0 Then
		strWhereClause = strWhereClause & " AND X.PACKAGE_ID = " & strkenanPK
	End If

	If Len(chkActiveOnly) <> 0 Then
		strWhereClause = strWhereClause & " AND C.RECORD_STATUS_IND = 'A'"
	End If


	strOrderBy = " ORDER BY A.LOB_DESC ASC, B.SERVICE_CATEGORY_DESC ASC, C.SERVICE_TYPE_DESC ASC, " &_
				"x.component_id, x.package_id ASC"
	strSQL = strSQL & strWhereClause & strOrderBy

'response.write(strSQL)
'response.end

	Set objRS = objConn.Execute(strSQL)
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
	'Print the first row for debug
	'response.write(aList(4,0)& "|" & aList(5,0) &"|" & aList(0,0) &"|"& aList(2,0) &"|")
	'response.end


' strSQL = "SELECT PACKAGE_ID, PACKAGE_NAME FROM ARBOR.V_PKG_COMPONENTS " &_
'    		" order by PACKAGE_ID, PACKAGE_NAME"
' set objRsSTKenanPK = objKenanConn.Execute(strSQL)



   'calculate page number
	intPageCount = Int(UBound(aList, 2) / intConstDisplayPageSize) + 1
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
					Set objTxtStream = objFSO.CreateTextFile(strExportPath & strRealUserID & "-STypeKenan.xls", True, False)

					With objTxtStream
						.WriteLine "<TABLE border=1>"
						.WriteLine "<THEAD>"
						.WriteLine "<TH>Service Type</TH>"
						.WriteLine "<TH>Kenan Package ID</TH>"
						.WriteLine "<TH>Kenan Package Name</TH>"
						.WriteLine "<TH>Kenan Component ID</TH>"
						.WriteLine "<TH>Kenan Component Name</TH>"
						.WriteLine "<TH>LOB</TH>"
						.WriteLine "<TH>Service Category</TH>"
						.WriteLine "</THEAD>"

						'export the body
						For k = 0 To UBound(aList, 2)
							.WriteLine "<TR>"
							.WriteLine "<TD NOWRAP>" & routineHtmlString(aList(4, k)) & "&nbsp;</TD>"
							if aList(6,k) <>"" then
								strSQL = "SELECT PACKAGE_id, PACKAGE_NAME " &_
									"FROM ARBOR.V_PKG_COMPONENTS " &_
   									"where package_id= " & Clng(aList(6,k))
									set objRsSTKenanPK = objKenanConn.Execute(strSQL)
									strkenPKID = objRsSTKenanPK(0)
									strkenPK = objRsSTKenanPK(1)
 							'		strkenPK = objRsSTKenanPK(0)& " | " & objRsSTKenanPK(1)
		    							objRsSTKenanPK.close
	  						 		set objRsSTKenanPK = Nothing
							else
								strkenPKID = ""
								strkenPK = ""
		    					end if

							strSQL = "SELECT component_id, COMPONENT_NAME " &_
								"FROM ARBOR.V_PKG_COMPONENTS " &_
    								"where component_id= " & Clng(aList(5,k))
 								set objRsSTKenanCP = objKenanConn.Execute(strSQL)
								strkenCPID = objRsSTKenanCP(0)
								strkenCP = objRsSTKenanCP(1)
 							'	strkenCP = objRsSTKenanCP(0)& " | " & objRsSTKenanCP(1)
		    						objRsSTKenanCP.close
	  						 	set objRsSTKenanCP = Nothing

							.WriteLine "<TD NOWRAP>" & routineHtmlString(strkenPKID) & "&nbsp;</TD>"
							.WriteLine "<TD NOWRAP>" & routineHtmlString(strkenPK) & "&nbsp;</TD>"
							.WriteLine "<TD NOWRAP>" & routineHtmlString(strkenCPID) & "&nbsp;</TD>"
							.WriteLine "<TD NOWRAP>" & routineHtmlString(strkenCP) & "&nbsp;</TD>"
							.WriteLine "<TD NOWRAP>" & routineHtmlString(aList(0, k)) & "&nbsp;</TD>"
							.WriteLine "<TD NOWRAP>" & routineHtmlString(aList(2, k)) & "&nbsp;</TD>"
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
						strsql = "<script type=""text/javascript"">document.location=""export/"&strRealUserID&"-STypeKenan.xls"";</script>"
						Response.Write strsql
						Response.End
'					Response.Redirect "export/" & strRealUserID & "-STypeKenan.xls"

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
<FORM id="frmSTKenanList" name="frmSTKenanList" method="post" action="SKenanList.asp">
	<INPUT type="hidden" id="hdnWinName" name="hdnWinName" value="<%=strWinName%>">
	<INPUT type="hidden" id="selLOB" name="selLOB" value="<%=strLOBID%>">
	<INPUT type="hidden" id="selSCategory" name="selSCategory" value="<%=strSCategoryID%>">
	<INPUT type="hidden" id="selKenanPK" name="selSTKenanPK" value="<%=strKenanPK%>">
	<INPUT type="hidden" id="selKenanCP" name="selSTKenanCP" value="<%=strKenanCP%>">
	<INPUT type="hidden" id="txtSTypeDescription" name="txtSTypeDescription" value="<%=strSTypeDescription%>">
    <INPUT type="hidden" id="chkActiveOnly" name="chkActiveOnly" value="<%=chkActiveOnly%>">
	<INPUT type="hidden" id="chkPrefLangOnly" name="chkPrefLangOnly" value="<%=chkPrefLangOnly%>">
	<INPUT type="hidden" id="hdnExport" name="hdnExport" value="">
<TABLE border="1" cellpadding="2" cellspacing="0"  width="100%">
<THEAD>
	<TH align="left" nowrap>Service Type</TH>
	<TH align="left" nowrap>Kenan Package ID and Name</TH>
	<TH align="left" nowrap>Kenan Component ID and Name</TH>
	<TH align="left" nowrap>LOB</TH>
	<TH align="left" nowrap>Service Category</TH>
</THEAD>

<%
For k = m To n
	If Int(k/2) = k/2 Then
		Response.Write "<TR>"
	Else
		Response.Write "<TR bgcolor='white'>"
	End If

'	dim rsSTypeEN, txtSTypeEN

'	strSQL = "select service_type_desc from crp.service_type where service_type_id = '" & aList(3,k) & "'"

'	response.write strSQL

'	Set objRS = objConn.Execute(strSQL)
'	If Not objRS.EOF Then
'		txtSTypeEN = objRS("service_type_desc")
		'release and kill the recordset and the connection objects
'		objRS.Close
'		Set objRS = Nothing
'		objConn.Close
'		Set objConn = Nothing
'	End If

'	response.write txtSTypeEN
'	response.end

	If UCase(strWinName) <> UCase("Popup") Then%>
		<TD align="left" nowrap>
		<%'  <A href="STypeKenanDetail.asp?ServiceTypeID=<%=routineHtmlString(aList(3, k))%>
		<%' "  target="_parent">  %>
		<%=routineHtmlString(aList(4, k))%></A></TD>
		<TD align="left" >
		<%  if aList(6,k) <>"" then
			  strSQL = "SELECT PACKAGE_id, PACKAGE_NAME " &_
			  "FROM ARBOR.V_PKG_COMPONENTS " &_
    		  " where PACKAGE_id=" & Clng(aList(6,k))
 			  set objRsSTKenanPK = objKenanConn.Execute(strSQL)
 			  'response.write strSQL
			  if not objRsSTKenanPK.Eof then
 			   response.write objRsSTKenanPK(0)& " | " & objRsSTKenanPK(1)
			  end if
 			  response.write strken
		      objRsSTKenanPK.close
		      set objRsSTKenanPK = Nothing
		    else
		      response.write "NULL"
		    end if
		  %>
		</A></TD>
		<TD align="left" >
		<% strSQL = "SELECT component_id, COMPONENT_NAME " &_
			"FROM ARBOR.V_PKG_COMPONENTS " &_
    		" where component_id=" & Clng(aList(5,k))
 			set objRsSTKenanCP = objKenanConn.Execute(strSQL)
			 if not objRsSTKenanCP.Eof then
				 response.write objRsSTKenanCP(0)& " | " & objRsSTKenanCP(1)
    		end if
 			'response.write strken
		    objRsSTKenanCP.close
		    set objRsSTKenanCP = Nothing
		    %>
</A></TD>
		<TD align="left" nowrap><%=routineHtmlString(aList(0, k))%></A></TD>
		<TD align="left" nowrap><%=routineHtmlString(aList(2, k))%></A></TD>
	<%End If
Next
'release and kill the recordset and the connection objects
	'objRS.Close
	'objRsSTKenanPK.close
	'Set objRS = Nothing
	'Set objRsSTKenanPK = Nothing
	'objConn.Close
	objKenanConn.Close
	Set objConn = Nothing
	Set objKenanConn = Nothing
%>
</TBODY>
<TFOOT>
<TR>
<TD align="left" colspan="8">
	<INPUT type="hidden" name="txtPageNumber" value="<%=intPageNumber%>">
	<INPUT type="submit" name="action" value="&lt;&lt;">
	<INPUT type="submit" name="action" value="&lt;">
	<INPUT type="text" name="txtGoToPageNo" onClick="document.forms[0].txtGoToPageNo.value=''" title="You can jump To a specific page by typing the page number in this box" value="page <%=intPageNumber%> of <%=intPageCount%>" style="HEIGHT: 22px; WIDTH: 150px">
	<INPUT type="submit" name="action" value="&gt;">
	<INPUT type="submit" name="action" value="&gt;&gt;">
	<IMG src="images/excel.gif" onClick="target='new';
	hdnExport.value='xls';
	submit();
	hdnExport.value=''
	target='_self';" width="32" height="32">
</TD>
</TR>
</TFOOT>
<CAPTION>Records <%=m+1%> To <%=n+1%> of <%=UBound(aList, 2) + 1 & " records"%></CAPTION>
</TABLE>
</FORM>
</BODY>
</HTML>

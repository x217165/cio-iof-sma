<%@ Language=VBScript %>
<%
option explicit
'on error resume next
'********************************************************************************************
'* Page name:	CorrReport.asp
'* Purpose:		Displays the correlation items hierarchically (i.e. a root service will be
'*				followed by an indented list of all the elements that make up that root service)
'* Created by:	Nancy Mooney 11/06/2000
'********************************************************************************************
%>
<!--#include file="smaConstants.inc"-->
<!--#include file="smaProcs.inc"-->
<!--#include file="databaseconnect.asp" -->

<html>
<head>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<meta HTTP-EQUIV="Pragma" CONTENT="no-cache">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Correlated Elements</title>
<link href="stylesheets/styles.css" rel="stylesheet" type="text/css">
</head>
<%

'check users access rights
dim intAccessLevel
intAccessLevel = CInt(CheckLogon(strConst_CorrelationCustomer))
if intAccessLevel and intConst_Access_ReadOnly <> intConst_Access_ReadOnly then
	DisplayError "BACK", "", 0, "ACCESS DENIED", "You don't have access to correlation management. Please contact your system administrator"
end if

dim lngCSID, strCSName, sql, aList, intPageNumber, intPageCount

'get customer service id
lngCSID = Request("CSID")
strCSName = Request("CSName")

sql =	"SELECT " & _
			"t2.lev " & _
			", decode(t1.network_element_id " &_
				",null,decode(t1.circuit_id " & _
					" ,null,decode(t1.root_customer_service_id,null,'z' " & _
							",'ROOT = ' || cs.customer_service_desc) " & _
			",c.circuit_type_code || ' = ' || c.circuit_name) " & _
            ",ne.network_element_type_code || ' = ' || ne.network_element_name) description " & _
		"FROM " & _
			"crp.managed_correlation t1, " &_
			"(SELECT  rownum ppp " & _
				",level lev " & _
				",root_customer_service_id " & _
				",customer_service_id " & _
				", managed_correlation_id " & _
			"FROM crp.managed_correlation " & _
			"START WITH customer_service_id = " & lngCSID & " " & _
			"CONNECT BY PRIOR root_customer_service_id = customer_service_id " & _
			") t2, " & _
			"crp.network_element ne, " & _
			"crp.circuit c, " & _
			"crp.customer_service cs " & _
		"WHERE t1.managed_correlation_id =t2.managed_correlation_id " & _
			"and t1.network_element_id = ne.network_element_id(+) " & _
			"and t1.root_customer_service_id = cs.customer_service_id(+) " & _
			"and t1.circuit_id = c.circuit_id(+) " & _
		"ORDER BY t2.ppp"

'Response.Write (sql & "<BR>")

dim rsCorr
set rsCorr = server.CreateObject("ADODB.Recordset")
rsCorr.Open sql, objConn
if err then
	DisplayError "BACK", "", err.Number, "Cannot create rsCorr recordset.", err.Description
end if

if not rsCorr.EOF then
	aList = rsCorr.GetRows
else
	Response.Write "0 records found. There are no elements correlated with this customer service."
	Response.end
end if

'release and kill the recordset and the connection objects
rsCorr.Close
set rsCorr = nothing

objConn.close
set objConn = nothing

if Request("hdnExport") <> "" then
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
	set objTxtStream = objFSO.CreateTextFile(strExportPath&strRealUserID&"-corrReport.xls", true, false)
	if err then
		DisplayError "CLOSE", "", err.Number, "CorrList.asp - Cannot create Excel spreadsheet file due to the following reasons. Please contact your website administrator.", err.Description
	end if
	with objTxtStream
		.WriteLine "<table border=1>"
		'export the header
		.WriteLine "<TR>"
		.WriteLine "<TH>Correlated Elements for Customer Service " & lngCSID & " : " & strCSName & "</TH>"
		.WriteLine "</TR>"

		'export the body
		for k = 0 to UBound(aList, 2)
			if (cInt(aList(0,k))>1) then
				dim strSpace, cnt
				strSpace = ""
				for cnt=1 to (cInt(alist(0,k))-1)*10
					strSpace = strSpace & "&nbsp;"
				next
			end if
			.WriteLine "<TR>"
			.WriteLine "<TD NOWRAP>"& strSpace & routineHtmlString(aList(1,k))&"</TD>"
			.WriteLine "</TR>"
		next
		.WriteLine "<tfoot><tr><td><BR>Each indented block represents a subset of elements that make up the element listed directly above.</td></tr></tfoot>"
		.WriteLine "</table>"
	end with
	objTxtStream.Close
	'Response.end
	Response.redirect "export/"&strRealUserID&"-corrReport.xls"
end if

'find number of records
dim k, n
n = UBound(aList,2)

'check if the client is still connected just before sending any html to the browser
if response.isclientconnected = false then
	Response.End
end if

'catch any unexpected error
if err then
	DisplayError "BACK", "", err.Number, "Unexpected error", err.Description
end if

%>


<body>
<form name="frmCorrReport" action="CorrReport.asp" method="POST">
<input type="hidden" name="hdnExport" value>
<input type="hidden" name="CSID" value=<%=lngCSID%>>
<input type="hidden" name="CSName" value=<%=strCSName%>>

<table border="0" cellspacing="0" cellpadding="0" width="100%">
<thead>
	<tr>
		<td align=left>Correlated Elements for Customer Service <%=lngCSID%> : <%=strCSName%></td>
		<td><img SRC="images/excel.gif" onclick="document.frmCorrReport.target='new';document.frmCorrReport.hdnExport.value='xls';document.frmCorrReport.submit();document.frmCorrReport.hdnExport.value='';document.frmCorrReport.target='_self';" WIDTH="32" HEIGHT="32"></td>
	</tr>
</thead>
<tbody>
<%
'display the table

for k = 0 to n
	'Alternate row background colour
	'if Int(k/2) = k/2 then
	'	Response.write "<TR >"
	'else
	'	Response.write "<TR bgcolor=White >"
	'end if
	Response.write "<TR>"
	strSpace = ""
	if (cInt(aList(0,k))>1) then
		for cnt=1 to ((cInt(alist(0,k))-1)*10)
			'Response.Write ("count =" & cnt & "<BR>")
			strSpace = strSpace & "&nbsp;"
		next
	end if
	Response.Write "<TD NOWRAP colspan=2>"& strSpace & routineHtmlString(aList(1,k))&"&nbsp;</a></TD>"&vbCrLf
	Response.Write "</TR>"
next
%>
</tbody>
<tfoot>
<tr><td colspan=2 ><BR>Each indented block represents a subset of elements that make up the element listed directly above.</td></tr>
</tfoot>
</table>
</form>
</body>
</html>

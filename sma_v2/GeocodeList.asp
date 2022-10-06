<%@ Language=VBScript %>
<% option explicit %>
<!--#include file="smaConstants.inc"-->
<!--#include file="smaProcs.inc"-->
<!--#include file="databaseconnect.asp"-->
<!--
********************************************************************************************
* Page name:	GeocodeList.asp
* Purpose:		To display the results of an address search.
*				Search criteria are chosen via AddressCriteria.asp
*
* Created by:	Sara Sangha	08/01/2000
*
********************************************************************************************
        Date		Author(s)		Changes/enhancements made
        -----		------			------------------------------------------------------
       12-May-08	ACheung, LChen		NGSM CLLI impelementation

***************************************************************************************************
-->
 <%

 dim aList, intPageNumber, intPageCount
 dim strGeocllicodeid, strAddress, strCity, strPostal, strProvince, strCllicode, strDescription, bolActiveOnly, strWinName
 dim strSQL, strSelectClause,strFromClause, strWhereClause, strRecordStatus, strOrderBy
 dim geoclliid

 ' read submitted values and create an sql
	strWinName = trim(Request("hdnWinName"))
	strGeocllicodeid = UCase(trim(Request("txtGeocllicodeid")))
	if(strGEocllicodeid<>"") then
		geoclliid = Clng(strGeocllicodeid)
	end if

	strAddress = UCase(trim(Request("txtAddress")))
	strCity = UCase(trim(Request("txtCity")))
	strPostal = UCase(trim(Request("txtPostal")))
	strProvince	= trim(Request("selProvince"))
    strDescription = UCase(trim(Request("txtDescription")))
	strCllicode = UCase(trim(Request("txtGeoclli")))
	bolActiveOnly = trim(Request("chkActiveOnly"))


	strSQL = "select CLLI_CODE, GEOCODEID_LCODE,  DESCRIPTION, "&_
			"ADDRESS, CITY, PROVINCE, POSTAL_CODE " &_
			"FROM CRP.LCODE_GEOCODEID "

	if len(strGeocllicodeid) > 0 then
	   if (strWhereClause="") then
		  strWhereClause = "where GEOCODEID_LCODE ="  & geoclliid & " "
       else
	   	  strWhereClause = "and GEOCODEID_LCODE ="  & geoclliid & " "
	   end if
	end IF


	if len(strAddress) > 0 then
	   if (strWhereClause="") then
		  strWhereClause = "where UPPER(ADDRESS) like '" & routineOraString(strAddress) & "%'"
	   else
	      strWhereClause = strWhereClause & "and UPPER(ADDRESS) like '" & routineOraString(strAddress) & "%'"
       end if
	end if

	if len(strCity) > 0 then
		if (strWhereClause="") then
     		strWhereClause = "where UPPER(CITY) like '" & routineOraString(strCity) & "%'"
     	else
     		strWhereClause = strWhereClause & "and UPPER(CITY) like '" & routineOraString(strCity) & "%'"
		end if
	end if

	if len(strPostal) > 0 then
		if (strWhereClause="") then
			strWhereClause = "where UPPER(POSTAL_CODE) like '" & strPostal & "%'"
		else
			strWhereClause = strWhereClause & "and UPPER(POSTAL_CODE) like '" & strPostal & "%'"
		end if
	end if


	if len(strProvince) > 0 then
		if (strWhereClause="") then
     		strWhereClause = "where PROVINCE = '" & routineOraString(strProvince) & "'"
     	else
			strWhereClause = strWhereClause & "and (PROVINCE) = '" & routineOraString(strProvince) & "'"
		end if
	end if

	if len(strCllicode) > 0 then
		if (strWhereClause="") then
     		strWhereClause = "where UPPER(CLLI_CODE) like '" & strCllicode & "%'"
     	else
			strWhereClause = strWhereClause & "and UPPER(CLLI_CODE) like '" & strCllicode & "%'"
		end if
	end if

	if len(strDescription) > 0 then
	   if (strWhereClause="") then
		  strWhereClause = "where DESCRIPTION  LIKE  '" & strDescription & "%'"
        else
		  strWhereClause = strWhereClause & " and DESCRIPTION  LIKE   '" & strDescription & "%'"
 	   end if
	end IF

	If bolActiveOnly = "yes" then
		if (strWhereClause="") then
     		strWhereClause = "where record_status_ind  = 'A' "
     	else
     		strRecordStatus = " and record_status_ind  = 'A' "
     	end if
	Else 'no
		strRecordStatus = " "
	End If

	'strOrderBy = " order by c.customer_name, " &_
	'			 " decode(primary_address_flag,'Y',0,1)  + decode(billing_address_flag,'Y',0,1) +  decode(mailing_address_flag,'Y',0,1), "  &_
	'			 " a.province_state_lcode, a.municipality_name, a.street "

	strsql = strsql & strWhereClause & strRecordStatus

'Response.Write(strsql)
'Response.End


	Dim objRs,Recordcnt,strbgcolor

	set objRS = objConn.Execute(StrSql)
	if not objRS.EOF then
		aList = objRS.GetRows
	else
		Response.Write "0 records found"
		Response.end
	end if
	'response.write("aList(0,0) is" & aList(0,0))
	'response.end


   'release and kill the recordset and the connection objects
	objRS.Close
	set objRS = nothing

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
						set objTxtStream = objFSO.CreateTextFile(strExportPath&strRealUserID&"-Address.xls", true, false)

						with objTxtStream
							.WriteLine "<table border=1>"

							'export the header
							.WriteLine "<TR>"
'							.WriteLine "<TR bgcolor=#ffcc99>"
							.WriteLine "<TH>CLLI CODE</TD>"
							.WriteLine "<TH>GEOCODE ID</TD>"
							.WriteLine "<TH>DESCRIPTION</TD>"
							.WriteLine "<TH>ADDRESS</TD>"
							.WriteLine "<TH>CITY</TD>"
							.WriteLine "<TH>PROVINCE</TD>"
							.WriteLine "<TH>POSTAL CODE</TD>"
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

								.WriteLine "<TD NOWRAP>"&routineHtmlString(aList(1,k))&" &nbsp; </TD>"
								.WriteLine "<TD NOWRAP>"&routineHtmlString(aList(2,k))&" &nbsp; </TD>"
								.WriteLine "<TD NOWRAP>"&routineHtmlString(aList(3,k))&" &nbsp; </TD>"
								.WriteLine "<TD NOWRAP>"&routineHtmlString(aList(4,k))&" &nbsp;</TD>"
								.WriteLine "<TD NOWRAP>"&routineHtmlString(aList(5,k))&" &nbsp; </TD>"
								.WriteLine "<TD NOWRAP>"&routineHtmlString(aList(6,k))&" &nbsp; </TD>"
								.WriteLine "<TD NOWRAP>"&routineHtmlString(aList(7,k))&" &nbsp; </TD>"
								.WriteLine "<TD NOWRAP>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</TD>"
								.WriteLine "</TR>"
							next
							.WriteLine "</table>"
						end with
						objTxtStream.Close
						strsql = "<script type=""text/javascript"">document.location=""export/"&strRealUserID&"-Address.xls"";</script>"
						Response.Write strsql
						Response.End

						'Response.redirect "export/"&strRealUserID&"-Address.xls"

				elseif Request("txtGoToPageNo") <> "" then
						intPageNumber = CInt(Request("txtGoToPageNo"))
					else
						intPageNumber = 1
					end if
	end select

	if intPageNumber < 1 then
		intPageNumber = 1
	end if
	if intPageNumber > intPageCount then
		intPageNumber = intPageCount
	end if

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
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<META HTTP-EQUIV="Pragma" CONTENT="no-cache">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<LINK rel="stylesheet" type="text/css" href="stylesheets/styles.css">
<script type="text/javascript" SRC="GeneralJavaFunctions.js"></script>
<SCRIPT LANGUAGE=javascript>
<!--
//**********************************************Java Functions ***********************************
function go_back(strCllicode, strGeocodeid, strDescription,  strAddress, strCity,  strProvince, strPostal){
//************************************************************************************************
// Function:	go_back
//
// Purpose:		To write the values of selected row into the base window that called the lookup
//				function. In addition, this function closes the Popup window.
//
// Created By:	Sara Sangha Aug 29th, 2000
//
// Updated By:
//************************************************************************************************

var strFullClliInfo ;
var exception;
	strFullClliInfo = strCllicode + '\n' + strGeocodeid + '\n' + strDescription + '\n' + strAddress + '\n' + strCity + ',' +  strProvince + '\n' + strPostal

	//alert (strProvinceCode);
	parent.opener.document.forms[0].hdnGeocode.value = strGeocodeid ;
	parent.opener.document.forms[0].textGeocllicode.value = strFullClliInfo;
	DeleteCookie("strSimple");

	try
	{
		//the following fields were added because they needed to be updated in ServLocDetail.asp
		parent.opener.document.forms[0].hdnGeoCodeid.value = strGeocodeid ;
		parent.opener.document.forms[0].txtGeoClliCode.value = strFullClliInfo ;
	}
	catch(exception)
	{}
	DeleteCookie("WinName");
	DeleteCookie("strSimple");
	parent.window.close ();

	}
//-->
//*********************************************** End of Java Functions****************************
</SCRIPT>

</HEAD>
<BODY>

<FORM method=post name=frmGeocodeList action="GeocodeList.asp">

	<input type=hidden name=hdnWinName value="<%=strWinName%>">
    <input type=hidden name=txtGeoclli value="<%=strCllicode%>">
    <input type=hidden name=txtGeocllicodeid value="<%=strGeocllicodeid%>">
    <input type=hidden name=txtDescription value="<%=strDescription%>">
    <input type=hidden name=txtAddress value="<%=strAddress%>">
    <input type=hidden name=txtCity value="<%=strCity%>">
    <input type=hidden name=selProvince value="<%=strProvince%>">
    <input type=hidden name=txtPostal value="<%=strPostal%>">
    <input type=hidden name=chkActiveOnly value="<%=bolActiveOnly%>">
    <input type=hidden name="hdnExport" value>

<TABLE  border=1 cellPadding=2 cellSpacing=0 width="100%">
 <THEAD>
    <TR>
		<TH align=left width="9%">CLLI CODE</TH>
		<TH align=center width="6%">GEOCODE ID</TH>
		<TH align=center width="21%">DESCRIPTION</TH>
		<TH align=center width="29%">ADDRESS</TH>
		<TH align=left width="11%">CITY</TH>
	    <TH align=left width="7%">PROVINCE</TH>
		<TH align=left width="12%">POSTAL CODE</TH>
     </TR>
  </THEAD>
  <TBODY>
<%

dim strBilling, strPrimary, strMailing
'display the table
for k = m to n
	'Alternate row background colour
	if Int(k/2) = k/2 then
		Response.write "<TR>"
	else
		Response.write "<TR bgcolor=White>"
	end if

	if (strWinName = "Popup" or strWinName = "Simple") then
	Response.Write "<td><a href=""#"" onClick=""return go_back('"&aList(0,k)& "', '" &routineJavascriptString(aList(1,k))& "','" &routineJavascriptString(aList(2,k))& "','" &routineJavascriptString(aList(3,k))& "','" &routineJavascriptString(aList(4,k))& "','" &routineJavascriptString(aList(5,k))& "','" &routineJavascriptString(aList(6,k))& "')"">" &aList(0,k)& "</a></td>"&vbCrLf
	Response.Write "<td><a href=""#"" onClick=""return go_back('"&aList(0,k)& "', '" &routineJavascriptString(aList(1,k))& "','" &routineJavascriptString(aList(2,k))& "','" &routineJavascriptString(aList(3,k))& "','" &routineJavascriptString(aList(4,k))& "','" &routineJavascriptString(aList(5,k))& "','" &routineJavascriptString(aList(6,k))& "')"">" &aList(1,k)& "</a></td>"&vbCrLf
	Response.Write "<td><a href=""#"" onClick=""return go_back('"&aList(0,k)& "', '" &routineJavascriptString(aList(1,k))& "','" &routineJavascriptString(aList(2,k))& "','" &routineJavascriptString(aList(3,k))& "','" &routineJavascriptString(aList(4,k))& "','" &routineJavascriptString(aList(5,k))& "','" &routineJavascriptString(aList(6,k))& "')"">" &aList(2,k)& "</a></td>"&vbCrLf
	Response.Write "<td><a href=""#"" onClick=""return go_back('"&aList(0,k)& "', '" &routineJavascriptString(aList(1,k))& "','" &routineJavascriptString(aList(2,k))& "','" &routineJavascriptString(aList(3,k))& "','" &routineJavascriptString(aList(4,k))& "','" &routineJavascriptString(aList(5,k))& "','" &routineJavascriptString(aList(6,k))& "')"">" &aList(3,k)& "</a></td>"&vbCrLf
	Response.Write "<td><a href=""#"" onClick=""return go_back('"&aList(0,k)& "', '" &routineJavascriptString(aList(1,k))& "','" &routineJavascriptString(aList(2,k))& "','" &routineJavascriptString(aList(3,k))& "','" &routineJavascriptString(aList(4,k))& "','" &routineJavascriptString(aList(5,k))& "','" &routineJavascriptString(aList(6,k))& "')"">" &aList(4,k)& "</a></td>"&vbCrLf
	Response.Write "<td><a href=""#"" onClick=""return go_back('"&aList(0,k)& "', '" &routineJavascriptString(aList(1,k))& "','" &routineJavascriptString(aList(2,k))& "','" &routineJavascriptString(aList(3,k))& "','" &routineJavascriptString(aList(4,k))& "','" &routineJavascriptString(aList(5,k))& "','" &routineJavascriptString(aList(6,k))& "')"">" &aList(5,k)& "</a></td>"&vbCrLf
	Response.Write "<td><a href=""#"" onClick=""return go_back('"&aList(0,k)& "', '" &routineJavascriptString(aList(1,k))& "','" &routineJavascriptString(aList(2,k))& "','" &routineJavascriptString(aList(3,k))& "','" &routineJavascriptString(aList(4,k))& "','" &routineJavascriptString(aList(5,k))& "','" &routineJavascriptString(aList(6,k))& "')"">" &aList(6,k)& "</a></td>"&vbCrLf
	else
	Response.Write "<TD NOWRAP><a target=""_parent"" href=""GeocodeDetail.asp?Geocode="&aList(1,k)&""">"&routineHtmlString(aList(0,k))&"</a></TD>"&vbCrLf
	Response.Write "<TD NOWRAP><a target=""_parent"" href=""GeocodeDetail.asp?Geocode="&aList(1,k)&""">"&routineHtmlString(aList(1,k))&"</a></TD>"&vbCrLf
	Response.Write "<TD NOWRAP><a target=""_parent"" href=""GeocodeDetail.asp?Geocode="&aList(1,k)&""">"&routineHtmlString(aList(2,k))&"</a></TD>"&vbCrLf
	Response.Write "<TD NOWRAP><a target=""_parent"" href=""GeocodeDetail.asp?Geocode="&aList(1,k)&""">"&routineHtmlString(aList(3,k))&"</a></TD>"&vbCrLf
	Response.Write "<TD NOWRAP><a target=""_parent"" href=""GeocodeDetail.asp?Geocode="&aList(1,k)&""">"&routineHtmlString(aList(4,k))&"</a></TD>"&vbCrLf
	Response.Write "<TD NOWRAP><a target=""_parent"" href=""GeocodeDetail.asp?Geocode="&aList(1,k)&""">"&routineHtmlString(aList(5,k))&"</a></TD>"&vbCrLf
	Response.Write "<TD NOWRAP><a target=""_parent"" href=""GeocodeDetail.asp?Geocode="&aList(1,k)&""">"&routineHtmlString(aList(6,k))&"</a></TD>"&vbCrLf
	Response.Write "</TR>"
	end if
next
%>
</TBODY>
<TFOOT>
<TR>
<TD align=left colSpan=8>
	<input type=hidden name=txtPageNumber value=<%=intPageNumber%>>
	<input type="submit" name="action" value="&lt;&lt;">
	<input type="submit" name="action" value="&lt;">
	<input type="text" name="txtGoToPageNo" title="You can jump to a specific page by typing the page number in this box" onclick="document.frmGeocodeList.txtGoToPageNo.value=''" value="page <%=intPageNumber%> of <%=intPageCount%>" style="HEIGHT: 22px; WIDTH: 150px">
	<input type="submit" name="action" value="&gt;">
	<input type="submit" name="action" value="&gt;&gt;">
</TD>
</TR>
</TFOOT>
<CAPTION>Records <%=m+1%> to <%=n+1%> of <%=UBound(aList, 2)+1 & " records"%></CAPTION>
</TABLE>
</FORM>
</BODY>
</HTML>

<%@ Language=VBScript %>
<% Option Explicit
  on error resume next
%>
<!--#include file="smaconstants.inc"-->
<!--#include file="databaseconnect.asp"-->
<!--#include file="smaProcs.inc"-->
<!--
*************************************************************************************
* File Name:	FacilityList.asp
*
* Purpose:
*
* In Param:
*
* Out Param:
*
* Created By:
* Edited by:    Adam Haydey Jan 25, 2001
*               Added Customer Service City, Customer Service Address, and Past Facility Start Date search fields.
*				Also Added ADSL Due Date and Facility Start Date to the results.
**************************************************************************************
		 Date		Author			Changes/enhancements made
		20-Jul-01	 DTy		When 'Active Only' is selected:
		                          Exclude Circuit that are marked as soft deleted or has
		                            status of 'TERM'.
		                          Exclude Customers that are marked as soft deleted.
		                          Exclude Customer Services that are marked as soft deleted.
		                          Exclude Service Locations that are marked as soft deleted.
		23-Jan-02	DTy			Add Facility Provider/On/Off Net search criterias.
        18-Feb-02	DTy		    Active customers are those whose status is either
                                  'Prospect', 'OnHold' or 'Current'.
       28-Feb-02	 DTy		Include Customer Service Desc Alias when searching for Customer
                                  Service names.
**************************************************************************************
-->
<%
Dim StrFacName,StrCustServA,StrServLocA,StrFacType,StrRegion,strWhereClause,strOrderBy
Dim StrActive,StrSql,StrUtAdsl,StrFacStat,StrCusta,StrOsPvc,StrFromClause, strWinName,StrAdslTyp, strToDate, strFromDate
Dim StrServiceCity, StrServiceAddress, StrPastFacStart
Dim strFacPrvdr, strOnOffNet

StrFacName = routineOraString(UCase(TRIM(Request("txtfacname"))))
StrFacType = routineOraString(UCase(TRIM(Request("selfactyp"))))
StrFacStat = routineOraString(UCase(TRIM(Request("selfacstat"))))
StrRegion  = routineOraString(UCase(TRIM(Request("selRegion"))))
StrCustServA = routineOraString(UCase(TRIM(Request("txtcuserva"))))
StrServLocA  = routineOraString(UCase(TRIM(Request("txtservloca"))))
StrActive = routineOraString(UCase(TRIM(Request("chkactive"))))
StrUtAdsl = routineOraString(UCase(TRIM(Request("chkutadsl"))))
StrCusta  = routineOraString(UCase(TRIM(Request("txtcusta"))))
strToDate = TRIM(Request("hdnToDate"))
strFromDate = TRIM(Request("hdnFromDate"))
StrOsPvc =  UCase(TRIM(Request("chkoutstpvc")))
strWinName = Request("hdnWinName")
StrAdslTyp = UCase(TRIM(Request("selfacadsltyp")))

StrServiceCity  = routineOraString(UCase(TRIM(Request("txtservcity"))))
StrServiceAddress  = routineOraString(UCase(TRIM(Request("txtservadd"))))
StrPastFacStart = routineOraString(UCase(TRIM(Request("chkPastFacStart"))))

strFacPrvdr = Request("selFacPrvdr")
strOnOffNet = Request("selOnOffNet")

StrSql ="SELECT A.CIRCUIT_ID, " &_
		"A.CIRCUIT_NUMBER, " &_
		"A.CIRCUIT_NAME, " &_
		"A.CIRCUIT_TYPE_CODE," &_
		"A.CIRCUIT_STATUS_CODE, " &_
		"A.NOC_REGION_LCODE, " &_
		"B.CUSTOMER_SERVICE_DESC," &_
		"C.SERVICE_LOCATION_NAME LOCATION_A, " &_
		"D.SERVICE_LOCATION_NAME LOCATION_B, " &_
		"E.CUSTOMER_NAME, " &_
		"TO_CHAR(A.CIRCUIT_START_DATE, 'MON-DD-YYYY'), " &_
		"DECODE(A.CIRCUIT_TYPE_CODE,'ADSL',A.ADSL_TYPE_CODE,A.CIRCUIT_TYPE_CODE) "

IF(StrFacType = "ATMPVC") THEN
	StrSql =StrSql & ", F.CUSTOMER_SERVICE_DESC SERV_B,  " &_
					"	TO_CHAR(A.CREATE_DATE_TIME, 'MON-DD-YYYY') "
ELSE
	StrSql =StrSql & ", TO_CHAR(A.ADSL_DUE_DATE, 'MON-DD-YYYY')  "
END IF

IF StrFacType <> "ATMPVC" THEN
	StrSql =StrSql & ", A.CIRCUIT_PROVIDER_CODE, DECODE(I.IS_ON_NET, 'Y', 'ON NET', ' ')"
END IF

StrFromClause = " FROM "&_
	"CRP.CIRCUIT A," &_
	"CRP.CUSTOMER_SERVICE B," &_
	"CRP.SERVICE_LOCATION C," &_
	"CRP.SERVICE_LOCATION D, " &_
	"CRP.CUSTOMER E"

IF (StrFacType = "ATMPVC") THEN
	StrFromClause= StrFromClause & ",CRP.CUSTOMER_SERVICE F"
ELSE 'Added by Adam Jan 2001
	IF (LEN(StrServiceCity) > 0) OR (LEN(StrServiceAddress) > 0) THEN
		StrFromClause= StrFromClause & ",CRP.ADDRESS G"
		StrFromClause= StrFromClause & ",CRP.ADDRESS H"
	END IF
END IF

IF StrFacType <> "ATMPVC" THEN
	StrFromClause= StrFromClause & ",CRP.CIRCUIT_PROVIDER I"
END IF

StrSql = StrSql & StrFromClause

IF StrFacType <> "ATMPVC" THEN
     IF StrUtAdsl = "YES" THEN
                 strWhereClause = "WHERE A.CUSTOMER_SERVICE_ID_A = B.CUSTOMER_SERVICE_ID(+) AND" &_
                 " A.SERVICE_LOCATION_ID_A = C.SERVICE_LOCATION_ID(+) AND" &_
                 " A.SERVICE_LOCATION_ID_B = D.SERVICE_LOCATION_ID(+) AND" &_
                 " A.CIRCUIT_START_DATE <= SYSDATE AND" &_
                 " A.CIRCUIT_TYPE_CODE = 'ADSL' AND" &_
                 " A.ADSL_TRAINED_SPEED IS NULL AND" &_
                 " B.CUSTOMER_ID != 49 AND" &_
                 " A.BILLING_CUSTOMER_ID_A = E.CUSTOMER_ID(+) "
     ELSE
                 strWhereClause =  " WHERE A.CUSTOMER_SERVICE_ID_A = B.CUSTOMER_SERVICE_ID(+) AND" &_
                  " A.SERVICE_LOCATION_ID_A = C.SERVICE_LOCATION_ID(+) AND" &_
                  " A.SERVICE_LOCATION_ID_B = D.SERVICE_LOCATION_ID(+) AND" &_
                  " A.BILLING_CUSTOMER_ID_A = E.CUSTOMER_ID(+) AND " &_
                  " A.CIRCUIT_TYPE_CODE NOT LIKE '%PVC%'"
   END IF
   'Added by Adam Jan 2001
	IF (LEN(StrServiceCity) > 0) OR (LEN(StrServiceAddress) > 0) THEN
		strWhereClause = strWhereClause & " AND  C.ADDRESS_ID = G.ADDRESS_ID (+) AND" &_
                  " D.ADDRESS_ID = H.ADDRESS_ID (+) "
	END IF
	'End Added by Adam Jan 2001

ELSE
   IF StrOsPvc = "YES" THEN
                strWhereClause = "WHERE A.CUSTOMER_SERVICE_ID_A = B.CUSTOMER_SERVICE_ID(+) AND" &_
                 " A.SERVICE_LOCATION_ID_A = C.SERVICE_LOCATION_ID(+) AND" &_
                 " A.SERVICE_LOCATION_ID_B = D.SERVICE_LOCATION_ID(+) AND" &_
                 " A.BILLING_CUSTOMER_ID_A = E.CUSTOMER_ID(+) AND " &_
                 " A.CUSTOMER_SERVICE_ID_B = F.CUSTOMER_SERVICE_ID(+) AND "&_
                 " (A.USAGE_CALCULATION_TYPE_CODE IS NULL OR A.CUSTOMER_SERVICE_ID_A IS NULL OR A.CUSTOMER_SERVICE_ID_B IS NULL) "
   ELSE
                  strWhereClause =  " WHERE A.CUSTOMER_SERVICE_ID_A = B.CUSTOMER_SERVICE_ID(+) AND" &_
                  " A.SERVICE_LOCATION_ID_A = C.SERVICE_LOCATION_ID(+) AND" &_
                  " A.SERVICE_LOCATION_ID_B = D.SERVICE_LOCATION_ID(+) AND" &_
                  " A.CUSTOMER_SERVICE_ID_B = F.CUSTOMER_SERVICE_ID(+) AND "&_
                  " A.BILLING_CUSTOMER_ID_A = E.CUSTOMER_ID(+) "
   END IF

END IF

IF  LEN(StrFacName) > 0 THEN
    strWhereClause = strWhereClause & " AND A.CIRCUIT_ID IN (SELECT CIRCUIT_ID FROM CRP.CIRCUIT X WHERE " &_
                     " UPPER(X.CIRCUIT_NAME) LIKE '" & StrFacName &"%' OR UPPER(X.CIRCUIT_NUMBER)  LIKE '" & StrFacName &"%'" &_
                     " UNION SELECT CIRCUIT_ID FROM CRP.CIRCUIT_NUMBER_ALIAS Y WHERE UPPER(Y.CIRCUIT_NUMBER_ALIAS) LIKE '" & StrFacName & "%'"

    IF  (StrActive="YES")  THEN
        strWhereClause = strWhereClause & " AND Y.RECORD_STATUS_IND = 'A'"
    END IF
    strWhereClause = strWhereClause & ")"
END IF


IF  LEN(StrFacType) > 0 AND StrUtAdsl <> "YES" THEN
    strWhereClause = strWhereClause & " AND UPPER(A.CIRCUIT_TYPE_CODE) LIKE '" & StrFacType &"%'"
END IF

' For Facility only
IF StrFacType <> "ATMPVC" THEN

   ' Extract Facility Providers
   strWhereClause = strWhereClause & " AND A.CIRCUIT_PROVIDER_CODE = I.CIRCUIT_PROVIDER_CODE"
   IF LEN(StrFacPrvdr) > 0 THEN
      strWhereClause = strWhereClause & " AND A.CIRCUIT_PROVIDER_CODE = '" & strFacPrvdr & "'"
   END IF

   ' Extract On/Off net
   IF LEN(strOnOffNet) > 0 THEN
      IF strOnOffNet = "ON" THEN
         strWhereClause = strWhereClause & " AND I.IS_ON_NET = 'Y'"
      ELSE
         strWhereClause = strWhereClause & " AND I.IS_ON_NET = 'N'"
      END IF
   END IF
END IF

IF  LEN(StrFacStat) > 0 THEN
    strWhereClause = strWhereClause & " AND UPPER(A.CIRCUIT_STATUS_CODE) LIKE '" & StrFacStat &"%'"
END IF

IF  (LEN(StrAdslTyp) > 0 AND (StrFacType <> "ATMPVC"))  THEN
    strWhereClause = strWhereClause & " AND UPPER(A.ADSL_TYPE_CODE) = '" & StrAdslTyp & "'"
END IF

IF  LEN(StrRegion) > 0 THEN
    strWhereClause = strWhereClause & " AND UPPER(A.NOC_REGION_LCODE) = '" & StrRegion & "'"
END IF

IF  (LEN(StrCustServA) > 0 AND (StrFacType <> "ATMPVC")) THEN
	strWhereClause = strWhereClause & " AND b.customer_service_id in (" &_
		            " select customer_service_id from crp.customer_service where " & rtRmvSpChr("customer_service_desc", "Y") & " like '%" & rtRmvSpChr(strCustServA, "N") & "%' union" &_
                    " select customer_service_id from crp.customer_service_desc_alias where " & rtRmvSpChr("customer_service_desc_alias", "Y") & " like '%" & rtRmvSpChr(strCustServA, "N") & "%')"
END IF

IF  (LEN(StrCustServA) > 0 AND (StrFacType = "ATMPVC")) THEN
	strWhereClause = strWhereClause & " AND (b.customer_service_id in (" &_
		            " select customer_service_id from crp.customer_service where " & rtRmvSpChr("customer_service_desc", "Y") & " like '%" & rtRmvSpChr(strCustServA, "N") & "%' union" &_
                    " select customer_service_id from crp.customer_service_desc_alias where " & rtRmvSpChr("customer_service_desc_alias", "Y") & " like '%" & rtRmvSpChr(strCustServA, "N") & "%')" &_
	                " OR f.customer_service_id in (" &_
		            " select customer_service_id from crp.customer_service where " & rtRmvSpChr("customer_service_desc", "Y") & " like '%" & rtRmvSpChr(strCustServA, "N") & "%' union" &_
                    " select customer_service_id from crp.customer_service_desc_alias where " & rtRmvSpChr("customer_service_desc_alias", "Y") & " like '%" & rtRmvSpChr(strCustServA, "N") & "%'))"
END IF


IF  LEN(StrServLocA) > 0 THEN
    strWhereClause = strWhereClause & " AND (UPPER(C.SERVICE_LOCATION_NAME) LIKE '" & StrServLocA &"%' OR  UPPER(D.SERVICE_LOCATION_NAME) LIKE '" & StrServLocA &"%')"
END IF

IF (StrFacType = "ATMPVC") THEN

	IF len(strToDate) > 0  and len(strFromDate) > 0 THEN
		strWhereClause = strWhereClause & " and A.CREATE_DATE_TIME BETWEEN TO_DATE('" &  strFromDate & " 00:01', 'MM/DD/YYYY HH24:MI') AND TO_DATE('" & strToDate & " 23:59', 'MM/DD/YYYY HH24:MI') "

	END IF

END IF


IF  LEN(StrCusta) > 0 THEN
    strWhereClause = strWhereClause & " AND (A.BILLING_CUSTOMER_ID_A IN " &_
                    "(SELECT A.CUSTOMER_ID FROM CRP.CUSTOMER_NAME_ALIAS A, CRP.CUSTOMER C" &_
                    " WHERE A.CUSTOMER_NAME_ALIAS_UPPER LIKE '" & StrCusta & "%' AND " &_
                    " A.CUSTOMER_ID = C.CUSTOMER_ID "
    IF  (StrActive="YES")  THEN
        strWhereClause = strWhereClause & " AND C.RECORD_STATUS_IND = 'A' AND " &_
	                 "C.CUSTOMER_STATUS_LCODE IN ('Prospect', 'Current', 'OnHold')"
    END IF

    strWhereClause = strWhereClause &  ") OR A.BILLING_CUSTOMER_ID_B IN " &_
                    "(SELECT A.CUSTOMER_ID FROM CRP.CUSTOMER_NAME_ALIAS A, CRP.CUSTOMER C" &_
                    " WHERE A.CUSTOMER_NAME_ALIAS_UPPER LIKE '" & StrCusta & "%' AND " &_
                    " A.CUSTOMER_ID = C.CUSTOMER_ID "
    IF  (StrActive="YES")  THEN
        strWhereClause = strWhereClause & " AND C.RECORD_STATUS_IND = 'A' AND " &_
	                 "C.CUSTOMER_STATUS_LCODE IN ('Prospect', 'Current', 'OnHold')"
    END IF

    strWhereClause = strWhereClause & "))"
END IF
' Added by Adam Jan 2001
IF ((LEN(StrServiceCity) > 0) AND (StrFacType <> "ATMPVC")) THEN
	strWhereClause = strWhereClause & " AND (UPPER(G.MUNICIPALITY_NAME) LIKE '" & StrServiceCity &"%' OR  UPPER(H.MUNICIPALITY_NAME) LIKE '" & StrServiceCity &"%')"
END IF

IF ((LEN(StrPastFacStart) > 0) AND (StrFacType <> "ATMPVC")) THEN
	strWhereClause = strWhereClause & " AND A.CIRCUIT_START_DATE < SYSDATE " &_
									  " AND A.CIRCUIT_STATUS_CODE in ('DEFINE', 'UNKNWN') "
END IF

IF ((LEN(StrServiceAddress) > 0) AND (StrFacType <> "ATMPVC")) THEN
	      strWhereClause = strWhereClause & " AND ((Upper(NVL(G.BUILDING_NAME,'NO BUILDING NAME') ||CHR(13)||CHR(10)|| " &_
					"decode(G.APARTMENT_NUMBER, null, null, rtrim(G.APARTMENT_NUMBER) || ' ') || " &_
					"decode(to_char(G.HOUSE_NUMBER) || G.HOUSE_NUMBER_SUFFIX, null, null, rtrim(to_char(G.house_number) || G.house_number_suffix)  || ' ') || " &_
					"decode(G.STREET_VECTOR, null, null, rtrim(G.STREET_VECTOR) || ' ') || " &_
					"NVL(G.LONG_STREET_NAME,'NO STREET ADDRESS')||CHR(13)||CHR(10)|| " &_
					"NVL(G.MUNICIPALITY_NAME,'NO MUNICIPALITY')||' '|| " &_
					"NVL(G.PROVINCE_STATE_LCODE,'NO PROVINCE')||CHR(13)||CHR(10)|| " &_
					"NVL(G.POSTAL_CODE_ZIP,'NO POSTAL CODE'))  LIKE '" & routineOraString(strServiceAddress) & "%' )"
		  strWhereClause = strWhereClause & " OR (Upper(NVL(H.BUILDING_NAME,'NO BUILDING NAME') ||CHR(13)||CHR(10)|| " &_
					"decode(H.APARTMENT_NUMBER, null, null, rtrim(H.APARTMENT_NUMBER) || ' ') || " &_
					"decode(to_char(H.HOUSE_NUMBER) || H.HOUSE_NUMBER_SUFFIX, null, null, rtrim(to_char(H.house_number) || H.house_number_suffix)  || ' ') || " &_
					"decode(H.STREET_VECTOR, null, null, rtrim(H.STREET_VECTOR) || ' ') || " &_
					"NVL(H.LONG_STREET_NAME,'NO STREET ADDRESS')||CHR(13)||CHR(10)|| " &_
					"NVL(H.MUNICIPALITY_NAME,'NO MUNICIPALITY')||' '|| " &_
					"NVL(H.PROVINCE_STATE_LCODE,'NO PROVINCE')||CHR(13)||CHR(10)|| " &_
					"NVL(H.POSTAL_CODE_ZIP,'NO POSTAL CODE'))  LIKE '" & routineOraString(strServiceAddress) & "%' ))"
END IF
' End added by Adam Jan 2001

IF  (StrActive="YES")  THEN
    IF  LEN(StrFacStat) = 0 THEN
		strWhereClause = strWhereClause & " AND A.CIRCUIT_STATUS_CODE <> 'TERM'"
	END IF
	strWhereClause = strWhereClause & " AND A.RECORD_STATUS_IND = 'A' AND " & _
	                 "B.RECORD_STATUS_IND (+) = 'A' AND C.RECORD_STATUS_IND (+) = 'A' AND " & _
	                 "D.RECORD_STATUS_IND (+) = 'A' AND E.RECORD_STATUS_IND (+) = 'A' AND " & _
	                 "(E.customer_status_lcode IS NULL OR E.customer_status_lcode IN ('Prospect', 'Current', 'OnHold'))"

	IF (LEN(StrServiceCity) > 0) OR (LEN(StrServiceAddress) > 0) THEN
	    strWhereClause = strWhereClause & " AND F.RECORD_STATUS_IND (+) = 'A' and G.RECORD_STATUS_IND (+) = 'A'"
	END IF
END IF

IF (StrFacType <> "ATMPVC" AND StrUtAdsl = "YES") THEN
	strOrderBy = "ORDER BY UPPER(A.CIRCUIT_NAME)"
ELSEIF (StrFacType = "ATMPVC" AND StrOsPvc = "YES") THEN
	strOrderBy = "ORDER BY UPPER(A.CIRCUIT_NAME)"
ELSE
	strOrderBy = "ORDER BY UPPER(A.CIRCUIT_NAME),A.CIRCUIT_NUMBER,A.CIRCUIT_TYPE_CODE"
END IF

StrSql =  StrSql & " "& strWhereClause&" "& strOrderBy

'Response.Write StrSql
'Response.end
'Create the command object
Dim objRs,Recordcnt,strbgcolor,aRecordSet,intPageCount,intPageNumber

'Create Recordset object
set objRS = objConn.Execute(StrSql)
Recordcnt = 0
strbgcolor = "<TR>"

if not objRS.EOF then
	aRecordSet = objRS.GetRows
else
	Response.Write "0 records found"
	Response.end
end if


'Response.write "<b>Total=" & Recordcnt & "</b>"
'Clean up our ADO objects
objRS.close
set objRS = Nothing

objConn.close
set ObjConn = Nothing

intPageCount = Int(UBound(aRecordSet, 2) / intConstDisplayPageSize)+1
on error resume next
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

						if (StrFacType <> "ATMPVC") then
						set objTxtStream = objFSO.CreateTextFile(strExportPath&strRealUserID&"-facility.xls", true, false)
						else
						set objTxtStream = objFSO.CreateTextFile(strExportPath&strRealUserID&"-pvc.xls", true, false)
						end if

						If err Then
						DisplayError "CLOSE", err.Number, "FacilityList.asp - Cannot create Excel spreadsheet file due to the following errors.  Please contact your system administrator.", err.Description
					    End If

						with objTxtStream
							.WriteLine "<table border=1>"

							'export the header
							.WriteLine "<TR>"
'							.WriteLine "<TR bgcolor=#ffcc99>"
							if (StrFacType <> "ATMPVC") then
							.WriteLine "<TH>Facility Name</TH>"
							.WriteLine "<TH>Facility Number</TH>"
							.WriteLine "<TH>Type</TH>"
							.WriteLine "<TH>Facility Provider</TH>"
							.WriteLine "<TH>ON/OFF Net</TH>"
							.WriteLine "<TH>Customer Service A</TH>"
							.WriteLine "<TH>Service Location A</TH>"
							.WriteLine "<TH>Service Location B</TH>"
							.WriteLine "<TH>Region</TH>"
							.WriteLine "<TH>Circuit Status</TH>"
								'IF ( StrUtAdsl = "YES") then
									'.WriteLine "<TH>Start Date</TD>"
								'end if
							.WriteLine "<TH>ADSL Due Date</TH>"
							.WriteLine "<TH>Facility Start Date</TH>"
							.WriteLine "</TR>"
							else
							.WriteLine "<TH>PVC Name</TH>"
							.WriteLine "<TH>PVC Number</TH>"
							.WriteLine "<TH>PVC Type</TH>"
							.WriteLine "<TH>Customer Service A</TH>"
							.WriteLine "<TH>Service Location A</TH>"
							.WriteLine "<TH>Customer Service B</TH>"
							.WriteLine "<TH>Service Location B</TH>"
							.WriteLine "<TH>Region</TH>"
							.WriteLine "<TH>PVC Status</TH>"
							.WriteLine "<TH>Create Date</TH>"
							.WriteLine "</TR>"

							end if
							'export the body
							for k = 0 to UBound(aRecordSet, 2)
								'Alternate row background colour
								if Int(k/2) = k/2 then
'									.WriteLine "<TR bgcolor=#ffffcc>"
									.WriteLine "<TR>"
								else
'									.WriteLine "<TR bgcolor=#ffffff>"
									.WriteLine "<TR>"
								end if

								.WriteLine "<TD NOWRAP>"&routineHtmlString(aRecordSet(2,k))&"</TD>"
								.WriteLine "<TD NOWRAP>"&routineHtmlString(aRecordSet(1,k))&"</TD>"
								.WriteLine "<TD NOWRAP>"&routineHtmlString(aRecordSet(11,k))&"</TD>"
								if (StrFacType <> "ATMPVC") then
								   .WriteLine "<TD NOWRAP>"&routineHtmlString(aRecordSet(13,k))&"</TD>"
								   .WriteLine "<TD NOWRAP>"&routineHtmlString(aRecordSet(14,k))&"</TD>"
								end if

								.WriteLine "<TD NOWRAP>"&routineHtmlString(aRecordSet(6,k))&"</TD>"
								.WriteLine "<TD NOWRAP>"&routineHtmlString(aRecordSet(7,k))&"</TD>"
								if (StrFacType = "ATMPVC") then
								.WriteLine "<TD NOWRAP>"&routineHtmlString(aRecordSet(12,k))&"</TD>"
								end if
								.WriteLine "<TD NOWRAP>"&routineHtmlString(aRecordSet(8,k))&"</TD>"
								.WriteLine "<TD NOWRAP>"&routineHtmlString(aRecordSet(5,k))&"</TD>"
								.WriteLine "<TD NOWRAP>"&routineHtmlString(aRecordSet(4,k))&"</TD>"
								.WriteLine "<TD NOWRAP>"&routineHtmlString(aRecordSet(13,k))&"</TD>"
								'if ((StrUtAdsl = "YES") and (StrFacType <> "ATMPVC")) then
								'.WriteLine "<TD NOWRAP>"&routineHtmlString(aRecordSet(10,k))&"</TD>"
								'end if
								if (StrFacType <> "ATMPVC") then
								.WriteLine "<TD NOWRAP>"&routineHtmlString(aRecordSet(12,k))&"</TD>"
								.WriteLine "<TD NOWRAP>"&routineHtmlString(aRecordSet(10,k))&"</TD>"
								end if
								.WriteLine "</TR>"
							next
							.WriteLine "</table>"
						end with
						objTxtStream.Close
						if (StrFacType <> "ATMPVC") then
							strsql = "<script type=""text/javascript"">document.location=""export/"&strRealUserID&"-facility.xls"";</script>"
							Response.Write strsql
							Response.End
	'						Response.redirect "export/"&strRealUserID&"-facility.xls"
						else
							strsql = "<script type=""text/javascript"">document.location=""export/"&strRealUserID&"-pvc.xls"";</script>"
							Response.Write strsql
							Response.End
	'						Response.redirect "export/"&strRealUserID&"-pvc.xls"
						end if
					elseif Request("txtGoToPageNo") <> "" then
						intPageNumber = CInt(Request("txtGoToPageNo"))
					else
						intPageNumber = 1
					end if
end select

if intPageNumber < 1 then intPageNumber = 1
if intPageNumber > intPageCount then intPageNumber = intPageCount
dim k, m, n
m = (intPageNumber-1) * intConstDisplayPageSize
n = (intPageNumber) * intConstDisplayPageSize-1
if n > UBound(aRecordSet, 2) then
	n = UBound(aRecordSet, 2)
end if

'check if the client is still connected
if response.isclientconnected = false then
	Response.End
end if
%>

<HTML>
<HEAD>
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<LINK REL="stylesheet" TYPE="text/css" HREF="stylesheets/styles.css">
<script type="text/javascript">
function go_back(strFAC_ID, strFAC_NAME, strFAC_TYPE) {
	parent.opener.document.forms[0].hdnNewElementID.value = strFAC_ID;
	parent.opener.document.forms[0].hdnNewElementName.value = unescape(strFAC_NAME);
	parent.opener.document.forms[0].hdnNewElementType.value = "FAC";
	parent.opener.btn_iFrmAddNewElement();
	parent.window.close();
}
</script>
</HEAD>

<BODY>
<FORM  id=fmFacList name=fmFacList action="FacilityList.asp" method=post>
<input type="hidden" name="hdnExport" value>
<TABLE border=1 cellPadding=2 cellSpacing=0 width="100%">
<THEAD>
	<TR>
		<%if (StrFacType <> "ATMPVC") THEN %>
			<TH align=left NOWRAP  WIDTH="10%"><B>Facility Name</B></TH>
			<TH align=left NOWRAP WIDTH="10%"><B>Facility Number</B></TH>
			<TH align=left NOWRAP WIDTH="10%"><B>Type</B></TH>
			<TH align=left NOWRAP WIDTH="10%"><B>Facility Provider </B></TH>
			<TH align=left NOWRAP WIDTH="10%"><B>ON/OFF Net </B></TH>
			<!--<TH align=left NOWRAP WIDTH="10%"><B>Customer A</B></TH> -->
			<TH align=left NOWRAP  WIDTH="10%"><B>Customer Service A</B></TH>
			<TH align=left NOWRAP WIDTH="10%"><B>Service Location A</B></TH>
			<TH align=left NOWRAP WIDTH="10%"><B>Service Location B</B></TH>
			<TH align=left NOWRAP  WIDTH="10%"><B>Region</B></TD>
			<TH align=left NOWRAP WIDTH="10%"><B>Circuit Status</B></TD>
			<%'IF ( StrUtAdsl = "YES") then %>
				<!--TH align=left NOWRAP  WIDTH="10%"><B>Start Date</B></TD-->

			<%' end if %>
			<TH align=left NOWRAP  WIDTH="10%"><B>ADSL Due Date</B></TD>
			<TH align=left NOWRAP  WIDTH="10%"><B>Facility Start Date</B></TD>
		<% else %>
			<TH align=left  NOWRAP WIDTH="10%"><B>PVC Name</B></TH>
			<TH align=left NOWRAP WIDTH="10%"><B>PVC Number</B></TH>
			<TH align=left NOWRAP WIDTH="10%"><B>PVC Type</B></TH>
			<!--<TH align=left NOWRAP WIDTH="10%"><B>Customer A</B></TH> -->
			<TH align=left NOWRAP WIDTH="10%"><B>Customer Service A</B></TH>
			<TH align=left NOWRAP WIDTH="10%"><B>Service Location A</B></TH>
			<TH align=left NOWRAP WIDTH="10%"><B>Customer Service B</B></TH>
			<TH align=left NOWRAP WIDTH="10%"><B>Service Location B</B></TH>
			<TH align=left NOWRAP WIDTH="10%"><B>Region</B></TD>
			<TH align=left NOWRAP WIDTH="10%"><B>PVC Status</B></TD>
			<TH align=left NOWRAP WIDTH="10%"><B>Create Date</B></TD>
		<% end if %>
	</TR>
</THEAD>

<TBODY>
<%

for k = m to n
'Do while Not objRS.EOF
	'Alternate table background colour
	if Int(k/2) = k/2 then
		Response.Write "<TR>"
	else
		Response.Write "<TR bgcolor=White>"
	end if

	if strWinName = "Popup" then
	%>
		<TD align=left NOWRAP WIDTH="10%"><a href="#" onClick="go_back(<%=aRecordSet(0,k)%>, '<%=escape(aRecordSet(1,k))%>', '<%=escape(aRecordSet(3,k))%>');"><%=aRecordSet(2,k)%></A>&nbsp;</TD>
		<TD align=left NOWRAP WIDTH="10%"><a href="#" onClick="go_back(<%=aRecordSet(0,k)%>, '<%=escape(aRecordSet(1,k))%>', '<%=escape(aRecordSet(3,k))%>');"><%=aRecordSet(1,k)%></A>&nbsp;</TD>
		<TD align=left NOWRAP WIDTH="10%"><a href="#" onClick="go_back(<%=aRecordSet(0,k)%>, '<%=escape(aRecordSet(1,k))%>', '<%=escape(aRecordSet(3,k))%>');"><%=aRecordSet(3,k)%></A>&nbsp;</TD>
		<TD align=left NOWRAP WIDTH="10%"><a href="#" onClick="go_back(<%=aRecordSet(0,k)%>, '<%=escape(aRecordSet(1,k))%>', '<%=escape(aRecordSet(3,k))%>');"><%=aRecordSet(6,k)%></A>&nbsp;</TD>
		<TD align=left NOWRAP WIDTH="10%"><a href="#" onClick="go_back(<%=aRecordSet(0,k)%>, '<%=escape(aRecordSet(1,k))%>', '<%=escape(aRecordSet(3,k))%>');"><%=aRecordSet(7,k)%></A>&nbsp;</TD>
		<%if (StrFacType = "ATMPVC") then %>
			<TD align=left NOWRAP WIDTH="10%"><a href ="#" onClick="go_back(<%=aRecordSet(0,k)%>, '<%=escape(aRecordSet(1,k))%>', '<%=escape(aRecordSet(3,k))%>');"><%=aRecordSet(12,k)%></A>&nbsp;</TD>
		<%end if%>
		<TD align=left NOWRAP WIDTH="10%"><a href="#" onClick="go_back(<%=aRecordSet(0,k)%>, '<%=escape(aRecordSet(1,k))%>', '<%=escape(aRecordSet(3,k))%>');"><%=aRecordSet(8,k)%></A>&nbsp;</TD>
		<TD align=left NOWRAP WIDTH="10%"><a href="#" onClick="go_back(<%=aRecordSet(0,k)%>, '<%=escape(aRecordSet(1,k))%>', '<%=escape(aRecordSet(3,k))%>');"><%=aRecordSet(5,k)%></A>&nbsp;</TD>
		<TD align=left NOWRAP WIDTH="10%"><a href="#" onClick="go_back(<%=aRecordSet(0,k)%>, '<%=escape(aRecordSet(1,k))%>', '<%=escape(aRecordSet(3,k))%>');"><%=aRecordSet(4,k)%></A>&nbsp;</TD>

		<%'if ((StrUtAdsl = "YES") and (StrFacType <> "ATMPVC")) then%>
			<!--TD align=left NOWRAP WIDTH="10%"><a href="#" onClick="go_back(<%=aRecordSet(0,k)%>, '<%=escape(aRecordSet(1,k))%>', '<%=escape(aRecordSet(3,k))%>');"><%=aRecordSet(10,k)%></A>&nbsp;</TD-->
		<%'end if
		if (StrFacType <> "ATMPVC") then%>
			<TD align=left NOWRAP WIDTH="10%"><a href="#" onClick="go_back(<%=aRecordSet(0,k)%>, '<%=escape(aRecordSet(1,k))%>', '<%=escape(aRecordSet(3,k))%>');"><%=aRecordSet(12,k)%></A>&nbsp;</TD>
			<TD align=left NOWRAP WIDTH="10%"><a href="#" onClick="go_back(<%=aRecordSet(0,k)%>, '<%=escape(aRecordSet(1,k))%>', '<%=escape(aRecordSet(3,k))%>');"><%=aRecordSet(10,k)%></A>&nbsp;</TD>
		<%else %>
			<TD align=left NOWRAP WIDTH="10%"><a href="#" onClick="go_back(<%=aRecordSet(0,k)%>, '<%=escape(aRecordSet(1,k))%>', '<%=escape(aRecordSet(3,k))%>');"><%=aRecordSet(13,k)%></A>&nbsp;</TD>
		<%end if
	else
	%>
		<TD align=left NOWRAP  WIDTH="10%"><a href="FacilityDetail.asp?CircuitID=<%=aRecordSet(0,k)%>&CircuitTyp=<%=aRecordSet(3,k)%>" TARGET="_parent"><%=aRecordSet(2,k)%></A>&nbsp;</TD>
		<TD align=left NOWRAP WIDTH="10%"><a href="FacilityDetail.asp?CircuitID=<%=aRecordSet(0,k)%>&CircuitTyp=<%=aRecordSet(3,k)%>" TARGET="_parent"><%=aRecordSet(1,k)%></A>&nbsp;</TD>
		<TD align=left NOWRAP WIDTH="10%"><a href="FacilityDetail.asp?CircuitID=<%=aRecordSet(0,k)%>&CircuitTyp=<%=aRecordSet(3,k)%>" TARGET="_parent"><%=aRecordSet(11,k)%></A>&nbsp;</TD>

		<%if (StrFacType <> "ATMPVC") then %>
		     <TD align=left NOWRAP WIDTH="10%"><a href="FacilityDetail.asp?CircuitID=<%=aRecordSet(0,k)%>&CircuitTyp=<%=aRecordSet(3,k)%>" TARGET="_parent"><%=aRecordSet(13,k)%></A>&nbsp;</TD>
		     <TD align=left NOWRAP WIDTH="10%"><a href="FacilityDetail.asp?CircuitID=<%=aRecordSet(0,k)%>&CircuitTyp=<%=aRecordSet(3,k)%>" TARGET="_parent"><%=aRecordSet(14,k)%></A>&nbsp;</TD>
		<%end if%>

		<TD align=left NOWRAP WIDTH="10%"><a href="FacilityDetail.asp?CircuitID=<%=aRecordSet(0,k)%>&CircuitTyp=<%=aRecordSet(3,k)%>" TARGET="_parent"><%=aRecordSet(6,k)%></A>&nbsp;</TD>
		<TD align=left NOWRAP WIDTH="10%"><a href="FacilityDetail.asp?CircuitID=<%=aRecordSet(0,k)%>&CircuitTyp=<%=aRecordSet(3,k)%>" TARGET="_parent"><%=aRecordSet(7,k)%></A>&nbsp;</TD>
		<%if (StrFacType = "ATMPVC") then %>
			<TD align=left NOWRAP WIDTH="10%"><a href="FacilityDetail.asp?CircuitID=<%=aRecordSet(0,k)%>&CircuitTyp=<%=aRecordSet(3,k)%>" TARGET="_parent"><%=aRecordSet(12,k)%></A>&nbsp;</TD>
		<%end if%>
		<TD align=left NOWRAP  WIDTH="10%"><a href="FacilityDetail.asp?CircuitID=<%=aRecordSet(0,k)%>&CircuitTyp=<%=aRecordSet(3,k)%>" TARGET="_parent"><%=aRecordSet(8,k)%></A>&nbsp;</TD>
		<TD align=left NOWRAP WIDTH="10%"><a href="FacilityDetail.asp?CircuitID=<%=aRecordSet(0,k)%>&CircuitTyp=<%=aRecordSet(3,k)%>" TARGET="_parent"><%=aRecordSet(5,k)%></A>&nbsp;</TD>
		<TD align=left NOWRAP WIDTH="10%"><a href="FacilityDetail.asp?CircuitID=<%=aRecordSet(0,k)%>&CircuitTyp=<%=aRecordSet(3,k)%>" TARGET="_parent"><%=aRecordSet(4,k)%></A>&nbsp;</TD>

		<%if (StrFacType <> "ATMPVC") then%>
			<TD align=left  NOWRAP WIDTH="10%"><a href="FacilityDetail.asp?CircuitID=<%=aRecordSet(0,k)%>&CircuitTyp=<%=aRecordSet(3,k)%>" TARGET="_parent"><%=aRecordSet(12,k)%></A>&nbsp;</TD>
			<TD align=left  NOWRAP WIDTH="10%"><a href="FacilityDetail.asp?CircuitID=<%=aRecordSet(0,k)%>&CircuitTyp=<%=aRecordSet(3,k)%>" TARGET="_parent"><%=aRecordSet(10,k)%></A>&nbsp;</TD>
		<%else%>
			<TD align=left NOWRAP WIDTH="10%"><a href="FacilityDetail.asp?CircuitID=<%=aRecordSet(0,k)%>&CircuitTyp=<%=aRecordSet(3,k)%>" TARGET="_parent"><%=aRecordSet(13,k)%></A>&nbsp;</TD>
		<%end if %>

		<%'if ((StrUtAdsl = "YES") and (StrFacType <> "ATMPVC")) then%>
			<!--TD align=left  NOWRAP WIDTH="10%"><a href="FacilityDetail.asp?CircuitID=<%=aRecordSet(0,k)%>&CircuitTyp=<%=aRecordSet(3,k)%>" TARGET="_parent"><%=aRecordSet(10,k)%></A>&nbsp;</TD-->
		<%'end if

	end if
next
%>
</TBODY>

<tfoot>
<tr>
<td align=left colSpan=10>
    <input type=hidden name=txtfacname value="<%=StrFacName%>">
    <input type=hidden name=selfactyp value="<%=StrFacType%>">
    <input type=hidden name=selfacstat value="<%=StrFacStat%>">
    <input type=hidden name=selRegion value="<%=StrRegion%>">
    <input type=hidden name=txtcuserva value="<%= StrCustServA%>">
    <input type=hidden name=txtservloca value="<%= StrServLocA%>">
    <input type=hidden name=chkactive value="<%= StrActive%>">
    <input type=hidden name=chkutadsl value="<%= StrUtAdsl%>">
    <input type=hidden name=txtcusta value="<%= StrCusta%>">
    <input type=hidden name=chkoutstpvc value="<%=StrOsPvc%>">
    <input type=hidden name=hdnWinName value="<%=strWinName%>">
    <input type=hidden name=selfacadsltyp value="<%=StrAdslTyp%>">
    <input type=hidden name=hdnToDate value="<%=strToDate%>">
    <input type=hidden name=hdnFromDate value="<%=strFromDate%>">
	<input type=hidden name=txtservadd value="<%= StrServiceAddress%>">
    <input type=hidden name=chkPastFacStart value="<%= StrPastFacStart%>">
    <input type=hidden name=txtservcity value="<%= StrServiceCity%>">

	<input type="hidden" name="selfacPrvdr" value="<%=strfacPrvdr%>">
	<input type="hidden" name="selOnOffNet" value="<%=strOnOffNet%>">

    <input type=hidden name=txtPageNumber value=<%=intPageNumber%>>
	<input type="submit" name="action" value="&lt;&lt;">
	<input type="submit" name="action" value="&lt;">
	<input type="text" name="txtGoToPageNo" onClick="document.fmFacList.txtGoToPageNo.value=''" title="You can jump to a specific page by typing the page number in this box" value="page <%=intPageNumber%> of <%=intPageCount%>" style="HEIGHT: 22px; WIDTH: 150px">
	<input type="submit" name="action" value="&gt;">
	<input type="submit" name="action" value="&gt;&gt;">
	<img SRC="images/excel.gif" onclick="document.fmFacList.target='new';document.fmFacList.hdnExport.value='xls';document.fmFacList.submit();document.fmFacList.hdnExport.value='';document.fmFacList.target='_self';" WIDTH="32" HEIGHT="32">
</td>

</tr>
</tfoot>
<caption align=left>
Records <%=m+1%> to <%=n+1%> of <%=UBound(aRecordSet, 2)+1 & " records"%>
</caption>

</TABLE>
</FORM>
</BODY>

</HTML>

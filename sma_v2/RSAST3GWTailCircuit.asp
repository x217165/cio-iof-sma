<%@ LANGUAGE=VBSCRIPT %>
<%  
OPTION EXPLICIT
'on error resume next
Response.CacheControl="Private"
%>
<!-- #include file="smaConstants.inc" -->
<!-- #include file="smaProcs.inc" -->
<!-- #include file="databaseconnect.asp" -->
<!--
*************************************************************************************
* File Name:	RSAST3GWTailCircuit.asp
*
* Purpose:		List Gateway Tail Circuits
*
* In Param:
*
* Out Param:
*
* Created By:
**************************************************************************************

        Date		Author			Changes/enhancements made
        -----		------		------------------------------------------------------
       10-22-01	     DTy		Change field names and variables.

                                WAN IP             to WAN IP Port Address
                                PNG IP             to LAN IP Port Address
                                PNG_IP_ID          to LAN_IP_ID
                                PNG_IP             to LAN_ID
                                PNG_IP_ADDRESS     to LAN_IP_ADDRESS

                                Delete WAN IP DLCI (txtWANDLCI & strWANIPDLCI)
                                Delete POS IP DLCI (txtPOSIPDLCI & strPOSIPDLCI)
								Correct index pointers.
								Add 'Node Number'. 
								Retrieve Customer ID.
**************************************************************************************
-->
<%
dim strGWID, sql, intTCCount, lngCustID, lngAddrID
strGWID = Request("GWID")
lngCustID = Request("CustID")
lngAddrID = Request("AddrID")

'Response.Write ("GWID=" & strGWID)
	'Response.end
'get the tail circuit recordset
if isNumeric(strGWID)then
	dim rsTC
			sql = "SELECT DISTINCT TC.TAIL_CIRCUIT_ID, " &_
			"WAN_IP.IP_ADDRESS WAN_IP_ADDRESS, " &_
			"LAN_IP.IP_ADDRESS LAN_IP_ADDRESS, " &_		
			"TC.NODE_NAME, "&_
			"TC.TAIL_CIRCUIT_NUMBER, "&_
			"TC.WAN_IP_DLCI, "&_
			"TC.POS_IP_DLCI, "&_
					"NVL(AD.BUILDING_NAME,'<NO BUILDING SPECIFIED>') ||CHR(13)||CHR(10)|| " &_
					"decode(AD.APARTMENT_NUMBER, null, null, rtrim(AD.APARTMENT_NUMBER) || ' ') || " &_
					"decode(to_char(AD.HOUSE_NUMBER) || AD.HOUSE_NUMBER_SUFFIX, null, null, rtrim(to_char(AD.house_number) || AD.house_number_suffix)  || ' ') || " &_
					"decode(AD.STREET_VECTOR, null, null, rtrim(AD.STREET_VECTOR) || ' ') || " &_
					"NVL(AD.LONG_STREET_NAME,'NO STREET ADDRESS')||CHR(13)||CHR(10)|| " &_
					"NVL(AD.MUNICIPALITY_NAME,'NO MUNICIPALITY')||' '|| " &_
					"NVL(AD.PROVINCE_STATE_LCODE,'NO PROVINCE')||CHR(13)||CHR(10)|| " &_
					"NVL(AD.POSTAL_CODE_ZIP,'NO POSTAL CODE') ADDRESS, " &_
			"TC.ORDER_NUMBER, " &_
			"TC.NODE_NUMBER, " &_
			"TC.UPDATE_DATE_TIME, " &_
			"AD.MUNICIPALITY_NAME, AD.PROVINCE_STATE_LCODE, AD.COUNTRY_LCODE " &_
			"FROM " &_
			"CRP.RSAS_TAIL_CIRCUIT	TC, "&_
			"CRP.RSAS_IP_ADDRESS	WAN_IP, "&_
			"CRP.RSAS_IP_ADDRESS	LAN_IP, "&_				
			"CRP.ADDRESS			AD "&_	
			"WHERE " &_
		"TC.GATEWAY_ID= "& strGWID & " " &_
		"AND TC.WAN_IP_ID = WAN_IP.IP_ADDRESS_ID (+) "&_
		"AND TC.LAN_IP_ID = LAN_IP.IP_ADDRESS_ID (+) "&_						
		"AND TC.SITE_ADDRESS_ID = AD.ADDRESS_ID (+) "

	'Response.Write (sql)
	'Response.end

	set rsTC=server.CreateObject("ADODB.Recordset")
	rsTC.CursorLocation = adUseClient
	rsTC.Open sql, objConn
	if err then
		DisplayError "BACK", "", err.Number, "CANNOT CREATE RECORDSET", err.Description
	end if
	set rsTC.ActiveConnection = nothing
end if
%>
<LINK href="stylesheets/styles.css" rel="stylesheet" type="text/css">
<HTML>
<TITLE>In-line Frame Page</TITLE>
<STYLE>
.regularItem {
	cursor: hand;
}
.whiteItem {
	cursor: hand;
	background-color: white;
}
.Highlight {
	cursor: hand; 
	background-color: #00974f;	
	color: white;
}
</STYLE>

<script type="text/javascript">
var oldHighlightedElement;
var oldHighlightedElementClassName;

function cell_onClick(intTCID, intAddrID, strMunicipality, strProvince, strCountry, strLastUpdate){
	document.frmIFR.hdnTCID.value = intTCID;
	document.frmIFR.hdnAddressID.value = intAddrID;
	document.frmIFR.hdnMunicipality.value = strMunicipality;
	document.frmIFR.hdnProvince.value = strProvince;
	document.frmIFR.hdnCountry.value = strCountry;
	document.frmIFR.hdnLastUpdate.value = strLastUpdate;

	//highlight current record
	if (oldHighlightedElement != null) {oldHighlightedElement.className = oldHighlightedElementClassName}
	oldHighlightedElement = window.event.srcElement.parentElement;
	oldHighlightedElementClassName = oldHighlightedElement.className;
	oldHighlightedElement.className = "Highlight";
}

</script>

<body>
<form name="frmIFR" action="RSAST3GWTailCircuit.asp" method="POST">
<input type="hidden" name="hdnTCID" value="">
<input type="hidden" name="hdnLastUpdate" value="">
<input type="hidden" name="hdnAddressID" value="">
<input type="hidden" name="hdnMunicipality" value="">
<input type="hidden" name="hdnProvince" value="">
<input type="hidden" name="hdnCountry" value="">
<input type="hidden" name="hdnTCCount" value="<%if isNumeric(strGWID)then Response.write (rsTC.RecordCount) else Response.write (0)%>">
<TABLE border=0 cellspacing=0 frame=void cellpadding=0 width="100%">
	<thead>
	Count: <%if isNumeric(strGWID)then Response.write (rsTC.RecordCount) else Response.write (0)%>
 </thead>
 </TABLE>
<TABLE border=1 cellspacing=0 frame=void cellpadding=2 width="100%">
	<thead>
		<th nowrap>WAN IP Port Address</th>
		<th nowrap>LAN IP Port Address</th>
		<th nowrap>Node Name</th>
		<th nowrap>Node Number</th>
		<th nowrap>Tail Circuit Number</th>
		<th nowrap>WAN IP DLCI </th>
		<th nowrap>POS IP DLCI</th>
		<th nowrap>Address</th>
		<th nowrap>Order Number</th>
	</thead>
	<tbody>
		<%
		if isNumeric(strGWID) then
			dim k
			k = 0
			while not rsTC.EOF
				if Int(k/2) = k/2 then
					Response.Write "<tr class=""regularItem"">"
				else
					Response.Write "<tr class=""whiteItem"">"
				end if
				k = k+1
		%>
				<td nowrap onClick="cell_onClick(<%=rsTC(0)%>,'<%=lngAddrID%>','<%=rsTC(11)%>','<%=rsTC(12)%>','<%=rsTC(13)%>','<%=rsTC(10)%>')"><%=rsTC(1)%>&nbsp;</td>
				<td nowrap onClick="cell_onClick(<%=rsTC(0)%>,'<%=lngAddrID%>','<%=rsTC(11)%>','<%=rsTC(12)%>','<%=rsTC(13)%>','<%=rsTC(10)%>')"><%=rsTC(2)%>&nbsp;</td>
				<td nowrap onClick="cell_onClick(<%=rsTC(0)%>,'<%=lngAddrID%>','<%=rsTC(11)%>','<%=rsTC(12)%>','<%=rsTC(13)%>','<%=rsTC(10)%>')"><%=rsTC(3)%>&nbsp;</td>
				<td nowrap onClick="cell_onClick(<%=rsTC(0)%>,'<%=lngAddrID%>','<%=rsTC(11)%>','<%=rsTC(12)%>','<%=rsTC(13)%>','<%=rsTC(10)%>')"><%=rsTC(9)%>&nbsp;</td>
				<td nowrap onClick="cell_onClick(<%=rsTC(0)%>,'<%=lngAddrID%>','<%=rsTC(11)%>','<%=rsTC(12)%>','<%=rsTC(13)%>','<%=rsTC(10)%>')"><%=rsTC(4)%>&nbsp;</td>
				<td nowrap onClick="cell_onClick(<%=rsTC(0)%>,'<%=lngAddrID%>','<%=rsTC(11)%>','<%=rsTC(12)%>','<%=rsTC(13)%>','<%=rsTC(10)%>')"><%=rsTC(5)%>&nbsp;</td>
				<td nowrap onClick="cell_onClick(<%=rsTC(0)%>,'<%=lngAddrID%>','<%=rsTC(11)%>','<%=rsTC(12)%>','<%=rsTC(13)%>','<%=rsTC(10)%>')"><%=rsTC(6)%>&nbsp;</td>
				<td nowrap onClick="cell_onClick(<%=rsTC(0)%>,'<%=lngAddrID%>','<%=rsTC(11)%>','<%=rsTC(12)%>','<%=rsTC(13)%>','<%=rsTC(10)%>')"><%=rsTC(7)%>&nbsp;</td>
				<td nowrap onClick="cell_onClick(<%=rsTC(0)%>,'<%=lngAddrID%>','<%=rsTC(11)%>','<%=rsTC(12)%>','<%=rsTC(13)%>','<%=rsTC(10)%>')"><%=rsTC(8)%>&nbsp;</td>
			</tr>
		<%
			rsTC.MoveNext
			wend
			rsTC.Close
		end if
		%>
	</tbody>
</TABLE>

</FORM>
</BODY>
</HTML>
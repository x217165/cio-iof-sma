<%@ Language=VBScript %>
<%  
OPTION EXPLICIT
on error resume next
Response.CacheControl="Private"
%>
<% Response.Buffer = true %>
<!-- #include file="smaConstants.inc" -->
<!-- #include file="smaProcs.inc" -->
<!-- #include file="databaseconnect.asp" -->
<!--        
  ***************************************************************************************************
  * Name:		CustVpnList.asp
  * Purpose:	This page list the vpn list information.
  * Created By:	Anthony Cheung
  ***************************************************************************************************

        Date		Author			Changes/enhancements made
        -----		------		------------------------------------------------------
       24-June-13	ACheung		(Adpoted from CustServVpnList.asp)
       							Provide VPN information via the Netcracker customer web services.

  ***************************************************************************************************
-->
<%
Const ASP_NAME = "CustVpnList.asp" 'only need to change this value when changing the filename

dim strCustServID, strServTypeID
'dim objRsServiceOrderInfo, StrSql, objRs, objRsSTAtt, objRsSTAvalue, StrWhereClasuse, objRs2

dim strCustomerID
strCustomerID = Request("CustomerID")

if isnumeric(strCustomerID)  then
	
' Debug info
	'Response.Write "CID :" & strCustomerID & "<P>"				  

	'Getting VPN Name, Type, Topology
	Dim  vindex, vpn_record_count, vpnname(50), vpntype(50), vpntopology(50), vpn_map, vrflist(50),RDlist(50),cTvpnlist(50), cTcustomerName(50)
	Dim  strvpnws2Status, vpninfo(50,6)
'	set vpn_map = CreateObject("Scripting.Dictionary")

	strvpnws2Status = nc_getVPNListByCustomer(strCustomerID, vpn_record_count, vpnname, vpntype, vpntopology, vrflist, RDlist, cTvpnlist, cTcustomerName)
' Debug info
	'Response.write "<p>Status2 = " & strvpnws2Status & "</p>"	
	'Response.write "<p>Testing strCustServID = " & strCustServID & "</p>"
	'Response.write "<p>Post = " & Request.ServerVariables("REQUEST_METHOD") & "</p>"	
	'Response.write "<p>Size = " & vpn_record_count & "</p>"
	vindex = 0
	for vindex = 0 to vpn_record_count-1
' Debug info
'		Response.write "<p>Testing 1 strCustomerID  = " & strCustomerID & "</p>"
'		Response.write "<p>vindex = " & vindex & "</p>"
'		Response.write "<p>keyCID(vindex) = " & keyCID(vindex) & "</p>"					
'		Response.write "<p>vpnlist(vindex) = " & vpnname(vindex) & "</p>"
'		Response.write "<p>vpntypelist(vindex) = " & vpntype(vindex) & "</p>"					
'		Response.write "<p>vpntoptypelist(vindex) = " & vpntopology(vindex) & "</p>"
'		Response.write "<p>Before if keyCID(vindex) = " & keyCID(vindex) & "</p>"		
'				'vpn_map.add(k, vpnlist(vindex))
		vpninfo(vindex,0) = vpnname(vindex)
		vpninfo(vindex,1) = vpntype(vindex)
		vpninfo(vindex,2) = vpntopology(vindex)				
		vpninfo(vindex,3) = vrflist(vindex)
		vpninfo(vindex,4) = RDlist(vindex)
		vpninfo(vindex,5) = cTvpnlist(vindex)
		vpninfo(vindex,6) = cTcustomerName(vindex)
' Debug info
'		Response.write "<p>vpninfo(0) = " & vpninfo(0) & "</p>"								
'		Response.write "<p>vpninfo(1) = " & vpninfo(1) & "</p>"	
'		Response.write "<p>vpninfo(2) = " & vpninfo(2) & "</p>"
'		Response.write "<p>vpninfo(3) = " & vpninfo(3) & "</p>"
'		Response.write "<p>vpninfo(4) = " & vpninfo(4) & "</p>"
'		Response.write "<p>vpninfo(5) = " & vpninfo(5) & "</p>"
'		Response.write "<p>vpninfo(6) = " & vpninfo(6) & "</p>"		
	next									
				
	if response.isclientconnected = false then
		Response.End
	end if

end if
%>

<HTML>
<HEAD>
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<LINK href="stylesheets/styles.css" rel="stylesheet" type="text/css">
<TITLE>In-line Frame Customer VPN Page</TITLE>
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
</script>
</HEAD>
<BODY>
<form method=post name=frmIFRvpn action="<%=ASP_NAME%>">
<input type="hidden" name="hdnRefID" value="">
<input type="hidden" name="hdnServID" value="">
<input type="hidden" name="hdnUpdateDateTime" value="">
<input name="hdnExport" type=hidden value>

<TABLE border=7 cellspacing=0 cellpadding=2 width="100%">
	<thead>
		<th nowrap title="VPN Name">VPN Name</th>
		<th nowrap title="VPN Type">VPN Type</th>
		<th nowrap title="VPN Topology">VPN Topology</th>
		<th nowrap title="VRF">VRF</th>
		<th nowrap title="RD">RD</th>
		<th nowrap title="Target VPN (VRF, RD)">Target VPN</th>	
		<th nowrap title="Target NC Cutsomer Name">Target NC Customer Name</th>
	</thead>
	<tbody>
		<%
		dim k
		'Response.write "<p>strCustServID  = " & strCustServID & "</p>"
		'if strCustServID <> "" then
			'if Int(k/2) = k/2 then
			'	Response.Write "<tr class=""regularItem"">"
			'else
			'	Response.Write "<tr class=""whiteItem"">"
			'end if
			'k = k+1

			'for k = 0 to UBound(aList,2)
			for k = 0 to vpn_record_count-1			
			%>
			<tr> 
				<td nowrap><%=routineHTMLString(vpninfo(k,0))%>&nbsp;</td>
				<td nowrap><%=routineHTMLString(vpninfo(k,1))%>&nbsp;</td>
				<td nowrap><%=routineHTMLString(vpninfo(k,2))%>&nbsp;</td>
				<td nowrap><%=routineHTMLString(vpninfo(k,3))%>&nbsp;</td>
				<td nowrap><%=routineHTMLString(vpninfo(k,4))%>&nbsp;</td>
				<td nowrap><%=routineHTMLString(vpninfo(k,5))%>&nbsp;</td>
				<td nowrap><%=routineHTMLString(vpninfo(k,6))%>&nbsp;</td>
			</tr>
			<%
			next
			'objRs.MoveNext
			'next
			'objRs.Close
			'set objRs = Nothing
		'end if
		%>
		</tbody>
		<TFOOT>
		<TR>
		</TR>		
		</TFOOT>
</table>
</form>
</BODY>
</HTML>



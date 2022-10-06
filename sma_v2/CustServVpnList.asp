<%@ Language=VBScript %>
<%  
OPTION EXPLICIT
on error resume next
Response.CacheControl="Private"
%>
<% Response.Buffer = true %>
<!-- #include file="smaConstants.inc" -->
<!--#include file="sma_env.inc"-->
<!-- #include file="smaProcs.inc" -->
<!-- #include file="databaseconnect.asp" -->
<!--        
  ***************************************************************************************************
  * Name:		CustServVpnList.asp
  * Purpose:	This page list the vpn list information.
  * Created By:	Anthony Cheung
  ***************************************************************************************************

        Date		Author			Changes/enhancements made
        -----		------		------------------------------------------------------
       17-June-13	ACheung		(Adpoted from CorrSOWInstList.asp)
       							Provide VPN information via the Netcracker customer service web services.

  ***************************************************************************************************
-->
<%
Const ASP_NAME = "CustServVpnList.asp" 'only need to change this value when changing the filename

dim strCustServID, strServTypeID
dim objRsServiceOrderInfo, StrSql, objRs, objRsSTAtt, objRsSTAvalue, StrWhereClasuse, objRs2

strServTypeID = Request("ServiceTypeID")
strCustServID = Request("CustomerServiceID")

if isnumeric(strCustServID)  then
	
	'Response.Write "Service :" & strServTypeID & "<P>"	
	'Response.Write "CSID :" & strCustServID & "<P>"				  

	'Getting VPN Name, Type, Topology

	Dim  vindex, vpn_record_count, vpnname(5), vpntype(5), vpntopology(5), vpn_map, keyCSID(5), vrflist(5),RDlist(5), cTvpnlist(5), cTcustomerName(5)
	Dim  strvpnws2Status, vpninfo(50)
'		set vpn_map = CreateObject("Scripting.Dictionary")

	strvpnws2Status = nc_getVPNListByService(strCustServID, vpn_record_count, vpnname, vpntype, vpntopology, keyCSID, vrflist, RDlist, cTvpnlist, cTcustomerName)
'	Response.write "<p>Status2 = " & strvpnws2Status & "</p>"	
'	Response.write "<p>Testing strCustServID = " & strCustServID & "</p>"
'	Response.write "<p>Post = " & Request.ServerVariables("REQUEST_METHOD") & "</p>"	
'	Response.write "<p>Size = " & vpn_record_count & "</p>"
	vindex = 0
	for vindex = 0 to vpn_record_count-1
'		Response.write "<p>Testing 1 strCustServID = " & strCustServID & "</p>"
'		Response.write "<p>vindex = " & vindex & "</p>"
'		Response.write "<p>keyCSID(vindex) = " & keyCSID(vindex) & "</p>"					
'		Response.write "<p>vpnlist(vindex) = " & vpnname(vindex) & "</p>"
'		Response.write "<p>vpntypelist(vindex) = " & vpntype(vindex) & "</p>"					
''		Response.write "<p>vpntoptypelist(vindex) = " & vpntopology(vindex) & "</p>"
'		Response.write "<p>Before if keyCSID(vindex) = " & keyCSID(vindex) & "</p>"		
'		Response.write "<p>cTvpnlist(vindex) = " & cTvpnlist(vindex) & "</p>"					
'		Response.write "<p>cTcustomerName(vindex) = " & cTcustomerName(vindex) & "</p>"
		if keyCSID(vindex) = strCustServID then
''			vpn_map.add(k, vpnlist(vindex))
'			Response.write "<p>Inside if keyCSID(vindex) = " & keyCSID(vindex) & "</p>"	
			vpninfo(0) = vpnname(vindex)
'			Response.write "<p>vpninfo(0) = " & vpninfo(0) & "</p>"								
			vpninfo(1) = vpntype(vindex)
			vpninfo(2) = vpntopology(vindex)					
			vpninfo(3) = vrflist(vindex)
			vpninfo(4) = RDlist(vindex)
			vpninfo(5) = cTvpnlist(vindex)
			vpninfo(6) = cTcustomerName(vindex)
'			Response.write "<p>vpninfo(1) = " & vpninfo(1) & "</p>"	
'			Response.write "<p>vpninfo(2) = " & vpninfo(2) & "</p>"
'			Response.write "<p>vpninfo(3) = " & vpninfo(3) & "</p>"
'			Response.write "<p>vpninfo(4) = " & vpninfo(4) & "</p>"
'			Response.write "<p>vpninfo(5) = " & vpninfo(5) & "</p>"
'			Response.write "<p>vpninfo(6) = " & vpninfo(6) & "</p>"
		end if	
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
		for vindex = 0 to vpn_record_count-1
	      if keyCSID(vindex) = strCustServID then
	       		vpninfo(0) = vpnname(vindex)
				vpninfo(1) = vpntype(vindex)
				vpninfo(2) = vpntopology(vindex)					
				vpninfo(3) = vrflist(vindex)
				vpninfo(4) = RDlist(vindex)
				vpninfo(5) = cTvpnlist(vindex)
				vpninfo(6) = cTcustomerName(vindex)
				if vpninfo(0) <>"" then 
					response.write "<tr>"
					response.write "<td nowrap>" + vpninfo(0) + "&nbsp;</td>"
					response.write "<td nowrap>" + vpninfo(1) + "&nbsp;</td>"
					response.write "<td nowrap>" + vpninfo(2) + "&nbsp;</td>"
					response.write "<td nowrap>" + vpninfo(3) + "&nbsp;</td>"
					response.write "<td nowrap>" + vpninfo(4) + "&nbsp;</td>"
					response.write "<td nowrap>" + vpninfo(5) + "&nbsp;</td>"
					response.write "<td nowrap>" + vpninfo(6) + "&nbsp;</td>"
					response.write "<tr>"
				end if
			end if
		next
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



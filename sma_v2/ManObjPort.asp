<%@  language="VBSCRIPT" %>
<%
OPTION EXPLICIT
on error resume next
%>
<!-- #include file="smaConstants.inc" -->
<!-- #include file="smaProcs.inc" -->
<!-- #include file="databaseconnect.asp" -->
<!--
********************************************************************************************
* Page name:	ManObjPort.asp
*
* Purpose:		To display Managed Object Port Name and LAN IP.
*
* Created by:	Dan S. Ty	03/13/2002
*
********************************************************************************************
        Date		Author			Changes/enhancements made
        -----		------		------------------------------------------------------
	   15-Oct-03	DTy			Add field required for IP Mediation:
								Customer Service ID & Name & Billable Port.
***************************************************************************************************
-->
<%
'get strNE_ID - Network Element ID
dim strNE_ID
strNE_ID = Request("ne_id")
    
            
dim sql

'get the Port Name and LAN IP recordset
if strNE_ID <> "" then
	dim rsPort
	sql = "SELECT NE.NETWORK_ELEMENT_PORT_ID, NE.NETWORK_ELEMENT_PORT_NAME, NE.NETWORK_ELEMENT_PORT_IP,  " &_
          "  NE.BILLABLE_PORT, NE.CUSTOMER_SERVICE_ID, CS.CUSTOMER_SERVICE_DESC, NE.UPDATE_DATE_TIME, NE.REPORTABLE," &_
            "LCI.CTR_IN_NAME," &_
                 "LCO.CTR_OUT_NAME,"&_
                 "NE.VN_NAME," &_
            "LVI.VTR_IN_NAME," &_
          "LVO.VTR_OUT_NAME," &_
          "NE.QOS_NAME," &_
          "LEI.ETR_IN_NAME," &_
          "LEO.ETR_OUT_NAME," &_
          "LCS.CI_STATUS_NAME," &_
          "CO.ORGANIZATION_NAME," &_
          "CO.ORGANIZATION_CODE," &_
          "SNC.SITE_NAME,"   &_
           "SNC.SITE_CODE,"   &_
           "NPNA.NETWORK_PORT_NAME_ALIAS NAME_ALIAS," &_
           "NE.PORT_NAME_ALIAS, " &_
           "( select listagg(S.MGMT_SYSTEM_NAME,',')   within group (order by S.MGMT_SYSTEM_ID)  MGMT_SYSTEMS " &_
    " from CRP.NETWORK_ELEMENT_MGMT_SYS NEMS, CRP.LCODE_MGMT_SYSTEMS S " &_
  "where NEMS.MGMT_SYSTEM_ID = S.MGMT_SYSTEM_ID and  NEMS.NETWORK_ELEMENT_Port_ID = NE.NETWORK_ELEMENT_Port_ID) as MGMT_VALUE, "&_
            "NE.MSUID,"   &_
    "LNP.NE_PORT_FUNCTION_NAME , NE.RECORD_STATUS_IND"   &_
          "  FROM CRP.NETWORK_ELEMENT_PORT NE, CRP.CUSTOMER_SERVICE CS,CRP.LCODE_NE_PORT_FUNCTION LNP, CRP.CUSTOMER_ORGANIZATION CO, CRP.SITE_NAME_CODE SNC,CRP.LCODE_CTR_IN LCI,CRP.LCODE_CTR_OUT LCO,CRP.NETWORK_PORT_NAME_ALIAS NPNA,CRP.LCODE_ETR_IN LEI,CRP.LCODE_ETR_OUT LEO,CRP.LCODE_CI_STATUS LCS,CRP.LCODE_VTR_IN LVI,CRP.LCODE_VTR_OUT LVO" &_
          "  WHERE NE.CUSTOMER_SERVICE_ID = CS.CUSTOMER_SERVICE_ID (+) AND NETWORK_ELEMENT_ID = " & strNE_ID &_
		  " and NE.ORGANIZATION_ID = CO.ORGANIZATION_ID(+) and NE.SITE_ID =SNC.SITE_ID(+) and NE.ETR_IN_ID = LEI.ETR_IN_ID(+) and NE.ETR_OUT_ID = LEO.ETR_OUT_ID(+) and NE.CTR_IN_ID=LCI.CTR_IN_ID(+) and NE.CTR_OUT_ID= LCO.CTR_OUT_ID(+) and NE.VTR_IN_ID=LVI.VTR_IN_ID(+) and NE.VTR_OUT_ID= LVO.VTR_OUT_ID(+) and NE.CI_STATUS_ID= LCS.CI_STATUS_ID(+) and NE.NE_PORT_FUNCTION_LCODE= LNP.NE_PORT_FUNCTION_LCODE(+)  and NE.network_element_port_id = NPNA.network_element_port_id(+) ORDER BY NE.NETWORK_ELEMENT_PORT_ID, NE.NETWORK_ELEMENT_PORT_IP"
    
           
	set rsPort=server.CreateObject("ADODB.Recordset")
	rsPort.CursorLocation = adUseClient
	rsPort.Open sql, objConn
	if err then
		DisplayError "BACK", "", err.Number, "CANNOT OPEN RECORDSET", err.Description
	end if
	set rsPort.ActiveConnection = nothing
end if
%>
<html>
<link href="stylesheets/styles.css" rel="stylesheet" type="text/css">
<meta http-equiv="Pragma" content="no-cache">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>In-line Frame Page</title>
<style>
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
</style>

<script type="text/javascript">
    var oldHighlightedElement;
    var oldHighlightedElementClassName;

    function cell_onClick(intPortID, strLastUpdate, record_ind) {
        
        document.frmIFR2.hdnPortID.value = intPortID;
        document.frmIFR2.hdnLastUpdate.value = strLastUpdate;

        window.parent.document.getElementsByName("btn_iFrame2Delete")[0].disabled = false;

        if (record_ind == 'D') {
            window.parent.document.getElementsByName("btn_iFrame2Delete")[0].value = "UnDelete";
        }
        else {
            window.parent.document.getElementsByName("btn_iFrame2Delete")[0].value = "Delete";
        }
        //highlight current record
        if (oldHighlightedElement != null) {
            oldHighlightedElement.className = oldHighlightedElementClassName
        }
        oldHighlightedElement = window.event.srcElement.parentElement;
        oldHighlightedElementClassName = oldHighlightedElement.className;
        oldHighlightedElement.className = "Highlight";
    }

</script>

<body>
    <form name="frmIFR2" action="ManObjPort.asp" method="POST">
        <input type="hidden" name="hdnPortID" value="">
        <input type="hidden" name="hdnLastUpdate" value="">

        <table border="1" cellspacing="0" cellpadding="2" width="100%">
            <thead>
                <tr>
                    <th style="align-content: center;" colspan="25" title="Port Information">Port Information</th>
                </tr>
                <tr>
                    <th nowrap title="Port Name">Port Name</th>
                    <th nowrap title="LAN IP">LAN IP</th>
                    <th nowrap title="Billable Port">Billable?</th>
                    <th nowrap title="Reporting Required">Reporting?</th>
                    <th nowrap title="CS ID">CS ID</th>
                    <th nowrap title="Customer Service Name">Customer Service Name</th>
                    <th nowrap title="CTR_IN_ID ">CTR IN</th>
                    <th nowrap title="CTR_OUT_ID">CTR OUT</th>
                    <th nowrap title="VN_NAME">VN NAME</th>
                    <th nowrap title="VTR_IN_ID">VTR IN</th>
                    <th nowrap title="VTR_OUT_ID">VTR OUT</th>
                    <th nowrap title="QOS_NAME">QOS NAME</th>
                    <th nowrap title="ETR_IN_ID">ETR IN</th>
                    <th nowrap title="ETR_OUT_ID">ETR OUT</th>
                    <th nowrap title="CI_STATUS">CI_STATUS</th>
                    <th nowrap title="ORGANIZATION_Name">ORGANIZATION NAME</th>
                    <th nowrap title="ORGANZATION_CODE">ORGANIZATION CODE</th>
                    <th nowrap title="SITE_NAME">SITE NAME</th>
                    <th nowrap title="SITE_CODE">SITE CODE</th>
                    <th nowrap title="SITE_CODE">NAME ALIAS</th>
                    <th nowrap title="SITE_CODE">MGMT SYSTEM</th>
                    <th nowrap title="PORT_IDENTIFICATION">PORT FUNCTION</th>
                    <th nowrap title="PORT_FUNCTION">PORT IDENTIFICATION</th>
                     <th nowrap title="PORT_FUNCTION">RECORD STATUS INDICATOR</th>
                     <th nowrap title="PORT_NAME_ALIAS">PORT NAME ALIAS</th>
                </tr>
            </thead>
            <tbody>
                <%
		if strNE_ID <> "" then
		dim k
		k = 0
		while not rsPort.EOF
			if Int(k/2) = k/2 then
				Response.Write "<tr class=""regularItem"">"
			else
				Response.Write "<tr class=""whiteItem"">"
			end if
			k = k+1
                %>
                <td nowrap onclick="cell_onClick(<%=rsPort(0)%>, '<%=rsPort(6)%>' , '<%=rsPort("RECORD_STATUS_IND")%>')"><%=rsPort(1)%>&nbsp;</td>
                <td nowrap onclick="cell_onClick(<%=rsPort(0)%>, '<%=rsPort(6)%>', '<%=rsPort("RECORD_STATUS_IND")%>')"><%=rsPort(2)%>&nbsp;</td>
                <td nowrap onclick="cell_onClick(<%=rsPort(0)%>, '<%=rsPort(6)%>', '<%=rsPort("RECORD_STATUS_IND")%>')"><%=rsPort(3)%>&nbsp;</td>
                <td nowrap onclick="cell_onClick(<%=rsPort(0)%>, '<%=rsPort(6)%>', '<%=rsPort("RECORD_STATUS_IND")%>')"><%=rsPort(7)%>&nbsp;</td>
                <td nowrap onclick="cell_onClick(<%=rsPort(0)%>, '<%=rsPort(6)%>', '<%=rsPort("RECORD_STATUS_IND")%>')"><%=rsPort(4)%>&nbsp;</td>
                <td nowrap onclick="cell_onClick(<%=rsPort(0)%>, '<%=rsPort(6)%>', '<%=rsPort("RECORD_STATUS_IND")%>')"><%=rsPort(5)%>&nbsp;</td>
                <td nowrap onclick="cell_onClick(<%=rsPort(0)%>, '<%=rsPort(6)%>', '<%=rsPort("RECORD_STATUS_IND")%>')"><%=rsPort("CTR_IN_NAME")%>&nbsp;</td>
                <td nowrap onclick="cell_onClick(<%=rsPort(0)%>, '<%=rsPort(6)%>', '<%=rsPort("RECORD_STATUS_IND")%>')"><%=rsPort("CTR_OUT_NAME")%>&nbsp;</td>
                <td nowrap onclick="cell_onClick(<%=rsPort(0)%>, '<%=rsPort(6)%>', '<%=rsPort("RECORD_STATUS_IND")%>')"><%=rsPort("VN_NAME")%>&nbsp;</td>
                <td nowrap onclick="cell_onClick(<%=rsPort(0)%>, '<%=rsPort(6)%>', '<%=rsPort("RECORD_STATUS_IND")%>')"><%=rsPort("VTR_IN_NAME")%>&nbsp;</td>
                <td nowrap onclick="cell_onClick(<%=rsPort(0)%>, '<%=rsPort(6)%>', '<%=rsPort("RECORD_STATUS_IND")%>')"><%=rsPort("VTR_OUT_NAME")%>&nbsp;</td>
                <td nowrap onclick="cell_onClick(<%=rsPort(0)%>, '<%=rsPort(6)%>', '<%=rsPort("RECORD_STATUS_IND")%>')"><%=rsPort("QOS_NAME")%>&nbsp;</td>
                <td nowrap onclick="cell_onClick(<%=rsPort(0)%>, '<%=rsPort(6)%>', '<%=rsPort("RECORD_STATUS_IND")%>')"><%=rsPort("ETR_IN_NAME")%>&nbsp;</td>
                <td nowrap onclick="cell_onClick(<%=rsPort(0)%>, '<%=rsPort(6)%>', '<%=rsPort("RECORD_STATUS_IND")%>')"><%=rsPort("ETR_OUT_NAME")%>&nbsp;</td>
                <td nowrap onclick="cell_onClick(<%=rsPort(0)%>, '<%=rsPort(6)%>', '<%=rsPort("RECORD_STATUS_IND")%>')"><%=rsPort("CI_STATUS_NAME")%>&nbsp;</td>
                <td nowrap onclick="cell_onClick(<%=rsPort(0)%>, '<%=rsPort(6)%>', '<%=rsPort("RECORD_STATUS_IND")%>')"><%=rsPort("ORGANIZATION_NAME")%>&nbsp;</td>
                <td nowrap onclick="cell_onClick(<%=rsPort(0)%>, '<%=rsPort(6)%>', '<%=rsPort("RECORD_STATUS_IND")%>')"><%=rsPort("ORGANIZATION_CODE")%>&nbsp;</td>
                <td nowrap onclick="cell_onClick(<%=rsPort(0)%>, '<%=rsPort(6)%>', '<%=rsPort("RECORD_STATUS_IND")%>')"><%=rsPort("SITE_NAME")%>&nbsp;</td>
                <td nowrap onclick="cell_onClick(<%=rsPort(0)%>, '<%=rsPort(6)%>', '<%=rsPort("RECORD_STATUS_IND")%>')"><%=rsPort("SITE_CODE")%>&nbsp;</td>
                <td nowrap onclick="cell_onClick(<%=rsPort(0)%>, '<%=rsPort(6)%>', '<%=rsPort("RECORD_STATUS_IND")%>')"><%=rsPort("NAME_ALIAS")%>&nbsp;</td>
                <td nowrap onclick="cell_onClick(<%=rsPort(0)%>, '<%=rsPort(6)%>', '<%=rsPort("RECORD_STATUS_IND")%>')"><%=rsPort("MGMT_VALUE")%>&nbsp;</td>
                <td nowrap onclick="cell_onClick(<%=rsPort(0)%>, '<%=rsPort(6)%>', '<%=rsPort("RECORD_STATUS_IND")%>')"><%=rsPort("NE_PORT_FUNCTION_NAME")%>&nbsp;</td>
                <td nowrap onclick="cell_onClick(<%=rsPort(0)%>, '<%=rsPort(6)%>', '<%=rsPort("RECORD_STATUS_IND")%>')"><%=rsPort("MSUID")%>&nbsp;</td>
                <td nowrap onclick="cell_onClick(<%=rsPort(0)%>, '<%=rsPort(6)%>', '<%=rsPort("RECORD_STATUS_IND")%>')"><%=rsPort("RECORD_STATUS_IND")%>&nbsp;</td>
                <td nowrap onclick="cell_onClick(<%=rsPort(0)%>, '<%=rsPort(6)%>', '<%=rsPort("RECORD_STATUS_IND")%>')"><%=rsPort("PORT_NAME_ALIAS")%>&nbsp;</td>
                </tr>
		<%
		rsPort.MoveNext
		wend
		rsPort.Close
		end if
        %>
            </tbody>
        </table>

    </form>
</body>
</html>

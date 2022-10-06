<%
const strConstConnectString = "DRIVER={Microsoft ODBC for Oracle};UID=sma_userid;Password=sma;server=ESDAT"	
const strConstSConnectString = "DRIVER={Microsoft ODBC for Oracle};UID=crp_sec;Password=gcnet_crp;server=ESDSAT"   
const CRP = "DRIVER={Microsoft ODBC for Oracle};UID=crp;Password=gcnet_crp;server=ESDSAT" 
const GCNET = "DRIVER={Microsoft ODBC for Oracle};UID=crp;Password=gcnet_crp;server=ESDSAT" 

public function getConnString(str)
dim connstring
connstring=""
Select case str
	case "strConstConnectString"
		connstring=strConstConnectString 
	case "strConstSConnectString"
		connstring=strConstSConnectString
	case "GCNET"
		connstring=GCNET
	case "CRP"
		connstring=CRP
End Select
getConnString = connstring
end function
%>

<%@ Language=VBScript %>
<%Option Explicit%>
<!-- #include file = "smaConstants.inc" 
**************************************************************************************

        Date		Author			Changes/enhancements made
        -----		------		------------------------------------------------------
       31-Dec-01	  DTy		Add Case "RSAST3Addr".
								Add Case "RSAST3Admin".
       22-Feb-02	  DTy		Adjust pop-up window sizes to remove vertical scroll
                                   bars.
       29-Mar-02	  DTy		Add Case "ContactClean", "CustClean" & "XLS".
       24-May-13	ACheung		Add Customer Profile info on Customer and Customer Service Search.
**************************************************************************************
-->

<%
'check if the user is logged on
'If Request.Cookies("UserAccessLevel")(strConst_Logon) = "" Then
If Session(strConst_Logon) = "" Then
	Response.Redirect "default.asp?redir=Y"
End If

Dim StrAppOption,StrSearchPg, intSearchFrameRows
'intSearchFrameRows = height of search frame

 StrAppOption = Request.QueryString("fraSrc")
 
 Select Case StrAppOption
  Case "Address"
	StrSearchPg="AddressCriteria.asp"
	intSearchFrameRows = 210
  Case "Asset"
	strSearchPg="AssetCriteria.asp"
	intSearchFrameRows = 225
  Case "AssetCatalogue"
	strSearchPg="AssetCatCriteria.asp"
	intSearchFrameRows = 180
  Case "AssetClass"
	strSearchPg="AssetClassCriteria.asp"
	intSearchFrameRows = 180
  Case "AssetSubclass"
	strSearchPg="AssetSubClassCriteria.asp"
	intSearchFrameRows = 165
  Case "AssetType"
	strSearchPg="AssetTypeCriteria.asp"
	intSearchFrameRows = 220	
  Case "City"
	StrSearchPg="CityCriteria.asp"
	intSearchFrameRows = 180
  Case "Contact"
	StrSearchPg="ContactCriteria.asp" 
	intSearchFrameRows = 200
  Case "ContactClean"
	StrSearchPg="ContactCleanEntry.asp" 
	intSearchFrameRows = 200
  Case "ContactRole"
	StrSearchPg="ContactRoleCriteria.asp"
	intSearchFrameRows = 220
  Case "Correlation"   
    	StrSearchPg="CorrCriteria.asp" 
	intSearchFrameRows = 240  
  Case "Country"
	StrSearchPg="CountryCriteria.asp"
	intSearchFrameRows = 180
  Case "Cust" 
	StrSearchPg="CustCriteria.asp"
	intSearchFrameRows = 220
  Case "Cust_CP" 
	StrSearchPg="CustCPCriteria.asp"
	intSearchFrameRows = 240
  Case "Cust_Profile" 
	StrSearchPg="CPCriteria.asp"
	intSearchFrameRows = 150
  Case "CustClean" 
	StrSearchPg="CustCleanEntry.asp"
	intSearchFrameRows = 220
  Case "CustServ"
	StrSearchPg="CustServCriteria.asp"
	intSearchFrameRows = 380
  Case "FCustServ"
	StrSearchPg="CustServfCriteria.asp"
	intSearchFrameRows = 380
  Case "CustServ_CP"
	StrSearchPg="CustServCPCriteria.asp"
	intSearchFrameRows = 390
  Case "CustServCont"
    StrSearchPg="CustServContCriteria.asp" 
	intSearchFrameRows = 210
  Case "CustServPVC"
	StrSearchPg="CustServPVCCriteria.asp"
	intSearchFrameRows = 260
  Case "Facility"
    StrSearchPg="FacilityCriteria.asp?CktType=OTHER" 
	intSearchFrameRows = 300
  Case "FacilityPVC"
    StrSearchPg="FacilityCriteria.asp?CktType=PVC" 
	intSearchFrameRows = 255
  Case "Geocode"
	StrSearchPg="GeocodeCriteria.asp"
	intSearchFrameRows = 210
  Case "Holiday"
	StrSearchPg="HolidayCriteria.asp"
	intSearchFrameRows = 140
  Case "LOB"
	StrSearchPg = "LOBCriteria.asp"
	intSearchFrameRows = 140
  Case "Make"
	StrSearchPg="MakeCriteria.asp"
	intSearchFrameRows = 120
  Case "Model"
	StrSearchPg="ModelCriteria.asp"
	intSearchFrameRows = 120	
  Case "ManagedObjects"
    StrSearchPg="manobjsearch.asp" 
	intSearchFrameRows = 310
  Case "Municipalities"
	StrSearchPg="CityCriteria.asp"
	intSearchFrameRows = 180
  Case "PartNum"
	StrSearchPg="PartNumCriteria.asp"
	intSearchFrameRows = 120
  Case "Province"
	StrSearchPg="ProvinceCriteria.asp"
	intSearchFrameRows = 180
  Case "OrderHistory"
	StrSearchPg="OrderHistoryCriteria.asp"
	intSearchFrameRows = 150
  Case "RSAST3Admin"
    StrSearchPg="RSAST3IPGenCriteria.asp" 
	intSearchFrameRows = 260
  Case "RSAST3Tier3"
    StrSearchPg="RSAST3search.asp" 
	intSearchFrameRows = 260
  Case "RSAST3Addr"
	StrSearchPg="RSAST3AddrCriteria.asp"
	intSearchFrameRows = 260
  Case "RSAST3AddrDetail" 
	StrSearchPg="RSAST3AddrDetail.asp"
	intSearchFrameRows = 200
  Case "ServiceCategory"
	StrSearchPg = "SCategoryCriteria.asp"
	intSearchFrameRows = 210
  Case "ServiceType"
	StrSearchPg = "STypeCriteria.asp"
	intSearchFrameRows = 210
  Case "ServiceTypeA"
	StrSearchPg = "STypeACriteria.asp"
	intSearchFrameRows = 300
  Case "ServiceInstA"
	StrSearchPg = "SInstACriteria.asp"
	intSearchFrameRows = 210
  Case "ServiceTypeKA"
	StrSearchPg = "SKenanACriteria.asp"
	intSearchFrameRows = 210
  Case "SLA"
	StrSearchPg = "SLACriteria.asp"
	intSearchFrameRows = 210
  Case "ServLoc"
	StrSearchPg="ServLocCriteria.asp"
	intSearchFrameRows = 230
  Case "Schedule"
	StrSearchPg="ScheduleCriteria.asp"
	intSearchFrameRows = 210
  Case "Staff"
	StrSearchPg="StaffCriteria.asp" 
	intSearchFrameRows = 230  	
  Case "StaffRole"
	StrSearchPg="StaffRoleCriteria.asp" 
	intSearchFrameRows = 170  	

  Case "XLSEntry"
	StrSearchPg="XLSEntry.asp"
	intSearchFrameRows = 180

  End Select 
  %>
<html>
<head>
<title>Service Management Application</title>
<frameset border="0" rows="<%=intSearchFrameRows%>,*" id="frasetSearch" name="frasetSearch">
  <frame scrolling="AUTO" src="<%=StrSearchPg%>" id="fraCriteria" name="fraCriteria">
  <frame scrolling="AUTO" src="blank.htm" id="fraResult" name="fraResult">
</frameset>
</head>
<body></body>
</html>

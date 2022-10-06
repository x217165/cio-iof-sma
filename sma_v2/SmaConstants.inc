<%
'database related
const intRowsReturned = 501                                                     'number of rows returned by a query

'const strConstConnectString = "DRIVER={Microsoft ODBC for Oracle};UID=sma_userid;Password=sma;server=ESDAT"		'connection string
'const strConstConnectString = "DRIVER={Oracle in OraClient11g_home};UID=SMA_USERID;PWD=SMA;Data Source=esdat"
'const strConstConnectString = "DRIVER={Oracle in OraClient11g_home};Dbq=rac_ln98090-scan:41521/ESDATsva;UID=SMA_USERID;PWD=SMA"

const intConstDisplayPageSize = 10                                              'the number of records displayed in a html page at a time (page size)
const strDelimiter = "¿"                                                        'delimiter used in complex strings
const strTelus = "TELUS"                                                        'used for hard coding of TELUS in queries

const intConst_CustShortNameLength = 5                                          'number of characters allowed for customer short name (database field = 15)
const intConst_MessageDisplay = 5000                                            'period of time status bar displays message before cleared
const intBaseYear = 1990                                                        'Base year for generating pulldown year select

const strConst_ServiceTypeEmailTo = "Mark Kuziej <mark.kuziej@telus.com>; Kevin Flemming <kevin.flemming@telus.com>; Steve Mullen <steve.mullen@telus.com>; Phil Brien <phil.brien@telus.com>"
'const strConst_ServiceTypeEmailTo = "Winnie Yue <winnie.yue@telus.com>; Gilles Archer <gilles.archer@telus.com>"

const strConst_NamePrefix = "Miss¿Mrs.¿Ms.¿Mr.¿Dr.¿Capt.¿Prof."

'Business Function Name/ID
const strConst_Logon = "Access_Granted"
const strConst_Customer = "100"
const strConst_CustomerName = "101"
const strConst_CustomerNameAlias = "102"
const strConst_Address = "110"
const strConst_ServiceLocation = "120"
const strConst_ServiceLocationContact = "121"
const strConst_Contact = "130"
const strConst_ContactRole = "131"
const strConst_CustomerService = "140"
const strConst_CustomerServiceContact = "141"
const strConst_CustomerServiceEmailSetup = "142"
const strConst_Facilities = "150"
const strConst_PVC = "160"
const strConst_Asset = "170"
const strConst_AssetAdditionalCosts = "171"
const strConst_AssetCatalogue = "172"
const strConst_AssetTypeClassification = "173"
const strConst_ManagedObjects = "180"
const strConst_SNMP = "1801"
const strConst_SNMP_write = "1802"
const strConst_COM = "1001"
const strConst_COM_write = "1002"
const strConst_MO_LCODE="4"
const strConst_CorrelationCustomer = "190"
const strConst_CorrelationElements = "191"
const strConst_Security = "300"
const strConst_ServiceCatalogue = "200"
const strConst_CityProvinceCountry = "220"
const strConst_RSAS = "210454"
const strConst_ESDCleanup = "220000"
const strConst_HolidayDefinition = "SMA - Holiday Definition"
const strConst_LOB = "SMA - Line of Business"
const strConst_ScheduleDefinition = "SMA - Schedule Definition"
const strConst_ServiceType = "SMA - Service Type"
const strConst_ServiceDefinition = "SMA - Service Definition"
const strConst_ServiceLevelAgreement = "SMA - Service Level Agreements"
const strConst_StaffManagement = "SMA - Staff Management"
const strConst_UsageCalculation = "SMA - Usage Calculation"
'please see Winnie before adding/changing these codes

'Access Levels
const intConst_Access_ReadOnly = 1
const intConst_Access_Create = 2
const intConst_Access_Update = 4
const intConst_Access_Delete = 8

dim arrRegionXRef(2, 5)
arrRegionXRef(0,0) = "AB"
arrRegionXRef(1,0) = "Alberta"
arrRegionXRef(0,1) = "BC"
arrRegionXRef(1,1) = "BC"
arrRegionXRef(0,2) = "EAST"
arrRegionXRef(1,2) = "National"
arrRegionXRef(0,3) = "QUEBEC"
arrRegionXRef(1,3) = "National"
arrRegionXRef(0,4) = "EAST"
arrRegionXRef(1,4) = "US / Int'l"


Const propLDAPPeoplePath = "ou=people,ou=teamMembers,ou=internal,o=telus"
Const propLDAPApplicationsPath = "ou=people,ou=systems,ou=internal,o=telus"

Const SXH_OPTION_IGNORE_SERVER_SSL_CERT_ERROR_FLAGS = 2
Const SXH_SERVER_CERT_IGNORE_ALL_SERVER_ERRORS = 13056
Const strcpwsid = "APP_SMA2"
Const strcppw = "soaorgid"


Const int_max_csidpervpn = 300	


Const strPermItwsid = "APP_SMA2"
Const strPermItpw = "soaorgid"
Const SMA2 = "sma2"

%>

<%
const strConstConnectString = "MIH3BgkrBgEEAYI3WAOggekwgeYGCisGAQQBgjdYAwGggdcwgdQCAwIAAQICZhACAgEABBAAAAAAAAAAAAAAAAAAAAAABBCBttxNmdmgvSXJGCgbK9AUBIGgQcpMat0OdP4nPBS+m3z7TDuiOoYZKUeKciI0RESQ/kAzqyZv8lTHow2ko4WkLxh61Xve2COmbcI0F4fpIulHg80A4HQfZGyTdVw8z3oU/UT2Yeky94kHCsAt/QNY3ubbSUrWueDbHQMuna1l6XwnUtINRrUSn4uiX935Rjmo4Q+9cJr7OnS0gaXukhIEHK6PHaSxmtIPezwVcWt/MNXGHA=="
const strConstSConnectString ="MIH3BgkrBgEEAYI3WAOggekwgeYGCisGAQQBgjdYAwGggdcwgdQCAwIAAQICZhACAgEABBAAAAAAAAAAAAAAAAAAAAAABBCCxWQkCnl/bQsm63naDDyzBIGg4FQDbzajis9iTKOGwyvt7u4gm6TBlkxEr8jpHWLZ+HFgFJn1EbQ3YnNzF5FMf7Axh6FHzN15Wj0jVTn+XIhby88+YLDlYEsVjIuna/q7/AIg9QzX9NZsfvXMh3vaVVMRmPheMAHsoMKSfqvYN1a6CavcvzYDymnsZdS2kW6WZObLbebC76d2IPNrajYEu+OCVTaExJGNiJEtxgNfP+Twpg=="
const APP_SMA_ALL_LCD="MIIBBwYJKwYBBAGCN1gDoIH5MIH2BgorBgEEAYI3WAMBoIHnMIHkAgMCAAECAmYQAgIBAAQQAAAAAAAAAAAAAAAAAAAAAAQQ8Yd7TXGQwmP6Uwq2JC6giQSBsGzUgGcvlk6DAQAj1g1rt3rzwwUCQT6ZeklVez/MSP1oKgzNt69nYjKFjno/jqhqHuz0mqi7BAwxuI92YsBjGip1XqY+/UveaDlNKY2Fjsj131lDgN39iBdS5WkoE1XE5ewP3Ht9XYHrTIRDx91Ug/pbEOTyLpxdEW9R7PqwzOzS9oWmp7DcV/89KXXHFDSlTAOF2JgWt3mP1fDLjkRDukadccAHB5Bec4fwhGJtGvmJ"
const APP_SMA_GCNET="MIH3BgkrBgEEAYI3WAOggekwgeYGCisGAQQBgjdYAwGggdcwgdQCAwIAAQICZhACAgEABBAAAAAAAAAAAAAAAAAAAAAABBBRBkf3y6SmPrRjHkK+wmEXBIGgdbM/ZX4yzIfPFgxmzL3zfopZRqMEKqOaQTlQCVbS3sx2AH7AmG0SLimg1hBkz3/0e5GFnR54Vy0gkF7yv2uLn5DVaK7vqztOufs0oswaGOuk6CngcxZgw0mSA5BDDb3kHIt63wYDZOTDUjSQn9hGQjVBWBekGUWAWkS8my7eiCwdb0XNRa2OmY/MeVCv5LQd55t7q6ya/WMqAoVUSjAUSw=="
const APP_SMA_OPSNAS="MIIBBwYJKwYBBAGCN1gDoIH5MIH2BgorBgEEAYI3WAMBoIHnMIHkAgMCAAECAmYQAgIBAAQQAAAAAAAAAAAAAAAAAAAAAAQQCWzkyioHTk1DvG1tf4K9ngSBsJVV3qLgYwgirsKXQLh0oH+T4+9vud/7KKqmAswnYAMgwsWJbHeM6u8SfYube/N2EXrlsRr7rtgD9bdMTMsO4Kr3RuFp/eOwaeeeUIPCOCkhdZZuqTBseYzQpqwY//9zJQ4eBmjxgRgl6AhEMXEPjD9zqqI51vaxxbILiUENJx414mSIy3kpkvljDtvKzU6w6i12yhKVpzBx9EF0JkyW59Y2q7L8LHvGjZTYRzH0CZou"
const APP_SMA_GOCWCS="MIIBBwYJKwYBBAGCN1gDoIH5MIH2BgorBgEEAYI3WAMBoIHnMIHkAgMCAAECAmYQAgIBAAQQAAAAAAAAAAAAAAAAAAAAAAQQQVYz6ZWVELjQxU4fwlGXxQSBsLcGqAdwzs+KCM0Opwmk7D4octNpueRpj71z3zJnZs6DeEgoD7Pr17+CytoEAS1C8tSVu1OS8jloaOOHXrHCROTrbqpLH+ojwBy7gCvEn+7v2mfpRyo/gwHaSHyOIG6YO3anvW3NWWbllD5sFriN3vnlfOnplDF18bv+zsPyWOTYXKYT7UJ8rKzSR5HgDTn+wv6M317uki2pWolRRCEKJi/0akysLsaIY242rEEMGpjM"
const APP_SMA_GCGONET="MIIBBwYJKwYBBAGCN1gDoIH5MIH2BgorBgEEAYI3WAMBoIHnMIHkAgMCAAECAmYQAgIBAAQQAAAAAAAAAAAAAAAAAAAAAAQQMEBbUkWQ+hh/f9utkzk3CwSBsKfVpoKJ+NWTAup6jJX0m6NmtrziMq0yP9XpUU+YLI8teIYyf4yKiJdk0eV2o1h6Bu47arn3Lal0BPZ8k3JnFy7vQnEcRMO2RF5Jjm6JeGLv9fuj11WWIOgFiVmt2NAxslrVh5VRfr8IP4xByty6KpiM4Gf+drnDiDNvrl8h6ytIO9Y5LYqOnLqESPSMG1N2g/PPTobyVJCWjhOAelllutwV4shVhVfmkqQ+a1zx4ghs"
const APP_SMA_WCSGONET="MIIBGQYJKwYBBAGCN1gDoIIBCjCCAQYGCisGAQQBgjdYAwGggfcwgfQCAwIAAQICZhACAgEABBAAAAAAAAAAAAAAAAAAAAAABBD77fcBjR7zqL8LfYdED3AKBIHAaAl0ZYNlJ6LUbhfpHoTxhDrUbwV/4SJ2IMgEKLKrBzzed95VdaMATr9W5UbIETzbZJqylU2UxrVDFovdf3RjKAQ5Iil0NbpoFhpzoOlGchsyxgzv9gpSFlyB7PEOpORUq0XmlC2FYg1Z7Wgu0Av0wh+PHjyNbdCvDoZRadeRYYv6XBX6EfXK9DJfsy90h0sh6udJ4a9yxY+lN6ROHKL4feu+DLkNUVjfa/IOHzD4G/Hh2cPRNCOq3vDRg15xmJLS"
const APP_SMA_WCSGCNET="MIIBGQYJKwYBBAGCN1gDoIIBCjCCAQYGCisGAQQBgjdYAwGggfcwgfQCAwIAAQICZhACAgEABBAAAAAAAAAAAAAAAAAAAAAABBAyuIoSf6JARcsFWrDStcrBBIHAVPqdr11UlpVY8K/ODEmIChRkTa2VQqz2Wd7qrsFDkLdGdRore+chW7EkMf5AL2X2Q5QR5xLYCZEHLZfzT0MO3SSezSl6LguvW3R9E9aHqgD7wLxJcrzlltITpHaTBJ6VIGz0ps4i5SNiEtQ8FhNUfVO4BCnyN0Npl6AG7orfJT5RpQnLM5AvLL6uYTOBvr++2vi33zmSCbTcpxHHy7lN6erpBwVJ/aSUL9Jy7VdcAFbJpnXbrQ8qDp8X59OVIY7K"

'const strConstNetcrackerURL = "http://ln98485.ent.agt.ab.ca:23102/solutions/customers/telus/ed/navigate.jsp?to=si&byCSID="	'AT
'const strConstNetcrackerURL = "http://ln98485.ent.agt.ab.ca:23122/solutions/customers/telus/ed/navigate.jsp?to=si&byCSID="	'AT new
'const strConstNetcrackerURL = "http://ln98485.corp.ads:23122:/solutions/customers/telus/ed/navigate.jsp?to=si&byCSID="	'AT new Dec 2013
'const strConstNetcrackerURL = "http://ln98485.corp.ads:23112:/solutions/customers/telus/ed/navigate.jsp?to=si&byCSID="	'PT new Dec 2013
'const strConstNetcrackerURL = "http://ln98485.ent.agt.ab.ca:23122/solutions/customers/telus/ed/navigate.jsp?to=si&byCSID="	'PT
		
 const strConstNetcrackerURL = "http://ncrpr.tsl.telus.com/solutions/customers/telus/ed/navigate.jsp?to=si&byCSID="	'PR
' LDAP Parameters (New)
Const BypassLDAP = False
'Const propLDAPURL = "LDAP://ldap-pr.tsl.telus.com:636/"
'Const propLDAPURL = "LDAP://ldap-pr.tsl.telus.com:389/"
Const propLDAPURL = "LDAP://ldap-pr.tsl.telus.com:389/"

'Customer Profile Web Service
'Const strCPURL = "http://cm-smbsvc-west-st02.tsl.telus.com/CMO/InformationMgmt/CustomerMgmtSvc/CustomerManagementService_v2_2?wsdl"
'Const strCPURL = "http://cm-smbsvc-west-pr.tsl.telus.com/CMO/InformationMgmt/CustomerMgmtSvc/CustomerManagementService_v2_2?wsdl"
'Const strCPURL = "https://servicemngr-st01.tsl.telus.com:443/CMO/InformationMgmt/CustomerMgmtSvc/CustomerManagementService_v2_2_vs0?wsdl"
'Const strCPURL ="https://soa-mp-toll-pr.tsl.telus.com/CMO/InformationMgmt/CustomerMgmtSvc/CustomerManagementService_v2_2_1_vs0?wsdl"
Const strCPURL =""

'NC Web Service
'Const strNCWSURL = "https://apigw-soap-st.tsl.telus.com/soap/itn01/RMO/InventoryMgmt/Inventory/VPNInventoryService_v1_0_1_RP?wsdl" 'Kong itn01
'Const strNCWSURL = "https://apigw-soap-st.tsl.telus.com/soap/itn02/RMO/InventoryMgmt/Inventory/VPNInventoryService_v1_0_1_RP?wsdl" 'Kong itn02
'Const strNCWSURL = "https://apigw-soap-st.tsl.telus.com/soap/itn03/RMO/InventoryMgmt/Inventory/VPNInventoryService_v1_0_1_RP?wsdl" 'Kong itn03
'Const strNCWSURL = "https://apigw-soap-st.tsl.telus.com/soap/is05/RMO/InventoryMgmt/Inventory/VPNInventoryService_v1_0_1_RP?wsdl" 'Kong is05
Const strNCWSURL = "https://apigw-soap-st.tsl.telus.com/soap/is06/RMO/InventoryMgmt/Inventory/VPNInventoryService_v1_0_1_RP?wsdl" 'Kong is06

'Const strNCWSURL = "http://L032310.corp.ads:7001/VPNInventoryService_v1_0_VPNInventoryServicePortImpl/VPNInventoryService_v1_0?WSDL"
'Const strNCWSURL = "http://ln98485.corp.ads:23062/RMO/InventoryMgmt/Inventory/VPNInventoryService_v1_0?wsdl" ' pre-AT
'Const strNCWSURL = "http://ln98485.corp.ads:23102/RMO/InventoryMgmt/Inventory/VPNInventoryService_v1_0?wsdl" ' AT
'Const strNCWSURL = "http://ln98486.corp.ads:23112/RMO/InventoryMgmt/Inventory/VPNInventoryService_v1_0?wsdl" ' AT01
'Const strNCWSURL = "https://servicemngr-dv03.tsl.telus.com:443/RMO/InventoryMgmt/Inventory/VPNInventoryService_v1_0_vs0?wsdl" ' DV
'Const strNCWSURL = "https://servicemngr-pt61.tsl.telus.com:443/RMO/InventoryMgmt/Inventory/VPNInventoryService_v1_0_vs0?wsdl" ' AT01 virtual                      
'Const strNCWSURL = "https://soa-mp-rmsk-it01.tsl.telus.com/RMO/InventoryMgmt/Inventory/VPNInventoryService_v1_0_1_vs0?wsdl" ' AT01 virtual
'Const strNCWSURL ="https://soa-mp-rmsk-it03.tsl.telus.com/RMO/InventoryMgmt/Inventory/VPNInventoryService_v1_0_1_vs0?wsdl"
'Const strNCWSURL ="https://soa-mp-rmsk-it02.tsl.telus.com/RMO/InventoryMgmt/Inventory/VPNInventoryService_v1_0_1_vs0?wsdl"
'Const strNCWSURL = "https://soa-mp-rmsk-pr.tsl.telus.com/RMO/InventoryMgmt/Inventory/VPNInventoryService_v1_0_1_vs0?wsdl"
'Const strNCWSURL = "https://servicemngr-pt41.tsl.telus.com:443/RMO/InventoryMgmt/Inventory/VPNInventoryService_v1_0_vs0?wsdl" ' AT06 virtual
'Const strNCWSURL = "https://servicemngr-st01.tsl.telus.com:443/RMO/InventoryMgmt/Inventory/VPNInventoryService_v1_0_vs0?wsdl" ' ST01 virtual
'Const strNCWSURL = "https://servicemngr1-pr.tsl.telus.com:443/RMO/InventoryMgmt/Inventory/VPNInventoryService_v1_0_vs0?wsdl" ' PR virtual
'Const strNCWSURL ="https://soa-mp-rmsk-is09.tsl.telus.com/RMO/InventoryMgmt/Inventory/VPNInventoryService_v1_0_1_vs0?wsdl"

Const int_max_vpnpercustomer = 300                                         'Maximum number of vpns/customer records returned from the NC web service call (getVPNListByCustomer)
'Const strPermItWSURL= "https://soa-mp-toll-pr.tsl.telus.com/EO/SecurityMgmt/AuthorizationPolicyService_v1_1_vs0?wsdl"
'Const strPermItWSURL="http://ln98949.corp.ads:50010/EO/SecurityMgmt/Authz/AuthorizationPolicyService_v1_1?wsdl"
Const strPermItWSURL = "https://soa-mp-toll-pt02.tsl.telus.com/EO/SecurityMgmt/AuthorizationPolicyService_v1_1_vs0?WSDL"
'Const strPermItWSURL= "https://soa-mp-toll-pr.tsl.telus.com/EO/SecurityMgmt/AuthorizationPolicyService_v1_1_vs0?wsdl"


const SRT2_URL="https://srt2.tsl.telus.com/" 		
const ESD_REP_URL="http://tacabcbrs01v.corp.ads/ESDReports/"

%>

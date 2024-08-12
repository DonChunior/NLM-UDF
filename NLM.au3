; AU3Check settings
#AutoIt3Wrapper_Run_AU3Check=Y
#AutoIt3Wrapper_Au3Check_Parameters=-d -w 1 -w 2 -w 3 -w 4 -w 5 -w 6 -w 7
#AutoIt3Wrapper_AU3Check_Stop_OnWarning=Y
#include-once

; #INDEX# =======================================================================================================================
; Title .........: Network List Manager UDF
; Description ...: The Network List Manager API enables applications to retrieve a list of available network connections.
;                  Applications can filter networks, based on attributes and signatures, and choose the networks best suited to
;                  their task.
;                  The Network List Manager infrastructure notifies applications of changes in the network environment, thus
;                  enabling applications to dynamically update network connections.
; Links .........: https://learn.microsoft.com/en-us/windows/win32/NLA/portal
; Version .......: 1.0.0-alpha
; AutoIt Version : 3.3.16.1
; Author(s) .....: Domenic Laritz
; Dll ...........:
; ===============================================================================================================================

; #CURRENT# =====================================================================================================================
;_NLM_NetworkListManager_ClearSimulatedProfileInfo (TO DO)
;_NLM_NetworkListManager_GetConnectivity
;_NLM_NetworkListManager_GetNetworkConnection (TO DO)
;_NLM_NetworkListManager_GetNetworkConnections
;_NLM_NetworkListManager_GetNetwork (TO DO)
;_NLM_NetworkListManager_GetNetworks
;_NLM_NetworkListManager_IsConnected
;_NLM_NetworkListManager_IsConnectedToInternet
;_NLM_NetworkListManager_SetSimulatedProfileInfo (TO DO)
;_NLM_Network_GetCategory
;_NLM_Network_GetConnectivity (TO DO)
;_NLM_Network_GetDescription
;_NLM_Network_GetDomainType
;_NLM_Network_GetName
;_NLM_Network_GetNetworkConnections (TO DO)
;_NLM_Network_GetTimeCreatedAndConnected (TO DO)
;_NLM_Network_GetNetworkId (TO DO)
;_NLM_Network_IsConnected (TO DO)
;_NLM_Network_IsConnectedToInternet (TO DO)
;_NLM_Network_SetCategory (TO DO)
;_NLM_Network_SetDescription (TO DO)
;_NLM_Network_SetName (TO DO)
; ===============================================================================================================================

; #INTERNAL_USE_ONLY# ===========================================================================================================
;__NLM_InitializeObject
; ===============================================================================================================================

; #VARIABLES# ===================================================================================================================
Global $__g_oNLM = 0
; ===============================================================================================================================

; #CONSTANTS# ===================================================================================================================

; NLM_CONNECTIVITY enumeration (https://learn.microsoft.com/en-us/windows/win32/api/netlistmgr/ne-netlistmgr-nlm_connectivity)
Global Const $NLM_CONNECTIVITY_DISCONNECTED = 0x0000
Global Const $NLM_CONNECTIVITY_IPV4_NOTRAFFIC = 0x0001
Global Const $NLM_CONNECTIVITY_IPV6_NOTRAFFIC = 0x0002
Global Const $NLM_CONNECTIVITY_IPV4_SUBNET = 0x0010
Global Const $NLM_CONNECTIVITY_IPV4_LOCALNETWORK = 0x0020
Global Const $NLM_CONNECTIVITY_IPV4_INTERNET = 0x0040
Global Const $NLM_CONNECTIVITY_IPV6_SUBNET = 0x0100
Global Const $NLM_CONNECTIVITY_IPV6_LOCALNETWORK = 0x0200
Global Const $NLM_CONNECTIVITY_IPV6_INTERNET = 0x0400

; NLM_DOMAIN_TYPE enumeration (https://learn.microsoft.com/en-us/windows/win32/api/netlistmgr/ne-netlistmgr-nlm_domain_type)
Global Const $NLM_DOMAIN_TYPE_NON_DOMAIN_NETWORK = 0x0
Global Const $NLM_DOMAIN_TYPE_DOMAIN_NETWORK = 0x01
Global Const $NLM_DOMAIN_TYPE_DOMAIN_AUTHENTICATED = 0x02

; NLM_ENUM_NETWORK enumeration (https://learn.microsoft.com/en-us/windows/win32/api/netlistmgr/ne-netlistmgr-nlm_enum_network)
Global Const $NLM_ENUM_NETWORK_CONNECTED = 0x01
Global Const $NLM_ENUM_NETWORK_DISCONNECTED = 0x02
Global Const $NLM_ENUM_NETWORK_ALL = 0x03

; NLM_NETWORK_CATEGORY enumeration (https://learn.microsoft.com/en-us/windows/win32/api/netlistmgr/ne-netlistmgr-nlm_network_category)
Global Const $NLM_NETWORK_CATEGORY_PUBLIC = 0x00
Global Const $NLM_NETWORK_CATEGORY_PRIVATE = 0x01
Global Const $NLM_NETWORK_CATEGORY_DOMAIN_AUTHENTICATED = 0x02

Global Const $__NLMCONSTANT_sIID_IBackgroundCopyError = "{DCB00C01-570F-4A9B-8D69-199FDBA5723B}"

; ===============================================================================================================================

; #FUNCTION# ====================================================================================================================
; Name ..........: _NLM_NetworkListManager_GetConnectivity
; Description ...: Returns the aggregated connectivity state of all networks on this machine.
; Syntax ........: _NLM_NetworkListManager_GetConnectivity()
; Parameters ....: None
; Return values .: None
; Author ........: Domenic Laritz
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........: https://learn.microsoft.com/en-us/windows/win32/api/netlistmgr/nf-netlistmgr-inetworklistmanager-getconnectivity
; Example .......: No
; ===============================================================================================================================
Func _NLM_NetworkListManager_GetConnectivity()
	If Not $__g_oNLM Then __NLM_InitializeObject()
	Local $iConnectivity = $__g_oNLM.GetConnectivity
	Return SetError(0, 0, $iConnectivity)
EndFunc   ;==>_NLM_NetworkListManager_GetConnectivity

; #FUNCTION# ====================================================================================================================
; Name ..........: _NLM_NetworkListManager_GetNetworkConnections
; Description ...: Enumerate the complete list of all connections in your compartment.
; Syntax ........: _NLM_NetworkListManager_GetNetworkConnections()
; Parameters ....: None
; Return values .: None
; Author ........: Domenic Laritz
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........: https://learn.microsoft.com/en-us/windows/win32/api/netlistmgr/nf-netlistmgr-inetworklistmanager-getnetworkconnections
; Example .......: No
; ===============================================================================================================================
Func _NLM_NetworkListManager_GetNetworkConnections()
	If Not $__g_oNLM Then __NLM_InitializeObject()
	Local $oNetworkConnections = $__g_oNLM.GetNetworkConnections
	Local $aoNetworkConnections[0]
	For $oNetworkConnection In $oNetworkConnections
		ReDim $aoNetworkConnections[UBound($aoNetworkConnections) + 1]
		$aoNetworkConnections[UBound($aoNetworkConnections) - 1] = $oNetworkConnection
	Next
	Return SetError(0, 0, $aoNetworkConnections)
EndFunc   ;==>_NLM_NetworkListManager_GetNetworkConnections

; #FUNCTION# ====================================================================================================================
; Name ..........: _NLM_NetworkListManager_GetNetworks
; Description ...: Enumerate the list of networks in your compartment.
; Syntax ........: _NLM_NetworkListManager_GetNetworks(Const Byref $iFlags)
; Parameters ....: $iFlags              - [in/out and const] an integer value.
; Return values .: None
; Author ........: Domenic Laritz
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........: https://learn.microsoft.com/en-us/windows/win32/api/netlistmgr/nf-netlistmgr-inetworklistmanager-getnetworks
; Example .......: No
; ===============================================================================================================================
Func _NLM_NetworkListManager_GetNetworks(Const ByRef $iFlags)
	If Not $__g_oNLM Then __NLM_InitializeObject()
	Local $oNetworks = $__g_oNLM.GetNetworks($iFlags)
	Local $aoNetworks[0]
	For $oNetwork In $oNetworks
		ReDim $aoNetworks[UBound($aoNetworks) + 1]
		$aoNetworks[UBound($aoNetworks) - 1] = $oNetwork
	Next
	Return SetError(0, 0, $aoNetworks)
EndFunc   ;==>_NLM_NetworkListManager_GetNetworks

; #FUNCTION# ====================================================================================================================
; Name ..........: _NLM_NetworkListManager_IsConnected
; Description ...: Returns whether this machine has any network connectivity.
; Syntax ........: _NLM_NetworkListManager_IsConnected()
; Parameters ....: None
; Return values .: None
; Author ........: Domenic Laritz
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........: https://learn.microsoft.com/en-us/windows/win32/api/netlistmgr/nf-netlistmgr-inetworklistmanager-get_isconnected
; Example .......: No
; ===============================================================================================================================
Func _NLM_NetworkListManager_IsConnected()
	If Not $__g_oNLM Then __NLM_InitializeObject()
	Local $bIsConnected = $__g_oNLM.IsConnected
	Return SetError(0, 0, $bIsConnected)
EndFunc   ;==>_NLM_NetworkListManager_IsConnected

; #FUNCTION# ====================================================================================================================
; Name ..........: _NLM_NetworkListManager_IsConnectedToInternet
; Description ...: Returns whether this machine has Internet connectivity.
; Syntax ........: _NLM_NetworkListManager_IsConnectedToInternet()
; Parameters ....: None
; Return values .: None
; Author ........: Domenic Laritz
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........: https://learn.microsoft.com/en-us/windows/win32/api/netlistmgr/nf-netlistmgr-inetworklistmanager-get_isconnectedtointernet
; Example .......: No
; ===============================================================================================================================
Func _NLM_NetworkListManager_IsConnectedToInternet()
	If Not $__g_oNLM Then __NLM_InitializeObject()
	Local $bIsConnectedToInternet = $__g_oNLM.IsConnectedToInternet
	Return SetError(0, 0, $bIsConnectedToInternet)
EndFunc   ;==>_NLM_NetworkListManager_IsConnectedToInternet

; #FUNCTION# ====================================================================================================================
; Name ..........: _NLM_Network_GetCategory
; Description ...: Returns the category of this network.
; Syntax ........: _NLM_Network_GetCategory(Const Byref $oNetwork)
; Parameters ....: $oNetwork            - [in/out and const] an object.
; Return values .: None
; Author ........: Domenic Laritz
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........: https://learn.microsoft.com/en-us/windows/win32/api/netlistmgr/nf-netlistmgr-inetwork-getcategory
; Example .......: No
; ===============================================================================================================================
Func _NLM_Network_GetCategory(Const ByRef $oNetwork)
	Local $iCategory = $oNetwork.GetCategory
	Return SetError(0, 0, $iCategory)
EndFunc   ;==>_NLM_Network_GetCategory

; #FUNCTION# ====================================================================================================================
; Name ..........: _NLM_Network_GetDescription
; Description ...: Get the network description.
; Syntax ........: _NLM_Network_GetDescription(Const Byref $oNetwork)
; Parameters ....: $oNetwork            - [in/out and const] an object.
; Return values .: None
; Author ........: Domenic Laritz
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........: https://learn.microsoft.com/en-us/windows/win32/api/netlistmgr/nf-netlistmgr-inetwork-getdescription
; Example .......: No
; ===============================================================================================================================
Func _NLM_Network_GetDescription(Const ByRef $oNetwork)
	Local $sDescription = $oNetwork.GetDescription
	Return SetError(0, 0, $sDescription)
EndFunc   ;==>_NLM_Network_GetDescription

; #FUNCTION# ====================================================================================================================
; Name ..........: _NLM_Network_GetDomainType
; Description ...: Get the domain type.
; Syntax ........: _NLM_Network_GetDomainType(Const Byref $oNetwork)
; Parameters ....: $oNetwork            - [in/out and const] an object.
; Return values .: None
; Author ........: Domenic Laritz
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........: https://learn.microsoft.com/en-us/windows/win32/api/netlistmgr/nf-netlistmgr-inetwork-getdomaintype
; Example .......: No
; ===============================================================================================================================
Func _NLM_Network_GetDomainType(Const ByRef $oNetwork)
	Local $iDomainType = $oNetwork.GetDomainType
	Return SetError(0, 0, $iDomainType)
EndFunc   ;==>_NLM_Network_GetDomainType

; #FUNCTION# ====================================================================================================================
; Name ..........: _NLM_Network_GetName
; Description ...: Get the name of this network.
; Syntax ........: _NLM_Network_GetName(Const Byref $oNetwork)
; Parameters ....: $oNetwork            - [in/out and const] an object.
; Return values .: None
; Author ........: Domenic Laritz
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........: https://learn.microsoft.com/en-us/windows/win32/api/netlistmgr/nf-netlistmgr-inetwork-getname
; Example .......: No
; ===============================================================================================================================
Func _NLM_Network_GetName(Const ByRef $oNetwork)
	Local $sName = $oNetwork.GetName
	Return SetError(0, 0, $sName)
EndFunc   ;==>_NLM_Network_GetName

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __NLM_InitializeObject
; Description ...:
; Syntax ........: __NLM_InitializeObject()
; Parameters ....: None
; Return values .: None
; Author ........: Domenic Laritz
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __NLM_InitializeObject()
	$__g_oNLM = ObjCreate($__NLMCONSTANT_sIID_IBackgroundCopyError)
EndFunc   ;==>__NLM_InitializeObject

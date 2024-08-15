; AU3Check settings
#AutoIt3Wrapper_Run_AU3Check=Y
#AutoIt3Wrapper_Au3Check_Parameters=-d -w 1 -w 2 -w 3 -w 4 -w 5 -w 6 -w 7
#AutoIt3Wrapper_AU3Check_Stop_OnWarning=Y
#include-once

#include <StructureConstants.au3>

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
; ===============================================================================================================================

; #INTERNAL_USE_ONLY# ===========================================================================================================
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

Global Const $__NLMCONSTANT_sCLSID_NetworkListManager = "{DCB00C01-570F-4A9B-8D69-199FDBA5723B}"
Global Const $__NLMCONSTANT_sIID_INetworkConnection = "{DCB00005-570F-4A9B-8D69-199FDBA5723B}"
Global Const $__NLMCONSTANT_sIID_INetworkListManager = "{DCB00000-570F-4A9B-8D69-199FDBA5723B}"
Global Const $__NLMCONSTANT_sTagIDispatch = _
		"GetTypeInfoCount hresult(uint_ptr*);" & _
		"GetTypeInfo hresult(uint;dword;ptr*);" & _
		"GetIDsOfNames hresult(struct*;wstr*;uint;dword;long_ptr*);" & _
		"Invoke hresult(long;struct*;dword;word;struct*;variant*;ptr*;uint_ptr*);"
Global Const $__NLMCONSTANT_sTag_INetworkConnection = _
		$__NLMCONSTANT_sTagIDispatch & _
		"GetNetwork hresult(ptr*);" & _ ; TO DO
		"IsConnectedToInternet hresult(bool*);" & _ ; TO DO
		"IsConnected hresult(bool*);" & _ ; TO DO
		"GetConnectivity hresult(int_ptr*);" & _ ; TO DO
		"GetConnectionId hresult(struct*);" & _ ; TO DO
		"GetAdapterId hresult(struct*);" & _ ; TO DO
		"GetDomainType hresult(int_ptr*);" ; TO DO
Global Const $__NLMCONSTANT_sTag_INetworkListManager = _
		$__NLMCONSTANT_sTagIDispatch & _
		"GetNetworks hresult(int;ptr*);" & _ ; TO DO
		"GetNetwork hresult(" & $tagGUID & ";ptr*);" & _ ; TO DO
		"GetNetworkConnections hresult(ptr*);" & _ ; TO DO
		"GetNetworkConnection hresult(" & $tagGUID & ";ptr*);" & _ ; TO DO
		"IsConnectedToInternet hresult(bool*);" & _
		"IsConnected hresult(bool*);" & _
		"GetConnectivity hresult(int_ptr*);" & _
		"SetSimulatedProfileInfo hresult(struct*);" & _ ; TO DO
		"ClearSimulatedProfileInfo hresult();" ; TO DO

; ===============================================================================================================================

Func _NLM_NetworkListManager_GetConnectivity()
	If Not $__g_oNLM Then __NLM_InitializeNLMObject()
	Local $iConnectivity = 0
	$__g_oNLM.GetConnectivity($iConnectivity)
	Return SetError(0, 0, $iConnectivity)
EndFunc   ;==>_NLM_NetworkListManager_GetConnectivity

Func _NLM_NetworkListManager_GetNetworkConnections()
	If Not $__g_oNLM Then __NLM_InitializeNLMObject()
	Local $pNetworkConnections = 0
	$__g_oNLM.GetNetworkConnections($pNetworkConnections)
	Return SetError(0, 0, $pNetworkConnections)
EndFunc   ;==>_NLM_NetworkListManager_GetNetworkConnections

Func _NLM_NetworkListManager_IsConnected()
	If Not $__g_oNLM Then __NLM_InitializeNLMObject()
	Local $bConnected = 0
	$__g_oNLM.IsConnected($bConnected)
	Return SetError(0, 0, $bConnected)
EndFunc   ;==>_NLM_NetworkListManager_IsConnected

Func _NLM_NetworkListManager_IsConnectedToInternet()
	If Not $__g_oNLM Then __NLM_InitializeNLMObject()
	Local $bConnected = 0
	$__g_oNLM.IsConnectedToInternet($bConnected)
	Return SetError(0, 0, $bConnected)
EndFunc   ;==>_NLM_NetworkListManager_IsConnectedToInternet

Func __NLM_InitializeNLMObject()
	$__g_oNLM = ObjCreateInterface($__NLMCONSTANT_sCLSID_NetworkListManager, $__NLMCONSTANT_sIID_INetworkListManager, $__NLMCONSTANT_sTag_INetworkListManager)
	Return SetError(0, 0, 0)
EndFunc   ;==>__NLM_InitializeNLMObject

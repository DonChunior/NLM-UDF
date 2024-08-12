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
;$tagSTRUCT
;_NLM_GetConnectivity
;_NLM_GetNetworks
;_NLM_IsConnected
;_NLM_IsConnectedToInternet
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

; NLM_ENUM_NETWORK enumeration (https://learn.microsoft.com/en-us/windows/win32/api/netlistmgr/ne-netlistmgr-nlm_enum_network)
Global Const $NLM_ENUM_NETWORK_CONNECTED = 0x01
Global Const $NLM_ENUM_NETWORK_DISCONNECTED = 0x02
Global Const $NLM_ENUM_NETWORK_ALL = 0x03

Global Const $__NLMCONSTANT_sIID_IBackgroundCopyError = "{DCB00C01-570F-4A9B-8D69-199FDBA5723B}"

; ===============================================================================================================================

; #FUNCTION# ====================================================================================================================
; Name ..........: _NLM_GetConnectivity
; Description ...: Returns the aggregated connectivity state of all networks on this machine.
; Syntax ........: _NLM_GetConnectivity()
; Parameters ....: None
; Return values .: None
; Author ........: Domenic Laritz
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........: https://learn.microsoft.com/en-us/windows/win32/api/netlistmgr/nf-netlistmgr-inetworklistmanager-getconnectivity
; Example .......: No
; ===============================================================================================================================
Func _NLM_GetConnectivity()
	If Not $__g_oNLM Then __NLM_InitializeObject()
	Local $iConnectivity = $__g_oNLM.GetConnectivity
	Return SetError(0, 0, $iConnectivity)
EndFunc   ;==>_NLM_GetConnectivity

; #FUNCTION# ====================================================================================================================
; Name ..........: _NLM_GetNetworks
; Description ...: Enumerate the list of networks in your compartment.
; Syntax ........: _NLM_GetNetworks(Const Byref $iFlags)
; Parameters ....: $iFlags              - [in/out and const] an integer value.
; Return values .: None
; Author ........: Domenic Laritz
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........: https://learn.microsoft.com/en-us/windows/win32/api/netlistmgr/nf-netlistmgr-inetworklistmanager-getnetworks
; Example .......: No
; ===============================================================================================================================
Func _NLM_GetNetworks(Const ByRef $iFlags)
	If Not $__g_oNLM Then __NLM_InitializeObject()
	Local $oNetworks = $__g_oNLM.GetNetworks($iFlags)
	Local $aoNetworks[0]
	For $oNetwork In $oNetworks
		ReDim $aoNetworks[UBound($aoNetworks) + 1]
		$aoNetworks[UBound($aoNetworks) - 1] = $oNetwork
	Next
	Return SetError(0, 0, $aoNetworks)
EndFunc   ;==>_NLM_GetNetworks

; #FUNCTION# ====================================================================================================================
; Name ..........: _NLM_IsConnected
; Description ...: Returns whether this machine has any network connectivity.
; Syntax ........: _NLM_IsConnected()
; Parameters ....: None
; Return values .: None
; Author ........: Domenic Laritz
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........: https://learn.microsoft.com/en-us/windows/win32/api/netlistmgr/nf-netlistmgr-inetworklistmanager-get_isconnected
; Example .......: No
; ===============================================================================================================================
Func _NLM_IsConnected()
	If Not $__g_oNLM Then __NLM_InitializeObject()
	Local $bIsConnected = $__g_oNLM.IsConnected
	Return SetError(0, 0, $bIsConnected)
EndFunc   ;==>_NLM_IsConnected

; #FUNCTION# ====================================================================================================================
; Name ..........: _NLM_IsConnectedToInternet
; Description ...: Returns whether this machine has Internet connectivity.
; Syntax ........: _NLM_IsConnectedToInternet()
; Parameters ....: None
; Return values .: None
; Author ........: Domenic Laritz
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........: https://learn.microsoft.com/en-us/windows/win32/api/netlistmgr/nf-netlistmgr-inetworklistmanager-get_isconnectedtointernet
; Example .......: No
; ===============================================================================================================================
Func _NLM_IsConnectedToInternet()
	If Not $__g_oNLM Then __NLM_InitializeObject()
	Local $bIsConnectedToInternet = $__g_oNLM.IsConnectedToInternet
	Return SetError(0, 0, $bIsConnectedToInternet)
EndFunc   ;==>_NLM_IsConnectedToInternet

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

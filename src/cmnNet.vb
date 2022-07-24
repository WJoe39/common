Option Strict Off
Option Explicit On 
Module cmnNet
#Region "Head"
    '******************************************************************************
    '*  MODULE        : cmnNet
    '*  FILE          : cmnNet.vb
    '*  PROJECT       : non-specific
    '*  AUTHOR        : Christoph A. Lutz
    '*  CREATED       : 01-Apr-2007
    '*  COPYRIGHT     : Copyright (c) 2007-2011 Christoph A. Lutz.
    '*                  All Rights Reserved.
    '*
    '*                  This module is free software; you can redistribute it
    '*                  and/or modify it under the terms of the GNU General
    '*                  Public License as published by the Free Software
    '*                  Foundation; either version 2 of the License, or any later
    '*                  version.
    '*
    '*                  All copyright notices regarding Chris A. Lutz must remain
    '*                  intact in the source code and in the outputted text.
    '*
    '*                  This program is distributed in the hope that it will be
    '*                  useful, but WITHOUT ANY WARRANTY; without even the implied
    '*                  warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR
    '*                  PURPOSE. See the GNU General Public License for more details.
    '*
    '*  DESCRIPTION   : Common network operations
    '*
    '*  MODIFICATION HISTORY:
    '*  AUTHOR:             DATE:       CHANGES:
    '*  Christoph A. Lutz   01-Apr-2007 Initial Version
    '*
    '******************************************************************************
#End Region
#Region "Declaration"
    '-------------- constants definition ------------------------------------------
    Private Const VB_MODULE As String = "cmnNet"

    Private Const MAX_WSADescription As Integer = 256
    Private Const MAX_WSASYSStatus As Integer = 128
    Private Const WS_VERSION_REQD As Integer = &H101S

    '   Ping
    Private Const IP_SUCCESS As Integer = 0
    Private Const INADDR_NONE As Integer = &HFFFFFFFF
    Private Const IP_BUF_TOO_SMALL As Integer = (11000 + 1)
    Private Const IP_DEST_NET_UNREACHABLE As Integer = (11000 + 2)
    Private Const IP_DEST_HOST_UNREACHABLE As Integer = (11000 + 3)
    Private Const IP_DEST_PROT_UNREACHABLE As Integer = (11000 + 4)
    Private Const IP_DEST_PORT_UNREACHABLE As Integer = (11000 + 5)
    Private Const IP_NO_RESOURCES As Integer = (11000 + 6)
    Private Const IP_BAD_OPTION As Integer = (11000 + 7)
    Private Const IP_HW_ERROR As Integer = (11000 + 8)
    Private Const IP_PACKET_TOO_BIG As Integer = (11000 + 9)
    Private Const IP_REQ_TIMED_OUT As Integer = (11000 + 10)
    Private Const IP_BAD_REQ As Integer = (11000 + 11)
    Private Const IP_BAD_ROUTE As Integer = (11000 + 12)
    Private Const IP_TTL_EXPIRED_TRANSIT As Integer = (11000 + 13)
    Private Const IP_TTL_EXPIRED_REASSEM As Integer = (11000 + 14)
    Private Const IP_PARAM_PROBLEM As Integer = (11000 + 15)
    Private Const IP_SOURCE_QUENCH As Integer = (11000 + 16)
    Private Const IP_OPTION_TOO_BIG As Integer = (11000 + 17)
    Private Const IP_BAD_DESTINATION As Integer = (11000 + 18)
    Private Const IP_ADDR_DELETED As Integer = (11000 + 19)
    Private Const IP_SPEC_MTU_CHANGE As Integer = (11000 + 20)
    Private Const IP_MTU_CHANGE As Integer = (11000 + 21)
    Private Const IP_UNLOAD As Integer = (11000 + 22)
    Private Const IP_ADDR_ADDED As Integer = (11000 + 23)
    Private Const IP_NEGOTIATING_IPSEC As Integer = (11000 + 32)
    Private Const IP_GENERAL_FAILURE As Integer = (11000 + 50)
    Private Const IP_PENDING As Integer = (11000 + 255)
    Private Const PING_TIMEOUT As Integer = 500

    '   Network connection
    Private Const NO_ERROR As Integer = 0
    Private Const CONNECT_UPDATE_PROFILE As Integer = &H1S

    Private Const RESOURCETYPE_DISK As Integer = &H1S
    'Private Const RESOURCETYPE_PRINT As Integer = &H2S
    'Private Const RESOURCETYPE_ANY As Integer = &H0S
    'Private Const RESOURCE_CONNECTED As Integer = &H1S
    'Private Const RESOURCE_REMEMBERED As Integer = &H3S
    'Private Const RESOURCE_GLOBALNET As Integer = &H2S
    'Private Const RESOURCEDISPLAYTYPE_DOMAIN As Integer = &H1S
    'Private Const RESOURCEDISPLAYTYPE_GENERIC As Integer = &H0S
    'Private Const RESOURCEDISPLAYTYPE_SERVER As Integer = &H2S
    'Private Const RESOURCEDISPLAYTYPE_SHARE As Integer = &H3S
    'Private Const RESOURCEUSAGE_CONNECTABLE As Integer = &H1S
    'Private Const RESOURCEUSAGE_CONTAINER As Integer = &H2S

    '   Network connection errors
    Private Const ERROR_ACCESS_DENIED As Integer = 5
    Private Const ERROR_ALREADY_ASSIGNED As Integer = 85
    Private Const ERROR_BAD_DEV_TYPE As Integer = 66
    Private Const ERROR_BAD_DEVICE As Integer = 1200
    Private Const ERROR_BAD_NET_NAME As Integer = 67
    Private Const ERROR_BAD_PROFILE As Integer = 1206
    Private Const ERROR_BAD_PROVIDER As Integer = 1204
    Private Const ERROR_BUSY As Integer = 170
    Private Const ERROR_CANCELLED As Integer = 1223
    Private Const ERROR_CANNOT_OPEN_PROFILE As Integer = 1205
    Private Const ERROR_DEVICE_ALREADY_REMEMBERED As Integer = 1202
    Private Const ERROR_EXTENDED_ERROR As Integer = 1208
    Private Const ERROR_INVALID_PASSWORD As Integer = 86
    Private Const ERROR_NO_NET_OR_BAD_PATH As Integer = 1203
    Private Const ERROR_NO_NETWORK As Integer = 1222
    Private Const ERROR_SESSION_CREDENTIAL_CONFLICT As Integer = 1219

    '-------------- libraries -----------------------------------------------------
    Friend Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, ByRef nSize As Integer) As Integer

    Private Declare Function WSAStartup Lib "WSOCK32.DLL" (ByVal wVersionRequired As Integer, ByRef lpWSADATA As WSADATA) As Integer
    Private Declare Function gethostbyname Lib "WSOCK32.DLL" (ByVal hostname As String) As Integer
    Private Declare Function inet_addr Lib "WSOCK32.DLL" (ByVal s As String) As Integer
    Private Declare Function IcmpCreateFile Lib "icmp.dll" () As Integer
    Private Declare Function IcmpSendEcho Lib "icmp.dll" (ByVal IcmpHandle As Integer, ByVal DestinationAddress As Integer, ByVal RequestData As String, ByVal RequestSize As Integer, ByVal RequestOptions As Integer, ByRef ReplyBuffer As ICMP_ECHO_REPLY, ByVal ReplySize As Integer, ByVal Timeout As Integer) As Integer
    Private Declare Function IcmpCloseHandle Lib "icmp.dll" (ByVal IcmpHandle As Integer) As Integer
    Private Declare Function WNetAddConnection2 Lib "mpr.dll" Alias "WNetAddConnection2A" (ByRef lpNetResource As NETRESOURCE, ByVal lpPassword As String, ByVal lpUsername As String, ByVal dwFlags As Integer) As Integer
    Private Declare Function WNetCancelConnection2 Lib "mpr.dll" Alias "WNetCancelConnection2A" (ByVal lpName As String, ByVal dwFlags As Integer, ByVal fForce As Integer) As Integer
    Private Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Integer

    '-------------- types definition ----------------------------------------------
    Friend Structure ICMP_OPTIONS
        Dim Ttl As Byte
        Dim Tos As Byte
        Dim Flags As Byte
        Dim OptionsSize As Byte
        Dim OptionsData As Integer
    End Structure
    Friend Structure ICMP_ECHO_REPLY
        Dim Address As Integer
        Dim Status As Integer
        Dim RoundTripTime As Integer
        Dim DataSize As Integer
        Dim DataPointer As Integer
        Dim Options As ICMP_OPTIONS
        <VBFixedString(250), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValTStr, SizeConst:=250)> Public Data As String
    End Structure
    Private Structure WSADATA
        Dim wVersion As Short
        Dim wHighVersion As Short
        <VBFixedArray(MAX_WSADescription)> Dim szDescription() As Byte
        <VBFixedArray(MAX_WSASYSStatus)> Dim szSystemStatus() As Byte
        Dim wMaxSockets As Integer
        Dim wMaxUDPDG As Integer
        Dim dwVendorInfo As Integer

        'UPGRADE_TODO: Zum Initialisieren der Instanzen dieser Struktur muss "Initialize" aufgerufen werden. Klicken Sie hier für weitere Informationen: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1026"'
        Public Sub Initialize()
            ReDim szDescription(MAX_WSADescription)
            ReDim szSystemStatus(MAX_WSASYSStatus)
        End Sub
    End Structure
    Private Structure NETRESOURCE
        Dim dwScope As Integer
        Dim dwType As Integer
        Dim dwDisplayType As Integer
        Dim dwUsage As Integer
        Dim lpLocalName As String
        Dim lpRemoteName As String
        Dim lpComment As String
        Dim lpProvider As String
    End Structure

    '-------------- symbol definition ---------------------------------------------
    Private lpNetResource As NETRESOURCE
#End Region
#Region "Properties"
#End Region
#Region "Methods"
    '-------------- procedure & function definition--------------------------------
    Friend Function bIsSocketInitialized() As Boolean
        '**************************************************************************
        ' bIsSocketInitialized (FUNCTION)
        '
        '  PURPOSE      : Initializes network socket
        '  PARAMETERS   : -
        '  RETURN VALUE : Returns false if an error occured
        '
        '**************************************************************************

        'UPGRADE_WARNING: Arrays in Struktur udtWinSocketData müssen möglicherweise initialisiert werden, bevor sie verwendet werden können. Klicken Sie hier für weitere Informationen: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1063"'
        Dim udtWinSocketData As WSADATA

        '   Error handling
        On Error Resume Next

        bIsSocketInitialized = WSAStartup(WS_VERSION_REQD, udtWinSocketData) = IP_SUCCESS

    End Function
    Friend Function sHostToIP(ByVal sHostName As String) As String
        '**************************************************************************
        ' sHostToIP (FUNCTION)
        '
        '  PURPOSE      : Lookup IP address by hostname
        '  PARAMETERS   : (IN) sHostName(String) - Host name
        '  RETURN VALUE : IP-Address
        '
        '**************************************************************************

        Try
            Dim clHostInfo As System.Net.IPHostEntry
            Dim objIPAddress As System.Net.IPAddress()
            Dim sIPAddress As String

            clHostInfo = System.Net.Dns.Resolve(sHostName)
            objIPAddress = clHostInfo.AddressList
            sIPAddress = objIPAddress(0).ToString()
            Return sIPAddress

        Catch err As Exception
            RaiseErr(VB_MODULE & ".sIPToHost", err.Message)
        End Try

    End Function
    Private Function sIPToHost(ByVal sIPAddress As String) As String
        '**************************************************************************
        ' sIPToHost (FUNCTION)
        '
        '  PURPOSE      : Lookup hostname by IP address
        '  PARAMETERS   : (IN) sIPAddress(String) - IP address
        '  RETURN VALUE : IP-Address
        '
        '**************************************************************************
        Try
            Dim clHostInfo As New System.Net.IPHostEntry
            Dim sHostname As String

            clHostInfo = System.Net.Dns.GetHostByAddress(sIPAddress)
            sHostname = clHostInfo.HostName
            Return sHostname
        Catch err As Exception
            RaiseErr(VB_MODULE & ".sIPToHost", err.Message)
        End Try
    End Function
    Friend Function sPing(ByRef sIPAddress As String, ByRef sDataToSend As String, ByRef udtEcho As ICMP_ECHO_REPLY) As Integer
        '**************************************************************************
        ' sPing (FUNCTION)
        '
        '  PURPOSE      : Pings an address
        '  PARAMETERS   : (IN) sIPAddress(String) - IP address of destination host
        '  RETURN VALUE : Pong
        '
        '**************************************************************************

        Dim lDestinationAddress As Integer
        Dim hWnd As Integer

        '   Error handling
        On Error Resume Next

        lDestinationAddress = inet_addr(sIPAddress)
        If lDestinationAddress <> INADDR_NONE Then
            hWnd = IcmpCreateFile()
            If hWnd Then
                IcmpSendEcho(hWnd, lDestinationAddress, sDataToSend, Len(sDataToSend), 0, udtEcho, Len(udtEcho), PING_TIMEOUT)
                sPing = udtEcho.Status
                IcmpCloseHandle(hWnd)
            End If
        Else
            sPing = INADDR_NONE
        End If

    End Function
    Friend Function sConnectedNetworkDrive(ByVal sNetworkConnection As String, ByVal sUsername As String, ByVal sPassword As String, Optional ByVal sDriveLetter As String = "") As String
        '**************************************************************************
        ' sConnectedNetworkDrive (FUNCTION)
        '
        '  PURPOSE      : Maps a network resource to a local drive letter
        '  PARAMETERS   : (IN) sNetworkConnection(String) - UNC path of the net res
        '                 (IN) sUsername(String) - Username
        '                 (IN) sPassword(String) - Password
        '                 (IN) sDriveLetter(String) - Drive letter
        '  RETURN VALUE : Drive letter of the connected network drive
        '
        '**************************************************************************

        Dim sLocalDrive As String
        Dim lReturnCode As Integer
        Dim sErrorMessage As String

        '   Error handling
        On Error Resume Next

        '   Local Init
        sErrorMessage = ""

        '   drive letter
        If sDriveLetter = "" Or Asc(UCase(sDriveLetter)) < 65 Or Asc(UCase(sDriveLetter)) > 100 Then
            sLocalDrive = sFreeDriveLetter()
        Else
            sLocalDrive = UCase(Left(sDriveLetter, 1))
        End If

        '   map drive
        lpNetResource.dwType = RESOURCETYPE_DISK
        lpNetResource.lpLocalName = sLocalDrive & vbNullChar
        lpNetResource.lpRemoteName = sNetworkConnection & vbNullChar
        lpNetResource.lpProvider = vbNullChar
        lReturnCode = WNetAddConnection2(lpNetResource, sPassword & vbNullChar, sUsername & vbNullChar, CONNECT_UPDATE_PROFILE)

        Select Case lReturnCode
            Case NO_ERROR
                sConnectedNetworkDrive = sLocalDrive
            Case ERROR_ACCESS_DENIED
                sErrorMessage = "Access is denied."
            Case ERROR_ALREADY_ASSIGNED
                sErrorMessage = "The local device name is already in use."
            Case ERROR_BAD_DEV_TYPE
                sErrorMessage = "The network resource type is not correct."
            Case ERROR_BAD_DEVICE
                sErrorMessage = "The specified device name is invalid."
            Case ERROR_BAD_NET_NAME
                sErrorMessage = "The network name cannot be found."
            Case ERROR_BAD_PROFILE
                sErrorMessage = "The network connection profile is corrupted."
            Case ERROR_BAD_PROVIDER
                sErrorMessage = "The specified network provider name is invalid."
            Case ERROR_BUSY
                sErrorMessage = "The specified network provider name is invalid."
            Case ERROR_CANCELLED
                sErrorMessage = "The operation was canceled by the user."
            Case ERROR_CANNOT_OPEN_PROFILE
                sErrorMessage = "Unable to open the network connection profile."
            Case ERROR_DEVICE_ALREADY_REMEMBERED
                sErrorMessage = "An attempt was made to remember a device that had previously been remembered."
            Case ERROR_EXTENDED_ERROR
                sErrorMessage = "An extended error has occurred."
                ' Call the WNetGetLastError function to get a description of the error.
            Case ERROR_INVALID_PASSWORD
                sErrorMessage = "The specified network password is not correct."
            Case ERROR_NO_NET_OR_BAD_PATH
                sErrorMessage = "No network provider accepted the given network path."
            Case ERROR_NO_NETWORK
                sErrorMessage = "The network is not present or not started."
            Case ERROR_SESSION_CREDENTIAL_CONFLICT
                sErrorMessage = "The credentials supplied conflict with an existing set of credentials."
            Case Else
                sErrorMessage = "Unknown Error, return code = " & lReturnCode.ToString
        End Select

        If sErrorMessage <> "" Then System.Diagnostics.Debug.WriteLine(CStr(lReturnCode) & "   [ " & sErrorMessage & " " & Mid(sNetworkConnection, 3, 3) & " ]")

    End Function
    Friend Function bIsDriveDisconnected(ByVal lsDrive As String) As Boolean
        '**************************************************************************
        ' bIsDriveDisconnected (FUNCTION)
        '
        '  PURPOSE      : Disconnects a mapped network resource
        '  PARAMETERS   : (IN) lsDrive(String) - Drive letter to be disconnected
        '  RETURN VALUE : Returns false if an error occured
        '
        '**************************************************************************

        Dim lReturnCode As Integer

        '   Error handling
        On Error Resume Next

        bIsDriveDisconnected = False

        lReturnCode = WNetCancelConnection2(lsDrive & vbNullChar, CONNECT_UPDATE_PROFILE, 1)

        If lReturnCode = NO_ERROR Then
            bIsDriveDisconnected = True
        End If

    End Function

    Friend Function sGetIPError(ByRef lStatus As Integer) As String
        '**************************************************************************
        ' sGetIPError (FUNCTION)
        '
        '  PURPOSE      : Evaluates an error message by number
        '  PARAMETERS   : (IN) lStatus(Long) - Error number
        '  RETURN VALUE : Returns corresponding error message
        '
        '**************************************************************************

        Dim sErrorMessage As String

        '   Error handling
        On Error Resume Next

        Select Case lStatus
            Case IP_SUCCESS
                sErrorMessage = "The status was success."
            Case INADDR_NONE
                sErrorMessage = "Invalid format of IP-address"
            Case IP_BUF_TOO_SMALL
                sErrorMessage = "The reply buffer was too small."
            Case IP_DEST_NET_UNREACHABLE
                sErrorMessage = "The destination network was unreachable. In IPv6 terminology, this status value is also defined as IP_DEST_NO_ROUTE."
            Case IP_DEST_HOST_UNREACHABLE
                sErrorMessage = "The destination host was unreachable. In IPv6 terminology, this status value is also defined as IP_DEST_ADDR_UNREACHABLE."
            Case IP_DEST_PROT_UNREACHABLE
                sErrorMessage = "The destination protocol was unreachable. In IPv6 terminology, this status value is also defined as IP_DEST_PROHIBITED."
            Case IP_DEST_PORT_UNREACHABLE
                sErrorMessage = "The destination port was unreachable."
            Case IP_NO_RESOURCES
                sErrorMessage = "No IP resources were available."
            Case IP_BAD_OPTION
                sErrorMessage = "A bad IP option was specified."
            Case IP_HW_ERROR
                sErrorMessage = "A hardware error occurred."
            Case IP_PACKET_TOO_BIG
                sErrorMessage = "The packet was too big."
            Case IP_REQ_TIMED_OUT
                sErrorMessage = "The request timed out."
            Case IP_BAD_REQ
                sErrorMessage = "A bad request."
            Case IP_BAD_ROUTE
                sErrorMessage = "A bad route."
            Case IP_TTL_EXPIRED_TRANSIT
                sErrorMessage = "The Time to Live (IPv4) or Hop Limit (IPv6) expired in transit. In IPv6 terminology, this status value is also defined as IP_HOP_LIMIT_EXCEEDED."
            Case IP_TTL_EXPIRED_REASSEM
                sErrorMessage = "The Time to Live (IPv4) or Hop Limit (IPv6) expired during fragment reassembly. In IPv6 terminology, this status value is also defined as IP_REASSEMBLY_TIME_EXCEEDED."
            Case IP_PARAM_PROBLEM
                sErrorMessage = "A parameter problem. In IPv6 terminology, this status value is also defined as IP_PARAMETER_PROBLEM."
            Case IP_SOURCE_QUENCH
                sErrorMessage = "Datagrams are arriving too fast to be processed and datagrams may have been discarded."
            Case IP_OPTION_TOO_BIG
                sErrorMessage = "An IP option was too big."
            Case IP_BAD_DESTINATION
                sErrorMessage = "A bad destination."
            Case IP_ADDR_DELETED
                sErrorMessage = "vbnet.mvps.org: ip addr deleted"
            Case IP_SPEC_MTU_CHANGE
                sErrorMessage = "ip spec mtu change"
            Case IP_MTU_CHANGE
                sErrorMessage = "ip mtu_change"
            Case IP_UNLOAD
                sErrorMessage = "ip unload"
            Case IP_ADDR_ADDED
                sErrorMessage = "ip addr added"
            Case IP_NEGOTIATING_IPSEC
                sErrorMessage = "Negotiating IPSEC"
            Case IP_GENERAL_FAILURE
                sErrorMessage = "ip general failure"
            Case IP_PENDING
                sErrorMessage = "ip pending"
            Case PING_TIMEOUT
                sErrorMessage = "ping timeout"
            Case Else
                sErrorMessage = "Unknown Error, return code = " & lStatus.ToString
        End Select

        sGetIPError = CStr(lStatus) & "   [ " & sErrorMessage & " ]"

    End Function
    Private Function sFreeDriveLetter() As String
        '**************************************************************************
        ' sFreeDriveLetter (FUNCTION)
        '
        '  PURPOSE      : Evaluates the next free drive letter in alphabetical order
        '  PARAMETERS   : -
        '  RETURN VALUE : Next free drive letter
        '
        '**************************************************************************

        Dim lDriveCounter As Integer
        Dim lDriveType As Integer

        '   Error handling
        On Error Resume Next

        '   Local Init
        lDriveCounter = 64

        Do Until lDriveCounter > 67 And lDriveType = 1
            lDriveCounter = lDriveCounter + 1
            lDriveType = GetDriveType(Chr(lDriveCounter) & ":\")
        Loop

        sFreeDriveLetter = Chr(lDriveCounter) & ":"

    End Function
#End Region
End Module

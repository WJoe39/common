Option Strict Off
Option Explicit On 
Module cmnGlobal
#Region "Head"
    '******************************************************************************
    '*  MODULE        : cmnGlobal
    '*  FILE          : cmnGlobal.vb
    '*  PROJECT       : non-specific
    '*  AUTHOR        : Chris A. Lutz
    '*  CREATED       : 14-May-2003
    '*  COPYRIGHT     : Copyright (c) 2003 Chris A. Lutz. All Rights Reserved.
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
    '*  DESCRIPTION   : Common general operations
    '*
    '*  MODIFICATION HISTORY:
    '*  AUTHOR:         DATE:       CHANGES:
    '*  Chris A. Lutz   14-May-2003 Initial Version
    '*
    '******************************************************************************
#End Region
#Region "Declaration"
    '-------------- constants definition ------------------------------------------
    Private Const VB_MODULE As String = "cmnGlobal"

    Private Const FORMAT_MESSAGE_FROM_SYSTEM As Integer = &H1000S
    Private Const FORMAT_MESSAGE_IGNORE_INSERTS As Integer = &H200S

    '-------------- libraries -----------------------------------------------------
    Private Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Integer, ByRef lpSource As IntPtr, ByVal dwMessageId As Integer, ByVal dwLanguageId As Integer, ByVal lpBuffer As String, ByVal nSize As Integer, ByRef Arguments As IntPtr) As Integer
    'Private Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Integer, ByVal lpSource As IntPtr, ByVal dwMessageId As Integer, ByVal dwLanguageId As Integer, ByVal lpBuffer As System.Text.StringBuilder, ByVal nSize As Integer, ByVal Arguments As IntPtr) As Integer

    '-------------- symbol definition ---------------------------------------------
    '   country-specific
    Private g_sLanguage As String
#End Region
#Region "Properties"
#End Region
#Region "Methods"
    '-------------- procedure & function definition--------------------------------
    Friend Sub Debug(ByRef sContext As String _
                   , ByRef sDump As String _
                   , Optional ByRef sProc As String = "")
        'TODO define debug proc
    End Sub
    Friend Sub RaiseErr(ByRef sContext As String _
                      , ByRef sDump As String _
                      , Optional ByRef sProc As String = "")
        'TODO define debug proc
    End Sub
    Friend Function sGetAPIError(ByVal LastDLLError As Integer) As String
        '******************************************************************************
        ' sGetAPIError (FUNCTION)
        '
        '  PURPOSE      : Evaluates an error message by number
        '  PARAMETERS   : (IN) LastDLLError(Long) - Error number
        '  RETURN VALUE : Returns corresponding error message
        '
        '******************************************************************************

        Dim sErrorMessage As String
        Dim lReturnCode As Integer
        Dim objNull As IntPtr

        '   Error handling
        On Error Resume Next

        sErrorMessage = New String(Chr(0), 256)
        lReturnCode = FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM Or FORMAT_MESSAGE_IGNORE_INSERTS, objNull, LastDLLError, 0, sErrorMessage, Len(sErrorMessage), objNull)
        If lReturnCode Then sGetAPIError = CStr(LastDLLError) & "   [ " & Left(sErrorMessage, lReturnCode - 1) & " ]"

    End Function
#End Region
End Module
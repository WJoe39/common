Option Strict Off
Option Explicit On 
Module cmnFile
#Region "Head"
    '******************************************************************************
    '*  MODULE        : cmnFile
    '*  FILE          : cmnFile.vb
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
    '*  DESCRIPTION   : Common file operations
    '*
    '*  MODIFICATION HISTORY:
    '*  AUTHOR:             DATE:       CHANGES:
    '*  Christoph A. Lutz   01-Apr-2007 Initial Version
    '*
    '******************************************************************************
#End Region
#Region "Declaration"
    '-------------- constants definition ------------------------------------------
    Private Const VB_MODULE As String = "cmnFile"

    Private Const MAX_PATH As Short = 260
    Private Const SHGFI_DISPLAYNAME As Short = &H200S

    '-------------- libraries -----------------------------------------------------
    Friend Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Integer
    Friend Declare Function CopyFile Lib "kernel32" Alias "CopyFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal bFailIfExists As Integer) As Integer

    Private Declare Function SHGetFileInfo Lib "shell32" Alias "SHGetFileInfoA" (ByVal pszPath As String, ByVal dwFileAttributes As Integer, ByRef psfi As SHFILEINFO, ByVal cbFileInfo As Integer, ByVal uFlags As Integer) As Integer
    Private Declare Function CreateDirectory Lib "kernel32" Alias "CreateDirectoryA" (ByVal lpPathName As String, ByRef lpSecurityAttributes As SECURITY_ATTRIBUTES) As Integer

    '-------------- types definition ----------------------------------------------
    Private Structure SHFILEINFO
        Dim hIcon As Integer
        Dim iIcon As Integer
        Dim dwAttributes As Integer
        <VBFixedString(MAX_PATH), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValTStr, SizeConst:=MAX_PATH)> Public szDisplayName As String
        <VBFixedString(80), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValTStr, SizeConst:=80)> Public szTypeName As String
    End Structure
    Private Structure SECURITY_ATTRIBUTES
        Dim nLength As Integer
        Dim pSecurityDescriptor As Integer
        Dim bInheritHandle As Integer
    End Structure
#End Region
#Region "Properties"
#End Region
#Region "Methods"
    '-------------- procedure & function definition--------------------------------
    Friend Function bDoesFileExist(ByRef lsFile As String) As Boolean
        'TODO copperplate

        Dim shfi As SHFILEINFO

        '**********************************************   Fehler ausschalten
        On Error Resume Next

        If SHGetFileInfo(lsFile, 0, shfi, Len(shfi), SHGFI_DISPLAYNAME) Then
            Return True
        Else
            Return False
        End If

    End Function
    Friend Function bHasSameTreeStruct(ByRef lsFolderIn As Object, ByRef lsHostShare As Object) As Boolean
        'TODO copperplate

        Dim lsFolderArray() As String
        Dim lsNextFolder As String
        Dim lvSubFolder As Object
        Dim SA As SECURITY_ATTRIBUTES
        Dim llCreateSuccess As Integer

        '**********************************************   Fehler ausschalten
        On Error Resume Next

        'UPGRADE_WARNING: Die Standardeigenschaft des Objekts lsFolderIn konnte nicht aufgelöst werden. Klicken Sie hier für weitere Informationen: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
        lsFolderArray = Split(lsFolderIn, "\", -1, CompareMethod.Text)
        'UPGRADE_WARNING: Die Standardeigenschaft des Objekts lsHostShare konnte nicht aufgelöst werden. Klicken Sie hier für weitere Informationen: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
        lsNextFolder = Left(lsHostShare, Len(lsHostShare) - 1)

        For Each lvSubFolder In lsFolderArray
            'UPGRADE_WARNING: Die Standardeigenschaft des Objekts lvSubFolder konnte nicht aufgelöst werden. Klicken Sie hier für weitere Informationen: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
            lsNextFolder = lsNextFolder & "\" & lvSubFolder
            If Not bDoesFileExist(lsNextFolder) Then
                llCreateSuccess = CreateDirectory(lsNextFolder, SA)
            End If
        Next lvSubFolder
        If llCreateSuccess = 0 Then
            Return False
        Else
            Return True
        End If

    End Function
#End Region
End Module
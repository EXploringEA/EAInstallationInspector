' Copyright (C) 2015 - 2022 Adrian LINCOLN, EXploringEA - All Rights Reserved
'
'   This program is free software: you can redistribute it and/or modify it under the terms of the GNU General Public License as published by
'   the Free Software Foundation, either version 3 of the License, or (at your option) any later version.

'   This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of
'   MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU General Public License for more details.
'
'    You should have received a copy of the GNU General Public License along with this program.  If not, see <http://www.gnu.org/licenses/>.
' =============================================================================================================================================

Module Constants




    ' LOCATION OF EA program keys
    ' 32-bit
    Friend Const EAHKCU32 As String = "HKEY_CURRENT_USER\Software\Sparx Systems\EA400\EA"
    Friend Const EAHKLM32 As String = "HKEY_LOCAL_MACHINE\Software\Sparx Systems\EA400\EA"
    ' 64-bit
    Friend Const EAHKCU64 As String = "HKEY_CURRENT_USER\Software\Sparx Systems\EA64\EA"
    Friend Const EAHKLM64 As String = "HKEY_LOCAL_MACHINE\Software\Sparx Systems\EA64\EA"


    ''' <summary>
    ''' Text not set
    ''' </summary>
    Friend Const cNotSet As String = "NOT SET"
    Friend Const cNotFound As String = "Not found"

    Friend Const cLogFile As String = "log file "

    Friend Const cFilePrefix As String = "file:///"
    Friend Const fileprefixlength As Integer = 8

    Friend cCodeBase As String = "CodeBase"




    Friend Const cCLSID As String = "CLSID"
    '   Friend Const cProgID As String = "ProgID"

    ' Strings that are used
    Friend Const cBackSlash As String = "\"
    Friend Const cForwardSlash As String = "/"
    Friend Const cInprocServer32 As String = "InprocServer32"
    Friend Const cWow6432Node As String = "Wow6432Node"



End Module

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
    Friend cAssembly As String = "Assembly"
    Friend cClass As String = "Class"
    Friend cRuntimeVersion As String = "RuntimeVersion"

    Friend Const cHKCUCLSID As String = "HKCU : CLSID = "
    Friend Const cHKCUClassname As String = "HKCU : Classname = "
    Friend Const cHKCUFilename As String = "HKCU : Filename = "
    Friend Const cHKCUVersion As String = "HKCU : Version = "
    Friend Const cHKCURuntimeVersion As String = "HKCU : Runtime version = "
    Friend Const cHKCUProgID As String = "HKCU : ProgId = "

    Friend Const cHKLMCLSID As String = "HKLM : CLSID = "
    Friend Const cHKLMClassname As String = "HKLM : Classname = "
    Friend Const cHKLMFilename As String = "HKLM : Filename = "
    Friend Const cHKLMVersion As String = "HKLM : Version = "
    Friend Const cHKLMRuntimeVersion As String = "HKLM : Runtime version = "
    Friend Const cHKLMProgID As String = "HKLM : ProgId = "

    Friend Const cHKLMWowCLSID As String = "HKLMWow : CLSID = "
    Friend Const cHKLMWowClassname As String = "HKLMWow : Classname = "
    Friend Const cHKLMWowFilename As String = "HKLMWow : Filename = "
    Friend Const cHKLMWowVersion As String = "HKLMWow : Version = "
    Friend Const cHKLMWowRuntimeVersion As String = "HKLMWow : Runtime version = "
    Friend Const cHKLMWowProgID As String = "HKLMWow : ProgId = "




    Friend Const cCLSID As String = "CLSID"
    Friend Const cProgID As String = "ProgID"







    ' Strings that are used
    Friend Const cBackSlash As String = "\"
    Friend Const cForwardSlash As String = "/"
    Friend Const cInprocServer32 As String = "InprocServer32"
    Friend Const cWow6432Node As String = "Wow6432Node"

    ' colours used for output
    Friend cQueryRowColorDefault As Color = Color.White
    Friend cQueryRowColorNew As Color = Color.Yellow
    Friend cQueryRowColorComplete As Color = Color.SpringGreen
    Friend cQueryRowWarning As Color = Color.Orange


End Module

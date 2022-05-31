' Copyright (C) 2015 - 2018 Adrian LINCOLN, EXploringEA - All Rights Reserved
'
'   This program is free software: you can redistribute it and/or modify it under the terms of the GNU General Public License as published by
'   the Free Software Foundation, either version 3 of the License, or (at your option) any later version.

'   This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of
'   MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU General Public License for more details.
'
'    You should have received a copy of the GNU General Public License along with this program.  If not, see <http://www.gnu.org/licenses/>.
' =============================================================================================================================================

Module Constants

    ''' <summary>
    ''' The sparx keys - 32-bit 
    ''' </summary>
    Friend Const SparxKeys32 As String = "Software\Sparx Systems\EAAddins"
    ''' <summary>
    ''' The sparx keys - 32-bit addin for 64-bit OS
    ''' </summary>
    Friend Const WSparxKeys32 As String = "Software\Wow6432Node\Sparx Systems\EAAddins"

    ''' <summary>
    ''' HKCU Keys
    ''' </summary>
    Friend Const HKCUfullKey32 As String = "HKEY_CURRENT_USER\" & SparxKeys32
    ''' <summary>
    ''' HKLM Keys
    ''' </summary>
    Friend Const HKLMfullKey32 As String = "HKEY_LOCAL_MACHINE\" & WSparxKeys32

    ''' <summary>
    ''' The sparx keys - x64
    ''' </summary>
    Friend Const SparxKeys64 As String = "Software\Sparx Systems\EAAddins64"
    ''' <summary>
    ''' HKCU Keys
    ''' </summary>6
    Friend Const HKCUfullKey64 As String = "HKEY_CURRENT_USER\" & SparxKeys64
    ''' <summary>
    ''' HKLM Keys
    ''' </summary>
    Friend Const HKLMfullKey64 As String = "HKEY_LOCAL_MACHINE\" & SparxKeys64

    ' Registry locations

    ''' <summary>
    ''' HKCR Classes
    ''' </summary>
    Friend Const HKCUClasses As String = "HKEY_CURRENT_USER\SOFTWARE\Classes"

    ''' <summary>
    ''' HKLM Classes
    ''' </summary>
    Friend Const HKLMClasses As String = "HKEY_LOCAL_MACHINE\SOFTWARE\Classes"

    Friend Const HKLMWowClasses As String = "HKEY_LOCAL_MACHINE\SOFTWARE\Classes\WOW6432Node"




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


    'Registry related strings

    Friend Const cHKCU As String = "HKCU"
    Friend Const cHKLM As String = "HKLM"

    Friend Const cHKCUWow As String = "HKCUWow"
    Friend Const cHKLMWow As String = "HKLMWow"

    Friend Const cHKCU64 As String = "HKCU64"
    Friend Const cHKLM64 As String = "HKLM64"


    Friend Const cCLSID As String = "CLSID"
    Friend Const cProgID As String = "ProgID"

    ' LOCATION Sparx AddIns keys
    ' 32-bit


    Friend cAddins32 As String = "SOFTWARE\Sparx Systems\EAAddins"
    Friend cWowAddins32 As String = "SOFTWARE\WOW6432Node\Sparx Systems\EAAddins"

    Friend Const cHKCUAddins32 As String = "HKCU\Software\Sparx Systems\EAAddins"
    Friend Const cHKLMAddins32 As String = "HLKM\SOFTWARE\Sparx Systems\EAAddins"
    Friend Const cHKLMWowAddins32 As String = "HLKM\SOFTWARE\WOW6432Node\Sparx Systems\EAAddins"

    Friend Const cHKCUSparxAddinKeys32 As String = "HKEY_CURRENT_USER\Software\Sparx Systems\EAAddins"
    Friend Const cHKLMSparxAddinKeys32 As String = "HKEY_LOCAL_MACHINE\Software\Sparx Systems\EAAddins"
    Friend Const cHKLMWowSparxAddinKeys32 As String = "HKEY_LOCAL_MACHINE\Software\Wow6432Node\Sparx Systems\EAAddins"

    ' 64-bit

    Friend Const cAddins64 As String = "SOFTWARE\Sparx Systems\EAAddins64"
    Friend Const cHKCUAddins64 As String = "HKCU\Software\Sparx Systems\EAAddins64"
    Friend Const cHKLMAddins64 As String = "HLKM\SOFTWARE\Sparx Systems\EAAddins64"
    ' Keys
    Friend Const cHKCUSparxAddinKeys64 As String = "HKEY_CURRENT_USER\Software\Sparx Systems\EAAddins64"
    Friend Const cHKLMSparxAddinKeys64 As String = "HKEY_LOCAL_MACHINE\Software\Sparx Systems\EAAddins64"

    ' LOCATION OF EA program keys
    ' 32-bit
    Friend Const EAHKCU32 As String = "HKEY_CURRENT_USER\Software\Sparx Systems\EA400\EA"
    Friend Const EAHKLM32 As String = "HKEY_LOCAL_MACHINE\Software\Sparx Systems\EA400\EA"
    ' 64-bit
    Friend Const EAHKCU64 As String = "HKEY_CURRENT_USER\Software\Sparx Systems\EA64\EA"
    Friend Const EAHKLM64 As String = "HKEY_LOCAL_MACHINE\Software\Sparx Systems\EA64\EA"


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

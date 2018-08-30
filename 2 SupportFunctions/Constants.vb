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
    ''' Root to look in HKCU
    ''' </summary>
    Friend Const EA As String = "HKEY_CURRENT_USER\Software\Sparx Systems\EA400\EA"


    ''' <summary>
    ''' The sparx keys - x86
    ''' </summary>
    Friend Const SparxKeys As String = "Software\Sparx Systems\EAAddins"
    ''' <summary>
    ''' The sparx keys - x64
    ''' </summary>
    Friend Const WSparxKeys As String = "Software\Wow6432Node\Sparx Systems\EAAddins"

    ''' <summary>
    ''' HKCU Keys
    ''' </summary>
    Friend Const HKCUfullKey As String = "HKEY_CURRENT_USER\" & SparxKeys
    ''' <summary>
    ''' HKLM Keys
    ''' </summary>
    Friend Const HKLMfullKey As String = "HKEY_LOCAL_MACHINE\" & WSparxKeys

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



    Friend Const cHKCU As String = "HKCU"
    Friend Const cHKLM As String = "HKLM"
    Friend Const cHKLMWow As String = "HKLMWow"
    Friend Const cCLSID As String = "CLSID"
    Friend Const cProgID As String = "ProgID"

    Friend cAddins As String = "SOFTWARE\Sparx Systems\EAAddins"
    Friend cWowAddins As String = "SOFTWARE\WOW6432Node\Sparx Systems\EAAddins"

    Friend Const cHKCUAddins As String = "HKCU\Software\Sparx Systems\EAAddins"
    Friend Const cHKLMAddins As String = "HLKM\SOFTWARE\Sparx Systems\EAAddins"
    Friend Const cHKLMWowAddins As String = "HLKM\SOFTWARE\WOW6432Node\Sparx Systems\EAAddins"

    ' Keys
    Friend Const cHKCUSparxAddinKeys As String = "HKEY_CURRENT_USER\Software\Sparx Systems\EAAddins"
    Friend Const cHKLMSparxAddinKeys As String = "HKEY_LOCAL_MACHINE\Software\Sparx Systems\EAAddins"
    Friend Const cHKLMWowSparxAddinKeys As String = "HKEY_LOCAL_MACHINE\Software\Wow6432Node\Sparx Systems\EAAddins"

    Friend Const cBackSlash As String = "\"
    Friend Const cForwardSlash As String = "/"
    Friend Const cInprocServer32 As String = "InprocServer32"
    Friend Const cWow6432Node As String = "Wow6432Node"

    Friend cQueryRowColorDefault As Color = Color.White
    Friend cQueryRowColorNew As Color = Color.Yellow
    Friend cQueryRowColorComplete As Color = Color.SpringGreen
    Friend cQueryRowWarning As Color = Color.Orange
End Module

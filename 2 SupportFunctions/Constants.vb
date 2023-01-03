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

    ''' <summary>
    ''' Text not set
    ''' </summary>
    Friend Const cNotSet As String = "NOT SET"
    Friend Const cNotFound As String = "Not found"
    Friend Const cFilePrefix As String = "file:///"
    Friend Const fileprefixlength As Integer = 8
    Friend cCodeBase As String = "CodeBase"
    Friend Const cCLSID As String = "CLSID"
    '   Friend Const cProgID As String = "ProgID"

    ' Strings that are used
    Friend Const cBackSlash As String = "\"
    Friend Const cForwardSlash As String = "/"

#Region "Usedin AddIn information and treeviewer"

#Region "Sparx keys"


    ' Sparx subkey folders - which exist in both HKCU and HKLM
    ''' <summary>
    ''' The sparx keys - 32-bit  
    ''' </summary>
    Friend Const cSparxKeys32 As String = "Software\Sparx Systems\EAAddins"
    ''' <summary>
    ''' The sparx keys - 32-bit addin for 64-bit OS
    ''' </summary>
    Friend Const cSparxKeysWOW32 As String = "Software\Wow6432Node\Sparx Systems\EAAddins"

    ' 64-bit AddIns

    ''' <summary>
    ''' The sparx keys - x64
    ''' </summary>
    Friend Const cSparxKeys64 As String = "Software\Sparx Systems\EAAddins64"

    ' Exact registry locations for Sparx keys
    ' 32-bit AddIns

    ''' <summary>
    ''' HKCU Keys
    ''' </summary>
    Friend Const ceaHKCU32AddInKeys As String = "HKEY_CURRENT_USER\" & cSparxKeys32

    ''' <summary>
    ''' HKLM Keys
    ''' </summary>
    Friend Const ceaHKLM32AddInKey64 As String = "HKEY_LOCAL_MACHINE\" & cSparxKeysWOW32
    Friend Const ceaHKLM32AddInKeys As String = "HKEY_LOCAL_MACHINE\" & cSparxKeys32


    ' 64-bit AddIns
    ''' <summary>
    ''' HKCU Keys
    ''' </summary>6
    Friend Const ceaHKCU64AddInKeys As String = "HKEY_CURRENT_USER\" & cSparxKeys64
    ''' <summary>
    ''' HKLM Keys
    ''' </summary>
    Friend Const ceaHKLM64AddInKeys As String = "HKEY_LOCAL_MACHINE\" & cSparxKeys64
#End Region






    Sub constantcheck()

        '  ccheck("cHKCU_Root", cHKCU_Root)
        ' ccheck("cHKCU_Software", cHKCU_Software)
        '    ccheck("cHKCU_Classes", cHKCU_Classes)
        '   ccheck("cHKCU_ClassesCLSID", cHKCU_ClassesCLSID)
        '    ccheck("cHKCUWOW_ClassesCLSID", cHKCUWOW_ClassesCLSID)
        '    ccheck("cHKCUWOW_Classes", cHKCUWOW_Classes)
        ' ccheck("cHKLM_Root", cHKLM_Root)
        '   ccheck("cHKLM_Software", cHKLM_Software)
        '  ccheck("cHKLM_Classes", cHKLM_Classes)
        '    ccheck("cHKLM_ClassesCLSID", cHKLM_ClassesCLSID)
        '   ccheck("cHKLMWow_Classes", cHKLMWow_Classes)
        '    ccheck("cHKLMWow_ClassesCLSID", cHKLMWow_ClassesCLSID)


    End Sub

    'Sub ccheck(p1 As String, p2 As String)
    '    Debug.Print("Constant: " & p1 & " = " & p2)

    'End Sub

#End Region
End Module

' Copyright (C) 2015 - 2017 Adrian LINCOLN, EXploringEA - All Rights Reserved
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
    Friend Const HKLMfullKey As String = "HKEY_LOCAL_MACHINE\" & SparxKeys

    ''' <summary>
    ''' HKCR Classes
    ''' </summary>
    Friend Const HKCRClasses As String = "HKEY_CURRENT_USER\SOFTWARE\Classes"
    ''' <summary>
    ''' HKLM Classes
    ''' </summary>
    Friend Const HKLMClasses As String = "HKEY_LOCAL_MACHINE\SOFTWARE\Classes"
    ''' <summary>
    ''' Text not set
    ''' </summary>
    Friend Const cNotSet As String = "NOT SET"


End Module

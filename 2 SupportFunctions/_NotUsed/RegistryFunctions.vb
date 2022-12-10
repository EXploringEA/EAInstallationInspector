' Copyright (C) 2022 Adrian LINCOLN, EXploringEA - All Rights Reserved
'
'   This program is free software: you can redistribute it and/or modify it under the terms of the GNU General Public License as published by
'   the Free Software Foundation, either version 3 of the License, or (at your option) any later version.

'   This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of
'   MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU General Public License for more details.
'
'    You should have received a copy of the GNU General Public License along with this program.  If not, see <http://www.gnu.org/licenses/>.
' =============================================================================================================================================
'


Imports Microsoft.Win32

Public Class RegistryFunctions


    '' This function will return the ClassID
    ''
    'Function returnClassID(pHive As String, pClassName As String) As KeyEventArgs

    '    Select Case pHive
    '        Case AddInEntry.cHKCU32
    '        Case AddInEntry.cHKCU64
    '        Case AddInEntry.cHKLM32
    '        Case AddInEntry.cHKLM64
    '    End Select

    '    '_location = "HKEY_CLASSES_ROOT\Wow6432Node" & cBackSlash & cCLSID & cBackSlash & ClassID & cBackSlash & cInprocServer32
    '    'Filename = Registry.GetValue(_location, cCodeBase, cNotSet) 'using the class try to find the DLL path
    '    'If Filename IsNot Nothing Then DLLSource = AddInEntry.cHKCR32

    'End Function

    ' NOT USED
    ''' <summary>
    ''' Search and Find Registry Function
    ''' </summary>
    Private Function SearchRegistry(ByVal location As String) As String

        'Open the HKEY_CLASSES_ROOT\CLSID which contains the list of all registered COM files (.ocx,.dll, .ax) 
        'on the system no matters if is 32 or 64 bits.
        Dim t_clsidKey As RegistryKey = Registry.ClassesRoot.OpenSubKey("CLSID")

        'Get all the sub keys it contains, wich are the generated GUID of each COM.
        '   For Each subKey In t_clsidKey.GetSubKeyNames.ToList

        Dim t_clsidSubKey As RegistryKey = Registry.ClassesRoot.OpenSubKey(location) '"CLSID\" & subKey & "\InProcServer32")

        If Not t_clsidSubKey Is Nothing Then
            'in the case InProcServer32 exist we get the default value wich contains the path of the COM file.
            Dim t_valueName As String = (From value In t_clsidSubKey.GetValueNames() Where value = "")(0).ToString

            'Now gets the value.
            Dim t_value As String = t_clsidSubKey.GetValue(t_valueName).ToString

            'And finnaly if the value ends with the name of the dll (include .dll) we return it
            If t_value.EndsWith(".dll") Then

                Return t_value

            End If

        End If



        'if not exist, return nothing
        Return Nothing

    End Function


End Class

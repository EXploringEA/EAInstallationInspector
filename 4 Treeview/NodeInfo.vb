' Copyright (C) 2015 - 2018 Adrian LINCOLN, EXploringEA - All Rights Reserved
'
'   This program is free software: you can redistribute it and/or modify it under the terms of the GNU General Public License as published by
'   the Free Software Foundation, either version 3 of the License, or (at your option) any later version.

'   This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of
'   MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU General Public License for more details.
'
'    You should have received a copy of the GNU General Public License along with this program.  If not, see <http://www.gnu.org/licenses/>.
' =============================================================================================================================================

' this class defines the information that will be stored in each node
' need to review in relation to registry entries as it may be that we have different types of information
' so what is the structure 
Friend Class NodeInfo

    Public Name As String = "NODE>>>"
    Public NodeType As NodeType ' different types of node are possible use this to define the type from which some of the attributes will have no meaning
    Public SparxEntryLocation As String = "" ' HKCU, HKLM, HKLMWow
    Public ClassName As String = "" ' e.g. fred.fred
    Public ClassNameLocation As String = "" ' e.g. HLKM\Software\Sparx Systems 
    Public CLSIDSrc As String = "" ' Source of ClassID HKCU, HKLM, HKLMWow
    Public CLSID As String = "" ' class ID for the class name
    Public CLSIDLocation As String = "" ' where the CLSID was defined
    Public childCount As Integer = 0 ' number of children for current node
    Public Filename As String = "" ' the filename inc ext
    Public FilePathName As String = "" ' filename inc fullpath
    Public ProgID As String = ""


    Sub New(pType As NodeType, pChildCount As Integer)
        NodeType = pType
        childCount = pChildCount
    End Sub

End Class

Friend Enum NodeType
    SparxRoot
    HIVE
    ClassNameNode
    CLSIDNode
    CLSIDNode_HIVE
    CLSIDNode_ClassName
    CLSIDNode_Assembly
    CLSIDNode_CodeBase
    CLSIDNode_RunTimeVersion
    CLSIDNode_ProgID
    CLSIDNode_Version
End Enum



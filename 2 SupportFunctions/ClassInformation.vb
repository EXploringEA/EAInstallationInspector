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

Imports System.Reflection
Imports Microsoft.Win32


Public Class ClassInformation


    ' Registry class locations
    ''' <summary>
    ''' HKCU Classes for 
    ''' * 32-bit on 32-bit OS
    ''' * 64-bit on 64-bit OS
    ''' </summary>
    Friend Const HKCU_Classes As String = "HKEY_CURRENT_USER\SOFTWARE\Classes"

    '''' <summary>
    '''' HKCU Classes for 
    '''' * 32-bit on 64-bit OS
    '''' </summary>
    'Friend Const HKCUWOW_Classes As String = "HKEY_CURRENT_USER\SOFTWARE\Classes\Wow6432Node"

    ''' <summary>
    ''' HKLM Classes
    ''' * 32-bit on 32-bit OS
    ''' * 64-bit on 64-bit OS
    ''' </summary>
    Friend Const HKLM_Classes As String = "HKEY_LOCAL_MACHINE\SOFTWARE\Classes"


    ''' <summary>
    ''' HKLM Classes for 
    ''' * 32-bit on 64-bit OS
    ''' </summary>
    Friend Const HKLMWow1_Classes As String = "HKEY_LOCAL_MACHINE\SOFTWARE\Classes\WOW6432Node"
    Friend Const HKLMWow2_Classes As String = "HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Classes"


    Property ClassSource As String = ""
    Property ClassID As String = ""
    Property DLLSource As String = ""
    Property Version As String = ""
    Property Filename As String = ""
    Property Colour As Color = Color.Magenta



    Friend Sub GetClassInformation(pAddInName As String, pSequence As String())

        getClassID(pAddInName, pSequence)

        If ClassID <> "" Then getClassDLLFilename()

        ' If we find no Class found Then the DLL cannot be found
        ' _ClassSource indicates where the class entry is defined, hence the HIVE where we would expect to find class details
        ' NOTE: this may not always be the case, we have seen cases where the library is put in the "wrong place"
        ' Also depending on whether it is a 32-bit or 64-bit machine will determine where the DLL detail is located
    End Sub



    Private Sub getClassID(pAddInName As String, s As String())

        ClassSource = ""
        ClassID = ""

        For Each loc As String In s
            Select Case loc
                ' 32-bit OS
                Case AddInEntry.cHKCU32 ' 32-bit on 32-bit OS
                    ClassID = Registry.GetValue(HKCU_Classes & cBackSlash & pAddInName & cBackSlash & cCLSID, "", cNotFound)
                    ClassSource = AddInEntry.cHKCU32
                    If ClassID <> "" Then Return

                Case AddInEntry.cHKLM32 ' 32-bit on 32-bit OS
                    ClassSource = AddInEntry.cHKLM32
                    ClassID = Registry.GetValue(HKLM_Classes & cBackSlash & pAddInName & cBackSlash & cCLSID, "", cNotFound)
                    If ClassID <> "" Then Return

                    ' 64-bit OS
                Case AddInEntry.cHKLM32Wow ' 32-bit on 64-bit OS
                    ' There are 2 locations which can be inspected
                    ClassID = Registry.GetValue(HKLMWow1_Classes & cBackSlash & pAddInName & cBackSlash & cCLSID, "", cNotFound)
                    If ClassID <> "" Then
                        ClassSource = AddInEntry.cHKLM32Wow
                        Return
                    End If
                    ClassID = Registry.GetValue(HKLMWow2_Classes & cBackSlash & pAddInName & cBackSlash & cCLSID, "", cNotFound)
                    If ClassID <> "" Then
                        ClassSource = AddInEntry.cHKLM32Wow
                        Return
                    End If
                Case AddInEntry.cHKCU64
                    ClassSource = AddInEntry.cHKCU64
                    ClassID = Registry.GetValue(HKCU_Classes & cBackSlash & pAddInName & cBackSlash & cCLSID, "", cNotFound)
                    If ClassID <> "" Then Return
                Case AddInEntry.cHKLM64
                    ClassSource = AddInEntry.cHKLM64
                    ClassID = Registry.GetValue(HKLM_Classes & cBackSlash & pAddInName & cBackSlash & cCLSID, "", cNotFound)
                    If ClassID <> "" Then Return
                Case Else
                    Debug.Print("Unknown Addin string " & loc)
            End Select

        Next
        Return
    End Sub

    ''' <summary>
    ''' Get the DLL filename which we assume is in the relevant place for HKCU / HKLM - the hive is provided by the caller
    ''' </summary>
    ''' <remarks>There are some instance where the file would not be found as expected - 
    ''' e.g. the COMServer class but don't think this is a likely case have included some checks</remarks>
    Private Sub getClassDLLFilename()

        Try
            ' 1. Use the Select the HIVE where the classID was found and get the information
            Dim _location As String = ""
            DLLSource = cNotFound
            Select Case ClassSource
                ' 32-bit OS
                Case AddInEntry.cHKCU32
                    _location = HKCU_Classes & cBackSlash & cCLSID & cBackSlash & ClassID & cBackSlash & cInprocServer32
                    Filename = Registry.GetValue(_location, cCodeBase, cNotSet) 'using the class try to find the DLL path

                    If Filename IsNot Nothing Then DLLSource = AddInEntry.cHKCU32

                Case AddInEntry.cHKLM32 ' 32-bit on 32-bit OS
                    _location = HKLM_Classes & cBackSlash & cWow6432Node & cBackSlash & cCLSID & cBackSlash & ClassID & cBackSlash & cInprocServer32
                    Filename = Registry.GetValue(_location, cCodeBase, cNotSet) 'using the class try to find the DLL path
                    If Filename IsNot Nothing Then DLLSource = AddInEntry.cHKLM32


                Case AddInEntry.cHKLM32Wow ' 32-bit AddIn on 64-bit OS
                    _location = HKLM_Classes & cBackSlash & cWow6432Node & cBackSlash & cCLSID & cBackSlash & ClassID & cBackSlash & cInprocServer32
                    Filename = Registry.GetValue(_location, cCodeBase, cNotSet) 'using the class try to find the DLL path
                    If Filename IsNot Nothing Then DLLSource = AddInEntry.cHKLM32Wow


                Case AddInEntry.cHKCU64
                    _location = HKCU_Classes & cBackSlash & cCLSID & cBackSlash & ClassID & cBackSlash & cInprocServer32
                    Filename = Registry.GetValue(_location, cCodeBase, cNotSet) 'using the class try to find the DLL path
                    If Filename IsNot Nothing Then
                        DLLSource = AddInEntry.cHKCU64
                    Else
                        _location = HKLM_Classes & cBackSlash & cCLSID & cBackSlash & ClassID & cBackSlash & cInprocServer32
                        Filename = Registry.GetValue(_location, "", cNotSet)
                        If Filename IsNot Nothing Then DLLSource = "HKLM ROOT KEY"
                    End If

                Case AddInEntry.cHKLM64
                    _location = HKLM_Classes & cBackSlash & cCLSID & cBackSlash & ClassID & cBackSlash & cInprocServer32
                    Filename = Registry.GetValue(_location, cCodeBase, cNotSet) 'using the class try to find the DLL path
                    If Filename IsNot Nothing Then
                        DLLSource = AddInEntry.cHKLM64
                    Else
                        _location = HKLM_Classes & cBackSlash & cCLSID & cBackSlash & ClassID & cBackSlash & cInprocServer32
                        Filename = Registry.GetValue(_location, "", cNotSet)
                        If Filename IsNot Nothing Then DLLSource = "HKLM ROOT KEY"
                    End If

                Case Else
                    ' ERROR
                    Filename = "Error"
                    DLLSource = "Error"
            End Select
            If Filename <> "Error" Then getDLLAssembly()

        Catch ex As Exception
#If DEBUG Then
            Debug.Print(ex.ToString)
#End If
        End Try
        Return
    End Sub



    Private Sub getDLLAssembly()

        Dim _filename As String = Filename
        If _filename IsNot Nothing And _filename <> cNotSet Then
            If Strings.Left(_filename, fileprefixlength) = cFilePrefix Then _filename = Strings.Right(_filename, _filename.Length - fileprefixlength)
            Try
                Dim ass As AssemblyName = AssemblyName.GetAssemblyName(_filename)
                Version = ass.Version.ToString
            Catch ex As Exception
                Version = "Unable to determine"
            End Try
            If _filename <> "" Then _filename = _filename.Replace("/", "\")
            Filename = _filename
            ' now depending on where items were found flag accordingly
            If ClassSource <> DLLSource Then
                Colour = Color.Red
            ElseIf Filename = cNotSet Then
                Colour = Color.Yellow
            Else
                Colour = If(DLLexists(Filename), Color.LightGreen, Color.Cyan)  ' File does not exist
            End If

        End If

    End Sub


    ''' <summary>
    ''' Dls the lexists.
    ''' </summary>
    ''' <param name="pFilePath">DLL Filename path.</param>
    ''' <returns>True if exists else false</returns>
    Private Function DLLexists(pFilePath As String) As Boolean
        Try
            ' remove file from front of string
            If Strings.InStr(pFilePath, "file:///") Then
                Dim myNewFN As String = Strings.Right$(pFilePath, Len(pFilePath) - 8)
                If System.IO.File.Exists(myNewFN) Then Return True
            ElseIf Strings.InStr(pFilePath, ":\") Or Strings.InStr(pFilePath, ":/") Then
                If System.IO.File.Exists(pFilePath) Then Return True
            End If

        Catch ex As Exception

        End Try
        Return False
    End Function

    ' used by treeviewer
    ' function to get class information from the registry location and classID
    ' HIVE NEEDS TO be HKCU32, HKCU64
    Friend Shared Function OLDgetClassInformation(pClassLocation As String, pID As String) As ClassRegistryInformation
        Dim myClassInfo As New ClassRegistryInformation
        Try

            Dim _keylocation As String = cNotSet
            Select Case pClassLocation
                Case AddInEntry.cHKCU32
                    _keylocation = HKCU_Classes & cBackSlash & cCLSID & cBackSlash & pID

                Case AddInEntry.cHKLM32
                    _keylocation = HKLM_Classes & cBackSlash & cCLSID & cBackSlash & pID

                Case AddInEntry.cHKLM32Wow
                    _keylocation = HKLMWow2_Classes & cBackSlash & cCLSID & cBackSlash & pID
                    If _keylocation = "" Then _keylocation = HKLMWow1_Classes & cBackSlash & cWow6432Node & cBackSlash & cCLSID & cBackSlash & pID

                Case AddInEntry.cHKCU64
                    _keylocation = HKCU_Classes & cBackSlash & cCLSID & cBackSlash & pID
                Case AddInEntry.cHKLM64
                    _keylocation = HKLM_Classes & cBackSlash & cCLSID & cBackSlash & pID
                Case Else
                    Debug.Print("HIVE - " & pClassLocation)
            End Select
            myClassInfo.HIVE = pClassLocation
            If _keylocation = cNotSet Then
                myClassInfo.CodeBase = cNotSet
                myClassInfo.Assembly = cNotSet
                myClassInfo.ClassName = cNotSet
                myClassInfo.RunTimeVersion = cNotSet
                myClassInfo.ProgID = cNotSet

            Else
                myClassInfo.CodeBase = Registry.GetValue(_keylocation & cBackSlash & cInprocServer32, cCodeBase, cNotSet)
                myClassInfo.Assembly = Registry.GetValue(_keylocation & cBackSlash & cInprocServer32, cAssembly, cNotSet)
                myClassInfo.ClassName = Registry.GetValue(_keylocation & cBackSlash & cInprocServer32, cClass, cNotSet)
                myClassInfo.RunTimeVersion = Registry.GetValue(_keylocation & cBackSlash & cInprocServer32, cRuntimeVersion, cNotSet)
                myClassInfo.ProgID = Registry.GetValue(_keylocation & cBackSlash & cProgID, "", cNotSet)

            End If

        Catch ex As Exception
#If DEBUG Then
            Debug.Print(ex.ToString)
#End If
        End Try
        Return myClassInfo

    End Function



End Class


Friend Class ClassRegistryInformation
    Friend HIVE As String = ""
    Friend ClassName As String = ""
    Friend Assembly As String = ""
    Friend CodeBase As String = ""
    Friend RunTimeVersion As String = ""
    Friend ProgID As String = ""
End Class


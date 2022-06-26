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

''' <summary>
''' Class to capture information about a specific class entry - treat multiple class entriese.g. 32-bit 64-bti as separate entries
''' Only looks for information for a single AddIn based on it's Sparx key location
''' </summary>
Public Class ClassInformation

    ' Friend cClass As String = "Class"
    Private Const cAssembly As String = "Assembly"
    Private Const cRuntimeVersion As String = "RuntimeVersion"
    Private Const cThreadingModel As String = "ThreadingModel"


    Property ClassSource As String = ""


    ' Information from Spaxr registry keys
    ' AddIn name - the name of the Key in Sparx entry
    Property AddInName As String = ""
    ' Location of the AddIn key within Sparx keys
    Property AddInSource As String = ""

    'Property LibraryType As String = "" ' 64-bit / 32-bit /both


    ' Class information
    ' ClassID - captured from HKCR - note Class ID will be the same for the same named class even if multiple verssion 32-bit/ 64-bit
    Property ClassID As String = ""
    Property DLLSource As String = ""
    Property DLLVersion As String = ""
    ' run time version for DLL
    Property RunTimeVersion As String = ""
    ' Filename as captured from HKCR
    Property Filename As String = ""
    Property ThreadingModel As String = ""
    Property AssemblyAsString As String = ""

    ' Location of key in registry
    Property RegistryLocation As String = ""
    ' Colour used to indicate status 
    Property Colour As Color = Color.Magenta



    Friend Const cHKCR_ClassesRoot As String = "HKEY_CLASSES_ROOT\"
    Friend Const cHKCR_ClassesRootWow32CLSID As String = "HKEY_CLASSES_ROOT\Wow6432Node\CLSID\"
    Friend Const cHKCR_ClassesRootCLSID As String = "HKEY_CLASSES_ROOT\CLSID\"

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
    Friend Const HKCUWOW_Classes As String = "HKEY_CURRENT_USER\SOFTWARE\Classes\Wow6432Node"

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




    Friend OS64Bit As Boolean = False
    ''' <summary>
    ''' Constructor used to save Sparx information and get the class ID
    ''' </summary>
    ''' <param name="pAddInName"></param>
    ''' <param name="pAddInSource"></param>
    Sub New(pAddInName As String, pAddInSource As String)
        AddInName = pAddInName
        AddInSource = pAddInSource
        If Environment.Is64BitOperatingSystem Then OS64Bit = True
        getClassID()
        Debug.Print("Addin " & AddInName & " Class ID = " & ClassID)
    End Sub

    ' we have the addin name and location of Sparx key
    ' Note that the Sparx Key does not mean that the class will be registered correctly
    ' For 64-bit OS - 32bit classes are stored under WOW6432Node and 64-bit under HKCU,HKLM
    ' 
    ' Interesting note that there are 
    ' \HKEY_CURRENT_USER\SOFTWARE\WOW6432Node\ -> Not a lot under this so no need to check
    ' \HKEY_CURRENT_USER\SOFTWARE\Classes\WOW6432Node\CLSID
    ' The follow are mirrors?
    ' \HKEY_LOCAL_MACHINE\SOFTWARE\Classes\WOW6432Node\CLSID
    ' \HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Classes\CLSID

    ' for 64-bit on 64-bit os and 32-bit on 32-bit os
    ' HKEY_CURRENT_USER\SOFTWARE\Classes\CLSID
    ' HKEY_LOCAL_MACHINE\SOFTWARE\Classes\CLSID
    ' so to search 
    ' returns classID found from HKCR - the classID is then used to search for the relevant key elsewhere

    ''' <summary>
    ''' Returns the classID based on the Classname from the HKEY_CLASSES_ROOT
    ''' This hive is a merge of the others so will produce a classID for a classname regardless if it is 32-bit or 64-bit
    ''' Within HKCR there are CLSID entries and on 64-bit systems WOW6432NODE CLSID
    ''' Hence it is possibke to determing if there are 32, 64 or 2 classes
    ''' </summary>

    Private Sub getClassID()


        ClassSource = ""
        ClassID = ""
        ' for 64-bit OS we would expect the class information to be in the relevant hive - we could get from HKCR
        ' so use the Spaxr AddIn location PLUS the OS to identify where we would expect to find or not - get the information from 
        ' get all information from HKCR
        ' 64-bit OS
        ClassSource = cNotFound
        Debug.Print(AddInName)
        Select Case AddInSource
               '32-bit add in
            Case AddInEntry.cHKCU32
                ClassID = Registry.GetValue(cHKCR_ClassesRoot & AddInName & cBackSlash & cCLSID, "", cNotFound)
                If ClassID <> "" Then
                    ClassSource = AddInEntry.cHKCR32
                    populateClassInformation()
                    Dim c As String = CheckHKCU32()
                    If c <> "" Then ClassSource += "," & c
                    c = CheckHKLM32()
                    If c <> "" Then ClassSource += "," & c
                End If

            Case AddInEntry.cHKLM32
                ClassID = Registry.GetValue(cHKCR_ClassesRoot & AddInName & cBackSlash & cCLSID, "", cNotFound)
                If ClassID <> "" Then
                    ClassSource = AddInEntry.cHKCR32
                    populateClassInformation()
                    Dim c As String = CheckHKLM32()
                    If c <> "" Then ClassSource += "," & c
                    c = CheckHKCU32()
                    If c <> "" Then ClassSource += "," & c
                End If
                    '64-bit addins
            Case AddInEntry.cHKCU64
                ClassID = Registry.GetValue(cHKCR_ClassesRoot & AddInName & cBackSlash & cCLSID, "", cNotFound)
                If ClassID <> "" Then
                    ClassSource = AddInEntry.cHKCR64
                    populateClassInformation()
                    Dim c As String = CheckHKCU()
                    If c <> "" Then ClassSource += "," & c
                    c = CheckHKLM()
                    If c <> "" Then ClassSource += "," & c

                End If
            Case AddInEntry.cHKLM64
                ClassID = Registry.GetValue(cHKCR_ClassesRoot & AddInName & cBackSlash & cCLSID, "", cNotFound)
                If ClassID <> "" Then
                    ClassSource = AddInEntry.cHKCR64
                    populateClassInformation()
                    Dim c As String = CheckHKLM()
                    If c <> "" Then ClassSource += "," & c
                    c = CheckHKCU()
                    If c <> "" Then ClassSource += "," & c

                End If
            Case Else
                MsgBox("Unable to get class ID for " & AddInName & " Source " & AddInSource)
        End Select


        Return
    End Sub
    ''' <summary>
    ''' Check of there is an HKCU entry the current classID
    ''' </summary>
    ''' <returns>HKCR if yes else ""</returns>
    Private Function CheckHKCU32() As String
        Dim _location As String = IIf(OS64Bit, HKCUWOW_Classes & cBackSlash & cCLSID & cBackSlash & ClassID & cBackSlash & cInprocServer32,
                               HKCU_Classes & cBackSlash & cCLSID & cBackSlash & ClassID & cBackSlash & cInprocServer32)

        If _location <> "" Then
            Dim _filename As String = Registry.GetValue(_location, cCodeBase, "") 'using the class try to find the DLL path
            _filename = cleanFilename(_filename)
            If _filename = Filename Then
                Debug.Print("CheckHKCU32() same filename")
                Return AddInEntry.cHKCU32
            Else
                Debug.Print("HKCR " & Filename & "HKCU Filename= " & _filename)
            End If
        End If
        Return ""
    End Function
    ''' <summary>
    ''' Check if there are any HKLM entries for the current classID
    ''' </summary>
    ''' <returns>HKLM if yes else ""</returns>
    Private Function CheckHKLM32()
        ' there are 2 locations for WOW6432 in HKLM - it is believed they mirror each other but not sure so check both
        Dim _location As String = IIf(OS64Bit, HKLMWow1_Classes & cBackSlash & cCLSID & cBackSlash & ClassID & cBackSlash & cInprocServer32,
                       HKLM_Classes & cBackSlash & cCLSID & cBackSlash & ClassID & cBackSlash & cInprocServer32)
        If _location = "" Then ' try other location
            _location = IIf(OS64Bit, HKLMWow2_Classes & cBackSlash & ClassID & cBackSlash & cInprocServer32,
                       HKLM_Classes & cBackSlash & cCLSID & cBackSlash & ClassID & cBackSlash & cInprocServer32)
        End If
        If _location <> "" Then
            Dim _filename As String = Registry.GetValue(_location, cCodeBase, "") 'using the class try to find the DLL path
            _filename = cleanFilename(_filename)
            If _filename = Filename Then
                Debug.Print("CheckHKLM32() same filename")
                Return AddInEntry.cHKLM32
            Else
                Debug.Print("HKCR " & Filename & "HKCU Filename= " & _filename)
            End If
        End If
        Return ""
    End Function


    Private Function CheckHKCU() As String
        Dim _location As String = HKCU_Classes & cBackSlash & cCLSID & cBackSlash & ClassID & cBackSlash & cInprocServer32
        If RegistryLocation <> "" Then
            Dim _filename As String = Registry.GetValue(_location, cCodeBase, "") 'using the class try to find the DLL path
            _filename = cleanFilename(_filename)
            If _filename = Filename Then
                Debug.Print("CheckHKCU() same filename")
                Return AddInEntry.cHKCU64
            Else
                Debug.Print("HKCR " & Filename & "HKCU Filename= " & _filename)
            End If
        End If
        Return ""
    End Function
    Private Function CheckHKLM() As String
        Dim _location As String = HKLM_Classes & cBackSlash & cCLSID & cBackSlash & ClassID & cBackSlash & cInprocServer32

        If RegistryLocation <> "" Then
            Dim _filename As String = Registry.GetValue(_location, cCodeBase, "") 'using the class try to find the DLL path
            _filename = cleanFilename(_filename)
            If _filename = Filename Then
                Debug.Print("CheckHKLM() same filename")
                Return AddInEntry.cHKLM64
            Else
                Debug.Print("HKCR " & Filename & "HKCU Filename= " & _filename)
            End If
        End If
        Return ""

    End Function

    ' use the classID to get the details if there is a 32-bit class
    ' information extracted from HKCR - so only 32-bit or 64-bit
    Private Sub populateClassInformation()
        Select Case ClassSource
            Case AddInEntry.cHKCR32
                RegistryLocation = IIf(OS64Bit, cHKCR_ClassesRootWow32CLSID & ClassID & cBackSlash & cInprocServer32,
                               cHKCR_ClassesRootCLSID & ClassID & cBackSlash & cInprocServer32)
                DLLSource = AddInEntry.cHKCR32

            Case AddInEntry.cHKCR64
                RegistryLocation = cHKCR_ClassesRootCLSID & ClassID & cBackSlash & cInprocServer32
                DLLSource = AddInEntry.cHKCR64

            Case Else
                DLLSource = cNotFound
                MsgBox("No class entry in HKCR for Addin - " & AddInName & " Source " & AddInSource)
        End Select

        Try
            If RegistryLocation <> "" Then
                Filename = Registry.GetValue(RegistryLocation, cCodeBase, cNotSet) 'using the class try to find the DLL path
                If Filename IsNot Nothing Then getDLLAssembly()
                RunTimeVersion = Registry.GetValue(RegistryLocation, cRuntimeVersion, cNotSet)
                ThreadingModel = Registry.GetValue(RegistryLocation, cThreadingModel, cNotSet)
                AssemblyAsString = Registry.GetValue(RegistryLocation, cAssembly, cNotSet)
            End If


        Catch ex As Exception
            ' if the entry doesn't exist then we may get an exception
        End Try

    End Sub

    ' Get information for DLL file
    Private Sub getDLLAssembly()

        Dim _filename As String = Filename
        If _filename IsNot Nothing And _filename <> cNotSet Then
            If Strings.Left(_filename, fileprefixlength) = cFilePrefix Then _filename = Strings.Right(_filename, _filename.Length - fileprefixlength)
            Try
                Dim ass As AssemblyName = AssemblyName.GetAssemblyName(_filename)
                DLLVersion = ass.Version.ToString
            Catch ex As Exception
                DLLVersion = "Unable to determine"
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

    Private Function cleanFilename(_Filename As String) As String

        If Strings.Left(_Filename, fileprefixlength) = cFilePrefix Then _Filename = Strings.Right(_Filename, _Filename.Length - fileprefixlength)
        If _Filename <> "" Then _Filename = _Filename.Replace("/", "\")
        Return _Filename
    End Function

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


End Class


Friend Class ClassRegistryInformation
    Friend HIVE As String = ""
    Friend ClassName As String = ""
    Friend Assembly As String = ""
    Friend CodeBase As String = ""
    Friend RunTimeVersion As String = ""
    Friend ProgID As String = ""
End Class


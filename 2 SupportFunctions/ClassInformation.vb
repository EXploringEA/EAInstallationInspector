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

Imports System.IO
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




    'Classification of issues
    ' use booleans to flag 
    Private SparxKeyExists As Boolean = False
    Private ClassIDExists As Boolean = False
    Private DLLPathExists As Boolean = False
    Private MismatchedHives As Boolean = False
    Private DLLExistsInSpecifiedLocation As Boolean = False


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
    Private Property Filename As String = ""
    Property DisplayFilename As String
        Get
            Return cleanFilename(Filename)
        End Get
        Set(value As String)

        End Set
    End Property
    Property ThreadingModel As String = ""
    Property AssemblyAsString As String = ""

    ' Location of key in registry
    Property RegistryLocation As String = ""

    Friend Const cHKCR_ClassesRoot As String = "HKEY_CLASSES_ROOT\"
    'Friend Const cHKCR_ClassesRootWow32CLSID As String = "HKEY_CLASSES_ROOT\Wow6432Node\CLSID\"
    ' Friend Const cHKCR_ClassesRootCLSID As String = "HKEY_CLASSES_ROOT\CLSID\"

    ' Registry class locations
    ''' <summary>
    ''' HKCU Classes for 
    ''' * 32-bit on 32-bit OS
    ''' * 64-bit on 64-bit OS
    ''' </summary>
    Friend Const cHKCU_Classes As String = "HKEY_CURRENT_USER\SOFTWARE\Classes"
    Friend Const cHKCU_ClassesCLSID As String = "HKEY_CURRENT_USER\SOFTWARE\Classes\CLSID\"
    Friend Const cHKCUWOW_ClassesCLSID As String = "HKEY_CURRENT_USER\SOFTWARE\Classes\Wow6432Node\CLSID\"


    '''' <summary>
    '''' HKCU Classes for 
    '''' * 32-bit on 64-bit OS
    '''' </summary>
    Friend Const cHKCUWOW_Classes As String = "HKEY_CURRENT_USER\SOFTWARE\Classes\Wow6432Node"

    ''' <summary>
    ''' HKLM Classes
    ''' * 32-bit on 32-bit OS
    ''' * 64-bit on 64-bit OS
    ''' </summary>
    Friend Const cHKLM_Classes As String = "HKEY_LOCAL_MACHINE\SOFTWARE\Classes"
    Friend Const cHKLM_ClassesCLSID As String = "HKEY_LOCAL_MACHINE\SOFTWARE\Classes\CLSID\"

    ''' <summary>
    ''' HKLM Classes for 
    ''' * 32-bit on 64-bit OS
    ''' </summary>
    Friend Const cHKLMWow1_Classes As String = "HKEY_LOCAL_MACHINE\SOFTWARE\Classes\WOW6432Node"
    Friend Const cHKLMWow2_Classes As String = "HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Classes"
    Friend Const cHKLMWow_ClassesCLSID As String = "HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Classes\CLSID\"




    Private OS64Bit As Boolean = False
    ''' <summary>
    ''' Constructor used to save Sparx information and get the class ID
    ''' </summary>
    ''' <param name="pAddInName"></param>
    ''' <param name="pAddInSource"></param>
    Sub New(pAddInName As String, pAddInSource As String)
        AddInName = pAddInName
        AddInSource = pAddInSource
        If Environment.Is64BitOperatingSystem Then OS64Bit = True
        If AddInSource <> "" Then SparxKeyExists = True
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
    ''' The classID is then used to check in the relevant hives starting with that where the Sparx entry is defined and then if not found checked in the other HIVE
    ''' NOTE: 32-biut and 64-bit are treated separately
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
                    ClassIDExists = True
                    ClassSource = AddInEntry.cHKCU32
                    populateClassInformation(AddInEntry.cHKCU32)
                    Dim c As String = CheckHKLM32()
                    If c <> "" Then ClassSource += "," & c
                End If

            Case AddInEntry.cHKLM32
                ClassID = Registry.GetValue(cHKCR_ClassesRoot & AddInName & cBackSlash & cCLSID, "", cNotFound)
                If ClassID <> "" Then
                    ClassIDExists = True
                    ClassSource = AddInEntry.cHKLM32
                    populateClassInformation(AddInEntry.cHKLM32)
                    Dim c As String = CheckHKCU32()
                    If c <> "" Then ClassSource += "," & c
                End If
                    '64-bit addins
            Case AddInEntry.cHKCU64
                ClassID = Registry.GetValue(cHKCR_ClassesRoot & AddInName & cBackSlash & cCLSID, "", cNotFound)
                If ClassID <> "" Then
                    ClassIDExists = True
                    ClassSource = AddInEntry.cHKCU64
                    populateClassInformation(AddInEntry.cHKCU64)
                    Dim c As String = CheckHKLM()
                    If c <> "" Then ClassSource += "," & c

                End If
            Case AddInEntry.cHKLM64
                ClassID = Registry.GetValue(cHKCR_ClassesRoot & AddInName & cBackSlash & cCLSID, "", cNotFound)
                If ClassID <> "" Then
                    ClassIDExists = True
                    ClassSource = AddInEntry.cHKLM64
                    populateClassInformation(AddInEntry.cHKLM64)
                    Dim c As String = CheckHKCU()
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
        Dim _location As String = IIf(OS64Bit, cHKCUWOW_Classes & cBackSlash & cCLSID & cBackSlash & ClassID & cBackSlash & cInprocServer32,
                               cHKCU_Classes & cBackSlash & cCLSID & cBackSlash & ClassID & cBackSlash & cInprocServer32)

        If _location <> "" Then
            Dim _filename As String = Registry.GetValue(_location, cCodeBase, "") 'using the class try to find the DLL path
            If _filename <> "" And Filename = "" Then
                populateClassInformation(AddInEntry.cHKCU32)
            ElseIf _filename = Filename Then
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
        Dim _location As String = IIf(OS64Bit, cHKLMWow1_Classes & cBackSlash & cCLSID & cBackSlash & ClassID & cBackSlash & cInprocServer32,
                       cHKLM_Classes & cBackSlash & cCLSID & cBackSlash & ClassID & cBackSlash & cInprocServer32)
        If _location = "" Then ' try other location
            _location = IIf(OS64Bit, cHKLMWow2_Classes & cBackSlash & ClassID & cBackSlash & cInprocServer32,
                       cHKLM_Classes & cBackSlash & cCLSID & cBackSlash & ClassID & cBackSlash & cInprocServer32)
        End If
        If _location <> "" Then
            Dim _filename As String = Registry.GetValue(_location, cCodeBase, "") 'using the class try to find the DLL path
            If _filename <> "" And Filename = "" Then
                populateClassInformation(AddInEntry.cHKLM32)
            ElseIf _filename = Filename Then
                Debug.Print("CheckHKLM32() same filename")
                Return AddInEntry.cHKLM32
            Else
                Debug.Print("HKCR " & Filename & "HKCU Filename= " & _filename)
            End If

        End If
        Return ""
    End Function


    Private Function CheckHKCU() As String
        Dim _location As String = cHKCU_Classes & cBackSlash & cCLSID & cBackSlash & ClassID & cBackSlash & cInprocServer32
        If RegistryLocation <> "" Then
            Dim _filename As String = Registry.GetValue(_location, cCodeBase, "") 'using the class try to find the DLL path

            If _filename <> "" And Filename = "" Then
                populateClassInformation(AddInEntry.cHKCU64)
            ElseIf _filename = Filename Then
                Debug.Print("CheckHKCU64() same filename")
                Return AddInEntry.cHKCU64
            Else
                Debug.Print("HKCR " & Filename & "HKCU Filename= " & _filename)
            End If
        End If
        Return ""
    End Function
    Private Function CheckHKLM() As String
        Dim _location As String = cHKLM_Classes & cBackSlash & cCLSID & cBackSlash & ClassID & cBackSlash & cInprocServer32

        If RegistryLocation <> "" Then
            Dim _filename As String = Registry.GetValue(_location, cCodeBase, "") 'using the class try to find the DLL path

            If _filename <> "" And Filename = "" Then
                populateClassInformation(AddInEntry.cHKLM64)
            ElseIf _filename = Filename Then
                Debug.Print("CheckHKLM64() same filename")
                Return AddInEntry.cHKLM64
            Else
                Debug.Print("HKCR " & Filename & "HKCU Filename= " & _filename)
            End If
        End If
        Return ""

    End Function

    ''' <summary>
    ''' use the classID to get the details if there is a 32-bit class
    ''' information extracted from HKCU or HKLM - so only 32-bit or 64-bit
    ''' </summary>
    ''' <param name="pSource">This must be one of cHKCU32, cHKLM32, cHKCU64,cHKLM64</param>
    Private Sub populateClassInformation(pSource As String)
        Select Case pSource
            Case AddInEntry.cHKCU32 ' get information from HKCU
                RegistryLocation = IIf(OS64Bit, cHKCUWOW_ClassesCLSID & ClassID & cBackSlash & cInprocServer32,
                        cHKCU_ClassesCLSID & ClassID & cBackSlash & cInprocServer32)
                DLLSource = AddInEntry.cHKCU32

            Case AddInEntry.cHKLM32 ' get information from HKLM
                RegistryLocation = IIf(OS64Bit, cHKLMWow_ClassesCLSID & ClassID & cBackSlash & cInprocServer32,
                      cHKLM_ClassesCLSID & ClassID & cBackSlash & cInprocServer32)
                DLLSource = AddInEntry.cHKLM32

            Case AddInEntry.cHKCU64 ' get information from HKCU
                RegistryLocation = cHKCU_ClassesCLSID & ClassID & cBackSlash & cInprocServer32
                DLLSource = AddInEntry.cHKCU64

            Case AddInEntry.cHKLM64 ' get information from HKLM
                RegistryLocation = cHKLM_ClassesCLSID & ClassID & cBackSlash & cInprocServer32
                DLLSource = AddInEntry.cHKLM64

            Case Else
                DLLSource = cNotFound
                MsgBox("No class entry in HKCR for Addin - " & AddInName & " Source " & AddInSource)
        End Select

        Try
            If RegistryLocation <> "" Then
                Filename = Registry.GetValue(RegistryLocation, cCodeBase, cNotSet) 'using the class try to find the DLL path
                If Filename = "" Or Filename = cNotSet Then
                    Filename = Registry.GetValue(RegistryLocation, "", cNotSet)
                End If
                If Filename IsNot Nothing Then getDLLAssembly()
                RunTimeVersion = Registry.GetValue(RegistryLocation, cRuntimeVersion, cNotSet)
                ThreadingModel = Registry.GetValue(RegistryLocation, cThreadingModel, cNotSet)
                AssemblyAsString = Registry.GetValue(RegistryLocation, cAssembly, cNotSet)

            End If


        Catch ex As Exception
            ' if the entry doesn't exist then we may get an exception
        End Try

    End Sub

    ''' <summary>
    ''' Get information for DLL file
    ''' </summary>
    Private Sub getDLLAssembly()

        Dim _filename As String = cleanFilename(Filename)

        If _filename IsNot Nothing And _filename <> cNotSet Then
            If File.Exists(_filename) Then
                If Strings.Left(_filename, fileprefixlength) = cFilePrefix Then _filename = Strings.Right(_filename, _filename.Length - fileprefixlength)
                Try
                    Dim ass As AssemblyName = AssemblyName.GetAssemblyName(_filename)
                    DLLVersion = ass.Version.ToString
                Catch ex As Exception
                    DLLVersion = "Unable to determine"
                End Try
                'If _filename <> "" Then _filename = _filename.Replace("/", "\")
                Filename = cleanFilename(_filename)

                ' now depending on where items were found flag accordingly
            End If

            MismatchedHives = ClassSource <> DLLSource
            DLLPathExists = IIf(Filename = cNotSet, False, True)
            If DLLPathExists Then DLLExistsInSpecifiedLocation = DLLexists(Filename)

        End If

    End Sub


    ' was used to ensure the comparison between filenames was fair - however all entries that are checked are from registry so unless we need to 
    ' use the address other than checking OK to leave raw
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

    ' function uses the properies of the entry to define the colour that should be set
    ' precendence of colours
    ' SPARX present
    ' CLASSID
    ' ?loads
    ' Mismatched HIVES
    ' DLL exists
    ' DLL wrong location

    Friend Function getLineColour() As Color
        Dim c As Color = Color.White
        Try
            If Not SparxKeyExists Then Return Color.Red
            If Not ClassIDExists Then Return Color.HotPink
            If MismatchedHives Then Return Color.Wheat
            If Not DLLExistsInSpecifiedLocation Then Return Color.Pink
            Return Color.Lime

        Catch ex As Exception

        End Try
        Return Color.Red
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


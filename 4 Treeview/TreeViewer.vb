' Copyright (C) 2015 - 2022 Adrian LINCOLN, EXploringEA - All Rights Reserved
'
'   This program is free software: you can redistribute it and/or modify it under the terms of the GNU General Public License as published by
'   the Free Software Foundation, either version 3 of the License, or (at your option) any later version.

'   This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of
'   MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU General Public License for more details.
'
'    You should have received a copy of the GNU General Public License along with this program.  If not, see <http://www.gnu.org/licenses/>.
' =============================================================================================================================================

Imports System.Threading
Imports System.ComponentModel
Imports Microsoft.Win32
Imports System.Reflection
Imports System.IO

Partial Class frmInspector
    ''' <summary>
    ''' Handles the after expand to launch - this results in creating a background thread to do the processing
    ''' setting up the DoWork method and the workercompleted method
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub ProjectBrowser_AfterExpand(sender As System.Object, e As System.Windows.Forms.TreeViewEventArgs) Handles Browser.AfterExpand
        Try

            If e.Node.Nodes.ContainsKey(VIRTUALNODE) Then
                Dim bw As New BackgroundWorker()
                AddHandler bw.DoWork, AddressOf bw_DoWork
                AddHandler bw.RunWorkerCompleted, AddressOf bw_RunWorkerCompleted

                Dim oArgs As Object() = New Object() {e.Node, "Some information..."}
                bw.RunWorkerAsync(oArgs)
            End If
        Catch ex As Exception
#If DEBUG Then
            Debug.Print(ex.ToString)
#End If


        End Try
    End Sub

    ''' <summary>
    ''' Gets a list of the node children in its own thread
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub bw_DoWork(sender As Object, e As DoWorkEventArgs)
        Try

            Dim oArgs As Object() = TryCast(e.Argument, Object())
            Dim tNodeParent As TreeNode = TryCast(oArgs(0), TreeNode)
            Dim sInfo As String = oArgs(1).ToString()

            ' Note you can't use tNodeParent in here because
            ' we're not on the the UI thread (see Invoke).  We've only
            ' passed it in so we can round trip it to the 
            ' bw_RunWorkerCompleted event.

            ' We should get the parent and want to get the children of the element

            ' Use sInfo argument to load the data
            '   Dim r As New Random()
            '   Thread.Sleep(r.[Next](500, 2500))
            ' Dim arrChildren As String() = New String() {"Grapes", "Apples", "Tomatoes", "Kiwi"}

            Dim arrChildren As ArrayList = getNodeChildren(tNodeParent.Tag)
            ' Return the Parent Tree Node and the list of children to the
            ' UI thread.
            e.Result = New Object() {tNodeParent, arrChildren}
        Catch ex As Exception
#If DEBUG Then
            Debug.Print(ex.ToString)
#End If
        End Try
    End Sub

    ''' <summary>
    ''' Handle the completion from the background worker which has found the node children
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub bw_RunWorkerCompleted(sender As Object, e As RunWorkerCompletedEventArgs)
        ' Get the Parent Tree Node and the list of children
        ' from the Background Worker Thread
        Try

            Dim oResult As Object() = TryCast(e.Result, Object())
            Dim tNodeParent As TreeNode = TryCast(oResult(0), TreeNode)
            Dim arrChildren As ArrayList = TryCast(oResult(1), ArrayList)

            tNodeParent.Nodes.Clear()

            For Each sChild As NodeInfo In arrChildren
                Dim tNode As TreeNode = tNodeParent.Nodes.Add(sChild.Name)
                tNode.Tag = sChild
                If sChild.childCount > 0 Then AddVirtualNode(tNode)
            Next
        Catch ex As Exception
#If DEBUG Then
            Debug.Print(ex.ToString)
#End If
        End Try
    End Sub

    Private Const VIRTUALNODE As String = "virt"

    ''' <summary>
    ''' Add a virtual node at the specified node
    ''' </summary>
    ''' <param name="tNode"></param>
    ''' <remarks>No content is added - this will occur in the background after the node has been expanded</remarks>
    Private Sub AddVirtualNode(tNode As TreeNode)
        Try

            Dim tVirt As New TreeNode()
            tVirt.Text = "Loading..."
            tVirt.Name = VIRTUALNODE
            tVirt.ForeColor = Color.Blue
            tVirt.NodeFont = New Font("Microsoft Sans Serif", 8.25F, FontStyle.Underline)
            tNode.Nodes.Add(tVirt)
        Catch ex As Exception
#If DEBUG Then
            Debug.Print(ex.ToString)
#End If
        End Try
    End Sub

    Const cNumClassLibraryEntries As Integer = 3 ' this is the number of items that will be display for each class
    ' classname
    ' filename
    ' version

    ''' <summary>
    ''' Get an array of children give the parent package/element
    ''' </summary>
    ''' <param name="pParent"></param>
    ''' <returns>Nodes to add to parent</returns>
    ''' <remarks></remarks>
    Private Function getNodeChildren(pParent As NodeInfo) As ArrayList
        Dim myNodes As ArrayList = New ArrayList
        Try
            Dim myCUKey32 As RegistryKey = Nothing
            Dim myLMKey32 As RegistryKey = Nothing
            Dim myCUKey64 As RegistryKey = Nothing
            Dim myLMKey64 As RegistryKey = Nothing
            Dim myNode As NodeInfo = Nothing

            Select Case pParent.NodeType

'1. Top level Sparx AddIn information
                Case NodeType.SparxRoot ' this is the top level which captures information for each group of addins
                    ' 32-bit CU
                    myCUKey32 = Registry.CurrentUser.OpenSubKey(AddInInformation.SparxKeys32) ' get 32-bit SparxKeys for current user
                    myNode = New NodeInfo(NodeType.HIVE, myCUKey32.GetSubKeyNames.Count)
                    myNode.ClassNameLocation = AddInEntry.cHKCU32 ' "HKCU\Software\Sparx Systems\EAAddins"
                    myNode.Name = AddInEntry.cHKCU32
                    myNodes.Add(myNode)

                    ' 32-bit LM
                    myLMKey32 = Registry.LocalMachine.OpenSubKey(AddInInformation.SparxKeysWOW32) '"SOFTWARE\WOW6432Node\Sparx Systems\EAAddins") 'SparxKeys)
                    myNode = New NodeInfo(NodeType.HIVE, myLMKey32.GetSubKeyNames.Count)

                    If Environment.Is64BitOperatingSystem Then
                        myNode.ClassNameLocation = AddInEntry.cHKLM32Wow ' "HLKM\SOFTWARE\WOW6432Node\Sparx Systems\EAAddins"
                        myNode.Name = AddInEntry.cHKLM32Wow
                    Else ' for 32-bit system
                        myNode.ClassNameLocation = AddInEntry.cHKLM32 '"HLKM\SOFTWARE\Sparx Systems\EAAddins"
                        myNode.Name = AddInEntry.cHKLM32
                    End If
                    myNodes.Add(myNode)

                    ' 64bit Addins
                    If Environment.Is64BitOperatingSystem Then
                        myCUKey64 = Registry.CurrentUser.OpenSubKey(AddInInformation.SparxKeys64) '"Software\Sparx Systems\EAAddins64") 'SparxKeys) ' get the sparx keys listing the addins for the current user
                        myNode = New NodeInfo(NodeType.HIVE, myCUKey64.GetSubKeyNames.Count)
                        myNode.ClassNameLocation = AddInEntry.cHKCU64 ' "HKCU\Software\Sparx Systems\EAAddins64"
                        myNode.Name = AddInEntry.cHKCU64
                        myNodes.Add(myNode)

                        myLMKey64 = Registry.LocalMachine.OpenSubKey(AddInInformation.SparxKeys64) '"SOFTWARE\Sparx Systems\EAAddins64") 'SparxKeys)
                        myNode = New NodeInfo(NodeType.HIVE, myLMKey32.GetSubKeyNames.Count)
                        myNode.ClassNameLocation = AddInEntry.cHKLM64 '"HLKM\SOFTWARE\Sparx Systems\EAAddins64"
                        myNode.Name = AddInEntry.cHKLM64

                        myNodes.Add(myNode)
                    End If

'2. Individual References
                Case NodeType.HIVE ' this gets the names and classes for the AddIns
                    ' we now need to get the content for each of the children
                    Select Case pParent.ClassNameLocation
                        Case AddInEntry.cHKCU32 ' "HKCU\Software\Sparx Systems\EAAddins"
                            myCUKey32 = Registry.CurrentUser.OpenSubKey(AddInInformation.SparxKeys32)
                            For Each pEntryKey In myCUKey32.GetSubKeyNames
                                Dim newKey As New NodeInfo(NodeType.ClassNameNode, cNumClassLibraryEntries)
                                newKey.SparxEntryLocation = AddInEntry.cHKCU32
                                newKey.ClassNameLocation = AddInInformation.eaHKCU32AddInKeys
                                newKey.ClassName = Registry.GetValue(AddInInformation.eaHKCU32AddInKeys & cBackSlash & pEntryKey, "", cNotSet) ' look for the class name 
                                newKey.Name = pEntryKey & " : " & newKey.ClassName ' registry location for addin name 
                                myNodes.Add(newKey)
                            Next

                        Case AddInEntry.cHKLM32
                            myLMKey32 = Registry.LocalMachine.OpenSubKey(AddInInformation.SparxKeys32)

                            For Each pEntryKey In myLMKey32.GetSubKeyNames
                                Dim newKey As New NodeInfo(NodeType.ClassNameNode, cNumClassLibraryEntries)
                                newKey.SparxEntryLocation = AddInEntry.cHKLM32
                                newKey.ClassNameLocation = AddInInformation.eaHKLM32AddInKeys
                                newKey.ClassName = Registry.GetValue(AddInInformation.eaHKLM32AddInKeys & cBackSlash & pEntryKey, "", cNotSet) ' look for the class name
                                newKey.Name = pEntryKey & " : " & newKey.ClassName ' registry location for addin name 
                                myNodes.Add(newKey)
                            Next


                        Case AddInEntry.cHKLM32Wow ' "HLKM\SOFTWARE\WOW6432Node\Sparx Systems\EAAddins" ' 64-bit systems
                            myLMKey32 = Registry.LocalMachine.OpenSubKey(AddInInformation.SparxKeysWOW32)
                            ' myCUKey32 = Registry.CurrentUser.OpenSubKey(AddInInformation.SparxKeys32)

                            For Each pEntryKey In myLMKey32.GetSubKeyNames
                                Dim newKey As New NodeInfo(NodeType.ClassNameNode, cNumClassLibraryEntries)
                                newKey.SparxEntryLocation = AddInEntry.cHKLM32
                                newKey.ClassNameLocation = AddInInformation.eaHKLM32AddInKey64
                                newKey.ClassName = Registry.GetValue(AddInInformation.eaHKLM32AddInKey64 & cBackSlash & pEntryKey, "", cNotSet) ' look for the class name
                                newKey.Name = pEntryKey & " : " & newKey.ClassName ' registry location for addin name 
                                myNodes.Add(newKey)
                            Next

                            '' HKCU and HKLM 64bit addins
                        Case AddInEntry.cHKCU64 ' "HKCU\Software\Sparx Systems\EAAddins"
                            myCUKey64 = Registry.CurrentUser.OpenSubKey(AddInInformation.SparxKeys64)
                            For Each pEntryKey In myCUKey64.GetSubKeyNames
                                Dim newKey64 As New NodeInfo(NodeType.ClassNameNode, cNumClassLibraryEntries)
                                newKey64.SparxEntryLocation = AddInInformation.eaHKCU64AddInKeys
                                newKey64.ClassNameLocation = AddInInformation.eaHKCU64AddInKeys

                                newKey64.ClassName = Registry.GetValue(AddInInformation.eaHKCU64AddInKeys & cBackSlash & pEntryKey, "", cNotSet) ' look for the class name 
                                newKey64.Name = pEntryKey & " : " & newKey64.ClassName ' registry location for addin name 
                                myNodes.Add(newKey64)
                            Next

                        Case AddInEntry.cHKLM64  ' "HLKM\SOFTWARE\Sparx Systems\EAAddins64" ' 64-bit systems
                            myLMKey64 = Registry.LocalMachine.OpenSubKey(AddInInformation.SparxKeys64)
                            For Each pEntryKey In myLMKey64.GetSubKeyNames
                                Dim newKey64 As New NodeInfo(NodeType.ClassNameNode, cNumClassLibraryEntries)
                                newKey64.SparxEntryLocation = AddInEntry.cHKLM64
                                newKey64.ClassNameLocation = AddInInformation.eaHKLM64AddInKeys
                                newKey64.ClassName = Registry.GetValue(AddInInformation.eaHKLM64AddInKeys & cBackSlash & pEntryKey, "", cNotSet) ' look for the class name CLSID name
                                newKey64.Name = pEntryKey & " : " & newKey64.ClassName ' registry location for addin name 
                                myNodes.Add(newKey64)
                            Next

                    End Select

'3. Detail for each class
                Case NodeType.ClassNameNode ' this populates information relating to each class
                    ' we use the class name
                    ' if the source is HKCU
                    ' then look for CLSID in HKCU and only if not found then HKLM

                    myCUKey32 = Registry.CurrentUser.OpenSubKey(AddInInformation.SparxKeys32) '"Software\Sparx Systems\EAAddins"
                    myLMKey32 = If(Environment.Is64BitOperatingSystem, Registry.LocalMachine.OpenSubKey(AddInInformation.SparxKeysWOW32), Registry.LocalMachine.OpenSubKey(AddInInformation.SparxKeys32))
                    myCUKey64 = Registry.CurrentUser.OpenSubKey(AddInInformation.SparxKeys64) '"Software\Sparx Systems\EAAddins64"
                    myLMKey64 = Registry.LocalMachine.OpenSubKey(AddInInformation.SparxKeys64)

                    myNode = Nothing
                    Dim CLSIDsrc As String = "" ' Source for classID Current User or Local Machine
                    Dim myCLSIDLocation As String = "" ' location of CLSID
                    Dim myCLSID As String = "" ' The class ID

                    Select Case pParent.ClassNameLocation
                        Case AddInInformation.eaHKCU32AddInKeys
                            myCLSIDLocation = ClassInformation.HKCU_Classes & cBackSlash & pParent.ClassName & cBackSlash & cCLSID
                            myCLSID = Registry.GetValue(myCLSIDLocation, "", cNotSet) ' get the CLSID
                            If myCLSID <> "" Then
                                CLSIDsrc = AddInEntry.cHKCU32
                            Else
                                myCLSIDLocation = ClassInformation.HKLM_Classes & cBackSlash & pParent.ClassName & cBackSlash & cCLSID
                                myCLSID = Registry.GetValue(myCLSIDLocation, "", cNotSet) ' get the CLSID
                                If myCLSID <> "" Then CLSIDsrc = AddInEntry.cHKLM32Wow ' Other locations??
                            End If

                        Case AddInInformation.eaHKCU64AddInKeys
                            myCLSIDLocation = ClassInformation.HKCU_Classes & cBackSlash & pParent.ClassName & cBackSlash & cCLSID
                            myCLSID = Registry.GetValue(myCLSIDLocation, "", cNotSet) ' get the CLSID
                            If myCLSID <> "" Then
                                CLSIDsrc = AddInEntry.cHKCU64
                            Else
                                myCLSIDLocation = ClassInformation.HKLM_Classes & cBackSlash & pParent.ClassName & cBackSlash & cCLSID
                                myCLSID = Registry.GetValue(myCLSIDLocation, "", cNotSet) ' get the CLSID
                                If myCLSID <> "" Then CLSIDsrc = AddInEntry.cHKLM64
                            End If


                        Case AddInInformation.eaHKLM32AddInKeys
                            myCLSIDLocation = ClassInformation.HKLM_Classes & cBackSlash & pParent.ClassName & cBackSlash & cCLSID
                            myCLSID = Registry.GetValue(myCLSIDLocation, "", cNotSet) ' get the CLSID
                            If myCLSID <> "" Then
                                CLSIDsrc = AddInEntry.cHKLM32
                            Else
                                myCLSIDLocation = ClassInformation.HKCU_Classes & cBackSlash & pParent.ClassName & cBackSlash & cCLSID
                                myCLSID = Registry.GetValue(myCLSIDLocation, "", cNotSet) ' get the CLSID
                                If myCLSID <> "" Then CLSIDsrc = AddInEntry.cHKCU32
                            End If

                        Case AddInInformation.eaHKLM32AddInKey64
                            myCLSIDLocation = ClassInformation.HKLM_Classes & cBackSlash & pParent.ClassName & cBackSlash & cCLSID
                            myCLSID = Registry.GetValue(myCLSIDLocation, "", cNotSet) ' get the CLSID
                            If myCLSID <> "" Then
                                CLSIDsrc = AddInEntry.cHKLM32Wow
                            Else
                                myCLSIDLocation = ClassInformation.HKCU_Classes & cBackSlash & pParent.ClassName & cBackSlash & cCLSID
                                myCLSID = Registry.GetValue(myCLSIDLocation, "", cNotSet) ' get the CLSID
                                If myCLSID <> "" Then CLSIDsrc = AddInEntry.cHKCU32
                            End If
                        Case AddInInformation.eaHKLM64AddInKeys
                            myCLSIDLocation = ClassInformation.HKLM_Classes & cBackSlash & pParent.ClassName & cBackSlash & cCLSID
                            myCLSID = Registry.GetValue(myCLSIDLocation, "", cNotSet) ' get the CLSID
                            If myCLSID <> "" Then
                                CLSIDsrc = AddInEntry.cHKLM32
                            Else
                                myCLSIDLocation = ClassInformation.HKCU_Classes & cBackSlash & pParent.ClassName & cBackSlash & cCLSID
                                myCLSID = Registry.GetValue(myCLSIDLocation, "", cNotSet) ' get the CLSID
                                If myCLSID <> "" Then CLSIDsrc = AddInEntry.cHKCU32
                            End If
                        Case Else
#If DEBUG Then
                            Debug.Print("Treeview Class information unhandled classname " & pParent.ClassNameLocation)
#End If

                    End Select

                    myNode = New NodeInfo(NodeType.CLSIDNode, 0) ' no children
                    Select Case CLSIDsrc
                        Case AddInEntry.cHKCU32
                            myNode.Name = cHKCUCLSID & myCLSID
                        Case AddInEntry.cHKLM32
                            myNode.Name = cHKLMCLSID & myCLSID
                        Case AddInEntry.cHKLM32Wow
                            myNode.Name = cHKLMWowCLSID & myCLSID
                        Case AddInEntry.cHKCU64
                            myNode.Name = cHKCUCLSID & myCLSID
                        Case AddInEntry.cHKLM64
                            myNode.Name = cHKLMCLSID & myCLSID

                    End Select
                    myNode.CLSIDSrc = CLSIDsrc
                    myNode.CLSID = myCLSID
                    myNode.CLSIDLocation = myCLSIDLocation
                    myNode.ClassName = pParent.ClassName
                    myNode.ClassNameLocation = pParent.ClassNameLocation
                    myNodes.Add(myNode)


                    Dim _ClassInformation As ClassRegistryInformation = ClassInformation.OLDgetClassInformation(CLSIDsrc, myCLSID)

                    ' Classname
                    myNode = New NodeInfo(NodeType.CLSIDNode_ClassName, 0)
                    Select Case CLSIDsrc
                        Case AddInEntry.cHKCU32
                            myNode.Name = cHKCUClassname & _ClassInformation.ClassName
                        Case AddInEntry.cHKLM32
                            myNode.Name = cHKLMClassname & _ClassInformation.ClassName
                        Case AddInEntry.cHKLM32Wow
                            myNode.Name = cHKLMWowClassname & _ClassInformation.ClassName
                        Case AddInEntry.cHKCU64
                            myNode.Name = cHKCUClassname & _ClassInformation.ClassName
                        Case AddInEntry.cHKLM64
                            myNode.Name = cHKLMClassname & _ClassInformation.ClassName
                    End Select
                    myNode.CLSIDSrc = CLSIDsrc
                    myNode.ClassName = _ClassInformation.ClassName
                    myNode.ClassNameLocation = pParent.ClassNameLocation
                    myNode.CLSID = myCLSID
                    myNode.CLSIDLocation = myCLSIDLocation
                    myNodes.Add(myNode)


                    Dim _filename As String = _ClassInformation.CodeBase ' BUt this doesn't handle the case where the file is not in
                    If Strings.Left(_filename, fileprefixlength) = cFilePrefix Then _filename = Strings.Right(_filename, _filename.Length - fileprefixlength)

                    'Class ID


                    myNode = New NodeInfo(NodeType.CLSIDNode_CodeBase, 0)
                    Select Case CLSIDsrc
                        Case AddInEntry.cHKCU32
                            myNode.Name = cHKCUFilename & _filename
                        Case AddInEntry.cHKLM32
                            myNode.Name = cHKLMFilename & _filename
                        Case AddInEntry.cHKLM32Wow
                            myNode.Name = cHKLMWowFilename & _filename
                        Case AddInEntry.cHKCU64
                            myNode.Name = cHKCUFilename & _filename
                        Case AddInEntry.cHKLM64
                            myNode.Name = cHKLMFilename & _filename

                    End Select
                    myNode.Filename = Path.GetFileName(_filename)
                    If _filename <> "" Then _filename = _filename.Replace("/", "\")
                    myNode.FilePathName = _filename

                    myNode.ClassName = _ClassInformation.ClassName
                    myNode.ClassNameLocation = pParent.ClassNameLocation
                    myNode.CLSID = myCLSID
                    myNode.CLSIDSrc = CLSIDsrc
                    myNode.CLSIDLocation = myCLSIDLocation
                    myNodes.Add(myNode)

                    ' Filename

                    myNode = New NodeInfo(NodeType.CLSIDNode_Version, 0)
                    Dim AssVersion As String = "Unable to identify"
                    If _filename <> cNotSet And _filename IsNot Nothing Then

                        Try
                            Dim ass As AssemblyName = AssemblyName.GetAssemblyName(_filename)


                            ' Dim ass As Assembly = Assembly.LoadFile(_filename)
                            AssVersion = ass.Version.ToString() ' ass.GetName().Version.ToString
                        Catch ex As Exception
#If DEBUG Then
                            Debug.Print("Treeviewer assembly load failure " & ex.ToString)
#End If
                        End Try

                    End If

                    Select Case CLSIDsrc
                        Case AddInEntry.cHKCU32
                            myNode.Name = cHKCUVersion & AssVersion
                        Case AddInEntry.cHKLM32
                            myNode.Name = cHKLMVersion & AssVersion
                        Case AddInEntry.cHKLM32Wow
                            myNode.Name = cHKLMWowVersion & AssVersion
                        Case AddInEntry.cHKCU64
                            myNode.Name = cHKCUVersion & AssVersion
                        Case AddInEntry.cHKLM64
                            myNode.Name = cHKLMVersion & AssVersion
                    End Select


                    myNode.ClassName = _ClassInformation.ClassName
                    myNode.ClassNameLocation = pParent.ClassNameLocation
                    myNode.CLSID = myCLSID
                    myNode.CLSIDSrc = CLSIDsrc
                    myNode.CLSIDLocation = myCLSIDLocation
                    myNode.Filename = Path.GetFileName(_filename)
                    myNode.FilePathName = _filename
                    myNodes.Add(myNode)

                    ' Run time version
                    myNode = New NodeInfo(NodeType.CLSIDNode_RunTimeVersion, 0)
                    Select Case CLSIDsrc
                        Case AddInEntry.cHKCU32
                            myNode.Name = cHKCURuntimeVersion & _ClassInformation.RunTimeVersion
                        Case AddInEntry.cHKLM32
                            myNode.Name = cHKLMRuntimeVersion & _ClassInformation.RunTimeVersion
                        Case AddInEntry.cHKLM32Wow
                            myNode.Name = cHKLMWowRuntimeVersion & _ClassInformation.RunTimeVersion
                        Case AddInEntry.cHKCU64
                            myNode.Name = cHKCURuntimeVersion & _ClassInformation.RunTimeVersion
                        Case AddInEntry.cHKLM64
                            myNode.Name = cHKLMRuntimeVersion & _ClassInformation.RunTimeVersion
                    End Select
                    myNode.CLSIDSrc = CLSIDsrc
                    myNodes.Add(myNode)

                    ' ProgID
                    myNode = New NodeInfo(NodeType.CLSIDNode_ProgID, 0)
                    Select Case CLSIDsrc
                        Case AddInEntry.cHKCU32
                            myNode.Name = cHKCUProgID & _ClassInformation.ProgID
                        Case AddInEntry.cHKLM32
                            myNode.Name = cHKLMProgID & _ClassInformation.ProgID
                        Case AddInEntry.cHKLM32Wow
                            myNode.Name = cHKLMWowProgID & _ClassInformation.ProgID
                        Case AddInEntry.cHKCU64
                            myNode.Name = cHKCUProgID & _ClassInformation.ProgID
                        Case AddInEntry.cHKLM64
                            myNode.Name = cHKLMProgID & _ClassInformation.ProgID
                    End Select
                    myNode.ClassName = _ClassInformation.ClassName
                    myNode.ClassNameLocation = pParent.ClassNameLocation
                    myNode.CLSID = myCLSID
                    myNode.CLSIDSrc = CLSIDsrc
                    myNode.CLSIDLocation = myCLSIDLocation
                    myNode.ProgID = _ClassInformation.ProgID
                    myNodes.Add(myNode)

                Case Else
#If DEBUG Then
                    Debug.Print(pParent.NodeType)
#End If
            End Select

        Catch ex As Exception
#If DEBUG Then
            Debug.Print(ex.ToString)
#End If
        End Try
        Return myNodes
    End Function



    ''' <summary>
    ''' called after the collapse of a node - to force a reload next time around
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub ProjectBrowser_AfterCollapse(sender As System.Object, e As System.Windows.Forms.TreeViewEventArgs) Handles Browser.AfterCollapse
        Try
            ' if e.node has children then put in place a virtual node else
            If e.Node.Nodes.Count > 0 Then e.Node.Name = VIRTUALNODE
        Catch ex As Exception
#If DEBUG Then
            Debug.Print(ex.ToString)
#End If
        End Try
    End Sub

End Class

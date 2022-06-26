' Copyright (C) 2015 - 2022 Adrian LINCOLN, EXploringEA - All Rights Reserved
'
'   This program is free software: you can redistribute it and/or modify it under the terms of the GNU General Public License as published by
'   the Free Software Foundation, either version 3 of the License, or (at your option) any later version.

'   This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of
'   MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU General Public License for more details.
'
'    You should have received a copy of the GNU General Public License along with this program.  If not, see <http://www.gnu.org/licenses/>.
' =============================================================================================================================================

Imports System.ComponentModel
Imports Microsoft.Win32


Partial Class frmInspector


    Friend Const cNodeCLSID As String = "CLSID = "
    Friend Const cNodeClassname As String = "Classname = "
    Friend Const cNodeClassSources As String = "Class sources = "
    Friend Const cNodeDLLFilename As String = "DLL Filename = "
    Friend Const cNodeDLLVersion As String = "DLL Version = "
    Friend Const cNodeRuntimeVersion As String = "Runtime version = "
    Friend Const cNodeProgID As String = "ProgId = "

    Private Shared _64bitSystem As Boolean = False
    Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        ' keep a flag for type of system
        If Environment.Is64BitOperatingSystem Then _64bitSystem = True

    End Sub
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
                                newKey64.SparxEntryLocation = AddInEntry.cHKCU64
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
                    Dim _C As New ClassInformation(pParent.ClassName, pParent.SparxEntryLocation)

                    ' Classname
                    myNode = New NodeInfo(NodeType.CLSIDNode, 0) ' no children
                    myNode.Name = cNodeClassname & _C.AddInName
                    myNode.CLSIDSrc = _C.ClassSource 'CLSIDsrc
                    myNode.CLSID = _C.ClassID 'myCLSID
                    myNode.CLSIDLocation = _C.ClassSource 'myCLSIDLocation
                    myNode.ClassName = pParent.ClassName
                    myNode.ClassNameLocation = pParent.ClassNameLocation
                    myNodes.Add(myNode)

                    ' CLSID
                    myNode = New NodeInfo(NodeType.CLSIDNode_ClassName, 0)
                    myNode.Name = cNodeCLSID & _C.ClassID
                    myNode.CLSIDSrc = _C.ClassSource 'CLSIDsrc
                    myNode.ClassName = _C.AddInName '_ClassInformation.ClassName
                    myNode.ClassNameLocation = pParent.ClassNameLocation
                    myNode.CLSID = _C.ClassID 'myCLSID
                    myNode.CLSIDLocation = _C.ClassSource 'myCLSIDLocation
                    myNodes.Add(myNode)


                    ' CLSID
                    myNode = New NodeInfo(NodeType.CLSIDNode_ClassName, 0)
                    myNode.Name = cNodeClassSources & _C.ClassSource
                    myNode.CLSIDSrc = _C.ClassSource 'CLSIDsrc
                    myNode.ClassName = _C.AddInName '_ClassInformation.ClassName
                    myNode.ClassNameLocation = pParent.ClassNameLocation
                    myNode.CLSID = _C.ClassID 'myCLSID
                    myNode.CLSIDLocation = _C.ClassSource 'myCLSIDLocation
                    myNodes.Add(myNode)

                    ' Filename
                    myNode = New NodeInfo(NodeType.CLSIDNode_CodeBase, 0)
                    myNode.Name = CNodeDLLFilename & _C.Filename
                    myNode.FilePathName = _C.Filename '_filename
                    myNode.ClassName = _C.AddInName '.ClassName
                    myNode.ClassNameLocation = pParent.ClassNameLocation
                    myNode.CLSID = _C.ClassID
                    myNode.CLSIDLocation = _C.ClassSource 'myCLSIDLocation
                    myNodes.Add(myNode)

                    ' version
                    myNode = New NodeInfo(NodeType.CLSIDNode_Version, 0)
                    myNode.Name = cNodeDLLVersion & _C.DLLVersion
                    myNode.CLSIDSrc = _C.ClassID
                    myNodes.Add(myNode)

                    ' Can we get run time version of the DLL
                    myNode = New NodeInfo(NodeType.CLSIDNode_RunTimeVersion, 0)
                    myNode.Name = cNodeRuntimeVersion & _C.RunTimeVersion
                    myNode.CLSIDSrc = _C.ClassID
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

' Copyright (C) 2015 - 2018 Adrian LINCOLN, EXploringEA - All Rights Reserved
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

            Select Case pParent.NodeType

                Case NodeType.SparxRoot ' this is the top level which captures information about the number of addins
                    ' if we have this we know we have 2 children
                    Dim myCUKey As RegistryKey = Registry.CurrentUser.OpenSubKey(cAddins) '"Software\Sparx Systems\EAAddins") 'SparxKeys) ' get the sparx keys listing the addins for the current user
                    Dim myNode As New NodeInfo(NodeType.HIVE, myCUKey.GetSubKeyNames.Count)
                    myNode.ClassNameLocation = cHKCUAddins ' "HKCU\Software\Sparx Systems\EAAddins"
                    myNode.Name = cHKCUAddins ' "HKCU\Software\Sparx Systems\EAAddins"
                    myNodes.Add(myNode)

                    Dim myLMKey As RegistryKey = Registry.LocalMachine.OpenSubKey(cWowAddins) '"SOFTWARE\WOW6432Node\Sparx Systems\EAAddins") 'SparxKeys)
                    myNode = New NodeInfo(NodeType.HIVE, myLMKey.GetSubKeyNames.Count)

                    If Environment.Is64BitOperatingSystem Then
                        myNode.ClassNameLocation = cHKLMWowAddins ' "HLKM\SOFTWARE\WOW6432Node\Sparx Systems\EAAddins"
                        myNode.Name = cHKLMWowAddins ' "HLKM\SOFTWARE\WOW6432Node\Sparx Systems\EAAddins"
                    Else
                        myNode.ClassNameLocation = cHKLMAddins '"HLKM\SOFTWARE\Sparx Systems\EAAddins"
                        myNode.Name = cHKLMAddins '"HLKM\SOFTWARE\Sparx Systems\EAAddins"
                    End If
                    myNodes.Add(myNode)


                Case NodeType.HIVE ' this gets the names and classes for the AddIns
                    ' we now need to get the content for each of the children
                    Select Case pParent.ClassNameLocation
                        Case cHKCUAddins ' "HKCU\Software\Sparx Systems\EAAddins"
                            Dim myCUKey As RegistryKey = Registry.CurrentUser.OpenSubKey(cAddins) '"Software\Sparx Systems\EAAddins") 'SparxKeys) ' get the sparx keys listing the addins for the current user
                            For Each pEntryKey In myCUKey.GetSubKeyNames
                                Dim newKey As New NodeInfo(NodeType.ClassNameNode, cNumClassLibraryEntries)
                                newKey.SparxEntryLocation = cHKCU
                                newKey.ClassNameLocation = HKCUfullKey
                                newKey.ClassName = Registry.GetValue(HKCUfullKey & cBackSlash & pEntryKey, "", "NOT SET") ' look for the class name CLSID name
                                newKey.Name = pEntryKey & " : " & newKey.ClassName ' registry location for addin name ' this only works for top level not the 6432
                                myNodes.Add(newKey)
                            Next

                        Case cHKLMWowAddins ' "HLKM\SOFTWARE\WOW6432Node\Sparx Systems\EAAddins" ' 64-bit systems
                            Dim myLMKey As RegistryKey = Registry.LocalMachine.OpenSubKey(cWowAddins) '"SOFTWARE\WOW6432Node\Sparx Systems\EAAddins") 'SparxKeys)
                            Dim myCUKey As RegistryKey = Registry.CurrentUser.OpenSubKey(cAddins) '"Software\Sparx Systems\EAAddins") 'SparxKeys) ' get the sparx keys listing the addins for the current user

                            For Each pEntryKey In myLMKey.GetSubKeyNames
                                Dim newKey As New NodeInfo(NodeType.ClassNameNode, cNumClassLibraryEntries)
                                newKey.SparxEntryLocation = cHKLM
                                newKey.ClassNameLocation = HKLMfullKey
                                newKey.ClassName = Registry.GetValue(HKLMfullKey & cBackSlash & pEntryKey, "", "NOT SET") ' look for the class name CLSID name
                                newKey.Name = pEntryKey & " : " & newKey.ClassName ' registry location for addin name ' this only works for top level not the 6432
                                myNodes.Add(newKey)
                            Next

                        Case cHKLMAddins  ' "HLKM\SOFTWARE\Sparx Systems\EAAddins" ' 32-bit systems
                            Dim myLMKey As RegistryKey = Registry.LocalMachine.OpenSubKey(cAddins) '"SOFTWARE\Sparx Systems\EAAddins") 'SparxKeys)
                            Dim myCUKey As RegistryKey = Registry.CurrentUser.OpenSubKey(cAddins) '"Software\Sparx Systems\EAAddins") 'SparxKeys) ' get the sparx keys listing the addins for the current user

                            For Each pEntryKey In myLMKey.GetSubKeyNames
                                Dim newKey As New NodeInfo(NodeType.ClassNameNode, cNumClassLibraryEntries)
                                newKey.SparxEntryLocation = cHKLM
                                newKey.ClassNameLocation = HKLMfullKey
                                newKey.ClassName = Registry.GetValue(HKLMfullKey & cBackSlash & pEntryKey, "", "NOT SET") ' look for the class name CLSID name
                                newKey.Name = pEntryKey & " : " & newKey.ClassName ' registry location for addin name ' this only works for top level not the 6432
                                myNodes.Add(newKey)
                            Next

                    End Select


                Case NodeType.ClassNameNode ' this populates information relating to each class
                    ' we use the class name
                    ' if the source is HKCU
                    ' then look for CLSID in HKCU and only if not found then HKLM

                    Dim myCUKey As RegistryKey = Registry.CurrentUser.OpenSubKey(cAddins) '"Software\Sparx Systems\EAAddins") 'SparxKeys) ' get the sparx keys listing the addins for the current user


                    Dim myLMKey As RegistryKey = If(Environment.Is64BitOperatingSystem, _
                                   Registry.LocalMachine.OpenSubKey(cWowAddins), Registry.LocalMachine.OpenSubKey(cAddins))


                    Dim myNode As NodeInfo
                    Dim CLSIDsrc As String = "" ' Source for classID Current User or Local Machine
                    Dim myCLSIDLocation As String = "" ' location of CLSID
                    Dim myCLSID As String = "" ' The class ID

                    ' "HKEY_CURRENT_USER\SOFTWARE\Classes\eaDocXAddIn.eaDocX_AddIn\CLSID"

                    If pParent.SparxEntryLocation = cHKCU Then
                        If pParent.ClassNameLocation = cHKCUSparxAddinKeys Then
                            myCLSIDLocation = HKCUClasses & cBackSlash & pParent.ClassName & cBackSlash & cCLSID
                            myCLSID = Registry.GetValue(myCLSIDLocation, "", cNotSet) ' get the CLSID
                            If myCLSID <> "" Then
                                CLSIDsrc = cHKCU
                            Else
                                myCLSIDLocation = HKLMClasses & cBackSlash & pParent.ClassName & cBackSlash & cCLSID
                                myCLSID = Registry.GetValue(myCLSIDLocation, "", cNotSet) ' get the CLSID
                                If myCLSID <> "" Then CLSIDsrc = cHKLM
                            End If
                        ElseIf pParent.ClassNameLocation = cHKLMWowSparxAddinKeys Then
                            myCLSIDLocation = HKLMClasses & cBackSlash & pParent.ClassName & cBackSlash & cCLSID
                            myCLSID = Registry.GetValue(myCLSIDLocation, "", cNotSet) ' get the CLSID
                            If myCLSID <> "" Then
                                CLSIDsrc = cHKLMWow
                            Else
                                myCLSIDLocation = HKCUClasses & cBackSlash & pParent.ClassName & cBackSlash & cCLSID
                                myCLSID = Registry.GetValue(myCLSIDLocation, "", cNotSet) ' get the CLSID
                                If myCLSID <> "" Then CLSIDsrc = cHKCU
                            End If

                        ElseIf pParent.ClassNameLocation = cHKLMSparxAddinKeys Then
                            myCLSIDLocation = HKLMClasses & cBackSlash & pParent.ClassName & cBackSlash & cCLSID
                            myCLSID = Registry.GetValue(myCLSIDLocation, "", cNotSet) ' get the CLSID
                            If myCLSID <> "" Then
                                CLSIDsrc = cHKLM
                            Else
                                myCLSIDLocation = HKCUClasses & cBackSlash & pParent.ClassName & cBackSlash & cCLSID
                                myCLSID = Registry.GetValue(myCLSIDLocation, "", cNotSet) ' get the CLSID
                                If myCLSID <> "" Then CLSIDsrc = cHKCU
                            End If
                        End If

                    Else
                        If pParent.ClassNameLocation = cHKLMWowSparxAddinKeys Then
                            myCLSIDLocation = HKLMClasses & cBackSlash & pParent.ClassName & cBackSlash & cCLSID
                            myCLSID = Registry.GetValue(myCLSIDLocation, "", cNotSet) ' get the CLSID
                            If myCLSID <> "" Then
                                CLSIDsrc = cHKLMWow
                            Else
                                myCLSIDLocation = HKCUClasses & cBackSlash & pParent.ClassName & cBackSlash & cCLSID
                                myCLSID = Registry.GetValue(myCLSIDLocation, "", cNotSet) ' get the CLSID
                                If myCLSID <> "" Then CLSIDsrc = cHKCU
                            End If

                        ElseIf pParent.ClassNameLocation = cHKLMSparxAddinKeys Then
                            myCLSIDLocation = HKLMClasses & cBackSlash & pParent.ClassName & cBackSlash & cCLSID
                            myCLSID = Registry.GetValue(myCLSIDLocation, "", cNotSet) ' get the CLSID
                            If myCLSID <> "" Then
                                CLSIDsrc = cHKLM
                            Else
                                myCLSIDLocation = HKCUClasses & cBackSlash & pParent.ClassName & cBackSlash & cCLSID
                                myCLSID = Registry.GetValue(myCLSIDLocation, "", cNotSet) ' get the CLSID
                                If myCLSID <> "" Then CLSIDsrc = cHKCU
                            End If

                        ElseIf pParent.ClassNameLocation = cHKCUSparxAddinKeys Then
                            myCLSIDLocation = HKCUClasses & cBackSlash & pParent.ClassName & cBackSlash & cCLSID
                            myCLSID = Registry.GetValue(myCLSIDLocation, "", cNotSet) ' get the CLSID
                            If myCLSID <> "" Then
                                CLSIDsrc = cHKCU
                            Else
                                myCLSIDLocation = HKLMClasses & cBackSlash & pParent.ClassName & cBackSlash & cCLSID
                                myCLSID = Registry.GetValue(myCLSIDLocation, "", cNotSet) ' get the CLSID
                                If myCLSID <> "" Then CLSIDsrc = cHKLM
                            End If
                        End If
                    End If

                  
                    myNode = New NodeInfo(NodeType.CLSIDNode, 0) ' no children
                    Select Case CLSIDsrc
                        Case cHKCU
                            myNode.Name = cHKCUCLSID & myCLSID
                        Case cHKLM
                            myNode.Name = cHKLMCLSID & myCLSID
                        Case cHKLMWow
                            myNode.Name = cHKLMWowCLSID & myCLSID
                    End Select
                    myNode.CLSIDSrc = CLSIDsrc
                    myNode.CLSID = myCLSID
                    myNode.CLSIDLocation = myCLSIDLocation
                    myNode.ClassName = pParent.ClassName
                    myNode.ClassNameLocation = pParent.ClassNameLocation
                    myNodes.Add(myNode)


                    Dim _ClassInformation As ClassRegistryInformation = getClassInformation(CLSIDsrc, myCLSID)

                    ' Classname
                    myNode = New NodeInfo(NodeType.CLSIDNode_ClassName, 0)
                    Select Case CLSIDsrc
                        Case cHKCU
                            myNode.Name = cHKCUClassname & _ClassInformation.ClassName
                        Case cHKLM
                            myNode.Name = cHKLMClassname & _ClassInformation.ClassName
                        Case cHKLMWow
                            myNode.Name = cHKLMWowClassname & _ClassInformation.ClassName
                    End Select
                    myNode.CLSIDSrc = CLSIDsrc
                    myNode.ClassName = _ClassInformation.ClassName
                    myNode.ClassNameLocation = pParent.ClassNameLocation
                    myNode.CLSID = myCLSID
                    myNode.CLSIDLocation = myCLSIDLocation
                    myNodes.Add(myNode)


                    Dim _filename As String = _ClassInformation.CodeBase
                    If Strings.Left(_filename, fileprefixlength) = cFilePrefix Then _filename = Strings.Right(_filename, _filename.Length - fileprefixlength)

                    'Class ID
                    myNode = New NodeInfo(NodeType.CLSIDNode_CodeBase, 0)
                    Select Case CLSIDsrc
                        Case cHKCU
                            myNode.Name = cHKCUFilename & _filename
                        Case cHKLM
                            myNode.Name = cHKLMFilename & _filename
                        Case cHKLMWow
                            myNode.Name = cHKLMWowFilename & _filename
                    End Select
                    myNode.Filename = Path.GetFileName(_filename)
                    myNode.FilePathName = _filename

                    myNode.ClassName = _ClassInformation.ClassName
                    myNode.ClassNameLocation = pParent.ClassNameLocation
                    myNode.CLSID = myCLSID
                    myNode.CLSIDSrc = CLSIDsrc
                    myNode.CLSIDLocation = myCLSIDLocation
                    myNodes.Add(myNode)

                    ' Filename
                    Dim ass As Assembly = Assembly.LoadFile(_filename)
                    myNode = New NodeInfo(NodeType.CLSIDNode_Version, 0)
                    Select Case CLSIDsrc
                        Case cHKCU
                            myNode.Name = cHKCUVersion & ass.GetName().Version.ToString
                        Case cHKLM
                            myNode.Name = cHKLMVersion & ass.GetName().Version.ToString
                        Case cHKLMWow
                            myNode.Name = cHKLMWowVersion & ass.GetName().Version.ToString
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
                        Case cHKCU
                            myNode.Name = cHKCURuntimeVersion & _ClassInformation.RunTimeVersion
                        Case cHKLM
                            myNode.Name = cHKLMRuntimeVersion & _ClassInformation.RunTimeVersion
                        Case cHKLMWow
                            myNode.Name = cHKLMWowRuntimeVersion & _ClassInformation.RunTimeVersion
                    End Select
                    myNode.CLSIDSrc = CLSIDsrc
                    myNodes.Add(myNode)

                    ' ProgID
                    myNode = New NodeInfo(NodeType.CLSIDNode_ProgID, 0)
                    Select Case CLSIDsrc
                        Case cHKCU
                            myNode.Name = cHKCUProgID & _ClassInformation.ProgID
                        Case cHKLM
                            myNode.Name = cHKLMProgID & _ClassInformation.ProgID
                        Case cHKLMWow
                            myNode.Name = cHKLMWowProgID & _ClassInformation.ProgID
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

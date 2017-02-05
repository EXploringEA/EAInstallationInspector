'   Copyright (C) 2015-2016 EXploringEA
'
' This program is free software: you can redistribute it and/or modify it under the terms of the GNU General Public License as published by
' the Free Software Foundation, either version 3 of the License, or (at your option) any later version.
'
' This program is distributed in the hope that it will be useful,but WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  
' See the GNU General Public License for more details.
' You should have received a copy of the GNU General Public Licensealong with this program.  If not, see <http://www.gnu.org/licenses/>.
'

Public Class frmEntryDetail

    Protected Friend Sub New(pEntryDetail As AddInDetail)
        Try
            InitializeComponent()
            tbAddInName.Text = pEntryDetail.AddInName
            tbSparxRef.Text = pEntryDetail.SparxEntry
            tbAssemblyName.Text = pEntryDetail.ClassDefinition
            tbClassSource.Text = pEntryDetail.ClassSource
            tbCLSID.Text = pEntryDetail.CLSID
            tbCLSIDSRC.Text = pEntryDetail.CLSIDSource

            tbDLL.Text = pEntryDetail.DLL

        Catch ex As Exception

        End Try
    End Sub
    Private Sub btClose_Click(sender As Object, e As EventArgs) Handles btClose.Click
        Me.Close()
    End Sub
End Class
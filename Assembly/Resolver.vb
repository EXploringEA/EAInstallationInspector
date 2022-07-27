
Imports System.IO
Imports System.Reflection

Namespace Modules
    Module Resolver
        Function ResolveEventHandler(sender As Object, args As ResolveEventArgs) As Assembly

            Dim executingAssemblies As Assembly = Assembly.GetExecutingAssembly()
            Dim referencedAssembliesNames() As AssemblyName = executingAssemblies.GetReferencedAssemblies()
            Dim assemblyName As AssemblyName
            Dim dllAssembly As Assembly = Nothing

            For Each assemblyName In referencedAssembliesNames

                'Look for the assembly names that have raised the "AssemblyResolve" event.
                If (assemblyName.FullName.Substring(0, assemblyName.FullName.IndexOf(",", StringComparison.Ordinal)) = args.Name.Substring(0, args.Name.IndexOf(",", StringComparison.Ordinal))) Then

                    If My.Settings.DllFolder <> "" Then

                        ' Build path to place DLL
                        Dim tempAssemblyPath As String = Path.Combine(My.Settings.DllFolder, args.Name.Substring(0, args.Name.IndexOf(",", StringComparison.Ordinal)) & ".dll")
                        dllAssembly = Assembly.LoadFrom(tempAssemblyPath)

                        Exit For
                    End If


                End If
            Next

            Return dllAssembly

        End Function
    End Module
End Namespace

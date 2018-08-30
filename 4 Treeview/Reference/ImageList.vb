Friend Module ImageReference

#Region "Imagelist"

    Friend Const r_Action As String = "Action"
    Friend Const r_ActionPin As String = "ActionPin"
    Friend Const r_Activity As String = "Activity"
    Friend Const r_ActivityParameter As String = "ActivityParameter"
    Friend Const r_ActivityPartition As String = "ActivityPartition"
    Friend Const r_ActivityRegion As String = "ActivityRegion"
    Friend Const r_Actor As String = "Actor"
    Friend Const r_Artifact As String = "Artifact"
    Friend Const r_Attribute As String = "Attribute"
    Friend Const r_Association As String = "Association"
    Friend Const r_Boundary As String = "Boundary"
    Friend Const r_CentralBufferNode As String = "CentralBufferNode"
    Friend Const r_Change As String = "Change"
    Friend Const r_Class As String = "Class"
    Friend Const r_Collaboration As String = "Collaboration"
    Friend Const r_CollaborationOccurrence As String = "CollaborationOccurrence"
    Friend Const r_Comment As String = "Comment"
    Friend Const r_Component As String = "Component"
    Friend Const r_ConditionalNode As String = "ConditionalNode"
    Friend Const r_Constraint As String = "Constraint"
    Friend Const r_DataStore As String = "DataStore"
    Friend Const r_DataType As String = "DataType"
    Friend Const r_Decision As String = "Decision"
    Friend Const r_DeploymentSpecification As String = "DeploymentSpecification"
    Friend Const r_Device As String = "Device"
    Friend Const r_DiagramFrame As String = "DiagramFrame"
    Friend Const r_Entity As String = "Entity"
    Friend Const r_EntryPoint As String = "EntryPoint"
    Friend Const r_Enumeration As String = "Enumeration"
    Friend Const r_Event As String = "Event"
    Friend Const r_ExceptionHandler As String = "ExceptionHandler"
    Friend Const r_ExecutionEnvironment As String = "ExecutionEnvironment"
    Friend Const r_ExitPoint As String = "ExitPoint"
    Friend Const r_ExpansionNode As String = "ExpansionNode"
    Friend Const r_ExpansionRegion As String = "ExpansionRegion"
    Friend Const r_Feature As String = "Feature"
    Friend Const r_GUIElement As String = "GUIElement"
    Friend Const r_InformationItem As String = "InformationItem"
    Friend Const r_Interaction As String = "Interaction"
    Friend Const r_InteractionFragment As String = "InteractionFragment"
    Friend Const r_InteractionOccurrence As String = "InteractionOccurrence"
    Friend Const r_InteractionState As String = "InteractionState"
    Friend Const r_Interface As String = "Interface"
    Friend Const r_InterruptibleActivityRegion As String = "InterruptibleActivityRegion"
    Friend Const r_Issue As String = "Issue"
    Friend Const r_Label As String = "Label"
    Friend Const r_LoopNode As String = "LoopNode"
    Friend Const r_MergeNode As String = "MergeNode"
    Friend Const r_MessageEndpoint As String = "MessageEndpoint"
    Friend Const r_Method As String = "Method"
    Friend Const r_Node As String = "Node"
    Friend Const r_Note As String = "Note"
    Friend Const r_Object As String = "Object"
    Friend Const r_ObjectNode As String = "ObjectNode"
    Friend Const r_Package As String = "Package"
    Friend Const r_Parameter As String = "Parameter"
    Friend Const r_Part As String = "Part"
    Friend Const r_Port As String = "Port"
    Friend Const r_PrimitiveType As String = "PrimitiveType"
    Friend Const r_ProtocolStateMachine As String = "ProtocolStateMachine"
    Friend Const r_ProvidedInterface As String = "ProvidedInterface"
    Friend Const r_Region As String = "Region"
    Friend Const r_Report As String = "Report"
    Friend Const r_RequiredInterface As String = "RequiredInterface"
    Friend Const r_Requirement As String = "Requirement"
    Friend Const r_Risk As String = "Risk"
    Friend Const r_Screen As String = "Screen"
    Friend Const r_Sequence As String = "Sequence"
    Friend Const r_Signal As String = "Signal"
    Friend Const r_State As String = "State"
    Friend Const r_StateMachine As String = "StateMachine"
    Friend Const r_StateNode As String = "StateNode"
    Friend Const r_Synchronization As String = "Synchronization"
    Friend Const r_Task As String = "Task"
    Friend Const r_Text As String = "Text"
    Friend Const r_TimeLine As String = "TimeLine"
    Friend Const r_Trigger As String = "Trigger"
    Friend Const r_UMLDiagram As String = "UMLDiagram"
    Friend Const r_UseCase As String = "UseCase"
    Friend Const r_User As String = "User"
    Friend Const r_classview As String = "classview"
    Friend Const r_componentview As String = "componentview"
    Friend Const r_deploymentview As String = "deploymentview"
    Friend Const r_dynamicview As String = "dynamicview"
    Friend Const r_simpleview As String = "simpleview"
    Friend Const r_usecaseview As String = "usecaseview"
    Friend Const r_model As String = "model"

    Friend Const r_ActivityDiagram As String = "Activity"
    Friend Const r_AnalysisDiagram As String = "Analysis"
    Friend Const r_ClassDiagram As String = "Logical"
    Friend Const r_CommunicationDiagram As String = "Communication"
    Friend Const r_ComponentDiagram As String = "Component"
    Friend Const r_CompositeStructureDiagram As String = "CompositeStructure"
    Friend Const r_DataModellingDiagram As String = "DataModelling"
    Friend Const r_DocumentationDiagram As String = "Documentation"
    Friend Const r_InteractionOverviewDiagram As String = "InteractionOverview"
    Friend Const r_ObjectDiagram As String = "Object"
    Friend Const r_PackageDiagram As String = "Package"
    Friend Const r_SequenceDiagram As String = "Sequence"
    Friend Const r_StateMachineDiagram As String = "State Machine"
    Friend Const r_TimingDiagram As String = "Timing"
    Friend Const r_UseCaseDiagram As String = "Use Case"
    Friend Const r_DeploymentDiagram As String = "Deployement"
    Friend Const r_CustomDiagram As String = "Custom"
    Friend Const r_RequirementsDiagram As String = "Requirements"
    Friend Const r_MaintenanceDiagram As String = "Maintenance"


    ' index

    Friend Const i_Action As Integer = 1
    Friend Const i_ActionPin As Integer = 2
    Friend Const i_Activity As Integer = 3
    Friend Const i_ActivityParameter As Integer = 4
    Friend Const i_ActivityPartition As Integer = 5
    Friend Const i_ActivityRegion As Integer = 6
    Friend Const i_Actor As Integer = 7
    Friend Const i_Artifact As Integer = 8
    Friend Const i_Association As Integer = 9
    Friend Const i_Boundary As Integer = 10
    Friend Const i_CentralBufferNode As Integer = 11
    Friend Const i_Change As Integer = 12
    Friend Const i_Class As Integer = 13
    Friend Const i_Collaboration As Integer = 14
    Friend Const i_CollaborationOccurrence As Integer = 15
    Friend Const i_Comment As Integer = 16
    Friend Const i_Component As Integer = 17
    Friend Const i_ConditionalNode As Integer = 18
    Friend Const i_Constraint As Integer = 19
    Friend Const i_DataStore As Integer = 20
    Friend Const i_DataType As Integer = 21
    Friend Const i_Decision As Integer = 22
    Friend Const i_DeploymentSpecification As Integer = 23
    Friend Const i_Device As Integer = 24
    Friend Const i_DiagramFrame As Integer = 25
    Friend Const i_Entity As Integer = 26
    Friend Const i_EntryPoint As Integer = 27
    Friend Const i_Enumeration As Integer = 28
    Friend Const i_Event As Integer = 29
    Friend Const i_ExceptionHandler As Integer = 30
    Friend Const i_ExecutionEnvironment As Integer = 31
    Friend Const i_ExitPoint As Integer = 32
    Friend Const i_ExpansionNode As Integer = 33
    Friend Const i_ExpansionRegion As Integer = 34
    Friend Const i_Feature As Integer = 35
    Friend Const i_GUIElement As Integer = 36
    Friend Const i_InformationItem As Integer = 37
    Friend Const i_Interaction As Integer = 38
    Friend Const i_InteractionFragment As Integer = 39
    Friend Const i_InteractionOccurrence As Integer = 40
    Friend Const i_InteractionState As Integer = 41
    Friend Const i_Interface As Integer = 42
    Friend Const i_InterruptibleActivityRegion As Integer = 43
    Friend Const i_Issue As Integer = 44
    Friend Const i_Label As Integer = 45
    Friend Const i_LoopNode As Integer = 46
    Friend Const i_MergeNode As Integer = 47
    Friend Const i_MessageEndpoint As Integer = 48
    Friend Const i_Node As Integer = 49
    Friend Const i_Note As Integer = 50
    Friend Const i_Object As Integer = 51
    Friend Const i_ObjectNode As Integer = 52
    Friend Const i_Package As Integer = 53
    Friend Const i_Parameter As Integer = 54
    Friend Const i_Part As Integer = 55
    Friend Const i_Port As Integer = 56
    Friend Const i_PrimitiveType As Integer = 57
    Friend Const i_ProtocolStateMachine As Integer = 58
    Friend Const i_ProvidedInterface As Integer = 59
    Friend Const i_Region As Integer = 60
    Friend Const i_Report As Integer = 61
    Friend Const i_RequiredInterface As Integer = 62
    Friend Const i_Requirement As Integer = 63
    Friend Const i_Risk As Integer = 64
    Friend Const i_Screen As Integer = 65
    Friend Const i_Sequence As Integer = 66
    Friend Const i_Signal As Integer = 67
    Friend Const i_State As Integer = 68
    Friend Const i_StateMachine As Integer = 69
    Friend Const i_StateNode As Integer = 70
    Friend Const i_Synchronization As Integer = 71
    Friend Const i_Task As Integer = 72
    Friend Const i_Text As Integer = 73
    Friend Const i_TimeLine As Integer = 74
    Friend Const i_Trigger As Integer = 75
    Friend Const i_UMLDiagram As Integer = 76
    Friend Const i_UseCase As Integer = 77
    Friend Const i_User As Integer = 78
    Friend Const i_classview As Integer = 79
    Friend Const i_componentview As Integer = 80
    Friend Const i_deploymentview As Integer = 81
    Friend Const i_dynamicview As Integer = 82
    Friend Const i_simpleview As Integer = 83
    Friend Const i_usecaseview As Integer = 84
    Friend Const i_openPackage As Integer = 85
    Friend Const i_model As Integer = 86
    Friend Const i_Attribute As Integer = 87
    Friend Const i_Method As Integer = 88

    Friend Const i_ActivityDiagram = 89
    Friend Const i_AnalysisDiagram = 90
    Friend Const i_ClassDiagram = 91
    Friend Const i_CommunicationDiagram = 92
    Friend Const i_ComponentDiagram = 93
    Friend Const i_CompositeStructureDiagram = 94
    Friend Const i_DataModellingDiagram = 95
    Friend Const i_DocumentationDiagram = 96
    Friend Const i_InteractionOverviewDiagram = 97
    Friend Const i_ObjectDiagram = 98
    Friend Const i_PackageDiagram = 99
    Friend Const i_SequenceDiagram = 100
    Friend Const i_StateMachineDiagram = 101
    Friend Const i_TimingDiagram = 102
    Friend Const i_UseCaseDiagram = 103
    Friend Const i_DeploymentDiagram = 104
    Friend Const i_CustomDiagram = 105
    Friend Const i_RequirementsDiagram = 106
    Friend Const i_MaintenanceDiagram = 107



    Friend Function loadImageList() As ImageList

        Dim myImageList As New ImageList
        Try
            myImageList.Images.Add(My.Resources.Undefined) ' index 0
            myImageList.Images.Add(My.Resources.Action)
            myImageList.Images.Add(My.Resources.Undefined) 'ActionPin)
            myImageList.Images.Add(My.Resources.Activity)
            myImageList.Images.Add(My.Resources.Undefined) 'ActivityParameter)
            myImageList.Images.Add(My.Resources.ActivityPartition)
            myImageList.Images.Add(My.Resources.Undefined) 'ActivityRegion)
            myImageList.Images.Add(My.Resources.Actor)
            myImageList.Images.Add(My.Resources.Artifact)
            myImageList.Images.Add(My.Resources.Association)
            myImageList.Images.Add(My.Resources.Undefined) 'Boundary)
            myImageList.Images.Add(My.Resources.CentralBufferNode)
            myImageList.Images.Add(My.Resources.Change)
            myImageList.Images.Add(My.Resources._Class)
            myImageList.Images.Add(My.Resources.Collaboration)
            myImageList.Images.Add(My.Resources.Undefined) 'CollaborationOccurrence)
            myImageList.Images.Add(My.Resources.Undefined) 'Comment)
            myImageList.Images.Add(My.Resources.Component)
            myImageList.Images.Add(My.Resources.Undefined) 'ConditionalNode)
            myImageList.Images.Add(My.Resources.Undefined) 'Constraint)
            myImageList.Images.Add(My.Resources.DataStore)
            myImageList.Images.Add(My.Resources.DataType)
            myImageList.Images.Add(My.Resources.Decision)
            myImageList.Images.Add(My.Resources.DeploymentSpecification)
            myImageList.Images.Add(My.Resources.Device)
            myImageList.Images.Add(My.Resources.Undefined) 'DiagramFrame)
            myImageList.Images.Add(My.Resources.Undefined) 'Entity)
            myImageList.Images.Add(My.Resources.Undefined) 'EntryPoint)
            myImageList.Images.Add(My.Resources.Undefined) 'Enumeration)
            myImageList.Images.Add(My.Resources._Event)
            myImageList.Images.Add(My.Resources.Undefined) 'ExceptionHandler)
            myImageList.Images.Add(My.Resources.ExecutionEnvironment)
            myImageList.Images.Add(My.Resources.Undefined) 'ExitPoint)
            myImageList.Images.Add(My.Resources.Undefined) 'ExpansionNode)
            myImageList.Images.Add(My.Resources.Undefined) 'ExpansionRegion)
            myImageList.Images.Add(My.Resources.Feature)
            myImageList.Images.Add(My.Resources.Undefined) 'GUIElement)
            myImageList.Images.Add(My.Resources.Undefined) 'InformationItem)
            myImageList.Images.Add(My.Resources.Undefined) 'Interaction)
            myImageList.Images.Add(My.Resources.Undefined) 'InteractionFragment)
            myImageList.Images.Add(My.Resources.Undefined) 'InteractionOccurrence)
            myImageList.Images.Add(My.Resources.Undefined) 'InteractionState)
            myImageList.Images.Add(My.Resources._Interface)
            myImageList.Images.Add(My.Resources.Undefined) 'InterruptibleActivityRegion)
            myImageList.Images.Add(My.Resources.Issue)
            myImageList.Images.Add(My.Resources.Undefined) 'Label)
            myImageList.Images.Add(My.Resources.Undefined) 'LoopNode)
            myImageList.Images.Add(My.Resources.Undefined) 'MergeNode)
            myImageList.Images.Add(My.Resources.Undefined) 'MessageEndpoint)
            myImageList.Images.Add(My.Resources.Node)
            myImageList.Images.Add(My.Resources.Note)
            myImageList.Images.Add(My.Resources._Object)
            myImageList.Images.Add(My.Resources.Undefined) 'ObjectNode)
            myImageList.Images.Add(My.Resources.Package) ' box_front) ' was package
            myImageList.Images.Add(My.Resources.Undefined) 'Parameter)
            myImageList.Images.Add(My.Resources.Undefined) 'Part)
            myImageList.Images.Add(My.Resources.Undefined) 'Port)
            myImageList.Images.Add(My.Resources.PrimitiveType)
            myImageList.Images.Add(My.Resources.Undefined) 'ProtocolStateMachine)
            myImageList.Images.Add(My.Resources.ProvidedInterface)
            myImageList.Images.Add(My.Resources.Undefined) 'Region)
            myImageList.Images.Add(My.Resources.Undefined) 'Report)
            myImageList.Images.Add(My.Resources.RequiredInterface)
            myImageList.Images.Add(My.Resources.Requirement)
            myImageList.Images.Add(My.Resources.Risk)
            myImageList.Images.Add(My.Resources.Undefined) 'Screen)
            myImageList.Images.Add(My.Resources.Undefined) 'Sequence)
            myImageList.Images.Add(My.Resources.Signal)
            myImageList.Images.Add(My.Resources.Undefined) 'State)
            myImageList.Images.Add(My.Resources.StateMachine)
            myImageList.Images.Add(My.Resources.Undefined) 'StateNode)
            myImageList.Images.Add(My.Resources.Undefined) 'Synchronization)
            myImageList.Images.Add(My.Resources.Task)
            myImageList.Images.Add(My.Resources.Undefined) 'Text)
            myImageList.Images.Add(My.Resources.Undefined) 'TimeLine)
            myImageList.Images.Add(My.Resources.Undefined) 'Trigger)
            myImageList.Images.Add(My.Resources.Undefined) 'UMLDiagram)
            myImageList.Images.Add(My.Resources.UseCase)
            myImageList.Images.Add(My.Resources.User)
            myImageList.Images.Add(My.Resources.classview)
            myImageList.Images.Add(My.Resources.componentview)
            myImageList.Images.Add(My.Resources.deploymentview)
            myImageList.Images.Add(My.Resources.dynamicview)
            myImageList.Images.Add(My.Resources.simpleview)
            myImageList.Images.Add(My.Resources.usecaseview)
            myImageList.Images.Add(My.Resources.OpenPackage) 'box_front_open) 'openPackage)
            myImageList.Images.Add(My.Resources.Model) 'drawer) ' model)
            myImageList.Images.Add(My.Resources.Attribute)
            myImageList.Images.Add(My.Resources.Method)

            myImageList.Images.Add(My.Resources.ActivityDiagram)
            myImageList.Images.Add(My.Resources.AnalysisDiagram)
            myImageList.Images.Add(My.Resources.ClassDiagram)
            myImageList.Images.Add(My.Resources.CommunicationDiagram)
            myImageList.Images.Add(My.Resources.ComponentDiagram)
            myImageList.Images.Add(My.Resources.CompositeStructureDiagram)
            myImageList.Images.Add(My.Resources.DataModellingDiagram)
            myImageList.Images.Add(My.Resources.DocumentationDiagram)
            myImageList.Images.Add(My.Resources.InteractionOverviewDiagram)
            myImageList.Images.Add(My.Resources.ObjectDiagram)
            myImageList.Images.Add(My.Resources.PackageDiagram)
            myImageList.Images.Add(My.Resources.SequenceDiagram)
            myImageList.Images.Add(My.Resources.StateMachineDiagram)
            myImageList.Images.Add(My.Resources.TimingDiagram)
            myImageList.Images.Add(My.Resources.UseCaseDiagram) ' 103

            myImageList.Images.Add(My.Resources.ObjectDiagram) ' Deployment
            myImageList.Images.Add(My.Resources.ObjectDiagram) ' custom
            myImageList.Images.Add(My.Resources.ObjectDiagram) ' req
            myImageList.Images.Add(My.Resources.ObjectDiagram) ' maintenance


        Catch ex As Exception
            XEALogForm.reportAppException(ex.ToString)
        End Try
        Return myImageList
    End Function

    Friend Function ImageIndex(pType As String) As Integer

        Select Case pType
            Case r_Action
                Return i_Action
            Case r_ActionPin
                Return i_ActionPin
            Case r_Activity
                Return i_Activity
            Case r_ActivityParameter
                Return i_ActivityParameter
            Case r_ActivityPartition
                Return i_ActivityPartition
            Case r_ActivityRegion
                Return i_ActivityRegion
            Case r_Actor
                Return i_Actor
            Case r_Artifact
                Return i_Artifact
            Case r_Association
                Return i_Association
            Case r_Boundary
                Return i_Boundary
            Case r_CentralBufferNode
                Return i_CentralBufferNode
            Case r_Change
                Return i_Change
            Case r_Class
                Return i_Class
            Case r_Collaboration
                Return i_Collaboration
            Case r_CollaborationOccurrence
                Return i_CollaborationOccurrence
            Case r_Comment
                Return i_Comment
            Case r_Component
                Return i_Component
            Case r_ConditionalNode
                Return i_ConditionalNode
            Case r_Constraint
                Return i_Constraint
            Case r_DataStore
                Return i_DataStore
            Case r_DataType
                Return i_DataType
            Case r_Decision
                Return i_Decision
            Case r_DeploymentSpecification
                Return i_DeploymentSpecification
            Case r_Device
                Return i_Device
            Case r_DiagramFrame
                Return i_DiagramFrame
            Case r_Entity
                Return i_Entity
            Case r_EntryPoint
                Return i_EntryPoint
            Case r_Enumeration
                Return i_Enumeration
            Case r_Event
                Return i_Event
            Case r_ExceptionHandler
                Return i_ExceptionHandler
            Case r_ExecutionEnvironment
                Return i_ExecutionEnvironment
            Case r_ExitPoint
                Return i_ExitPoint
            Case r_ExpansionNode
                Return i_ExpansionNode
            Case r_ExpansionRegion
                Return i_ExpansionRegion
            Case r_Feature
                Return i_Feature
            Case r_GUIElement
                Return i_GUIElement
            Case r_InformationItem
                Return i_InformationItem
            Case r_Interaction
                Return i_Interaction
            Case r_InteractionFragment
                Return i_InteractionFragment
            Case r_InteractionOccurrence
                Return i_InteractionOccurrence
            Case r_InteractionState
                Return i_InteractionState
            Case r_Interface
                Return i_Interface
            Case r_InterruptibleActivityRegion
                Return i_InterruptibleActivityRegion
            Case r_Issue
                Return i_Issue
            Case r_Label
                Return i_Label
            Case r_LoopNode
                Return i_LoopNode
            Case r_MergeNode
                Return i_MergeNode
            Case r_MessageEndpoint
                Return i_MessageEndpoint
            Case r_Node
                Return i_Node
            Case r_Note
                Return i_Note
            Case r_Object
                Return i_Object
            Case r_ObjectNode
                Return i_ObjectNode
            Case r_Package
                Return i_Package
            Case r_Parameter
                Return i_Parameter
            Case r_Part
                Return i_Part
            Case r_Port
                Return i_Port
            Case r_PrimitiveType
                Return i_PrimitiveType
            Case r_ProtocolStateMachine
                Return i_ProtocolStateMachine
            Case r_ProvidedInterface
                Return i_ProvidedInterface
            Case r_Region
                Return i_Region
            Case r_Report
                Return i_Report
            Case r_RequiredInterface
                Return i_RequiredInterface
            Case r_Requirement
                Return i_Requirement
            Case r_Risk
                Return i_Risk
            Case r_Screen
                Return i_Screen
            Case r_Sequence
                Return i_Sequence
            Case r_Signal
                Return i_Signal
            Case r_State
                Return i_State
            Case r_StateMachine
                Return i_StateMachine
            Case r_StateNode
                Return i_StateNode
            Case r_Synchronization
                Return i_Synchronization
            Case r_Task
                Return i_Task
            Case r_Text
                Return i_Text
            Case r_TimeLine
                Return i_TimeLine
            Case r_Trigger
                Return i_Trigger
            Case r_UMLDiagram
                Return i_UMLDiagram
            Case r_UseCase
                Return i_UseCase
            Case r_User
                Return i_User
            Case r_classview
                Return i_classview
            Case r_componentview
                Return i_componentview
            Case r_deploymentview
                Return i_deploymentview
            Case r_dynamicview
                Return i_dynamicview
            Case r_simpleview
                Return i_simpleview
            Case r_usecaseview
                Return i_usecaseview
            Case r_model
                Return i_model
            Case r_Attribute
                Return i_Attribute
            Case r_Method
                Return i_Method

            Case r_ActivityDiagram
                Return i_ActivityDiagram
            Case r_AnalysisDiagram
                Return i_AnalysisDiagram
            Case r_ClassDiagram
                Return i_ClassDiagram
            Case r_CommunicationDiagram
                Return i_CommunicationDiagram
            Case r_ComponentDiagram
                Return i_ComponentDiagram
            Case r_CompositeStructureDiagram
                Return i_CompositeStructureDiagram
            Case r_DataModellingDiagram
                Return i_DataModellingDiagram
            Case r_DocumentationDiagram
                Return i_DocumentationDiagram
            Case r_InteractionOverviewDiagram
                Return i_InteractionOverviewDiagram
            Case r_ObjectDiagram
                Return i_ObjectDiagram
            Case r_PackageDiagram
                Return i_PackageDiagram
            Case r_SequenceDiagram
                Return i_SequenceDiagram
            Case r_StateMachineDiagram
                Return i_StateMachineDiagram
            Case r_TimingDiagram
                Return i_TimingDiagram
            Case r_UseCaseDiagram
                Return i_UseCaseDiagram
            Case r_DeploymentDiagram
                Return i_DeploymentDiagram
            Case r_CustomDiagram
                Return i_CustomDiagram
            Case r_RequirementsDiagram
                Return i_RequirementsDiagram
            Case r_MaintenanceDiagram
                Return i_MaintenanceDiagram

            Case Else
                XEALogForm.reportAppError("No icon for " & pType)
        End Select
        Return 0
    End Function
#End Region



End Module

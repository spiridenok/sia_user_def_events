Private WithEvents events_tree As clsTreeView
Private WithEvents vars_tree As clsTreeView
Private selected_node As clsNode
Public confirmed_node As clsNode
Public confirmed_node_is_variable As Boolean
Public show_variable As Boolean

Private Sub ButtonCancel_Click()
    FormSelectItem.Hide
End Sub

Private Sub ButtonHelp_Click()
    MsgBox "This will pop-up a window with explanation what can be done using this window."
End Sub

Private Sub ButtonSelect_Click()
    If is_udev(selected_node) Then
        MsgBox "It is not possible to use user defined event nodes in node conditions! please choose another node."
    Else
        Set confirmed_node = selected_node
        FormSelectItem.Hide
    End If
End Sub

Private Sub TreeSelectVariable_Click()

End Sub

Private Sub UserForm_Activate()
    If Me.show_variable Then
        MultiPage1.Pages.Item(0).Enabled = True
        MultiPage1.Pages.Item(1).Enabled = False
    Else
        MultiPage1.Pages.Item(1).Enabled = True
        MultiPage1.Pages.Item(0).Enabled = False
    End If
    
End Sub

Private Sub vars_tree_click(cNode As clsNode)
    If cNode.ChildNodes Is Nothing Then
        ButtonSelect.Enabled = True
    Else
        ButtonSelect.Enabled = False
    End If
    Set selected_node = cNode
    confirmed_node_is_variable = True
End Sub

Private Sub events_tree_click(cNode As clsNode)
    If cNode.ChildNodes Is Nothing Then
        ButtonSelect.Enabled = True
    Else
        ButtonSelect.Enabled = False
    End If
    Set selected_node = cNode
    confirmed_node_is_variable = False
End Sub



Private Sub UserForm_Initialize()
    Set events_tree = New clsTreeView
    
    Dim events_root As clsNode
    Dim cNode As clsNode
    
    With events_tree
    
        Set .TreeControl = Me.TreeSelectEvent
        .AppName = "DNDM Events Tree"
        .RootButton = True
        '.Form = Me
        
        Set events_root = .AddRoot("DNDM:ROOT", "DNDMxDEF_EVENT_POOL")
        Set cNode = events_root.AddChild("MI", "MI")
        cNode.AddChild "MI:E_1", "MI event ONE"
        cNode.AddChild "MI:E_2", "MI event TWO"
        Set cNode = events_root.AddChild("MPDM", "MPDM")
        cNode.AddChild "MPDM:E_A", "PALM event A"
        cNode.AddChild "MPDM:E_B", "PALM event B"
        Set events_pgma_root = events_root.AddChild("PGMA", "PGMA_stages")
        Set cPgmaRegNode = events_pgma_root.AddChild("PGMA:OTHER_EVENTS", "Regular Events")
        Set cNode = cPgmaRegNode.AddChild("PGMA:OTHER_EVENT_1", "Some Regular Event")
        Set cPgmaUdevNode = events_pgma_root.AddChild("PGMA:UDEV_EVENTS", "User Defined Events Events")
        Set cNode = cPgmaUdevNode.AddChild("PGMA:UDEV_EVENT_A", "User Defined Event A")
        Set cNode = cPgmaUdevNode.AddChild("PGMA:UDEV_EVENT_B", "User Defined Event B")
        
        .Refresh
        
    End With

    Set vars_tree = New clsTreeView
    
    Dim vars_root As clsNode
    'Dim cNode As clsNode
    
    With vars_tree
    
        Set .TreeControl = Me.TreeSelectVariable
        .AppName = "DNDM Events Tree"
        .RootButton = True
        
        Set vars_root = .AddRoot("DNDM:ROOT", "DNDMxDEF_VARIABLE_POOL")
        Set cNode = vars_root.AddChild("MI", "MI")
        cNode.AddChild "MI:V_11", "MI variable 11"
        cNode.AddChild "MI:V_22", "MI variable 22"
        Set cNode = vars_root.AddChild("MPDM", "MPDM")
        cNode.AddChild "MPDM:V_AA", "PALM variable AA"
        cNode.AddChild "MPDM:V_BB", "PALM variable BB"
        Set vars_pgma_root = vars_root.AddChild("PGMA", "PGMA_stages")
        Set vars_pgma_ss_root = vars_pgma_root.AddChild("PGMA:SS_VARS", "Short Stroke Variables")
        Set cNode = vars_pgma_ss_root.AddChild("PGMA:SS_POS_X", "Position X")
        Set cNode = vars_pgma_ss_root.AddChild("PGMA:SS_POS_Y", "Position Y")
        Set vars_pgma_ls_root = vars_pgma_root.AddChild("PGMA:LS_VARS", "Long Stroke Variables")
        Set cNode = vars_pgma_ls_root.AddChild("PGMA:LS_POS_X", "Speed Z")
        Set cNode = vars_pgma_ls_root.AddChild("PGMA:LS_POS_Y", "Acceleration Z")
        
        .Refresh
        
    End With
    
    Set confirmed_node = Nothing

End Sub


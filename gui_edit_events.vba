Private WithEvents mcTree As clsTreeView

Private Sub mcTree_Click(cNode As clsNode)
    UdevPropertiesButton.Enabled = is_udev(cNode)
End Sub

Private Sub UdevPropertiesButton_Click()
    FormUdevProp.Show
End Sub

Private Sub UserForm_Initialize()
    Set mcTree = New clsTreeView
    
    Dim cRoot As clsNode
    Dim cNode As clsNode
    
    With mcTree
    
        Set .TreeControl = Me.EventsTree
        .AppName = "DNDM Events Tree"
        .RootButton = True
        '.Form = Me
        
        Set cRoot = .AddRoot("DNDM:ROOT", "DNDMxDEF_EVENT_POOL")
        Set cNode = cRoot.AddChild("MI", "MI")
        Set cNode = cNode.AddChild("MI:1", "Nothing interesting")
        Set cNode = cRoot.AddChild("MPDM", "MPDM")
        Set cNode = cNode.AddChild("MPDM:1", "Nothing interesting")
        Set cPgmaNode = cRoot.AddChild("PGMA", "PGMA_stages")
        Set cPgmaRegNode = cPgmaNode.AddChild("PGMA:OTHER_EVENTS", "Regular Events")
        Set cNode = cPgmaRegNode.AddChild("PGMA:OTHER_EVENT_1", "Some Regular Event")
        Set cPgmaUdevNode = cPgmaNode.AddChild("PGMA:UDEV_EVENTS", "User Defined Events Events")
        Set cNode = cPgmaUdevNode.AddChild("PGMA:UDEV_EVENT_A", "User Defined Event A")
        Set cNode = cPgmaUdevNode.AddChild("PGMA:UDEV_EVENT_B", "User Defined Event B")
        
        .Refresh
        
    End With
    
    
End Sub


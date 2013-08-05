Const SA_STR = "Select action..."
Const SV_STR = "Select variable..."
Const SE_STR = "Select event..."
Const SC_STR = "Select Constant..."
Const SR_OP1_STR = "Result op operation#1"
Const SR_OP2_STR = "Result op operation#2"
Private var_selected As Boolean
' used as a workaround for click() events that are caused by selecting active values of ComboBoxes
Private ignore_update As Boolean
Const INIT_NUM = 5

Sub check_counters(cb As ComboBox, ByRef op_counter As Integer, ByRef inp_counter As Integer)
    If op_selected(cb.value) Then
        op_counter = op_counter - 1
    Else
        If cb.value <> SA_STR And cb.value <> SV_STR And cb.value <> SE_STR And cb.value <> SC And InStr(1, cb.value, ":") > 0 Then inp_counter = inp_counter - 1
    End If
End Sub


Sub update_counters()
    Dim op_counter As Integer
    Dim inp_counter As Integer
    
    op_counter = 10
    inp_counter = 10
    
    check_counters CbFirstAction, op_counter, inp_counter
    check_counters CbSecondAction, op_counter, inp_counter
    check_counters cbThirdAction, op_counter, inp_counter
    check_counters cbFirstAction_2, op_counter, inp_counter
    check_counters cbSecondAction_2, op_counter, inp_counter
    check_counters cbThirdAction_2, op_counter, inp_counter
    check_counters cbFirstAction_3, op_counter, inp_counter
    check_counters cbSecondAction_3, op_counter, inp_counter
    check_counters cbThirdAction_3, op_counter, inp_counter
    
    LabelNumOperations.Caption = op_counter
    LabelNumInputs.Caption = inp_counter
End Sub

Private Function result_selected(val As String) As Boolean
    result_selected = (val = SR_OP1_STR) Or (val = SR_OP2_STR)
End Function


Private Sub ButtonClearConditions_Click()
    CbFirstAction.Clear
    CbFirstAction.Text = SA_STR
    CbFirstAction.AddItem "ABS"
    CbFirstAction.AddItem "SQRT"
    CbFirstAction.AddItem SV_STR
    CbFirstAction.AddItem SE_STR

    CbSecondAction.Clear
    CbSecondAction.Text = SA_STR
    CbSecondAction.Enabled = False
    
    cbThirdAction.Clear
    cbThirdAction.Text = SA_STR
    cbThirdAction.Enabled = False
    
    cbFirstAction_2.Clear
    cbFirstAction_2.Text = SA_STR
    cbFirstAction_2.Enabled = False
    
    cbSecondAction_2.Clear
    cbSecondAction_2.Text = SA_STR
    cbSecondAction_2.Enabled = False
    
    cbThirdAction_2.Clear
    cbThirdAction_2.Text = SA_STR
    cbThirdAction_2.Enabled = False
    
    update_counters
End Sub

Private Sub ButtonUdevPropCancel_Click()
    FormUdevProp.Hide
End Sub

Private Sub ButtonUdevPropOk_Click()
    FormUdevProp.Hide
End Sub

Private Function op_selected(value As String) As Boolean
    op_selected = (value = "PLUS" Or value = "MINUS" Or value = "MUL" Or value = "DIV" Or _
                        value = "ABS" Or value = "LESS THAN" Or value = "GREATER THAN" Or _
                        value = "OR" Or value = "SQRT" Or value = "AND")
End Function

Public Sub handle_1st_action(ByRef cb As ComboBox, ByRef next_cb As ComboBox)
    If ignore_update Then Exit Sub
    next_cb.Clear
    If Not op_selected(cb.value) Then
        If result_selected(cb.value) Then
            LabelNumInputs.Caption = LabelNumInputs.Caption - 1
            next_cb.Enabled = True
            If var_selected Then
                next_cb.AddItem "PLUS"
                next_cb.AddItem "MINUS"
                next_cb.AddItem "MUL"
                next_cb.AddItem "DIV"
                next_cb.AddItem "LESS THAN"
                next_cb.AddItem "GREATER THAN"
            Else
                next_cb.AddItem "OR"
                next_cb.AddItem "AND"
            End If
        Else
            If FormSelectItem.confirmed_node Is Nothing Or cb.value = SV_STR Or cb.value = SE_STR Then
                var_selected = (cb.value = SV_STR)
                FormSelectItem.show_variable = var_selected
                FormSelectItem.Show
            Else
                If Not in_list(cb, FormSelectItem.confirmed_node.Key) Then
                    var_selected = (cb.value = SV_STR)
                    FormSelectItem.show_variable = var_selected
                    FormSelectItem.Show
                End If
            End If
            If FormSelectItem.confirmed_node Is Nothing Then
                next_cb.Enabled = False
            Else
                LabelNumInputs.Caption = LabelNumInputs.Caption - 1
                next_cb.Enabled = True
                If Not in_list(cb, FormSelectItem.confirmed_node.Key) Then
                    cb.AddItem FormSelectItem.confirmed_node.Key
                    ignore_update = True
                    cb.value = FormSelectItem.confirmed_node.Key
                    ignore_update = False
                End If
                If FormSelectItem.confirmed_node_is_variable Then
                    next_cb.AddItem "PLUS"
                    next_cb.AddItem "MINUS"
                    next_cb.AddItem "MUL"
                    next_cb.AddItem "DIV"
                    next_cb.AddItem "LESS THAN"
                    next_cb.AddItem "GREATER THAN"
                Else
                    next_cb.AddItem "OR"
                    next_cb.AddItem "AND"
                End If
            End If
        End If
    Else
        next_cb.Enabled = True
        next_cb.AddItem SV_STR
    End If
    update_counters
End Sub

Private Sub cbFirstAction_2_Click()
    handle_1st_action cbFirstAction_2, cbSecondAction_2
End Sub

Private Sub cbFirstAction_3_Click()
    handle_1st_action cbFirstAction_3, cbSecondAction_3
End Sub


Private Sub CbFirstAction_Click()
    handle_1st_action CbFirstAction, CbSecondAction
End Sub

Public Sub handle_2st_action(ByRef cb As ComboBox, ByRef next_cb As ComboBox, ByRef next_row_cb As Variant)
    next_cb.Clear
    If ignore_update Then Exit Sub
    
    If op_selected(cb.value) Then
        next_cb.Enabled = True
        If var_selected Then
            next_cb.AddItem SV_STR
            next_cb.AddItem SC_STR
        Else
            next_cb.AddItem SE_STR
        End If
    Else
        If cb.value = SV_STR Or cb.value = SE_STR Then
            var_selected = True
            FormSelectItem.show_variable = var_selected
            FormSelectItem.Show
        Else
            If Not in_list(cb, FormSelectItem.confirmed_node.Key) Then
                var_selected = True
                FormSelectItem.show_variable = var_selected
                FormSelectItem.Show
            End If
        End If
        If Not FormSelectItem.confirmed_node Is Nothing Then
            LabelNumInputs.Caption = LabelNumInputs.Caption - 1
            var_selected = (cb.value = SV_STR)
            If Not in_list(CbSecondAction, FormSelectItem.confirmed_node.Key) Then
                cb.AddItem FormSelectItem.confirmed_node.Key
                ignore_update = True
                cb.value = FormSelectItem.confirmed_node.Key
                ignore_update = False
            End If
            next_cb.Enabled = False
            If Not next_row_cb Is Nothing And Not next_row_cb.Enabled Then
                next_row_cb.AddItem "ABS"
                next_row_cb.AddItem "SQRT"
                next_row_cb.AddItem SV_STR
                next_row_cb.AddItem SE_STR
                next_row_cb.AddItem SR_OP1_STR
                next_row_cb.Enabled = True
                If cbFirstAction_3.Enabled Then next_row_cb.AddItem SR_OP2_STR
            End If
        End If
    End If
    update_counters
End Sub

Private Sub CbSecondAction_Click()
    handle_2st_action CbSecondAction, cbThirdAction, cbFirstAction_2
End Sub

Private Sub CbSecondAction_2_Click()
    handle_2st_action cbSecondAction_2, cbThirdAction_2, cbFirstAction_3
End Sub
Private Sub CbSecondAction_3_Click()
    handle_2st_action cbSecondAction_3, cbThirdAction_3, Nothing
End Sub


Private Sub handle_3rd_action(ByRef cb As ComboBox, ByRef next_row_cb As Variant)
    If ignore_update Then Exit Sub
    
    If cb.value = SC_STR Then
        FormSelectConst.Show
        If Not in_list(cb, FormSelectConst.InputVal.value) Then cb.AddItem FormSelectConst.InputVal.value
        ignore_update = True
        cb.value = FormSelectConst.InputVal.value
        ignore_update = False
        If Not next_row_cb Is Nothing Then
            If Not next_row_cb.Enabled Then
                next_row_cb.Enabled = True
                next_row_cb.AddItem "ABS"
                next_row_cb.AddItem "SQRT"
                next_row_cb.AddItem SV_STR
                next_row_cb.AddItem SE_STR
                next_row_cb.AddItem SR_OP1_STR
                If cbFirstAction_3.Enabled Then next_row_cb.AddItem SR_OP2_STR
            End If
        End If
    Else
        If cb.value = SV_STR Or cb.value = SE_STR Then
            var_selected = (cb.value = SV_STR)
            FormSelectItem.show_variable = var_selected
            FormSelectItem.Show
        Else
            If Not in_list(cb, FormSelectItem.confirmed_node.Key) Then
                var_selected = (cb.value = SV_STR)
                FormSelectItem.show_variable = var_selected
                FormSelectItem.Show
            End If
        End If
        If Not FormSelectItem.confirmed_node Is Nothing Then
            LabelNumInputs.Caption = LabelNumInputs.Caption - 1
            If Not in_list(cb, FormSelectItem.confirmed_node.Key) Then
                cb.AddItem FormSelectItem.confirmed_node.Key
                ignore_update = True
                cb.value = FormSelectItem.confirmed_node.Key
                ignore_update = False
            End If
            If Not next_row_cb Is Nothing Then
                If Not next_row_cb.Enabled Then
                    next_row_cb.AddItem "ABS"
                    next_row_cb.AddItem "SQRT"
                    next_row_cb.AddItem SV_STR
                    next_row_cb.AddItem SE_STR
                    next_row_cb.AddItem SR_OP1_STR
                    next_row_cb.Enabled = True
                    If cbFirstAction_3.Enabled Then next_row_cb.AddItem SR_OP2_STR
                End If
            End If
        End If
    End If
    update_counters
End Sub

Private Sub cbThirdAction_Click()
    handle_3rd_action cbThirdAction, cbFirstAction_2
End Sub

Private Sub cbThirdAction_2_Click()
    handle_3rd_action cbThirdAction_2, cbFirstAction_3
End Sub

Private Sub cbThirdAction_3_Click()
    handle_3rd_action cbThirdAction_3, Nothing
End Sub


Private Sub CommandButton1_Click()
    MsgBox "Here you will get a text box with information about how to specify conditions of a user defined event using this window"
End Sub

Private Sub UserForm_Initialize()
    CbFirstAction.AddItem "ABS"
    CbFirstAction.AddItem "SQRT"
    CbFirstAction.AddItem SV_STR
    CbFirstAction.AddItem SE_STR
    
    update_counters
End Sub


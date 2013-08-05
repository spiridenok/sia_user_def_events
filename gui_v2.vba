Private Sub ButtonAbs_Click()
    If Result.Caption = "<Select operation>" Then Result.Caption = "<?>"
    Result.Caption = Replace(Result.Caption, "<?>", "ABS")
    Result.Caption = Result.Caption & " <?> "
    
    ButtonPlus.Enabled = False
    ButtonMin.Enabled = False
    ButtonMul.Enabled = False
    ButtonDiv.Enabled = False
    
    ButtonVariable.Enabled = True
    ButtonEvent.Enabled = False
    ButtonConstant.Enabled = False
    
    ButtonSqrt.Enabled = True
    ButtonAbs.Enabled = True
    
    ButtonAnd.Enabled = False
    ButtonOr.Enabled = False
    ButtonGreater.Enabled = False
    ButtonLess.Enabled = False

    update_counters
End Sub

Private Sub ButtonAnd_Click()
    ButtonOr_Click
    
    Result.Caption = StrReverse(Replace(StrReverse(Result.Caption), StrReverse("OR"), StrReverse("AND"), , 1))
End Sub

Private Sub ButtonClearConditions_Click()
    Result.Caption = "<Select operation>"
    UserForm_Initialize
End Sub

Private Sub ButtonConstant_Click()
    FormSelectConst.Show

    Result.Caption = Replace(Result.Caption, "<?>", FormSelectConst.InputVal.value)
    Result.Caption = Result.Caption & " <?> "
    
    ButtonPlus.Enabled = True
    ButtonMin.Enabled = True
    ButtonMul.Enabled = True
    ButtonDiv.Enabled = True
    
    ButtonVariable.Enabled = True
    ButtonEvent.Enabled = False
    ButtonConstant.Enabled = False
    
    ButtonSqrt.Enabled = False
    ButtonAbs.Enabled = False
    
    ButtonAnd.Enabled = True
    ButtonOr.Enabled = True
    ButtonGreater.Enabled = True
    ButtonLess.Enabled = True
    
    update_counters
End Sub

Private Sub ButtonDiv_Click()
    ButtonPlus_Click
    
    Result.Caption = StrReverse(Replace(StrReverse(Result.Caption), StrReverse("+"), StrReverse("/"), , 1))
End Sub

Private Sub ButtonEvent_Click()
    FormSelectItem.show_variable = False
    FormSelectItem.Show
    
    If FormSelectItem.confirmed_node Is Nothing Then Exit Sub
    
    If Result.Caption = "<Select operation>" Then
        ' If this is the first variable in the whole condition
        Result.Caption = ""
        
        ButtonAnd.Enabled = True
        ButtonOr.Enabled = True
        ButtonGreater.Enabled = False
        ButtonLess.Enabled = False
    Else
'        ButtonAnd.Enabled = True
'        ButtonOr.Enabled = True
'        ButtonGreater.Enabled = True
'        ButtonLess.Enabled = True
    End If
    
    If InStr(Result.Caption, "<?>") > 0 Then
        Result.Caption = Replace(Result.Caption, "<?>", FormSelectItem.confirmed_node.Key)
    Else
        Result.Caption = Result.Caption & FormSelectItem.confirmed_node.Key
    End If
    Result.Caption = Result.Caption & " <?> "

    ButtonSqrt.Enabled = False
    ButtonAbs.Enabled = False
    
    ButtonPlus.Enabled = False
    ButtonMin.Enabled = False
    ButtonMul.Enabled = False
    ButtonDiv.Enabled = False
        
    ButtonVariable.Enabled = False
    ButtonEvent.Enabled = False
    ButtonConstant.Enabled = False
    
    update_counters
End Sub

Private Sub ButtonGreater_Click()
    Result.Caption = Replace(Result.Caption, "<?>", ">>")
    Result.Caption = Result.Caption & " <?> "
    
    ButtonPlus.Enabled = False
    ButtonMin.Enabled = False
    ButtonMul.Enabled = False
    ButtonDiv.Enabled = False
    
    ButtonVariable.Enabled = True
    ButtonEvent.Enabled = False
    ButtonConstant.Enabled = True
    
    ButtonSqrt.Enabled = True
    ButtonAbs.Enabled = True
    
    ButtonAnd.Enabled = False
    ButtonOr.Enabled = False
    ButtonGreater.Enabled = False
    ButtonLess.Enabled = False
    
    update_counters
End Sub

Private Sub ButtonLess_Click()
    ButtonGreater_Click
    
    Result.Caption = StrReverse(Replace(StrReverse(Result.Caption), StrReverse(">>"), StrReverse("<<"), , 1))
End Sub

Private Sub ButtonMin_Click()
    ButtonPlus_Click
    
    Result.Caption = StrReverse(Replace(StrReverse(Result.Caption), StrReverse("+"), StrReverse("-"), , 1))
End Sub

Private Sub ButtonMul_Click()
    ButtonPlus_Click
    
    Result.Caption = StrReverse(Replace(StrReverse(Result.Caption), StrReverse("+"), StrReverse("*"), , 1))

End Sub

Private Sub ButtonOr_Click()
    Result.Caption = Replace(Result.Caption, "<?>", "OR")
    Result.Caption = Result.Caption & " <?> "
    
    ButtonPlus.Enabled = False
    ButtonMin.Enabled = False
    ButtonMul.Enabled = False
    ButtonDiv.Enabled = False
    
    ButtonVariable.Enabled = True
    ButtonEvent.Enabled = True
    ButtonConstant.Enabled = False
    
    ButtonSqrt.Enabled = True
    ButtonAbs.Enabled = True
    
    update_counters
End Sub

Private Sub ButtonPlus_Click()
    Result.Caption = Replace(Result.Caption, "<?>", "+")
    Result.Caption = Result.Caption & " <?> "
    
    ButtonPlus.Enabled = False
    ButtonMin.Enabled = False
    ButtonMul.Enabled = False
    ButtonDiv.Enabled = False
    
    ButtonVariable.Enabled = True
    ButtonEvent.Enabled = False
    ButtonConstant.Enabled = True
    
    ButtonAnd.Enabled = False
    ButtonOr.Enabled = False
    ButtonGreater.Enabled = False
    ButtonLess.Enabled = False
    
    update_counters
End Sub

Private Sub ButtonSqrt_Click()
    If Result.Caption = "<Select operation>" Then Result.Caption = "<?>"
    Result.Caption = Replace(Result.Caption, "<?>", "SQRT")
    Result.Caption = Result.Caption & " <?> "
    
    ButtonPlus.Enabled = False
    ButtonMin.Enabled = False
    ButtonMul.Enabled = False
    ButtonDiv.Enabled = False
    
    ButtonVariable.Enabled = True
    ButtonEvent.Enabled = False
    ButtonConstant.Enabled = False
    
    ButtonSqrt.Enabled = False
    ButtonAbs.Enabled = True
    
    ButtonAnd.Enabled = False
    ButtonOr.Enabled = False
    ButtonGreater.Enabled = False
    ButtonLess.Enabled = False
    
    update_counters
End Sub

Private Sub ButtonUdevPropCancel_Click()
    FormUdevPropV2.Hide
End Sub

Private Sub ButtonUdevPropOk_Click()
    FormUdevPropV2.Hide
    MsgBox "You have selected a user defined event with some conditions! This event will be applied to your trace definition."
End Sub

Private Sub ButtonVariable_Click()
    FormSelectItem.show_variable = True
    FormSelectItem.Show
    
    If FormSelectItem.confirmed_node Is Nothing Then Exit Sub
    
    If Result.Caption = "<Select operation>" Then
        ' If this is the first variable in the whole condition
        Result.Caption = "<?>"
        
        ButtonAnd.Enabled = False
        ButtonOr.Enabled = False
    Else
        ButtonAnd.Enabled = True
        ButtonOr.Enabled = True
    End If
    
    Result.Caption = Replace(Result.Caption, "<?>", FormSelectItem.confirmed_node.Key)
    Result.Caption = Result.Caption & " <?> "
        
    ButtonGreater.Enabled = True
    ButtonLess.Enabled = True
        
    ButtonSqrt.Enabled = False
    ButtonAbs.Enabled = False
    
    ButtonPlus.Enabled = True
    ButtonMin.Enabled = True
    ButtonMul.Enabled = True
    ButtonDiv.Enabled = True
        
    ButtonVariable.Enabled = False
    ButtonEvent.Enabled = False
    ButtonConstant.Enabled = False
    
    update_counters
End Sub

Private Sub CommandButton1_Click()
    MsgBox "Here you will get a text box with information about how to specify conditions of a user defined event using this window"
End Sub

Private Sub Frame3_Click()

End Sub

Private Sub UserForm_Initialize()
    ButtonAnd.Enabled = False
    ButtonPlus.Enabled = False
    ButtonMin.Enabled = False
    ButtonMul.Enabled = False
    ButtonDiv.Enabled = False
    ButtonOr.Enabled = False
    ButtonGreater.Enabled = False
    ButtonLess.Enabled = False
    ButtonConstant.Enabled = False
    
    ButtonVariable.Enabled = True
    ButtonEvent.Enabled = True
    
    update_counters
End Sub

Sub update_counters()
    Dim op_counter As Integer
    Dim inp_counter As Integer
    
    op_counter = 10
    inp_counter = 10
    
    If Result.Caption <> "<Select operation>" Then
    
        Dim current_expression As String
        current_expression = Result.Caption
        current_expression = Replace(current_expression, " <?> ", "")
        
        s = Split(current_expression, " ")
        For i = 0 To UBound(s)
            If s(i) <> "" Then
                If s(i) = "+" Or s(i) = "*" Or s(i) = "/" Or s(i) = "-" Or s(i) = "OR" Or s(i) = "AND" Or s(i) = "SQRT" Or s(i) = "ABS" Or s(i) = ">> " Or s(i) = "<<" Then
                    op_counter = op_counter - 1
                Else
                    If InStr(1, s(i), ":") > 0 Then inp_counter = inp_counter - 1
                End If
            End If
        Next i
    End If
        
    LabelNumOperations.Caption = op_counter
    LabelNumInputs.Caption = inp_counter
End Sub


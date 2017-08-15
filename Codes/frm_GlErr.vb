Dim lng_DebugSwitch As Long 
 
Private Sub btn_Debug_Click() 
     
    lng_DebugSwitch = True 
     
    Unload Me 
End Sub 
 
Private Sub btn_Exit_Click() 
     
    lng_DebugSwitch = False 
 
    Unload Me 
End Sub 
 
Public Function z_Show(ByRef Err As Object) As Long 
     
    'Preset form 
    Me.txt_ErrCode.Caption = Err.Number 
    Me.txt_ErrDesc.Caption = Err.Description 
 
    If Application.Statusbar <> False Then 
        Me.txt_StatusBar.Caption = Application.Statusbar 
    End If 
     
    'Display modad dialog wait for action 
    Me.Show 
     
    'Event pressed 
    If lng_DebugSwitch = True Then 
        z_Show = lng_DebugSwitch 
     
        Stop 'break to code here 'Error handling dialog on error resume proceed with Ctrl+Shift+F8 
        Call shtATk.MacroFinish(True) 'optional - restore excel prerun state - skip if not needed 
         
    Else 
        Call shtATk.MacroFinish(True) 
        End 'exit all further running macros 
     
    End If 
     
End Function

Option Explicit

Private Const PAGE_EXECUTE_READWRITE = &H40
Private Declare PtrSafe Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As LongPtr, Source As LongPtr, ByVal Length As LongPtr)
Private Declare PtrSafe Function VirtualProtect Lib "kernel32" (lpAddress As LongPtr, ByVal dwSize As LongPtr, ByVal flNewProtect As LongPtr, lpflOldProtect As LongPtr) As LongPtr
Private Declare PtrSafe Function GetModuleHandleA Lib "kernel32" (ByVal lpModuleName As String) As LongPtr
Private Declare PtrSafe Function GetProcAddress Lib "kernel32" (ByVal hModule As LongPtr, ByVal lpProcName As String) As LongPtr
Private Declare PtrSafe Function DialogBoxParam Lib "user32" Alias "DialogBoxParamA" (ByVal hInstance As LongPtr, ByVal pTemplateName As LongPtr, ByVal hWndParent As LongPtr, ByVal lpDialogFunc As LongPtr, ByVal dwInitParam As LongPtr) As Integer
Private HookBytes(0 To 11) As Byte
Private OriginBytes(0 To 11) As Byte
Private pFunc As LongPtr
Private Flag As Boolean

Private Function GetPtr(ByVal Value As LongPtr) As LongPtr
    GetPtr = Value
End Function

Private Sub RecoverBytes()
    If Flag Then MoveMemory ByVal pFunc, ByVal VarPtr(OriginBytes(0)), 12
End Sub

Private Function MyDialogBoxParam(ByVal hInstance As LongPtr, ByVal pTemplateName As LongPtr, ByVal hWndParent As LongPtr, ByVal lpDialogFunc As LongPtr, ByVal dwInitParam As LongPtr) As Integer
    If pTemplateName = 4070 Then
        MyDialogBoxParam = 1
    Else
        RecoverBytes
        MyDialogBoxParam = DialogBoxParam(hInstance, pTemplateName, _
                   hWndParent, lpDialogFunc, dwInitParam)
        UnlockVBAProjec
    End If
End Function

Public Sub UnlockVBAProjects()
    Dim TmpBytes(0 To 11) As Byte
    Dim p As LongPtr, osi As Byte
    Dim OriginProtect As LongPtr
    #If Win64 Then
        osi = 1
    #Else
        osi = 0
    #End If
    pFunc = GetProcAddress(GetModuleHandleA("user32.dll"), "DialogBoxParamA")
    If VirtualProtect(ByVal pFunc, 12, PAGE_EXECUTE_READWRITE, OriginProtect) <> 0 Then
        MoveMemory ByVal VarPtr(TmpBytes(0)), ByVal pFunc, osi + 1
        If TmpBytes(osi) <> &HB8 Then
            MoveMemory ByVal VarPtr(OriginBytes(0)), ByVal pFunc, 12
            p = GetPtr(AddressOf MyDialogBoxParam)
            If osi Then HookBytes(0) = &H48
            HookBytes(osi) = &HB8
            osi = osi + 1
            MoveMemory ByVal VarPtr(HookBytes(osi)), ByVal VarPtr(p), 4 * osi
            HookBytes(osi + 4 * osi) = &HFF
            HookBytes(osi + 4 * osi + 1) = &HE0
            MoveMemory ByVal pFunc, ByVal VarPtr(HookBytes(0)), 12
            Flag = True
        End If
    End If
End Sub

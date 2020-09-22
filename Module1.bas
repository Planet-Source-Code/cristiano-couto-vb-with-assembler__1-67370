Attribute VB_Name = "Module1"
Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Function GetTickCount Lib "kernel32" () As Long

Public pVal As Long
Public pCount As Long

Public AsmOPCodes() As Byte

Public PtrpVal As String
Public PtrpCount As String

Sub ASMSum()
'The sum procedure in assembler is
'PUSHAD                         ;Save all registers in stack
'MOV ECX,DWORD PTR DS:[403028]  ;Move to ECX the value at pCounter variable
'MOV EAX,DWORD PTR DS:[403024]  ;Move to EAX the value at pVal variable
'ADD EAX,1                      ;Sum EAX+1
'SUB ECX,1                      ;Decrease ECX-1
'JNZ 00402A77                   ;Repeat if ECX<>0
'MOV DWORD PTR DS:[403024],EAX  ;Move to pCounter variable the value at ECX
'MOV DWORD PTR DS:[403028],ECX  ;Move to pVal variable the value at EAX
'POPAD                          ;Retrieve all registers from stack
'RETN                           ;Return to the normal code execution

pVal = 0
'Call the asm routine that we wrote at AsmOpCodes array
Call CallWindowProc(VarPtr(AsmOPCodes(1)), VarPtr(pVal), VarPtr(pVal), VarPtr(pVal), pVal)

End Sub

Function ReverseBytes(pVal As String) As String
Dim pTemp As String
For x = Len(pVal) - 1 To 1 Step -2
    pTemp = pTemp & Mid(pVal, x, 2)
Next
'Must be 8 bytes, if is less, fill with 0
ReverseBytes = pTemp & String(8 - Len(pTemp), "0")
End Function


Sub VBSum()
pVal = 0
Do While pCount <> 0
    pVal = pVal + 1
    pCount = pCount - 1
Loop
End Sub



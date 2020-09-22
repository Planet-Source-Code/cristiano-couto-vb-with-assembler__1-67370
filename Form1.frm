VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2895
   ClientLeft      =   2490
   ClientTop       =   1530
   ClientWidth     =   3405
   LinkTopic       =   "Form1"
   ScaleHeight     =   2895
   ScaleWidth      =   3405
   Begin VB.CommandButton Command1 
      Caption         =   "Go"
      Height          =   375
      Left            =   1200
      TabIndex        =   7
      Top             =   2160
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1200
      TabIndex        =   4
      Text            =   "10000000"
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   240
      TabIndex        =   9
      Top             =   1320
      Width           =   45
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   240
      TabIndex        =   8
      Top             =   720
      Width           =   45
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "0"
      Height          =   195
      Left            =   2850
      TabIndex        =   6
      Top             =   1080
      Width           =   90
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "0"
      Height          =   195
      Left            =   2835
      TabIndex        =   5
      Top             =   480
      Width           =   90
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Cycles:"
      Height          =   195
      Left            =   240
      TabIndex        =   3
      Top             =   1680
      Width           =   510
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Time"
      Height          =   195
      Left            =   2280
      TabIndex        =   2
      Top             =   120
      Width           =   345
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "ASM Sum (loop)"
      Height          =   195
      Left            =   240
      TabIndex        =   1
      Top             =   1080
      Width           =   1140
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "VB Sum (loop)"
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   1005
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim pTime As Long

pCount = CLng(Text1.Text)
pTime = GetTickCount
VBSum 'Call the sum procedure made with pure VB
Label5.Caption = GetTickCount - pTime & "ms"
Label7.Caption = pVal

pCount = CLng(Text1.Text)
pTime = GetTickCount
ASMSum 'Call the sum procedure made with assembler
Label6.Caption = GetTickCount - pTime & "ms"
Label8.Caption = pVal

End Sub


Private Sub Form_Load()
Dim StrOPCodes As String
Dim i As Long

PtrpVal = ReverseBytes(Hex(VarPtr(pVal))) 'Retrieve the pointer of pVal, convert to Hexadecimal and revert it to build the opcode
PtrpCount = ReverseBytes(Hex(VarPtr(pCount))) 'Retrieve the pointer of pCount, convert to Hexadecimal and revert it to build the opcode

'Write the entire asm routine into string
StrOPCodes = "608B0D" & PtrpCount & "A1" & PtrpVal & "83C00183E90175F8A3" & PtrpVal & "890D" & PtrpCount & "61C3"
ReDim AsmOPCodes(1 To Len(StrOPCodes) / 2)
i = 0
For x = 1 To Len(StrOPCodes) Step 2
    i = i + 1
    AsmOPCodes(i) = Val("&H" & Mid(StrOPCodes, x, 2)) 'Convert the hexadecimal opcodes into bytes in array
Next
End Sub



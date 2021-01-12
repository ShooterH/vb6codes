VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   495
      Left            =   3000
      TabIndex        =   4
      Top             =   960
      Width           =   735
   End
   Begin VB.CommandButton Command3 
      Caption         =   "product"
      Height          =   375
      Left            =   2160
      TabIndex        =   3
      Top             =   960
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      Caption         =   "sum"
      Height          =   495
      Left            =   3000
      TabIndex        =   2
      Top             =   240
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   360
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   1680
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "find"
      Height          =   615
      Left            =   2040
      TabIndex        =   0
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Function sum(nend As Integer, step As Integer) As Integer '等差数列求和
    If nend - step <= 0 Then
        sum = 0
        Exit Function
    End If
    sum = nend + sum(nend - step, step)
End Function
Function product(nend As Integer, step As Integer) As Long '等比数列求和
    If nend - step <= 0 Then
        product = 1
        Exit Function
    End If
    product = nend * product(nend - step, step)
End Function

Private Sub Command1_Click()
Dim a, b, c As Integer
Form1.Cls
'999内水仙花数
For a = 1 To 9
    For b = 0 To 9
        For c = 0 To 9
            If a ^ 3 + b ^ 3 + c ^ 3 = 100 * a + 10 * b + c Then Print 100 * a + 10 * b + c
        Next c
    Next b
Next a
End Sub

Private Sub Command2_Click()
Form1.Cls
Print sum(Text1.Text, 2)
End Sub

Private Sub Command3_Click()
Form1.Cls
Print product(Text1.Text, 1)
End Sub

Private Sub Command4_Click()
'教材p54Q2
Dim l As Double, t As Integer
l = 0.0001
Do
    l = l * 2
    n = n + 1
Loop Until l > 8848
Print n
Print l
End Sub

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
   Begin VB.CommandButton Command5 
      Caption         =   "IsPrimeNumOrNot"
      Height          =   495
      Left            =   2160
      TabIndex        =   5
      Top             =   1200
      Width           =   1935
   End
   Begin VB.CommandButton Command4 
      Caption         =   "P54Q2"
      Height          =   495
      Left            =   2160
      TabIndex        =   4
      Top             =   600
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      Caption         =   "product"
      Height          =   375
      Left            =   3240
      TabIndex        =   3
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "sum"
      Height          =   375
      Left            =   2160
      TabIndex        =   2
      Top             =   120
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "find"
      Height          =   495
      Left            =   3240
      TabIndex        =   0
      Top             =   600
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Function sum(nend As Integer, step As Integer) As Integer  '等差数列求和
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
Function IsPrimeNumOrNot(target As Integer) As Boolean '素数检测
    If target <= 2 Then
        IsPrimeNumOrNot = True
        Exit Function
    End If
    For i = 2 To Int(Sqr(target)) + 1                  '算法来自网络
        If target Mod i = 0 Then
            IsPrimeNumOrNot = False
            Exit Function
        End If
    Next i
    IsPrimeNumOrNot = True
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
'教材P54Q2
Dim l As Double, t As Integer
l = 0.0001
Do
    l = l * 2
    n = n + 1
Loop Until l > 8848
Print n
Print l
End Sub

Private Sub Command5_Click()
    If Text1.Text < 0 Then
        MsgBox ("Bad inpution")
    Else
        MsgBox (IsPrimeNumOrNot(Text1.Text))
    End If
End Sub


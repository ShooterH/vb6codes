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
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.CommandButton Command3 
      Caption         =   "product"
      Height          =   375
      Left            =   2520
      TabIndex        =   3
      Top             =   2040
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      Caption         =   "sum"
      Height          =   495
      Left            =   2520
      TabIndex        =   2
      Top             =   1200
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
      Left            =   2160
      TabIndex        =   0
      Top             =   360
      Width           =   735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Function sum(nend As Integer, step As Integer) As Integer
    If nend - step <= 0 Then
        sum = 0
        Exit Function
    End If
    sum = nend + sum(nend - step, step)
End Function
Function product(nend As Integer, step As Integer) As Long
    If nend - step <= 0 Then
        product = 1
        Exit Function
    End If
    product = nend * product(nend - step, step)
End Function

Private Sub Command1_Click()
Dim a, b, c As Integer
Form1.Cls
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

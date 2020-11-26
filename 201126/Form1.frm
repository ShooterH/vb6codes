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
   Begin VB.CommandButton Command1 
      Caption         =   "transform"
      Height          =   495
      Left            =   3240
      TabIndex        =   1
      Top             =   2040
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   720
      TabIndex        =   0
      Top             =   480
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Function transform(target As String) As String
    If Len(target) > 1 Then
        transform = "error"
        Exit Function
    End If
    If Asc(target) >= Asc("A") And Asc(target) <= Asc("Z") Then
        transform = Chr(Asc(target) + 32)
    Else
    If Asc(target) >= Asc("a") And Asc(target) <= Asc("z") Then
        transform = Chr(Asc(target) - 32)
    Else
        transform = "error"
    End If
    End If
End Function

Private Sub Command1_Click()
    Print transform(Text1.Text)
End Sub

Private Sub Command2_Click()
    Dim a As String
    a = "error"
    Print a
    Print Chr(Asc(a))
    Print Len(a)
End Sub

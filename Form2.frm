VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form2"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox texto1 
      Height          =   375
      Left            =   720
      TabIndex        =   0
      Text            =   "subir bajar 1"
      Top             =   360
      Width           =   1335
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub texto1_KeyDown(KeyCode As Integer, Shift As Integer)

    texto1.Text = subirBajar(KeyCode)
    

End Sub

Function subirBajar(tecla As Integer) As Variant

    If (tecla = 38) Then
        MsgBox "subir"
    End If
    If (tecla = 40) Then
        MsgBox "bajar"
    End If
    
    subirBajar = "funciona"

End Function

VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6675
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11865
   LinkTopic       =   "Form1"
   ScaleHeight     =   6675
   ScaleWidth      =   11865
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton guardarMesas 
      Caption         =   "guardar mesas"
      Height          =   375
      Left            =   4440
      TabIndex        =   19
      Top             =   120
      Width           =   3015
   End
   Begin VB.CommandButton Command11 
      Caption         =   "cambiar estado"
      Height          =   375
      Left            =   120
      TabIndex        =   18
      Top             =   3360
      Width           =   3855
   End
   Begin VB.CommandButton Command10 
      Caption         =   "yo elijo el nombre"
      Height          =   375
      Left            =   120
      TabIndex        =   17
      Top             =   2400
      Width           =   3855
   End
   Begin VB.CommandButton Command9 
      Caption         =   "todos los nombres"
      Height          =   375
      Left            =   120
      TabIndex        =   16
      Top             =   2040
      Width           =   3855
   End
   Begin VB.CommandButton Command8 
      Caption         =   "ver nombre"
      Height          =   375
      Left            =   2760
      TabIndex        =   15
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton Command7 
      Caption         =   "nombre mesa 3"
      Height          =   375
      Left            =   2760
      TabIndex        =   14
      Top             =   1320
      Width           =   1215
   End
   Begin VB.TextBox txtprecio 
      Height          =   285
      Left            =   120
      TabIndex        =   13
      Text            =   "3.3"
      Top             =   960
      Width           =   1455
   End
   Begin VB.CommandButton Command6 
      Caption         =   "destruir"
      Height          =   375
      Left            =   1440
      TabIndex        =   12
      Top             =   2880
      Width           =   1095
   End
   Begin VB.CommandButton Command5 
      Caption         =   "mesa 1"
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   2880
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "nombre mesa 2"
      Height          =   375
      Left            =   1440
      TabIndex        =   10
      Top             =   1320
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "ver nombre"
      Height          =   375
      Left            =   1440
      TabIndex        =   9
      Top             =   1680
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "ver nombre"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   1680
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000000&
      Caption         =   "nombre mesa 1"
      Height          =   375
      Left            =   120
      MaskColor       =   &H80000000&
      TabIndex        =   7
      Top             =   1320
      Width           =   1335
   End
   Begin VB.CommandButton ver2 
      Caption         =   "mer mesa 2"
      Height          =   375
      Left            =   2640
      TabIndex        =   6
      Top             =   480
      Width           =   1335
   End
   Begin VB.CommandButton guardarmesa2 
      Caption         =   "guardar2"
      Height          =   375
      Left            =   2640
      TabIndex        =   5
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton ver1 
      Caption         =   "ver mesa 1"
      Height          =   375
      Left            =   1680
      TabIndex        =   4
      Top             =   480
      Width           =   975
   End
   Begin VB.TextBox txtcantidad 
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Text            =   "3.3"
      Top             =   720
      Width           =   1455
   End
   Begin VB.TextBox txtdetalle 
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Text            =   "detalle"
      Top             =   480
      Width           =   1455
   End
   Begin VB.TextBox txtcodigo 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Text            =   "codigo"
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton guardarmesa1 
      Caption         =   "guardar1"
      Height          =   375
      Left            =   1680
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mesa1 As mesa
Dim mesa2 As mesa
Dim mesa3 As mesa





''''''''''''''''''
Dim Vcodigo() As String
Dim Vdetalle() As String
Dim Vcantidad() As Double
Dim Vprecio() As Double

Dim objetos As Collection




Private Sub Command1_Click()
Set mesa1 = New mesa
objetos.Add mesa1

mesa1.asignarNombre InputBox(nombre)

Me.colores

mesa1.leerTxt



End Sub

Private Sub Command10_Click()
    Dim nombre As String
    nombre = InputBox(nombre)
    
    Dim i As Integer
    
    MsgBox objetos.Count
    
    For i = 1 To objetos.Count
    
        If objetos.Item(i).getNombre = nombre Then
            MsgBox objetos.Item(i).getNombre
        End If
        
    Next i
End Sub

Private Sub Command11_Click()
Me.cambiarEstado (InputBox(ESTADO))
Me.colores

End Sub

Private Sub Command2_Click()
mesa1.getNombre

End Sub

Private Sub Command3_Click()
mesa2.getNombre

End Sub

Private Sub Command4_Click()
Set mesa2 = New mesa
objetos.Add mesa2

mesa2.asignarNombre InputBox(nombre)

Me.colores
End Sub

Private Sub Command5_Click()
Set mesa3 = New mesa

End Sub

Private Sub Command6_Click()
    destruir (InputBox("nombre mesa"))
    Me.colores
    
End Sub

Function destruir(mesa As String)
    
        'Set mesa1 = Nothing
        Dim nombre As String
    
    
    Dim i As Integer
    
    MsgBox objetos.Count
    
    For i = 1 To objetos.Count
    
        If objetos.Item(i).getNombre = mesa Then
        MsgBox i
        
            Set objetos.Item(i) = Nothing
            
            objetos.Remove (i)
        End If
        
        
    Next i
End Function

Private Sub Command7_Click()
Set mesa3 = New mesa
objetos.Add mesa3

mesa3.asignarNombre InputBox(nombre)

Me.colores
End Sub

Private Sub Command8_Click()
mesa3.getNombre

End Sub

Private Sub Command9_Click()
Dim i As Integer

MsgBox objetos.Count

For i = 1 To objetos.Count
    MsgBox objetos.Item(i).getNombre
Next i



End Sub

Private Sub Form_Load()





Set objetos = New Collection



Form2.Show



End Sub

Function colores()
    If mesa1 Is Nothing Then
    Command1.Caption = "CERRADO"
    Else
    Command1.Caption = mesa1.getEstado
    
    End If
    If mesa2 Is Nothing Then
    Command4.Caption = "CERRADO"
    Else
    Command4.Caption = mesa2.getEstado
    
    End If
    If mesa3 Is Nothing Then
    Command7.Caption = "CERRADO"
    Else
    Command7.Caption = mesa3.getEstado
    End If
    
    
    

End Function

Function cambiarEstado(nombre As String)
Dim i As Integer


For i = 1 To objetos.Count



    If objetos.Item(i).getNombre = nombre Then
        objetos.Item(i).asignarEstado ("ACTIVA")
        'MsgBox "activa"
    Else
        If objetos.Item(i).getEstado = "CERRADA" Then
        'MsgBox "cerrada"
        Else
        
        objetos.Item(i).asignarEstado ("PASIVA")
        'MsgBox "pasiva"
        End If
        
    End If
    
    'MsgBox objetos.Item(i).getNombre & " estado " & objetos.Item(i).getEstado
    
Next i
End Function

Private Sub guardarmesa1_Click()

Dim codigo As String
Dim detalle As String
Dim cantidad As Double

codigo = Me.txtcodigo.Text
detalle = Me.txtdetalle.Text
cantidad = Val(Me.txtcantidad.Text)



mesa1.cargarVector codigo, detalle, cantidad, Me.txtprecio.Text




End Sub



Private Sub guardarmesa2_Click()

Dim codigo As String
Dim detalle As String
Dim cantidad As Double

codigo = Me.txtcodigo.Text
detalle = Me.txtdetalle.Text
cantidad = Val(Me.txtcantidad.Text)



mesa2.cargarVector codigo, detalle, cantidad, Me.txtprecio.Text

End Sub

Private Sub guardarMesas_Click()
'Open archivo For Output As #nFich
    
    Dim i As Integer
      
    MsgBox App.Path
    
    
    
    For i = 1 To objetos.Count
    
    MsgBox objetos.Item(i).guardarTxt
    
    
      
      
        
        
    Next i
    
    
    
End Sub

Private Sub ver1_Click()
mesa1.getVectores

End Sub

Private Sub ver2_Click()
mesa2.getVectores

End Sub

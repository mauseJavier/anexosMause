VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MESA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"

Dim nombreMesa As String
Dim ESTADO As String


Dim Vcodigo() As String
Dim Vdetalle() As String
Dim Vcantidad() As Double
Dim Vprecio() As Double

Function extraerVcodigo() As Variant


    extraerVcodigo = Vcodigo()
       
    
    
End Function





 Sub getVectores()

      
    
     
    
    For i = 0 To (UBound(Vcodigo) - 1)
    
    
        
        MsgBox "Cantidad: " & Vcantidad(i) & " Detalle " & Vdetalle(i) & " precio " & Vprecio(i)
        
    

    Next i
    

    

End Sub

Function guardarTxt() As Variant

    Open App.Path + "/mesasTEMPORAL/" + nombreMesa + ".txt" For Output As #1
    
    For i = 0 To (UBound(Vcodigo) - 1)
    
        MsgBox "Codigo " & Vcodigo(i) & " Detalle " & Vdetalle(i) & " Cantidad " & Vcantidad(i)
        Write #1, Vcodigo(i); Vdetalle(i); Vcantidad(i)
        
    
    Next i
    
    Close #1
    
    guardarTxt = True
End Function

Function leerTxt() As Variant

    fileToCheck = App.Path + "/mesasTEMPORAL/" + nombreMesa + ".txt"
    Dim datos As String
    
    directoryName = Dir(fileToCheck)
    
    If directoryName = "" Then
    
            leerTxt = "No existe Archivo"
            
    Else
            'CUANDO EXISTE EL ARCHIVO
            Open App.Path + "/mesasTEMPORAL/" + nombreMesa + ".txt" For Input As #nombreMesa
            
                While Not EOF(nombreMesa)
                    Line Input #nombreMesa, datos
                    MsgBox datos
                Wend
    
    Close #nombreMesa
    
    End If
    
    




End Function


 Sub getVectoresOtro()

    For i = 0 To (UBound(Vcodigo) - 1)
    
    MsgBox "Codigo " & Vcodigo(i) & " Detalle " & Vdetalle(i) & " Cantidad " & Vcantidad(i)
    
    Next i
    

End Sub

 Sub cargarVector(codigo As String, detalle As String, cantidad As Double, precio As Double)

    
    
    Vcodigo(UBound(Vcodigo)) = codigo
    Vdetalle(UBound(Vdetalle)) = detalle
    Vcantidad(UBound(Vcantidad)) = cantidad
    Vprecio(UBound(Vprecio)) = precio
    
    ReDim Preserve Vcodigo(UBound(Vcodigo) + 1)
    ReDim Preserve Vdetalle(UBound(Vdetalle) + 1)
    ReDim Preserve Vcantidad(UBound(Vcantidad) + 1)
    ReDim Preserve Vprecio(UBound(Vprecio) + 1)
    

End Sub

Private Sub Class_Initialize()

'MsgBox "hola nueva clase"

ESTADO = "CERRADA"

ReDim Vcodigo(0)

ReDim Vdetalle(0)

ReDim Vcantidad(0)

ReDim Vprecio(0)

End Sub

Public Sub asignarNombre(nombre As String)
nombreMesa = nombre

End Sub
Public Sub asignarEstado(tipo As String)
ESTADO = tipo

End Sub
Function getEstado() As String


    getEstado = ESTADO
    

End Function

Function getNombre() As String


    getNombre = nombreMesa
    

End Function




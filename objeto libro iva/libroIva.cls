VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "libroIva"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Dim nombreArchivo As String
Dim vFacturas() As Variant

Function datoFactura(fecha As Date, tipo As String, pVenta As Integer, nComp As String, nCompHasta As String, cDocumento As String, Ndoc As String, nombreCliente As String, importe As String)

'Fecha de comprobante
'20210802
'Tipo de comprobante Según tabla Comprobantes Ventas
'https://www.afip.gob.ar/canasta-alimentaria/documentos/Tipos-de-comprobantes-de-ventas.pdf
'001
'Punto de venta
'00002
'Número de comprobante
'00000000000000008034
'numero de comprobante hasta
'00000000000000008034
'Código de documento del comprador Según tabla Documentos
'80
'Número de identificación del comprador Completar con ceros a izquierda
'00000000027273229979
'Apellido y nombre o denominación del comprador 30 espacios
'COLIPAN MARIA CLOTILDE
'Importe total de la operación 13 enteros 2 decimales sin punto decimal  LEN 15
'000000000224092
'Importe total de conceptos que no integran el
'precio neto gravado 13 enteros 2 decimales sin punto decimal LEN 15
'000000000000000
'Percepción a no categorizados 13 enteros 2 decimales sin punto decimal LEN 15
'000000000000000
'Importe de operaciones exentas 13 enteros 2 decimales sin punto decimal LEN 15
'000000000000000
'Importe de percepciones o pagos a cuenta de
'impuestos Nacionales 13 enteros 2 decimales sin punto decimal LEN 15
'000000000000000
'Importe de percepciones de Ingresos Brutos 13 enteros 2 decimales sin punto decimal LEN 15
'000000000000000
'Importe de percepciones impuestos Municipales 13 enteros 2 decimales sin punto decimal LEN 15
'000000000000000
'Importe impuestos internos 13 enteros 2 decimales sin punto decimal LEN 15
'000000000000000
'Código de moneda Según tabla Tipo de Monedas LEN 3
'https://www.afip.gob.ar/fe/documentos/TABLA%20MONEDAS%20V.0%20%2025082010.xls
'PES
'Tipo de cambio 4 enteros 6 decimales sin punto decimal LEN 10
'0001000000
'Cantidad de alícuotas de IVA LEN 1
'1
'Otros Tributos 13 enteros 2 decimales sin punto decimal LEN 15
'000000000000000
'Fecha de Vencimiento o Pago AAAAMMDD LEN 8
'00000000

    Dim f As String
    f = formatoFecha(fecha)
    Dim t As String
    t = formatoTipoComprobante(tipo)
    Dim p As String
    p = formatoPuntoVenta(pVenta)
    Dim n As String
    n = formatoNumeroComprobante(nComp)
    Dim nH As String
    nH = formatoNumeroComprobanteHasta(nCompHasta)
    Dim doc As String
    doc = formatoDocumento(cDocumento)
    Dim nd As String
    nd = formatoNumeroDocumento(Ndoc)
    Dim nc As String
    dc = formatoNombreCliente(nombreCliente)
    Dim im As String
    im = formatoImporte(importe)


    vFacturas(UBound(vFacturas)) = f & t & p & n & nH & doc & nd & dc & im & "000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000PES00010000001 00000000000000000000000"
    
    MsgBox "guardado " & vFacturas(UBound(vFacturas))
    
    ReDim Preserve vFacturas(UBound(vFacturas) + 1)

End Function

Function formatoImporte(importe As String) As String

    Dim t As String
    t = Replace(importe, ".", "") 'REMPLAZA EL PUNTO
    t = Replace(t, ",", "") ' REMPLAZA LA COMA
    
    Do While Len(t) < 15
     t = "0" & t
    Loop

    MsgBox " formatoImporte  " & Len(t) & " dato  " & t
    'MsgBox t
    
    formatoImporte = t
End Function

Function formatoNombreCliente(nombre As String) As String

    Dim t As String
    t = nombre
    
    Do While Len(t) < 30
     t = t & " "
    Loop

    'MsgBox " formatoNombreCliente  " & Len(t) & " dato  " & t
    'MsgBox t
    
    formatoNombreCliente = t
End Function

Function formatoNumeroDocumento(documento As String) As String

    Dim t As String
    t = Trim(Str(documento))
    
    Do While Len(t) < 20
     t = "0" & t
    Loop

    'MsgBox " formatoNumeroDocumento  " & Len(t) & " dato  " & t
    'MsgBox t
    
    formatoNumeroDocumento = t
End Function

Function formatoDocumento(documento As String) As String

    Dim t As String
    t = Trim(Str(documento))
    
    Do While Len(t) < 2
     t = "0" & t
    Loop

    'MsgBox " formatoDocumento  " & Len(t) & " dato  " & t
    'MsgBox t
    
    formatoDocumento = t
End Function

Function formatoNumeroComprobanteHasta(numero As String) As String

    Dim t As String
    t = Trim(Str(numero))
    
    Do While Len(t) < 20
     t = "0" & t
    Loop

    'MsgBox " formatoNumeroComprobanteHasta  " & Len(t) & " dato  " & t
    'MsgBox t
    
    formatoNumeroComprobanteHasta = t
End Function

Function formatoNumeroComprobante(numero As String) As String

    Dim t As String
    t = Trim(Str(numero))
    
    Do While Len(t) < 20
     t = "0" & t
    Loop

    'MsgBox " formatoNumeroComprobante  " & Len(t) & " dato  " & t
    'MsgBox t
    
    formatoNumeroComprobante = t
End Function

Function formatoPuntoVenta(punto As Integer) As String

    Dim t As String
    t = Trim(Str(punto))
    
    Do While Len(t) < 5
     t = "0" & t
    Loop

    'MsgBox "formatoPuntoVenta  " & Len(t) & " dato  " & t
    'MsgBox t
    
    formatoPuntoVenta = t
End Function
Function formatoTipoComprobante(tipo As String) As String

    Dim t As String
    t = Trim(Str(tipo))
    
    Do While Len(t) < 3
     t = "0" & t
    Loop

     'MsgBox "formatoTipoComprobante  " & Len(t) & " dato  " & t
    'MsgBox t
    
    formatoTipoComprobante = t
End Function
Function formatoFecha(fecha As Date) As String
    
    'MsgBox "formatoFecha " & Len(Format(Date, "yyyymmdd")) & " dato " & Format(fecha, "yyyymmdd")
    formatoFecha = Format(fecha, "yyyymmdd")

End Function
Private Sub Class_Initialize()

ReDim vFacturas(0)

End Sub

Function guardarTxt() As Variant

      

    Open App.Path + "/libroIva/" + nombreArchivo + ".txt" For Output As #1
    
    For i = 0 To (UBound(vFacturas) - 1)
    
        
        'Write #1, registro
        Print #1, vFacturas(i)
    
    
    Next i
    


    Close #1
    
    guardarTxt = True
End Function

Function Archivo(nombre As String)
    
        nombreArchivo = nombre

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

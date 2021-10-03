VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   9105
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15225
   LinkTopic       =   "Form1"
   ScaleHeight     =   9105
   ScaleWidth      =   15225
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "CODIGO INSERT"
      Height          =   855
      Left            =   3480
      TabIndex        =   16
      Top             =   240
      Width           =   2295
   End
   Begin VB.Frame Frame1 
      Caption         =   "Columnas"
      Height          =   3615
      Left            =   120
      TabIndex        =   9
      Top             =   1320
      Width           =   6855
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   3480
         TabIndex        =   2
         Text            =   "2"
         Top             =   720
         Width           =   2295
      End
      Begin VB.TextBox Text6 
         Height          =   375
         Left            =   3480
         TabIndex        =   6
         Text            =   "6"
         Top             =   2160
         Width           =   2295
      End
      Begin VB.TextBox Text5 
         Height          =   375
         Left            =   3480
         TabIndex        =   5
         Text            =   "5"
         Top             =   1800
         Width           =   2295
      End
      Begin VB.TextBox Text4 
         Height          =   375
         Left            =   3480
         TabIndex        =   4
         Text            =   "4"
         Top             =   1440
         Width           =   2295
      End
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   3480
         TabIndex        =   3
         Text            =   "3"
         Top             =   1080
         Width           =   2295
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   3480
         TabIndex        =   1
         Text            =   "1"
         Top             =   360
         Width           =   2295
      End
      Begin VB.Label Label6 
         Caption         =   "Columna Rubro"
         Height          =   375
         Left            =   240
         TabIndex        =   15
         Top             =   2160
         Width           =   1935
      End
      Begin VB.Label Label5 
         Caption         =   "Columna Proveedor"
         Height          =   375
         Left            =   240
         TabIndex        =   14
         Top             =   1800
         Width           =   1935
      End
      Begin VB.Label Label4 
         Caption         =   "Columna Precio Final"
         Height          =   375
         Left            =   240
         TabIndex        =   13
         Top             =   1440
         Width           =   1935
      End
      Begin VB.Label Label3 
         Caption         =   "Columna Stock"
         Height          =   375
         Left            =   240
         TabIndex        =   12
         Top             =   1080
         Width           =   1935
      End
      Begin VB.Label Label2 
         Caption         =   "Columna Detalle"
         Height          =   375
         Left            =   240
         TabIndex        =   11
         Top             =   720
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "Columna Codigo"
         Height          =   375
         Left            =   240
         TabIndex        =   10
         Top             =   360
         Width           =   1935
      End
   End
   Begin VB.CheckBox TITULOS 
      Caption         =   "TITULOS?"
      Height          =   375
      Left            =   600
      TabIndex        =   8
      Top             =   720
      Value           =   1  'Checked
      Width           =   1215
   End
   Begin VB.ListBox lista 
      Height          =   2010
      Left            =   120
      TabIndex        =   7
      Top             =   6360
      Width           =   14655
   End
   Begin VB.CommandButton Command1 
      Caption         =   "COMENZAR"
      Height          =   495
      Left            =   1440
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
   Begin MSComDlg.CommonDialog DialogoProductos 
      Left            =   480
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim CODIGO As String
Dim DETALLE As String
Dim STOCK As Double
Dim PRECIO As Double
Dim COSTO As Double
Dim IVA As Integer
Dim PROVEEDOR As String
Dim RUBRO As String


Dim ListaPrecio As Integer
Dim idPROVEEDOR As Integer
Dim idRUBRO As Integer




Private Sub Command1_Click()


'Me.DialogoProductos.Filter = "(.xlsx)|.xlsx|Todos los ficheros (.)|.|prueba (*.pdf)"
DialogoProductos.ShowOpen
If DialogoProductos.FileName <> "" Then
'MsgBox DialogoProductos.Filename
End If
'-------------------------
Dim excel_app As Object
Dim excel_sheet As Object
'Dim db As Database
Dim New_Value As String
Dim row As Integer
'-------------------------
Dim dato1 As String
Dim dato2 As String

'-----------------------------------
   ' Screen.MousePointer = vbHourglass
    DoEvents
 
    Set excel_app = CreateObject("Excel.Application")
 
'    excel_app.Visible = True
 
    excel_app.Workbooks.Open FileName:=DialogoProductos.FileName
 
    If Val(excel_app.Application.Version) >= 8 Then
        Set excel_sheet = excel_app.ActiveSheet
    Else
        Set excel_sheet = excel_app
    End If
 

    'SI POSEE TITULOS ARRANCA EN LA SEGUNDA FILA
    If Me.TITULOS.Value = Checked Then
        row = 2
        
    Else
        row = 1
    
    End If
    
        MsgBox ("INICIO PROCESO  ARCHIVO: " & DialogoProductos.FileName)
         
         Do
 
        'COLUMNA 1
        dato1 = Trim$(excel_sheet.Cells(row, 1))
        'MsgBox ("TAMAÑO DE DATO " & Len(dato1) & " este es el dato : " & dato1)
        
         CODIGO = Trim$(excel_sheet.Cells(row, Val(Me.Text1.Text)))
         DETALLE = Trim$(excel_sheet.Cells(row, Val(Me.Text2.Text)))
         STOCK = Val(Trim$(excel_sheet.Cells(row, Val(Me.Text3.Text))))
         PRECIO = Val(Trim$(excel_sheet.Cells(row, Val(Me.Text4.Text))))
         COSTO = PRECIO
         IVA = 0
         PROVEEDOR = Trim$(excel_sheet.Cells(row, Val(Me.Text5.Text)))
         RUBRO = Trim$(excel_sheet.Cells(row, Val(Me.Text6.Text)))
        
        Me.lista.AddItem (CODIGO & "  |  " & DETALLE & "  |  " & STOCK & "  |  " & PRECIO & "  |  " & PROVEEDOR & "  |  " & RUBRO)
        
        
         '--------------------------------------------------------------
         Set rs = cn.Execute(" SELECT * FROM INVENTARIO WHERE  a.codigo ='" & CODIGO & "'")
         '--------------------------------------------------------------
         If rs.EOF = True Then
         'no existe y lo agrega
         
         'los metodos debuelven el id
             ListaPrecio = DameListaPrecio()
             idPROVEEDOR = DameProveedor(PROVEEDOR)
             idRUBRO = DameRubro(RUBRO)
             
             MsgBox ListaPrecio
             MsgBox idPROVEEDOR
             MsgBox idRUBRO
             
             

         cn.Execute ("insert into  inventario (codigo,detalle,costo,lista_precio,stock,proveedor,rubro,ubicacion,iva,fecha_modificacion,stock_max,stock_min)" _
         & "values('" & CODIGO & "', '" & DETALLE & "','" & COSTO & "','" & ListaPrecio & "','" & STOCK & "','" & PROVEEDOR & "','" & RUBRO & "','" & "1" & "','" & IVA & "','" & Date & "','" & "50" & "','" & "3" & "')")
         Else
         'si ya existe
         cn.Execute ("update inventario set " _
                  & "  detalle ='" & DETALLE & "'," _
                  & "  Stock = '" & STOCK & "'," _
                   & " Costo = '" & Round(COSTO, Module1.N_DECIMALES) & "'," _
                  & "  rubro='" & RUBRO & "'," _
                  & "  fecha_modificacion= '" & Date & "'," _
                 & "   where Codigo = '" & CODIGO & "'")
         End If
         '--------------------------------------------------------------
        
       
         
        If Len(dato1) = 0 Then Exit Do
         row = row + 1
         Loop
         
         MsgBox ("PROCESO TERMINADO")

End Sub


Function DameListaPrecio()

DameListaPrecio = 0
End Function


Function DameProveedor()

DameProveedor = 0
End Function
Function DameRubro()

DameRubro = 0
End Function

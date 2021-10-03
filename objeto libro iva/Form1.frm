VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6585
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9540
   LinkTopic       =   "Form1"
   ScaleHeight     =   6585
   ScaleWidth      =   9540
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox importe 
      Height          =   375
      Left            =   240
      TabIndex        =   13
      Text            =   "2240.92"
      Top             =   3840
      Width           =   2295
   End
   Begin VB.TextBox nombreCliente 
      Height          =   375
      Left            =   240
      TabIndex        =   12
      Text            =   "COLIPAN MARIA CLOTILDE"
      Top             =   3480
      Width           =   2535
   End
   Begin VB.TextBox numeroDocumento 
      Height          =   375
      Left            =   240
      TabIndex        =   11
      Text            =   "35833716"
      Top             =   3120
      Width           =   2055
   End
   Begin VB.TextBox cDocumento 
      Height          =   405
      Left            =   240
      TabIndex        =   10
      Text            =   "80"
      Top             =   2760
      Width           =   1455
   End
   Begin VB.TextBox nComprobanteHasta 
      Height          =   375
      Left            =   240
      TabIndex        =   7
      Text            =   "8034"
      Top             =   2400
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Pruebas"
      Height          =   375
      Left            =   2160
      TabIndex        =   6
      Top             =   600
      Width           =   2175
   End
   Begin VB.TextBox nComprobante 
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Text            =   "8034"
      Top             =   2040
      Width           =   1935
   End
   Begin VB.TextBox pVenta 
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Text            =   "2"
      Top             =   1680
      Width           =   1455
   End
   Begin VB.TextBox tipoComp 
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Text            =   "1"
      Top             =   1320
      Width           =   1455
   End
   Begin VB.TextBox nombreArchivo 
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Text            =   "nombre archivo"
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton btnComenzar 
      Caption         =   "crear txt"
      Height          =   375
      Left            =   2160
      TabIndex        =   1
      Top             =   240
      Width           =   2175
   End
   Begin VB.TextBox fecha 
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Text            =   "fecha"
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "COMP HASTA"
      Height          =   375
      Left            =   2280
      TabIndex        =   9
      Top             =   2400
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "COMP DESDE"
      Height          =   375
      Left            =   2280
      TabIndex        =   8
      Top             =   2040
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim libroIva As libroIva


Private Sub btnComenzar_Click()

libroIva.Archivo Me.nombreArchivo.Text

libroIva.datoFactura Me.fecha.Text, Me.tipoComp.Text, Me.pVenta.Text, Me.nComprobante.Text, Me.nComprobanteHasta.Text, Me.cDocumento.Text, Me.numeroDocumento.Text, Me.nombreCliente.Text, Me.importe.Text

libroIva.guardarTxt


End Sub

Private Sub Command1_Click()
'libroIva.formatoPuntoVenta Me.pVenta.Text
'MsgBox (libroIva.formatoTipoComprobante(Me.tipoComp.Text))

MsgBox (libroIva.formatoNumeroComprobante(Me.nComprobante.Text))

End Sub

Private Sub Form_Load()
Set libroIva = New libroIva


Me.fecha.Text = Date






End Sub

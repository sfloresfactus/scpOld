VERSION 5.00
Begin VB.Form ProductosImportar 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Importacion de Productos"
   ClientHeight    =   1890
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2640
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1890
   ScaleWidth      =   2640
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnAccion 
      Caption         =   "Importar"
      Height          =   375
      Left            =   840
      TabIndex        =   0
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "archivo productos.csv, separado por ; (archivo productos.csv, en carpeta de aplicacion)"
      Height          =   855
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2415
   End
End
Attribute VB_Name = "ProductosImportar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Db As Database, Rs As Recordset
Private salida(99) As String, totalSalida As Integer
Private Sub Form_Load()
' abre archivos
Set Db = OpenDatabase(data_file)
Set Rs = Db.OpenRecordset("Productos")
Rs.Index = "Codigo"
End Sub
Private Sub btnAccion_Click()
Dim i As Integer
Dim reg

Open App.Path & "\productos.csv" For Input As #1
With Rs
Do While Not EOF(1)

    Line Input #1, reg
    split reg, ";"
    
    .Seek "=", salida(7) ' codigo producto
    If .NoMatch Then
        ' producto nio existe
        .AddNew
        ![Codigo] = salida(7)
        ![Descripcion] = salida(8)
        ![unidad de medida] = UCase(Left(salida(9), 3))
        .Update
    End If
    
'    For i = 1 To totalSalida
'        Debug.Print "|" & salida(i) & "|"
'    Next
    
Loop
End With

MsgBox "Listo"

End Sub
Private Sub split(cadena, separador)
' transforma linea de string con separadores,
' en arreglo
' ejemplo:
' linea = "plano | rev | marca | cantidad"
' txt=explode(linea, "|")
' queda como:
' txt[1]=plano
' txt[2]=rev, etc
Dim txt As String
Dim i As Integer, indice As Integer
Dim largo As Integer
Dim pca As Integer ' posicion de caracter anterior

txt = Trim(cadena)
largo = Len(txt)
indice = 0
pca = 0

For i = 1 To largo
    If Mid(txt, i, 1) = separador Then
        indice = indice + 1
        salida(indice) = Mid(txt, pca + 1, i - pca - 1)
        pca = i
    End If
Next
If largo > pca Then
    indice = indice + 1
    salida(indice) = Mid(txt, pca + 1)
End If

totalSalida = indice

End Sub

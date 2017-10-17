VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form Familiares 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Familiares"
   ClientHeight    =   3945
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5070
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3945
   ScaleWidth      =   5070
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox telefono 
      Height          =   300
      Left            =   3720
      TabIndex        =   11
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Frame FrameParentesco 
      Caption         =   "Parentesco"
      ForeColor       =   &H00FF0000&
      Height          =   1095
      Left            =   240
      TabIndex        =   12
      Top             =   2040
      Width           =   2295
      Begin VB.OptionButton OpParentesco 
         Caption         =   "Hijo(a)"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   14
         Top             =   720
         Width           =   1335
      End
      Begin VB.OptionButton OpParentesco 
         Caption         =   "Cónyuge"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   13
         Top             =   360
         Width           =   1215
      End
   End
   Begin MSMask.MaskEdBox fnacimiento 
      Height          =   300
      Left            =   1440
      TabIndex        =   9
      Top             =   1680
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   529
      _Version        =   327680
      PromptChar      =   "_"
   End
   Begin VB.TextBox nombres 
      Height          =   300
      Left            =   1440
      TabIndex        =   7
      Top             =   1200
      Width           =   3375
   End
   Begin VB.TextBox materno 
      Height          =   300
      Left            =   2760
      TabIndex        =   5
      Top             =   840
      Width           =   2055
   End
   Begin VB.TextBox paterno 
      Height          =   300
      Left            =   240
      TabIndex        =   3
      Top             =   840
      Width           =   2175
   End
   Begin VB.TextBox rut 
      Height          =   300
      Left            =   840
      TabIndex        =   1
      Top             =   240
      Width           =   1335
   End
   Begin VB.CommandButton btnGrabar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   1920
      TabIndex        =   18
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Frame FrameSexo 
      Caption         =   "Sexo"
      ForeColor       =   &H00FF0000&
      Height          =   1095
      Left            =   2760
      TabIndex        =   15
      Top             =   2040
      Width           =   2055
      Begin VB.OptionButton OpSexo 
         Caption         =   "Femenino"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   1
         Left            =   360
         TabIndex        =   17
         Top             =   720
         Width           =   1095
      End
      Begin VB.OptionButton OpSexo 
         Caption         =   "Masculino"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   16
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Label lbl 
      Caption         =   "Teléfono"
      ForeColor       =   &H00FF0000&
      Height          =   300
      Index           =   6
      Left            =   2760
      TabIndex        =   10
      Top             =   1680
      Width           =   735
   End
   Begin VB.Label lbl 
      Caption         =   "Fecha Nac."
      ForeColor       =   &H00FF0000&
      Height          =   300
      Index           =   4
      Left            =   240
      TabIndex        =   8
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Label lbl 
      Caption         =   "Nombres"
      ForeColor       =   &H00FF0000&
      Height          =   300
      Index           =   3
      Left            =   240
      TabIndex        =   6
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label lbl 
      Caption         =   "Apellido Materno"
      ForeColor       =   &H00FF0000&
      Height          =   300
      Index           =   2
      Left            =   2760
      TabIndex        =   4
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label lbl 
      Caption         =   "Apellido Paterno"
      ForeColor       =   &H00FF0000&
      Height          =   300
      Index           =   1
      Left            =   240
      TabIndex        =   2
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label lbl 
      Caption         =   "RUT"
      ForeColor       =   &H00FF0000&
      Height          =   300
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   495
   End
End
Attribute VB_Name = "Familiares"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_Rut As String, m_RutCarga As String, m_Accion As String
Private Rs As Recordset
Private d As Variant
Public Property Let Accion(ByVal NuevoValor As String)
m_Accion = NuevoValor
End Property
Public Property Let ruttrabajador(ByVal NuevoValor As String)
m_Rut = NuevoValor
End Property
Public Property Let RutCarga(ByVal NuevoValor As String)
m_RutCarga = NuevoValor
End Property
Public Property Let RecSet(RecZet As Recordset)
Set Rs = RecZet
End Property
Private Sub Form_Load()

If m_Accion = "agregar" Then
    rut.Enabled = True
Else ' editar
    rut.Text = m_RutCarga
    rut.Enabled = False
    With Rs
    .Seek "=", m_Rut, m_RutCarga
    If .NoMatch Then
        MsgBox "Rut No Existe"
    Else
        paterno.Text = !paterno
        materno.Text = !materno
        nombres.Text = !nombres
        
        fnacimiento.Text = ![Fecha Nacimiento]
        telefono.Text = !telefono
        
        If !Sexo = "M" Then
            OpSexo(0).Value = True
        Else
            OpSexo(1).Value = True
        End If
        
        If !Parentesco = "C" Then
            OpParentesco(0).Value = True
        Else
            OpParentesco(1).Value = True
        End If
        
'        If !Carga = False Then
'            Carga.Value = 0
'            finicio.Text = ""
'            ftermino.Text = ""
'        Else
'            Carga.Value = 1
'            finicio.Text = ![Fecha Inicio]
'            ftermino.Text = ![Fecha Vencimiento]
'        End If
        
    End If
    End With
    
End If

rut.MaxLength = 10
paterno.MaxLength = 20
materno.MaxLength = 20
nombres.MaxLength = 40

End Sub
Private Sub Carga_Click()
'If Carga.Value = 0 Then
'    finicio.Enabled = False
'    ftermino.Enabled = False
'Else
'    finicio.Enabled = True
'    ftermino.Enabled = True
'End If
End Sub
Private Function Valida() As Boolean
Valida = False
m_RutCarga = rut.Text

If Rut_Verifica(m_RutCarga) = False Then
    MsgBox "RUT no Válido"
    rut.SetFocus
    Exit Function
End If

m_RutCarga = Rut_Formato(rut.Text)

m_Rut = Rut_Formato(m_Rut)

' nuevo
If m_Accion = "agregar" Then
    Rs.Seek "=", m_Rut, m_RutCarga
    If Not Rs.NoMatch Then
        MsgBox "RUT Ya Existe"
        rut.SetFocus
        Exit Function
    End If
Else ' editar
    Rs.Seek "=", m_Rut, m_RutCarga
    If Rs.NoMatch Then
        MsgBox "RUT NO Existe"
        rut.SetFocus
        Exit Function
    End If
End If

If Trim(paterno.Text) = "" Then
    MsgBox "Debe digitar Apellido Paterno"
    paterno.SetFocus
    Exit Function
End If

If Trim(nombres.Text) = "" Then
    MsgBox "Debe digitar Nombres"
    nombres.SetFocus
    Exit Function
End If

If fnacimiento.Text = "__/__/__" Or fnacimiento.Text = "" Then
    MsgBox "Debe Digitar Fecha Nacimiento"
    fnacimiento.SetFocus
    Exit Function
End If

If OpParentesco(0).Value = False And OpParentesco(1).Value = False Then
    MsgBox "Debe Digitar Parentesco"
    Exit Function
End If

If OpSexo(0).Value = False And OpSexo(1).Value = False Then
    MsgBox "Debe Digitar Sexo"
    Exit Function
End If

'If Carga.Value = 0 Then
'    ' ok
'Else
'    If finicio.Text = "__/__/__" Or finicio.Text = "" Then
'        MsgBox "Debe Digitar Fecha Inicio"
'        finicio.SetFocus
'        Exit Function
'    End If
'    If ftermino.Text = "__/__/__" Or ftermino.Text = "" Then
'        MsgBox "Debe Digitar Fecha Término"
'        ftermino.SetFocus
'        Exit Function
'    End If
'End If

Valida = True

End Function
Private Sub btnGrabar_Click()
If Valida Then
    
    If m_Accion = "agregar" Then ' nuevo
    
        Rs.Seek "=", m_Rut, m_RutCarga
        If Rs.NoMatch Then
            ' ok
            With Rs
            .AddNew
            !rut = m_Rut
            ![Rut Carga] = m_RutCarga
            !paterno = paterno.Text
            !materno = materno.Text
            !nombres = nombres.Text
            ![Fecha Nacimiento] = fnacimiento.Text
            !telefono = telefono.Text
            !Parentesco = IIf(OpParentesco(0).Value = True, "C", "H")
            !Sexo = IIf(OpSexo(0).Value = True, "M", "F")
            
'            If Carga.Value = 1 Then
'                !Carga = True
'                ![Fecha Inicio] = finicio.Text
'                ![Fecha Vencimiento] = ftermino.Text
'            Else
'                !Carga = False
'            End If
            
            .Update
            End With
        Else
            MsgBox "RUT Ya Existe"
            Exit Sub
        End If
    
        Unload Me
        
    Else
        Rs.Seek "=", m_Rut, m_RutCarga
        If Not Rs.NoMatch Then
            With Rs
            .Edit
            !paterno = paterno.Text
            !materno = materno.Text
            !nombres = nombres.Text
            ![Fecha Nacimiento] = fnacimiento.Text
            !telefono = telefono.Text
            !Parentesco = IIf(OpParentesco(0).Value = True, "C", "H")
            !Sexo = IIf(OpSexo(0).Value = True, "M", "F")
            
'            If Carga.Value = 1 Then
'                !Carga = True
'                ![Fecha Inicio] = finicio.Text
'                ![Fecha Vencimiento] = ftermino.Text
'            Else
'                !Carga = False
'            End If
            .Update
            End With
        Else
            MsgBox "RUT NO Existe"
            Exit Sub
        End If
        
        Unload Me
        
    End If
    
End If
End Sub
Private Sub rut_KeyPress(KeyAscii As Integer)
Enter KeyAscii
End Sub
Private Sub paterno_KeyPress(KeyAscii As Integer)
Enter KeyAscii
End Sub
Private Sub materno_KeyPress(KeyAscii As Integer)
Enter KeyAscii
End Sub
Private Sub nombres_KeyPress(KeyAscii As Integer)
Enter KeyAscii
End Sub
Private Sub fnacimiento_KeyPress(KeyAscii As Integer)
Enter KeyAscii
End Sub
Private Sub fnacimiento_LostFocus()
d = Fecha_Valida(fnacimiento)
End Sub
Private Sub telefono_KeyPress(KeyAscii As Integer)
Enter KeyAscii
End Sub
Private Sub finicio_KeyPress(KeyAscii As Integer)
Enter KeyAscii
End Sub
Private Sub finicio_LostFocus()
'If Carga.Value = 1 Then
'    d = Fecha_Valida(finicio)
'End If
End Sub
Private Sub ftermino_KeyPress(KeyAscii As Integer)
Enter KeyAscii
End Sub
Private Sub ftermino_LostFocus()
'If Carga.Value = 1 Then
'    d = Fecha_Valida(ftermino)
'End If
End Sub

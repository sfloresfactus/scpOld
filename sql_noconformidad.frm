VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "msinet.ocx"
Begin VB.Form sql_noconformidad 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registro de No Conformidad"
   ClientHeight    =   9450
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8325
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9450
   ScaleWidth      =   8325
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Crystal.CrystalReport cr 
      Left            =   7320
      Top             =   480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   7800
      Top             =   600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.Toolbar Toolbar 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8325
      _ExtentX        =   14684
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      ImageList       =   "ImageList"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   9
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Nuevo INC"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Modificar INC"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Eliminar INC"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir INC"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "DesHacer"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Grabar "
            ImageIndex      =   6
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Clientes"
            ImageIndex      =   7
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.Frame Frame_ER 
      Height          =   615
      Left            =   4320
      TabIndex        =   5
      Top             =   450
      Width           =   2895
      Begin VB.OptionButton Op_ER 
         Caption         =   "Receptor"
         Height          =   255
         Index           =   1
         Left            =   1440
         TabIndex        =   7
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton Op_ER 
         Caption         =   "Emisior"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   6
         Top             =   240
         Width           =   855
      End
   End
   Begin InetCtlsObjects.Inet Inet 
      Left            =   6960
      Top             =   600
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin TabDlg.SSTab SSTab 
      Height          =   8175
      Left            =   120
      TabIndex        =   8
      Top             =   1200
      Width           =   8085
      _ExtentX        =   14261
      _ExtentY        =   14420
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Emisor"
      TabPicture(0)   =   "sql_noconformidad.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lbl(3)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lbl(2)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label6"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lbl(5)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblGerente"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "e_FechaCierre"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "frame(0)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Frame1"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "frame(1)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "e_Cerrado"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "e_clave"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "CbEmisores"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "CbGerencias"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "CbAreas"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).ControlCount=   15
      TabCaption(1)   =   "Receptor"
      TabPicture(1)   =   "sql_noconformidad.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lbl(4)"
      Tab(1).Control(1)=   "lbl(8)"
      Tab(1).Control(2)=   "lbl(9)"
      Tab(1).Control(3)=   "lbl(6)"
      Tab(1).Control(4)=   "Label1"
      Tab(1).Control(5)=   "lblFechaPrimeraRespuesta"
      Tab(1).Control(6)=   "frame_r_causas"
      Tab(1).Control(7)=   "frame(3)"
      Tab(1).Control(8)=   "Frame5"
      Tab(1).Control(9)=   "CbEncargado"
      Tab(1).Control(10)=   "r_clave"
      Tab(1).Control(11)=   "encargado_clave"
      Tab(1).Control(12)=   "r_CostoEstimado"
      Tab(1).Control(13)=   "r_archivo"
      Tab(1).Control(14)=   "Frame_accionesCorrectivas"
      Tab(1).Control(15)=   "Frame_comentarios"
      Tab(1).ControlCount=   16
      Begin VB.Frame Frame_comentarios 
         Caption         =   "Comentarios"
         Height          =   1095
         Left            =   -74760
         TabIndex        =   74
         Top             =   6360
         Width           =   7455
         Begin VB.TextBox r_comentarios 
            Height          =   495
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   75
            Top             =   240
            Width           =   7095
         End
      End
      Begin VB.Frame Frame_accionesCorrectivas 
         Caption         =   "Acciones Correctivas"
         Height          =   1335
         Left            =   -74760
         TabIndex        =   68
         Top             =   2520
         Width           =   7455
         Begin VB.TextBox r_ac 
            Height          =   615
            Left            =   240
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   73
            Top             =   600
            Width           =   6975
         End
         Begin VB.CheckBox op_ac 
            Caption         =   "Otra"
            Height          =   375
            Index           =   3
            Left            =   4200
            TabIndex        =   72
            Top             =   240
            Width           =   1095
         End
         Begin VB.CheckBox op_ac 
            Caption         =   "RFI"
            Height          =   375
            Index           =   2
            Left            =   3240
            TabIndex        =   71
            Top             =   240
            Width           =   735
         End
         Begin VB.CheckBox op_ac 
            Caption         =   "Reparación"
            Height          =   375
            Index           =   1
            Left            =   1680
            TabIndex        =   70
            Top             =   240
            Width           =   1215
         End
         Begin VB.CheckBox op_ac 
            Caption         =   "Reproceso"
            Height          =   375
            Index           =   0
            Left            =   240
            TabIndex        =   69
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.Frame r_archivo 
         Height          =   855
         Left            =   -74760
         TabIndex        =   60
         Top             =   3840
         Width           =   7455
         Begin VB.CommandButton btnREliminarAdjuntos 
            Caption         =   "Eliminar Adjuntos"
            Height          =   495
            Left            =   1320
            TabIndex        =   81
            Top             =   240
            Width           =   855
         End
         Begin VB.CommandButton btnRPdfVer 
            Caption         =   "Ver PDF"
            Height          =   300
            Left            =   2400
            TabIndex        =   78
            Top             =   480
            Width           =   1095
         End
         Begin VB.TextBox lblRAdjuntos 
            Height          =   495
            Left            =   3720
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   77
            Top             =   240
            Width           =   3495
         End
         Begin VB.CommandButton BtnRFotosVer 
            Caption         =   "Ver Foto(s)"
            Height          =   300
            Left            =   2400
            TabIndex        =   62
            Top             =   120
            Width           =   1095
         End
         Begin VB.CommandButton btnRAdjuntar 
            Caption         =   "Adjuntar Archivo"
            Height          =   495
            Left            =   240
            TabIndex        =   61
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.TextBox r_CostoEstimado 
         Height          =   285
         Left            =   -73440
         TabIndex        =   55
         Top             =   7560
         Width           =   1095
      End
      Begin VB.TextBox encargado_clave 
         Height          =   300
         Left            =   -68880
         TabIndex        =   45
         Top             =   1080
         Width           =   975
      End
      Begin VB.TextBox r_clave 
         Height          =   300
         Left            =   -68880
         TabIndex        =   43
         Top             =   720
         Width           =   975
      End
      Begin VB.ComboBox CbAreas 
         Height          =   315
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   1320
         Width           =   5655
      End
      Begin VB.ComboBox CbGerencias 
         Height          =   315
         Left            =   2880
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   960
         Width           =   3375
      End
      Begin VB.ComboBox CbEmisores 
         Height          =   315
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   480
         Width           =   3855
      End
      Begin VB.TextBox e_clave 
         Height          =   300
         Left            =   6600
         TabIndex        =   12
         Top             =   480
         Width           =   975
      End
      Begin VB.CheckBox e_Cerrado 
         Caption         =   "Cerrado"
         Height          =   255
         Left            =   360
         TabIndex        =   37
         Top             =   7560
         Width           =   975
      End
      Begin VB.ComboBox CbEncargado 
         Height          =   315
         Left            =   -74640
         Style           =   2  'Dropdown List
         TabIndex        =   41
         Top             =   1200
         Width           =   3855
      End
      Begin VB.Frame Frame5 
         Caption         =   "Fecha Comp"
         Height          =   1335
         Left            =   -74760
         TabIndex        =   48
         Top             =   4920
         Width           =   1575
         Begin MSMask.MaskEdBox r_fecha3 
            Height          =   300
            Left            =   120
            TabIndex        =   51
            Top             =   960
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   529
            _Version        =   327680
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox r_fecha2 
            Height          =   300
            Left            =   120
            TabIndex        =   50
            Top             =   600
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   529
            _Version        =   327680
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox r_fecha1 
            Height          =   300
            Left            =   120
            TabIndex        =   49
            Top             =   240
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   529
            _Version        =   327680
            PromptChar      =   "_"
         End
      End
      Begin VB.Frame frame 
         Caption         =   "Acciones Preventivas (Explique)"
         Height          =   1455
         Index           =   3
         Left            =   -73080
         TabIndex        =   52
         Top             =   4800
         Width           =   5775
         Begin VB.TextBox r_acciones 
            Height          =   975
            Left            =   240
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   53
            Top             =   240
            Width           =   5295
         End
      End
      Begin VB.Frame frame_r_causas 
         Caption         =   "Investigación de las Causas (Explique el (los) motivo(s) que causa(n) la No Conformidad)"
         Height          =   975
         Left            =   -74760
         TabIndex        =   46
         Top             =   1560
         Width           =   7455
         Begin VB.TextBox r_causas 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   240
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   47
            Top             =   240
            Width           =   6975
         End
      End
      Begin VB.Frame frame 
         Caption         =   "Evidencia Objetiva de las Causas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Index           =   1
         Left            =   240
         TabIndex        =   19
         Top             =   3840
         Width           =   7695
         Begin VB.CommandButton BtnEPdfVer 
            Caption         =   "Ver PDF"
            Height          =   300
            Left            =   3120
            TabIndex        =   76
            Top             =   600
            Width           =   1095
         End
         Begin VB.TextBox lblEAdjuntos 
            Height          =   735
            Left            =   4320
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   63
            Top             =   240
            Width           =   3255
         End
         Begin VB.CommandButton BtnEFotosVer 
            Caption         =   "Ver Foto(s)"
            Height          =   300
            Left            =   3120
            TabIndex        =   59
            Top             =   240
            Width           =   1095
         End
         Begin VB.CommandButton btnEAdjuntar 
            Caption         =   "Adjuntar Archivo (JPG, PDF o DWG)"
            Height          =   660
            Left            =   360
            TabIndex        =   20
            Top             =   240
            Width           =   2415
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Revision Nº y Realizada Por"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2415
         Left            =   240
         TabIndex        =   21
         Top             =   5040
         Width           =   7695
         Begin VB.ComboBox CbRevisadoPor 
            Height          =   315
            Index           =   2
            Left            =   840
            Style           =   2  'Dropdown List
            TabIndex        =   27
            Top             =   1200
            Width           =   3255
         End
         Begin VB.ComboBox CbRevisadoPor 
            Height          =   315
            Index           =   1
            Left            =   840
            Style           =   2  'Dropdown List
            TabIndex        =   25
            Top             =   840
            Width           =   3255
         End
         Begin VB.ComboBox CbRevisadoPor 
            Height          =   315
            Index           =   0
            Left            =   840
            Style           =   2  'Dropdown List
            TabIndex        =   23
            Top             =   480
            Width           =   3255
         End
         Begin VB.Frame frame 
            Caption         =   "Efect de las Acciones"
            Height          =   1455
            Index           =   4
            Left            =   4320
            TabIndex        =   28
            Top             =   240
            Width           =   1815
            Begin VB.Frame Frame4 
               BorderStyle     =   0  'None
               Height          =   400
               Left            =   120
               TabIndex        =   58
               Top             =   880
               Width           =   1575
               Begin VB.OptionButton Op_efe3 
                  Caption         =   "No"
                  Height          =   200
                  Index           =   1
                  Left            =   840
                  TabIndex        =   34
                  Top             =   120
                  Width           =   615
               End
               Begin VB.OptionButton Op_efe3 
                  Caption         =   "Si"
                  Height          =   200
                  Index           =   0
                  Left            =   120
                  TabIndex        =   33
                  Top             =   120
                  Width           =   615
               End
            End
            Begin VB.Frame Frame3 
               BorderStyle     =   0  'None
               Height          =   400
               Left            =   240
               TabIndex        =   57
               Top             =   520
               Width           =   1455
               Begin VB.OptionButton Op_efe2 
                  Caption         =   "No"
                  Height          =   200
                  Index           =   1
                  Left            =   720
                  TabIndex        =   32
                  Top             =   150
                  Width           =   615
               End
               Begin VB.OptionButton Op_efe2 
                  Caption         =   "Si"
                  Height          =   200
                  Index           =   0
                  Left            =   0
                  TabIndex        =   31
                  Top             =   120
                  Width           =   615
               End
            End
            Begin VB.Frame Frame2 
               BorderStyle     =   0  'None
               Height          =   400
               Left            =   120
               TabIndex        =   56
               Top             =   160
               Width           =   1575
               Begin VB.OptionButton Op_efe1 
                  Caption         =   "No"
                  Height          =   200
                  Index           =   1
                  Left            =   840
                  TabIndex        =   30
                  Top             =   150
                  Width           =   615
               End
               Begin VB.OptionButton Op_efe1 
                  Caption         =   "Si"
                  Height          =   200
                  Index           =   0
                  Left            =   120
                  TabIndex        =   29
                  Top             =   150
                  Width           =   495
               End
            End
         End
         Begin VB.TextBox e_comentario 
            Height          =   300
            Left            =   120
            MaxLength       =   50
            TabIndex        =   36
            Top             =   1920
            Width           =   7380
         End
         Begin VB.Label Label5 
            Caption         =   "3ra"
            Height          =   300
            Left            =   360
            TabIndex        =   26
            Top             =   1200
            Width           =   255
         End
         Begin VB.Label Label4 
            Caption         =   "2da"
            Height          =   300
            Left            =   360
            TabIndex        =   24
            Top             =   840
            Width           =   375
         End
         Begin VB.Label Label3 
            Caption         =   "1ra"
            Height          =   300
            Left            =   360
            TabIndex        =   22
            Top             =   480
            Width           =   375
         End
         Begin VB.Label lbl 
            Caption         =   "Comentarios"
            Height          =   255
            Index           =   7
            Left            =   120
            TabIndex        =   35
            Top             =   1680
            Width           =   1575
         End
      End
      Begin VB.Frame frame 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1935
         Index           =   0
         Left            =   240
         TabIndex        =   17
         Top             =   1680
         Width           =   7635
         Begin VB.OptionButton Op_NC 
            Caption         =   "NC del Cliente"
            Height          =   255
            Index           =   3
            Left            =   5160
            TabIndex        =   79
            Top             =   200
            Width           =   1455
         End
         Begin VB.OptionButton Op_NC 
            Caption         =   "NC Potencial"
            Height          =   255
            Index           =   2
            Left            =   3480
            TabIndex        =   67
            Top             =   200
            Width           =   1455
         End
         Begin VB.OptionButton Op_NC 
            Caption         =   "NC PNC"
            Height          =   255
            Index           =   1
            Left            =   1800
            TabIndex        =   66
            Top             =   200
            Width           =   1215
         End
         Begin VB.OptionButton Op_NC 
            Caption         =   "NC SGC"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   65
            Top             =   200
            Width           =   1095
         End
         Begin VB.TextBox e_descripcion 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1275
            Left            =   240
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   18
            Top             =   480
            Width           =   7215
         End
      End
      Begin MSMask.MaskEdBox e_FechaCierre 
         Height          =   300
         Left            =   2880
         TabIndex        =   39
         Top             =   7560
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   529
         _Version        =   327680
         PromptChar      =   "_"
      End
      Begin VB.Label lblGerente 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   495
         Left            =   6480
         TabIndex        =   82
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label lblFechaPrimeraRespuesta 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   -72480
         TabIndex        =   80
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha Primera Respuesta"
         Height          =   255
         Left            =   -74640
         TabIndex        =   64
         Top             =   600
         Width           =   2055
      End
      Begin VB.Label lbl 
         Caption         =   "Costo Estimado"
         Height          =   255
         Index           =   6
         Left            =   -74640
         TabIndex        =   54
         Top             =   7560
         Width           =   1215
      End
      Begin VB.Label lbl 
         Caption         =   "Clave Encargado"
         Height          =   255
         Index           =   9
         Left            =   -70440
         TabIndex        =   44
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label lbl 
         Caption         =   "Fecha de Cierre"
         Height          =   255
         Index           =   5
         Left            =   1680
         TabIndex        =   38
         Top             =   7560
         Width           =   1215
      End
      Begin VB.Label Label6 
         Caption         =   "Clave"
         Height          =   255
         Left            =   6000
         TabIndex        =   11
         Top             =   480
         Width           =   495
      End
      Begin VB.Label lbl 
         Caption         =   "Clave Receptor"
         Height          =   255
         Index           =   8
         Left            =   -70440
         TabIndex        =   42
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Area"
         Height          =   255
         Left            =   360
         TabIndex        =   15
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label lbl 
         Caption         =   "Gerencia donde se detectó la NC"
         Height          =   255
         Index           =   2
         Left            =   360
         TabIndex        =   13
         Top             =   960
         Width           =   2535
      End
      Begin VB.Label lbl 
         Caption         =   "Nombre del Emisor"
         Height          =   255
         Index           =   3
         Left            =   360
         TabIndex        =   9
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label lbl 
         Caption         =   "Persona Encargada de dar Solucion"
         Height          =   255
         Index           =   4
         Left            =   -74640
         TabIndex        =   40
         Top             =   960
         Width           =   2655
      End
   End
   Begin VB.TextBox Numero 
      Height          =   300
      Left            =   840
      TabIndex        =   2
      Top             =   600
      Width           =   975
   End
   Begin MSMask.MaskEdBox e_fecha 
      Height          =   300
      Left            =   3000
      TabIndex        =   4
      Top             =   600
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   529
      _Version        =   327680
      PromptChar      =   "_"
   End
   Begin MSComctlLib.ImageList ImageList 
      Left            =   7560
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "sql_noconformidad.frx":0038
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "sql_noconformidad.frx":014A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "sql_noconformidad.frx":025C
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "sql_noconformidad.frx":036E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "sql_noconformidad.frx":0480
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "sql_noconformidad.frx":0592
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "sql_noconformidad.frx":06A4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label lbl 
      Caption         =   "Número"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   735
   End
   Begin VB.Label lbl 
      Caption         =   "Fecha"
      Height          =   255
      Index           =   1
      Left            =   2280
      TabIndex        =   3
      Top             =   600
      Width           =   735
   End
End
Attribute VB_Name = "sql_noconformidad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Obj As String, Objs As String
Private btnAgregar As Button, btnModificar As Button, btnEliminar As Button
Private btnImprimir As Button, btnDesHacer As Button, btnGrabar As Button
Private DbD As Database, RsTra As Recordset
Private Accion As String, d As Variant, prt As Printer, i As Integer
Private a_Gerencias(3, 19) As String, a_Areas(1, 99) As String
Private a_Emisores(2, 99) As String, a_Encargados(3, 299) As String
Private m_PathDestino As String, m_NombreArchivoDestino As String
' para sql
Private RsTabla As New ADODB.Recordset, TablaNombre As String
Private RsGer As New ADODB.Recordset, RsAreas As New ADODB.Recordset
Private ac(21, 1) As String, av(21) As String, asn(21) As Boolean

' para servidor web acr3006-dualpro
' originalmente esta en i:\scp\mdb\nc_files
Private Const Carpeta_Adjuntos As String = "nc_files"
Private Test As Boolean, aTexto(19) As String
Private Const ClaveAdminNC As String = "2005" ' barbara lagos
Private Sub btnEAdjuntar_Click()

' emisor adjunta archivos

Dim p As Integer
Dim m_PathyArchivoOrigen As String
Dim m_ArchivoOrigenNombre As String

' formsto de archivo adjunto:
' nc_N_archivousuario.extensionusuario
' N es el numero de la NO Conformidad

If Op_ER(0).Value = False And Op_ER(1).Value = False Then
    MsgBox "Indique si UD. es Emisor o Receptor de la NO Conformidad"
    Exit Sub
End If

' primero verifica si digito clave y si es correcta
If e_clave.Text = "" Then
    Beep
    MsgBox "DEBE DIGITAR CLAVE"
    e_clave.SetFocus
    Exit Sub
End If

' verifica si clave es correcta
If a_Emisores(2, CbEmisores.ListIndex) <> e_clave.Text Then
    Beep
    MsgBox "CLAVE DE EMISOR INCORRECTA"
    e_clave.SetFocus
    Exit Sub
End If

' adjunta archivo

cd.DialogTitle = "Buscar Archivo"
cd.Filter = "Imagen (*.jpg)|*.jpg|Documento PDF (*.pdf)|*.pdf|Documento DWG (*.dwg)|*.dwg|Todos los Archivos (*.*)|*.*"

cd.ShowOpen

m_PathyArchivoOrigen = cd.filename

If m_PathyArchivoOrigen = "" Then
'    MsgBox "Debe escoger Archivo"
    Exit Sub
End If

'MsgBox Cd.filename ' viene con ruta clompleta
m_PathyArchivoOrigen = cd.filename

' separa path y archivo
p = InStrLast(m_PathyArchivoOrigen, ".")
'extension = Mid(m_PathyArchivoOrigen, p)

p = InStrLast(m_PathyArchivoOrigen, "\")

If p > 0 Then

'    m_PathArchivo = Left(m_PathArchivo, p)
    m_ArchivoOrigenNombre = Mid(m_PathyArchivoOrigen, p + 1)
    
    ' copia archivo a carpeta de servidor con nombre "numero"
    ' es decir, queda con el numbre del numero de la No Conformidad
'    m_PathDestino = Drive_Server & Path_Server & "nc_img\"
'    m_Destino = "C:\foto.jpg"

'    For i = 1 To 4 ' hasta cuatro fotos
    
'       m_NombreArchivoDestino = "nc_" & Numero.Text & "_" & i & "." & extension

        m_NombreArchivoDestino = "nc_" & Numero.Text & "_" & m_ArchivoOrigenNombre
        
        m_NombreArchivoDestino = Replace(m_NombreArchivoDestino, " ")
        
        If Archivo_Existe(m_PathDestino, m_NombreArchivoDestino) Then
            
        Else
        
            FileCopy m_PathyArchivoOrigen, m_PathDestino & m_NombreArchivoDestino
'            Exit For
        
        End If

'    Next
    
If False Then

    If i = 5 Then
    
        ' quiere decir que ya subio cuatro archivos
        ' inicio de renombre de fotos desde:
        
        ' eliminar foto_1
        Kill m_PathDestino & Numero.Text & "_1.jpg"
        
        ' REN foto_2 a foto_1
        Name m_PathDestino & Numero.Text & "_2.jpg" As m_PathDestino & Numero.Text & "_1.jpg"
        
        ' REN foto_3 a foto_2
        Name m_PathDestino & Numero.Text & "_3.jpg" As m_PathDestino & Numero.Text & "_2.jpg"
        
        ' REN foto_4 a foto_3
        Name m_PathDestino & Numero.Text & "_4.jpg" As m_PathDestino & Numero.Text & "_3.jpg"
        
        ' SUBE FOTO 4
        i = 4
        m_NombreArchivoDestino = Numero.Text & "_" & i & ".jpg"
        FileCopy m_PathyArchivoOrigen, m_PathDestino & m_NombreArchivoDestino
       
    End If
    
End If
    
'    lblCarpeta.Caption = m_Path
    
    ' guarda ultima ruta usada
'    SaveSetting "scp", "planos", "ruta", m_Path
    
End If

Archivos_Subidos_Buscar "E"

End Sub
Private Sub BtnEFotosVer_Click()
frmJPGMostrar.Tipo = "E" ' emisor
frmJPGMostrar.Numero = Numero.Text
frmJPGMostrar.Show 1
End Sub
Private Sub BtnEPdfVer_Click()
' muestra form para ver pdfs
frmPDFMostrar.Tipo = "E" ' emisor
frmPDFMostrar.Numero = Numero.Text
frmPDFMostrar.Show 1
End Sub
Private Sub btnRAdjuntar_Click()
' receptor adjunta archivo
Dim p As Integer
Dim m_PathyArchivoOrigen As String
Dim m_ArchivoOrigenNombre As String

' formato de archivo adjunto:
' ncr_N_archivousuario.extensionusuario
' N es el numero de la NO Conformidad

If Not Receptor_Validar() Then
    Exit Sub
End If

' adjunta archivo

cd.DialogTitle = "Buscar Imagen JPG"
cd.Filter = "Imagen (*.jpg)|*.jpg|Documento PDF (*.pdf)|*.pdf|Documento DWG (*.dwg)|*.dwg|Todos los Archivos (*.*)|*.*"

cd.ShowOpen

m_PathyArchivoOrigen = cd.filename

If m_PathyArchivoOrigen = "" Then
'    MsgBox "Debe escoger Archivo"
    Exit Sub
End If

'MsgBox Cd.filename ' viene con ruta clompleta
m_PathyArchivoOrigen = cd.filename

' separa path y archivo
p = InStrLast(m_PathyArchivoOrigen, ".")
'extension = Mid(m_PathyArchivoOrigen, p)

p = InStrLast(m_PathyArchivoOrigen, "\")

If p > 0 Then

    m_ArchivoOrigenNombre = Mid(m_PathyArchivoOrigen, p + 1)
    
    m_NombreArchivoDestino = "ncr_" & Numero.Text & "_" & m_ArchivoOrigenNombre
    If Archivo_Existe(m_PathDestino, m_NombreArchivoDestino) Then

    Else
        FileCopy m_PathyArchivoOrigen, m_PathDestino & m_NombreArchivoDestino
    End If

End If

Archivos_Subidos_Buscar "R"

End Sub
Private Sub btnREliminarAdjuntos_Click()

If Not Receptor_Validar() Then
    Exit Sub
End If

If MsgBox("¿ Elimina TODOS los archivos adjuntos del Receptor ?", vbOKCancel) = vbOK Then
    'If Tipo = "R" Then
    Dim ArchivoNombre
    ArchivoNombre = Dir(m_PathDestino & "ncr_" & Numero.Text & "_*.*", vbArchive)
    If ArchivoNombre <> "" Then
        Do
            'MsgBox "voy a eliminar " & ArchivoNombre
            Kill m_PathDestino & ArchivoNombre
            ArchivoNombre = Dir()
        Loop Until ArchivoNombre = ""
    End If
    lblRAdjuntos.Text = ""
End If
End Sub
Private Sub BtnRFotosVer_Click()
frmJPGMostrar.Tipo = "R" ' receptor
frmJPGMostrar.Numero = Numero.Text
frmJPGMostrar.Show 1
End Sub
Private Sub btnRPdfVer_Click()
' muestra form para ver pdfs
frmPDFMostrar.Tipo = "R" ' receptor
frmPDFMostrar.Numero = Numero.Text
frmPDFMostrar.Show 1
End Sub
'Private m_GerenteClave As String ' clave del gerente que recibe la NC, ej: erwin
Private Sub Form_Load()

Dim j As Integer, m_Rut As String

If False Then
    RsGer.Open "SELECT * FROM track WHERE documento_numero=2681", CnxSqlServer_scp0
    With RsGer
    Do While Not .EOF
        Debug.Print !fechahora, !Usuario_Win, !usuario_scp, !documento_tipo, !documento_numero, !Operacion
        .MoveNext
    Loop
    End With
    RsGer.Close
End If

'Test = True
Test = False

TablaNombre = "noconformidad"

i = 0
j = 0 ' para receptores
Set DbD = OpenDatabase(data_file)
Set RsTra = DbD.OpenRecordset("Trabajadores")

With RsTra
.Index = "apellidos"

CbEmisores.AddItem " "
CbEncargado.AddItem " "
CbRevisadoPor(0).AddItem " "
CbRevisadoPor(1).AddItem " "
CbRevisadoPor(2).AddItem " "

Do While Not .EOF

    If !emisor_nc Then
'    If True Then
    
        i = i + 1
        a_Emisores(0, i) = !rut
        a_Emisores(1, i) = !appaterno & " " & !apmaterno & " " & !nombres
        a_Emisores(2, i) = NoNulo(!emisor_nc_clave)
        
        CbEmisores.AddItem a_Emisores(1, i)
        
        CbRevisadoPor(0).AddItem a_Emisores(1, i)
        CbRevisadoPor(1).AddItem a_Emisores(1, i)
        CbRevisadoPor(2).AddItem a_Emisores(1, i)
        
    End If
    
'    If True Then ' no es toda la fabrica
    If False Then
    If !emisor_nc Then
    
        j = j + 1
        a_Encargados(0, j) = !rut
        a_Encargados(1, j) = !appaterno & " " & !apmaterno & " " & !nombres
        a_Encargados(2, j) = NoNulo(!emisor_nc_clave)
        a_Encargados(3, j) = NoNulo(!dato20)
        
        CbEncargado.AddItem a_Encargados(1, j)
        
    End If
    End If
    
    .MoveNext
    
Loop
'.Close
End With

' para password
e_clave.PasswordChar = "*"
e_clave.MaxLength = 10
e_comentario.MaxLength = 1000
r_clave.PasswordChar = "*"
r_clave.MaxLength = 10
encargado_clave.PasswordChar = "*"
encargado_clave.MaxLength = 10

'Set Dbm2 = OpenDatabase(mpro2_file)
'Set RsINC = Dbm2.OpenRecordset("INC")
'RsINC.Index = "numero"

RsTra.Index = "rut"

Inicializa

Me.Caption = Obj

If Usuario.ReadOnly Then
    Botones_Enabled 0, 0, 0, 1, 0, 0
Else
    Botones_Enabled 1, 1, 1, 1, 0, 0
End If

i = 0
a_Gerencias(0, 0) = ""
a_Gerencias(1, 0) = " "
'a_Gerencias(2, 0) ' clave
'a_Gerencias(3, 0) ' nombre de gerente
CbGerencias.AddItem a_Gerencias(1, i)
RsGer.Open "SELECT * FROM maestros WHERE tipo='GER' AND activo='S'", CnxSqlServer_scp0
With RsGer
Do While Not .EOF

    i = i + 1
    a_Gerencias(0, i) = !Codigo
    a_Gerencias(1, i) = !Descripcion
    m_Rut = !dato1
    m_Rut = PadL(m_Rut, 10)
    RsTra.Seek "=", m_Rut
    If RsTra.NoMatch Then
        a_Gerencias(2, i) = ""
        a_Gerencias(3, i) = ""
    Else
        a_Gerencias(2, i) = RsTra!emisor_nc_clave ' clave de NC del gerente
        a_Gerencias(3, i) = Trim(RsTra![nombres]) & " " & Trim(RsTra![appaterno])
    End If
        
    CbGerencias.AddItem a_Gerencias(1, i)
    
    .MoveNext
Loop
.Close
End With

' el CbAreas se puebla de acuerdo a lo elegido en area

e_descripcion.MaxLength = 1000
r_CostoEstimado.MaxLength = 8
r_causas.MaxLength = 1000
r_acciones.MaxLength = 1000

m_PathDestino = Drive_Server & Path_Mdb & Carpeta_Adjuntos & "\"

End Sub
Private Sub Campos_Emisor_Definir()

ac(1, 0) = "e_numero"
ac(1, 1) = ""
ac(2, 0) = "e_fecha"
ac(2, 1) = "'"
ac(3, 0) = "e_rut"
ac(3, 1) = "'"
ac(4, 0) = "e_gerencia"
ac(4, 1) = "'"
ac(5, 0) = "e_area"
ac(5, 1) = "'"
ac(6, 0) = "e_tipo" ' tipo de no conformidad: N: normal, P: potencial
ac(6, 1) = "'"
ac(7, 0) = "e_descripcion"  ' [varchar]   (1000),
ac(7, 1) = "'"
ac(8, 0) = "e_rutrev1"      ' [varchar]     (10),"
ac(8, 1) = "'"
ac(9, 0) = "e_efectividad1" ' [varchar]      (1),"
ac(9, 1) = "'"
ac(10, 0) = "e_rutrev2"     ' [varchar]     (10),"
ac(10, 1) = "'"
ac(11, 0) = "e_efectividad2" ' [varchar]     (1),"
ac(11, 1) = "'"
ac(12, 0) = "e_rutrev3"      ' [varchar]    (10),"
ac(12, 1) = "'"
ac(13, 0) = "e_efectividad3" ' [varchar]     (1),"
ac(13, 1) = "'"
ac(14, 0) = "e_comentarios"  ' [varchar]  (1000),"
ac(14, 1) = "'"
ac(15, 0) = "e_cerrado"      ' [varchar]     (1),"
ac(15, 1) = "'"
ac(16, 0) = "e_fechacierre"       ' [datetime],"
ac(16, 1) = "'"

End Sub
Private Sub Campos_Receptor_Definir()

ac(1, 0) = "e_numero"
ac(1, 1) = ""
ac(2, 0) = "r_rut"          ' [varchar]   (10),"
ac(2, 1) = "'"
ac(3, 0) = "r_causas"       ' [varchar] (1000),"
ac(3, 1) = "'"
ac(4, 0) = "r_fecha1"       ' [datetime],"
ac(4, 1) = "'"
ac(5, 0) = "r_fecha2"       ' [datetime],"
ac(5, 1) = "'"
ac(6, 0) = "r_fecha3"       ' [DateTime]"
ac(6, 1) = "'"
ac(7, 0) = "r_acciones"     ' [varchar] (1000),"
ac(7, 1) = "'"
ac(8, 0) = "costoestimado"  ' [int]'
ac(8, 1) = ""
ac(9, 0) = "r_acReproceso"  ' [varchar] (1),"
ac(9, 1) = "'"
ac(10, 0) = "r_acReparacion"  ' [varchar] (1),"
ac(10, 1) = "'"
ac(11, 0) = "r_acRFI"  ' [varchar] (1),"
ac(11, 1) = "'"
ac(12, 0) = "r_acOtra"  ' [varchar] (1),"
ac(12, 1) = "'"
ac(13, 0) = "r_accionCorrectiva"  ' [varchar] (1000),"
ac(13, 1) = "'"
ac(14, 0) = "r_comentarios"  ' [varchar] (1000),"
ac(14, 1) = "'"
ac(15, 0) = "r_fechaPrimeraRespuesta" ' [DateTime]"
ac(15, 1) = "'"

End Sub
Private Sub Archivos_Subidos_Buscar(Tipo As String)
' busca nombres de archivos subidos
Dim ArchivoNombre As String


If Tipo = "E" Then

'    lblEAdjuntos.Caption = ""
    lblEAdjuntos.Text = ""
    ArchivoNombre = Dir(m_PathDestino & "nc_" & Numero.Text & "_*.*", vbArchive)
    If ArchivoNombre <> "" Then
        Do
'            lblEAdjuntos.Caption = lblEAdjuntos.Caption & " " & ArchivoNombre
            lblEAdjuntos.Text = lblEAdjuntos.Text & " " & ArchivoNombre
            ArchivoNombre = Dir()
        Loop Until ArchivoNombre = ""
    End If
End If
If Tipo = "R" Then
    lblRAdjuntos.Text = ""
    ArchivoNombre = Dir(m_PathDestino & "ncr_" & Numero.Text & "_*.*", vbArchive)
    If ArchivoNombre <> "" Then
        Do
            lblRAdjuntos.Text = lblRAdjuntos.Text & " " & ArchivoNombre
            ArchivoNombre = Dir()
        Loop Until ArchivoNombre = ""
    End If
End If

'D:\acr3006-dualpro\scp\mdb\nc_files

End Sub

Private Sub CbGerencias_Click()

' puebla combo de areas
Dim m_Gerencia As String
i = CbGerencias.ListIndex

If i = -1 Then Exit Sub

'MsgBox i

lblGerente.Caption = a_Gerencias(3, i)

CbAreas.Clear
CbAreas.AddItem " "
m_Gerencia = a_Gerencias(0, i)
i = 0

RsAreas.Open "SELECT * FROM maestros WHERE tipo='GAR' AND dato1='" & m_Gerencia & "' AND activo='S'", CnxSqlServer_scp0
With RsAreas
Do While Not .EOF
    i = i + 1
    a_Areas(0, i) = !Codigo
    a_Areas(1, i) = !Descripcion
    CbAreas.AddItem a_Areas(1, i)
    .MoveNext
Loop
.Close
End With

End Sub
Private Sub e_fecha_LostFocus()
d = Fecha_Valida(e_fecha)
End Sub

Private Sub Op_NC_Click(Index As Integer)

If Op_NC(0).Value Then
    frame_r_causas.Caption = "Investigación de las Causas (Explique el (los) motivo(s) que causa(n) la No Conformidad) NC SGC"
End If

If Op_NC(1).Value Then
    frame_r_causas.Caption = "Investigación de las Causas (NC PNC)"
End If

If Op_NC(2).Value Then
    frame_r_causas.Caption = "Investigación de las Causas (NC Potencial)"
End If

If Op_NC(3).Value Then
    frame_r_causas.Caption = "Investigación de las Causas (NC del Cliente)"
End If

End Sub
Private Sub r_fecha1_LostFocus()
d = Fecha_Valida(r_fecha1)
End Sub
Private Sub r_fecha2_LostFocus()
d = Fecha_Valida(r_fecha2)
End Sub
Private Sub r_fecha3_LostFocus()
d = Fecha_Valida(r_fecha3)
End Sub
Private Sub Form_LostFocus()
'valida fecha
End Sub

Private Sub Form_Unload(Cancel As Integer)
DataBases_Cerrar
End Sub
Private Sub Inicializa()

Obj = "Registro"
Objs = "Registros"

Set btnAgregar = Toolbar.Buttons(1)
Set btnModificar = Toolbar.Buttons(2)
Set btnEliminar = Toolbar.Buttons(3)
Set btnImprimir = Toolbar.Buttons(4)
Set btnDesHacer = Toolbar.Buttons(6)
Set btnGrabar = Toolbar.Buttons(7)

Accion = ""

'btnSearch.Visible = False
'btnSearch.ToolTipText = "Busca Cliente"

'Numero.Mask = "##########"
'Numero.PromptInclude = False
Numero.MaxLength = 10

e_fecha.Mask = Fecha_Mask

e_FechaCierre.Mask = Fecha_Mask

r_fecha1.Mask = Fecha_Mask
r_fecha2.Mask = Fecha_Mask
r_fecha3.Mask = Fecha_Mask

Campos_Enabled False

End Sub
Private Sub Botones_Enabled(btn_Agregar As Boolean, btn_Modificar As Boolean, _
                            btn_Eliminar As Boolean, btn_Imprimir As Boolean, _
                            btn_DesHacer As Boolean, btn_Grabar As Boolean)
                            
btnAgregar.Enabled = btn_Agregar
btnModificar.Enabled = btn_Modificar
btnEliminar.Enabled = btn_Eliminar
btnImprimir.Enabled = btn_Imprimir
btnDesHacer.Enabled = btn_DesHacer
btnGrabar.Enabled = btn_Grabar

btnAgregar.Value = tbrUnpressed
btnModificar.Value = tbrUnpressed
btnEliminar.Value = tbrUnpressed
btnImprimir.Value = tbrUnpressed
btnDesHacer.Value = tbrUnpressed
btnGrabar.Value = tbrUnpressed

End Sub
Private Sub Campos_Enabled(Si As Boolean)

Numero.Enabled = Si
Frame_ER.Enabled = Si

Emisor_Campos_Enabled Si

Receptor_Campos_Enabled Si

End Sub
Private Sub Emisor_Campos_Enabled(Si As Boolean)

e_fecha.Enabled = Si

' emisor
CbEmisores.Enabled = Si
e_clave.Enabled = Si
CbGerencias.Enabled = Si
CbAreas.Enabled = Si

Op_NC(0).Enabled = Si
Op_NC(1).Enabled = Si
Op_NC(2).Enabled = Si
Op_NC(3).Enabled = Si

e_descripcion.Enabled = Si

'SSTab.Enabled = Si
'If Si Then
'    SSTab.Tab = 1
'    SSTab.Tab = 0
'End If

btnEAdjuntar.Enabled = Si
BtnEFotosVer.Enabled = Si
BtnEPdfVer.Enabled = Si

CbRevisadoPor(0).Enabled = Si
CbRevisadoPor(1).Enabled = Si
CbRevisadoPor(2).Enabled = Si

Op_efe1(0).Enabled = Si
Op_efe1(1).Enabled = Si

Op_efe2(0).Enabled = Si
Op_efe2(1).Enabled = Si

Op_efe3(0).Enabled = Si
Op_efe3(1).Enabled = Si

e_comentario.Enabled = Si
e_Cerrado.Enabled = Si
e_FechaCierre.Enabled = Si

lblEAdjuntos.Enabled = False

End Sub
Private Sub Campos_Limpiar()

Numero.Text = ""
e_fecha.Text = Fecha_Vacia

'm_GerenteClave = ""

Op_ER(0).Value = False
Op_ER(1).Value = False

CbEmisores.ListIndex = -1
e_clave.Text = ""

lblGerente.Caption = ""

CbGerencias.ListIndex = -1
CbAreas.ListIndex = -1

Op_NC(0).Value = False
Op_NC(1).Value = False
Op_NC(2).Value = False
Op_NC(3).Value = False

e_descripcion.Text = ""
r_CostoEstimado.Text = ""

SSTab.Tab = 0

CbRevisadoPor(0).ListIndex = -1
CbRevisadoPor(1).ListIndex = -1
CbRevisadoPor(2).ListIndex = -1

Op_efe1(0).Value = False
Op_efe1(1).Value = False

Op_efe2(0).Value = False
Op_efe2(1).Value = False

Op_efe3(0).Value = False
Op_efe3(1).Value = False

e_comentario.Text = ""
e_Cerrado.Value = 0

CbEncargado.ListIndex = -1

r_clave.Text = ""

lblFechaPrimeraRespuesta.Caption = "" ' "10/06/15"

r_fecha1.Text = Fecha_Vacia
r_fecha2.Text = Fecha_Vacia
r_fecha3.Text = Fecha_Vacia

r_causas.Text = ""
r_acciones.Text = ""

'lblEAdjuntos.Caption = ""
lblEAdjuntos.Text = ""
lblRAdjuntos.Text = ""

op_ac(0).Value = False
op_ac(1).Value = False
op_ac(2).Value = False
op_ac(3).Value = False

r_ac.Text = ""
r_comentarios.Text = ""

'Activa.Value = True

End Sub
Private Sub Numero_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then After_Enter
End Sub
Private Sub After_Enter()
If Val(Numero.Text) < 1 Then Beep: Exit Sub
Select Case Accion
Case "Agregando"

'    RsINC.Seek "=", Numero.Text
'    If Not RsINC.NoMatch Then
    If Registro_Existe(TablaNombre, "e_numero=" & Numero.Text) Then
    
        Doc_Leer
        
        MsgBox Obj & " YA EXISTE"
        Campos_Limpiar
        Numero.Enabled = True
        Numero.SetFocus
        
    Else
    
        Campos_Enabled True
        
        Numero.Enabled = False
        
        Op_ER(0).Value = True ' no conformidad nueva la genera un Emisor
        
        e_fecha.Text = Format(Now, Fecha_Format)
        e_fecha.SetFocus
        
        Botones_Enabled 0, 0, 0, 0, 1, 1
        
    End If
    
Case "Modificando"

'    RsINC.Seek "=", Numero.Text
'    If RsINC.NoMatch Then
    If Not Registro_Existe(TablaNombre, "e_numero=" & Numero.Text) Then
        MsgBox Obj & " NO EXISTE"
    Else
    
        Botones_Enabled 0, 0, 0, 0, 1, 1
        
        Doc_Leer
        
        Campos_Enabled True
        Numero.Enabled = False
        
'        CbEmisores.Enabled = False
'        clave.text
        
'        Emisor_Enabled
        
        ' no se puede modificar emisor, ni siquiera el emisor
        CbEmisores.Enabled = False
        
'        btnSearch.Visible = True
        
    End If

Case "Eliminando"

'    RsINC.Seek "=", Numero.Text
'    If RsINC.NoMatch Then
    If Not Registro_Existe(TablaNombre, "e_numero=" & Numero.Text) Then

        MsgBox Obj & " NO EXISTE"
    Else
    
        Doc_Leer
        
        Campos_Enabled False
        e_clave.Enabled = True
'        e_clave.SetFocus
        
        Botones_Enabled 0, 0, 1, 0, 1, 0
        
    End If
   
Case "Imprimiendo"

'    RsINC.Seek "=", Numero
'    If RsINC.NoMatch Then
    If Not Registro_Existe(TablaNombre, "e_numero=" & Numero.Text) Then
        MsgBox Obj & " NO EXISTE"
    Else
        Doc_Leer
        Numero.Enabled = False
        'If MsgBox("¿ Imprime ?", vbYesNo) = vbYes Then
        
        Dim m_ImpresoraNombre As String

        prt_escoger.ImpresoraNombre = ""
        prt_escoger.Show 1
        m_ImpresoraNombre = prt_escoger.ImpresoraNombre
        
        If m_ImpresoraNombre <> "" Then
        
            NC_Prepara Numero.Text, m_ImpresoraNombre
            NC_Print
            
        End If
            
        'End If
        'Campos_Limpiar
        'Numero.Enabled = True
        'Numero.SetFocus
    End If
End Select

End Sub
Private Sub Toolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
Case 1 ' agregar
    Accion = "Agregando"
    Botones_Enabled 0, 0, 0, 0, 1, 0
    Campos_Enabled False
    
    Numero.Text = sql_Documento_Numero_Nuevo(CnxSqlServer_scp0, TablaNombre, "", "e_numero")
    Numero.Enabled = True
    Numero.SetFocus
    
Case 2 ' modificar
    Accion = "Modificando"
    Botones_Enabled 0, 0, 0, 0, 1, 0
    Campos_Limpiar
    Campos_Enabled False
    Numero.Enabled = True
    Numero.SetFocus ''
'    Obras_Leer
'    ComboNV.Visible = True
Case 3 ' Eliminar

    Accion = "Eliminando"
    If Numero.Text = "" Then
        Botones_Enabled 0, 0, 0, 0, 1, 0
        Campos_Enabled False
        Numero.Enabled = True
        Numero.SetFocus
    Else
    
        If e_clave.Text = "" Then
            Beep
            MsgBox "DEBE DIGITAR CLAVE"
'            e_clave.SetFocus
            Exit Sub
        End If
        
        ' verifica si clave es correcta
        If a_Emisores(2, CbEmisores.ListIndex) <> e_clave.Text Then
            Beep
            MsgBox "CLAVE DE EMISOR NO CORRECTA"
            e_clave.SetFocus
            Exit Sub
        End If

        If MsgBox("¿ ELIMINA " & Obj & " ?", vbYesNo) = vbYes Then
            Doc_Eliminar
        End If
        
        Campos_Limpiar
        Campos_Enabled False
        Numero.Enabled = True
        Numero.SetFocus

    End If
    
Case 4 ' Imprimir

    Accion = "Imprimiendo"
    If Numero.Text = "" Then
        Botones_Enabled 0, 0, 0, 0, 1, 0
        Campos_Enabled False
        Numero.Enabled = True
        Numero.SetFocus
    Else
        ' imprime
    End If
'    Obras_Leer
'    ComboNV.Visible = True
Case 5 ' Separador
Case 6 ' DesHacer
    If Numero = "" Then
        If Usuario.ReadOnly Then '01/06/98
            Botones_Enabled 0, 0, 0, 1, 0, 0
        Else
            Botones_Enabled 1, 1, 1, 1, 0, 0
        End If
        Campos_Limpiar
        Campos_Enabled False
        Accion = ""
'        btnSearch.Visible = False
    Else
        If MsgBox("¿Seguro que quiere DesHacer?", vbYesNo) = vbYes Then
            Botones_Enabled 1, 1, 1, 1, 0, 0
            Campos_Limpiar
            Campos_Enabled False
            Accion = ""
'            btnSearch.Visible = False
        End If
    End If
'    ComboNV.Visible = False
Case 7 ' grabar

    If Doc_Validar Then
        If MsgBox("¿ GRABA " & Obj & " ?", vbYesNo) = vbYes Then
        
            If Accion = "Agregando" Then
                Doc_Grabar True
            Else
                Doc_Grabar False
            End If
'            If Accion = "Agregando" Then Obras_Leer
            
'            If MsgBox("¿ IMPRIME " & Obj & " ?", vbYesNo) = vbYes Then
'                Doc_Imprimir
'            End If
            
            If Usuario.ReadOnly Then
                Botones_Enabled 0, 0, 0, 1, 0, 0
            Else
                Botones_Enabled 1, 1, 1, 1, 0, 0
            End If
            Campos_Limpiar
            Campos_Enabled False
            Accion = ""
'            btnSearch.Visible = False

'            btnDesHacer.Value = tbrPressed
            
'            Numero.Enabled = True
'            Numero.SetFocus
            
        End If
    End If
Case 8 'separador
Case 9
    MousePointer = vbHourglass
    Load Clientes
    MousePointer = vbDefault
    Clientes.Show 1
End Select

If Accion = "" Then
    Me.Caption = StrConv(Objs, vbProperCase)
Else
    Me.Caption = StrConv(Objs, vbProperCase) & " [" & Accion & "]"
End If

End Sub
Private Sub Doc_Leer()
'On Error Resume Next
Dim j As Integer

RsTabla.Open "SELECT * FROM " & TablaNombre & " WHERE e_numero=" & Numero.Text, CnxSqlServer_scp0

With RsTabla

If Not .EOF Then

    e_fecha.Text = Format(!e_fecha, Fecha_Format)
'    activo.Value = IIf(RsTabla!activo = "S", 1, 0)

    ' busca persona
    For i = 1 To 99
        If a_Emisores(0, i) = !e_rut Then
            CbEmisores.Text = a_Emisores(1, i)
            Exit For
        End If
    Next
    
    ' busca gerencias
    For i = 1 To 19
        If a_Gerencias(0, i) = !e_gerencia Then
            CbGerencias.Text = a_Gerencias(1, i)
            Exit For
        End If
    Next
    
    ' busca areas
    For i = 1 To 99
        If a_Areas(0, i) = !e_area Then
            CbAreas.Text = a_Areas(1, i)
            Exit For
        End If
    Next
    
    If !e_tipo = "N" Then ' sgc
        Op_NC(0).Value = True
    End If
    If !e_tipo = "R" Then ' pnc
        Op_NC(1).Value = True
    End If
    If !e_tipo = "P" Then ' potencial
        Op_NC(2).Value = True
    End If
    If !e_tipo = "C" Then ' del cliente
        Op_NC(3).Value = True
    End If
    
    e_descripcion.Text = NoNulo(!e_descripcion)
    
    r_CostoEstimado.Text = !costoestimado
    
    If !e_rutrev1 <> "" Then
        For i = 1 To 99
            If a_Emisores(0, i) = !e_rutrev1 Then
                CbRevisadoPor(0).Text = a_Emisores(1, i)
                Exit For
            End If
        Next
        Op_efe1(0).Value = IIf(NoNulo(!e_efectividad1) = "S", True, False)
        Op_efe1(1).Value = Not Op_efe1(0).Value
    End If
    If !e_rutrev2 <> "" Then
        For i = 1 To 99
            If a_Emisores(0, i) = !e_rutrev2 Then
                CbRevisadoPor(1).Text = a_Emisores(1, i)
                Exit For
            End If
        Next
        Op_efe2(0).Value = IIf(NoNulo(!e_efectividad2) = "S", True, False)
        Op_efe2(1).Value = Not Op_efe2(0).Value
    End If
    If !e_rutrev3 <> "" Then
        For i = 1 To 99
            If a_Emisores(0, i) = !e_rutrev3 Then
                CbRevisadoPor(2).Text = a_Emisores(1, i)
                Exit For
            End If
        Next
        Op_efe3(0).Value = IIf(NoNulo(!e_efectividad3) = "S", True, False)
        Op_efe3(1).Value = Not Op_efe3(0).Value
    End If
    
    e_comentario.Text = NoNulo(![e_comentarios])
    
    e_Cerrado.Value = IIf(NoNulo(!e_Cerrado) = "S", 1, 0)
    
    e_FechaCierre.Text = Format(!e_FechaCierre, Fecha_Format)
    
    ' RECEPTOR
    '////////////////////////////////
    ' los encargados se cargan solo en modificacion de nc
    If Accion = "Modificando" Or Accion = "Imprimiendo" Then
        RsTra.MoveFirst
        RsTra.Index = "apellidos"
        j = 0
        CbEncargado.Clear
        CbEncargado.AddItem " "
        Do While Not RsTra.EOF
            If RsTra!emisor_nc Then
    '            If !a = a_Gerencias(0, CbGerencias.ListIndex) Then
                    j = j + 1
                    a_Encargados(0, j) = RsTra!rut
                    a_Encargados(1, j) = RsTra!appaterno & " " & RsTra!apmaterno & " " & RsTra!nombres
                    a_Encargados(2, j) = NoNulo(RsTra!emisor_nc_clave)
                    a_Encargados(3, j) = NoNulo(RsTra!dato20)
                    CbEncargado.AddItem a_Encargados(1, j)
    '            End If
            End If
            RsTra.MoveNext
        Loop
        RsTra.Index = "rut"
    End If
    ' busca persona
    If Trim(!r_rut) <> "" Then
        For i = 1 To 299
            If a_Encargados(0, i) = !r_rut Then
                CbEncargado.Text = a_Encargados(1, i)
                Exit For
            End If
        Next
    End If
    '////////////////////////////////
    
    r_causas.Text = NoNulo(!r_causas)
    r_acciones.Text = NoNulo(!r_acciones)
    
    If Not IsNull(!r_fechaPrimeraRespuesta) Then
        If Format(!r_fechaPrimeraRespuesta, "yyyymmdd") <> "19000101" Then ' fecha vacia
            lblFechaPrimeraRespuesta.Caption = Format(!r_fechaPrimeraRespuesta, Fecha_Format)
        End If
    End If
   
    
    If Not IsNull(!r_fecha1) Then
        If Format(!r_fecha1, "yyyymmdd") <> "19000101" Then ' fecha vacia
            r_fecha1.Text = Format(!r_fecha1, Fecha_Format)
        End If
    End If
    If Not IsNull(!r_fecha2) Then
        If Format(!r_fecha2, "yyyymmdd") <> "19000101" Then ' fecha vacia
            r_fecha2.Text = Format(!r_fecha2, Fecha_Format)
        End If
    End If
    If Not IsNull(!r_fecha3) Then
        If Format(!r_fecha3, "yyyymmdd") <> "19000101" Then ' fecha vacia
            r_fecha3.Text = Format(!r_fecha3, Fecha_Format)
        End If
    End If

    op_ac(0).Value = IIf(!r_acReproceso = "S", 1, 0)
    op_ac(1).Value = IIf(!r_acReparacion = "S", 1, 0)
    op_ac(2).Value = IIf(!r_acRFI = "S", 1, 0)
    op_ac(3).Value = IIf(!r_acOtra = "S", 1, 0)
    
    r_ac.Text = NoNulo(!r_accionCorrectiva)
    r_comentarios.Text = NoNulo(!r_comentarios)

End If

.Close

End With

Archivos_Subidos_Buscar "E"
Archivos_Subidos_Buscar "R"

'txt2arreglo e_descripcion.Text, 40

End Sub
Private Function Doc_Validar() As Boolean

Doc_Validar = False

If Op_ER(0).Value = False And Op_ER(1).Value = False Then
    MsgBox "Indique si UD. es Emisor o Receptor de la NO Conformidad"
    Exit Function
End If

If Op_ER(0).Value Then
    If Not Emisor_Validar() Then
        Exit Function
    End If
End If

If Op_ER(1).Value Then
    If Not Receptor_Validar() Then
        Exit Function
    End If
End If

If False Then
    ' verifica si existe evidencia
    ' "sin evidencia no hay NO Conformidad"
'    If lblEAdjuntos.Caption = "" Then
    If lblEAdjuntos.Text = "" Then
        MsgBox "Debe Adjuntar Evidencia, es decir, adjuntar archivo(s)"
        btnEAdjuntar.SetFocus
        Exit Function
    End If
End If

Doc_Validar = True

End Function
Private Function Emisor_Validar() As Boolean

Emisor_Validar = False

If CbEmisores.ListIndex < 1 Then
    Beep
    MsgBox "DEBE ELEGIR EMISOR"
    On Error Resume Next
    CbEmisores.SetFocus
    On Error GoTo 0
    Exit Function
End If

' primero verifica si digito clave y si es correcta
If e_clave.Text = "" Then
    Beep
    MsgBox "DEBE DIGITAR CLAVE"
    e_clave.SetFocus
    Exit Function
End If

If ClaveAdminNC = e_clave.Text Then
    ' el administrador de las NC puede cerrasr cualquier NC
Else

    ' verifica si clave es correcta
    If a_Emisores(2, CbEmisores.ListIndex) <> e_clave.Text Then
        Beep
        MsgBox "CLAVE DE EMISOR NO CORRECTA"
        e_clave.SetFocus
        Exit Function
    End If
    
End If

If CbGerencias.ListIndex < 1 Then
    Beep
    MsgBox "DEBE ELEGIR GERENCIA"
    CbGerencias.SetFocus
    Exit Function
End If
If CbAreas.ListIndex < 1 Then
    Beep
    MsgBox "DEBE ELEGIR AREA"
    CbAreas.SetFocus
    Exit Function
End If

If Op_NC(0).Value = False And Op_NC(1).Value = False And Op_NC(2).Value = False And Op_NC(3).Value = False Then
    Beep
    MsgBox "DEBE indicar si es NC SGC o NC PNC o NC Potencial o NC de Cliente"
'    Op_NC(0).SetFocus
    Exit Function
End If

If Len(e_descripcion.Text) = 0 Then
    MsgBox "DEBE DIGITAR DESCRIPCIÓN"
    e_descripcion.SetFocus
    Exit Function
End If

If e_Cerrado.Value = 1 Then
    ' ticket en CERRADO
    If e_FechaCierre = Fecha_Vacia Then
        MsgBox "DEBE DIGITAR FECHA DE CIERRE"
        e_FechaCierre.SetFocus
        Exit Function
    End If
End If

Emisor_Validar = True

End Function
Private Function Receptor_Validar() As Boolean

Receptor_Validar = False

SSTab.Tab = 1

If CbEncargado.ListIndex < 1 Then
    Beep
    MsgBox "DEBE ELEGIR ENCARGADO"
    CbEncargado.SetFocus
    Exit Function
End If

' debe digitar una de las dos claves ya sea el del receptor (gerente) o el del encargado
If r_clave.Text = "" And encargado_clave.Text = "" Then
    Beep
    MsgBox "Debe digitar clave de Receptor o Encargado"
    r_clave.SetFocus
    Exit Function
End If

' verifica clave receptor ( gerente )
If r_clave.Text <> "" Then
    If a_Gerencias(2, CbGerencias.ListIndex) <> r_clave.Text Then
        Beep
        MsgBox "Clave Receptor incorrecta", , "Error"
        r_clave.SetFocus
        Exit Function
    End If
End If

' verifica si clave encargado es correcta
If encargado_clave.Text <> "" Then
    If a_Encargados(2, CbEncargado.ListIndex) <> encargado_clave.Text Then
        Beep
        MsgBox "Clave de Encargado incorrecta"
        encargado_clave.SetFocus
        Exit Function
    End If
End If

Receptor_Validar = True

End Function
Private Sub Receptor_Campos_Enabled(Si As Boolean)

CbEncargado.Enabled = Si

r_clave.Enabled = Si
encargado_clave.Enabled = Si

r_fecha1.Enabled = Si
r_fecha2.Enabled = Si
r_fecha3.Enabled = Si

r_causas.Enabled = Si
r_acciones.Enabled = Si

r_CostoEstimado.Enabled = Si

op_ac(0).Enabled = Si
op_ac(1).Enabled = Si
op_ac(2).Enabled = Si
op_ac(3).Enabled = Si
r_ac.Enabled = Si
r_comentarios.Enabled = Si

btnRAdjuntar.Enabled = Si
btnREliminarAdjuntos.Enabled = Si
BtnRFotosVer.Enabled = Si
btnRPdfVer.Enabled = Si

lblRAdjuntos.Enabled = False

End Sub
Private Sub e_FechaCierre_LostFocus()
d = Fecha_Valida(e_FechaCierre)
End Sub
Private Sub Doc_Grabar(nuevo As Boolean)

Dim TotalCampos As Integer

If Op_ER(0).Value Then ' graba datos del emisor

    Campos_Emisor_Definir

    ' limpia campos "valores"
    asn(1) = False
    TotalCampos = 16
    For i = 2 To TotalCampos
        av(i) = ""
        asn(i) = True
    Next
    
    av(1) = Numero.Text
    av(2) = Format(e_fecha.Text, sql_Fecha_Formato)
    av(3) = a_Emisores(0, CbEmisores.ListIndex)
    av(4) = a_Gerencias(0, CbGerencias.ListIndex)
    av(5) = a_Areas(0, CbAreas.ListIndex)
    
    If Op_NC(0).Value Then
        av(6) = "N" ' nc normal SGC
    End If
    If Op_NC(1).Value Then
        av(6) = "R" ' nc PNC
    End If
    If Op_NC(2).Value Then
        av(6) = "P" ' nc potencial
    End If
    If Op_NC(3).Value Then
        av(6) = "C" ' nc del cliente
    End If
    
    av(7) = Left(e_descripcion.Text, 1000)
    
    If CbRevisadoPor(0).ListIndex > 0 Then av(8) = a_Emisores(0, CbRevisadoPor(0).ListIndex)
    av(9) = IIf(Op_efe1(0).Value, "S", "N")
    If CbRevisadoPor(1).ListIndex > 0 Then av(10) = a_Emisores(0, CbRevisadoPor(1).ListIndex)
    av(11) = IIf(Op_efe2(0).Value, "S", "N")
    If CbRevisadoPor(2).ListIndex > 0 Then av(12) = a_Emisores(0, CbRevisadoPor(2).ListIndex)
    av(13) = IIf(Op_efe3(0).Value, "S", "N")
    
    av(14) = e_comentario.Text
    
    If e_Cerrado.Value = 1 Then
        av(15) = "S"
        av(16) = Format(e_FechaCierre.Text, sql_Fecha_Formato)
    Else
        av(15) = "N"
        av(16) = ""
    End If

End If

If Op_ER(1).Value Then ' graba datos del receptor

    Campos_Receptor_Definir
    asn(1) = False
    TotalCampos = 14 '8
    For i = 2 To TotalCampos
        av(i) = ""
        asn(i) = True
    Next

    ' receptor
    av(1) = Numero.Text
    av(2) = a_Encargados(0, CbEncargado.ListIndex)
    av(3) = r_causas.Text

    If r_fecha1.Text <> Fecha_Vacia Then av(4) = Format(r_fecha1.Text, sql_Fecha_Formato)
    If r_fecha2.Text <> Fecha_Vacia Then av(5) = Format(r_fecha2.Text, sql_Fecha_Formato)
    If r_fecha3.Text <> Fecha_Vacia Then av(6) = Format(r_fecha3.Text, sql_Fecha_Formato)

    av(7) = r_acciones
    av(8) = Val(r_CostoEstimado.Text)

    ' 17/07/2014
    ' campos nuvos agregados receptor
    av(9) = IIf(op_ac(0).Value, "S", "N") ' reproceso
    av(10) = IIf(op_ac(1).Value, "S", "N") ' reparacion
    av(11) = IIf(op_ac(2).Value, "S", "N") ' RFI
    av(12) = IIf(op_ac(3).Value, "S", "N") ' Otro
    av(13) = Trim(r_ac.Text)
    av(14) = Trim(r_comentarios.Text)

    ' fecha primera respuesta
    ' para que sea aceptada como fecha de primera respuesta del receptor
    ' debe digitar obligatoriamente estos cuatro campos:
    ' r_causas
    ' r_ac (accion correctiva)
    ' lblRAdjuntos (adjuntar evidencia)
    ' r_acciones (acciones preventivas)
    If lblFechaPrimeraRespuesta.Caption = "" Then
        If Len(r_causas.Text) > 0 And Len(r_ac.Text) > 0 _
        And Len(lblRAdjuntos.Text) > 0 And Len(r_acciones.Text) > 0 Then
        
            TotalCampos = 15
            asn(15) = True
            'av(15) = Date
            av(15) = Format(Date, sql_Fecha_Formato)
        End If
    End If

End If

If nuevo Then

    MousePointer = vbHourglass

    Registro_Agregar CnxSqlServer_scp0, TablaNombre, ac, av, TotalCampos

    'Email_Generar

    MousePointer = vbDefault

Else

    MousePointer = vbHourglass

    Registro_Modificar CnxSqlServer_scp0, TablaNombre, ac, av, asn, TotalCampos ' 12 ok

    'Email_Generar ' aqui va, cuando receptor modifica o responde NC

    MousePointer = vbDefault

End If

Email_Generar

Botones_Enabled 0, 0, 0, 0, 1, 0

End Sub
Private Sub Doc_Eliminar()
' borra cabecera

Registro_Eliminar CnxSqlServer_scp0, TablaNombre, Numero.Text
   
' elimina archivos
' ojo borrado de archivos
For i = 1 To 4
    m_NombreArchivoDestino = Numero.Text & "_" & i & ".jpg"
    If Archivo_Existe(m_PathDestino, m_NombreArchivoDestino) Then
        Kill m_PathDestino & m_NombreArchivoDestino
    End If
Next

End Sub
Private Sub Doc_ImprimirXXX()
Dim indice As Integer
' cabecera
prt.Print Empresa.Razon
prt.Print "GIRO: " & Empresa.Giro
prt.Print Empresa.Direccion
prt.Print "Teléfono: " & Empresa.Telefono1 & " - " & Empresa.Comuna
prt.Print ""

prt.Font.Size = 15
prt.Print Tab(10); "NO CONFORMIDAD Nº" & Numero.Text
prt.Font.Size = 10
prt.Print ""

prt.Print Tab(0); "FECHA     : " & e_fecha.Text
prt.Print Tab(0); "Nombre del Emisior : "; CbEmisores.Text
prt.Print Tab(0); "Gerencia donde se detectó la NO Conformidad: "; CbGerencias.Text
prt.Print Tab(0); "Area : "; CbAreas.Text
prt.Print ""
prt.Print Tab(0); "1.- Descripcion de la No Conformidad : "

txt2arreglo e_descripcion.Text, 85
Do While True
    prt.Print Tab(0); aTexto(0)
Loop

'prt.Print Tab(tab0); e_descripcion.Text
prt.Print ""

prt.Print Tab(0); "2.- Evidencia objetiva de las causas : "
prt.Print ""

prt.Print Tab(0); "3.- Investigación de las Causas (Explique el (los) motivo(s) que causa(n) la No Conformidad)"
txt2arreglo r_causas.Text, 85
indice = 0
Do While True
    indice = indice + 1
    If Trim(aTexto(indice)) = "" Then Exit Do
    prt.Print Tab(0); aTexto(indice)
Loop
'prt.Print Tab(tab0); r_causas.Text
prt.Print ""

prt.Print Tab(0); "4.- Acciones Adoptadas (Acciones correctivas / preventivas): (Explique)"
txt2arreglo r_acciones.Text, 85
indice = 0
Do While True
    indice = indice + 1
    If Trim(aTexto(indice)) = "" Then Exit Do
    prt.Print Tab(0); aTexto(indice)
Loop
prt.Print ""

'prt.Print Tab(tab0); r_acciones.Text
prt.Print Tab(0); "Fechas de Comprobación :"
prt.Print Tab(0); r_fecha1.Text; " " & r_fecha2.Text; " " & r_fecha3.Text
prt.Print Tab(0); "Persona Encargada de dar Solucion :"; CbEncargado.Text
prt.Print ""

prt.Print Tab(0); "5.- Revision Nº y Realizada Por"; Tab(40); "Efectividad de las Acciones"
prt.Print Tab(0); "1ª: "; CbRevisadoPor(0).Text; Tab(40); IIf(Op_efe1(0), "SI", "NO")
prt.Print Tab(0); "2ª: "; CbRevisadoPor(1).Text; Tab(40); IIf(Op_efe2(0), "SI", "NO")
prt.Print Tab(0); "3ª: "; CbRevisadoPor(2).Text; Tab(40); IIf(Op_efe3(0), "SI", "NO")
prt.Print ""

prt.Print Tab(0); "Comentarios:"
prt.Print Tab(0); e_comentario.Text

For i = 1 To 28
    prt.Print ""
Next

prt.Print Tab(0); Tab(14), "__________________", Tab(56), "__________________"
prt.Print Tab(0); Tab(14), "       VºBº       ", Tab(56), "       VºBº       "

prt.EndDoc

Impresora_Predeterminada "default"

MousePointer = vbDefault
End Sub
Private Sub Email_Generar()
' genera email
Dim intranet As String
Dim ServidorHTTP As String

Dim pagina_emi2rec As String ' de emisor a receptor
Dim pagina_rec2emi As String ' de receptor a emisor
Dim pagina_rec2enc As String ' de recpetor a encargado

Dim Parametros As String, textoWeb As String, txt As String

Inet.Protocol = icHTTP

intranet = ReadIniValue(Path_Local & "scp.ini", "Path", "intranet_server")
'ServidorHTTP = "HTTP://acr3006-dualpro/intranet/"
ServidorHTTP = intranet & "intranet/"

If Op_ER(0).Value Then  ' emisor

    If Test Then
        pagina_emi2rec = "correo_nc_e2r_test.asp" ' emisor a receptor
    Else
        pagina_emi2rec = "correo_nc_e2r.asp" ' emisor a receptor
    End If
    Parametros = "?nc=" & Numero.Text

    textoWeb = ServidorHTTP & pagina_emi2rec & Parametros
    Debug.Print textoWeb
    txt = Inet.OpenURL(textoWeb)
    Debug.Print txt
    
    If Left(txt, 2) = "OK" Then
        MsgBox "Se ha enviado un correo al Gerente de " & CbGerencias.Text & vbLf & "a: " & Mid(txt, 4)
    Else
        MsgBox "NO se pudo enviar correo" & vbLf & "Error|" & txt & "|e1|"
    End If

End If

If Op_ER(1).Value Then  ' receptor

    Parametros = "?nc=" & Numero.Text

    If Test Then
        pagina_rec2emi = "correo_nc_r2e_test.asp" ' receptor a emisor
    Else
        pagina_rec2emi = "correo_nc_r2e.asp" ' receptor a emisor
    End If

    textoWeb = ServidorHTTP & pagina_rec2emi & Parametros
    'Debug.Print textoWeb
    txt = Inet.OpenURL(textoWeb)
    'Debug.Print txt

    If Left(txt, 2) = "OK" Then
        MsgBox "Se ha enviado un correo a " & CbEmisores.Text & vbLf & "a: " & Mid(txt, 4)
    Else
        MsgBox "NO se pudo enviar correo" & vbLf & "Error: " & txt & "|r1|"
    End If

    ' envia correo a encargado (de dar solucion)
    pagina_rec2enc = "correo_nc_r2enc.asp" ' receptor a encargado (de dar solucion)
    textoWeb = ServidorHTTP & pagina_rec2enc & Parametros
    txt = Inet.OpenURL(textoWeb)
    If Left(txt, 2) = "OK" Then
        MsgBox "Se ha enviado un correo a " & CbEncargado.Text & vbLf & "a: " & Mid(txt, 4)
    Else
        MsgBox "NO se pudo enviar correo" & vbLf & "Error: " & txt & "|r2|"
    End If

End If




End Sub
Private Function txt2arreglo(ByVal Texto As String, ByVal LargoLinea As Integer) As Integer
' traspasa un texto largo en varias lineas, cada linea es un elelmento del arreglo
' sirve pra imprimir textos muy largos en varias lineas
' devuelve arreglo "publico" atexto() as string
Dim txt As String, inicio As Integer, lineatxt As String, pub As Integer, TextoRestante As String, Sigue As Boolean
Dim indice As Integer

If False Then

    Debug.Print "txt|" & txt & "|"
    Debug.Print ""
    
    inicio = 1
    TextoRestante = Mid(txt, inicio)
    txt = Mid(TextoRestante, inicio, LargoLinea)
    pub = InStrLast(txt, " ")
    Debug.Print "TextoRestante|" & TextoRestante & "|"
    txt = Mid(TextoRestante, 1, pub - 1)
    
    Debug.Print "pub1|" & pub & "|"
    Debug.Print "txt1a|" & txt & "|"
    Debug.Print ""
    
    inicio = pub + 1
    TextoRestante = Mid(TextoRestante, inicio)
    txt = Mid(TextoRestante, 1, LargoLinea)
    pub = InStrLast(txt, " ")
    Debug.Print "TextoRestante|" & TextoRestante & "|"
    txt = Mid(TextoRestante, 1, pub - 1)
    Debug.Print "pub2|" & pub & "|"
    Debug.Print "txt2a|" & txt & "|"
    Debug.Print ""
    
    inicio = pub + 1
    TextoRestante = Mid(TextoRestante, inicio)
    txt = Mid(TextoRestante, 1, LargoLinea)
    pub = InStrLast(txt, " ")
    Debug.Print "TextoRestante|" & TextoRestante & "|"
    txt = Mid(TextoRestante, 1, pub - 1)
    Debug.Print "pub3|" & pub & "|"
    Debug.Print "txt3a|" & txt & "|"
    
End If

'For indice = 1 To Len(Texto)
'    txt = Mid(Texto, indice, 1)
'    Debug.Print txt, Asc(txt)
'Next

' limpia arreglo
For indice = 1 To 9
    aTexto(indice) = ""
Next

' cambio de enter por vblf
Texto = Replace(Texto, Chr(13), " ")
Texto = Replace(Texto, Chr(10), " ")

TextoRestante = Texto

pub = 0
Sigue = True
indice = 0
Do While Sigue

    inicio = pub + 1
    TextoRestante = Mid(TextoRestante, inicio)
    
    If Len(TextoRestante) > LargoLinea Then
        txt = Mid(TextoRestante, 1, LargoLinea)
        pub = InStrLast(txt, " ")
        txt = Mid(TextoRestante, 1, pub - 1)
    Else
        txt = TextoRestante
        Sigue = False
    End If
    indice = indice + 1
    aTexto(indice) = txt
'    Debug.Print "TextoRestante|" & TextoRestante & "|"
'    Debug.Print "pub|" & pub & "|"
'    Debug.Print "txt|" & txt & "|"
'    Debug.Print ""
        
Loop

txt2arreglo = 0

End Function
Public Sub NC_Prepara(Numero As String, ImpresoraNombre As String)
' prepara archivo de reporte NC para imprimir

Dim fi As Double, m_desc As String, Tipo As String, pos As Integer
Dim p_dir As String, p_com As String, p_tel As String, p_fax As String
Dim pDesc_a_Dinero As Double

'////////////////////////////////////////
Dim Dbi As Database, RsRNC As Recordset
Set Dbi = OpenDatabase(repo_file)
Set RsRNC = Dbi.OpenRecordset("nc")
'////////////////////////////////////////

ImpresoraNombre = UCase(ImpresoraNombre)

With RsRNC

Dbi.Execute "delete * from [nc]"

.AddNew

!Numero = Numero
!e_fecha = e_fecha.Text
!e_nombre = CbEmisores.Text
!e_gerencia = CbGerencias.Text
!e_tipo = IIf(Op_NC(0).Value, Op_NC(0).Caption, "") & IIf(Op_NC(1).Value, Op_NC(1).Caption, "") & IIf(Op_NC(2).Value, Op_NC(2).Caption, "")
!e_descripcion = e_descripcion.Text + " "
!e_evidencia = Left(lblEAdjuntos.Text, 254) + " "
!r_investigacion = r_causas.Text + IIf(Len(r_causas.Text) = 0, " ", "")
!r_accionCorrectivaOpciones = IIf(op_ac(0).Value, op_ac(0).Caption, "") & "  " & IIf(op_ac(1).Value, op_ac(1).Caption, "") & "  " & IIf(op_ac(2).Value, op_ac(2).Caption, "") & "  " & IIf(op_ac(3).Value, op_ac(3).Caption, "")
!r_accionCorrectiva = r_ac.Text + " "
!r_accionPreventiva = r_acciones.Text + IIf(Len(r_acciones.Text) = 0, " ", "")
!r_evidencia = Left(lblRAdjuntos.Text, 254)
!r_comentarios = r_comentarios.Text + IIf(Len(r_comentarios.Text) = 0, " ", "")
!r_encargadonombre = CbEncargado.Text
!r_gerencianombre = CbGerencias.Text
!r_areanombre = CbAreas.Text
If (e_Cerrado.Value) Then
    If e_FechaCierre.Text = "__/__/__" Then
    Else
        !e_FechaCierre = e_FechaCierre.Text
    End If
End If

.Update

End With

End Sub
Private Sub NC_Print()
cr.WindowTitle = "NC Nº " & Numero.Text
cr.ReportSource = crptReport
cr.WindowState = crptMaximized
'Cr.WindowBorderStyle = crptFixedSingle
cr.WindowMaxButton = False
cr.WindowMinButton = False
cr.DataFiles(0) = repo_file & ".MDB"
cr.ReportFileName = Drive_Server & Path_Rpt & "nc.rpt"
cr.Action = 1
End Sub

VERSION 5.00
Begin VB.Form Usr_SubMenu 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Usuarios - SubMenú"
   ClientHeight    =   9600
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   7470
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9600
   ScaleWidth      =   7470
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame 
      BorderStyle     =   0  'None
      Height          =   255
      Index           =   24
      Left            =   120
      TabIndex        =   152
      Top             =   9120
      Width           =   1935
      Begin VB.OptionButton A0 
         Height          =   255
         Index           =   24
         Left            =   120
         TabIndex        =   156
         Top             =   0
         Width           =   255
      End
      Begin VB.OptionButton A1 
         Height          =   255
         Index           =   24
         Left            =   600
         TabIndex        =   155
         Top             =   0
         Width           =   255
      End
      Begin VB.OptionButton A2 
         Caption         =   "Option1"
         Height          =   255
         Index           =   24
         Left            =   1080
         TabIndex        =   154
         Top             =   0
         Width           =   255
      End
      Begin VB.OptionButton A3 
         Height          =   255
         Index           =   24
         Left            =   1560
         TabIndex        =   153
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.Frame Frame 
      BorderStyle     =   0  'None
      Height          =   255
      Index           =   23
      Left            =   120
      TabIndex        =   146
      Top             =   8760
      Width           =   1935
      Begin VB.OptionButton A0 
         Height          =   255
         Index           =   23
         Left            =   120
         TabIndex        =   150
         Top             =   0
         Width           =   255
      End
      Begin VB.OptionButton A1 
         Height          =   255
         Index           =   23
         Left            =   600
         TabIndex        =   149
         Top             =   0
         Width           =   255
      End
      Begin VB.OptionButton A2 
         Caption         =   "Option1"
         Height          =   255
         Index           =   23
         Left            =   1080
         TabIndex        =   148
         Top             =   0
         Width           =   255
      End
      Begin VB.OptionButton A3 
         Height          =   255
         Index           =   23
         Left            =   1560
         TabIndex        =   147
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.Frame Frame 
      BorderStyle     =   0  'None
      Height          =   255
      Index           =   22
      Left            =   120
      TabIndex        =   140
      Top             =   8400
      Width           =   1935
      Begin VB.OptionButton A0 
         Height          =   255
         Index           =   22
         Left            =   120
         TabIndex        =   144
         Top             =   0
         Width           =   255
      End
      Begin VB.OptionButton A1 
         Height          =   255
         Index           =   22
         Left            =   600
         TabIndex        =   143
         Top             =   0
         Width           =   255
      End
      Begin VB.OptionButton A2 
         Caption         =   "Option1"
         Height          =   255
         Index           =   22
         Left            =   1080
         TabIndex        =   142
         Top             =   0
         Width           =   255
      End
      Begin VB.OptionButton A3 
         Height          =   255
         Index           =   22
         Left            =   1560
         TabIndex        =   141
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.Frame Frame 
      BorderStyle     =   0  'None
      Height          =   255
      Index           =   21
      Left            =   120
      TabIndex        =   135
      Top             =   8040
      Width           =   1935
      Begin VB.OptionButton A3 
         Height          =   255
         Index           =   21
         Left            =   1560
         TabIndex        =   139
         Top             =   0
         Width           =   375
      End
      Begin VB.OptionButton A2 
         Caption         =   "Option1"
         Height          =   255
         Index           =   21
         Left            =   1080
         TabIndex        =   138
         Top             =   0
         Width           =   255
      End
      Begin VB.OptionButton A1 
         Height          =   255
         Index           =   21
         Left            =   600
         TabIndex        =   137
         Top             =   0
         Width           =   255
      End
      Begin VB.OptionButton A0 
         Height          =   255
         Index           =   21
         Left            =   120
         TabIndex        =   136
         Top             =   0
         Width           =   255
      End
   End
   Begin VB.Frame Frame 
      BorderStyle     =   0  'None
      Height          =   255
      Index           =   20
      Left            =   120
      TabIndex        =   129
      Top             =   7680
      Width           =   1935
      Begin VB.OptionButton A0 
         Height          =   255
         Index           =   20
         Left            =   120
         TabIndex        =   133
         Top             =   0
         Width           =   255
      End
      Begin VB.OptionButton A1 
         Height          =   255
         Index           =   20
         Left            =   600
         TabIndex        =   132
         Top             =   0
         Width           =   255
      End
      Begin VB.OptionButton A2 
         Caption         =   "Option1"
         Height          =   255
         Index           =   20
         Left            =   1080
         TabIndex        =   131
         Top             =   0
         Width           =   255
      End
      Begin VB.OptionButton A3 
         Height          =   255
         Index           =   20
         Left            =   1560
         TabIndex        =   130
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.Frame Frame 
      BorderStyle     =   0  'None
      Height          =   255
      Index           =   19
      Left            =   120
      TabIndex        =   115
      Top             =   7320
      Width           =   1935
      Begin VB.OptionButton A3 
         Height          =   255
         Index           =   19
         Left            =   1560
         TabIndex        =   119
         Top             =   0
         Width           =   375
      End
      Begin VB.OptionButton A2 
         Caption         =   "Option1"
         Height          =   255
         Index           =   19
         Left            =   1080
         TabIndex        =   118
         Top             =   0
         Width           =   255
      End
      Begin VB.OptionButton A1 
         Height          =   255
         Index           =   19
         Left            =   600
         TabIndex        =   117
         Top             =   0
         Width           =   255
      End
      Begin VB.OptionButton A0 
         Height          =   255
         Index           =   19
         Left            =   120
         TabIndex        =   116
         Top             =   0
         Width           =   255
      End
   End
   Begin VB.Frame Frame 
      BorderStyle     =   0  'None
      Caption         =   "109"
      Height          =   255
      Index           =   18
      Left            =   120
      TabIndex        =   109
      Top             =   6960
      Width           =   1935
      Begin VB.OptionButton A3 
         Height          =   255
         Index           =   18
         Left            =   1560
         TabIndex        =   113
         Top             =   0
         Width           =   255
      End
      Begin VB.OptionButton A2 
         Height          =   255
         Index           =   18
         Left            =   1080
         TabIndex        =   112
         Top             =   0
         Width           =   255
      End
      Begin VB.OptionButton A1 
         Height          =   255
         Index           =   18
         Left            =   600
         TabIndex        =   111
         Top             =   0
         Width           =   255
      End
      Begin VB.OptionButton A0 
         Height          =   255
         Index           =   18
         Left            =   120
         TabIndex        =   110
         Top             =   0
         Width           =   255
      End
   End
   Begin VB.Frame Frame 
      BorderStyle     =   0  'None
      Height          =   255
      Index           =   17
      Left            =   120
      TabIndex        =   103
      Top             =   6600
      Width           =   1935
      Begin VB.OptionButton A3 
         Height          =   255
         Index           =   17
         Left            =   1560
         TabIndex        =   107
         Top             =   0
         Width           =   255
      End
      Begin VB.OptionButton A0 
         Height          =   255
         Index           =   17
         Left            =   120
         TabIndex        =   104
         Top             =   0
         Width           =   255
      End
      Begin VB.OptionButton A1 
         Height          =   255
         Index           =   17
         Left            =   600
         TabIndex        =   105
         Top             =   0
         Width           =   255
      End
      Begin VB.OptionButton A2 
         Height          =   255
         Index           =   17
         Left            =   1080
         TabIndex        =   106
         Top             =   0
         Width           =   255
      End
   End
   Begin VB.Frame Frame 
      BorderStyle     =   0  'None
      Height          =   255
      Index           =   16
      Left            =   120
      TabIndex        =   97
      Top             =   6240
      Width           =   1935
      Begin VB.OptionButton A3 
         Caption         =   "Option12"
         Height          =   255
         Index           =   16
         Left            =   1560
         TabIndex        =   101
         Top             =   0
         Width           =   255
      End
      Begin VB.OptionButton A2 
         Caption         =   "Option11"
         Height          =   255
         Index           =   16
         Left            =   1080
         TabIndex        =   100
         Top             =   0
         Width           =   255
      End
      Begin VB.OptionButton A1 
         Caption         =   "Option10"
         Height          =   255
         Index           =   16
         Left            =   600
         TabIndex        =   99
         Top             =   0
         Width           =   255
      End
      Begin VB.OptionButton A0 
         Caption         =   "Option9"
         Height          =   255
         Index           =   16
         Left            =   120
         TabIndex        =   98
         Top             =   0
         Width           =   255
      End
   End
   Begin VB.Frame Frame 
      BorderStyle     =   0  'None
      Height          =   255
      Index           =   15
      Left            =   120
      TabIndex        =   91
      Top             =   5880
      Width           =   1935
      Begin VB.OptionButton A0 
         Caption         =   "Option9"
         Height          =   255
         Index           =   15
         Left            =   120
         TabIndex        =   92
         Top             =   0
         Width           =   255
      End
      Begin VB.OptionButton A1 
         Caption         =   "Option10"
         Height          =   255
         Index           =   15
         Left            =   600
         TabIndex        =   93
         Top             =   0
         Width           =   255
      End
      Begin VB.OptionButton A2 
         Caption         =   "Option11"
         Height          =   255
         Index           =   15
         Left            =   1080
         TabIndex        =   94
         Top             =   0
         Width           =   255
      End
      Begin VB.OptionButton A3 
         Caption         =   "Option12"
         Height          =   255
         Index           =   15
         Left            =   1560
         TabIndex        =   95
         Top             =   0
         Width           =   255
      End
   End
   Begin VB.Frame Frame 
      BorderStyle     =   0  'None
      Height          =   255
      Index           =   14
      Left            =   120
      TabIndex        =   85
      Top             =   5520
      Width           =   1935
      Begin VB.OptionButton A0 
         Caption         =   "Option9"
         Height          =   255
         Index           =   14
         Left            =   120
         TabIndex        =   86
         Top             =   0
         Width           =   255
      End
      Begin VB.OptionButton A1 
         Caption         =   "Option10"
         Height          =   255
         Index           =   14
         Left            =   600
         TabIndex        =   87
         Top             =   0
         Width           =   255
      End
      Begin VB.OptionButton A2 
         Caption         =   "Option11"
         Height          =   255
         Index           =   14
         Left            =   1080
         TabIndex        =   88
         Top             =   0
         Width           =   255
      End
      Begin VB.OptionButton A3 
         Caption         =   "Option12"
         Height          =   255
         Index           =   14
         Left            =   1560
         TabIndex        =   89
         Top             =   0
         Width           =   255
      End
   End
   Begin VB.Frame Frame 
      BorderStyle     =   0  'None
      Height          =   255
      Index           =   13
      Left            =   120
      TabIndex        =   79
      Top             =   5160
      Width           =   1935
      Begin VB.OptionButton A0 
         Caption         =   "Option9"
         Height          =   255
         Index           =   13
         Left            =   120
         TabIndex        =   80
         Top             =   0
         Width           =   255
      End
      Begin VB.OptionButton A1 
         Caption         =   "Option10"
         Height          =   255
         Index           =   13
         Left            =   600
         TabIndex        =   81
         Top             =   0
         Width           =   255
      End
      Begin VB.OptionButton A2 
         Caption         =   "Option11"
         Height          =   255
         Index           =   13
         Left            =   1080
         TabIndex        =   82
         Top             =   0
         Width           =   255
      End
      Begin VB.OptionButton A3 
         Caption         =   "Option12"
         Height          =   255
         Index           =   13
         Left            =   1560
         TabIndex        =   83
         Top             =   0
         Width           =   255
      End
   End
   Begin VB.Frame Frame 
      BorderStyle     =   0  'None
      Height          =   255
      Index           =   12
      Left            =   120
      TabIndex        =   73
      Top             =   4800
      Width           =   1935
      Begin VB.OptionButton A3 
         Caption         =   "Option12"
         Height          =   255
         Index           =   12
         Left            =   1560
         TabIndex        =   77
         Top             =   0
         Width           =   255
      End
      Begin VB.OptionButton A2 
         Caption         =   "Option11"
         Height          =   255
         Index           =   12
         Left            =   1080
         TabIndex        =   76
         Top             =   0
         Width           =   255
      End
      Begin VB.OptionButton A1 
         Caption         =   "Option10"
         Height          =   255
         Index           =   12
         Left            =   600
         TabIndex        =   75
         Top             =   0
         Width           =   255
      End
      Begin VB.OptionButton A0 
         Caption         =   "Option9"
         Height          =   255
         Index           =   12
         Left            =   120
         TabIndex        =   74
         Top             =   0
         Width           =   255
      End
   End
   Begin VB.Frame Frame 
      BorderStyle     =   0  'None
      Height          =   255
      Index           =   11
      Left            =   120
      TabIndex        =   67
      Top             =   4440
      Width           =   1935
      Begin VB.OptionButton A3 
         Caption         =   "Option8"
         Height          =   255
         Index           =   11
         Left            =   1560
         TabIndex        =   71
         Top             =   0
         Width           =   255
      End
      Begin VB.OptionButton A2 
         Caption         =   "Option7"
         Height          =   255
         Index           =   11
         Left            =   1080
         TabIndex        =   70
         Top             =   0
         Width           =   255
      End
      Begin VB.OptionButton A1 
         Caption         =   "Option6"
         Height          =   255
         Index           =   11
         Left            =   600
         TabIndex        =   69
         Top             =   0
         Width           =   255
      End
      Begin VB.OptionButton A0 
         Caption         =   "Option5"
         Height          =   255
         Index           =   11
         Left            =   120
         TabIndex        =   68
         Top             =   0
         Width           =   255
      End
   End
   Begin VB.Frame Frame 
      BorderStyle     =   0  'None
      Height          =   255
      Index           =   10
      Left            =   120
      TabIndex        =   61
      Top             =   4080
      Width           =   1935
      Begin VB.OptionButton A3 
         Caption         =   "Option4"
         Height          =   195
         Index           =   10
         Left            =   1560
         TabIndex        =   65
         Top             =   0
         Width           =   255
      End
      Begin VB.OptionButton A2 
         Caption         =   "Option3"
         Height          =   195
         Index           =   10
         Left            =   1080
         TabIndex        =   64
         Top             =   0
         Width           =   255
      End
      Begin VB.OptionButton A1 
         Caption         =   "Option2"
         Height          =   195
         Index           =   10
         Left            =   600
         TabIndex        =   63
         Top             =   0
         Width           =   255
      End
      Begin VB.OptionButton A0 
         Caption         =   "Option1"
         Height          =   195
         Index           =   10
         Left            =   120
         TabIndex        =   62
         Top             =   0
         Width           =   255
      End
   End
   Begin VB.Frame Framel 
      Caption         =   "Leyenda"
      Height          =   1815
      Left            =   5760
      TabIndex        =   121
      Top             =   1440
      Width           =   1575
      Begin VB.Label lbl 
         Caption         =   "3 Total"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   125
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label lbl 
         Caption         =   "2 Lee y Escribe"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   124
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label lbl 
         Caption         =   "1 Sólo Lectura"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   123
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label lbl 
         Caption         =   "0 Nulo"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   122
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.CommandButton btnCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   6000
      TabIndex        =   127
      Top             =   3960
      Width           =   1095
   End
   Begin VB.CommandButton btnAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   6000
      TabIndex        =   126
      Top             =   3480
      Width           =   1095
   End
   Begin VB.Frame Frame 
      BorderStyle     =   0  'None
      Height          =   255
      Index           =   9
      Left            =   120
      TabIndex        =   55
      Top             =   3720
      Width           =   1935
      Begin VB.OptionButton A3 
         Caption         =   "Option36"
         Height          =   255
         Index           =   9
         Left            =   1560
         TabIndex        =   59
         Top             =   0
         Width           =   255
      End
      Begin VB.OptionButton A2 
         Caption         =   "Option35"
         Height          =   255
         Index           =   9
         Left            =   1080
         TabIndex        =   58
         Top             =   0
         Width           =   255
      End
      Begin VB.OptionButton A1 
         Caption         =   "Option34"
         Height          =   255
         Index           =   9
         Left            =   600
         TabIndex        =   57
         Top             =   0
         Width           =   255
      End
      Begin VB.OptionButton A0 
         Caption         =   "Option33"
         Height          =   255
         Index           =   9
         Left            =   120
         TabIndex        =   56
         Top             =   0
         Width           =   255
      End
   End
   Begin VB.Frame Frame 
      BorderStyle     =   0  'None
      Height          =   255
      Index           =   8
      Left            =   120
      TabIndex        =   49
      Top             =   3360
      Width           =   1935
      Begin VB.OptionButton A3 
         Caption         =   "Option32"
         Height          =   255
         Index           =   8
         Left            =   1560
         TabIndex        =   53
         Top             =   0
         Width           =   255
      End
      Begin VB.OptionButton A2 
         Caption         =   "Option31"
         Height          =   255
         Index           =   8
         Left            =   1080
         TabIndex        =   52
         Top             =   0
         Width           =   255
      End
      Begin VB.OptionButton A1 
         Caption         =   "Option30"
         Height          =   255
         Index           =   8
         Left            =   600
         TabIndex        =   51
         Top             =   0
         Width           =   255
      End
      Begin VB.OptionButton A0 
         Caption         =   "Option29"
         Height          =   255
         Index           =   8
         Left            =   120
         TabIndex        =   50
         Top             =   0
         Width           =   255
      End
   End
   Begin VB.Frame Frame 
      BorderStyle     =   0  'None
      Height          =   255
      Index           =   7
      Left            =   120
      TabIndex        =   43
      Top             =   3000
      Width           =   1935
      Begin VB.OptionButton A3 
         Caption         =   "Option28"
         Height          =   255
         Index           =   7
         Left            =   1560
         TabIndex        =   47
         Top             =   0
         Width           =   255
      End
      Begin VB.OptionButton A2 
         Caption         =   "Option27"
         Height          =   255
         Index           =   7
         Left            =   1080
         TabIndex        =   46
         Top             =   0
         Width           =   255
      End
      Begin VB.OptionButton A1 
         Caption         =   "Option26"
         Height          =   255
         Index           =   7
         Left            =   600
         TabIndex        =   45
         Top             =   0
         Width           =   255
      End
      Begin VB.OptionButton A0 
         Caption         =   "Option25"
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   44
         Top             =   0
         Width           =   255
      End
   End
   Begin VB.Frame Frame 
      BorderStyle     =   0  'None
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   37
      Top             =   2640
      Width           =   1935
      Begin VB.OptionButton A3 
         Caption         =   "Option24"
         Height          =   255
         Index           =   6
         Left            =   1560
         TabIndex        =   41
         Top             =   0
         Width           =   255
      End
      Begin VB.OptionButton A2 
         Caption         =   "Option23"
         Height          =   255
         Index           =   6
         Left            =   1080
         TabIndex        =   40
         Top             =   0
         Width           =   255
      End
      Begin VB.OptionButton A1 
         Caption         =   "Option22"
         Height          =   255
         Index           =   6
         Left            =   600
         TabIndex        =   39
         Top             =   0
         Width           =   255
      End
      Begin VB.OptionButton A0 
         Caption         =   "Option21"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   38
         Top             =   0
         Width           =   255
      End
   End
   Begin VB.Frame Frame 
      BorderStyle     =   0  'None
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   31
      Top             =   2280
      Width           =   1935
      Begin VB.OptionButton A3 
         Caption         =   "Option20"
         Height          =   255
         Index           =   5
         Left            =   1560
         TabIndex        =   35
         Top             =   0
         Width           =   255
      End
      Begin VB.OptionButton A2 
         Caption         =   "Option19"
         Height          =   255
         Index           =   5
         Left            =   1080
         TabIndex        =   34
         Top             =   0
         Width           =   255
      End
      Begin VB.OptionButton A1 
         Caption         =   "Option18"
         Height          =   255
         Index           =   5
         Left            =   600
         TabIndex        =   33
         Top             =   0
         Width           =   255
      End
      Begin VB.OptionButton A0 
         Caption         =   "Option17"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   32
         Top             =   0
         Width           =   255
      End
   End
   Begin VB.Frame Frame 
      BorderStyle     =   0  'None
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   25
      Top             =   1920
      Width           =   1935
      Begin VB.OptionButton A3 
         Caption         =   "Option16"
         Height          =   255
         Index           =   4
         Left            =   1560
         TabIndex        =   29
         Top             =   0
         Width           =   255
      End
      Begin VB.OptionButton A2 
         Caption         =   "Option15"
         Height          =   255
         Index           =   4
         Left            =   1080
         TabIndex        =   28
         Top             =   0
         Width           =   255
      End
      Begin VB.OptionButton A1 
         Caption         =   "Option14"
         Height          =   255
         Index           =   4
         Left            =   600
         TabIndex        =   27
         Top             =   0
         Width           =   255
      End
      Begin VB.OptionButton A0 
         Caption         =   "Option13"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   26
         Top             =   0
         Width           =   255
      End
   End
   Begin VB.Frame Frame 
      BorderStyle     =   0  'None
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   19
      Top             =   1560
      Width           =   1935
      Begin VB.OptionButton A3 
         Caption         =   "Option12"
         Height          =   255
         Index           =   3
         Left            =   1560
         TabIndex        =   23
         Top             =   0
         Width           =   255
      End
      Begin VB.OptionButton A2 
         Caption         =   "Option11"
         Height          =   255
         Index           =   3
         Left            =   1080
         TabIndex        =   22
         Top             =   0
         Width           =   255
      End
      Begin VB.OptionButton A1 
         Caption         =   "Option10"
         Height          =   255
         Index           =   3
         Left            =   600
         TabIndex        =   21
         Top             =   0
         Width           =   255
      End
      Begin VB.OptionButton A0 
         Caption         =   "Option9"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   20
         Top             =   0
         Width           =   255
      End
   End
   Begin VB.Frame Frame 
      BorderStyle     =   0  'None
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   13
      Top             =   1200
      Width           =   1935
      Begin VB.OptionButton A3 
         Caption         =   "Option8"
         Height          =   255
         Index           =   2
         Left            =   1560
         TabIndex        =   17
         Top             =   0
         Width           =   255
      End
      Begin VB.OptionButton A2 
         Caption         =   "Option7"
         Height          =   255
         Index           =   2
         Left            =   1080
         TabIndex        =   16
         Top             =   0
         Width           =   255
      End
      Begin VB.OptionButton A1 
         Caption         =   "Option6"
         Height          =   255
         Index           =   2
         Left            =   600
         TabIndex        =   15
         Top             =   0
         Width           =   255
      End
      Begin VB.OptionButton A0 
         Caption         =   "Option5"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   14
         Top             =   0
         Width           =   255
      End
   End
   Begin VB.Frame Frame 
      BorderStyle     =   0  'None
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   7
      Top             =   840
      Width           =   1935
      Begin VB.OptionButton A3 
         Caption         =   "Option4"
         Height          =   255
         Index           =   1
         Left            =   1560
         TabIndex        =   11
         Top             =   0
         Width           =   255
      End
      Begin VB.OptionButton A2 
         Caption         =   "Option3"
         Height          =   255
         Index           =   1
         Left            =   1080
         TabIndex        =   10
         Top             =   0
         Width           =   255
      End
      Begin VB.OptionButton A1 
         Caption         =   "Option2"
         Height          =   255
         Index           =   1
         Left            =   600
         TabIndex        =   9
         Top             =   0
         Width           =   255
      End
      Begin VB.OptionButton A0 
         Caption         =   "Option1"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   8
         Top             =   0
         Width           =   255
      End
   End
   Begin VB.Frame Frame 
      BorderStyle     =   0  'None
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   1935
      Begin VB.OptionButton A0 
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   0
         Width           =   255
      End
      Begin VB.OptionButton A1 
         Height          =   255
         Index           =   0
         Left            =   600
         TabIndex        =   3
         Top             =   0
         Width           =   255
      End
      Begin VB.OptionButton A2 
         Height          =   255
         Index           =   0
         Left            =   1080
         TabIndex        =   4
         Top             =   0
         Width           =   255
      End
      Begin VB.OptionButton A3 
         Height          =   255
         Index           =   0
         Left            =   1560
         TabIndex        =   5
         Top             =   0
         Width           =   255
      End
   End
   Begin VB.Label Opcion 
      Caption         =   "Opcion(24)"
      Height          =   255
      Index           =   24
      Left            =   2160
      TabIndex        =   157
      Top             =   9120
      Width           =   3255
   End
   Begin VB.Label Opcion 
      Caption         =   "Opcion(23)"
      Height          =   255
      Index           =   23
      Left            =   2160
      TabIndex        =   151
      Top             =   8760
      Width           =   3255
   End
   Begin VB.Label Opcion 
      Caption         =   "Opcion(22)"
      Height          =   255
      Index           =   22
      Left            =   2160
      TabIndex        =   145
      Top             =   8400
      Width           =   3255
   End
   Begin VB.Label Opcion 
      Caption         =   "Opcion(21)"
      Height          =   255
      Index           =   21
      Left            =   2160
      TabIndex        =   134
      Top             =   8040
      Width           =   3255
   End
   Begin VB.Label Opcion 
      Caption         =   "Opcion(20)"
      Height          =   255
      Index           =   20
      Left            =   2160
      TabIndex        =   128
      Top             =   7680
      Width           =   3255
   End
   Begin VB.Label Opcion 
      Caption         =   "Opcion(19)"
      Height          =   255
      Index           =   19
      Left            =   2160
      TabIndex        =   120
      Top             =   7320
      Width           =   3255
   End
   Begin VB.Label Opcion 
      Caption         =   "Opcion(18)"
      Height          =   255
      Index           =   18
      Left            =   2160
      TabIndex        =   114
      Top             =   6960
      Width           =   3255
   End
   Begin VB.Label Opcion 
      Caption         =   "Opcion(17)"
      Height          =   255
      Index           =   17
      Left            =   2160
      TabIndex        =   108
      Top             =   6600
      Width           =   3255
   End
   Begin VB.Label Opcion 
      Caption         =   "Opcion(16)"
      Height          =   255
      Index           =   16
      Left            =   2160
      TabIndex        =   102
      Top             =   6240
      Width           =   3255
   End
   Begin VB.Label Opcion 
      Caption         =   "Opcion(15)"
      Height          =   255
      Index           =   15
      Left            =   2160
      TabIndex        =   96
      Top             =   5880
      Width           =   3255
   End
   Begin VB.Label Opcion 
      Caption         =   "Opcion(14)"
      Height          =   255
      Index           =   14
      Left            =   2160
      TabIndex        =   90
      Top             =   5520
      Width           =   3255
   End
   Begin VB.Label Opcion 
      Caption         =   "Opcion(13)"
      Height          =   255
      Index           =   13
      Left            =   2160
      TabIndex        =   84
      Top             =   5160
      Width           =   3255
   End
   Begin VB.Label Opcion 
      Caption         =   "Opcion(12)"
      Height          =   255
      Index           =   12
      Left            =   2160
      TabIndex        =   78
      Top             =   4800
      Width           =   3255
   End
   Begin VB.Label Opcion 
      Caption         =   "Opcion(11)"
      Height          =   255
      Index           =   11
      Left            =   2160
      TabIndex        =   72
      Top             =   4440
      Width           =   3255
   End
   Begin VB.Label Opcion 
      Caption         =   "Opcion(10)"
      Height          =   255
      Index           =   10
      Left            =   2160
      TabIndex        =   66
      Top             =   4080
      Width           =   3255
   End
   Begin VB.Label lbl 
      Caption         =   " 0      1      2      3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Opcion 
      Caption         =   "Opcion(9)"
      Height          =   255
      Index           =   9
      Left            =   2160
      TabIndex        =   60
      Top             =   3720
      Width           =   3255
   End
   Begin VB.Label Opcion 
      Caption         =   "Opcion(8)"
      Height          =   255
      Index           =   8
      Left            =   2160
      TabIndex        =   54
      Top             =   3360
      Width           =   3255
   End
   Begin VB.Label Opcion 
      Caption         =   "Opcion(7)"
      Height          =   255
      Index           =   7
      Left            =   2160
      TabIndex        =   48
      Top             =   3000
      Width           =   3255
   End
   Begin VB.Label Opcion 
      Caption         =   "Opcion(6)"
      Height          =   255
      Index           =   6
      Left            =   2160
      TabIndex        =   42
      Top             =   2640
      Width           =   3255
   End
   Begin VB.Label Opcion 
      Caption         =   "Opcion(5)"
      Height          =   255
      Index           =   5
      Left            =   2160
      TabIndex        =   36
      Top             =   2280
      Width           =   3255
   End
   Begin VB.Label Opcion 
      Caption         =   "Opcion(4)"
      Height          =   255
      Index           =   4
      Left            =   2160
      TabIndex        =   30
      Top             =   1920
      Width           =   3255
   End
   Begin VB.Label Opcion 
      Caption         =   "Opcion(3)"
      Height          =   255
      Index           =   3
      Left            =   2160
      TabIndex        =   24
      Top             =   1560
      Width           =   3255
   End
   Begin VB.Label Opcion 
      Caption         =   "Opcion(2)"
      Height          =   255
      Index           =   2
      Left            =   2160
      TabIndex        =   18
      Top             =   1200
      Width           =   3255
   End
   Begin VB.Label Opcion 
      Caption         =   "Opcion(1)"
      Height          =   255
      Index           =   1
      Left            =   2160
      TabIndex        =   12
      Top             =   840
      Width           =   3255
   End
   Begin VB.Label Opcion 
      Caption         =   "Opcion(0)"
      Height          =   255
      Index           =   0
      Left            =   2160
      TabIndex        =   6
      Top             =   480
      Width           =   3255
   End
End
Attribute VB_Name = "Usr_SubMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' privilegios de usuario
' es decir, opciones del menu que estan disponibles
Option Explicit
Private i As Integer, j As Integer, k As Integer
Private m_Menu As Integer
'Private m_TotalOpciones As Integer ' del submenu
Private m_Sm As String
'////////////////////////////////////////////////////////////////////
Public Property Get Usr_Menu() As Integer
Usr_Menu = m_Menu
End Property
Public Property Let Usr_Menu(ByVal vNewValue As Integer)
m_Menu = vNewValue
End Property
Public Property Get Usr_SMenu() As String
Usr_SMenu = m_Sm
End Property
Public Property Let Usr_SMenu(ByVal vNewValue As String)
m_Sm = vNewValue
End Property

'////////////////////////////////////////////////////////////////////
Private Sub Form_Load()
Dim Fin As Integer
'm_TotalOpciones = 15
On Error GoTo Sigue
For i = 0 To Menu_Filas - 1
    Select Case m_Menu
    Case 1
        Opcion(i).Caption = menu.Mnu1(i + 1).Caption
    Case 2
        Opcion(i).Caption = menu.Mnu2(i + 1).Caption
    Case 3
        Opcion(i).Caption = menu.Mnu3(i + 1).Caption
    Case 4
        Opcion(i).Caption = menu.Mnu4(i + 1).Caption
    Case 5
        Opcion(i).Caption = menu.Mnu5(i + 1).Caption
    Case 6
        Opcion(i).Caption = menu.Mnu6(i + 1).Caption
    Case 7
        Opcion(i).Caption = menu.Mnu7(i + 1).Caption
    Case 8
        Opcion(i).Caption = menu.Mnu8(i + 1).Caption
    Case 9
        Opcion(i).Caption = menu.Mnu9(i + 1).Caption
    Case 10
        Opcion(i).Caption = menu.Mnu10(i + 1).Caption
    Case 11
        Opcion(i).Caption = menu.Mnu11(i + 1).Caption
    End Select
    A0(i).Value = IIf(Val(Mid(m_Sm, i + 1, 1)) = 0, True, False)
    A1(i).Value = IIf(Val(Mid(m_Sm, i + 1, 1)) = 1, True, False)
    A2(i).Value = IIf(Val(Mid(m_Sm, i + 1, 1)) = 2, True, False)
    A3(i).Value = IIf(Val(Mid(m_Sm, i + 1, 1)) = 3, True, False)
    Fin = i
Next
Sigue:
For i = Fin + 1 To Menu_Filas - 1
    Opcion(i).visible = False
    frame(i).visible = False
Next

End Sub
Private Sub btnAceptar_Click()
m_Sm = ""
For i = 0 To Menu_Filas - 1
    Select Case True
    Case A0(i).Value
        m_Sm = m_Sm + "0"
    Case A1(i).Value
        m_Sm = m_Sm + "1"
    Case A2(i).Value
        m_Sm = m_Sm + "2"
    Case A3(i).Value
        m_Sm = m_Sm + "3"
    Case Else
        m_Sm = m_Sm + "0"
    End Select
Next

Unload Me
End Sub
Private Sub btnCancelar_Click()
Unload Me
End Sub

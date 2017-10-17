VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form email_frm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "email_frm"
   ClientHeight    =   960
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2100
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   960
   ScaleWidth      =   2100
   StartUpPosition =   3  'Windows Default
   Begin InetCtlsObjects.Inet Inet 
      Left            =   240
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
   End
End
Attribute VB_Name = "email_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

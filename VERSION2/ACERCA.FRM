VERSION 2.00
Begin Form Acerca_de 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Double
   ClientHeight    =   2445
   ClientLeft      =   2115
   ClientTop       =   2925
   ClientWidth     =   4170
   ControlBox      =   0   'False
   DrawMode        =   1  'Blackness
   FillStyle       =   0  'Solid
   Height          =   2850
   Left            =   2055
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2445
   ScaleWidth      =   4170
   Top             =   2580
   Width           =   4290
   Begin CommandButton Cancelar 
      Caption         =   "&Cerrar"
      Height          =   315
      Left            =   1635
      TabIndex        =   4
      Top             =   1980
      Width           =   840
   End
   Begin PictureBox Logo 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   480
      Left            =   90
      Picture         =   ACERCA.FRX:0000
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   0
      Top             =   75
      Width           =   480
   End
   Begin Label Label4 
      Caption         =   "Todos los Derechos Reservados"
      Height          =   240
      Left            =   645
      TabIndex        =   6
      Top             =   1260
      Width           =   2790
   End
   Begin Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "MSI-Two Way para  Telxon, Symbol y Norand "
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   195
      TabIndex        =   5
      Top             =   630
      Width           =   3915
   End
   Begin Label Acerca_de 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Arequipa - Perú"
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   1470
      TabIndex        =   3
      Top             =   1560
      Width           =   1380
   End
   Begin Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Copyright 1996 por Soluciones Informáticas "
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   195
      TabIndex        =   2
      Top             =   930
      Width           =   3825
   End
   Begin Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Programa de Comunicaciones PC-Terminal para Windows"
      ForeColor       =   &H00000000&
      Height          =   525
      Left            =   1110
      TabIndex        =   1
      Top             =   120
      Width           =   2535
   End
End

Sub Cancelar_Click ()
    Unload Me
End Sub


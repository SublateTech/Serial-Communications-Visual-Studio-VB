VERSION 2.00
Begin Form Profile 
   ClientHeight    =   2100
   ClientLeft      =   1470
   ClientTop       =   2235
   ClientWidth     =   6375
   Height          =   2505
   Left            =   1410
   LinkTopic       =   "Form1"
   ScaleHeight     =   2100
   ScaleWidth      =   6375
   Top             =   1890
   Width           =   6495
   Begin CommandButton Cancelar 
      Caption         =   "Cancelar"
      Height          =   330
      Left            =   4770
      TabIndex        =   7
      Top             =   1170
      Width           =   1485
   End
   Begin CommandButton Aceptar 
      Caption         =   "Aceptar"
      Height          =   345
      Left            =   4770
      TabIndex        =   6
      Top             =   690
      Width           =   1485
   End
   Begin TextBox Archivo 
      Height          =   300
      Left            =   60
      TabIndex        =   3
      Top             =   390
      Width           =   1935
   End
   Begin FileListBox Archivos 
      Height          =   1200
      Left            =   60
      Pattern         =   "*.ini"
      TabIndex        =   4
      Top             =   660
      Width           =   1935
   End
   Begin DirListBox Directorios 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   1155
      Left            =   2130
      TabIndex        =   5
      Top             =   705
      Width           =   2565
   End
   Begin Label Label3 
      Caption         =   "&Directorios"
      Height          =   255
      Left            =   2100
      TabIndex        =   0
      Top             =   120
      Width           =   1665
   End
   Begin Label Directorio_Label 
      Height          =   240
      Left            =   2115
      TabIndex        =   2
      Top             =   405
      Width           =   2565
   End
   Begin Label Label1 
      Caption         =   "&Nombre del archivo:"
      Height          =   285
      Left            =   75
      TabIndex        =   1
      Top             =   120
      Width           =   2055
   End
End

Sub Aceptar_Click ()
    
    If Len(Archivo.Text) <> 0 Then
       cProfile = Archivo.Text
    End If
    Unload Me
End Sub

Sub Archivos_Click ()
    Archivo.Text = Archivos.FileName
End Sub

Sub Archivos_DblClick ()
    Call Aceptar_Click
End Sub

Sub Archivos_PathChange ()
    Archivos.Refresh
End Sub

Sub Cancelar_Click ()
Unload Me
End Sub

Sub Directorios_Change ()
    Directorio_Label.Caption = Directorios.Path
    Archivos.Path = Directorios.Path
End Sub

Sub Form_Load ()
    Directorios.Path = cDirectorio_Programa '"C:\TWOWAY\"
    Directorio_Label.Caption = Directorios.Path
End Sub


VERSION 2.00
Begin Form Setup 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Configuración"
   ClientHeight    =   4515
   ClientLeft      =   1770
   ClientTop       =   1755
   ClientWidth     =   5640
   Height          =   4920
   Left            =   1710
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4515
   ScaleWidth      =   5640
   Top             =   1410
   Width           =   5760
   Begin CommandButton Cmd_Cancelar 
      Caption         =   "&Cancelar"
      Height          =   450
      Left            =   2820
      TabIndex        =   27
      Top             =   4005
      Width           =   2745
   End
   Begin CommandButton Cmd_Grabar 
      Caption         =   "&Grabar"
      Height          =   435
      Left            =   75
      TabIndex        =   20
      Top             =   4020
      Width           =   2700
   End
   Begin Frame Directorios 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Directorios"
      Height          =   1200
      Left            =   75
      TabIndex        =   15
      Top             =   30
      Width           =   5505
      Begin TextBox Directorio_Recibir 
         Height          =   285
         Left            =   1140
         TabIndex        =   19
         Top             =   675
         Width           =   4245
      End
      Begin TextBox Directorio_Enviar 
         Height          =   285
         Left            =   1140
         TabIndex        =   18
         Top             =   315
         Width           =   4245
      End
      Begin Label Directorio_Recibir_Label 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Recibir:"
         Height          =   255
         Left            =   105
         TabIndex        =   17
         Top             =   720
         Width           =   915
      End
      Begin Label Directorio_Enviar_Label 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Envíar:"
         Height          =   255
         Left            =   105
         TabIndex        =   16
         Top             =   330
         Width           =   1035
      End
   End
   Begin Frame Modo_Horas 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Modo/Horas"
      Height          =   1665
      Left            =   2805
      TabIndex        =   11
      Top             =   2295
      Width           =   2790
      Begin TextBox Horas_Recibir 
         Height          =   315
         Left            =   1500
         TabIndex        =   24
         Top             =   1200
         Width           =   1185
      End
      Begin TextBox Horas_Enviar 
         Height          =   330
         Left            =   1500
         TabIndex        =   23
         Top             =   750
         Width           =   1185
      End
      Begin OptionButton Modo_Manual 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Manual"
         Height          =   240
         Left            =   1485
         TabIndex        =   14
         Top             =   330
         Width           =   1215
      End
      Begin OptionButton Modo_Automático 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Automático"
         Height          =   240
         Left            =   90
         TabIndex        =   21
         Top             =   330
         Value           =   -1  'True
         Width           =   1350
      End
      Begin Label Horas_Recibir_Label 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Horas Recibir:"
         Height          =   240
         Left            =   120
         TabIndex        =   26
         Top             =   1245
         Width           =   1290
      End
      Begin Label Horas_Enviar_Label 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Horas Envíar:"
         Height          =   285
         Left            =   120
         TabIndex        =   25
         Top             =   780
         Width           =   1260
      End
   End
   Begin Frame Puerto_Velocidad 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Puerto/Velocidad"
      Height          =   975
      Left            =   2790
      TabIndex        =   10
      Top             =   1260
      Width           =   2790
      Begin ComboBox Comms 
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   285
         TabIndex        =   22
         Top             =   420
         Width           =   1110
      End
      Begin OptionButton Velocidad_9600 
         BackColor       =   &H00C0C0C0&
         Caption         =   "  9600"
         Height          =   240
         Left            =   1700
         TabIndex        =   13
         Top             =   585
         Width           =   900
      End
      Begin OptionButton Velocidad_19200 
         BackColor       =   &H00C0C0C0&
         Caption         =   "19200 "
         Height          =   255
         Left            =   1700
         TabIndex        =   12
         Top             =   315
         Value           =   -1  'True
         Width           =   825
      End
   End
   Begin Frame Archivos 
      BackColor       =   &H00C0C0C0&
      Caption         =   " Archivos"
      Height          =   2700
      Left            =   75
      TabIndex        =   1
      Top             =   1260
      Width           =   2670
      Begin TextBox FileText 
         Height          =   330
         Left            =   60
         TabIndex        =   9
         Top             =   585
         Width           =   1350
      End
      Begin ListBox FileList 
         Height          =   1590
         Left            =   60
         TabIndex        =   8
         Top             =   960
         Width           =   1335
      End
      Begin CommandButton Cmd_Añadir 
         Caption         =   "&Añadir"
         Height          =   345
         Left            =   1455
         TabIndex        =   7
         Top             =   1785
         Width           =   1095
      End
      Begin CommandButton Cmd_Borrar 
         Caption         =   "&Borrar"
         Height          =   345
         Left            =   1455
         TabIndex        =   6
         Top             =   2190
         Width           =   1095
      End
      Begin Frame Frame_Archivos 
         BackColor       =   &H00C0C0C0&
         Caption         =   "&Acción"
         ClipControls    =   0   'False
         Height          =   1245
         Left            =   1470
         TabIndex        =   2
         Top             =   465
         Width           =   1095
         Begin OptionButton Opcion_Enviar 
            BackColor       =   &H00C0C0C0&
            Caption         =   "&Enviar"
            Height          =   270
            Left            =   90
            TabIndex        =   5
            Top             =   330
            Value           =   -1  'True
            Width           =   915
         End
         Begin OptionButton Opcion_Borrar 
            BackColor       =   &H00C0C0C0&
            Caption         =   "&Borrar"
            Height          =   240
            Left            =   90
            TabIndex        =   4
            Top             =   840
            Width           =   855
         End
         Begin OptionButton Opcion_Recibir 
            BackColor       =   &H00C0C0C0&
            Caption         =   "&Recibir"
            Height          =   240
            Left            =   90
            TabIndex        =   3
            Top             =   600
            Width           =   960
         End
      End
      Begin Label Archivos_Label 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Nombre :"
         Height          =   240
         Left            =   75
         TabIndex        =   0
         Top             =   330
         Width           =   1035
      End
   End
End
Declare Function GetPrivateProfileString Lib "Kernel" (ByVal lpApplicationName As String, lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Integer, ByVal lpFilename As String) As Integer
Declare Function WritePrivateProfileString Lib "Kernel" (ByVal lpApplicationName As String, lpKeyName As Any, lpString As Any, ByVal lplFileName As String) As Integer



Dim cArchivos_Enviar As String
Dim cArchivos_Recibir As String
Dim cArchivos_Borrar As String

Dim Archivos_Enviar As String
Dim Archivos_Recibir As String
Dim Archivos_Borrar As String

Sub Cmd_Añadir_Click ()
    Dim cArchivos$
    If Len(FileText.Text) > 0 Then
       FileList.AddItem (FileText.Text)
       FileText.Text = ""
       cArchivos = FileList.List(0)
       For I = 1 To FileList.ListCount - 1
            cArchivos = cArchivos + "," + FileList.List(I)
       Next
       
       If Opcion_Enviar.Value = True Then
          Archivos_Enviar = cArchivos
       End If
       If Opcion_Recibir.Value = True Then
          Archivos_Recibir = cArchivos
       End If
       If Opcion_Borrar.Value = True Then
          Archivos_Borrar = cArchivos
       End If
    End If

End Sub

Sub Cmd_Borrar_Click ()
    Dim cArchivos$
    If FileList.ListCount > 0 And FileList.ListIndex <> -1 Then
        FileList.RemoveItem (FileList.ListIndex)
       cArchivos = FileList.List(0)
       For I = 1 To FileList.ListCount - 1
            cArchivos = cArchivos + "," + FileList.List(I)
       Next
       
       If Opcion_Enviar.Value = True Then
          Archivos_Enviar = cArchivos
       End If
       If Opcion_Recibir.Value = True Then
          Archivos_Recibir = cArchivos
       End If
       If Opcion_Borrar.Value = True Then
          Archivos_Borrar = cArchivos
       End If
    
    End If
End Sub

Sub Cmd_Cancelar_Click ()
    Unload Me
End Sub

Sub Cmd_Grabar_Click ()
Dim cBuffer$

cBuffer = Directorio_Enviar.Text
nCnt = WritePrivateProfileString(ByVal cInstancia, ByVal "Directorio_Enviar", ByVal cBuffer, ByVal "TwoWay.INI")

cBuffer = Directorio_Recibir.Text
nCnt = WritePrivateProfileString(ByVal cInstancia, ByVal "Directorio_Recibir", ByVal cBuffer, ByVal "TwoWay.INI")

cBuffer = Horas_Recibir.Text
nCnt = WritePrivateProfileString(ByVal cInstancia, ByVal "Horas_Recibir", ByVal cBuffer, ByVal "TwoWay.INI")

cBuffer = Horas_Enviar.Text
nCnt = WritePrivateProfileString(ByVal cInstancia, ByVal "Horas_Enviar", ByVal cBuffer, ByVal "TwoWay.INI")

cBuffer = IIf(Modo_Automático.Value = True, "Automatico", "Manual")
nCnt = WritePrivateProfileString(ByVal cInstancia, ByVal "Modo", ByVal cBuffer, ByVal "TwoWay.INI")

cBuffer = IIf(Velocidad_19200.Value = True, "19200", "9600")
nCnt = WritePrivateProfileString(ByVal cInstancia, ByVal "Velocidad", ByVal cBuffer, ByVal "TwoWay.INI")

cBuffer = Str(Comms.ListIndex + 1)
nCnt = WritePrivateProfileString(ByVal cInstancia, ByVal "Puerto", ByVal cBuffer, ByVal "TwoWay.INI")

cBuffer = Archivos_Enviar
nCnt = WritePrivateProfileString(ByVal cInstancia, ByVal "Archivos_Enviar", ByVal cBuffer, ByVal "TwoWay.INI")

cBuffer = Archivos_Recibir
nCnt = WritePrivateProfileString(ByVal cInstancia, ByVal "Archivos_Recibir", ByVal cBuffer, ByVal "TwoWay.INI")

cBuffer = Archivos_Borrar
nCnt = WritePrivateProfileString(ByVal cInstancia, ByVal "Archivos_Borrar", ByVal cBuffer, ByVal "TwoWay.INI")

lModificado = True
End Sub

Sub FileList_Click ()
    FileText.Text = FileList.List(FileList.ListIndex)
End Sub

Sub Form_Load ()


Comms.AddItem ("Comm1")
Comms.AddItem ("Comm2")
Comms.AddItem ("Comm3")
Comms.AddItem ("Comm4")
Comms.ListIndex = Val(Lee_Profile(cInstancia, "Puerto", "")) - 1

Directorio_Enviar.Text = Lee_Profile(cInstancia, "Directorio_Enviar", "")
Directorio_Recibir.Text = Lee_Profile(cInstancia, "Directorio_Recibir", "")

Horas_Enviar.Text = Lee_Profile(cInstancia, "Horas_Enviar", "")
Horas_Recibir.Text = Lee_Profile(cInstancia, "Horas_Recibir", "")

Modo_Automático.Value = IIf(Lee_Profile(cInstancia, "Modo", "") = "Automatico", True, False)
Modo_Manual.Value = IIf(Lee_Profile(cInstancia, "Modo", "") = "Automatico", False, True)

Velocidad_19200.Value = IIf(Lee_Profile(cInstancia, "Velocidad", "") = "19200", True, False)
Velocidad_9600.Value = IIf(Lee_Profile(cInstancia, "Velocidad", "") = "19200", False, True)

cArchivos_Enviar = Lee_Profile(cInstancia, "Archivos_Enviar", "")
cArchivos_Recibir = Lee_Profile(cInstancia, "Archivos_Recibir", "")
cArchivos_Borrar = Lee_Profile(cInstancia, "Archivos_Borrar", "")

Archivos_Enviar = cArchivos_Enviar
Archivos_Recibir = cArchivos_Recibir
Archivos_Borrar = cArchivos_Borrar

Call Opcion_Enviar_Click
End Sub

Function Lee_Profile (Aplicacion$, Tema$, Default$) As String
Dim cBuffer$
    cBuffer = Space(40)
    nCnt = GetPrivateProfileString(ByVal Aplicacion, ByVal Tema, Default, cBuffer, Len(cBuffer), "TwoWay.INI")

    Lee_Profile = Mid(cBuffer, 1, nCnt)
End Function

Function Lee_Submit (cOperacion$) As String
Dim nFound%, cFileSubmit$, cNombre$, cExt$

Select Case cOperacion
       Case "Enviar"
            If Len(cArchivos_Enviar) > 0 Then
               If InStr(1, cArchivos_Enviar, ",") > 0 Then
                  cFileSubmit = Left(cArchivos_Enviar, InStr(1, cArchivos_Enviar, ",") - 1)
                  cArchivos_Enviar = Right(cArchivos_Enviar, Len(cArchivos_Enviar) - InStr(1, cArchivos_Enviar, ","))
               Else
                  cFileSubmit = cArchivos_Enviar
                  cArchivos_Enviar = ""
               End If
            Else
                cFileSubmit = ""
            End If

       Case "Recibir"
            If Len(cArchivos_Recibir) > 0 Then
               If InStr(1, cArchivos_Recibir, ",") > 0 Then
                  cFileSubmit = Left(cArchivos_Recibir, InStr(1, cArchivos_Recibir, ",") - 1)
                  cArchivos_Recibir = Right(cArchivos_Recibir, Len(cArchivos_Recibir) - InStr(1, cArchivos_Recibir, ","))
               Else
                  cFileSubmit = cArchivos_Recibir
                  cArchivos_Recibir = ""
               End If
            Else
                cFileSubmit = ""
            End If

       Case "Borrar"
            If Len(cArchivos_Borrar) > 0 Then
               If InStr(1, cArchivos_Borrar, ",") > 0 Then
                  cFileSubmit = Left(cArchivos_Borrar, InStr(1, cArchivos_Borrar, ",") - 1)
                  cArchivos_Borrar = Right(cArchivos_Borrar, Len(cArchivos_Borrar) - InStr(1, cArchivos_Borrar, ","))
               Else
                  cFileSubmit = cArchivos_Borrar
                  cArchivos_Borrar = ""
               End If
            Else
                cFileSubmit = ""
            End If

End Select
'Ahora Descomponer Archivo en Nombre y Extensión
If Len(cFileSubmit) > 0 Then
    If InStr(1, cFileSubmit, ".") > 0 Then
        cNombre = Left(cFileSubmit, InStr(1, cFileSubmit, ".") - 1)
        cExt = Right(cFileSubmit, Len(cFileSubmit) - InStr(1, cFileSubmit, "."))
    Else
        cNombre = cFileSubmit
        cExt = ""
    End If
    cFileSubmit = RTrim(cNombre) + "." + cExt
End If

Lee_Submit = cFileSubmit

End Function

Sub Modo_Automático_Click ()
    Horas_Enviar.Enabled = True
    Horas_Recibir.Enabled = True
End Sub

Sub Modo_Manual_Click ()
    Horas_Enviar.Enabled = False
    Horas_Recibir.Enabled = False
End Sub

Sub Opcion_Borrar_Click ()
    cArchivos_Borrar = Archivos_Borrar
    FileList.Clear
    cFileSubmit = Lee_Submit("Borrar")
    Do While cFileSubmit <> ""
        FileList.AddItem (cFileSubmit)
        cFileSubmit = Lee_Submit("Borrar")
    Loop

End Sub

Sub Opcion_Enviar_Click ()
    cArchivos_Enviar = Archivos_Enviar
    FileList.Clear
    cFileSubmit = Lee_Submit("Enviar")
    Do While cFileSubmit <> ""
        FileList.AddItem (cFileSubmit)
        cFileSubmit = Lee_Submit("Enviar")
    Loop
End Sub

Sub Opcion_Recibir_Click ()
    cArchivos_Recibir = Archivos_Recibir
    FileList.Clear
    cFileSubmit = Lee_Submit("Recibir")
    Do While cFileSubmit <> ""
        FileList.AddItem (cFileSubmit)
        cFileSubmit = Lee_Submit("Recibir")
    Loop

End Sub


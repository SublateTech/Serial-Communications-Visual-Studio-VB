VERSION 2.00
Begin Form Setup 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Configuración"
   ClientHeight    =   4305
   ClientLeft      =   1785
   ClientTop       =   2025
   ClientWidth     =   5565
   Height          =   4995
   Left            =   1725
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4305
   ScaleWidth      =   5565
   Top             =   1395
   Width           =   5685
   Begin CommonDialog CMD 
      Left            =   5505
      Top             =   4290
   End
   Begin Frame Directorios 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Directorios"
      Height          =   1035
      Left            =   75
      TabIndex        =   15
      Top             =   30
      Width           =   5400
      Begin TextBox Directorio_Recibir 
         Height          =   285
         Left            =   1140
         TabIndex        =   19
         Top             =   615
         Width           =   4140
      End
      Begin TextBox Directorio_Enviar 
         Height          =   285
         Left            =   1140
         TabIndex        =   18
         Top             =   255
         Width           =   4140
      End
      Begin Label Directorio_Recibir_Label 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Recibir:"
         Height          =   255
         Left            =   105
         TabIndex        =   17
         Top             =   660
         Width           =   915
      End
      Begin Label Directorio_Enviar_Label 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Envíar:"
         Height          =   255
         Left            =   105
         TabIndex        =   16
         Top             =   270
         Width           =   1035
      End
   End
   Begin Frame Puerto_Velocidad_Frame 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Puerto/Velocidad"
      Height          =   735
      Left            =   2805
      TabIndex        =   10
      Top             =   1140
      Width           =   2670
      Begin ComboBox Comms 
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   180
         TabIndex        =   21
         Top             =   285
         Width           =   1110
      End
      Begin OptionButton Velocidad_9600 
         BackColor       =   &H00C0C0C0&
         Caption         =   "  9600"
         Height          =   240
         Left            =   1695
         TabIndex        =   13
         Top             =   465
         Width           =   900
      End
      Begin OptionButton Velocidad_19200 
         BackColor       =   &H00C0C0C0&
         Caption         =   "19200 "
         Height          =   225
         Left            =   1695
         TabIndex        =   12
         Top             =   195
         Value           =   -1  'True
         Width           =   825
      End
   End
   Begin Frame Formato_Frame 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Formato Bloque de Datos"
      Height          =   885
      Left            =   2805
      TabIndex        =   26
      Top             =   1935
      Width           =   2655
      Begin CheckBox QuitarCRLF 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Quitar CRLF"
         Height          =   225
         Left            =   195
         TabIndex        =   29
         Top             =   555
         Value           =   1  'Checked
         Width           =   1425
      End
      Begin OptionButton ASCII 
         BackColor       =   &H00C0C0C0&
         Caption         =   "ASCII"
         Height          =   255
         Left            =   180
         TabIndex        =   28
         Top             =   255
         Value           =   -1  'True
         Width           =   810
      End
      Begin OptionButton TRANSPARENT 
         BackColor       =   &H00C0C0C0&
         Caption         =   "&Transparente"
         Height          =   255
         Left            =   1125
         TabIndex        =   27
         Top             =   255
         Width           =   1440
      End
   End
   Begin Frame Modo_Horas_Frame 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Modo/Horas"
      Height          =   1365
      Left            =   2805
      TabIndex        =   11
      Top             =   2865
      Width           =   2670
      Begin TextBox Horas_Recibir 
         Height          =   285
         Left            =   1380
         TabIndex        =   23
         Top             =   990
         Width           =   1140
      End
      Begin TextBox Horas_Enviar 
         Height          =   285
         Left            =   1380
         TabIndex        =   22
         Top             =   630
         Width           =   1140
      End
      Begin OptionButton Modo_Manual 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Manual"
         Height          =   240
         Left            =   1440
         TabIndex        =   14
         Top             =   270
         Width           =   1080
      End
      Begin OptionButton Modo_Automático 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Automático"
         Height          =   240
         Left            =   75
         TabIndex        =   20
         Top             =   270
         Value           =   -1  'True
         Width           =   1350
      End
      Begin Label Horas_Recibir_Label 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Horas Recibir:"
         Height          =   240
         Left            =   105
         TabIndex        =   25
         Top             =   990
         Width           =   1290
      End
      Begin Label Horas_Enviar_Label 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Horas Envíar:"
         Height          =   285
         Left            =   105
         TabIndex        =   24
         Top             =   705
         Width           =   1260
      End
   End
   Begin Frame Archivos_Frame 
      BackColor       =   &H00C0C0C0&
      Caption         =   " Archivos/Inicio Manual"
      Height          =   3075
      Left            =   90
      TabIndex        =   1
      Top             =   1140
      Width           =   2670
      Begin CheckBox Archivo_LOG 
         BackColor       =   &H00C0C0C0&
         Caption         =   "&Archivo LOG"
         Height          =   330
         Left            =   75
         TabIndex        =   30
         Top             =   2625
         Width           =   1815
      End
      Begin TextBox FileText 
         Height          =   330
         Left            =   60
         TabIndex        =   9
         Top             =   585
         Width           =   1335
      End
      Begin ListBox FileList 
         Height          =   1590
         Left            =   60
         TabIndex        =   8
         Top             =   900
         Width           =   1335
      End
      Begin CommandButton Cmd_Añadir 
         Caption         =   "&Añadir"
         Height          =   345
         Left            =   1470
         TabIndex        =   7
         Top             =   1755
         Width           =   1095
      End
      Begin CommandButton Cmd_Borrar 
         Caption         =   "&Borrar"
         Height          =   345
         Left            =   1470
         TabIndex        =   6
         Top             =   2100
         Width           =   1095
      End
      Begin Frame Frame_Archivos 
         BackColor       =   &H00C0C0C0&
         Caption         =   "&Acción"
         ClipControls    =   0   'False
         Height          =   1245
         Left            =   1470
         TabIndex        =   2
         Top             =   480
         Width           =   1095
         Begin OptionButton Opcion_Enviar 
            BackColor       =   &H00C0C0C0&
            Caption         =   "&Enviar"
            Height          =   270
            Left            =   75
            TabIndex        =   5
            Top             =   330
            Value           =   -1  'True
            Width           =   915
         End
         Begin OptionButton Opcion_Borrar 
            BackColor       =   &H00C0C0C0&
            Caption         =   "&Borrar"
            Height          =   240
            Left            =   75
            TabIndex        =   4
            Top             =   885
            Width           =   915
         End
         Begin OptionButton Opcion_Recibir 
            BackColor       =   &H00C0C0C0&
            Caption         =   "&Recibir"
            Height          =   240
            Left            =   75
            TabIndex        =   3
            Top             =   615
            Width           =   915
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
   Begin Menu Menu_Profile 
      Caption         =   "&Profile"
      Begin Menu Op_Nuevo 
         Caption         =   "&Nuevo"
      End
      Begin Menu Op_Abrir 
         Caption         =   "&Abrir..."
      End
      Begin Menu Op_Guardar 
         Caption         =   "&Guardar "
      End
      Begin Menu Op_Guardar_Como 
         Caption         =   "&Guardar como..."
      End
      Begin Menu Linea1 
         Caption         =   "-"
      End
      Begin Menu Op_Salir 
         Caption         =   "&Salir"
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
Dim Formato As String
Dim Inicio As String
Dim Profile As String

Sub Abre_Profile ()
Me.Caption = "Configuración - " + UCase(Quita_Path(Profile))

Directorio_Enviar.Text = Lee_Profile(cInstancia, "DirectorioEnviar", "")
Directorio_Recibir.Text = Lee_Profile(cInstancia, "DirectorioRecibir", "")
Directorio_Enviar.Text = Directorio_Enviar.Text + IIf(Right(Directorio_Enviar.Text, 1) = "\", "", "\")
Directorio_Recibir.Text = Directorio_Recibir.Text + IIf(Right(Directorio_Recibir.Text, 1) = "\", "", "\")
cArchivos_Enviar = Lee_Profile(cInstancia, "ArchivosEnviar", "")
cArchivos_Recibir = Lee_Profile(cInstancia, "ArchivosRecibir", "")
cArchivos_Borrar = Lee_Profile(cInstancia, "ArchivosBorrar", "")

Archivos_Enviar = cArchivos_Enviar
Archivos_Recibir = cArchivos_Recibir
Archivos_Borrar = cArchivos_Borrar

If Profile = cDirectorio_Programa + "TWOWAY.INI" Then
    Comms.Clear
    Comms.AddItem ("Comm1")
    Comms.AddItem ("Comm2")
    Comms.AddItem ("Comm3")
    Comms.AddItem ("Comm4")
    Comms.ListIndex = Val(Lee_Profile(cInstancia, "Puerto", "1")) - 1

    Horas_Enviar.Text = Lee_Profile(cInstancia, "HorasEnviar", "")
    Horas_Recibir.Text = Lee_Profile(cInstancia, "HorasRecibir", "")

    Velocidad_19200.Value = IIf(Lee_Profile(cInstancia, "Velocidad", "") = "19200", True, False)
    Velocidad_9600.Value = IIf(Lee_Profile(cInstancia, "Velocidad", "") = "19200", False, True)

    Formato = Lee_Profile(cInstancia, "Formato", "ASCII")
    If Formato = "ASCII" Then
	ASCII.Value = True
    Else
	TRANSPARENT.Value = True
    End If

    Modo_Automático.Value = IIf(Lee_Profile(cInstancia, "Modo", "") = "Automatico", True, False)
    Modo_Manual.Value = IIf(Lee_Profile(cInstancia, "Modo", "") = "Automatico", False, True)
    Inicio = Lee_Profile(cInstancia, "InicioManual", "Enviar")
    If Inicio = "Enviar" Then
	Call Opcion_Enviar_Click
    ElseIf Inicio = "Recibir" Then
	Opcion_Recibir.Value = True
	Call Opcion_Recibir_Click
    Else
	Opcion_Borrar.Value = True
	Call Opcion_Borrar_Click
    End If

    QuitarCRLF.Value = IIf(Lee_Profile(cInstancia, "QuitarCRLF", "SI") = "SI", 1, 0)
    Archivo_LOG.Value = IIf(Lee_Profile(cInstancia, "Archivo_LOG", "SI") = "SI", 1, 0)
    
    Archivo_LOG.Enabled = True
    Formato_Frame.Enabled = True
    Puerto_Velocidad_Frame.Enabled = True
    Modo_Horas_Frame.Enabled = True

Else
   Call Opcion_Enviar_Click
   Archivo_LOG.Enabled = False
   Formato_Frame.Enabled = False
   Puerto_Velocidad_Frame.Enabled = False
   Modo_Horas_Frame.Enabled = False
End If
    
End Sub

Sub Abrir_Guardar_Profile (nOpcion)
Dim cFile$, nFound%
    CMD.Filter = "Archivos INI (*.ini)|*.ini"
    CMD.FilterIndex = 1
    CMD.DefaultExt = "INI"
    CMD.InitDir = Left(cDirectorio_Programa, Len(cDirectorio_Programa) - 1)
    CMD.Filename = "*.ini"
    CMD.Flags = OFN_OVERWRITEPROMPT
    CMD.Action = nOpcion
End Sub

Sub ASCII_Click ()
    Formato = "ASCII"
End Sub

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

Sub Directorio_Enviar_LostFocus ()
    If Dir$(Directorio_Enviar.Text, 16) = "" Then
	MsgBox "No Existe Directorio"
	Directorio_Enviar.SetFocus
    End If
End Sub

Sub Directorio_Recibir_LostFocus ()
    If Dir$(Directorio_Recibir.Text, 16) = "" Then
	MsgBox "No Existe Directorio"
	Directorio_Recibir.SetFocus
    End If
End Sub

Sub FileList_Click ()
    FileText.Text = FileList.List(FileList.ListIndex)
End Sub

Sub Form_Load ()

Profile = cDirectorio_Programa + cProfile
Abre_Profile

End Sub

Sub Form_Unload (Cancel As Integer)
    Profile = "TWOWAY.INI"
End Sub

Sub Guardar_Archivo ()
Dim cBuffer$
Screen.MousePointer = 11
If Profile = cDirectorio_Programa + "TWOWAY.INI" Then
    cBuffer = Str(Comms.ListIndex + 1)
    nCnt = WritePrivateProfileString(ByVal cInstancia, ByVal "Puerto", ByVal cBuffer, ByVal Profile)

    cBuffer = IIf(Velocidad_19200.Value = True, "19200", "9600")
    nCnt = WritePrivateProfileString(ByVal cInstancia, ByVal "Velocidad", ByVal cBuffer, ByVal Profile)

    cBuffer = Formato
    nCnt = WritePrivateProfileString(ByVal cInstancia, ByVal "Formato", ByVal cBuffer, ByVal Profile)

    cBuffer = IIf(QuitarCRLF.Value = 1, "SI", "NO")
    nCnt = WritePrivateProfileString(ByVal cInstancia, ByVal "QuitarCRLF", ByVal cBuffer, ByVal Profile)

    cBuffer = IIf(Modo_Automático.Value = True, "Automatico", "Manual")
    nCnt = WritePrivateProfileString(ByVal cInstancia, ByVal "Modo", ByVal cBuffer, ByVal Profile)

    cBuffer = IIf(Opcion_Enviar.Value = True, "Enviar", IIf(Opcion_Recibir.Value = True, "Recibir", "Borrar"))
    nCnt = WritePrivateProfileString(ByVal cInstancia, ByVal "InicioManual", ByVal cBuffer, ByVal Profile)

    cBuffer = Horas_Recibir.Text
    nCnt = WritePrivateProfileString(ByVal cInstancia, ByVal "HorasRecibir", ByVal cBuffer, ByVal Profile)

    cBuffer = Horas_Enviar.Text
    nCnt = WritePrivateProfileString(ByVal cInstancia, ByVal "HorasEnviar", ByVal cBuffer, ByVal Profile)
    
    cBuffer = IIf(Archivo_LOG.Value = 0, "NO", "SI")
    nCnt = WritePrivateProfileString(ByVal cInstancia, ByVal "Archivo_LOG", ByVal cBuffer, ByVal Profile)
    lModificado = True

End If
cBuffer = Directorio_Enviar.Text
nCnt = WritePrivateProfileString(ByVal cInstancia, ByVal "DirectorioEnviar", ByVal cBuffer, ByVal Profile)

cBuffer = Directorio_Recibir.Text
nCnt = WritePrivateProfileString(ByVal cInstancia, ByVal "DirectorioRecibir", ByVal cBuffer, ByVal Profile)

cBuffer = Archivos_Enviar
nCnt = WritePrivateProfileString(ByVal cInstancia, ByVal "ArchivosEnviar", ByVal cBuffer, ByVal Profile)

cBuffer = Archivos_Recibir
nCnt = WritePrivateProfileString(ByVal cInstancia, ByVal "ArchivosRecibir", ByVal cBuffer, ByVal Profile)

cBuffer = Archivos_Borrar
nCnt = WritePrivateProfileString(ByVal cInstancia, ByVal "ArchivosBorrar", ByVal cBuffer, ByVal Profile)
Screen.MousePointer = 0
End Sub

Function Lee_Profile (Aplicacion$, Tema$, Default$) As String
Dim cBuffer$
    cBuffer = Space(512)
    nCnt = GetPrivateProfileString(ByVal Aplicacion, ByVal Tema, Default, cBuffer, Len(cBuffer), Profile)

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

Sub Op_Abrir_Click ()
    Abrir_Guardar_Profile (1)

    If Len(Dir$(CMD.Filename)) > 0 And CMD.Filename <> "*.ini" Then ' And Err <> 32755 Then
       Profile = CMD.Filename
       Abre_Profile
    End If

End Sub

Sub Op_Guardar_Click ()
Dim cBuffer$
If Profile = "[Sin Nombre]" Then
   Abrir_Guardar_Profile (2)
   Profile = CMD.Filename
   If Err = 32755 Then Exit Sub
   If CMD.Filename = "*.ini" Then Exit Sub
End If
Guardar_Archivo
Abre_Profile
End Sub

Sub Op_Guardar_Como_Click ()
    Abrir_Guardar_Profile (2)
	
    If Err = 32755 Then Exit Sub
    If CMD.Filename <> "*.ini" Then ' And Err <> 32755 Then
       Profile = CMD.Filename
       Guardar_Archivo
       Abre_Profile
    Else
       Exit Sub
    End If
    
End Sub

Sub Op_Nuevo_Click ()
    Profile = "[Sin Nombre]"
    Abre_Profile
End Sub

Sub Op_Salir_Click ()
    Unload Me
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

Function Quita_Path (ByVal cFullPath) As String
    Dim nFound%, cFile$
    cFile = cFullPath
    Do
      nFound = InStr(1, cFile, "\")
      cFile = Right(cFile, Len(cFile) - nFound)
    Loop While nFound > 0
Quita_Path = cFile
End Function

Sub TRANSPARENT_Click ()
    Formato = "Transparente"
End Sub


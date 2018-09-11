VERSION 2.00
Begin Form TwoWay 
   BackColor       =   &H00800000&
   BorderStyle     =   0  'None
   Caption         =   "TwoWay"
   ClientHeight    =   1920
   ClientLeft      =   2670
   ClientTop       =   3540
   ClientWidth     =   2955
   DrawMode        =   1  'Blackness
   FillColor       =   &H00C0C0C0&
   FillStyle       =   0  'Solid
   FontBold        =   -1  'True
   FontItalic      =   -1  'True
   FontName        =   "Arial"
   FontSize        =   8.25
   FontStrikethru  =   -1  'True
   FontUnderline   =   0   'False
   Height          =   2325
   Icon            =   TWOWAY.FRX:0000
   Left            =   2610
   LinkMode        =   1  'Source
   LinkTopic       =   "Comandos"
   MaxButton       =   0   'False
   Picture         =   TWOWAY.FRX:0302
   ScaleHeight     =   1920
   ScaleWidth      =   2955
   Top             =   3195
   Width           =   3075
   WindowState     =   1  'Minimized
   Begin MSComm Puerto 
      InBufferSize    =   256
      Interval        =   1000
      Left            =   2745
      Top             =   1740
   End
   Begin PictureBox RTS 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   165
      Left            =   1395
      Picture         =   TWOWAY.FRX:163C
      ScaleHeight     =   165
      ScaleWidth      =   165
      TabIndex        =   6
      Top             =   1650
      Width           =   165
   End
   Begin PictureBox DSR 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   165
      Left            =   2115
      Picture         =   TWOWAY.FRX:170E
      ScaleHeight     =   165
      ScaleWidth      =   165
      TabIndex        =   5
      Top             =   1650
      Width           =   165
   End
   Begin PictureBox Receive 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   165
      Left            =   690
      Picture         =   TWOWAY.FRX:17E0
      ScaleHeight     =   165
      ScaleWidth      =   165
      TabIndex        =   4
      Top             =   1650
      Width           =   165
   End
   Begin PictureBox Send 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   165
      Left            =   1035
      Picture         =   TWOWAY.FRX:18B2
      ScaleHeight     =   165
      ScaleWidth      =   165
      TabIndex        =   3
      Top             =   1650
      Width           =   165
   End
   Begin PictureBox CTS 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   165
      Left            =   2475
      Picture         =   TWOWAY.FRX:1984
      ScaleHeight     =   165
      ScaleWidth      =   165
      TabIndex        =   2
      Top             =   1650
      Width           =   165
   End
   Begin PictureBox Port 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   165
      Left            =   315
      Picture         =   TWOWAY.FRX:1A56
      ScaleHeight     =   165
      ScaleWidth      =   165
      TabIndex        =   1
      Top             =   1650
      Width           =   165
   End
   Begin PictureBox DTR 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   165
      Left            =   1755
      Picture         =   TWOWAY.FRX:1B28
      ScaleHeight     =   165
      ScaleWidth      =   165
      TabIndex        =   0
      Top             =   1650
      Width           =   165
   End
   Begin PictureBox Conn_Off_Pic 
      BackColor       =   &H00800000&
      BorderStyle     =   0  'None
      Height          =   15
      Left            =   450
      Picture         =   TWOWAY.FRX:1BFA
      ScaleHeight     =   1
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   1
      TabIndex        =   9
      Top             =   1170
      Width           =   15
      Begin PictureBox On_Pic 
         AutoSize        =   -1  'True
         BackColor       =   &H00800000&
         BorderStyle     =   0  'None
         Height          =   165
         Left            =   1500
         Picture         =   TWOWAY.FRX:1EFC
         ScaleHeight     =   11
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   11
         TabIndex        =   14
         Top             =   345
         Width           =   165
      End
      Begin PictureBox Off_Pic 
         AutoSize        =   -1  'True
         BackColor       =   &H00800000&
         BorderStyle     =   0  'None
         Height          =   165
         Left            =   1500
         Picture         =   TWOWAY.FRX:1FCE
         ScaleHeight     =   11
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   11
         TabIndex        =   13
         Top             =   75
         Width           =   165
      End
      Begin PictureBox Conn_Snd_Pic 
         AutoSize        =   -1  'True
         BackColor       =   &H00800000&
         BorderStyle     =   0  'None
         Height          =   480
         Left            =   975
         Picture         =   TWOWAY.FRX:20A0
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   12
         Top             =   0
         Width           =   480
      End
      Begin PictureBox Conn_Rcv_Pic 
         AutoSize        =   -1  'True
         BackColor       =   &H00800000&
         BorderStyle     =   0  'None
         Height          =   480
         Left            =   495
         Picture         =   TWOWAY.FRX:23A2
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   11
         Top             =   0
         Width           =   480
      End
      Begin PictureBox Conn_On_Pic 
         AutoSize        =   -1  'True
         BackColor       =   &H00800000&
         BorderStyle     =   0  'None
         Height          =   480
         Left            =   15
         Picture         =   TWOWAY.FRX:26A4
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   10
         Top             =   0
         Width           =   480
      End
   End
   Begin Label Msg 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Esperando señal..."
      FontBold        =   -1  'True
      FontItalic      =   -1  'True
      FontName        =   "Times New Roman"
      FontSize        =   9
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H0000FFFF&
      Height          =   315
      Left            =   -15
      TabIndex        =   8
      Top             =   870
      Width           =   2955
   End
   Begin Label Titulos 
      BackColor       =   &H0000FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   " Port  Rx   Tx   RTS DTR DSR CTS"
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "Times New Roman"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H00FFFF00&
      Height          =   210
      Left            =   210
      TabIndex        =   7
      Top             =   1380
      Width           =   2610
   End
   Begin Menu Menu_Opciones 
      Caption         =   "&Opciones"
      Visible         =   0   'False
      Begin Menu Op_Enviar 
         Caption         =   "&Envíar"
      End
      Begin Menu Op_Recibir 
         Caption         =   "&Recibir"
      End
      Begin Menu Op_Borrar 
         Caption         =   "&Borrar"
      End
      Begin Menu Linea 
         Caption         =   "-"
      End
      Begin Menu Op_Automático 
         Caption         =   "&Automático"
         Checked         =   -1  'True
      End
      Begin Menu Linea1 
         Caption         =   "-"
      End
      Begin Menu Op_Configuracion 
         Caption         =   "&Configuración"
      End
      Begin Menu Linea2 
         Caption         =   "-"
      End
      Begin Menu Op_Salir 
         Caption         =   "&Salir"
      End
   End
   Begin Menu Menu_Acerca 
      Caption         =   "&Acerca de "
      Visible         =   0   'False
   End
End
Declare Function lclose Lib "Kernel" Alias "_lclose" (ByVal hFile As Integer) As Integer
Declare Function lcreat Lib "Kernel" Alias "_lcreat" (ByVal lpPathName As String, ByVal iAttribute As Integer) As Integer
Declare Function lopen Lib "Kernel" Alias "_lopen" (ByVal lpPathName As String, ByVal iReadWrite As Integer) As Integer
Declare Function lread Lib "Kernel" Alias "_lread" (ByVal hFile As Integer, ByVal lpBuffer As String, ByVal wBytes As Integer) As Integer
Declare Function lwrite Lib "Kernel" Alias "_lwrite" (ByVal hFile As Integer, ByVal lpBuffer As String, ByVal wBytes As Integer) As Integer

Declare Function GetTickCount Lib "User" () As Long

Declare Function GetPrivateProfileString Lib "Kernel" (ByVal lpApplicationName As String, lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Integer, ByVal lpFilename As String) As Integer
Declare Function WritePrivateProfileString Lib "Kernel" (ByVal lpApplicationName As String, lpKeyName As Any, lpString As Any, ByVal lplFileName As String) As Integer

Declare Function GetModuleHandle Lib "Kernel" (ByVal lpModuleName As String) As Integer
Declare Function GetModuleUsage Lib "Kernel" (ByVal hModule As Integer) As Integer
Declare Function GetModuleFileName Lib "Kernel" (ByVal hModule As Integer, ByVal lpFilename As String, ByVal nSize As Integer) As Integer

Declare Function UngetCommChar Lib "User" (ByVal nCid As Integer, ByVal cChar As Integer) As Integer

Dim TimeOutIni As Long
Dim TimeOut As Long
Dim lCRLF As Integer     'Quitar lCRLF
Dim lLastBlock As Integer
Dim BID As String        'Line Bid (ENQ)

Dim NUL As String         '0
Dim SOH As String         '1
Dim STX As String         '2
Dim ETX As String         '3
Dim EOT As String         '4
Dim ENQ As String         '5        'Enquiry
Dim DLE As String         '10
Dim NAK As String         '15
Dim ETB As String         '17

Dim ACK As String
Dim ACK0 As String                  '= Chr(16) + Chr(48)
Dim ACK1 As String                  '= Chr(16) + Chr(49)
Dim DISC As String                  ' Chr(16) + Chr(4)
Dim WACK As String                  'Chr(16) + Chr(59)
Dim TTD As String                   'Chr(2) + Chr(5)
Dim FATAL_ERROR As String
Dim OK As String

Dim ACKP As String

Dim cComando As String
Dim cBlock As String
Dim cLoginId As String

'Files
Dim LENFILE As Long
Dim FILENUM As Integer
Dim FileName As String
Dim CRLF As String
Dim BUFFERLEN As Integer
Dim cArchivos_Enviar As String
Dim cArchivos_Recibir As String
Dim cArchivos_Borrar As String
Dim cHora_Borrar As String
Dim cHora_Recibir As String
Dim cHora_Enviar As String
Dim cDirectorio_Enviar As String
Dim cDirectorio_Recibir As String
Dim lActivo As Integer
Dim nHora_Ini%, nHora_Fin%, nTime%
Dim cFormato  As String                'ASCII o TRANSPARENT
Dim cModo As String
Dim cInicio As String
Dim cLogFile As String
Dim lLog As Integer
Dim nPuerto As Integer
Dim cTipo As String

Function Abre_Archivo (cArchivo$)
    FileName = cArchivo
    FILENUM = lopen(cDirectorio_Enviar + cArchivo, 0)
    If FILENUM = -1 Then
       Mensaje "No existe Archivo..."
       Abre_Archivo = False
    Else
       LENFILE = FileLen(cDirectorio_Enviar + cArchivo)
       Abre_Archivo = True
    End If
End Function

Sub Activa_Linea ()
 On Error GoTo Activa_Linea_Err
    Puerto.CommPort = nPuerto
    Puerto.Settings = Lee_Profile(cInstancia, "Velocidad", "19200") + ",O," + IIf(UCase(Lee_Profile(cInstancia, "Formato", "ASCII")) = "ASCII", "7", "8") + ",2"
    
    'Set el Buffer de Entrada
    Puerto.OutBufferSize = BUFFERLEN + 3 + 10
    Puerto.InBufferSize = BUFFERLEN * 2     '1024
    Puerto.Interval = 0
    Puerto.NullDiscard = True
    Puerto.ParityReplace = "?"
    Puerto.Handshaking = MSCOMM_HANDSHAKE_RTS

    If Not Puerto.PortOpen Then
        Puerto.PortOpen = True
        Port.Picture = On_Pic.Picture
    End If

    
     'Ok Line
    Puerto.RTSEnable = True
    RTS.Picture = On_Pic.Picture

    'Listo para Recibir Data
    Puerto.DTREnable = True
    DTR.Picture = On_Pic.Picture

    'Para capturar se±ales de env_o y recepci_n
    Puerto.RThreshold = 1
    Puerto.SThreshold = 1
   

    If Puerto.CTSHolding = True Then
        CTS.Picture = On_Pic.Picture
    Else
        CTS.Picture = Off_Pic.Picture
    End If


    If Puerto.DSRHolding = True Then
        DSR.Picture = On_Pic.Picture
    Else
        DSR.Picture = Off_Pic.Picture
    End If

GoTo Activa_Linea_End
Activa_Linea_Err:
  If Err = 68 Then
     MsgBox "Puerto COMM" + Lee_Profile(cInstancia, "Puerto", "1") + " no Disponible"
     Setup.Show MODAL
  Else
     MsgBox Str(Err)
     End
  End If
  Resume Activa_Linea_End

Activa_Linea_End:

End Sub

Sub Adiciona_Registro (ByVal cRegistro$)
 nCount = lwrite(FILENUM, cRegistro, Len(cRegistro))
End Sub

Function CalBCC (cBuffer$, cControl$)
Dim nBCC%, nCnt%, nChar%, cCadena$
nBCC = 0
nCnt = 1
nChar = 0
cCadena = ""
For nCnt = 1 To Len(cBuffer)
   nChar = Asc(Mid(cBuffer, nCnt, 1))
   If cFormato = "ASCII" Then
    If nChar > 127 Then
      Select Case nChar
            Case 130
                nChar = Asc("e")
            Case 160
                nChar = Asc("a")
            Case 161
                nChar = Asc("i")
            Case 162
                nChar = Asc("o")
            Case 163
                nChar = Asc("u")
            Case 164
                nChar = Asc("n")
            Case 165
                nChar = Asc("N")
            Case Else
                nChar = Asc(".")
       End Select
    End If
   End If
   cCadena = cCadena + Chr(nChar)
   nBCC = nBCC Xor nChar
Next
cBuffer = cCadena
nBCC = nBCC Xor Asc(Right(cControl, 1))
If cFormato = "ASCII" Then
   CalBCC = Chr(nBCC And 255)
Else
   CalBCC = Chr(nBCC)
End If
End Function

Sub Cierra_Archivo ()
    Dim nResult%
    nResult = lclose(FILENUM)
    If nResult <> 0 Then
       Mensaje "No se pudo cerrar Archivo"
    End If
End Sub

Sub Crea_Archivo (cArchivo$)
    FileName = cArchivo
    FILENUM = lcreat(cDirectorio_Recibir + IIf(Right(cDirectorio_Recibir, 1) = "\", "", "\") + FileName, 0)
End Sub

Function Chequea_CTS ()
Dim Estado%
    Estado = Puerto.CTSHolding
    If Estado = False Then
       Mensaje "Error: Linea perdida ..."
    End If
    Chequea_CTS = Estado
End Function

Sub Desactiva_Linea ()

    TwoWay.Icon = Conn_Off_Pic.Picture
    
    'Drop Line
    Puerto.RTSEnable = False
    RTS.Picture = Off_Pic.Picture

    'NO Listo para transmitir
    Puerto.DTREnable = False
    DTR.Picture = Off_Pic.Picture

    If Puerto.CTSHolding = True Then
        CTS.Picture = On_Pic.Picture
    Else
        CTS.Picture = Off_Pic.Picture
    End If

    If Puerto.DSRHolding = True Then
        DSR.Picture = On_Pic.Picture
    Else
        DSR.Picture = Off_Pic.Picture
    End If

    If Puerto.PortOpen Then
        Puerto.PortOpen = False
        Port.Picture = Off_Pic.Picture
    End If
    Wait_Millisecs (200)
End Sub

Function Erase_File (cArchivo)
   cComando = "14"

   cChar = Send_Command(cComando + cArchivo + Space(6) + Space(10), True)
   If cChar = FATAL_ERROR Then
      Erase_File = FATAL_ERROR
      Exit Function
   End If
   If Receive_Status() = "80" Then      ' No existe archivo en Host
      Erase_File = OK
      Exit Function
   End If
Erase_File = FATAL_ERROR
End Function

Function Fetch_File (cFile)

   TwoWay.Icon = Conn_Rcv_Pic.Picture
   cComando = "10"
   cChar = Send_Command(cComando + cFile + Space(6) + Space(10), True)
   If cChar = FATAL_ERROR Then
      Fetch_File = FATAL_ERROR
      Exit Function
   End If
   If Receive_Status() = "81" Then      ' No existe archivo en Host
      Mensaje "  No existe archivo " & cFile
      Fetch_File = OK
      Exit Function
   End If
   cChar = Receive_Command()
   If cChar = FATAL_ERROR Then
        Fetch_File = FATAL_ERROR
        Exit Function
   End If
   LENFILE = Val(Mid(cBlock, 15, 10))
   Crea_Archivo Mid(cBlock, 3, 10)
   cChar = Receive_File()
   cChar = Send_Good_Status()
   Fetch_File = OK
   Exit Function
   
Fetch_File = FATAL_ERROR

End Function

Sub Form_DblClick ()
    If Menu_Opciones.Visible = True Then
       Menu_Opciones.Visible = False
       Menu_Acerca.Visible = False
       Me.Height = Me.Height - 300
    Else
       Menu_Opciones.Visible = True
       Menu_Acerca.Visible = True
       Me.Height = Me.Height + 300
    End If


End Sub

Sub Form_LinkExecute (CmdStr As String, Cancel As Integer)
    Select Case UCase(CmdStr)
            Case "[CERRAR()]"
                Unload Me
    End Select

End Sub

Sub Form_Load ()
Dim cFileSubMit$, cChar$
Me.Show
Do While True
    Set_Vars
    Mensaje "Esperando señal ..."
    Activa_Linea
    If cTipo = "Slave" Then
      If Slave() <> FATAL_ERROR Then
        If Op_Enviar.Checked = True Then
           cFileSubMit = Lee_Submit("Enviar")
           Do While cFileSubMit <> ""
              cChar = Send_File(cFileSubMit)
              If cChar = FATAL_ERROR Then
                 Exit Do
              End If
              cFileSubMit = Lee_Submit("Enviar")
           Loop
        End If
        If Op_Recibir.Checked = True Then
           cFileSubMit = Lee_Submit("Recibir")
           Do While cFileSubMit <> ""
              cChar = Fetch_File(cFileSubMit)
              If cChar = FATAL_ERROR Then
                 Exit Do
              End If
              cFileSubMit = Lee_Submit("Recibir")
           Loop
        End If
        If Op_Borrar.Checked = True Then
           cFileSubMit = Lee_Submit("Borrar")
           Do While cFileSubMit <> ""
              cChar = Erase_File(cFileSubMit)
              If cChar = FATAL_ERROR Then
                 Exit Do
              End If
              cFileSubMit = Lee_Submit("Borrar")
           Loop
        End If
        If cChar = FATAL_ERROR Then
           Graba_Profile Format$(Date, "yy-mm-dd"), cLoginId, "BAD", cDirectorio_Programa + cLogFile
        Else
           Graba_Profile Format$(Date, "yy-mm-dd"), cLoginId, "OK", cDirectorio_Programa + cLogFile
           Graba_Profile Format$(Date, "yy-mm-dd"), "Profile", cProfile, cDirectorio_Programa + cLogFile
           Graba_Profile Format$(Date, "yy-mm-dd"), "ArchivosEnviados", "", cDirectorio_Programa + cLogFile
           Graba_Profile Format$(Date, "yy-mm-dd"), "ArchivosRecibidos", "", cDirectorio_Programa + cLogFile
        End If
        Send_Sign_Off
     End If
    Else
     If Master() <> FATAL_ERROR Then
        'Ok
     End If
    End If
    Desactiva_Linea
Continuar_Form:
Loop
End Sub

Sub Form_Unload (Cancel As Integer)
    Mensaje "Cerrando aplicación ..."
    Desactiva_Linea
    TwoWay.WindowState = 1
    DoEvents
    Wait_Millisecs (1000)
    Unload Me
    End
End Sub

Sub Graba_Profile (ByVal Aplicacion$, ByVal Tema$, ByVal Cadena$, ByVal Archivo$)
  Dim cBuffer$
  If lLog Then
    cBuffer = Cadena
    nCnt = WritePrivateProfileString(ByVal Aplicacion, ByVal Tema, ByVal cBuffer, ByVal Archivo)
  End If
End Sub

Function IfTimeOut () As Integer
If TimeOut > GetTickCount() - TimeOutIni Then
    IfTimeOut = False
Else
    IfTimeOut = True
End If
End Function

Function Lee_Archivo (cBuffer$, nLenArchivo%)
   Dim cBlock$
   cBlock = Space(BUFFERLEN + 3)
   nCount = lread(FILENUM, cBlock, nLenArchivo)
   cBuffer = cBlock
   Lee_Archivo = nCount
End Function

Function Lee_Profile (Aplicacion$, Tema$, Default$) As String
Dim cBuffer$
    cBuffer = Space(512)
    nCnt = GetPrivateProfileString(ByVal Aplicacion, ByVal Tema, Default, cBuffer, Len(cBuffer), cDirectorio_Programa + cProfile)

    Lee_Profile = Mid(cBuffer, 1, nCnt)
End Function

Function Lee_Profile_Privado ()
  If cProfile <> ".INI" Then
    If Dir$(cDirectorio_Programa + cProfile) = "" Then
       Mensaje "Profile " + cProfile + " No existe..."
       Wait_Millisecs (1000)
       Lee_Profile_Privado = False
    Else
       cDirectorio_Enviar = Lee_Profile(cInstancia, "DirectorioEnviar", "ENVIA\")
       cDirectorio_Recibir = Lee_Profile(cInstancia, "DirectorioRecibir", "RECIBE\")
       cArchivos_Enviar = Lee_Profile(cInstancia, "ArchivosEnviar", "")
       cArchivos_Recibir = Lee_Profile(cInstancia, "ArchivosRecibir", "")
       cArchivos_Borrar = Lee_Profile(cInstancia, "ArchivosBorrar", "")
       Lee_Profile_Privado = True
    End If
  Else
    cProfile = "TWOWAY.INI"
    Lee_Profile_Privado = True
  End If
End Function

Function Lee_Submit (cOperacion$) As String
Dim nFound%, cFileSubMit$, cNombre$, cExt$

Select Case cOperacion
       Case "Enviar"
            If Len(cArchivos_Enviar) > 0 Then
               If InStr(1, cArchivos_Enviar, ",") > 0 Then
                  cFileSubMit = Left(cArchivos_Enviar, InStr(1, cArchivos_Enviar, ",") - 1)
                  cArchivos_Enviar = Right(cArchivos_Enviar, Len(cArchivos_Enviar) - InStr(1, cArchivos_Enviar, ","))
               Else
                  cFileSubMit = cArchivos_Enviar
                  cArchivos_Enviar = ""
               End If
            Else
                cFileSubMit = ""
            End If

       Case "Recibir"
            If Len(cArchivos_Recibir) > 0 Then
               If InStr(1, cArchivos_Recibir, ",") > 0 Then
                  cFileSubMit = Left(cArchivos_Recibir, InStr(1, cArchivos_Recibir, ",") - 1)
                  cArchivos_Recibir = Right(cArchivos_Recibir, Len(cArchivos_Recibir) - InStr(1, cArchivos_Recibir, ","))
               Else
                  cFileSubMit = cArchivos_Recibir
                  cArchivos_Recibir = ""
               End If
            Else
                cFileSubMit = ""
            End If

       Case "Borrar"
            If Len(cArchivos_Borrar) > 0 Then
               If InStr(1, cArchivos_Borrar, ",") > 0 Then
                  cFileSubMit = Left(cArchivos_Borrar, InStr(1, cArchivos_Borrar, ",") - 1)
                  cArchivos_Borrar = Right(cArchivos_Borrar, Len(cArchivos_Borrar) - InStr(1, cArchivos_Borrar, ","))
               Else
                  cFileSubMit = cArchivos_Borrar
                  cArchivos_Borrar = ""
               End If
            Else
                cFileSubMit = ""
            End If

End Select
'Ahora Descomponer Archivo en Nombre y Extensión
If Len(cFileSubMit) > 0 Then
    If InStr(1, cFileSubMit, ".") > 0 Then
        cNombre = Left(cFileSubMit, InStr(1, cFileSubMit, ".") - 1)
        cExt = Right(cFileSubMit, Len(cFileSubMit) - InStr(1, cFileSubMit, "."))
    Else
        cNombre = cFileSubMit
        cExt = ""
    End If

    nFound = InStr(1, cNombre, "?")
    cNombre = IIf(nFound = 0, cNombre, Left(cNombre, nFound - 1) + cLoginId + Right(cNombre, Len(cNombre) - nFound))
    cFileSubMit = cNombre + Space(8 - Len(cNombre)) + "." + RTrim(cExt)
End If

Lee_Submit = cFileSubMit

End Function

Function Master ()
cChar = Send_Command("01" + "00" + "X" + "AMG" + "  112233" + Space(8) + Space(6), True)
If Receive_Status() = "81" Then
   Master = FATAL_ERROR
   Exit Function
End If

Do While True
   If Receive_Command() = FATAL_ERROR Then
       Master = FATAL_ERROR
       Exit Function
   End If
   Select Case cComando
      Case "03"
           LENFILE = Val(Mid(cBlock, 15, 10))
           Crea_Archivo (Mid(cBlock, 3, 10))
           cChar = Receive_File()
           cChar = Send_Good_Status()

      Case "10"
           cChar = Send_Good_Status()
           cChar = Send_File(Mid(cBlock, 3, 10))

      Case "02"
           Master = FATAL_ERROR
           Exit Function
           
      Case Else
           MsgBox cComando
   End Select
Loop
Exit Function

End Function

Sub Mensaje (ByVal cMensaje)
Msg.Caption = cMensaje
Msg.Refresh
End Sub

Sub Menu_Acerca_Click ()
    Acerca_de.Show
End Sub

Sub Msg_DblClick ()
    Call Form_DblClick
End Sub

Sub Op_Automático_Click ()
    If Op_Automático.Caption = "&Automático" Then
       Op_Automático.Caption = "&Manual"
    Else
       Op_Automático.Caption = "&Automático"
    End If
    Op_Automático.Checked = True
End Sub

Sub Op_Borrar_Click ()
If Op_Automático.Caption = "&Manual" Then
    Op_Enviar.Checked = False
    Op_Recibir.Checked = False
    Op_Borrar.Checked = True
End If
TwoWay.WindowState = 1
End Sub

Sub Op_Configuracion_Click ()
    Setup.Show MODAL
End Sub

Sub Op_Enviar_Click ()
    If Op_Automático.Caption = "&Manual" Then
        Op_Enviar.Checked = True
        Op_Recibir.Checked = False
        Op_Borrar.Checked = False
     End If
     TwoWay.WindowState = 1

End Sub

Sub Op_Recibir_Click ()
    If Op_Automático.Caption = "&Manual" Then
        Op_Enviar.Checked = False
        Op_Recibir.Checked = True
        Op_Borrar.Checked = False
    End If
    TwoWay.WindowState = 1
End Sub

Sub Op_Salir_Click ()
    Unload Me
End Sub

Private Sub Puerto_OnComm ()
    Select Case Puerto.CommEvent
    ' Errors
        'Case comBreak   ' A Break was received.
        ' Code to handle a BREAK goes here.
        'Case comCDTO    ' CD (RLSD) Timeout.
        'Case comCTSTO   ' CTS Timeout.
        'Case comDSRTO   ' DSR Timeout.
         Case comFrame   ' Framing Error
              Mensaje "Error: Framing"
         Case comOverrun ' Data Lost.
              Mensaje "Error: Datos perdidos..."
         Case comRxOver  ' Receive buffer overflow.
              Mensaje "Error: Buffer de entrada lleno ..."
         Case comRxParity    ' Parity Error.
              Mensaje "Error: Paridad"
         Case comTxFull  ' Transmit buffer full.
              Mensaje "Error: Buffer de salida lleno ..."
    ' Events
        Case comEvCD    ' Change in the CD line.
             'Mensaje "Cambio en linea Carrier Detect "
        Case comEvCTS   ' Change in the CTS line.
             If Puerto.CTSHolding = True Then
                CTS.Picture = On_Pic.Picture
             Else
                CTS.Picture = Off_Pic.Picture
             End If
        Case comEvDSR   ' Change in the DSR line.
             If Puerto.DSRHolding = True Then
                DSR.Picture = On_Pic.Picture
             Else
                DSR.Picture = Off_Pic.Picture
             End If
        Case comEvRing  ' Change in the Ring Indicator.
             Mensaje "Cambio en Indicador Ring"
        Case comEvReceive   ' Received RThreshold # of chars.
            Receive.Picture = On_Pic.Picture
            Receive.Picture = Off_Pic.Picture
        Case comEvSend  ' There are SThreshold number of
                        ' characters in the transmit buffer.
            Send.Picture = On_Pic.Picture
            Send.Picture = Off_Pic.Picture
        Case comEvEOF
            'Mensaje "Recibiendo Fin de Archivo..."
        Case Else
            MsgBox Puerto.CommEvent
    End Select

End Sub

Function Quitar_CrLf (cBlock$) As String
Dim nFound%, nStart%, nLenBlock%

If lCRLF Then    '!::lCrLf
   nStart = 1
   Do While True
     nFound = InStr(nStart, cBlock, CRLF)
     nLenStr = Len(cBlock)
     If nFound > 0 Then
        cBlock = Left(cBlock, nFound - 1) + Mid(cBlock, nFound + 2, Len(cBlock) - nFound + 1)
        nStart = nFound - 1
     Else
        Exit Do
     End If
   Loop
End If
Quitar_CrLf = cBlock
End Function

Function Receive_Block (cBlock$)
Dim nCounter As Integer, cChar As String, nCntCar As Integer
Dim cBackStr$, Car1$, Car2$

  nCounter = 7
  cChar = ""
  
  cBlock = ""

Do While True
  nCntCar = 1
  If Send_Char(ACKP) = FATAL_ERROR Then
     Receive_Block = FATAL_ERROR
     Exit Function
  End If
  cChar = Receive_Control()
  If cChar <> STX Then
     MsgBox "PRIMER CARACTER MALO..."
     Receive_Block = FATAL_ERROR
     Exit Function
  End If
  Do While True
        cChar = ""
        Do While cChar <> EOT And cChar <> ENQ And cChar <> ETX And cChar <> ETB And cChar <> FATAL_ERROR 'And cChar <> STX
            nCntCar = nCntCar + 1
            cBlock = cBlock + cChar
            cChar = Receive_Control()
            DoEvents
        Loop
        
        Select Case cChar
               Case FATAL_ERROR
                    Receive_Block = FATAL_ERROR
                    Exit Function
               Case EOT
                    Receive_Block = EOT
                    Exit Function
               Case ENQ
                    nCounter = nCounter - 1
                    If nCounter = 0 Then
                        MsgBox "Fatal Error -> Receive Block ENQ En Receieve_Block"
                        Receive_Block = FATAL_ERROR
                        Exit Function
                    End If
                    Exit Do
               Case Else
                    If cChar = ETB Or cChar = ETX Then
                        cBackStr = cBlock
                        If Receive_Char() <> CalBCC(cBackStr, cChar) Then
                           nCounter = nCounter - 1
                           If nCounter > 0 Then
                              If Send_Char(NAK) = FATAL_ERROR Then
                          Receive_Block = FATAL_ERROR
                                 Exit Function
                              End If
                            Else
                                MsgBox "Fatal Error -> Receive Block Error BCC"
                                Receive_Block = FATAL_ERROR
                                Exit Function
                            End If
                        Else
                            Set_ACK
                            If cChar = ETX Then
                                If Receive_EOT() = FATAL_ERROR Then
                                   Receive_Block = FATAL_ERROR
                                   Exit Function
                                End If
                            End If
                            Receive_Block = cChar
                            Exit Function
                        End If
                    End If
        End Select
        
 Loop
Loop
Receive_Block = NUL

End Function

Function Receive_Command ()
Dim cChar$, cBuffer$
   ACKP = ACK0
   cChar = Wait_For_Bid()
   If cChar = FATAL_ERROR Then
      Receive_Command = FATAL_ERROR
      Exit Function
   End If
   
   cChar = ""
   cBuffer = ""
   cChar = Receive_Block(cBuffer)
   cBlock = cBuffer

   If cChar = ETB Then
      lLastBlock = False
   End If
   If cChar = ETX Then
      lLastBlock = True
   End If
   If cChar = EOT Then
      lLastBlock = True
   End If
   If cChar = FATAL_ERROR Then
      lLastBlock = True
   End If
   cComando = Mid(cBuffer, 1, 2)
Receive_Command = cChar
End Function

Function Receive_Control ()
Dim cControl As String, cChar As String, nStatus%
  cControl = Receive_Char()
  If cControl = Chr(16) Then  'DLE
     cChar = Receive_Char()
     If cChar = Chr(2) Or cChar = Chr(3) Or cChar = Chr(5) Or cChar = Chr(23) Or cChar = Chr(16) Or cChar = Chr(48) Or cChar = Chr(49) Or cChar = Chr(59) Then
        cControl = cControl + cChar
     Else
        nStatus = UngetCommChar(ByVal Puerto.CommID, ByVal Asc(cChar))
        If nStatus < 0 Then
           cControl = FATAL_ERROR
        End If
     End If
  End If
Receive_Control = cControl
End Function

Function Receive_Char () As String
  SetTimeOut (1000 * 21)
  Do While Puerto.InBufferCount = 0
     DoEvents
     If Chequea_CTS() = False Or IfTimeOut() Then
        Receive_Char = FATAL_ERROR
        Exit Function
     End If
  Loop
  Puerto.InputLen = 1
  Receive_Char = IIf(Chequea_CTS(), Puerto.Input, FATAL_ERROR)
End Function

Function Receive_EOT ()
Dim nCounter%, cChar$
nCounter = 7
cChar = ""

Do While True
   If Send_Char(ACKP) = FATAL_ERROR Then
      Receive_EOT = FATAL_ERROR
      Exit Function
   End If
   Do While True
      cChar = Receive_Char()

      If cChar = FATAL_ERROR Then
         Receive_EOT = FATAL_ERROR
         Exit Function
      End If
      If cChar = EOT Then
         Receive_EOT = OK
         Exit Function
      End If
      If cChar = ENQ Then
         nCounter = nCounter - 1
         If nCounter = 0 Then
            Exit Do
         Else
            MsgBox "ERROR FATAL -> Receive_EOT()"
            Receive_EOT = FATAL_ERROR
            Exit Function
         End If
      End If
   Loop
   
Loop

End Function

Function Receive_File ()
Dim cBuffer$, nLenFile%, nCntBlock%
     cBuffer = ""
     nCntBlock = 0
     nLenFile = 0
     Do While True
Continuar:
        cChar = Receive_Block(cBuffer) ' Por Referencia
        nLenFile = nLenFile + Len(cBuffer)
        Mensaje "Recibiendo " + Str(nLenFile) + "," + Str(Int(nLenFile / LENFILE * 100)) + "%" + " de " + Str(LENFILE)

        If cChar = ETB Then
           nCntBlock = nCntBlock + 1
           'Mensaje " Recibiendo Block Nro. " + Str(nCntBlock)

           Adiciona_Registro cBuffer
           GoTo Continuar
        End If
        If cChar = ETX Then
            nCntBlock = nCntBlock + 1
            Adiciona_Registro cBuffer + Chr(26)
            Cierra_Archivo
            'Mensaje "Recibiendo Ultimo Block Nro." + Str(nCntBlock)
            Exit Do
        End If
        If cChar = EOT Then
            Cierra_Archivo
            Exit Do
        End If
        If cChar = FATAL_ERROR Then
            Cierra_Archivo
            Exit Do
        End If
     Loop
     Receive_File = cChar
     Exit Function

End Function

Function Receive_Logon ()
cChar = ""

If Receive_Command() = FATAL_ERROR Then
   Receive_Logon = FATAL_ERROR
   Exit Function
End If
If cComando = "01" Then           ' Recibir confirmación.
    Mensaje "Sign On"
                
    cProfile = RTrim(Mid(cBlock, 17, 8)) + ".INI"
    cLoginId = LTrim(RTrim(Mid(cBlock, 9, 8)))

    If Mid(cBlock, 6, 3) <> "AMG" Then
        cChar = Send_Bad_Status()
        Receive_Logon = FATAL_ERROR
        Exit Function
    End If
                
    If Lee_Profile_Privado() Then
        cChar = Send_Good_Status()
        Receive_Logon = OK
    Else
        cChar = Send_Bad_Status()
        Receive_Logon = FATAL_ERROR
    End If

    Me.Caption = Right(Me.Caption, 1) + ":" + cLoginId + ":" + cProfile
    Exit Function
Else
    cChar = Send_Bad_Status()
    Receive_Logon = FATAL_ERROR
    Exit Function
End If
End Function

Function Receive_Status ()
  If Receive_Command() = FATAL_ERROR Then
        Receive_Status = FATAL_ERROR
        Exit Function
  End If
  Receive_Status = cComando
  Exit Function
  'Receive_Status = NULO
End Function

Function Send_Bad_Status ()
  cChar = Send_Command("81", True)
  If cChar = FATAL_ERROR Then
     Send_Bad_Status = cChar
     Exit Function
  End If
' ::Send_Char(cEOT)
Send_Bad_Status = NULO
End Function

Function Send_Bid ()
Dim nCounter%, cChar$, Tiempo As Long
 nCounter = 7
 cChar = ""
 Do While True
     Wait_Millisecs (500)
     If Send_Char(BID) = FATAL_ERROR Then
        Send_Bid = FATAL_ERROR
        Exit Function
     End If
     cChar = Receive_Control()

     If cChar = EOT Then
        GoTo Continuar_SB
     End If
     If cChar = ACK0 Then
        Send_Bid = ACK0
        Exit Function
     Else
        nCounter = nCounter - 1
        If nCounter = 0 Then
            MsgBox "ERROR FATAL -> Send_Bid()"
            Send_Bid = FATAL_ERROR
            Exit Function
        End If
     End If
Continuar_SB:
Loop
Send_Bid = NUL
End Function

Function Send_Block (ByVal cBlock$, lLastBlock%)        'SEND-BLOCK-ETX
Dim nRetry%, cBCC$
    nRetry = 6
    cBCC = ""
    If lLastBlock Then
       cBCC = CalBCC(cBlock, ETX)
       cBlock = STX + cBlock + ETX + cBCC
    Else
       cBCC = CalBCC(cBlock, ETB)
       cBlock = STX + cBlock + ETB + cBCC
    End If


    Set_ACK
    Do While True
       DoEvents
       If Send_Char(cBlock) = FATAL_ERROR Then
            Send_Block = FATAL_ERROR
            Exit Function
       End If
       DoEvents
       Do While True
          cChar = Receive_Control()
           Select Case cChar
                Case FATAL_ERROR
                    Send_Block = FATAL_ERROR
                    Exit Function
                Case ACKP
                    If lLastBlock Then
                        If Send_Char(EOT) = FATAL_ERROR Then
                           Send_Block = FATAL_ERROR
                           Exit Function
                        End If
                        Send_Block = EOT
                        Exit Function
                    End If
                    Send_Block = ACKP
                    Exit Function

                Case EOT
                    Send_Block = EOT
                    Exit Function
                
                Case NAK
                    nRetry = nRetry - 1
                    If nRetry = 0 Then
                        MsgBox "NAK Send_Block()"
                        Send_Block = FATAL_ERROR
                        Exit Function
                    Else
                        Exit Do
                    End If
                
                Case WACK
                    If Send_Char(ENQ) = FATAL_ERROR Then
                        Send_Block = FATAL_ERROR
                        Exit Function
                    End If
                    GoTo Continuar_Send_Block
                
                Case Else
                    If cChar <> ACKP Then
                        nRetry = nRetry - 1
                        If nRetry = 0 Then
                            MsgBox "BAD ACKP Send_Block()"
                            Send_Block = FATAL_ERROR
                            Exit Function
                        Else
                            Exit Do
                        End If
                    End If
            End Select

            If Send_Char(ENQ) = FATAL_ERROR Then
                Send_Block = FATAL_ERROR
                Exit Function
            End If

Continuar_Send_Block:
       Loop
    Loop
Send_Block = NUL
End Function

Function Send_Command (ByVal cBlock$, ByVal lLastBlock%)       'SEND-BLOCK-ETX
Dim cACKP$, cChar$
If Send_Bid() = FATAL_ERROR Then  'Pide Linea
   Send_Command = FATAL_ERROR
   Exit Function
End If

ACKP = ACK0
cComando = Mid(cBlock, 1, 2)

Select Case cComando
   Case "02"
        Mensaje "Desconectando"
   Case "03"
        Mensaje "Enviando Archivo ..."
   Case "10"
        Mensaje "Pidiendo Archivo..."
   Case "14"
        Mensaje "Borrando Archivo de Host"

   End Select
cChar = Send_Block(cBlock, lLastBlock)
Send_Command = cChar
End Function

Function Send_Char (ByVal Cadena$)
    Puerto.Output = Cadena
    Do While Puerto.OutBufferCount > 0
       If Chequea_CTS() = False Then
          Exit Do
       End If
       DoEvents
    Loop
    If Chequea_CTS() = False Then
       Send_Char = FATAL_ERROR
    Else
       Send_Char = OK
    End If

End Function

Function Send_File (cFile$)
Dim nBytesRead%, cChar$, cBuffer$, nBytesTrn%, nCntBlock%, nLenFile As Long
cBuffer = Space(BUFFERLEN + 4)
nCntBlock = 0
cComando = "03"
nLenFile = 0

Mensaje "Abriendo Archivo..."
If Abre_Archivo(cFile) = True Then
        TwoWay.Icon = Conn_Snd_Pic.Picture
        lLastBlock = False
        cChar = Send_Command(cComando + cFile + Space(2) + Format$(LENFILE, "0000000000") + Space(6), lLastBlock)
        If cChar = FATAL_ERROR Then
            Send_File = cChar
            Exit Function
        End If
        Do While True
            nCntBlock = nCntBlock + 1

            nBytesRead = Lee_Archivo(cBuffer, BUFFERLEN)
            nLenFile = nLenFile + nBytesRead
            Mensaje "Enviando  " + Str(nLenFile) + "," + Str(Int(nLenFile / LENFILE * 100)) + "%" + " de " + Str(LENFILE)
            If nBytesRead = BUFFERLEN Then
                'Mensaje "Enviando Block Nro. " + Str(nCntBlock)
                cBlock = Left(cBuffer, nBytesRead)
                lLastBlock = False
            Else
                'Mensaje "Enviando Ultimo Block Nro. " + Str(nCntBlock)

                If nBytesRead > 1 Then         ' Verificar alg·n error.
                    cBlock = Left(cBuffer, nBytesRead - 1)
                    lLastBlock = True
                    Cierra_Archivo
                Else
                    lLastBlock = True
                    cBlock = ""
                    Cierra_Archivo
                End If
            End If
            'nBytesTrn = nBytesTrn + Len(cBlock)

            cBlock = Quitar_CrLf(cBlock)
            cChar = Send_Block(cBlock, lLastBlock)
            If cChar = FATAL_ERROR Then
                Send_File = FATAL_ERROR
                Exit Function
            End If
            If cChar = EOT Then
                If Receive_Status() = "81" Then
                    Send_File = FATAL_ERROR
                    Exit Function
                Else
                    Send_File = OK
                    Exit Function
                End If
            End If
        Loop
End If

End Function

Function Send_Good_Status () As String
  Dim cChar$
  
  cChar = Send_Command("80", True)
  If cChar = FATAL_ERROR Then
     Send_Good_Status = cChar
     Exit Function
  End If
  Send_Good_Status = ""

End Function

Sub Send_Sign_Off ()
    cChar = Send_Command("0200" + Space(26), True)
End Sub

Function Send_String (ByVal cCadena$)
Dim nLen%, nCnt%
    nLen = Len(cCadena)
    For nCnt = 1 To nLen
       If Send_Char(Mid(cCadena, nCnt, 1)) = FATAL_ERROR Then
          Send_String = FATAL_ERROR
          Exit Function
       End If
       DoEvents
    Next
End Function

Sub Set_ACK ()
    ACKP = IIf(ACKP = ACK1, ACK0, ACK1)
End Sub

Sub Set_Time_Mode ()
Dim nHora_Ini%, nHora_Fin%, nTime%
If Op_Automático.Caption = "&Automático" Then
   nTime = Val(Left(Format(Time, "hh:mm"), 2)) * 60 + Val(Mid(Format(Time, "hh:mm"), 4, 2))
   If Len(cHora_Enviar) = 11 Then
        nHora_Ini = Val(Left(cHora_Enviar, 2)) * 60 + Val(Mid(cHora_Enviar, 4, 2))
        nHora_Fin = Val(Mid(cHora_Enviar, 7, 2)) * 60 + Val(Mid(cHora_Enviar, 10, 2)) ' + IIf(Val(Mid(cHora_Enviar, 10, 2)) < 24, 0, 24)
        If nHora_Ini <= nTime And nHora_Fin >= nTime Then
            Op_Enviar.Checked = True
        Else
            Op_Enviar.Checked = False
        End If
   End If
   If Len(cHora_Recibir) = 11 Then
        nHora_Ini = Val(Left(cHora_Recibir, 2)) * 60 + Val(Mid(cHora_Recibir, 4, 2))
        nHora_Fin = Val(Mid(cHora_Recibir, 7, 2)) * 60 + Val(Mid(cHora_Recibir, 10, 2))
        If nHora_Ini <= nTime And nHora_Fin >= nTime Then
            Op_Recibir.Checked = True
        Else
            Op_Recibir.Checked = False
        End If
   End If
   If Len(cHora_Borrar) = 11 Then
        nHora_Ini = Val(Left(cHora_Borrar, 2)) * 60 + Val(Mid(cHora_Borrar, 4, 2))
        nHora_Fin = Val(Mid(cHora_Borrar, 7, 2)) * 60 + Val(Mid(cHora_Borrar, 10, 2))
        If nHoraIni <= nTime And nHora_Fin >= nTime Then
            Op_Borrar.Checked = True
        Else
            Op_Borrar.Checked = False
        End If
   End If

End If
End Sub

Sub Set_Vars ()
On Error GoTo Set_Vars_Err
Dim nFound%
    
    cPrograma = "TWOWAY.EXE"
    cProfile = "TWOWAY.INI"
    cLogFile = "TWOWAY.LOG"

    NUL = Chr(0)
    lLastBlock = True
    lModificado = False
    If lActivo = False Then
        lCRLF = True
        SOH = Chr(1)
        EOT = Chr(4)
        BID = Chr(5)
        DLE = Chr(16)
        NAK = Chr(21)
        

        ACK0 = Chr(16) + Chr(48)
        ACK1 = Chr(16) + Chr(49)
        WACK = Chr(16) + Chr(59)
        'TTD = STX + ENQ

        ACKP = ACK0

        FATAL_ERROR = Chr(16) + Chr(4)
        OK = Chr(16)

        'cBlock = ""
        CRLF = Chr(13) + Chr(10)

        BUFFERLEN = 512


        Select Case GetModuleUsage(GetModuleHandle(cPrograma))
            Case 1
                cInstancia = "TwoWay Uno"
                Me.Caption = "TwoWay:1"
            Case 2
                cInstancia = "TwoWay Dos"
                Me.Caption = "TwoWay:2"
            Case 3
                cInstancia = "TwoWay Tres"
                Me.Caption = "TwoWay:3"
            Case 4
                cInstancia = "TwoWay Cuatro"
                Me.Caption = "TwoWay:4"
            Case Else
                cInstancia = "TwoWay Uno"
                Me.Caption = "TwoWay:1"
        End Select
        
        cDirectorio_Programa = Space(200)
        nFound = GetModuleFileName(GetModuleHandle(cPrograma), cDirectorio_Programa, Len(cDirectorio_Programa))
        cDirectorio_Programa = UCase(Mid(cDirectorio_Programa, 1, nFound))

        If Len(cDirectorio_Programa) = 0 Then
           cDirectorio_Programa = "C:\TWOWAY\TWOWAY.EXE"
        End If
        
        nFound = InStr(1, cDirectorio_Programa, cPrograma)
        cDirectorio_Programa = Left(cDirectorio_Programa, nFound - 1)
        lActivo = True
        Op_Enviar.Checked = IIf(Lee_Profile(cInstancia, "Modo", "Manual") = "Automatico", False, True)
    End If
    
    
    nPuerto = Val(Lee_Profile(cInstancia, "Puerto", "1"))
    
    cHora_Enviar = Lee_Profile(cInstancia, "HorasEnviar", "")
    cHora_Recibir = Lee_Profile(cInstancia, "HorasRecibir", "")
    cHora_Borrar = Lee_Profile(cInstancia, "HorasBorrar", "")
    
    cDirectorio_Enviar = Lee_Profile(cInstancia, "DirectorioEnviar", "")
    cDirectorio_Enviar = cDirectorio_Enviar + IIf(Right(cDirectorio_Enviar, 1) = "\", "", "\")
    cDirectorio_Recibir = Lee_Profile(cInstancia, "DirectorioRecibir", "")
    cDirectorio_Recibir = cDirectorio_Recibir + IIf(Right(cDirectorio_Recibir, 1) = "\", "", "\")

    cModo = Lee_Profile(cInstancia, "Modo", "Manual")
    cTipo = Lee_Profile(cInstancia, "Tipo", "Slave")
    cArchivos_Enviar = Lee_Profile(cInstancia, "ArchivosEnviar", "")
    cArchivos_Recibir = Lee_Profile(cInstancia, "ArchivosRecibir", "")
    cArchivos_Borrar = Lee_Profile(cInstancia, "ArchivosBorrar", "")
    lCRLF = IIf(Lee_Profile(cInstancia, "QuitarCRLF", "SI") = "SI", True, False)
    lLog = IIf(Lee_Profile(cInstancia, "Archivo_LOG", "SI") = "SI", True, False)
    If Lee_Profile(cInstancia, "Formato", "ASCII") = "ASCII" Then
       cFormato = "ASCII"
       STX = Chr(2)
       ETX = Chr(3)
       ENQ = Chr(5)
       ETB = Chr(23)
       DLE = Chr(16)
    Else              'TRANSPARENT
       cFormato = "TRANSPARENT"
       STX = Chr(16) + Chr(2)
       ETX = Chr(16) + Chr(3)
       ENQ = Chr(16) + Chr(5)
       ETB = Chr(16) + Chr(23)
       DLE = Chr(16) + Chr(16)
    End If
    Op_Automático.Caption = IIf(cModo = "Automatico", "&Automático", "&Manual")
    If cModo = "Manual" Then
        cInicio = Lee_Profile(cInstancia, "InicioManual", "Enviar")
        If cInicio = "Enviar" Then
            Op_Enviar.Checked = True
            Op_Recibir.Checked = False
            Op_Borrar.Checked = False
        ElseIf cInicio = "Recibir" Then
            Op_Enviar.Checked = False
            Op_Recibir.Checked = True
            Op_Borrar.Checked = False
        Else
            Op_Enviar.Checked = False
            Op_Recibir.Checked = False
            Op_Borrar.Checked = True
        End If
    End If
    Me.Caption = cInstancia + ":" + LTrim(Str(nPuerto))

GoTo Set_Vars_End

Set_Vars_Err:
  If Err = 53 Then
     MsgBox "Función DLL no soportado"
     End
  Else
     MsgBox Str(Err)
     End
  End If
  Resume Set_Vars_End

Set_Vars_End:

End Sub

Sub SetTimeOut (Millisecs As Long)
    TimeOutIni = GetTickCount()
    TimeOut = Millisecs
End Sub

Function Slave ()
Dim cChar$
TwoWay.Icon = Conn_Off_Pic.Picture
Do While True
    If Puerto.CTSHolding = True Then
        TwoWay.Icon = Conn_On_Pic.Picture
        Exit Do
    End If
    DoEvents
    Set_Time_Mode
    If lModificado = True Then
       Slave = FATAL_ERROR
       Exit Function
    End If
Loop
If Receive_Logon() = FATAL_ERROR Then
   Slave = FATAL_ERROR
End If
End Function

Sub Titulos_DblClick ()
 Call Form_DblClick
End Sub

Function Wait_For_Bid ()
   If Receive_Char() = BID Then
        Wait_For_Bid = OK
        Exit Function
   End If
Wait_For_Bid = FATAL_ERROR
End Function

Sub Wait_Millisecs (ByVal Millisecs As Long)
    Dim PauseTimer As Long, Start As Long
    PauseTime = Millisecs
    Start = GetTickCount()
    Do While GetTickCount() < Start + PauseTime
    Loop
End Sub


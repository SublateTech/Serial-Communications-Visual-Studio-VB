VERSION 2.00
Begin Form TwoWay 
   BackColor       =   &H00800000&
   BorderStyle     =   0  'None
   Caption         =   "Terminal"
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
      Height          =   105
      Left            =   1350
      Picture         =   TWOWAY.FRX:163C
      ScaleHeight     =   75
      ScaleWidth      =   210
      TabIndex        =   6
      Top             =   1605
      Width           =   240
   End
   Begin PictureBox DSR 
      AutoSize        =   -1  'True
      Height          =   105
      Left            =   2070
      Picture         =   TWOWAY.FRX:16DE
      ScaleHeight     =   75
      ScaleWidth      =   210
      TabIndex        =   5
      Top             =   1605
      Width           =   240
   End
   Begin PictureBox Receive 
      AutoSize        =   -1  'True
      Height          =   105
      Left            =   645
      Picture         =   TWOWAY.FRX:1780
      ScaleHeight     =   75
      ScaleWidth      =   210
      TabIndex        =   4
      Top             =   1605
      Width           =   240
   End
   Begin PictureBox Send 
      AutoSize        =   -1  'True
      Height          =   105
      Left            =   990
      Picture         =   TWOWAY.FRX:1822
      ScaleHeight     =   75
      ScaleWidth      =   210
      TabIndex        =   3
      Top             =   1605
      Width           =   240
   End
   Begin PictureBox CTS 
      AutoSize        =   -1  'True
      Height          =   105
      Left            =   2430
      Picture         =   TWOWAY.FRX:18C4
      ScaleHeight     =   75
      ScaleWidth      =   210
      TabIndex        =   2
      Top             =   1605
      Width           =   240
   End
   Begin PictureBox Port 
      AutoSize        =   -1  'True
      Height          =   105
      Left            =   270
      Picture         =   TWOWAY.FRX:1966
      ScaleHeight     =   75
      ScaleWidth      =   210
      TabIndex        =   1
      Top             =   1605
      Width           =   240
   End
   Begin PictureBox DTR 
      AutoSize        =   -1  'True
      Height          =   105
      Left            =   1710
      Picture         =   TWOWAY.FRX:1A08
      ScaleHeight     =   75
      ScaleWidth      =   210
      TabIndex        =   0
      Top             =   1605
      Width           =   240
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
Dim cDirectorio_Programa As String
Dim nHora_Ini%, nHora_Fin%, nTime%

Function Abre_Archivo (cArchivo$)
    FileName = cArchivo
    FILENUM = lopen(cDirectorio_Enviar + IIf(Right(cDirectorio_Enviar, 1) = "\", "", "\") + cArchivo, 0)
    If FILENUM = -1 Then
       Mensaje "No existe Archivo..."
       Abre_Archivo = False
    Else
       Abre_Archivo = True
    End If
End Function

Sub Activa_Linea ()
    
    Puerto.CommPort = Val(Lee_Profile(cInstancia, "Puerto", "1"))
    Puerto.Settings = Lee_Profile(cInstancia, "Velocidad", "19200") + ",O,7,2"
    
    'Set el Buffer de Entrada
    Puerto.OutBufferSize = BUFFERLEN + 3
    Puerto.InBufferSize = BUFFERLEN * 2  '1024
    Puerto.Interval = 0
    Puerto.NullDiscard = True
    Puerto.ParityReplace = "?"
    Puerto.Handshaking = MSCOMM_HANDSHAKE_RTS

    If Not Puerto.PortOpen Then
	Puerto.PortOpen = True
	Port.Picture = LoadPicture(cDirectorio_Programa + "\" + "ON.BMP")
    End If


     'Ok Line
    Puerto.RTSEnable = True
    RTS.Picture = LoadPicture(cDirectorio_Programa + "\" + "ON.BMP")

    'Listo para Recibir Data
    Puerto.DTREnable = True
    DTR.Picture = LoadPicture(cDirectorio_Programa + "\" + "ON.BMP")

    'Para capturar se±ales de env_o y recepci_n
    Puerto.RThreshold = 1
    Puerto.SThreshold = 1
   

    If Puerto.CTSHolding = True Then
	CTS.Picture = LoadPicture(cDirectorio_Programa + "\" + "ON.bmp")
    Else
	CTS.Picture = LoadPicture(cDirectorio_Programa + "\" + "OFF.bmp")
    End If


    If Puerto.DSRHolding = True Then
	DSR.Picture = LoadPicture(cDirectorio_Programa + "\" + "ON.bmp")
    Else
	DSR.Picture = LoadPicture(cDirectorio_Programa + "\" + "OFF.bmp")
    End If


End Sub

Sub Adiciona_Registro (ByVal cRegistro$)
 nCount = lwrite(FILENUM, cRegistro, Len(cRegistro))
End Sub

Function CalBCC (sStr$, cControl$)
Dim nBCC%, nCnt%, nChar%, cCadena$
nBCC = 0
nCnt = 1
nChar = 0
cCadena = ""
For nCnt = 1 To Len(sStr)
    nChar = Asc(Mid(sStr, nCnt, 1))
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
    cCadena = cCadena + Chr(nChar)
    nBCC = nBCC Xor nChar
Next
sStr = cCadena
nBCC = nBCC Xor Asc(cControl)
CalBCC = Chr(nBCC And 255)
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

    TwoWay.Icon = LoadPicture(cDirectorio_Programa + "\" + "Conn_Off.ico")
    'Drop Line
    Puerto.RTSEnable = False
    RTS.Picture = LoadPicture(cDirectorio_Programa + "\" + "OFF.BMP")

   ' 'NO Listo para transmitir
    Puerto.DTREnable = False
    DTR.Picture = LoadPicture(cDirectorio_Programa + "\" + "OFF.BMP")

    If Puerto.PortOpen Then
	Puerto.PortOpen = False
	Port.Picture = LoadPicture(cDirectorio_Programa + "\" + "OFF.BMP")
    End If

    If Puerto.CTSHolding = True Then
	CTS.Picture = LoadPicture(cDirectorio_Programa + "\" + "ON.bmp")
    Else
	CTS.Picture = LoadPicture(cDirectorio_Programa + "\" + "OFF.bmp")
    End If

    If Puerto.DSRHolding = True Then
	DSR.Picture = LoadPicture(cDirectorio_Programa + "\" + "ON.bmp")
    Else
	DSR.Picture = LoadPicture(cDirectorio_Programa + "\" + "OFF.bmp")
    End If


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
   TwoWay.Icon = LoadPicture(cDirectorio_Programa + "\" + "Conn_Rcv.ico")
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
   Crea_Archivo Mid(cBlock, 3, 10)
'   MsgBox Str(Puerto.InBufferCount)
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
    Dim cFileSubmit$

    Me.Show

Do While True
    Set_Vars
    Mensaje "Esperando señal ..."
    Activa_Linea

    If Slave() <> FATAL_ERROR Then
	If Op_Enviar.Checked = True Then
	   cFileSubmit = Lee_Submit("Enviar")
	   Do While cFileSubmit <> ""
	      cChar = Send_File(cFileSubmit)
	      If cChar = FATAL_ERROR Then
		 Exit Do
	      End If
	      cFileSubmit = Lee_Submit("Enviar")
	   Loop
	End If
	If Op_Recibir.Checked = True Then
	   cFileSubmit = Lee_Submit("Recibir")
	   Do While cFileSubmit <> ""
	      cChar = Fetch_File(cFileSubmit)
	      If cChar = FATAL_ERROR Then
		 'lOk = False
		 Exit Do
	      End If
	
	      cFileSubmit = Lee_Submit("Recibir")
	   Loop
	End If
	If Op_Borrar.Checked = True Then
	   cFileSubmit = Lee_Submit("Borrar")
	   Do While cFileSubmit <> ""
	      cChar = Erase_File(cFileSubmit)
	      If cChar = FATAL_ERROR Then
		 'lOk = False
		 Exit Do
	      End If
	      cFileSubmit = Lee_Submit("Borrar")
	   Loop
	End If

	Send_Sign_Off
    End If
    Desactiva_Linea
Loop
End Sub

Sub Form_Unload (Cancel As Integer)
    Mensaje "Cerrando aplicación ..."
    Desactiva_Linea
    Unload Me
    End
End Sub

Function Lee_Archivo (cBuffer$, nLenArchivo%)
   Dim cBlock$
   cBlock = Space(BUFFERLEN + 3)
   nCount = lread(FILENUM, cBlock, nLenArchivo)
   cBuffer = cBlock
   Lee_Archivo = nCount
End Function

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

    nFound = InStr(1, cNombre, "?")
    cFileSubmit = IIf(nFound = 0, cNombre, Left(cNombre, nFound - 1) + cLoginId + Right(cNombre, Len(cNombre) - nFound)) + "." + cExt
End If

Lee_Submit = cFileSubmit

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
	'Case comFrame   ' Framing Error
	'Case comOverrun ' Data Lost.
	 Case comRxOver  ' Receive buffer overflow.
	      Puerto.InBufferSize = Puerto.InBufferSize + 256
	'Case comRxParity    ' Parity Error.
	'Case comTxFull  ' Transmit buffer full.
	      
    ' Events
	'Case comEvCD    ' Change in the CD line.

	Case comEvCTS   ' Change in the CTS line.
	     If Puerto.CTSHolding = True Then
		CTS.Picture = LoadPicture(cDirectorio_Programa + "\" + "ON.bmp")
	     Else
		CTS.Picture = LoadPicture(cDirectorio_Programa + "\" + "OFF.bmp")
	     End If

	Case comEvDSR   ' Change in the DSR line.
	     If Puerto.DSRHolding = True Then
		DSR.Picture = LoadPicture(cDirectorio_Programa + "\" + "ON.bmp")
	     Else
		DSR.Picture = LoadPicture(cDirectorio_Programa + "\" + "OFF.bmp")
	     End If

	'Case comEvRing  ' Change in the Ring Indicator.
	Case comEvReceive   ' Received RThreshold # of chars.
	    Receive.Picture = LoadPicture(cDirectorio_Programa + "\" + "ON.BMP")
	    Receive.Picture = LoadPicture(cDirectorio_Programa + "\" + "OFF.BMP")

	Case comEvSend  ' There are SThreshold number of
			' characters in the transmit buffer.
	    Send.Picture = LoadPicture(cDirectorio_Programa + "\" + "ON.BMP")
	    Send.Picture = LoadPicture(cDirectorio_Programa + "\" + "OFF.BMP")
	Case Else
	    MsgBox Puerto.CommEvent
    End Select

End Sub

Function Quitar_CrLf (cBlock$) As String
Dim nFound%, nStart%, nLenBlock%

'IF !::lCrLf
   nStart = 1
   Do While True
     nFound = InStr(nStart, cBlock, CRLF)
     nLenStr = Len(cBlock)
     If nFound > 0 Then
	cBlock = Left(cBlock, nFound - 1) + Mid(cBlock, nFound + 2, Len(cBlock) - nFound + 1)
	Start = nFound
     Else
	Exit Do
     End If
   Loop
'End If
Quitar_CrLf = cBlock
End Function

Function Receive_Block (cBlock$)
Dim nCounter As Integer, cChar As String, nCntCar As Integer
Dim cBuffer$, cString$
  cBuffer = ""
  nCounter = 7
  cChar = ""
  nCntCar = 0
  cBlock = ""

Do While True
  If Send_Char(ACKP) = FATAL_ERROR Then
     Receive_Block = FATAL_ERROR
     Exit Function
  End If
  
  Do While True
	cChar = ""
	Do While cChar <> EOT And cChar <> ENQ And cChar <> STX And cChar <> ETX And cChar <> ETB And cChar <> FATAL_ERROR
	    nCntCar = nCntCar + 1
	    cBlock = cBlock + cChar
	    cChar = Receive_String(1)
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
	       Case STX
		    If nCntCar <> 1 Then
			MsgBox "PRIMER CARACTER MALO..."
			Receive_Block = FATAL_ERROR
			Exit Function
		    End If
	       Case Else
		    If cChar = ETB Or cChar = ETX Then
			If Receive_String(1) <> CalBCC(cBlock, cChar) Then
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
Dim cControl As String
  cControl = Receive_String(1)
  If cControl = ACK Then
     cControl = cControl + Receive_String(1)
  End If
Receive_Control = cControl
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
      cChar = Receive_String(1)

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
     nCntBlock = 0
     Do While True
Continuar:
	cChar = Receive_Block(cBuffer) ' Por Referencia
	'nLenFile = nLenFile + Len(cBuffer)
	If cChar = ETB Then
	   nCntBlock = nCntBlock + 1
	   Mensaje " Recibiendo Block Nro. " + Str(nCntBlock)

	   Adiciona_Registro cBuffer
	   GoTo Continuar
	End If
	If cChar = ETX Then
	    nCntBlock = nCntBlock + 1
	    Adiciona_Registro cBuffer
	    Cierra_Archivo
	    Mensaje "Recibiendo Ultimo Block Nro." + Str(nCntBlock)

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

Function Receive_Status ()
  If Receive_Command() = FATAL_ERROR Then
	Receive_Status = FATAL_ERROR
	Exit Function
  End If
  Receive_Status = cComando
  Exit Function
  'Receive_Status = NULO
End Function

Function Receive_String (nCantidad%) As String
  Do While Puerto.InBufferCount < nCantidad
     DoEvents
     If Chequea_CTS() = False Then
	Receive_String = FATAL_ERROR
	Exit Function
     End If
  Loop
  Puerto.InputLen = nCantidad
  Receive_String = IIf(Chequea_CTS(), Puerto.Input, FATAL_ERROR)
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
Dim nCounter%, cChar$
 nCounter = 7
 cChar = ""
 Do While True
     Wait_Millisecs (300)
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
	MsgBox "Aqui Send_Bid " + Str(Asc(cChar))
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
Dim nBytesRead%, cChar$, cBuffer$, nBytesTrn%, nCntBlock%
cBuffer = Space(BUFFERLEN + 4)
nCntBlock = 0
cComando = "03"

Mensaje "Abriendo Archivo..."
If Abre_Archivo(cFile) = True Then
	TwoWay.Icon = LoadPicture(cDirectorio_Programa + "\" + "Conn_Snd.ico")
	lLastBlock = False
	cChar = Send_Command(cComando + cFile + Right(Str(nLenFile), 6) + Space(10), lLastBlock)
	If cChar = FATAL_ERROR Then
	    Send_File = cChar
	    Exit Function
	End If
	Do While True
	    nCntBlock = nCntBlock + 1

	    nBytesRead = Lee_Archivo(cBuffer, BUFFERLEN)
	    If nBytesRead = BUFFERLEN Then
		Mensaje "Enviando Block Nro. " + Str(nCntBlock)
		cBlock = Left(cBuffer, nBytesRead)
		lLastBlock = False
	    Else
		Mensaje "Enviando Ultimo Block Nro. " + Str(nCntBlock)

		If nBytesRead > 1 Then         ' Verificar alg·n error.
		    cBlock = Left(cBuffer, nBytesRead)
		    lLastBlock = True
		    Cierra_Archivo
		Else
		    lLastBlock = True
		    cBlock = ""
		    Cierra_Archivo
		End If
	    End If
	    nBytesTrn = nBytesTrn + Len(cBlock)

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
Dim nFound%
    NUL = Chr(0)

    lLastBlock = True
    lModificado = False
    If lActivo = False Then
	SOH = Chr(1)
	STX = Chr(2)
	ETX = Chr(3)
	EOT = Chr(4)
	ENQ = Chr(5)       'Enquiry
	BID = Chr(5)
	DLE = Chr(16)
	NAK = Chr(21)
	ETB = Chr(23)

	ACK = Chr(16)
	ACK0 = Chr(16) + Chr(48)
	ACK1 = Chr(16) + Chr(49)
	WACK = Chr(16) + Chr(59)
	TTD = STX + ENQ

	ACKP = ACK0

	FATAL_ERROR = Chr(220)
	OK = Chr(200)

	'cBlock = ""
	CRLF = Chr(13) + Chr(10)

	BUFFERLEN = 512


	Select Case GetModuleUsage(GetModuleHandle("TwoWay.EXE"))
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
	nFound = GetModuleFileName(GetModuleHandle("TwoWay.EXE"), cDirectorio_Programa, Len(cDirectorio_Programa))
	cDirectorio_Programa = Mid(cDirectorio_Programa, 1, nFound)

	If Len(cDirectorio_Programa) = 0 Then
	   cDirectorio_Programa = "C:\TwoWay\TWOWAY.EXE"
	End If
	
	nFound = InStr(1, cDirectorio_Programa, "TWOWAY.EXE")
	cDirectorio_Programa = Left(cDirectorio_Programa, nFound - 2)
	lActivo = True
	Op_Enviar.Checked = IIf(Lee_Profile(cInstancia, "Modo", "False") = "Automatico", False, True)
    Else
	Me.Caption = Left(Me.Caption, 8)
    End If


    cHora_Enviar = Lee_Profile(cInstancia, "Horas_Enviar", "")
    cHora_Recibir = Lee_Profile(cInstancia, "Horas_Recibir", "")
    cHora_Borrar = Lee_Profile(cInstancia, "Horas_Borrar", "")
    cDirectorio_Enviar = Lee_Profile(cInstancia, "Directorio_Enviar", "ENVIA\")
    cDirectorio_Recibir = Lee_Profile(cInstancia, "Directorio_Recibir", "RECIBE\")
    Op_Automático.Caption = IIf(Lee_Profile(cInstancia, "Modo", "False") = "Automatico", "&Automático", "&Manual")
    cArchivos_Enviar = Lee_Profile(cInstancia, "Archivos_Enviar", "")
    cArchivos_Recibir = Lee_Profile(cInstancia, "Archivos_Recibir", "")
    cArchivos_Borrar = Lee_Profile(cInstancia, "Archivos_Borrar", "")

End Sub

Function Slave ()
Dim cChar$
Do While True
    If Puerto.CTSHolding = True Then
	TwoWay.Icon = LoadPicture(cDirectorio_Programa + "\" + "Conn_On.Ico")
	Exit Do
    End If
    DoEvents
    Set_Time_Mode
    If lModificado = True Then
       Slave = FATAL_ERROR
       Exit Function
    End If
Loop

cChar = ""
Do While True
       If Receive_Command() = FATAL_ERROR Then
	  Slave = FATAL_ERROR
	  Exit Function
       End If
       Select Case cComando
	    Case "01"            ' Recibir confirmación.
		Mensaje "Sign On"
		cChar = Send_Good_Status()
		
		cLoginId = RTrim(Mid(cBlock, 9, 8))
		Me.Caption = Me.Caption + ":" + cLoginId
	
		Slave = OK
		Exit Function

       End Select
    
 Loop
Slave = ""
End Function

Sub Titulos_DblClick ()
 Call Form_DblClick
End Sub

Function Wait_For_Bid ()
   If Receive_String(1) = ENQ Then
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


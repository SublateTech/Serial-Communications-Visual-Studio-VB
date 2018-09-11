*********** Sample programs for CA-Clipper Tools / 5.2 ************
*
* Program     : XMODEM.PRG
* Copyright   : (c) 1991-93, Computer Associates International.
*               All rights reserved.
* Date        : 07/20/91
* Author      : Michael Peters / Lothar Bongartz
* Description : XMODEM and XMODEM 1k file transfer
* Compiling   : CLIPPER XMODEM /N
* Linking     : - to use only in conjunction with other programs -
*               RTLINK FI <prog>, XMODEM LIB CT
* Remarks     :
*
* XMODEM.PRG is not an executable program!
*
* In this file you will find three functions for sending and receiving  n
* of files via XMODEM protocol. Compiled and linked with other programs
* this module will become an easy to use function for file transfer.
*
*
*       IMPORTANT IMPORTANT IMPORTANT IMPORTANT IMPORTANT IMPORTANT
*
* In conjunction with the communication protocols the terms XMODEM,
* XMODEM 1k, YMODEM and YMODEM Batch are used differently, which was
* very confusing. In the following you will find an extract of the
* original XMODEM/YMODEM documentation, written by Chuck Forsberg in
* 1987:
*
*  1.  TOWER OF BABEL
*
*  A "YMODEM Tower of Babel" has descended on the microcomputing community
*  bringing with it confusion, frustration, bloated phone bills, and wasted
*  man hours.  Sadly, I (Chuck Forsberg) am partly to blame for this mess.
*
*  As author of the early 1980s batch and 1k XMODEM extensions, I assumed
*  readers of earlier versions of this document would implement as much of
*  the YMODEM protocol as their programming skills and computing environments
*  would permit.  This proved a rather naive assumption as programmers
*  motivated by competitive pressure implemented as little of YMODEM as
*  possible.  Some have taken whatever parts of YMODEM that appealed to them,
*  applied them to MODEM7 Batch, Telink, XMODEM or whatever, and called the
*  result YMODEM.
*
*  Jeff Garbers (Crosstalk package development director) said it all: "With
*  protocols in the public domain, anyone who wants to dink around with them
*  can go ahead." [1]
*
*  Documents containing altered examples derived from YMODEM.DOC have added
*  to the confusion.  In one instance, the heading in YMODEM.DOC's Figure 1
*  has mutated from "1024 byte Packets" to "YMODEM/CRC File Transfer
*  Protocol".  None of the XMODEM and YMODEM examples shown in that document
*  were correct.
*
*  To put an end to this confusion, we must make "perfectly clear" what
*  YMODEM stands for, as Ward Christensen defined it in his 1985 coining of
*  the term.
*
*  To the majority of you who read, understood, and respected Ward's
*  definition of YMODEM, I apologize for the inconvenience.
*
*  1.1  Definitions
*
*  XMODEM  refers to the file transfer etiquette introduced by Ward
*          Christensen's 1977 MODEM.ASM program.  The name XMODEM comes from
*          Keith Petersen's XMODEM.ASM program, an adaptation of MODEM.ASM
*          for Remote CP/M (RCPM) systems.  It's also called the MODEM or
*          MODEM2 protocol.  Some who are unaware of MODEM7's unusual batch
*          file mode call it MODEM7.  Other aliases include "CP/M Users'
*          Group" and "TERM II FTP 3".  The name XMODEM caught on partly
*          because it is distinctive and partly because of media interest in
*          bulletin board and RCPM systems where it was accessed with an
*          "XMODEM" command.  This protocol is supported by every serious
*          communications program because of its universality, simplicity,
*          and reasonable performance.
*
*  XMODEM/CRC replaces XMODEM's 1 byte checksum with a two byte Cyclical
*          Redundancy Check (CRC-16), giving modern error detection
*          protection.
*
*  XMODEM-1k Refers to the XMODEM/CRC protocol with 1024 byte data blocks.
*
*  YMODEM  Refers to the XMODEM/CRC (optional 1k blocks) protocol with batch
*          transmission as described below.  In a nutshell, YMODEM means
*          BATCH.
*
*  YMODEM-g Refers to the streaming YMODEM variation described below.
*
*  True YMODEM(TM) In an attempt to sort out the YMODEM Tower of Babel, Omen
*          Technology has trademarked the term True YMODEM(TM) to represent
*          the complete YMODEM protocol described in this document, including
*          pathname, length, and modification date transmitted in block 0.
*          Please contact Omen Technology about certifying programs for True
*          YMODEM(TM) compliance.
*
*  ZMODEM  uses familiar XMODEM/CRC and YMODEM technology in a new protocol
*          that provides reliability, throughput, file management, and user
*          amenities appropriate to contemporary data communications.
**
**************************************************************************
*
* Control codes
*
 #define cCAN CHR(24)
 #define cACK CHR(6)
 #define cNAK CHR(21)
 #define cEOT CHR(4)
 #define cSOH CHR(1)
 #define cSTX CHR(2)
 #define cEOF CHR(26)
 #define nCanKey 27
*
**************************************************************************
*
* Function : XMO_TRANS()
* Returns  : LOGICAL returns .T., when the transfer has been
*            successful.
*
* Usage    : The function is able to transfer a file via XMODEM
*            protocol with a block length of 128 Byte.
*            Both, CHECKSUM and CRC can be used. The test mode to be
*            used is determined automatically by the function and the
*            remote station.
*            The function has been tested with PROCOMM 2.4.2 on the
*            remote station and a direct cable connection.
*
* Example  :
*            nPort := 1
*            nBaudRate := 2400
*
*            COM_OPEN(nPort, 1000)
*            COM_INIT(nPort, nBaudrate, "N", 8, 1)
*
*             ACCEPT "File : " TO cFile
*             ? XMO_TRANS(nPort, cFile)       // send file on port 1
*
*            COM_CLOSE(nPort)
*
**************************************************************************
FUNCTION XMO_TRANS(nPort, cFile)
LOCAL nBlockNum, nMaxBsize, cMessage, cBuffer, nHandle
LOCAL nErrCount := 0, Tx_Ok := .F., lChkMode := .F., nReadData := 0

 CSETREF(.T.)                              // highest performance

* transfer variable
*
 nBlockNum  := 1
 nMaxBsize  := 128
 cBuffer    := REPLICATE(cEOF, nMaxBsize)  // data input buffer
 nHandle    := FOPEN(cFile, 0)             // read only

 * Wait for begin of transfer (or termination)
 * When receiving character "C" empfangen, the transfer of the first
 * block starts immediately; the CRC protocol required by the remote
 * station is accepted.

 WAITPERIOD(6000)                          // loop for max. 60 seconds
 cChar := COM_READ(nPort, 1)               // read one character

 DO WHILE .NOT. cChar $ "C"+cNAK+cCAN .AND. WAITPERIOD()
   cChar := COM_READ(nPort, 1)             // read one character
 ENDDO

 lChkMode := (cChar = "C")                 // CRC (.T.), Checksum (.F.)

 MILLISEC(100)                             // wait for 0.1 second

 IF cChar = cNAK .OR. cChar = "C"          // answer from remote station
  nReadData := FREAD(nHandle, @cBuffer, nMaxBsize)
  cMessage  := "Sending Block:  "

  DO WHILE nReadData > 0 .AND. nErrCount <= 10 .AND. cChar <> cCAN
   ? cMessage + NTOC(nBlockNum, 10, 4)

   COM_SEND(nPort, XMOBLOCK(cBuffer, nBlockNum, lChkMode))

   * wait for block confirmation
   *
    WAITPERIOD(1000)                       // loop for max. 10 seconds
    cChar := COM_READ(nPort, 1)            // read one character

    DO WHILE .NOT. cChar $ cACK+cCAN .AND. WAITPERIOD()
     cChar := COM_READ(nPort, 1)           // read one character
    ENDDO

    IF cChar = cACK
     ?? " - verified "
     ++nBlockNum
     nErrCount := 0
     RANGEREPL(0, 255, @cBuffer, CHR(26))
     nReadData := FREAD(nHandle, @cBuffer, nMaxBsize)

     cMessage := "Sending Block:  "
    ELSE
     ++nErrCount                           // at NAK or timeout -> again
     cMessage := "Resending block:"
    ENDIF
  ENDDO

  IF cChar = cCAN
   ? "Terminated by remote station!"
  ELSE
    IF nErrCount > 10                      // transferred successfully?
     ? "too many errors, cancelled !"
    ELSE
     Tx_Ok := .T.
    ENDIF
   COM_SEND(nPort, cEOT)                   // End of transmission
  ENDIF
 ENDIF

 FCLOSE(nHandle)

RETURN(Tx_Ok)

**************************************************************************
* Function : XMO_TRA1k()
* Returns  : LOGICAL returns .T., when the transfer has been successful
*
* Usage    : The function is able to transfer a file via XMODEM
*            protocol with a block length of 1024 Byte.
*            Both, CHECKSUM and CRC can be used. The test mode to be
*            used is determined automatically by the function and the
*            remote station.
*            The function has been tested with PROCOMM 2.4.2 on the
*            remote station and a direct cable connection.
*            In PROCOMM, XMODEM 1k is called YMODEM, which is not
*            correct!
*
* Notes   :  As the number of the remaining bytes is less than 128
*            bytes, the function uses the 'small' block. Therefore,
*            the receive function must be able to recognize the two
*            different block lengths. This corresponds to the
*            definition of XMODEM 1k and true YMODEM, but we can not
*            be sure if this will be handled correctly by any remote
*            stations. If there are any problems with the shortened
*            last block, the fourth parameter of XMOBLOCK() within
*            the following functions can be set 2 constantly.
*
* Example  :
*            nPort := 1
*            nBaudRate := 2400
*
*            COM_OPEN(nPort, 1000)
*            COM_INIT(nPort, nBaudrate, "N", 8, 1)
*
*             ACCEPT "Datei: " TO cFile
*             ? XMO_TRA1k(nPort, cFile)       // send file on Port 1
*
*            COM_CLOSE(nPort)
*
**************************************************************************
FUNCTION XMO_TRA1k(nPort, cFile)
LOCAL nBlockNum, nMaxBsize, nBlockType, cMessage, cBuffer, nHandle
LOCAL nErrCount := 0, Tx_Ok := .F., lChkMode := .F., nReadData := 0

 CSETREF(.T.)                              // highest performance

* transfer variable
*
 nBlockNum  := 1
 nMaxBsize  := 1024
 cBuffer    := REPLICATE(cEOF, nMaxBsize)  // data input buffer
 nHandle    := FOPEN(cFile, 0)             // read only

 * Wait for begin of transfer (or termination)
 * When receiving character "C" empfangen, the transfer of the first
 * block starts immediately; the CRC protocol required by the remote
 * station is accepted.

 WAITPERIOD(6000)                          // loop for max. 60 seconds
 cChar := COM_READ(nPort, 1)               // read one character

 DO WHILE .NOT. cChar $ "C"+cNAK+cCAN .AND. WAITPERIOD()
   cChar := COM_READ(nPort, 1)             // read one character
 ENDDO

 lChkMode := (cChar = "C")                 // CRC (.T.), Checksum (.F.)

 MILLISEC(100)                             // wait 0.1 seconds

 IF cChar = cNAK .OR. cChar = "C"          // answer of remote station
  nReadData := FREAD(nHandle, @cBuffer, nMaxBsize)
  cMessage  := "Sending Block:  "

  DO WHILE nReadData > 0 .AND. nErrCount <= 10 .AND. cChar <> cCAN
   ? cMessage + NTOC(nBlockNum, 10, 4)

   nBlockType := IF(nReadData <= 128, 1, 2)
   COM_SEND(nPort, XMOBLOCK(cBuffer, nBlockNum, lChkMode, nBlockType))

   * wait for block confirmation
   *
    WAITPERIOD(1000)                       // loop for max. 10 seconds
    cChar := COM_READ(nPort, 1)            // read one character

    DO WHILE .NOT. cChar $ cACK+cCAN .AND. WAITPERIOD()
     cChar := COM_READ(nPort, 1)           // read one character
    ENDDO

    IF cChar = cACK
     ?? " - verified ", if(nBlockType = 1, " 128 Byte", " 1024 Byte")
     ++nBlockNum
     nErrCount := 0
     RANGEREPL(0, 255, @cBuffer, CHR(26))
     nReadData := FREAD(nHandle, @cBuffer, nMaxBsize)

     cMessage := "Sending Block:  "
    ELSE
     ++nErrCount                           // at NAK or timeout -> again
     cMessage := "Resending block:"
    ENDIF
  ENDDO

  IF cChar = cCAN
   ? "Cancelled by remote station!"
  ELSE
    IF nErrCount > 10                      // transferred successfully
     ? "Too many errors, cancelled !"
    ELSE
     Tx_Ok := .T.
    ENDIF
   COM_SEND(nPort, cEOT)                   // End of transmission
  ENDIF
 ENDIF

 FCLOSE(nHandle)

RETURN(Tx_Ok)


**************************************************************************
*
* Function : XMO_REC()
* Returns  : LOGICAL returns .T., when transfer has been successful.
*
* Usage    : The function is able to receive a file, transmitted via
*            XMODEM protocol.
*            The function can handle CHECKSUM as well as CRC, and block
*            lengths of 128 and 1024 bytes (XMODEM 1k). The test mode to
*            be used is determined automatically by the function and the
*            remote station. The block length received will be
*            recognized by XMOCHECK() in the first character of the
*            block header.
*            The function has been tested with PROCOMM 2.4.2 on the
*            remote station and a direct cable connection.
*
* Example  :
*            nPort := 1
*            nBaudRate := 2400
*
*            COM_OPEN(nPort, 1000)
*            COM_INIT(nPort, nBaudrate, "N", 8, 1)
*
*             ACCEPT "File : " TO cFile
*
*             IF FILE(cFile)
*              * !! WARNING, FILE ALREADY EXISTS !!
*             ELSE
*              ? Xmo1k_Rx(nPort, cFile)    // receive file on port 1
*             ENDIF
*
*            COM_CLOSE(nPort)
*
**************************************************************************
FUNCTION XMO_REC(nPort, cFile)
LOCAL lEOT := .F., nBlockSize := 128, nBlockNum := 1
LOCAL cBlock, nHandle, nNoRxCnt, lCRCMode := .F., nErrCount := 0

 COM_FLUSH(nPort)            // clear buffer
 nHandle := FCREATE(cFile)   // create / overwrite file

 IF nHandle = -1
  * send 5 CAN, if file can not be opened or created.
  *
   COM_SEND(nPort, REPLICATE(cCan, 5))
 ELSE
  *
  * To check, if the remote station is able to transfer with CRC,
  * character "C" is sent three times with a one second delay.
  * If the remote station begins the transmission directly, CRC
  * protocol is accepted; if not, another character has to be sent
  * to start a normal transmission without CRC.
  *
  Counter := 3
  DO WHILE Counter > 0 .AND. COM_COUNT(nPort) = 0
   COM_SEND(nPort, "C")
   --Counter
   WAITPERIOD(100)
   DO WHILE COM_COUNT(nPort) = 0 .AND. WAITPERIOD()
   ENDDO
  ENDDO

  IF COM_READ(nPort, 1) $ cSOH+cSTX  // remote station sends block?
   lCRCMode  := .T.                  // CRC protocol on
   nSmallBlk := 133
   nLargeBlk := 1029
   nOvrHead  := 5
  ELSE
   lCRCMode  := .F.                  // Checksum protocol on
   nSmallBlk := 132
   nLargeBlk := 1028
   nOvrHead  := 4
   COM_SEND(nPort, cNAK)             // send start character
  ENDIF

  nKey := 0
  lEOT := .F.

  DO WHILE !lEOT .AND. DCD() .AND. nKey <> nCanKey .AND. nErrCount <= 10
   WAITPERIOD(100)
   DO WHILE WAITPERIOD() .AND. DCD() .AND.;
                               COM_COUNT(nPort) = 0 .AND. nKey <> nCanKey
    nKey := INKEY()
   ENDDO

   * With the first received character the block type is determined!
   *
    x := y := COM_COUNT(nPort)

    IF x > 0
     nBlocksize := IF(COM_READ(nPort, 1, .T.) = cSTX, nLargeBlk, nSmallBlk)
    ENDIF

   nNoRxCnt := 1
   DO WHILE nNoRxCnt <= 10 .AND. DCD() .AND.;
                                  x < nBlockSize .AND. nKey <> nCanKey

    WAITPERIOD(100)
    DO WHILE WAITPERIOD() .AND. DCD() .AND.;
                                x < nBlockSize .AND. nKey <> nCanKey

     x := COM_COUNT(nPort)

     IF x > y
      y := x
      WAITPERIOD(100)
     ENDIF
      nKey := INKEY()
    ENDDO

    ++nNoRxCnt
   ENDDO

   IF DCD()
    IF nKey = nCanKey
     COM_SEND(nPort, REPLICATE(cCan, 5))
    ELSE
     cBlock = COM_READ(nPort, nBlockSize)
      IF LEN(cBlock) = nBlockSize
       IF XMOCHECK(cBlock, lCRCMode) = nBlockNum     // block size
                                                     // recognized
        COM_SEND(nPort, cACK)
        FWRITE(nHandle, SUBSTR(cBlock, 4, nBlockSize -nOvrHead))
        ? "Received block:", NTOC(nBlockNum, 10, 5), LEN(cBlock), FERROR()

        strfile(cBlock, ntoc(nBlockNum))

        nBlockNum := NUMAND(255, nBlockNum +1)
        nErrCount := 0
       ELSE
        COM_SEND(nPort, cNAK)
        ++nErrCount
       ENDIF
      ELSE
      IF cBlock == cEOT
       COM_SEND(nPort, cACK)
       lEOT := .T.
      ELSE
       COM_SEND(nPort, cNAK)
       ++nErrCount
      ENDIF
     ENDIF
    ENDIF
   ENDIF
  ENDDO

  FCLOSE(nHandle)

  IF !lEOT
    * Transmission failed!
    DELETEFILE(cFile)
  ENDIF
 ENDIF
RETURN(lEOT)

*---------------------------------------------------------------------------
*
* Data Carrier Detect
* The function checks, if a connection is still active.
* If, for example, a modem is used, the check of the Data Carrier
* Signal would be carried out with COM_DCD().
* At a direct local connection, the function always returns .T., as
* the Data Carrier Signal will not be checked. Within this function,
* however, any other kind of check can be done.
*
FUNCTION DCD
 *
 * Any kind of checking can take place.
 *
RETURN(.T.)

Attribute VB_Name = "MdlDwgPreview"
''BitMap Api
''This class was writen by http://www.vbdesign.net/cadpages/features/ good articles

Private Type BITMAPINFOHEADER
  biSize As Long
  biWidth As Long
  biHeight As Long
  biPlanes As Integer
  biBitCount As Integer
  biCompression As Long
  biSizeImage As Long
  biXPelsPerMeter As Long
  biYPelsPerMeter As Long
  biClrUsed As Long
  biClrImportant As Long
End Type


Private Type RGBQUAD
  rgbBlue As Byte
  rgbGreen As Byte
  rgbRed As Byte
  rgbReserved As Byte
End Type

Private Type IMGREC
  bytType As Byte
  lngStart As Long
  lngLen As Long
End Type

Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Public Function PaintPreview(ByRef strFile As String, PicBox As PictureBox) As Integer
  Dim lngSeeker As Long
  Dim lngImgLoc As Long
  Dim bytCnt As Byte
  Dim lngFile As Long
  Dim lngCurLoc As Long
  Dim intCnt As Integer
  Dim udtRec As IMGREC
  Dim bytBMPBuff() As Byte
  Dim udtColors() As RGBQUAD
  Dim udtColor As RGBQUAD
  Dim lngHwnd As Long
  Dim lngDc As Long
  Dim lngY As Long
  Dim lngX As Long
  Dim intRed As Integer
  Dim intGreen As Integer
  Dim intBlue As Integer
  Dim lngColor As Long
  Dim lngCnt As Long
  Dim udtHeader As BITMAPINFOHEADER
  
  On Error GoTo Err_Control
  If Len(Dir(strFile)) > 0 Then
    lngFile = FreeFile
    Open strFile For Binary As lngFile
    Seek lngFile, 14
    Get lngFile, , lngImgLoc
    Seek lngFile, lngImgLoc + 17
    lngCurLoc = Seek(lngFile)
    Seek lngFile, lngCurLoc + 4
    Get lngFile, , bytCnt
    If bytCnt > 1 Then
      For intCnt = 1 To bytCnt
        Get lngFile, , udtRec
        If udtRec.bytType = 2 Then
        
        'Now we begin the color extraction
        'The start value is the BYTE BEFORE
        'The BMP Header data (The RGBQUAD
        'And BMP Header are contained within
        'Another structure), so move the read/
        'Write marker to the next byte...
          Seek lngFile, udtRec.lngStart + 1
          'Pull out the BMP header data...
          Get lngFile, , udtHeader
          'Resize the Byte buffer to the full
          'Length of the data...
          ReDim bytBMPBuff(udtRec.lngLen)
          'Did you read Randall's article?
          If udtHeader.biBitCount = 8 Then
            'Resize the array of RGBQuad's, I
            'Could also have used the biClrUsed
            'Value of the udtHeader...
            ReDim udtColors(256)
            'Grab all of the color values
            Get lngFile, , udtColors
            'Now we grab the full record by
            'Moving the Read/Write marker
            'Back to the start of the data.
            'Don't worry about all of the data
            'We already grabbed...
            '(If you read Randall's article,
            'Remember that the data is reverse
            'Scan...
            Seek lngFile, udtRec.lngStart
            'Fill the buffer...
            Get lngFile, , bytBMPBuff
            lngDc = GetDC(PicBox.hWnd)
            
            'Clean any old paint off..
            PicBox.Cls
            'Begin Painting
            For lngY = 1 To udtHeader.biHeight
              For lngX = udtHeader.biWidth To _
              1 Step -1
                'See, we are reading the data
                'From THE END of the buffer...
                lngColor = _
                bytBMPBuff((UBound(bytBMPBuff) _
                - lngCnt))
                'Get the mapped value
                udtColor = udtColors(lngColor)
                'Break it into Red
                intRed = CInt(udtColor.rgbRed)
                'Green
                intGreen = CInt(udtColor.rgbGreen)
                'And Blue
                intBlue = CInt(udtColor.rgbBlue)
                'Get a color the API will accept
                lngColor = RGB(intRed, intGreen, _
                intBlue)
                'Paint this Pixel. The + 5 is to
                'Give a little offset from the edge
                'Of the form.
                'But before we do, would you like
                'To have Black backgrounds? Easy,
                'Swap the map:
                '///BLACK BACKGROUND///
                If lngColor = vbBlack Then
                  lngColor = vbWhite
                ElseIf lngColor = vbWhite Then
                  lngColor = vbBlack
                End If
                '//////////////////////
                'If your prefer White (the true
                'Value) Then just remove that..
                SetPixel lngDc, lngX + 5, lngY + 5, _
                lngColor
                'Increment the counter...
                lngCnt = lngCnt + 1
              Next lngX
            Next lngY
          End If
          Exit For
        ElseIf udtRec.bytType = 3 Then
          'Add your message f
          Exit For
        End If
      Next intCnt
    Else
      PicBox.Cls
      PicBox.Print "No Preview"
    End If
    'Close the file
    Close lngFile
    'Return the value
  End If
  'General Error control
Exit_Here:
  Exit Function
Err_Control:
  Select Case Err.Number
  'Add your Case selections here
    Case Else
    MsgBox Err.Description
    Resume Exit_Here
  End Select
End Function

Public Function DwgVers(strFullPath _
As String) As String
  Dim strText As String * 6
  Dim intFile As Integer
  Dim strVers As String
  If Len(Dir(strFullPath)) > 0 Then
    intFile = FreeFile
    Open strFullPath For Random As intFile
    Get #intFile, , strText
    Close intFile
    strVers = strText
  End If
  
Select Case strVers
Case "AC1015"
    strVers = "R15"
Case "AC1013"
    strVers = "R14"
Case "AC1014"
    strVers = "R14"
Case "AC1010"
    strVers = "R13"
Case "AC1011"
    strVers = "R13"
Case "AC1012"
    strVers = "R13"
Case Else
    strVers = "Unknown"
End Select
   
  DwgVers = strVers
End Function
Public Function ImagePreviewType(strFile As _
String) As Integer
  Dim lngSeeker As Long
  Dim lngImgLoc As Long
  Dim bytCnt As Byte
  Dim lngFile As Long
  Dim lngCurLoc As Long
  Dim intCnt As Integer
  Dim intTemp As Integer
  Dim udtRec As IMGREC
  On Error GoTo Err_Control
  If Len(Dir(strFile)) > 0 Then
    lngFile = FreeFile
    Open strFile For Binary As lngFile
    Seek lngFile, 14
    Get lngFile, , lngImgLoc
    Seek lngFile, lngImgLoc + 17
    lngCurLoc = Seek(lngFile)
    Seek lngFile, lngCurLoc + 4
    Get lngFile, , bytCnt
    If bytCnt > 1 Then
      For intCnt = 1 To bytCnt
        Get lngFile, , udtRec
        If udtRec.bytType = 2 Then
          intTemp = udtRec.bytType
          Exit For
        ElseIf udtRec.bytType = 3 Then
          intTemp = udtRec.bytType
          Exit For
        End If
      Next intCnt
    Else
      intTemp = 1
    End If
    Close lngFile
    ImagePreviewType = intTemp
  End If
Exit_Here:
  Exit Function
Err_Control:
  Select Case Err.Number
    Case Else
    MsgBox Err.Description
    Resume Exit_Here
  End Select
End Function







VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPalette"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    ByVal lpvDest As Long, lpvSource As Any, ByVal cbCopy As Long)

Private Type SAFEARRAYBOUND
    cElements As Long
    lLbound As Long
End Type

Private Type SAFEARRAY2D
    cDims As Integer
    fFeatures As Integer
    cbElements As Long
    cLocks As Long
    pvData As Long
    Bounds(0 To 1) As SAFEARRAYBOUND
End Type

Private Declare Function VarPtrArray Lib "msvbvm50.dll" Alias _
   "VarPtr" (Ptr() As Any) As Long
Private Type tNode                              'Octree node struct
    bIsLeaf As Boolean                          'Leaf flag
    bAddedReduce As Boolean                     'Linked list flag
    vRed As Long                                'Red Value
    vGreen As Long                              'Green Value
    vBlue As Long                               'Blue Value
    cClrs As Long                               'Pixel count
    iChildren(0 To 1, 0 To 1, 0 To 1) As Long   'Child pointers
    iNext As Long                               'Next reducable node
End Type

Private aNodes() As tNode, cNodes As Long
Private nDepth As Byte, TopGarbage As Long
Private cClr As Long, aReduce(1 To 8) As Long
Private cOps As Long

Private Declare Function timeGetTime Lib "winmm.dll" () As Long

Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As Any, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, lpOverlapped As Any) As Long
Private Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, lpOverlapped As Any) As Long
Private Declare Function SetFilePointer Lib "kernel32" (ByVal hFile As Long, ByVal lDistanceToMove As Long, lpDistanceToMoveHigh As Long, ByVal dwMoveMethod As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Const INVALID_HANDLE_VALUE = -1
Private Const CREATE_ALWAYS = 2
Private Const OPEN_EXISTING = 3
Private Const OPEN_ALWAYS = 4
Private Const GENERIC_READ = &H80000000
Private Const GENERIC_WRITE = &H40000000
Private Const FILE_ATTRIBUTE_ARCHIVE = &H20
Private Const FILE_ATTRIBUTE_COMPRESSED = &H800
Private Const FILE_ATTRIBUTE_DIRECTORY = &H10
Private Const FILE_ATTRIBUTE_HIDDEN = &H2
Private Const FILE_ATTRIBUTE_NORMAL = &H80
Private Const FILE_ATTRIBUTE_READONLY = &H1
Private Const FILE_ATTRIBUTE_SYSTEM = &H4
Private Const FILE_ATTRIBUTE_TEMPORARY = &H100
Private Const FILE_BEGIN = 0

' make private storage in compatible
' with windows storage in memory/on disk
Private Type RGBQUAD
   Red As Byte
   Green As Byte
   Blue As Byte
   Reserved As Byte
End Type
Private m_tPal() As RGBQUAD
Private m_iPalette As Long

Public Event InitProgress(ByVal Max As Long)
Public Event Progress(ByVal lPosition As Long)
Public Event Complete(ByVal lTimeMs As Long)

Public Sub Add(ByVal Red As Byte, ByVal Green As Byte, ByVal Blue As Byte)
    m_iPalette = m_iPalette + 1
    ReDim Preserve m_tPal(1 To m_iPalette) As RGBQUAD
    m_tPal(m_iPalette).Red = Red
    m_tPal(m_iPalette).Green = Green
    m_tPal(m_iPalette).Blue = Blue
End Sub

Public Property Get ClosestIndex( _
      ByVal Red As Long, _
      ByVal Green As Long, _
      ByVal Blue As Long _
   ) As Long
Dim i As Long
Dim lER As Long, lEB As Long, lEG As Long
Dim lMinER As Long, lMinEB As Long, lMinEG As Long
Dim lMinIndex As Long
   
   lMinER = 255: lMinEB = 255: lMinEG = 255
   For i = 1 To m_iPalette
      With m_tPal(i)
         If (Red = .Red) And (Blue = .Blue) And (Green = .Green) Then
            ClosestIndex = i
            Exit Property
         Else
            lER = Abs(Red - .Red)
            lEB = Abs(Blue - .Blue)
            lEG = Abs(Green - .Green)
            If (lER + lEB + lEG < lMinER + lMinEB + lMinEG) Then
               lMinER = lER
               lMinEB = lEB
               lMinEG = lEG
               lMinIndex = i
            End If
         End If
      End With
   Next i
   
   ClosestIndex = lMinIndex
      
End Property

Public Property Get Count() As Long
' Number of colours in this palette:
   Count = m_iPalette
End Property

Public Property Get Red(ByVal iIndex) As Byte
' Red value for palette index
   Red = m_tPal(iIndex).Red
End Property

Public Property Get Green(ByVal iIndex) As Byte
' Green value for palette index
   Green = m_tPal(iIndex).Green
End Property

Public Property Get Blue(ByVal iIndex) As Byte
' Blue value for palette index
   Blue = m_tPal(iIndex).Blue
End Property

'***************************************************************
'CreateOptimalPalette--Creates an optimal palette with the
'  specified number of colors using octree quantisation
'
'Inputs:
'  cDib
'  DIBSection for which palette is made
'
'  nLevels
'  Number of levels of color to create
'
'  PalSize
'  Size of palette to make
'
'Date Created:7/12/1999          Brian Schimpf
'
'Modifications:                                     Date:
'***************************************************************
Public Sub CreateOptimal( _
      ByRef cdib As cDIBSection, _
      Optional ByVal PalSize As Long = 236, _
      Optional ByVal nLevels As Long = 6 _
   )
Dim tSA As SAFEARRAY2D, aDib() As Byte
Dim x As Long, y As Long
Dim xMax As Long, yMax As Long
Dim sTime As Long, pPal As Double
Dim Inc As Double

   'Creates a VB array without copying data
   With tSA
       .cbElements = 1
       .cDims = 2
       .Bounds(0).lLbound = 0
       .Bounds(0).cElements = cdib.Height
       .Bounds(1).lLbound = 0
       .Bounds(1).cElements = cdib.BytesPerScanLine
       .pvData = cdib.DIBSectionBitsPtr
   End With
   CopyMemory VarPtrArray(aDib), VarPtr(tSA), 4
   
   xMax = (cdib.Width - 1) * 3
   Debug.Print "hptr = " & Hex(cdib.DIBSectionBitsPtr)
   Debug.Print "xMax = " & xMax
   Debug.Print "byte = " & cdib.BytesPerScanLine
   Debug.Print "ubound2 = " & UBound(aDib, 2)
   Debug.Print "ubound1 = " & UBound(aDib, 1)
   
   yMax = cdib.Height - 1
   
   nDepth = nLevels
   
   sTime = timeGetTime
   
   ReDim aNodes(1 To 50) 'Allocates initial storage
   TopGarbage = 0
   cNodes = 1
   cClr = 0
   
   RaiseEvent InitProgress(yMax + PalSize)
   
   For y = 0 To yMax
      For x = 0 To xMax
         'Adds the current pixel to the color octree
          AddClr 1, 1, 0, 255, 0, 255, 0, 255, _
            aDib(x + 2, y), aDib(x + 1, y), aDib(x, y)
         
         pPal = -1
         Do While cClr > PalSize
            'Combine the levels to get down to desired palette size
            pPal = CombineNodes
            If pPal = False Then Exit Do
         Loop
         If pPal = False Then
            If MsgBox("Error in reducing nodes", vbYesNo) = vbNo _
               Then Stop
         End If
      Next x
      RaiseEvent Progress(y)
   Next y
   
   Clear
   Debug.Print "Colors: " & cClr
   pPal = 0
   Inc = pPal / cNodes
   'Go through octree and extract colors
   For y = 1 To UBound(aNodes)
      If aNodes(y).bIsLeaf Then
         With aNodes(y)
            pPal = pPal + 1
            Add .vRed / .cClrs, .vGreen / .cClrs, .vBlue / .cClrs
            RaiseEvent Progress(y)
         End With
      End If
   Next y
   
   Debug.Print pPal
   
   RaiseEvent Complete(ByVal (timeGetTime - sTime))

End Sub

'***************************************************************
'                         --Recursive--
'AddClr--Adds a color to the OctTree palette
'  Will call itself if not in correct level
'
'Inputs:
'  iBranch                       nLevel
'  Branch to look down           Current level(depth) in tree
'
'  vMin(R, G, B)                 vMax(R, G, B)
'  The minimum branch value      The maximum branch value
'
'  R, G, B
'  The Red, Green, and Blue color components
'
'Date Created: 7/12/1999         Brian Schimpf
'
'Modifications:                                     Date:
'***************************************************************
Private Sub AddClr(ByVal iBranch As Long, ByVal nLevel As Long, _
      ByVal vMinR As Byte, ByVal vMaxR As Byte, _
      ByVal vMinG As Byte, ByVal vMaxG As Byte, _
      ByVal vMinB As Byte, ByVal vMaxB As Byte, _
      ByVal r As Byte, ByVal g As Byte, ByVal b As Byte)
        
Dim iR As Byte, iG As Byte, iB As Byte
Dim vMid As Long, iIndex As Long

   
   'Find mid values for colors and decide which path to take
   'Also update max and min values for later call to self
   vMid = vMinR / 2 + vMaxR / 2
   If r > vMid Then
      iR = 1
      vMinR = vMid
   Else
      iR = 0
      vMaxR = vMid
   End If
   
   vMid = vMinG / 2 + vMaxG / 2
   If g > vMid Then
      iG = 1
      vMinG = vMid
   Else
      iG = 0
      vMaxG = vMid
   End If
   
   vMid = vMinB / 2 + vMaxB / 2
   If b > vMid Then
      iB = 1
      vMinB = vMid
   Else
      iB = 0
      vMaxB = vMid
   End If
   
   If aNodes(iBranch).iChildren(iR, iG, iB) = 0 Then  'If no child here then...
      iIndex = GetFreeNode 'Get a new node index
      aNodes(iBranch).iChildren(iR, iG, iB) = iIndex
      aNodes(iBranch).cClrs = aNodes(iBranch).cClrs + 1
      With aNodes(iIndex)   'Clear/set data
         .bIsLeaf = (nLevel = nDepth): .vGreen = 0
         .iNext = 0: .cClrs = 0: .vBlue = 0: .vRed = 0
      End With
   Else  'Has a child here
      iIndex = aNodes(iBranch).iChildren(iR, iG, iB)
   End If
   
   
   If aNodes(iIndex).bIsLeaf Then  'If it is a leaf
      With aNodes(iIndex)
         If .cClrs = 0 Then cClr = cClr + 1
         .cClrs = .cClrs + 1
         .vRed = .vRed + r
         .vGreen = .vGreen + g
         .vBlue = .vBlue + b
      End With
   Else  'Otherwise
      With aNodes(iIndex)
         If .bAddedReduce = False Then    'If 2 or more colors, add to reducable aNodes list
            .iNext = aReduce(nLevel)
            aReduce(nLevel) = iIndex
            .bAddedReduce = True
            'Debug.Print "Reducable Added: " & iIndex, "Level = " & nLevel, "Top was = " & .iNext
         End If
      End With
      
      'Search a level deeper
      AddClr iIndex, nLevel + 1, vMinR, vMaxR, vMinG, vMaxG, _
         vMinB, vMaxB, r, g, b
   End If
End Sub

'***************************************************************
'CombineNodes--Combines octree aNodes to reduce the count of
'  colors.  Combines all children of a leaf into itself.
'
'Date Created: 7/13/1999         Brian Schimpf
'
'Modifications:                                     Date:
'***************************************************************
Private Function CombineNodes() As Boolean
Dim i As Long, iIndex As Long
Dim iR As Byte, iG As Byte, iB As Byte
Dim nR As Long, nG As Long, nB As Long, nPixel As Long
   
   
   
   For i = nDepth To 1 Step -1   'Find deepest reducable level
      If aReduce(i) <> 0 Then Exit For
   Next i
   
   'Debug.Print "aReduce Index = " & i
   
   If i = 0 Then Exit Function
   iIndex = aReduce(i)
   aReduce(i) = aNodes(iIndex).iNext
   
   For i = 0 To 7
      'iR = IIf((i And 1) = 1, 1, 0)
      If (i And 1) = 1 Then
         iR = 1
      Else
         iR = 0
      End If
      'iG = IIf((i And 2) = 2, 1, 0)
      If (i And 2) = 2 Then
         iG = 1
      Else
         iG = 0
      End If
      'iB = IIf((i And 4) = 4, 1, 0)
      If (i And 4) = 4 Then
         iB = 1
      Else
         iB = 0
      End If
      
      If aNodes(iIndex).iChildren(iR, iG, iB) <> 0 Then 'If there is a child
         With aNodes(aNodes(iIndex).iChildren(iR, iG, iB))
            'Add red, green, blue, and pixel count to running total
            nR = nR + .vRed
            nG = nG + .vGreen
            nB = nB + .vBlue
            nPixel = nPixel + .cClrs
            FreeNode aNodes(iIndex).iChildren(iR, iG, iB) 'Free the node
            cClr = cClr - 1
         End With
         aNodes(iIndex).iChildren(iR, iG, iB) = 0 'Clear the link
      End If
   Next i
   
   cClr = cClr + 1
   
   'Set the new node data
   With aNodes(iIndex)
      .cClrs = nPixel
      .bIsLeaf = True
      .vRed = nR
      .vBlue = nB
      .vGreen = nG
   End With
   CombineNodes = True
   
End Function

Public Sub EraseNodes()
  Erase aNodes
  TopGarbage = 0
End Sub

'***************************************************************
'FreeNode--Puts a node on the top of the garbage list
'
'Inputs:
'  iNode
'  Index of node to free
'
'Date Creates:7/13/1999          Brian Schimpf
'
'Modifications:                                     Date:
'***************************************************************
Private Sub FreeNode(ByVal iNode As Long)
   aNodes(iNode).iNext = TopGarbage
   TopGarbage = iNode
   aNodes(iNode).bIsLeaf = False 'Necessary for final loop through
   aNodes(iNode).bAddedReduce = False
   cNodes = cNodes - 1
End Sub

'***************************************************************
'GetFreeNode--Gets a new node index from the trash list or the
'  end of the list.  Clears child pointers
'
'Outputs:
'  Node Index
'
'Date Created:7/12/1999          Brian Schimpf
'
'Modifications:                                     Date:
'***************************************************************
Private Function GetFreeNode() As Long
Dim i As Long, iR As Byte, iG As Byte, iB As Byte
   cNodes = cNodes + 1
   If TopGarbage = 0 Then
      If cNodes > UBound(aNodes) Then
         i = cNodes * 1.1
         ReDim Preserve aNodes(1 To i)
      End If
      GetFreeNode = cNodes
   Else
      GetFreeNode = TopGarbage
      TopGarbage = aNodes(TopGarbage).iNext
      For i = 0 To 7
         'iR = IIf((i And 1) = 1, 1, 0)
         If (i And 1) = 1 Then
            iR = 1
         Else
            iR = 0
         End If
         'iG = IIf((i And 2) = 2, 1, 0)
         If (i And 2) = 2 Then
            iG = 1
         Else
            iG = 0
         End If
         'iB = IIf((i And 4) = 4, 1, 0)
         If (i And 4) = 4 Then
            iB = 1
         Else
            iB = 0
         End If
         aNodes(GetFreeNode).iChildren(iR, iG, iB) = 0
      Next i
   End If
End Function

Public Sub Clear()
   Erase m_tPal
   m_iPalette = 0
End Sub

Public Function Load(ByVal sFileName As String, Optional ByRef lPosition As Long) As Boolean
Dim hFile As Long
Dim lSize As Long
Dim lBytesRead As Long

   hFile = CreateFile(sFileName, _
                  GENERIC_READ, _
                  ByVal 0&, _
                  ByVal 0&, _
                  OPEN_EXISTING, _
                  FILE_ATTRIBUTE_NORMAL, _
                  ByVal 0&)
   If Not hFile = INVALID_HANDLE_VALUE Then
      If lPosition > 0 Then
         SetFilePointer hFile, lPosition, ByVal 0&, FILE_BEGIN
      Else
         lPosition = 0
      End If
      ReadFile hFile, lSize, 4, lBytesRead, ByVal 0&
      If lBytesRead = 4 Then
         If lSize > 0 And lSize <= 4097 Then
            m_iPalette = lSize
            ReDim m_tPal(1 To m_iPalette) As RGBQUAD
            lSize = LenB(m_tPal(1)) * m_iPalette
            ReadFile hFile, m_tPal(1), lSize, lBytesRead, ByVal 0&
            If lSize = lBytesRead Then
               lPosition = lPosition + 4 + lBytesRead
               Load = True
            End If
         End If
      End If
      CloseHandle hFile
   End If

End Function

Public Function Save(ByVal sFileName As String, Optional ByVal lPosition As Long) As Boolean
Dim hFile As Long
Dim lSize As Long
Dim lBytesWritten As Long

   hFile = CreateFile(sFileName, _
                  GENERIC_WRITE, _
                  ByVal 0&, _
                  ByVal 0&, _
                  CREATE_ALWAYS, _
                  FILE_ATTRIBUTE_NORMAL, _
                  ByVal 0&)
   If Not hFile = INVALID_HANDLE_VALUE Then
      If lPosition > 0 Then
         SetFilePointer hFile, lPosition, 0, FILE_BEGIN
      Else
         lPosition = 0
      End If
      WriteFile hFile, m_iPalette, 4, lBytesWritten, ByVal 0&
      If lBytesWritten = 4 Then
         lSize = LenB(m_tPal(1)) * m_iPalette
         WriteFile hFile, m_tPal(1), lSize, lBytesWritten, ByVal 0&
         If lSize = lBytesWritten Then
            lPosition = lPosition + 4 + lBytesWritten
            Save = True
         End If
      End If
      CloseHandle hFile
   End If
   
End Function

Private Sub SplitDelimitedString( _
        ByVal sString As String, _
        ByVal sDelim As String, _
        ByRef sValues() As String, _
        ByRef iCount As Long _
    )
' ==================================================================
' Splits sString into an array of parts which are
' delimited in the string by sDelim.  The array is
' indexed 1-iCount where iCount is the number of
' items.  If no items found iCount=1 and the array has
' one element, the original string.
'   sString : String to split
'   sDelim  : Delimiter
'   sValues : Return array of values
'   iCount  : Number of items returned in sValues()
' ==================================================================
Dim iPos As Long
Dim iNextPos As Long
Dim iDelimLen As Long

    iCount = 0
    Erase sValues
    iDelimLen = Len(sDelim)
    iPos = 1
    iNextPos = InStr(sString, sDelim)
    Do While iNextPos > 0
        iCount = iCount + 1
        ReDim Preserve sValues(1 To iCount) As String
        sValues(iCount) = Mid$(sString, iPos, (iNextPos - iPos))
        iPos = iNextPos + iDelimLen
        iNextPos = InStr(iPos, sString, sDelim)
    Loop
    iCount = iCount + 1
    ReDim Preserve sValues(1 To iCount) As String
    sValues(iCount) = Mid$(sString, iPos)
End Sub
Public Sub CreateWebSafe()
Dim lIndex As Long
Dim r As Long, g As Long, b As Long
Dim l As Long, i As Long

   ' Websafe (IE 216 colour) palette
   m_iPalette = 256
   ReDim m_tPal(1 To 256) As RGBQUAD
   p16ColourLow8 1
   p16ColourHigh8 248
   lIndex = 8
   For b = 0 To &HFF Step &H33
      For g = 0 To &HFF Step &H33
         For r = 0 To &HFF Step &H33
            ' ignore if the output is any combination of 0 and FF
            l = r + g + b
            If l = 0 Or l = &H2FD Then
               ' ignore
            ElseIf l = &H1FE And (r = 0 Or g = 0 Or b = 0) Then
               ' ignore
            ElseIf l = &HFF And ((r = 0 And g = 0) Or (r = 0 And b = 0) Or (g = 0 And b = 0)) Then
               ' ignore
            Else
               ' add
               lIndex = lIndex + 1
               With m_tPal(lIndex)
                  .Red = r: .Green = g: .Blue = b
               End With
            End If
         Next
      Next
   Next
   ' Fill the remain entries with gray shades:
   r = 8: g = 8: b = 8
   For i = 217 To 247
      With m_tPal(lIndex)
         .Red = r: .Green = g: .Blue = b
         r = r + 8: g = g + 8: b = b + 8
      End With
   Next i
End Sub

Private Sub p16ColourLow8(ByVal lStartIndex As Long)
   lStartIndex = lStartIndex - 1
   With m_tPal(lStartIndex + 2)
      .Red = &H80: .Green = 0: .Blue = 0
   End With
   With m_tPal(lStartIndex + 3)
      .Red = 0: .Green = &H80: .Blue = 0
   End With
   With m_tPal(lStartIndex + 4)
      .Red = &H80: .Green = &H80: .Blue = 0
   End With
   With m_tPal(lStartIndex + 5)
      .Red = 0: .Green = 0: .Blue = &H80
   End With
   With m_tPal(lStartIndex + 6)
      .Red = &H80: .Green = 0: .Blue = &H80
   End With
   With m_tPal(lStartIndex + 7)
      .Red = 0: .Green = &H80: .Blue = &H80
   End With
   With m_tPal(lStartIndex + 8)
      .Red = &HC0: .Green = &HC0: .Blue = &HC0
   End With

End Sub
Private Sub p16ColourHigh8(ByVal lStartIndex As Long)
   lStartIndex = lStartIndex - 9
   With m_tPal(lStartIndex + 9)
      .Red = &H80: .Green = &H80: .Blue = &H80
   End With
   With m_tPal(lStartIndex + 10)
      .Red = &HFF: .Green = 0: .Blue = 0
   End With
   With m_tPal(lStartIndex + 11)
      .Red = 0: .Green = &HFF: .Blue = 0
   End With
   With m_tPal(lStartIndex + 12)
      .Red = &HFF: .Green = &HFF: .Blue = 0
   End With
   With m_tPal(lStartIndex + 13)
      .Red = 0: .Green = 0: .Blue = &HFF
   End With
   With m_tPal(lStartIndex + 14)
      .Red = &HFF: .Green = 0: .Blue = &HFF
   End With
   With m_tPal(lStartIndex + 15)
      .Red = 0: .Green = &HFF: .Blue = &HFF
   End With
   With m_tPal(lStartIndex + 16)
      .Red = &HFF: .Green = &HFF: .Blue = &HFF
   End With
End Sub


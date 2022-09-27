VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form1 
   Caption         =   "Image Encode Compression"
   ClientHeight    =   10695
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   13470
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   10695
   ScaleWidth      =   13470
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text3 
      Height          =   372
      Left            =   9090
      TabIndex        =   31
      Text            =   "3"
      Top             =   7455
      Width           =   492
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Sign"
      Height          =   255
      Left            =   10395
      TabIndex        =   30
      Top             =   7425
      Width           =   855
   End
   Begin VB.PictureBox Picture6 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   2052
      Left            =   5175
      ScaleHeight     =   133
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   125
      TabIndex        =   29
      Top             =   3480
      Width           =   1932
   End
   Begin VB.PictureBox Picture5 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   2052
      Left            =   2850
      ScaleHeight     =   133
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   125
      TabIndex        =   28
      Top             =   3480
      Width           =   1932
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Save Picture"
      Height          =   372
      Left            =   9480
      TabIndex        =   27
      Top             =   10080
      Width           =   1452
   End
   Begin VB.Frame Frame6 
      Caption         =   "Save"
      Height          =   852
      Left            =   9000
      TabIndex        =   26
      Top             =   9720
      Width           =   2292
   End
   Begin VB.TextBox Text2 
      Height          =   372
      Left            =   9360
      TabIndex        =   25
      Text            =   "3"
      Top             =   3720
      Width           =   492
   End
   Begin VB.TextBox Text1 
      Height          =   288
      Left            =   10200
      TabIndex        =   22
      Text            =   "8"
      Top             =   2280
      Width           =   372
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   4920
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox Picture4 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   2052
      Left            =   240
      ScaleHeight     =   133
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   125
      TabIndex        =   21
      Top             =   3480
      Width           =   1932
   End
   Begin VB.PictureBox Picture3 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   2052
      Left            =   5205
      ScaleHeight     =   133
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   125
      TabIndex        =   20
      Top             =   840
      Width           =   1932
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   2052
      Left            =   2880
      ScaleHeight     =   133
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   125
      TabIndex        =   19
      Top             =   855
      Width           =   1932
   End
   Begin VB.CommandButton zigzagscan 
      Caption         =   "ZigZag"
      Height          =   372
      Left            =   9720
      TabIndex        =   18
      Top             =   5160
      Width           =   972
   End
   Begin VB.Frame zigzagframe 
      Caption         =   "ZigZag Scan"
      Height          =   972
      Left            =   9000
      TabIndex        =   17
      Top             =   4680
      Width           =   2292
   End
   Begin VB.CommandButton CHuffmanCoding 
      Caption         =   "Huffman Encoder"
      Height          =   492
      Left            =   9720
      TabIndex        =   16
      Top             =   8880
      Width           =   972
   End
   Begin VB.Frame Huffman 
      Caption         =   "Huffman Coding"
      Height          =   972
      Left            =   9000
      TabIndex        =   15
      Top             =   8520
      Width           =   2292
   End
   Begin VB.CommandButton ShiftCoding 
      Caption         =   "Shift Code"
      Height          =   492
      Left            =   9720
      TabIndex        =   14
      Top             =   7665
      Width           =   972
   End
   Begin VB.Frame Frame5 
      Caption         =   "Shift Coding"
      Height          =   1410
      Left            =   8985
      TabIndex        =   13
      Top             =   7110
      Width           =   2292
   End
   Begin VB.CommandButton encoding 
      Caption         =   "Encode"
      Height          =   372
      Left            =   9720
      TabIndex        =   12
      Top             =   6600
      Width           =   972
   End
   Begin VB.OptionButton Option5 
      Caption         =   "RLE"
      Height          =   372
      Left            =   10320
      TabIndex        =   11
      Top             =   6120
      Width           =   732
   End
   Begin VB.OptionButton Option4 
      Caption         =   "DPCM"
      Height          =   372
      Left            =   9360
      TabIndex        =   10
      Top             =   6120
      Width           =   852
   End
   Begin VB.Frame Frame4 
      Caption         =   "Encoding"
      Height          =   1212
      Left            =   9000
      TabIndex        =   9
      Top             =   5880
      Width           =   2292
   End
   Begin VB.CommandButton quantization 
      Caption         =   "Quantize"
      Height          =   372
      Left            =   10200
      TabIndex        =   8
      Top             =   3720
      Width           =   972
   End
   Begin VB.Frame Frame3 
      Caption         =   "Quantization"
      Height          =   1092
      Left            =   9000
      TabIndex        =   7
      Top             =   3360
      Width           =   2292
   End
   Begin VB.CommandButton DCT_Tran 
      Caption         =   "Transform"
      Height          =   372
      Left            =   9600
      TabIndex        =   6
      Top             =   2640
      Width           =   972
   End
   Begin VB.CommandButton ColorTransform 
      Caption         =   "Transform"
      Height          =   372
      Left            =   9600
      TabIndex        =   5
      Top             =   1200
      Width           =   972
   End
   Begin VB.OptionButton Option1 
      Caption         =   "RGB to YUV"
      Height          =   372
      Left            =   9240
      TabIndex        =   4
      Top             =   840
      Value           =   -1  'True
      Width           =   1452
   End
   Begin VB.Frame Frame2 
      Caption         =   "Color Transform"
      Height          =   1092
      Left            =   9000
      TabIndex        =   2
      Top             =   600
      Width           =   2292
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   2052
      Left            =   240
      Negotiate       =   -1  'True
      ScaleHeight     =   133
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   125
      TabIndex        =   0
      Top             =   840
      Width           =   1932
   End
   Begin VB.Frame Frame1 
      Height          =   9972
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   8292
   End
   Begin VB.Frame DCT 
      Caption         =   "DCT Tranform"
      Height          =   1332
      Left            =   9000
      TabIndex        =   3
      Top             =   1920
      Width           =   2292
      Begin VB.Label Label2 
         Caption         =   "Block Size"
         Height          =   252
         Left            =   360
         TabIndex        =   24
         Top             =   360
         Width           =   1092
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   372
      Left            =   5520
      TabIndex        =   23
      Top             =   4800
      Width           =   972
   End
   Begin VB.Menu file 
      Caption         =   "File"
      Begin VB.Menu open 
         Caption         =   "Open"
      End
      Begin VB.Menu save 
         Caption         =   "Save"
      End
      Begin VB.Menu exit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu help 
      Caption         =   "Help"
      Begin VB.Menu about 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'for color transform
Dim Wid, Hgt
Dim red(), green(), blue()
Dim red2(), green2(), blue2()
Dim YY(), UU(), Vv()
'for DCT
Dim II(), IIV()
Dim Block_Size, B2, xs, ys
Dim FDCT(), BDCT(), pi
'q
Dim Rq(), Gq(), Bq(), dd(), coefficients(), pixelvaluesY()
Dim Bt(), Buf(), Bit()
Dim HistII(), Histdd()

Dim Shortest, Steps
Dim VecTbl() As Integer, AvgTbl() As Long, CntTbl() As Long, MaxInd
Private zigzag(7, 7) As Long
''''''''''''''''''''''''''''''''''''''''''''''
Dim uniqueSymbols As New Collection
Dim shiftedCodes As New Collection
Dim HuffmanCodedData As String
''''''''''''''''''''''''''''''''''''''''''''''

Private Sub CHuffmanCoding_Click()
    If Option5.Value = True Then
        
    ElseIf Option4.Value = True Then
        DPCM2HuffmanCode
    End If
End Sub

Private Sub DPCM2HuffmanCode()
    Dim i As Long, j As Long
    Dim tempSymbols As New Collection
    
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Open App.Path & "\DPCMEncodingToHuffmanCoding.txt" For Output As #1
    Print #1, "Huffman Coding Input"
    
    Dim sRow As String
    
    i = 0
    While (i < Hgt)
        j = 0
        sRow = ""
        While (j < Wid)
            sRow = sRow & " " & Picture5.Point(i, j)
            j = j + 1
        Wend
        Print #1, sRow
        i = i + 1
    Wend
    Close #1
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    
    ' Finding Unique Codes and their quantity
    i = 0
    While (i < Hgt)
        j = 0
        While (j < Wid)
            If (isUnique(Picture6.Point(i, j))) Then
                addUnique Picture6.Point(i, j)
            Else
                incrementUnique Picture6.Point(i, j)
            End If
            j = j + 1
        Wend
        i = i + 1
    Wend
    
    'Creating the Tree
    
    'Initialising tempSymbols
    i = 1
    Dim x As UniqueCode
    While (i <= uniqueSymbols.Count)
        Set x = uniqueSymbols.Item(i)

        addSymbol tempSymbols, x
        i = i + 1
    Wend
    
    While tempSymbols.Count > 1
    
        'Sort tempSymbols
        Dim first As New UniqueCode, second As New UniqueCode, sTemp As New UniqueCode
        i = 1
        While (i <= tempSymbols.Count - 1)
            j = i + 1
            While (j <= tempSymbols.Count)
                                
                Set first = tempSymbols.Item(i)
                Set second = tempSymbols.Item(j)
                
                If (first.Quantity > second.Quantity) Then
                    sTemp.Code = first.Code
                    sTemp.Quantity = first.Quantity
                    
                    first.Code = second.Code
                    first.Quantity = second.Quantity
                
                    second.Code = sTemp.Code
                    second.Quantity = sTemp.Quantity
                End If
                
                j = j + 1
            Wend
            i = i + 1
        Wend
        
        'Merge Least Quantity Symbols
        
        Set first = tempSymbols.Item(1)
        Set second = tempSymbols.Item(2)
        
        sTemp.Code = "(" & first.Code & "," & second.Code & ")"
        sTemp.Quantity = first.Quantity + second.Quantity
    
        tempSymbols.Remove (1)
        tempSymbols.Remove (1)
        
        addSymbol tempSymbols, sTemp
        
    Wend
    
    ' Deriving Huffman code for each Symbol from Tree
    Dim sTree As String, sHuffCode As String, sSymbol As String, sChar As String
    
    Set first = tempSymbols.Item(1)
    sTree = first.Code

    While (Len(sTree) <> 0)
        sChar = Left(sTree, 1)
        sTree = Right(sTree, Len(sTree) - 1)
        
        If (StrComp(sChar, "(") = 0) Then
            sHuffCode = sHuffCode & "0"
        ElseIf (StrComp(sChar, ",") = 0) Then
            sHuffCode = Left(sHuffCode, Len(sHuffCode) - 1)
            sHuffCode = sHuffCode & "1"
        ElseIf (StrComp(sChar, ")") = 0) Then
            sHuffCode = Left(sHuffCode, Len(sHuffCode) - 1)
        Else
            While (StrComp(sChar, ")") <> 0 And StrComp(sChar, "(") <> 0 And StrComp(sChar, ",") <> 0)
                sSymbol = sSymbol & sChar
                
                sChar = Left(sTree, 1)
                If (StrComp(sChar, ")") <> 0 And StrComp(sChar, "(") <> 0 And StrComp(sChar, ",") <> 0) Then
                    sTree = Right(sTree, Len(sTree) - 1)
                End If
            Wend
            
        End If
        
        j = 1
        While (j <= uniqueSymbols.Count)
        
            Set first = uniqueSymbols.Item(j)
            
            If (StrComp(first.Code, sSymbol) = 0) Then
                sSymbol = ""
                first.HuffCode = sHuffCode
            End If
            
            j = j + 1
        Wend

    Wend
    
    ' Changing Original Data to Huffman Coded Data
    Dim k As Long
    
    i = 0
    While (i < Hgt)
        j = 0
        While (j < Wid)
            k = 1
            While (k <= uniqueSymbols.Count)
                Set first = uniqueSymbols.Item(k)
                
                If (StrComp(first.Code, (Picture6.Point(i, j) & "")) = 0) Then
                    HuffmanCodedData = HuffmanCodedData & first.HuffCode
                End If
                
                k = k + 1
            Wend
            
            j = j + 1
        Wend
    
        i = i + 1
    Wend
    
    MsgBox "Huffman Code for DPCM Generated, Length of Code is " & Len(HuffmanCodedData), vbInformation
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Open App.Path & "\DPCMEncodingToHuffmanCoding.txt" For Append As #1
    Print #1, "Huffman Coding Output"
    
    i = 1
    Print #1, "Point", "Quantity", "Huffman Code For Point"
    While (i <= uniqueSymbols.Count)
        Set x = uniqueSymbols.Item(i)
        Print #1, x.Code, x.Quantity, x.HuffCode
    
        i = i + 1
    Wend
    
    Print #1, "Huffman Code: "
    Print #1, HuffmanCodedData
    
    Close #1
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
End Sub

Private Sub SaveHuffmanCodeToFile()
    Open App.Path & "\HuffmanCode.txt" For Output As #1
    i = 1
    Print #1, "Point", "Quantity", "Huffman Code For Point"
    While (i <= uniqueSymbols.Count)
        Set x = uniqueSymbols.Item(i)
        Print #1, x.Code, x.Quantity, x.HuffCode
    
        i = i + 1
    Wend
    
    Print #1, "Huffman Code: "
    Print #1, HuffmanCodedData
    
    Close #1
    
    MsgBox "Successfully Saved Huffman Code !", vbInformation
End Sub



Private Sub addSymbol(symbolList As Collection, data As UniqueCode)
    Dim temp As New UniqueCode
    temp.Code = data.Code
    temp.Quantity = data.Quantity
    symbolList.Add temp
End Sub

Private Sub addUnique(sCode As String)
    
    Dim temp As New UniqueCode
    temp.Code = sCode
    temp.Quantity = 1
    uniqueSymbols.Add temp
End Sub

Private Sub incrementUnique(sCode As String)
    Dim i As Long
    Dim temp As New UniqueCode
    i = 1
    While (i <= uniqueSymbols.Count)
        Set temp = uniqueSymbols.Item(i)
        If (temp.Code = sCode) Then '
            temp.Quantity = temp.Quantity + 1
            'Set uniqueSymbols.Item(i) = temp
        End If
        i = i + 1
    Wend
End Sub

Private Function isUnique(sCode As String) As Boolean
    Dim i As Long
    Dim temp As New UniqueCode
    Dim bStatus As Boolean
    
    bStatus = True
    i = 1
    While (i <= uniqueSymbols.Count)
        Set temp = uniqueSymbols.Item(i)
        If (temp.Code = sCode) Then '
            bStatus = False
        End If
        i = i + 1
    Wend
    isUnique = bStatus
End Function


Private Sub ColorTransform_Click()
 
ReDim red(Wid, Hgt), green(Wid, Hgt), blue(Wid, Hgt)
 For i = 0 To Wid - 1
   For j = 0 To Hgt - 1
    
    pixel = Picture1.Point(i, j)
    red(i, j) = pixel And &HFF
    green(i, j) = (pixel \ 256) And &HFF
    blue(i, j) = (pixel \ 65536) And &HFF
   Next
 Next
 
ReDim YY(Wid, Hgt), UU(Wid, Hgt), Vv(Wid, Hgt)
 For i = 0 To Wid - 1
   For j = 0 To Hgt - 1
    YY(i, j) = 0.299 * red(i, j) + 0.299 * green(i, j) + 0.299 * blue(i, j)
    UU(i, j) = Abs(Int(blue(i, j) - YY(i, j)))
    Vv(i, j) = Abs(Int(red(i, j) - YY(i, j)))
    Next
 Next
 
 ReDim red2(Wid, Hgt), green2(Wid, Hgt), blue2(Wid, Hgt)

 For i = 0 To Wid - 1
   For j = 0 To Hgt - 1
    Yc = YY(i, j): Uc = UU(i, j): Vc = Vv(i, j)
    red2(i, j) = Uc + Yc
    green2(i, j) = Abs(Yc - (0.195 * Uc) - (0.509 * Vc))
    blue2(i, j) = Vc + Yc
    Picture2.PSet (i, j), RGB(red2(i, j), green2(i, j), blue2(i, j))
        
   Next
 Next
 
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Command2_Click()
    SaveHuffmanCodeToFile
End Sub

Private Sub DCT_Tran_Click()
Block_Size = Text1.Text
 B2 = 2 * Block_Size
 
 Wid = Picture1.ScaleWidth:  Hgt = Picture1.ScaleHeight

 ReDim II(Wid, Hgt), IIV(Wid, Hgt), FDCT(3, Wid, Hgt)
 pi = 3.14
 
 For xs = 0 To Wid - 1 Step Block_Size
   For ys = 0 To Hgt - 1 Step Block_Size
    
    Read_RGB
    
    FDCTrans
     
   Next
   DoEvents
  Next
End Sub
Private Sub FDCTrans()
 
 For u = 0 To Block_Size - 1:  For v = 0 To Block_Size - 1
     sumc = 0
     For x = 0 To Block_Size - 1:  For y = 0 To Block_Size - 1
         sumc = sumc + II(x + xs, y + ys) * _
                       Cos((((2 * x) + 1) * u * pi) / B2) * _
                       Cos((((2 * y) + 1) * v * pi) / B2)
     Next: Next
     s2 = Sqr(2) / Block_Size
     FDCT(0, u + xs, v + ys) = CC(u) * CC(v) * sumc * s2 ' / 8
     f = Abs(FDCT(0, u + xs, v + ys))
     g = From0To256(f)
     Picture3.Refresh: DoEvents
     Picture3.PSet (u + xs, v + ys), RGB(g, g, g)
     
 Next: Next

End Sub

Private Sub Read_RGB()

 For x = xs To xs + Block_Size
   For y = ys To ys + Block_Size
     pix = Picture2.Point(x, y)
     Rr = pix Mod 256
     Gg = (pix / 256) Mod 256
     Bb = (pix / 65536) Mod 256
     II(x, y) = (Rr + Gg + Bb) / 3
   Next
 Next
End Sub
Function CC(vlll)
  If vlll = 0 Then CC = 1 / Sqr(2) Else CC = 1
End Function
Function From0To256(Val)
  If Val < 0 Then From0To256 = 0
  If Val > 256 Then From0To256 = 256
End Function




Private Sub encoding_Click()
    If Option5.Value = True Then
        RLEEncoding
    ElseIf Option4.Value = True Then
        DPCMEncoding
    End If
    
End Sub

Private Sub DPCMEncoding()
    Dim i As Long, j As Long
    Dim xWidth As Long, xHeight As Long
            
    xWidth = Wid
    xHeight = Hgt
    
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Open App.Path & "\DPCMEncoding.txt" For Output As #1
    Print #1, "DPCM Encoding Input"
    
    Dim sRow As String
    
    i = 0
    While (i < Hgt)
        j = 0
        sRow = ""
        While (j < Wid)
            sRow = sRow & " " & Picture5.Point(i, j)
            j = j + 1
        Wend
        Print #1, sRow
        i = i + 1
    Wend
    Close #1
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    
    i = 0
    While (i < xHeight)
        j = 0
        While (j < xWidth)
            If (i = 0 And j = 0) Then
                Picture6.Refresh: DoEvents
                Picture6.PSet (0, 0), Picture5.Point(0, 0)
            Else
                Picture6.Refresh: DoEvents
                Picture6.PSet (i, j), Picture5.Point(i, j) - Picture5.Point(Int(((i * xWidth) + (j - 1)) / xWidth), ((i * xWidth) + (j - 1)) Mod xWidth)
            End If
            j = j + 1
        Wend
        i = i + 1
    Wend
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Open App.Path & "\DPCMEncoding.txt" For Append As #1
    Print #1, "DPCM Encoding Output"
    
    i = 0
    While (i < Hgt)
        j = 0
        sRow = ""
        While (j < Wid)
            sRow = sRow & " " & Picture6.Point(i, j)
            j = j + 1
        Wend
        Print #1, sRow
        i = i + 1
    Wend
    Close #1
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
End Sub

Private Sub RLEEncoding()
    Dim rleList As New Collection
    Dim i As Long, j As Long
    Dim xCount As Long
    
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Open App.Path & "\RLEncoding.txt" For Output As #1
    Print #1, "RLE Encoding Input"
    
    Dim sRow As String
    
    i = 0
    While (i < Hgt)
        j = 0
        sRow = ""
        While (j < Wid)
            sRow = sRow & " " & Picture5.Point(i, j)
            j = j + 1
        Wend
        Print #1, sRow
        i = i + 1
    Wend
    Close #1
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    xCount = 0
    i = 0
    
    While (i < Hgt)
        j = 0
        While (j < Wid)
            If (Picture5.Point(i, j) = 0) Then
                xCount = xCount + 1
            Else
                rleList.Add (xCount & "," & Picture5.Point(i, j))
                xCount = 0
            End If
            j = j + 1
        Wend
        i = i + 1
    Wend
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Open App.Path & "\RLEncoding.txt" For Append As #1
    i = 1
    Print #1, "RLE Encoding Output"
    While (i <= rleList.Count)
        Print #1, rleList.Item(i)
        i = i + 1
    Wend
    Close #1
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
End Sub

Private Sub open_Click()
CD1.ShowOpen
file_name = CD1.FileName
Picture1.Picture = LoadPicture(file_name)
Wid = Picture1.ScaleWidth
Hgt = Picture1.ScaleHeight
End Sub

Private Sub quantization_Click()
 ReDim Rq(Wid, Hgt), Gq(Wid, Hgt), Bq(Wid, Hgt)
   
   ReDim Bt(8), Bit(8), Buf(8)
   
   Bt(0) = 1
   For i = 1 To 8
    Bt(i) = Bt(i - 1) * 2
   Next i
   
   wl = Val(Text2.Text)
   
   For i = 0 To Wid - 1
     For j = 0 To Hgt - 1
         
         Pixl = Picture3.Point(i, j)
       
         Rq(i, j) = Abs(Pixl Mod 256)
         Gq(i, j) = (Pixl / 256) Mod 256
         Bq(i, j) = (Pixl / 65536) Mod 256
       
         Dece2Binry 8, Rq(i, j), Bt
         vr = Binry2Dece(8, Bit(), wl)
            
         Dece2Binry 8, Gq(i, j), Bt
         vg = Binry2Dece(8, Bit(), wl)
            
         Dece2Binry 8, Bq(i, j), Bt
         vbb = Binry2Dece(8, Bit(), wl - 1)
         
         Picture4.PSet (i, j), RGB(vr, vg, vbb)
                
     Next
     Picture4.Refresh
   Next
   
End Sub

Public Sub Dece2Binry(NB, Nu, Bt())
      
 For i = NB - 1 To 0 Step -1
   ss = Nu And Bt(i)
   If ss = 0 Then Bit(i) = 0 Else Bit(i) = 1
 Next
 
End Sub

Public Function Binry2Dece(NB, Bt(), wl)

 Xx = 8 - wl
 For i = NB - 1 To Xx Step -1
   Nu = Nu + (2 ^ i) * Bit(i)
 Next
 
 Binry2Dece = Nu
End Function


Private Sub save_Click()
    SaveHuffmanCodeToFile
End Sub

Private Sub ShiftCoding_Click()
    If Option5.Value = True Then
        
    ElseIf Option4.Value = True Then
        DPCMShifCoding
    End If
End Sub

Private Sub DPCMShifCoding()
Dim i As Long, j As Long, vect() As Long
 Dim Buf As Byte, BitNo As Integer, Bits(30) As Long
 Dim Vv As Long
 
    NoBits = Text3.Text
    Max = 2 ^ NoBits - 1
    
    Buf = 0: BitNo = 0: Bits(0) = 1
    For i = 1 To 8
        Bits(i) = Bits(i - 1) * 2
    Next i
  
    ReDim vect(Hgt * Wid) As Long
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Open App.Path & "\DPCMShiftCode.txt" For Output As #1
    Print #1, "DPCM Shift Code Input"
    
    Dim sRow As String
    
    i = 0
    While (i < Hgt)
        j = 0
        sRow = ""
        While (j < Wid)
            sRow = sRow & " " & Picture6.Point(i, j)
            j = j + 1
        Wend
        Print #1, sRow
        i = i + 1
    Wend
    Close #1
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    i = 0
    While (i < Hgt)
        j = 0
        While (j < Wid)
            vect((i * Wid) + j) = IIf(Picture6.Point(i, j) > 500, 500, Picture6.Point(i, j))
            j = j + 1
        Wend
        i = i + 1
    Wend
        
    If Check1.Value <> 1 Then
        For i = 1 To (Hgt * Wid)
            If vect(i) >= 0 Then
                
                vect(i) = 2 * vect(i)
            Else
                vect(i) = 2 * Abs(vect(i)) - 1
            End If
        Next
    End If
 
    Open App.Path & "\DPCMShiftCode.txt" For Append As #1
    Print #1, "DPCM Shift Code Output"
    
    For i = 1 To (Hgt * Wid)
        If Check1.Value = 1 Then
            If vect(i) < 0 Then
                Buf = (Buf Or Bits(BitNo))
                If BitNo = 7 Then
                    'Buff is ready Here
                    Print #1, Buf
                    'shiftedCodes.Add (Buf)
                    Buf = 0
                    BitNo = 0
                Else
                    BitNo = BitNo + 1
                End If
            Else
                If BitNo = 7 Then
                    'Buff is ready Here
                    Print #1, Buf
                    'shiftedCodes.Add (Buf)
                    Buf = 0
                    BitNo = 0
                Else
                    BitNo = BitNo + 1
                End If
            End If
        End If
    
        Vv = Abs(vect(i))
        While Vv >= Max
        
            For j = 0 To NoBits - 1
                If (CLng(Max) And Bits(j)) <> 0 Then
                    Buf = (Buf Or Bits(BitNo))
                    If BitNo = 7 Then
                        'Buff is ready Here
                        Print #1, Buf
                        'shiftedCodes.Add (Buf)
                        Buf = 0
                        BitNo = 0
                    Else
                        BitNo = BitNo + 1
                    End If
                Else
                    If BitNo = 7 Then
                        'Buff is ready Here
                        Print #1, Buf
                        'shiftedCodes.Add (Buf)
                        Buf = 0
                        BitNo = 0
                    Else
                        BitNo = BitNo + 1
                    End If
                End If
            Next j
            
            Vv = Vv - Max
        Wend
        
        For j = 0 To NoBits - 1
            If (Vv And Bits(j)) <> 0 Then
                Buf = (Buf Or Bits(BitNo))
                If BitNo = 7 Then
                    'Buff is ready Here
                    Print #1, Buf
                    'shiftedCodes.Add (Buf)
                    Buf = 0
                    BitNo = 0
                Else
                    BitNo = BitNo + 1
                End If
            Else
                If BitNo = 7 Then
                    'Buff is ready Here
                    Print #1, Buf
                    'shiftedCodes.Add (Buf)
                    Buf = 0
                    BitNo = 0
                Else
                    BitNo = BitNo + 1
                End If
            End If
        Next j
    Next i
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'    Open App.Path & "\DPCMShiftCode.txt" For Append As #1
'    Print #1, "DPCM Shift Code Output"
'    Dim vByte As Byte
'    i = 1
'    While (i <= shiftedCodes.Count)
'        vByte = shiftedCodes.Item(i)
'        Print #1, vByte
'        i = i + 1
'    Wend
    Close #1
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    MsgBox "DPCM Shift code Completed !", vbInformation

End Sub

Private Function addInteger(sValue As Long) As CInteger
    
    Dim temp As New CInteger
    temp.Value = sValue
        
    Set addInteger = temp
End Function


Private Sub zigzagscan_Click()

Dim i As Long, j As Long, ix As Long, jx As Long, xWidth As Long, xHeight As Long
Dim pix As Integer
Dim bDontExit As Boolean

    xWidth = Wid
    xHeight = Hgt
    
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Open App.Path & "\ZigzagScan.txt" For Output As #1
    Print #1, "Zigzag Input"
    
    Dim sRow As String
    
    i = 0
    While (i < Hgt)
        j = 0
        sRow = ""
        While (j < Wid)
            sRow = sRow & " " & Picture4.Point(i, j)
            j = j + 1
        Wend
        Print #1, sRow
        i = i + 1
    Wend
    Close #1
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    ix = 0
    jx = 0
    i = 0
    j = 0
    bDontExit = True
    
    While (bDontExit)
        'pix = Picture4.Point(i, j)
        Picture5.Refresh: DoEvents
        Picture5.PSet (ix, jx), Picture4.Point(i, j)
     '   Print #1, i, j
        
        If (jx = xWidth - 1) Then
            jx = 0
            ix = ix + 1
        Else
            jx = jx + 1
        End If
        
        'ZigZag Parsing
        If (i = 0) Then
            If (j = xWidth - 1) Then
                If (j Mod 2 = 0) Then
                    i = i + 1
                Else
                    j = j - 1
                    i = i + 1
                End If
            Else
                If (j Mod 2 = 0) Then
                    j = j + 1
                Else
                    j = j - 1
                    i = i + 1
                End If
            End If
        ElseIf (j = 0) Then
            If (i = xHeight - 1) Then
                If (i Mod 2 = 1) Then
                    j = j + 1
                Else
                    i = i - 1
                    j = j + 1
                End If
            Else
                If (i Mod 2 = 1) Then
                    i = i + 1
                Else
                    i = i - 1
                    j = j + 1
                End If
            End If
        ElseIf (j = xWidth - 1 And i = xHeight - 1) Then
            bDontExit = False
                    
        ElseIf (j = xWidth - 1) Then
            If (j Mod 2 = i Mod 2) Then
                i = i + 1
            Else
                j = j - 1
                i = i + 1
            End If
        ElseIf (i = xHeight - 1) Then
            If (j Mod 2 <> i Mod 2) Then
                j = j + 1
            Else
                i = i - 1
                j = j + 1
            End If
        Else
            If (j Mod 2 = i Mod 2) Then
                i = i - 1
                j = j + 1
            Else
                i = i + 1
                j = j - 1
            End If
        End If
    Wend
    Picture5.Refresh
    
    'Close #1
    
    
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Open App.Path & "\ZigzagScan.txt" For Append As #1
    Print #1, "Zigzag Output"
    
    i = 0
    While (i < Hgt)
        j = 0
        sRow = ""
        While (j < Wid)
            sRow = sRow & " " & Picture5.Point(i, j)
            j = j + 1
        Wend
        Print #1, sRow
        i = i + 1
    Wend
    Close #1
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
End Sub



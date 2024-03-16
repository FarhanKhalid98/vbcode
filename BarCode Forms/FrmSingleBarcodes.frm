VERSION 5.00
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Object = "{437AC048-61C7-4A9E-B976-DF5D12A9A884}#17.0#0"; "SITextBox.ocx"
Begin VB.Form FrmSingleBarcodes 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   11130
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15360
   Icon            =   "FrmSingleBarcodes.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   742
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1024
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox CmbPackName 
      Enabled         =   0   'False
      Height          =   315
      Left            =   6360
      Style           =   2  'Dropdown List
      TabIndex        =   29
      Top             =   2100
      Width           =   1425
   End
   Begin VB.ComboBox CmbPage 
      Height          =   315
      ItemData        =   "FrmSingleBarcodes.frx":0ECA
      Left            =   5160
      List            =   "FrmSingleBarcodes.frx":0F01
      Style           =   2  'Dropdown List
      TabIndex        =   23
      Top             =   7335
      Width           =   1740
   End
   Begin VB.ComboBox CmbPrinters 
      Height          =   315
      ItemData        =   "FrmSingleBarcodes.frx":0FBD
      Left            =   5168
      List            =   "FrmSingleBarcodes.frx":0FBF
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   8003
      Width           =   3315
   End
   Begin VB.Frame FraHelp 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Help"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4110
      Left            =   12840
      TabIndex        =   12
      Top             =   1200
      Visible         =   0   'False
      Width           =   4200
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3750
         Left            =   135
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   13
         Tag             =   "NC"
         Text            =   "FrmSingleBarcodes.frx":0FC1
         Top             =   360
         Width           =   3975
      End
      Begin VB.Label LblClose 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   3915
         TabIndex        =   14
         Top             =   90
         Width           =   135
      End
   End
   Begin JeweledBut.JeweledButton BtnClear 
      Height          =   420
      Left            =   7163
      TabIndex        =   5
      Top             =   8813
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   741
      TX              =   "Clear"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FOCUSR          =   -1  'True
      MPTR            =   0
      MICON           =   "FrmSingleBarcodes.frx":1091
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtQty 
      Height          =   315
      Left            =   10485
      TabIndex        =   1
      Top             =   2100
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   556
      Alignment       =   1
      Appearance      =   0
      MaxLength       =   6
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Masked          =   1
      IntegralPoint   =   5
   End
   Begin JeweledBut.JeweledButton BtnPrint 
      Height          =   420
      Left            =   3773
      TabIndex        =   4
      Top             =   8828
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   741
      TX              =   "Print"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FOCUSR          =   -1  'True
      MPTR            =   0
      MICON           =   "FrmSingleBarcodes.frx":10AD
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnClose 
      Height          =   420
      Left            =   8483
      TabIndex        =   6
      Top             =   8813
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   741
      TX              =   "Close"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FOCUSR          =   -1  'True
      MPTR            =   0
      MICON           =   "FrmSingleBarcodes.frx":10C9
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnPreview 
      Height          =   420
      Left            =   2468
      TabIndex        =   11
      Top             =   8813
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   741
      TX              =   "Preview"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FOCUSR          =   -1  'True
      MPTR            =   0
      MICON           =   "FrmSingleBarcodes.frx":10E5
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtX 
      Height          =   315
      Left            =   11603
      TabIndex        =   16
      Tag             =   "NC"
      Top             =   7238
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   556
      Alignment       =   2
      Appearance      =   0
      MaxLength       =   6
      Text            =   "0"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin SITextBox.Txt TxtY 
      Height          =   315
      Left            =   12338
      TabIndex        =   17
      Tag             =   "NC"
      Top             =   7238
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   556
      Alignment       =   2
      Appearance      =   0
      MaxLength       =   6
      Text            =   "0"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin SITextBox.Txt TxtBarcode 
      Height          =   315
      Left            =   7785
      TabIndex        =   0
      Top             =   2100
      Width           =   2700
      _ExtentX        =   4763
      _ExtentY        =   556
      Appearance      =   0
      MaxLength       =   50
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      IntegralPoint   =   7
   End
   Begin JeweledBut.JeweledButton BtnSave 
      Height          =   420
      Left            =   5843
      TabIndex        =   2
      Top             =   8813
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   741
      TX              =   "Save"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FOCUSR          =   -1  'True
      MPTR            =   0
      MICON           =   "FrmSingleBarcodes.frx":1101
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtProductID 
      Height          =   315
      Left            =   1125
      TabIndex        =   26
      Top             =   2100
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   556
      Appearance      =   0
      MaxLength       =   8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Masked          =   1
      IntegralPoint   =   7
   End
   Begin JeweledBut.JeweledButton BtnProduct 
      Height          =   330
      Left            =   2130
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   2100
      Width           =   360
      _ExtentX        =   635
      _ExtentY        =   582
      TX              =   "..."
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FOCUSR          =   -1  'True
      MPTR            =   0
      MICON           =   "FrmSingleBarcodes.frx":111D
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtProductName 
      Height          =   315
      Left            =   2490
      TabIndex        =   28
      Top             =   2100
      Width           =   3870
      _ExtentX        =   6826
      _ExtentY        =   556
      Appearance      =   0
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      IntegralPoint   =   7
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Pack Name"
      Height          =   195
      Left            =   6360
      TabIndex        =   25
      Top             =   1883
      Width           =   840
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Page Size"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   4185
      TabIndex        =   24
      Top             =   7395
      Width           =   870
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Barcode"
      Height          =   195
      Left            =   7785
      TabIndex        =   22
      Top             =   1883
      Width           =   600
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Printer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   4493
      TabIndex        =   21
      Top             =   8048
      Width           =   570
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "H Value"
      Height          =   195
      Left            =   11603
      TabIndex        =   20
      Top             =   7028
      Width           =   570
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "V Value"
      Height          =   195
      Left            =   12338
      TabIndex        =   19
      Top             =   7028
      Width           =   555
   End
   Begin VB.Label LblPrint 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "--- Print Settings ---"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   11603
      TabIndex        =   18
      Top             =   6788
      Width           =   1290
   End
   Begin VB.Label LblHelp 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Help"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   11295
      TabIndex        =   15
      Top             =   540
      Width           =   435
   End
   Begin VB.Label LblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Single BarCodes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Index           =   0
      Left            =   2700
      TabIndex        =   10
      Top             =   270
      Width           =   2355
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Qty"
      Height          =   195
      Left            =   10485
      TabIndex        =   9
      Top             =   1883
      Width           =   240
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Product Name"
      Height          =   195
      Left            =   2520
      TabIndex        =   8
      Top             =   1890
      Width           =   1020
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Product ID"
      Height          =   195
      Left            =   1110
      TabIndex        =   7
      Top             =   1890
      Width           =   765
   End
   Begin VB.Image ImgExit 
      Height          =   315
      Left            =   11625
      Top             =   60
      Width           =   345
   End
   Begin VB.Menu MnuDelete 
      Caption         =   "Delete"
      Visible         =   0   'False
      Begin VB.Menu mniRemoveRow 
         Caption         =   "Remove This Row"
      End
   End
End
Attribute VB_Name = "FrmSingleBarcodes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim DataFile As Integer, Fl As Long, Chunks As Integer
Dim Application1 As New CRAXDRT.Application
Dim Fragment As Integer, Chunk() As Byte, i As Integer, FileName As String
Const ChunkSize As Integer = 16384
Const conChunkSize = 100
Dim strFileNm As String
Dim vCounter As Integer
Dim vIsNewRecord As Boolean
Dim RsPic As New ADODB.Recordset
Dim RsReport As New ADODB.Recordset
Dim Flag As Boolean
Dim ssql As String
Dim vStrSQL As String
Dim vIsNewRow As Boolean
Dim vNoofPages As Integer
Dim vCurrentPage As Integer
Dim vRecNo As Integer
Dim vStartFrom As Integer
Dim vCurrentRecord As Integer
Dim vProductQty As Integer

Private Sub BtnClear_Click()
   SubClearFields
End Sub

Private Sub SubPrinterSetting()
   On Error GoTo ErrorHandler
   Dim vPrinter() As String
   vPrinter = Split(CmbPrinters.Text, ",")
   If cn.Execute("Select * From PrinterSetting where size =  '28*38'  ").RecordCount >= 1 Then
      cn.Execute "UPDATE PrinterSetting set x = " & Val(TxtX.Text) & " , y = " & Val(TxtY.Text) & ", DeviceName = '" & vPrinter(0) & "', DriverName = '" & vPrinter(1) & "', Port = '" & vPrinter(2) & "'"
   Else
      cn.Execute "INSERT INTO PrinterSetting Values(" & Val(TxtX.Text) & " ," & Val(TxtY.Text) & ",'28*38' ,'" & vPrinter(0) & "','" & vPrinter(1) & "','" & vPrinter(2) & "')"
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnClose_Click()
   On Error GoTo ErrorHandler
   Call SubPrinterSetting
   Unload Me
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub SubBarCodeGenerate()
   On Error GoTo ErrorHandler
   Dim vProductID  As String
   Load FrmBarcodeViewer
   cn.Execute ("delete from pic")
   For vCurrentPage = 1 To vCounter
      cn.Execute ("Insert Into Pic Values('" & vProductID & "',null," & vCurrentPage & ",null" & ")")
   Next vCurrentPage
   vProductQty = TxtQty.Text
   vProductID = TxtProductID.Text
   With cn.Execute("select *, isnull(discpc,0) as Disc from products where productid='" & vProductID & "'")
      FrmBarcodeViewer.TxtBarcode.Text = TxtBarcode.Text
      FrmBarcodeViewer.ParaInProductID = vProductID
      FrmBarcodeViewer.ParaInCompany = ObjRegistry.CompanyShortName
      FrmBarcodeViewer.ParaInProductName = !ProductName
'      FrmBarcodeViewer.ParaInRate = IIf(ChkPrice.Value = 1, !RetailPrice, IIf(ChkDiscountedPrice.Value = 1, !RetailPrice - !Disc, 0))
   End With
   FrmBarcodeViewer.cmdEANCreate.Value = True
   ssql = "Select * from ProductBarcodes where ProductID = '" & vProductID & "' and code = '" & Val(vCode) & "'"
   If cn.Execute(ssql).RecordCount = 0 Then
       'CN.Execute "Delete From ProductBarcodes where ProductID = '" & vProductID & "' and code like '11%'"
       cn.Execute "INSERT into ProductBarcodes(ProductID,Code) values ('" & vProductID & "','" & Val(vCode) & "')"
   End If
            
   'FrmBarcodeViewer.Show
   
   'For vCurrentPage = 1 To vNoofPages
      If Rs.State = adStateOpen Then Rs.Close
      Rs.CursorLocation = adUseClient
      Rs.Open "select * from pic", cn, adOpenStatic, adLockOptimistic
'
         Rs.AddNew
            strFileNm = "c:\" & Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(FrmBarcodeViewer.ParaInProductName, "/", "-"), """", "-"), "\", "-"), ".", "-"), "*", "-"), "?", "-"), "&", "-"), ":", "-") & " " & vProductID & ".bmp"
            DataFile = 1
            Open strFileNm For Binary Access Read As DataFile
             Fl = LOF(DataFile)   ' Length of data in file
             If Fl = 0 Then Close DataFile: Exit Sub
             Chunks = Fl \ ChunkSize
             Fragment = Fl Mod ChunkSize
             ReDim Chunk(Fragment)
             Get DataFile, , Chunk()
             Rs(1).AppendChunk Chunk()
             ReDim Chunk(ChunkSize)
             For i = 1 To Chunks
                 Get DataFile, , Chunk()
                 Rs(1).AppendChunk Chunk()
             Next i
             Close DataFile
            vProductQty = vProductQty - 1
            'vCurrentRecord = vCurrentRecord + 1
            'vCounter = vCounter + 1
            Rs(2) = vCounter + vCurrentPage
            Rs(0) = vProductID
            Rs(3) = Val(vCode)
            
         Rs.Update
      vCounter = 0
      Rs.Close
      Set Rs = Nothing
      'Call BtnPrint_Click
      'CN.Execute ("delete from pic")
   'Next vCurrentPage
   Kill "c:\*.bmp"
   'Call BtnClear_Click
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub SubBarCodeGenerate2()
   On Error GoTo ErrorHandler
   
   Dim vProductID  As String
   Load FrmBarcodeViewer
'   vCounter = IIf(Val(TxtStartFrom.Text) = 0, 0, Val(TxtStartFrom.Text) - 1)
'   vCounter = 0
'   If CmbPage.ListIndex < 7 Then
'       vNoofPages = IIf((Val(TxtTotQty.Text) + vCounter) Mod CmbPage.ItemData(CmbPage.ListIndex) = 0, (Val(TxtTotQty.Text) + vCounter) \ CmbPage.ItemData(CmbPage.ListIndex), ((Val(TxtTotQty.Text) + vCounter) \ CmbPage.ItemData(CmbPage.ListIndex)) + 1)
'    Else
'        vNoofPages = 1
'    End If
   
   cn.Execute ("delete from pic")
   cn.Execute ("Insert Into Pic Values('" & TxtProductID.Text & "',null," & vCurrentPage & ",null" & ")")
   
   
   vProductQty = TxtQty.Text
   vProductID = TxtProductID.Text
    With cn.Execute("select *, isnull(discpc,0) as Disc from products where productid='" & vProductID & "'")
      FrmBarcodeViewer.TxtBarcode.Text = TxtBarcode.Text
      FrmBarcodeViewer.ParaInProductID = vProductID
      FrmBarcodeViewer.ParaInCompany = ObjRegistry.CompanyShortName
      FrmBarcodeViewer.ParaInProductName = ""
'      FrmBarcodeViewer.ParaInRate = IIf(ChkPrice.Value = 1, !RetailPrice, IIf(ChkDiscountedPrice.Value = 1, !RetailPrice - !Disc, 0))
   End With
    FrmBarcodeViewer.cmdEANCreate.Value = True
   ssql = "Select * from ProductBarcodes where ProductID = '" & vProductID & "' and code = '" & Val(vCode) & "'"
   If cn.Execute(ssql).RecordCount = True Then
       'CN.Execute "Delete From ProductBarcodes where ProductID = '" & vProductID & "' and code like '11%'"
       ssql = "INSERT into ProductBarcodes(ProductID,Code,qty) values ('" & vProductID & "','" & Val(vCode) & "'," & TxtQty.Text & ")"
       cn.Execute ssql
   End If
   
  
   
         If RsPic.State = adStateOpen Then RsPic.Close
         RsPic.Open "select * from pic", cn, adOpenStatic, adLockOptimistic
         RsPic.AddNew
            strFileNm = "c:\" & Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(FrmBarcodeViewer.ParaInProductID, "/", "-"), """", "-"), "\", "-"), ".", "-"), "*", "-"), "?", "-"), "&", "-"), ":", "-") & " " & vProductID & ".bmp"
            DataFile = 1
            Open strFileNm For Binary Access Read As DataFile
             Fl = LOF(DataFile)   ' Length of data in file
            If Fl = 0 Then
               Close DataFile
               Exit Sub
            End If
             Chunks = Fl \ ChunkSize
             Fragment = Fl Mod ChunkSize
             ReDim Chunk(Fragment)
             Get DataFile, , Chunk()
             RsPic(1).AppendChunk Chunk()
             ReDim Chunk(ChunkSize)
             For i = 1 To Chunks
                 Get DataFile, , Chunk()
                 RsPic(1).AppendChunk Chunk()
             Next i
             Close DataFile
            vProductQty = vProductQty - 1
            'vCurrentRecord = vCurrentRecord + 1
            'vCounter = vCounter + 1
            RsPic(2) = vCounter + vCurrentPage
            RsPic(0) = vProductID
            RsPic(3) = Val(vCode)
            
         RsPic.Update
      
      vCounter = 0
      RsPic.Close
      Set RsPic = Nothing
      'Call BtnPrint_Click
      'CN.Execute ("delete from pic")
   'Next vCurrentPage
   Kill "c:\*.bmp"
   'Call BtnClear_Click
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnProduct_Click()
   If FunSelectProduct(ssButton, True) = True Then
      TxtBarcode.SetFocus
   Else
      TxtProductID.SetFocus
   End If
End Sub

Private Function FunSelectProduct(ByVal CallerName As SelectAccountCaller, ByVal IsOptional As Boolean) As Boolean
   On Error GoTo ErrorHandler
   Dim vStrSQL As String
   If CallerName = ssButton Or CallerName = ssFunctionKey Then
      SchProduct.ParaInWhere = " and isLocked = 0 and isNoCostProduct = 0"
      SchProduct.Show vbModal, Me
      If SchProduct.ParaOutID = "" Then FunSelectProduct = False: Exit Function
      TxtProductID.Text = SchProduct.ParaOutID
   End If
   '---------------------------

   CmbPackName.Clear
   vStrSQL = "select purchasepackingid PackingID, PackingName from Products p inner join packings pk on p.purchasepackingid = pk.packingid" & vbCrLf _
           + "left outer join ProductBarcodes b on b.productid = p.productid" & vbCrLf _
           + " where p.productid = '" & TxtProductID.Text & "' or code='" & TxtProductID.Text & "'"
   With cn.Execute(vStrSQL)
      While Not .EOF
         CmbPackName.AddItem !PackingName
         CmbPackName.ItemData(CmbPackName.NewIndex) = !PackingID
         CmbPackName.ListIndex = 0
         .MoveNext
      Wend
      .Close
   End With

   vStrSQL = "SELECT ProductName from Products where ProductID='" & TxtProductID.Text & "'" & " and isLocked = 0 and isNoCostProduct = 0"
   With cn.Execute(vStrSQL)
      If .RecordCount > 0 Then
         TxtProductName.Text = !ProductName
         FunSelectProduct = True
         .Close
         Exit Function
      Else
         FunSelectProduct = False
         .Close
         TxtProductID.Text = ""
         TxtProductName.Text = ""
         Exit Function
      End If
   End With
Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Private Sub BtnProductRange_Click()
   On Error GoTo ErrorHandler
   SchProductRange.Show vbModal, Me
   If SchProductRange.ParaOutFromID <> "" Then
   Dim vPID As Long, vCounter As Long
   vPID = SchProductRange.ParaOutFromID
   For vCounter = CLng(SchProductRange.ParaOutFromID) To CLng(SchProductRange.ParaOutToID)
      TxtProductID.Text = vPID
      FunSelectProduct ssValidate, False
      TxtQty.Text = SchProductRange.ParaOutQty
      
      vPID = vPID + 1
      DoEvents
   Next vCounter
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub Check1_Click()

End Sub

Private Sub BtnSave_Click()
On Error GoTo ErrorHandler
   If Trim(TxtProductID.Text) = "" Then Exit Sub
   If Trim(TxtBarcode.Text) = "" Then Exit Sub
   cn.BeginTrans
   
     If cn.Execute("select * from productbarcodes where code = '" & TxtBarcode.Text & "'").EOF Then
         cn.Execute ("Insert into productbarcodes (ProductID, code, qty) values ('" & TxtProductID.Text & "','" & TxtBarcode.Text & "'," & TxtQty.Text & ")")
     End If
     cn.CommitTrans
     Call BtnPreview_Click
      
   
   Call SubClearFields
Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub



Private Sub LblPrint_Click()
   TxtX.Enabled = Not TxtX.Enabled
   TxtY.Enabled = Not TxtY.Enabled
   If TxtX.Enabled Then TxtX.SetFocus
   If TxtX.Enabled Then
      LblPrint.ForeColor = vbBlack
   Else
      'TxtFirst.SetFocus
      LblPrint.ForeColor = &H800000
   End If
   If TxtX.Enabled = False Then
      Call SubPrinterSetting
   End If
End Sub

Private Sub TxtProductID_Change()
   If ActiveControl.Name <> TxtProductID.Name Then Exit Sub
   If TxtProductName.Text <> "" Then TxtProductName.Text = ""
End Sub

Private Sub TxtProductID_KeyDown(KeyCode As Integer, Shift As Integer)
'   If KeyCode = vbKeyDown Then Grid.SetFocus
End Sub

Private Sub TxtProductID_Validate(Cancel As Boolean)
   If TxtProductName.Text <> "" Then Exit Sub
   On Error GoTo ErrorHandler
   Dim vTemp As Boolean
   If Trim(TxtProductID.Text) = "" Then Exit Sub
   vTemp = FunSelectProduct(ssValidate, False)
   If vTemp = False Then
      vTemp = FunSelectProduct(ssButton, False)
      Cancel = False
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub LblClose_Click()
   FraHelp.Visible = False
End Sub

Private Sub LblHelp_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
   LblHelp.ForeColor = &H800000
   FraHelp.ZOrder 0
   FraHelp.Visible = True
End Sub

Private Sub LblHelp_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
   If LblHelp.FontUnderline = True Then Exit Sub
   LblHelp.FontUnderline = True
End Sub

Private Sub LblHelp_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
   LblHelp.ForeColor = vbWhite
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
   Dim lngReturnValue As Long
   If Button = 1 Then
      Call ReleaseCapture
      lngReturnValue = SendMessage(Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
   End If
   If LblHelp.FontUnderline = False Then Exit Sub
   LblHelp.FontUnderline = False
End Sub

Private Sub Form_Load()
   On Error GoTo ErrorHandler
    ShowPicture Me, 2
   AddLabelEffect Me, 2, vbWhite, vbBlack, lblEffectBorder
   SetWindowText Me.hWnd, "Multiple Barcodes"
   
   CmbPrinters.Clear
   CmbPrinters.AddItem "Default,winspool,LPT1"
   Dim p
   For Each p In Printers
      CmbPrinters.AddItem p.DeviceName & "," & p.DriverName & "," & p.Port
   Next p
   
'   CN.Execute ("UPDATE sysindexs Set Value = '" & vPrinter(0) & "' where RegistryKey = 'DeviceName'")
'   CN.Execute ("UPDATE sysindexs Set Value = '" & vPrinter(1) & "' where RegistryKey = 'DriverName'")
'   CN.Execute ("UPDATE sysindexs Set Value = '" & vPrinter(2) & "' where RegistryKey = 'Port'")


   CmbPage.ListIndex = 5
   With cn.Execute("select * from PrinterSetting")
     If .RecordCount > 0 Then
        TxtX.Text = !x
        TxtY.Text = !Y
        CmbPage.Text = !Size
        If Not IsNull(!DeviceName) Then
            CmbPrinters.Text = !DeviceName & "," & !DriverName & "," & !Port
        Else
            CmbPrinters.ListIndex = 0
        End If
     End If
     .Close
   End With
   TxtX.Enabled = False
   TxtY.Enabled = False
   LblPrint.ForeColor = &H800000
   HelpLocation Me
   SubClearFields
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   On Error GoTo ErrorHandler
   If KeyCode = vbKeyReturn Then
     
   ElseIf KeyCode = vbKeyEscape Then
      FraHelp.Visible = False
      If TxtProductID.Enabled Then TxtProductID.SetFocus: Call SubClearDetailArea
   ElseIf KeyCode = vbKeyF1 Then
      Select Case ActiveControl.Name
         Case TxtProductID.Name: If FunSelectProduct(ssFunctionKey, True) = True Then TxtBarcode.SetFocus
      End Select
   ElseIf KeyCode = vbKeyF12 And Me.ActiveControl.Name = TxtProductID.Name Then
         KeyCode = 0
'         TxtStartFrom.SetFocus
   ElseIf Shift = vbCtrlMask Then
      
      Select Case KeyCode
         Case vbKeyW
            If BtnClear.Enabled Then BtnClear_Click
            KeyCode = 0
         Case vbKeyP
            If BtnPrint.Enabled Then BtnPrint_Click
            KeyCode = 0
         Case vbKeyV
            If BtnPreview.Enabled Then BtnPreview_Click
            KeyCode = 0
         Case vbKeyQ
            If BtnClose.Enabled Then BtnClose_Click
            KeyCode = 0
      End Select
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub ImgExit_Click()
   Unload Me
End Sub

Private Sub BtnPreview_Click()
   If SetReport Then
       RptReportViewer.Caption = "Single Barcode"
       RptReportViewer.Show vbModal
   End If
End Sub

Private Sub BtnPrint_Click()
    If SetReport Then RptReportViewer.Report.PrintOut False
End Sub

Private Function SetReport() As Boolean
   On Error GoTo ErrorHandler
   SetReport = False
   SubBarCodeGenerate
   If RsReport.State = adStateOpen Then RsReport.Close
'   vStrSQL = "select pic.f1, pic.Barcode, p.ProductID, ProductName, " & IIf(ChkPrice.Value = 0, IIf(ChkDiscountedPrice.Value = 0, "'' as RetailPrice", "'RsPic.' + cast(cast(RetailPrice-isnull(discpc,0) as int) as varchar(10)) RetailPrice"), "'RsPic.' + cast(cast(RetailPrice as int) as varchar(10)) RetailPrice") & " from pic left outer join Products p on pic.productid = p.ProductID Order by Sr"
   vStrSQL = "select pic.f1, pic.Barcode, p.ProductID, '' as  ProductName,  0 as RetailPrice from pic left outer join Products p on pic.productid = p.ProductID Left Outer Join packings pk on p.purchasepackingid = pk.packingid Order by Sr"
   RsReport.Open vStrSQL, cn, adOpenDynamic, adLockReadOnly

   Set RptReportViewer.Report = New CrpMultiBarCodeContinues28X38
   RptReportViewer.Report.DiscardSavedData
   RptReportViewer.Report.Database.SetDataSource RsReport, 3, 1
   RptReportViewer.Report.ParameterFields(1).AddCurrentValue ObjRegistry.CompanyShortName 'CN.Execute("Select ShortName from company").Fields(0).Value
   
   Dim vPrinter() As String
   vPrinter = Split(CmbPrinters.Text, ",")
'   RptReportViewer.Report.PrinterName = vPrinter(0)
'   RptReportViewer.Report.DriverName = vPrinter(1)
'   RptReportViewer.Report.PortName = vPrinter(2)
   RptReportViewer.Report.SelectPrinter vPrinter(1), vPrinter(0), vPrinter(2)
   
'   If CmbPage.ListIndex < 8 Then
'      RptReportViewer.Report.PaperSize = crPaperA4
'      RptReportViewer.Report.LeftMargin = TxtX.Text
'      RptReportViewer.Report.TopMargin = TxtY.Text
'   Else
      RptReportViewer.Report.LeftMargin = TxtX.Text
      RptReportViewer.Report.TopMargin = TxtY.Text
'   End If
   
   
   'RptReportViewer.Show
   'RptReportViewer.Report.PrintOut False
   'MsgBox "Print has been sent to the printer", vbInformation + vbOKOnly, "Alert"
   SetReport = True
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Public Sub SubClearFields()
   On Error GoTo ErrorHandler
   Dim ctl As Control
   For Each ctl In Me.Controls
      If TypeOf ctl Is TextBox Then
         If ctl.Tag = "" Then ctl.Text = ""
      ElseIf TypeOf ctl Is SITextBox.txt Then
         If ctl.Tag = "" Then ctl.Text = ""
      ElseIf TypeOf ctl Is ComboBox Then
      End If
   Next
   
   CmbPackName.Clear
'   ChkPrice.Value = 1
   BtnProduct.Enabled = True
   TxtProductID.Enabled = True
   If TxtProductID.Visible = True Then TxtProductID.SetFocus
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub



Private Sub TxtQty_LostFocus()
   If Me.ActiveControl.Name = TxtProductID.Name Then Exit Sub
   
End Sub

Private Sub SubClearDetailArea()
   CmbPackName.Clear
   TxtProductID.Enabled = True
   TxtProductID.Text = ""
   TxtProductName.Text = ""
   TxtQty.Text = ""
End Sub



Private Sub Grid_BeforeDelete(Cancel As Integer, DispPromptMsg As Integer)
   On Error GoTo ErrorHandler
   DispPromptMsg = 0
'   TxtTotQty.Text = Val(TxtTotQty.Text) - Grid.Columns("Qty").Value
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub








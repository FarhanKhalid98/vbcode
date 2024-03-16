VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Form1"
   ClientHeight    =   5640
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   5640
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdMonth 
      Caption         =   "Update Month"
      Height          =   375
      Left            =   855
      TabIndex        =   5
      Top             =   3510
      Width           =   3405
   End
   Begin VB.CommandButton Command5 
      Caption         =   "5- Delete StoreID = 3"
      Height          =   375
      Left            =   675
      TabIndex        =   4
      Top             =   2205
      Width           =   3405
   End
   Begin VB.CommandButton Command4 
      Caption         =   "4- Delete StoreID = 4"
      Height          =   375
      Left            =   675
      TabIndex        =   3
      Top             =   1710
      Width           =   3405
   End
   Begin VB.CommandButton Command3 
      Caption         =   "3- Delete Opening Stock StoreID = 4"
      Height          =   375
      Left            =   630
      TabIndex        =   2
      Top             =   1125
      Width           =   3405
   End
   Begin VB.CommandButton Command2 
      Caption         =   "2- Delete Opening Stock StoreID = 3"
      Height          =   375
      Left            =   630
      TabIndex        =   1
      Top             =   585
      Width           =   3405
   End
   Begin VB.CommandButton Command1 
      Caption         =   "1- Update Opening Negative Stock Into Zero"
      Height          =   375
      Left            =   630
      TabIndex        =   0
      Top             =   90
      Width           =   3405
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim a As Date

Private Sub CmdMonth_Click()
   a = "03-30-2008"
   While a <= "12-31-2008"
      CN.Execute ("insert tempa values ('" & a & "')")
      a = a + 1
   Wend
   MsgBox "Ok", vbOKOnly, Command1.Caption
End Sub

Private Sub Command1_Click()
   With CN.Execute("select * from OpeningStock where qtyloose < 0")
      While Not .EOF
         CN.Execute ("update OpeningStock set QtyLoose = 0 , Amount = 0 where Productid = '" & !ProductID & "'")
         .MoveNext
      Wend
   End With
   MsgBox "Ok", vbOKOnly, Command1.Caption
End Sub

Private Sub Command2_Click()
   With CN.Execute("select * from OpeningStock where StoreID = 3")
      If .RecordCount Then .MoveFirst
      While Not .EOF
         CN.Execute ("Delete From OpeningStock where StoreID = 3 and Productid = '" & !ProductID & "'")
         .MoveNext
      Wend
   End With
   MsgBox "Ok", vbOKOnly, Command2.Caption
End Sub

Private Sub Command3_Click()
   With CN.Execute("select * from OpeningStock where StoreID = 4")
      While Not .EOF
         CN.Execute ("Delete From OpeningStock where StoreID = 4 and Productid = '" & !ProductID & "'")
         .MoveNext
      Wend
   End With
   MsgBox "Ok", vbOKOnly, Command3.Caption
End Sub

Private Sub Command4_Click()
 'On Error Resume Next
   Dim i As Integer, j As Integer
   Dim Rs1 As New ADODB.Recordset
   If Rs1.State = adStateOpen Then Rs1.Close
   Rs1.Open "select * From StockTransferBody where TransferID = 1 and TransferDate = '10-16-2007'", CN, adOpenStatic, adLockPessimistic
   For i = 1 To Rs1.RecordCount
      Rs1.Delete
   Next i
   CN.Execute ("Delete From StockTransferHeader where TransferID = 1 and TransferDate = '10-16-2007'")
   CN.Execute ("Delete from Stores where StoreID = 4")
   CN.Execute ("Delete from CurrentStockStore where StoreID = 4")
   MsgBox "Delete Store = 4.", vbOKOnly, "Alert"
End Sub

Private Sub Command5_Click()
   'On Error Resume Next
   Dim i As Integer, j As Integer
   Dim Rs1 As New ADODB.Recordset
   If Rs1.State = adStateOpen Then Rs1.Close
   Rs1.Open "select * From SaleHeader where StoreID = 3", CN, adOpenStatic, adLockPessimistic
   Dim Rs2 As New ADODB.Recordset
   For i = 1 To Rs1.RecordCount
      If Rs2.State = adStateOpen Then Rs2.Close
      Rs2.Open "select * From SaleBody where BillID = " & Rs1!BillID & " and BillDate ='" & Rs1!BillDate & "'", CN, adOpenStatic, adLockPessimistic
      For j = 1 To Rs2.RecordCount
         vStr = "Delete from SaleBody where BillID = " & Rs2!BillID & " and BillDate ='" & Rs2!BillDate & "' and ProductID = '" & Rs2!ProductID & "'"
         CN.Execute vStr
         Rs2.MoveNext
      Next j
      CN.Execute "Delete From SaleHeader Where BillID = " & Rs1!BillID & " and BillDate ='" & Rs1!BillDate & "'"
      Rs1.MoveNext
   Next i
   CN.Execute ("Delete from Stores where StoreID = 3")
   CN.Execute ("Delete from CurrentStockStore where StoreID = 3")
   MsgBox "Delete Store = 3.", vbOKOnly, "Alert"
End Sub

Private Sub Form_Click()
   Dim i As Integer, j As Integer
'   Dim Rs1 As New ADODB.Recordset
'   If Rs1.State = adStateOpen Then Rs1.Close
'   Rs1.Open "select * From SaleHeader where StoreID = 3", CN, adOpenStatic, adLockPessimistic
'   Dim Rs2 As New ADODB.Recordset
'   For i = 1 To Rs1.RecordCount
'      If Rs2.State = adStateOpen Then Rs2.Close
'      Rs2.Open "select * From SaleBody where BillID = " & Rs1!BillID & " and BillDate ='" & Rs1!BillDate & "'", CN, adOpenStatic, adLockPessimistic
'      For j = 1 To Rs2.RecordCount
'         vStr = "Delete from SaleBody where BillID = " & Rs2!BillID & " and BillDate ='" & Rs2!BillDate & "' and ProductID = '" & Rs2!ProductID & "'"
'         CN.Execute vStr
'         Rs2.MoveNext
'      Next j
'      CN.Execute "Delete From SaleHeader Where BillID = " & Rs1!BillID & " and BillDate ='" & Rs1!BillDate & "'"
'      Rs1.MoveNext
'   Next i
   Dim a As String, vSQL As String
   With CN.Execute("select * from Products$")
      For i = 1 To .RecordCount
         If IsNull(!AlternateCode) = False Then
            a = CStr(!AlternateCode)
            vSQL = "update Products$ set ACode = '" & a & "' where ProductCode = " & !ProductCode
            CN.Execute vSQL
         End If
         .MoveNext
      Next i
   End With
End Sub


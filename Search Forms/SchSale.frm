VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Object = "{437AC048-61C7-4A9E-B976-DF5D12A9A884}#17.0#0"; "SITextBox.ocx"
Object = "{E8671A8B-E5DD-11CD-836C-0000C0C14E92}#1.0#0"; "SSCALA32.OCX"
Begin VB.Form SchSale 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   11940
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15450
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   796
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1030
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TxtSID 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   1041
      TabIndex        =   30
      Top             =   2115
      Width           =   1590
   End
   Begin VB.CheckBox ChkIncludeDate 
      Appearance      =   0  'Flat
      BackColor       =   &H00EFC09E&
      Caption         =   "Include Date"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   2700
      TabIndex        =   20
      Top             =   2205
      Value           =   1  'Checked
      Width           =   1290
   End
   Begin VB.ComboBox CmbType 
      Height          =   315
      Left            =   9538
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   2723
      Visible         =   0   'False
      Width           =   1365
   End
   Begin VB.TextBox TxtTTLAmount 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   6073
      TabIndex        =   17
      Top             =   2723
      Width           =   945
   End
   Begin VB.TextBox TxtTableName 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   8278
      TabIndex        =   5
      Top             =   2723
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.TextBox TxtManualBillNo 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   7018
      TabIndex        =   4
      Top             =   2723
      Width           =   1260
   End
   Begin VB.TextBox TxtTag 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   10903
      TabIndex        =   7
      Top             =   2723
      Width           =   1905
   End
   Begin VB.TextBox TxtCustomerName 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   4138
      TabIndex        =   3
      Top             =   2723
      Width           =   1935
   End
   Begin VB.TextBox TxtBillID 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   1041
      TabIndex        =   1
      Top             =   2723
      Width           =   690
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid Grid 
      Height          =   6660
      Left            =   1044
      TabIndex        =   0
      Top             =   3060
      Width           =   13365
      ScrollBars      =   3
      _Version        =   196616
      RecordSelectors =   0   'False
      stylesets.count =   1
      stylesets(0).Name=   "SelectedRow"
      stylesets(0).ForeColor=   0
      stylesets(0).BackColor=   16579021
      stylesets(0).HasFont=   -1  'True
      BeginProperty stylesets(0).Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      stylesets(0).Picture=   "SchSale.frx":0000
      AllowUpdate     =   0   'False
      MultiLine       =   0   'False
      AllowRowSizing  =   0   'False
      AllowGroupSizing=   0   'False
      AllowColumnSizing=   0   'False
      AllowGroupMoving=   0   'False
      AllowColumnMoving=   2
      AllowGroupSwapping=   0   'False
      AllowColumnSwapping=   0
      AllowGroupShrinking=   0   'False
      AllowColumnShrinking=   0   'False
      AllowDragDrop   =   0   'False
      SelectTypeCol   =   0
      SelectTypeRow   =   1
      RowNavigation   =   1
      ForeColorEven   =   0
      BackColorOdd    =   15724527
      RowHeight       =   423
      ExtraHeight     =   26
      ActiveRowStyleSet=   "SelectedRow"
      Columns.Count   =   22
      Columns(0).Width=   1191
      Columns(0).Caption=   "Bill ID"
      Columns(0).Name =   "ID"
      Columns(0).CaptionAlignment=   2
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   2143
      Columns(1).Caption=   "Bill Date"
      Columns(1).Name =   "Date"
      Columns(1).CaptionAlignment=   2
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).NumberFormat=   "dd/MM/yyyy"
      Columns(1).FieldLen=   256
      Columns(2).Width=   1270
      Columns(2).Caption=   "Bill Time"
      Columns(2).Name =   "BillTime"
      Columns(2).DataField=   "Column 8"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      Columns(3).Width=   3413
      Columns(3).Caption=   "Customer Name"
      Columns(3).Name =   "CustomerName"
      Columns(3).CaptionAlignment=   2
      Columns(3).DataField=   "Column 4"
      Columns(3).DataType=   8
      Columns(3).FieldLen=   256
      Columns(4).Width=   3200
      Columns(4).Visible=   0   'False
      Columns(4).Caption=   "Order ID"
      Columns(4).Name =   "OrderID"
      Columns(4).DataField=   "Column 17"
      Columns(4).DataType=   8
      Columns(4).FieldLen=   256
      Columns(5).Width=   3200
      Columns(5).Visible=   0   'False
      Columns(5).Caption=   "Store Name"
      Columns(5).Name =   "StoreName"
      Columns(5).CaptionAlignment=   2
      Columns(5).DataField=   "Column 11"
      Columns(5).DataType=   8
      Columns(5).FieldLen=   256
      Columns(6).Width=   3200
      Columns(6).Visible=   0   'False
      Columns(6).Caption=   "Type"
      Columns(6).Name =   "Type"
      Columns(6).CaptionAlignment=   2
      Columns(6).DataField=   "Column 16"
      Columns(6).DataType=   8
      Columns(6).FieldLen=   256
      Columns(7).Width=   3200
      Columns(7).Visible=   0   'False
      Columns(7).Caption=   "Table Name"
      Columns(7).Name =   "TableName"
      Columns(7).CaptionAlignment=   2
      Columns(7).DataField=   "Column 13"
      Columns(7).DataType=   8
      Columns(7).FieldLen=   256
      Columns(8).Width=   3200
      Columns(8).Visible=   0   'False
      Columns(8).Caption=   "EmpName"
      Columns(8).Name =   "EmpName"
      Columns(8).DataField=   "Column 19"
      Columns(8).DataType=   8
      Columns(8).FieldLen=   256
      Columns(9).Width=   1296
      Columns(9).Caption=   "Ttl items"
      Columns(9).Name =   "TotalItems"
      Columns(9).Alignment=   1
      Columns(9).CaptionAlignment=   2
      Columns(9).DataField=   "Column 5"
      Columns(9).DataType=   5
      Columns(9).FieldLen=   256
      Columns(10).Width=   3200
      Columns(10).Caption=   "Ttl Qty"
      Columns(10).Name=   "TotalQtys"
      Columns(10).DataField=   "Column 20"
      Columns(10).DataType=   8
      Columns(10).FieldLen=   256
      Columns(11).Width=   1164
      Columns(11).Caption=   "Disc."
      Columns(11).Name=   "Disc"
      Columns(11).CaptionAlignment=   2
      Columns(11).DataField=   "Column 14"
      Columns(11).DataType=   8
      Columns(11).FieldLen=   256
      Columns(12).Width=   3200
      Columns(12).Visible=   0   'False
      Columns(12).Caption=   "SC"
      Columns(12).Name=   "SC"
      Columns(12).Alignment=   2
      Columns(12).DataField=   "Column 15"
      Columns(12).DataType=   8
      Columns(12).FieldLen=   256
      Columns(13).Width=   1667
      Columns(13).Caption=   "Ttl Amount"
      Columns(13).Name=   "Amount"
      Columns(13).Alignment=   1
      Columns(13).CaptionAlignment=   2
      Columns(13).DataField=   "Column 5"
      Columns(13).DataType=   8
      Columns(13).FieldLen=   256
      Columns(14).Width=   2805
      Columns(14).Caption=   "Bill Type"
      Columns(14).Name=   "BillType"
      Columns(14).CaptionAlignment=   2
      Columns(14).DataField=   "Column 6"
      Columns(14).DataType=   8
      Columns(14).FieldLen=   256
      Columns(15).Width=   1773
      Columns(15).Caption=   "CO"
      Columns(15).Name=   "CO"
      Columns(15).CaptionAlignment=   2
      Columns(15).DataField=   "Column 5"
      Columns(15).DataType=   8
      Columns(15).FieldLen=   256
      Columns(16).Width=   3200
      Columns(16).Visible=   0   'False
      Columns(16).Caption=   "Manual Bill No"
      Columns(16).Name=   "ManualBillNo"
      Columns(16).DataField=   "Column 12"
      Columns(16).DataType=   8
      Columns(16).FieldLen=   256
      Columns(17).Width=   1852
      Columns(17).Caption=   "Closed"
      Columns(17).Name=   "Closed"
      Columns(17).DataField=   "Column 7"
      Columns(17).DataType=   11
      Columns(17).FieldLen=   256
      Columns(17).Style=   2
      Columns(18).Width=   1508
      Columns(18).Caption=   "Replaced"
      Columns(18).Name=   "Replaced"
      Columns(18).DataField=   "Column 9"
      Columns(18).DataType=   8
      Columns(18).FieldLen=   256
      Columns(18).Style=   2
      Columns(19).Width=   3200
      Columns(19).Visible=   0   'False
      Columns(19).Caption=   "Tag"
      Columns(19).Name=   "Tag"
      Columns(19).CaptionAlignment=   0
      Columns(19).DataField=   "Column 10"
      Columns(19).DataType=   8
      Columns(19).FieldLen=   256
      Columns(20).Width=   3200
      Columns(20).Caption=   "StoreID"
      Columns(20).Name=   "StoreID"
      Columns(20).DataField=   "Column 18"
      Columns(20).DataType=   8
      Columns(20).FieldLen=   256
      Columns(21).Width=   3200
      Columns(21).Caption=   "SID"
      Columns(21).Name=   "SID"
      Columns(21).DataField=   "Column 21"
      Columns(21).DataType=   8
      Columns(21).FieldLen=   256
      TabNavigation   =   1
      _ExtentX        =   23574
      _ExtentY        =   11747
      _StockProps     =   79
      BackColor       =   15724527
      BeginProperty PageFooterFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty PageHeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin JeweledBut.JeweledButton BtnSelect 
      Height          =   420
      Left            =   5616
      TabIndex        =   8
      Top             =   9893
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   741
      TX              =   "Select"
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
      MICON           =   "SchSale.frx":001C
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnClose 
      Cancel          =   -1  'True
      Height          =   420
      Left            =   6936
      TabIndex        =   9
      Top             =   9893
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
      MICON           =   "SchSale.frx":0038
      BC              =   14737632
      FC              =   0
   End
   Begin SSCalendarWidgets_A.SSDateCombo DtpFromDate 
      Height          =   330
      Left            =   1738
      TabIndex        =   2
      Top             =   2723
      Width           =   1200
      _Version        =   65543
      _ExtentX        =   2117
      _ExtentY        =   582
      _StockProps     =   93
      BackColor       =   16777215
      BeginProperty DropDownFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DateSeparator   =   "/"
      Format          =   "dd/MM/yyyy"
      BackColorSelected=   16777215
      BevelColorFace  =   14737632
      DividerStyle    =   0
      ForeColorSelected=   6883113
      BevelType       =   0
      SpinButton      =   0
      Mask            =   2
   End
   Begin SSCalendarWidgets_A.SSDateCombo DtpToDate 
      Height          =   330
      Left            =   2940
      TabIndex        =   21
      Top             =   2730
      Width           =   1200
      _Version        =   65543
      _ExtentX        =   2117
      _ExtentY        =   582
      _StockProps     =   93
      BackColor       =   16777215
      BeginProperty DropDownFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DateSeparator   =   "/"
      Format          =   "dd/MM/yyyy"
      BackColorSelected=   16777215
      BevelColorFace  =   14737632
      DividerStyle    =   0
      ForeColorSelected=   6883113
      BevelType       =   0
      SpinButton      =   0
      Mask            =   2
   End
   Begin SITextBox.Txt TxtTotalSale 
      Height          =   315
      Left            =   9495
      TabIndex        =   22
      Tag             =   "D"
      Top             =   1365
      Width           =   1650
      _ExtentX        =   2910
      _ExtentY        =   556
      Alignment       =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin SITextBox.Txt TxtCash 
      Height          =   315
      Left            =   7680
      TabIndex        =   24
      Tag             =   "D"
      Top             =   1365
      Width           =   1650
      _ExtentX        =   2910
      _ExtentY        =   556
      Alignment       =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin SITextBox.Txt TxtBank 
      Height          =   315
      Left            =   5865
      TabIndex        =   26
      Tag             =   "D"
      Top             =   1365
      Width           =   1650
      _ExtentX        =   2910
      _ExtentY        =   556
      Alignment       =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin SITextBox.Txt TxtCredit 
      Height          =   315
      Left            =   4050
      TabIndex        =   28
      Tag             =   "D"
      Top             =   1350
      Width           =   1650
      _ExtentX        =   2910
      _ExtentY        =   556
      Alignment       =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Bill Barcode"
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
      Left            =   1035
      TabIndex        =   31
      Top             =   1890
      Width           =   1035
   End
   Begin VB.Label LblCredit 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Credit"
      Height          =   195
      Left            =   4050
      TabIndex        =   29
      Top             =   1170
      Width           =   405
   End
   Begin VB.Label LblBank 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Bank"
      Height          =   195
      Left            =   5865
      TabIndex        =   27
      Top             =   1170
      Width           =   375
   End
   Begin VB.Label LblCash 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Cash"
      Height          =   195
      Left            =   7680
      TabIndex        =   25
      Top             =   1170
      Width           =   360
   End
   Begin VB.Label LblTotalSale 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Total Sale"
      Height          =   195
      Left            =   9495
      TabIndex        =   23
      Top             =   1170
      Width           =   720
   End
   Begin VB.Label LblType 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Type"
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
      Left            =   9538
      TabIndex        =   19
      Top             =   2498
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Label LblTtlAmount 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Ttl Amount"
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
      Left            =   6073
      TabIndex        =   18
      Top             =   2498
      Width           =   930
   End
   Begin VB.Label LblTableName 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Table Name"
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
      Left            =   8278
      TabIndex        =   16
      Top             =   2498
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Label LblManualBillNo 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Manual Bill #"
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
      Left            =   7018
      TabIndex        =   15
      Top             =   2498
      Width           =   1125
   End
   Begin VB.Label LblTag 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Tag"
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
      Left            =   10843
      TabIndex        =   14
      Top             =   2498
      Width           =   345
   End
   Begin VB.Label LblCustomerName 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Customer Name"
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
      Left            =   4138
      TabIndex        =   13
      Top             =   2498
      Width           =   1335
   End
   Begin VB.Label LblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Search"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   0
      Left            =   3000
      TabIndex        =   12
      Top             =   270
      Width           =   1005
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Bill Date"
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
      Left            =   1738
      TabIndex        =   11
      Top             =   2498
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   345
      Left            =   13309
      Top             =   1628
      Width           =   330
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Bill ID"
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
      Left            =   1041
      TabIndex        =   10
      Top             =   2498
      Width           =   525
   End
End
Attribute VB_Name = "SchSale"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Rs As ADODB.Recordset
Dim vStrSQL As String, vPartyName, vIncludeDate, vIncludeReturnDate As String, vTag As String
Dim vManualBillNo As String, vTtlAmount As String, vTableName As String, vType As String
Dim vOrder As String, vDirection As String, vCol As Byte, vSearchInPreviousState As Boolean
Dim vDateDiff As Byte
Public ParaOutBillID As String
Public ParaOutBillDate As String
Public ParaInBillDate As String
Public ParaOutStoreID As String
Public ParaOutSID As String

Private Sub LoadGrid()
   On Error GoTo ErrorHandler
   Set Rs = New ADODB.Recordset
   
   If ChkIncludeDate = 1 And vDateDiff <> 0 And Val(TxtSID.Text) = 0 Then
    vIncludeDate = " and h.BillDate Between '" & DtpFromDate.DateValue & "' and '" & DtpToDate.DateValue & "'"
    vIncludeReturnDate = " and h.ReturnDate Between '" & DtpFromDate.DateValue & "' and '" & DtpToDate.DateValue & "'"
   ElseIf ChkIncludeDate = 1 And vDateDiff = 0 And Val(TxtSID.Text) = 0 Then
      vIncludeDate = " and h.BillDate = '" & DtpToDate.DateValue & "'"
      vIncludeReturnDate = " and h.ReturnDate = '" & DtpToDate.DateValue & "'"
   Else
    vIncludeDate = " And H.SID = " & Val(TxtSID.Text)
    vIncludeReturnDate = ""
   End If
   
   vStrSQL = "SELECT h.SID, h.BillID as SaleID, H.StoreID, OrderID, h.BillDate as SaleDate, TableName, Substring(CONVERT(varchar(20),isnull(BillTime,0)),13,7) as BillTime,  " & vbCrLf _
         + " case when cash = 1 then 'Cash' " & vbCrLf _
         + " when credit = 1 and isnull(CashReceived,0) > 0 and isnull(BankAmount,0) > 0 then 'Credit + Cash + Bankd card' When credit = 1 and isnull(CashReceived,0) > 0 and isnull(BankAmount,0) = 0 then 'Credit + Cash'  When credit = 1 and isnull(CashReceived,0) = 0 and isnull(BankAmount,0) > 0 then 'Credit + BankCard'  When credit = 1 and isnull(CashReceived,0) = 0 and isnull(BankAmount,0) = 0 then 'Credit'  " & vbCrLf _
         + " when BankCard = 1 and isnull(CashReceived,0) = 0 then 'Bank Card' when BankCard = 1 and isnull(CashReceived,0) > 0 then 'Bank Card + Cash' end as BillType, InvType, " & vbCrLf _
         + " EmpName, EmpName, Case when CustomerID = '621' then isnull(CustomerName,c.AccountName) Else AccountName + isnull(' (' + p.City + ')','') + isnull(' (' + P.address + ')','') End as CustomerName, Round(Amount-isnull(billdisc,0)+isnull(ServiceCharges,0)+isnull(STax,0)+isnull(othercharges,0),0) as TotalAmount, isnull(billdisc,0) + disc as Disc, isnull(ServiceCharges,0) as SC, TotalQtys, Totalitems, UserName, isPosted, isReplace, StoreName, Tag, isnull(ManualBillNo,'')ManualBillNo " & vbCrLf _
         + " FROM SaleHeader h INNER JOIN" & vbCrLf _
         + " (SELECT SID, BillDate, count(Productid) as TotalItems, sum(isnull(multiplier,0)* isnull(QtyPack,0) + Qty + isnull(Bonus,0)) as TotalQtys, sum(DiscVal) as disc, sum(amount) Amount FROM SaleBody h Where 1=1 " & vIncludeDate & " GROUP BY SID, BillDate) b" & vbCrLf _
         + " ON H.SID = B.SID and H.BillDate = B.BillDate" & vbCrLf _
         + " left outer JOIN chartofaccounts c ON h.CustomerID = c.AccountNo " & vbCrLf _
         + " left outer JOIN Parties p ON p.PartyID = c.AccountNo " & vbCrLf _
         + " left outer JOIN Employees emp ON Emp.EmpID = h.EmpID " & vbCrLf _
         + " left outer JOIN Tables tb ON tb.TableID = h.TableID " & vbCrLf _
         + " INNER JOIN users u ON h.userno = u.userno " & vbCrLf _
         + " INNER JOIN Stores s ON s.StoreID = h.StoreID " & vbCrLf _
         + " Where 1=1 " & vbCrLf _
         & vIncludeDate & IIf(ObjUserSecurity.IsAdministrator = False And ObjUserSecurity.IsManager = False, " and isPosted = 0 and h.userno=" & ObjUserSecurity.UserNo, "") & IIf(vSessionID = 0, "", " and SessionID = " & vSessionID) & vPartyName & vTag & vTableName & vType & vTtlAmount & vManualBillNo & vOrder & vDirection
   Rs.Open vStrSQL, cn, adOpenStatic, adLockReadOnly
   
   Set Grid.DataSource = Rs
   Grid.Columns("SID").DataField = "SID"
   Grid.Columns("ID").DataField = "SaleID"
   Grid.Columns("OrderID").DataField = "OrderID"
   Grid.Columns("Date").DataField = "SaleDate"
   Grid.Columns("BillTime").DataField = "BillTime"
   Grid.Columns("TableName").DataField = "TableName"
   Grid.Columns("CustomerName").DataField = "CustomerName"
   Grid.Columns("EmpName").DataField = "EmpName"
   Grid.Columns("TotalItems").DataField = "TotalItems"
   Grid.Columns("TotalQtys").DataField = "TotalQtys"
   Grid.Columns("Disc").DataField = "Disc"
   Grid.Columns("SC").DataField = "SC"
   Grid.Columns("Amount").DataField = "TotalAmount"
   Grid.Columns("CO").DataField = "UserName"
   Grid.Columns("BillType").DataField = "BillType"
   Grid.Columns("Type").DataField = "InvType"
   Grid.Columns("Closed").DataField = "isPosted"
   Grid.Columns("Replaced").DataField = "isReplace"
   Grid.Columns("StoreID").DataField = "StoreID"
   Grid.Columns("StoreName").DataField = "StoreName"
   Grid.Columns("Tag").DataField = "Tag"
   Grid.Columns("ManualBillNo").DataField = "ManualBillNo"
   
   If ObjUserSecurity.IsAdministrator = True Then
      vStrSQL = "Select isnull(Sum(BankNetAmount),0) BankNetAmount, isnull(Sum(CashNetAmount),0) CashNetAmount, isnull(Sum(CreditNetAmount),0) CreditNetAmount From " & vbCrLf _
            + "(  " & vbCrLf _
            + " Select isnull(BankAmount,0) +  CASE WHEN bankcard = 1 THEN SaleAmount - isnull(BillDisc,0) + isnull(OtherCharges,0) + isnull(servicecharges,0) + isnull(STax,0) - isnull(cashReceived,0) Else 0 End BankNetAmount, " & vbCrLf _
            + " Case WHEN Cash = 0 THEN isnull(cashReceived,0) Else 0 End + Case WHEN Cash = 1 THEN SaleAmount - isnull(BillDisc,0) + isnull(OtherCharges,0) + isnull(servicecharges,0) + isnull(STax,0)  Else 0 End CashNetAmount, " & vbCrLf _
            + " Case WHEN Credit = 1 THEN SaleAmount - isnull(BillDisc,0) + isnull(OtherCharges,0) + isnull(servicecharges,0) + isnull(STax,0) - isnull(BankAmount,0) - isnull(cashReceived,0)  Else 0 End CreditNetAmount " & vbCrLf _
            + " From SaleHeader h inner join (select SID, BillDate, sum(amount) SaleAmount from SaleBody group by SID, BillDate)b On H.SID = B.SID And H.BillDate = B.Billdate  " & vbCrLf _
            + " Where 1=1 " & vIncludeDate & IIf(ObjUserSecurity.IsAdministrator = False And ObjUserSecurity.IsManager = False, " and h.userno=" & ObjUserSecurity.UserNo, "") & vbCrLf _
            + " Union All  " & vbCrLf _
            + " Select CASE WHEN bankcard = 1 THEN -(SaleAmount - isnull(BillDisc,0) + isnull(OtherCharges,0) + isnull(servicecharges,0) + isnull(STax,0)) Else 0 End BankNetAmount," & vbCrLf _
            + " Case WHEN Cash = 1 THEN -(SaleAmount - isnull(BillDisc,0) + isnull(OtherCharges,0) + isnull(servicecharges,0) + isnull(STax,0)) Else 0 End CashNetAmount, " & vbCrLf _
            + " Case WHEN Credit = 1 THEN -(SaleAmount - isnull(BillDisc,0) + isnull(OtherCharges,0) + isnull(servicecharges,0) + isnull(STax,0)) Else 0 End CreditNetAmount " & vbCrLf _
            + " From SaleReturnHeader h inner join (select SID, ReturnDate, sum(amount) SaleAmount from SaleReturnBody group by SID, ReturnDate)b On H.SID = B.SID And H.ReturnDate = B.ReturnDate  " & vbCrLf _
            + " Where 1=1 " & vIncludeReturnDate & IIf(ObjUserSecurity.IsAdministrator = False And ObjUserSecurity.IsManager = False, " and h.userno=" & ObjUserSecurity.UserNo, "") & vbCrLf _
            + ")SaleAmount"
      With cn.Execute(vStrSQL)
         If .EOF = False Then
            TxtCredit.Text = !CreditNetAmount
            TxtBank.Text = !BankNetAmount
            TxtCash.Text = !CashNetAmount
            TxtTotalSale.Text = !CreditNetAmount + !BankNetAmount + !CashNetAmount
         End If
      End With
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnClose_Click()
   Me.ParaOutBillID = -1
   Me.ParaOutBillDate = ""
   Me.ParaOutSID = -1
   Unload Me
End Sub

Private Sub BtnSelect_Click()
   On Error GoTo ErrorHandler
   If Grid.rows = 0 Then Exit Sub
   If Abs(Rs!isReplace) = 1 Then
      Me.ParaOutBillID = -1
      Me.ParaOutSID = -1
      Me.ParaOutStoreID = -1
      Me.ParaOutBillDate = ""
   Else
      Me.ParaOutSID = Rs!SID
      Me.ParaOutBillID = Rs!SaleID
      Me.ParaOutStoreID = Rs!StoreID
      Me.ParaOutBillDate = Rs!SaleDate
   End If
   Unload Me
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub ChkIncludeDate_Click()
    Call TxtCustomerName_Change
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub DtpFromDate_Change()
   If DtpFromDate.IsDateValid = False Then Exit Sub
   If DtpToDate.Visible = False Then DtpToDate.DateValue = DtpFromDate.DateValue
   vOrder = " Order by SaleDate ASC, SaleID"
   vDirection = " ASc"
   Call LoadGrid
End Sub

Private Sub DtpToDate_Change()
   If DtpToDate.IsDateValid = False Then Exit Sub
   Call LoadGrid
End Sub

Private Sub Form_Load()
   On Error GoTo ErrorHandler
   ShowPicture Me, 2
   AddLabelEffect Me, 2, vbWhite, vbBlack, lblEffectBorder
   SetWindowText Me.hWnd, "Search"
   
   If ObjUserSecurity.ShowSumInSearchSaleInvoice = False Then
      LblTotalSale.Visible = False
      TxtTotalSale.Visible = False
      LblCredit.Visible = False
      TxtCredit.Visible = False
      LblBank.Visible = False
      TxtBank.Visible = False
      LblCash.Visible = False
      TxtCash.Visible = False
   End If
   
   
   
   
   DtpToDate.DateValue = Me.ParaInBillDate
   DtpFromDate.DateValue = DtpToDate.DateValue
   vDateDiff = ObjRegistry.SearchDateDifference
   If vDateDiff = 0 Then
      DtpFromDate.Visible = False
      DtpToDate.Left = DtpFromDate.Left
   Else
      DtpFromDate.DateValue = DateAdd("D", -vDateDiff, DtpFromDate.DateValue)
   End If
   
   vSearchInPreviousState = ObjRegistry.ProductSearchOpenInPreviousState

   LblTag.Visible = ObjRegistry.Tag
   TxtTag.Visible = ObjRegistry.Tag
   Grid.Columns("Tag").Visible = ObjRegistry.Tag
   
   LblType.Visible = ObjRegistry.InvType
   CmbType.Visible = ObjRegistry.InvType
   Grid.Columns("Type").Visible = ObjRegistry.InvType
   Grid.Columns("Type").Width = 80
   
   LblManualBillNo.Visible = ObjRegistry.ManualBillNoVisible
   TxtManualBillNo.Visible = ObjRegistry.ManualBillNoVisible
   Grid.Columns("ManualBillNo").Visible = ObjRegistry.ManualBillNoVisible
   Grid.Columns("ManualBillNo").Width = 70
   
   LblTableName.Visible = ObjRegistry.TableVisible
   TxtTableName.Visible = ObjRegistry.TableVisible
   Grid.Columns("TableName").Visible = ObjRegistry.TableVisible
   Grid.Columns("TableName").Width = 60
   Grid.Columns("SC").Visible = ObjRegistry.TableVisible
   Grid.Columns("SC").Width = 40
   
   Grid.Columns("StoreName").Visible = ObjRegistry.StoreVisible
   Grid.Columns("StoreName").Width = 80
   
   Grid.Columns("OrderID").Visible = ObjRegistry.SaleOrderVisible
   Grid.Columns("OrderID").Width = 50
   
   Grid.Columns("EmpName").Visible = True
   Grid.Columns("EmpName").Visible = ObjRegistry.EmpVisible
   
   If ObjUserSecurity.IsAdministrator = False Then
      Grid.Columns("Amount").Visible = Not ObjRegistry.HideSaleAmount
   End If
   
   Dim vWidth As Long, i As Integer
   vWidth = 0
   For i = 0 To Grid.Cols - 1
      If Grid.Columns(i).Visible = True Then
         vWidth = vWidth + Grid.Columns(i).Width
      End If
   Next i
'   Grid.Width = vWidth + 18
   If ObjUserSecurity.IsAdministrator = False Then
      Grid.Columns("Amount").Visible = Not ObjRegistry.HideSaleAmount
   End If

   DtpToDate.DateValue = Me.ParaInBillDate
   DtpFromDate.DateValue = DtpToDate.DateValue - vDateDiff
   
   If vDateDiff = 0 Then
      DtpToDate.Visible = True
'      LblCustomerName.Left = LblCustomerName.Left - DtpToDate.Width
'      TxtCustomerName.Left = TxtCustomerName.Left - DtpToDate.Width
'      LblTtlAmount.Left = LblTtlAmount.Left - DtpToDate.Width
'      TxtTTLAmount.Left = TxtTTLAmount.Left - DtpToDate.Width
'      LblTableName.Left = LblTableName.Left - DtpToDate.Width
'      TxtTableName.Left = TxtTableName.Left - DtpToDate.Width
'      LblTag.Left = LblTag.Left - DtpToDate.Width
'      TxtTag.Left = TxtTag.Left - DtpToDate.Width
'      LblManualBillNo.Left = LblManualBillNo.Left - DtpToDate.Width
'      TxtManualBillNo.Left = TxtManualBillNo.Left - DtpToDate.Width
'      LblType.Left = LblType.Left - DtpToDate.Width
'      CmbType.Left = CmbType.Left - DtpToDate.Width
      
      LblCustomerName.Left = LblCustomerName.Left
      TxtCustomerName.Left = TxtCustomerName.Left
      LblTtlAmount.Left = LblTtlAmount.Left
      TxtTTLAmount.Left = TxtTTLAmount.Left
      LblTableName.Left = LblTableName.Left
      TxtTableName.Left = TxtTableName.Left
      LblTag.Left = LblTag.Left
      TxtTag.Left = TxtTag.Left
      LblManualBillNo.Left = LblManualBillNo.Left
      TxtManualBillNo.Left = TxtManualBillNo.Left
      LblType.Left = LblType.Left
      CmbType.Left = CmbType.Left
      
   End If
   Me.ParaOutSID = -1
   Me.ParaOutBillID = -1
   Me.ParaOutBillDate = ""
   
   CmbType.Clear
   CmbType.AddItem ""
   With cn.Execute("select * from InvTypes")
      If .RecordCount > 0 Then
         While Not .EOF
            CmbType.AddItem ![InvType]
            .MoveNext
         Wend
      End If
   End With

   vOrder = " Order by SaleDate Desc, SaleID"
   vDirection = " Desc"
   If vSearchInPreviousState = False Then
      vTag = ""
      vTtlAmount = ""
      vPartyName = ""
      vManualBillNo = ""
      vTableName = ""
      vType = ""
   End If
   Call LoadGrid
   
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyEscape Then Call BtnClose_Click
   If KeyCode = vbKeyReturn Then
      If ActiveControl.Name = TxtSID.Name Then LoadGrid
      Select Case ActiveControl.Name
      Case Grid.Name, TxtSID.Name, TxtBillID.Name, DtpFromDate.Name, DtpToDate.Name
         Call BtnSelect_Click
      End Select
   End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   On Error GoTo ErrorHandler
   Dim frmObj As Object
   For Each frmObj In Forms
       Set frmObj = Nothing
   Next
   Set Rs = Nothing
Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub Grid_DblClick()
  If Grid.rows > 0 Then BtnSelect_Click
End Sub

Private Sub Grid_HeadClick(ByVal ColIndex As Integer)
   vOrder = " order by " & Grid.Columns(ColIndex).DataField
   If vCol = ColIndex Then
      vDirection = IIf(vDirection = " Asc", " Desc", " Asc")
   Else
      vDirection = " Asc"
   End If
   vCol = ColIndex
   LoadGrid
End Sub

Private Sub Grid_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
   Case vbKey0 To vbKey9
      TxtBillID.Text = Chr(KeyAscii): TxtBillID.SelStart = Len(TxtBillID.Text): TxtBillID.SetFocus
   End Select
End Sub

Private Sub Image1_Click()
   Unload Me
End Sub

Private Sub TxtBillID_Change()
   On Error GoTo ErrorHandler
   If Trim(TxtBillID.Text) = "" Then Grid.MoveFirst: Exit Sub
   Rs.Find "SaleID = " & TxtBillID.Text, , adSearchForward, 1
   If Rs.EOF Then Grid.MoveLast
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub TxtCustomerName_Change()
   On Error GoTo ErrorHandler
   vPartyName = " and (Case when CustomerID = '621' then isnull(CustomerName,AccountName) Else AccountName + isnull(' (' + P.City + ')','') + isnull(' (' + P.address + ')','') End) like '%" & TxtCustomerName.Text & "%'"
   LoadGrid
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub TxtSID_LostFocus()
On Error GoTo ErrorHandler
'   LoadGrid
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub TxtTableName_Change()
   On Error GoTo ErrorHandler
   vTableName = " and TableName Like '%" & TxtTableName.Text & "%'"
   LoadGrid
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub TxtTag_Change()
   On Error GoTo ErrorHandler
   vTag = " and Tag Like '%" & TxtTag.Text & "%'"
   LoadGrid
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub TxtTTLAmount_Change()
   On Error GoTo ErrorHandler
   vTtlAmount = IIf(Val(TxtTTLAmount.Text) = 0, "", " and amount-isnull(billdisc,0)+isnull(ServiceCharges,0)+isnull(STax,0)+isnull(othercharges,0)  = " & Val(TxtTTLAmount.Text))
'   vTtlAmount = IIf(Val(TxtTTLAmount.Text) = 0, "", " and TotalAmount  = " & Val(TxtTTLAmount.Text))
   LoadGrid
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub TxtManualBillNo_Change()
   On Error GoTo ErrorHandler
   vManualBillNo = " and ManualBillNo like '%" & (TxtManualBillNo.Text) & "%'"
   LoadGrid
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub CmbType_Click()
   On Error GoTo ErrorHandler
   If CmbType.ListIndex = 0 Then
      vType = ""
   Else
      vType = " and InvType = '" & CmbType.Text & "'"
   End If
   LoadGrid
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

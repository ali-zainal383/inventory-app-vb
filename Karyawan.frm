VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form karyawan 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Master Karyawan"
   ClientHeight    =   6135
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12600
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6135
   ScaleWidth      =   12600
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text9 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   8880
      TabIndex        =   25
      Text            =   "Text9"
      Top             =   2400
      Width           =   3420
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Align           =   2  'Align Bottom
      Height          =   3255
      Left            =   0
      TabIndex        =   24
      Top             =   2880
      Width           =   12600
      _ExtentX        =   22225
      _ExtentY        =   5741
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   18
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   5400
      Top             =   1440
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.TextBox Text8 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   8880
      TabIndex        =   22
      Text            =   "Text8"
      Top             =   1200
      Width           =   3465
   End
   Begin VB.TextBox Text7 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   8880
      TabIndex        =   21
      Text            =   "Text7"
      Top             =   840
      Width           =   3465
   End
   Begin VB.TextBox Text6 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   8880
      TabIndex        =   20
      Text            =   "Text6"
      Top             =   480
      Width           =   3465
   End
   Begin VB.TextBox Text5 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   8880
      TabIndex        =   19
      Text            =   "Text5"
      Top             =   120
      Width           =   3465
   End
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   1680
      TabIndex        =   18
      Text            =   "Text4"
      Top             =   1560
      Width           =   3300
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   1680
      TabIndex        =   17
      Text            =   "Text3"
      Top             =   1200
      Width           =   3300
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Perempuan"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   3360
      TabIndex        =   16
      Top             =   840
      Width           =   1335
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Laki-Laki"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   1680
      TabIndex        =   15
      Top             =   840
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   1680
      TabIndex        =   14
      Text            =   "Text2"
      Top             =   480
      Width           =   3300
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   1680
      TabIndex        =   13
      Text            =   "Text1"
      Top             =   120
      Width           =   3300
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Keluar"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   3720
      TabIndex        =   12
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Hapus"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   2520
      TabIndex        =   11
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Batal"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   1320
      TabIndex        =   10
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Simpan"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   120
      TabIndex        =   9
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label Label10 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Cari Data"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   7320
      TabIndex        =   23
      Top             =   2400
      Width           =   1500
   End
   Begin VB.Label Label9 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tanggal Lahir"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   7320
      TabIndex        =   8
      Top             =   1200
      Width           =   1500
   End
   Begin VB.Label Label8 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Telepon"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   7320
      TabIndex        =   7
      Top             =   840
      Width           =   1500
   End
   Begin VB.Label Label7 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Jabatan"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   7320
      TabIndex        =   6
      Top             =   120
      Width           =   1500
   End
   Begin VB.Label Label6 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Divisi"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   7320
      TabIndex        =   5
      Top             =   480
      Width           =   1500
   End
   Begin VB.Label Label5 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Kota Asal"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   120
      TabIndex        =   4
      Top             =   1560
      Width           =   1500
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Alamat"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   1500
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Jenis Kelamin"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   1500
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Nama Karyawan"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   1500
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Nik "
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1500
   End
End
Attribute VB_Name = "karyawan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim kelamin As String
If Option1.Value = True Then
    kelamin = "Laki-Laki"
Else
    kelamin = "Perempuan"
End If

Call CariData
    If RSKaryawan.EOF Then
        simpan = "insert into TBLKaryawan values('" & Text1 & "','" & Text2 & "','" & kelamin & "','" & Text3 & "','" & Text4 & "','" & Text6 & "','" & Text5 & "','" & Text7 & "','" & Text8 & "')"
        CONN.Execute simpan
        Call kosongkan
        Form_Activate
    Else
        edit = "update TBLKaryawan set nama = '" & Text2 & "',jenis= '" & kelamin & "',alamat='" & Text3 & "',kota='" & Text4 & "',divisi='" & Text6 & "',jabatan='" & Text5 & "',telepon='" & Text7 & "',ttl='" & Text8 & "' where nik='" & Text1 & "'"
        CONN.Execute edit
        Call kosongkan
        Form_Activate
End If
End Sub

Private Sub Command2_Click()
Call kosongkan
Text1.SetFocus
End Sub

Private Sub Command3_Click()
If Text1 = "" Then
    MsgBox "NIK tidak boleh kosong"
    Exit Sub
Else
    Call CariData
    If Not RSKaryawan.EOF Then
        pesan = MsgBox("Apakah anda yakin..??", vbYesNo)
        If pesan = vbYes Then
            hapus = "delete * from TBLKaryawan where nik='" & Text1 & "'"
            CONN.Execute hapus
            Call kosongkan
            Form_Activate
        Else
            Call kosongkan
        End If
    End If
End If
End Sub

Private Sub Command4_Click()
Unload Me
End Sub

Private Sub Form_Activate()
Text1.SetFocus
Call Koneksi
Adodc1.ConnectionString = LokasiData
Adodc1.RecordSource = "TBLKaryawan"
Adodc1.Refresh
Set DataGrid1.DataSource = Adodc1
DataGrid1.Columns(0).Width = 1000
DataGrid1.Columns(1).Width = 1900
DataGrid1.Columns(2).Width = 1050
DataGrid1.Columns(3).Width = 2200
DataGrid1.Columns(4).Width = 1100
DataGrid1.Columns(5).Width = 1300
DataGrid1.Columns(6).Width = 1400
DataGrid1.Columns(7).Width = 1180
DataGrid1.Columns(8).Width = 1100
DataGrid1.Refresh
End Sub

Sub kosongkan()
Text1 = ""
Text2 = ""
Option1.Value = False
Option2.Value = False
Text3 = ""
Text4 = ""
Text5 = ""
Text6 = ""
Text7 = ""
Text8 = ""
Text9 = ""
End Sub

Sub CariData()
Call Koneksi
RSKaryawan.Open "select * from TBLKaryawan where nik ='" & Text1 & "'", CONN
RSKaryawan.Requery
End Sub

Sub DataBaru()
Text2 = ""
Text3 = ""
Text4 = ""
Text5 = ""
Text6 = ""
Text7 = ""
Text8 = ""
End Sub

Sub Ketemu()
Dim kelamin, divisi As String
    Text2 = RSKaryawan!nama
    kelamin = RSKaryawan!jenis
        If kelamin = "Laki-Laki" Then
            Option1.Value = True
        Else
            Option2.Value = True
        End If
    Text3 = RSKaryawan!alamat
    Text4 = RSKaryawan!kota
    Text5 = RSKaryawan!jabatan
    Text6 = RSKaryawan!divisi
    Text7 = RSKaryawan!telepon
    Text8 = RSKaryawan!ttl
    Text2.SetFocus
End Sub

Private Sub Form_Load()
Call kosongkan
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Text1 = "" Then
        MsgBox "NIK tidak boleh kosong"
        Text1.SetFocus
        Exit Sub
    Else
        Call CariData
        If Not RSKaryawan.EOF Then
            Call Ketemu
        Else
            Call DataBaru
        End If
    End If
End If
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack Or vbkeyascii = vbKeyDelete Or vbkeyascii = vbKeySpace Or KeyAscii = vbKeyReturn) Then
    MsgBox "Hanya boleh diisi oleh angka", vbInformation + vbOKOnly, "Perhatian"
    KeyAscii = 0
End If
End Sub

Private Sub Text9_Change()
Call Koneksi
RSKaryawan.Open "select * from TBLKaryawan where nama like '%" & Text9 & "%'", CONN
RSKaryawan.Requery
If Not RSKaryawan.EOF Then
    Adodc1.ConnectionString = LokasiData
    Adodc1.RecordSource = "select * from TBLKaryawan where nama like '%" & Text9 & "%'"
    Adodc1.Refresh
    Set DataGrid1.DataSource = Adodc1
    DataGrid1.Columns(0).Width = 1000
    DataGrid1.Columns(1).Width = 1900
    DataGrid1.Columns(2).Width = 1050
    DataGrid1.Columns(3).Width = 2200
    DataGrid1.Columns(4).Width = 1100
    DataGrid1.Columns(5).Width = 1300
    DataGrid1.Columns(6).Width = 1400
    DataGrid1.Columns(7).Width = 1180
    DataGrid1.Columns(8).Width = 1100
    DataGrid1.Refresh
Else
    MsgBox "Data tidak ada"
End If
End Sub

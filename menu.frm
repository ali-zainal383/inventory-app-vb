VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form menu_utama 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Menu Utama"
   ClientHeight    =   4050
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   7350
   LinkTopic       =   "Form1"
   ScaleHeight     =   4050
   ScaleWidth      =   7350
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   450
      Left            =   0
      TabIndex        =   0
      Top             =   3600
      Width           =   7350
      _ExtentX        =   12965
      _ExtentY        =   794
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   2
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
         EndProperty
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
   End
   Begin VB.Menu mnkaryawan 
      Caption         =   "Karyawan"
   End
   Begin VB.Menu mnbarang 
      Caption         =   "Barang"
      Begin VB.Menu mnbarangmasuk 
         Caption         =   "Barang Masuk"
      End
      Begin VB.Menu mnbarangkeluar 
         Caption         =   "Barang Keluar"
      End
   End
   Begin VB.Menu mnmaintenance 
      Caption         =   "Maintenance"
   End
   Begin VB.Menu mnkeluar 
      Caption         =   "Keluar"
   End
End
Attribute VB_Name = "menu_utama"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub mnbarangkeluar_Click()
barang_keluar.Show
End Sub

Private Sub mnbarangmasuk_Click()
barang_masuk.Show
End Sub

Private Sub mnkaryawan_Click()
karyawan.Show
End Sub

Private Sub mnkeluar_Click()
pesan = MsgBox("Tutup Aplikasi..?", vbYesNo)
If pesan = vbYes Then End
End Sub

Private Sub mnmaintenance_Click()
maintenance.Show
End Sub

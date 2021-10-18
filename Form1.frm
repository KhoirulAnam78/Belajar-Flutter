VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form CRUD 
   Caption         =   "CRUD MAHASISWA"
   ClientHeight    =   9315
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9780
   LinkTopic       =   "Form1"
   ScaleHeight     =   9315
   ScaleWidth      =   9780
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Data Mahasiswa"
      Height          =   8775
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   9015
      Begin VB.CommandButton btnReset 
         Caption         =   "Reset"
         Height          =   495
         Left            =   2640
         TabIndex        =   19
         Top             =   4440
         Width           =   1335
      End
      Begin VB.Frame Frame4 
         Caption         =   "Cari Data Berdasarkan NIM"
         Height          =   735
         Left            =   480
         TabIndex        =   16
         Top             =   3240
         Width           =   7935
         Begin VB.CommandButton btnCari 
            Caption         =   "Cari"
            Height          =   375
            Left            =   5640
            TabIndex        =   18
            Top             =   240
            Width           =   1335
         End
         Begin VB.TextBox txtcari 
            Height          =   375
            Left            =   360
            TabIndex        =   17
            Top             =   240
            Width           =   5175
         End
      End
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   330
         Left            =   2640
         Top             =   8160
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   2
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
         Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\TugasVB6\databasemahasiswa.mdb;Persist Security Info=False"
         OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\TugasVB6\databasemahasiswa.mdb;Persist Security Info=False"
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "daftar_mahasiswa"
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
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "Form1.frx":0000
         Height          =   2295
         Left            =   480
         TabIndex        =   15
         Top             =   5640
         Width           =   7935
         _ExtentX        =   13996
         _ExtentY        =   4048
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   19
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Data Mahasiswa"
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
               LCID            =   14345
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
               LCID            =   14345
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
      Begin VB.Frame Frame3 
         Caption         =   "Actions"
         Height          =   1215
         Left            =   480
         TabIndex        =   2
         Top             =   4080
         Width           =   7935
         Begin VB.CommandButton btnKeluar 
            Caption         =   "Keluar"
            Height          =   615
            Left            =   6120
            TabIndex        =   10
            Top             =   1200
            Width           =   1575
         End
         Begin VB.CommandButton btnHapus 
            Caption         =   "Hapus"
            Height          =   495
            Left            =   5760
            TabIndex        =   9
            Top             =   360
            Width           =   1455
         End
         Begin VB.CommandButton btnUbah 
            Caption         =   "Ubah"
            Height          =   495
            Left            =   3960
            TabIndex        =   8
            Top             =   360
            Width           =   1335
         End
         Begin VB.CommandButton btnTambah 
            Caption         =   "Tambah"
            Height          =   495
            Left            =   240
            TabIndex        =   7
            Top             =   360
            Width           =   1455
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Input"
         Height          =   2655
         Left            =   480
         TabIndex        =   1
         Top             =   360
         Width           =   7935
         Begin VB.TextBox txtAngkatan 
            Height          =   375
            Left            =   4080
            TabIndex        =   6
            Top             =   1920
            Width           =   1935
         End
         Begin VB.TextBox txtProdi 
            Height          =   375
            Left            =   4080
            TabIndex        =   5
            Top             =   1440
            Width           =   1935
         End
         Begin VB.TextBox txtNama 
            Height          =   375
            Left            =   4080
            TabIndex        =   4
            Top             =   960
            Width           =   1935
         End
         Begin VB.TextBox txtNim 
            Height          =   375
            Left            =   4080
            TabIndex        =   3
            Top             =   360
            Width           =   1935
         End
         Begin VB.Label Label4 
            Caption         =   "Angkatan"
            Height          =   375
            Left            =   1680
            TabIndex        =   14
            Top             =   2040
            Width           =   2655
         End
         Begin VB.Label Label3 
            Caption         =   "Program Studi"
            Height          =   615
            Left            =   1680
            TabIndex        =   13
            Top             =   1560
            Width           =   2055
         End
         Begin VB.Label Label2 
            Caption         =   "Nama Mahasiswa"
            Height          =   615
            Left            =   1680
            TabIndex        =   12
            Top             =   960
            Width           =   1695
         End
         Begin VB.Label Label1 
            Caption         =   "Nim Mahasiswa"
            Height          =   375
            Left            =   1680
            TabIndex        =   11
            Top             =   480
            Width           =   1815
         End
      End
   End
End
Attribute VB_Name = "CRUD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim pencarian As Boolean

Sub bersih()
    txtNim.Text = ""
    txtAngkatan.Text = ""
    txtcari.Text = ""
    txtNama.Text = ""
    txtProdi.Text = ""
End Sub

Private Sub Adodc1_WillMove(ByVal adReason As ADODB.EventReasonEnum, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)

End Sub

Private Sub btnCari_Click()
    Adodc1.Recordset.Find "nim='" & txtcari.Text & "'", , adSearchForward, 1
    If Not Adodc1.Recordset.EOF Then
        txtNim.Text = Adodc1.Recordset!nim
        txtNama.Text = Adodc1.Recordset!nama_mahasiswa
        txtProdi.Text = Adodc1.Recordset!prodi
        txtAngkatan.Text = Adodc1.Recordset!angkatan
        pencarian = True
    Else
        Call bersih
        MsgBox "Data tidak ditemukan", vbOKOnly, "Pencarian"
        pencarian = False
    End If
End Sub

Private Sub btnHapus_Click()
    If pencarian = True Then
        Dim konfirmasi As VbMsgBoxResult
        konfirmasi = MsgBox("Anda akan Menghapus data Mahasiswa Dengan NIM " + txtNim.Text, vbYesNo, "Konfirmasi Hapus")
        If konfirmasi = vbYes Then
            Adodc1.Recordset.Delete
            pencarian = False
            Call bersih
        End If
    Else
        MsgBox "Cari terlebih dahulu data yang ingin dihapus", vbOKOnly, "Gagal Menghapus Data"
        Call bersih
    End If
End Sub

Private Sub btnReset_Click()
    Call bersih
End Sub

Private Sub btnTambah_Click()
    If txtNim.Text = "" Or txtNama.Text = "" Or txtProdi.Text = "" Or txtAngkatan.Text = "" Then
        MsgBox "Data Belum Lengkap", vbOKOnly, "Tambah Data"
    Else
        Adodc1.Recordset.Find "nim='" & txtNim.Text & "'", , adSearchForward, 1
        If Adodc1.Recordset.EOF Then
            Adodc1.Recordset.AddNew
            Adodc1.Recordset!nim = txtNim.Text
            Adodc1.Recordset!nama_mahasiswa = txtNama.Text
            Adodc1.Recordset!prodi = txtProdi.Text
            Adodc1.Recordset!angkatan = txtAngkatan.Text
            MsgBox "Berhasil Menambahkan Data", vbOKOnly, "Tambah Data"
            Call bersih
        Else
            MsgBox "Data dengan nim " + txtNim.Text + " sudah ada", vbOKOnly, "Tambah Data"
        End If
    End If
End Sub

Private Sub btnUbah_Click()
    If pencarian = True Then
        If txtNim.Text = "" Or txtNama.Text = "" Or txtProdi.Text = "" Or txtAngkatan.Text = "" Then
            MsgBox "Data Belum Lengkap", vbOKOnly, "Ubah Data"
        Else
            Adodc1.Recordset.Update
            Adodc1.Recordset!nim = txtNim.Text
            Adodc1.Recordset!nama_mahasiswa = txtNama.Text
            Adodc1.Recordset!prodi = txtProdi.Text
            Adodc1.Recordset!angkatan = txtAngkatan.Text
            pencarian = False
            MsgBox "Data Berhasil Diubah", vbOKOnly, "Ubah Data"
            Call bersih
        End If
    Else
        MsgBox "Cari data yang ingin diubah", vbOKOnly, "Gagal Mengubah Data"
        Call bersih
    End If
End Sub

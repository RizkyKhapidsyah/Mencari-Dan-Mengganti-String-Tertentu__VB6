VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Mencari dan Mengganti String Tertentu"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6045
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   6045
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   495
      Left            =   3000
      TabIndex        =   2
      Top             =   1920
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   1080
      TabIndex        =   1
      Top             =   1920
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1680
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   600
      Width           =   2415
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Kedua fungsi FindReplace() dan ReplaceFirstInstance() 'berikut ini digunakan untuk mencari dan mengganti 'semuanya secara bersamaan (fungsi dipanggil sekali dan 'langsung mengganti semua sub string ybt)
'Created by Rizky Khapidsyah
'Source Code dimulai dari sini

Dim vFixed As Boolean
Function FindReplace(SourceString, Searchstring, _
Replacestring)
  tmpString1 = SourceString
  Do Until vFixed
      tmpString2 = tmpString1
      tmpString1 = ReplaceFirstInstance(tmpString1, _
                   Searchstring, Replacestring)
      If tmpString1 = tmpString2 Then vFixed = True
  Loop
  FindReplace = tmpString1
  Text1.Text = FindReplace
  MsgBox "String tidak ditemukan!", vbCritical, _
         "Tidak Ditemukan"
  Exit Function
End Function

Function ReplaceFirstInstance(SourceString, _
Searchstring, Replacestring)
  Static StartLoc
  If StartLoc = 0 Then StartLoc = 1
  FoundLoc = InStr(StartLoc, SourceString, Searchstring) '*
  If FoundLoc <> 0 Then
     ReplaceFirstInstance = Left(SourceString, _
     FoundLoc - 1) & Replacestring & _
     Right(SourceString, Len(SourceString) - _
     (FoundLoc - 1) - Len(Searchstring))
     StartLoc = FoundLoc + Len(Replacestring)
  Else
     StartLoc = 1
     ReplaceFirstInstance = SourceString
  End If
End Function

'Fungsi sReplace() untuk mencari dan mengganti satu 'string tertentu saja bila fungsi ini dipanggil. String 'berikutnya akan dicari/diganti bila fungsi ini 'dipanggil lagi.

Function sReplace(SearchLine As String, SearchFor As _
String, ReplaceWith As String)

    Dim vSearchLine As String, found As Integer
found = InStr(SearchLine, SearchFor): _
vSearchLine = SearchLine
    If found <> 0 Then
        vSearchLine = ""
        If found > 1 Then vSearchLine = _
           Left(SearchLine, found - 1)
       vSearchLine = vSearchLine + ReplaceWith
        If found + Len(SearchFor) - 1 < _
           Len(SearchLine) Then _
           vSearchLine = vSearchLine + _
           Right$(SearchLine, Len(SearchLine) - _
           found - Len(SearchFor) + 1)
    Else
       MsgBox "String tidak ditemukan!", _
              vbCritical, "Tidak Ditemukan"
       Exit Function
    End If
    sReplace = vSearchLine
    Text1.Text = vSearchLine
End Function

'Dalam contoh ini, kita mengganti setiap sub string "ya" menjadi "yes"
'Mengganti satu per satu berurutan dari atas ke bawah

Private Sub Command1_Click()
  Call sReplace(Text1.Text, "ya", "yes")
End Sub

'Mengganti semuanya secara bersamaan
Private Sub Command2_Click()
  Call FindReplace(Text1.Text, "ya", "yes")
End Sub

Private Sub Form_Load()
  Text1.Text = "Halo apa kabar !" & _
  "ini semoga semuanya dalam keadaan baik " & _
  "dan sehat-sehat selalu tanpa kurang " & _
  "suatu apapun ya kalau begitu ya sudah " & _
  "kita akan tidur malam ini ya atau tidak " & _
  "itu semua terserah Anda karena ada dan tidak " & _
  "ada itu sudah hal yang biasa saja ya kan"
End Sub

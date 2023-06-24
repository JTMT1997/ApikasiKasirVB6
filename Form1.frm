VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Aplikasi Kasir"
   ClientHeight    =   5670
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11370
   LinkTopic       =   "Form1"
   ScaleHeight     =   5670
   ScaleWidth      =   11370
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame5 
      Height          =   615
      Left            =   9480
      TabIndex        =   28
      Top             =   4680
      Width           =   1695
      Begin VB.Label Label9 
         Height          =   255
         Left            =   240
         TabIndex        =   29
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Total Bayar"
      Height          =   375
      Left            =   7440
      TabIndex        =   24
      Top             =   3480
      Width           =   1095
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   8640
      TabIndex        =   19
      Top             =   2760
      Width           =   1215
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   7320
      TabIndex        =   18
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Clear"
      Height          =   375
      Left            =   10200
      TabIndex        =   16
      Top             =   360
      Width           =   1095
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Add Bayar"
      Height          =   375
      Left            =   7200
      TabIndex        =   15
      Top             =   360
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      Height          =   5655
      Left            =   6960
      TabIndex        =   14
      Top             =   0
      Width           =   4455
      Begin VB.TextBox Text9 
         Height          =   285
         Left            =   3000
         TabIndex        =   23
         Top             =   3120
         Width           =   1215
      End
      Begin VB.TextBox Text8 
         Height          =   285
         Left            =   1680
         TabIndex        =   22
         Top             =   3120
         Width           =   1215
      End
      Begin VB.TextBox Text7 
         Height          =   285
         Left            =   360
         TabIndex        =   21
         Top             =   3120
         Width           =   1215
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Left            =   3000
         TabIndex        =   20
         Top             =   2760
         Width           =   1215
      End
      Begin VB.Frame Frame3 
         Caption         =   "Data Beli"
         Height          =   3495
         Left            =   120
         TabIndex        =   17
         Top             =   2160
         Width           =   4215
         Begin VB.Frame Frame4 
            Height          =   615
            Left            =   120
            TabIndex        =   30
            Top             =   2400
            Width           =   1695
            Begin VB.Label Label8 
               Height          =   255
               Left            =   240
               TabIndex        =   31
               Top             =   240
               Width           =   1215
            End
         End
         Begin VB.Label Label7 
            Caption         =   "Label7"
            Height          =   255
            Left            =   240
            TabIndex        =   27
            Top             =   1920
            Width           =   3255
         End
         Begin VB.Label Label6 
            Caption         =   "Label6"
            Height          =   495
            Left            =   1920
            TabIndex        =   26
            Top             =   1320
            Width           =   1575
         End
      End
   End
   Begin VB.Frame Frame1 
      Height          =   5655
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   6735
      Begin VB.Timer Timer1 
         Interval        =   1000
         Left            =   4800
         Top             =   5040
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Exit"
         Height          =   375
         Left            =   4440
         TabIndex        =   13
         Top             =   4200
         Width           =   1095
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Clear"
         Height          =   375
         Left            =   2640
         TabIndex        =   12
         Top             =   4200
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Total Bayar"
         Height          =   375
         Left            =   3600
         TabIndex        =   11
         Top             =   3120
         Width           =   1095
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   4200
         TabIndex        =   9
         Top             =   2520
         Width           =   1335
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   4200
         TabIndex        =   8
         Top             =   2040
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   4200
         TabIndex        =   6
         Top             =   1680
         Width           =   1335
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Tambahkan Jenis Brng"
         Height          =   495
         Left            =   3000
         TabIndex        =   3
         Top             =   120
         Width           =   1335
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "Form1.frx":0000
         Left            =   720
         List            =   "Form1.frx":000A
         TabIndex        =   2
         Text            =   "-----"
         Top             =   120
         Width           =   1815
      End
      Begin VB.ListBox List1 
         Height          =   4155
         ItemData        =   "Form1.frx":001A
         Left            =   360
         List            =   "Form1.frx":001C
         TabIndex        =   1
         Top             =   960
         Width           =   1935
      End
      Begin VB.Label Label5 
         Caption         =   "Label5"
         Height          =   375
         Left            =   3720
         TabIndex        =   25
         Top             =   3720
         Width           =   1695
      End
      Begin VB.Label Label4 
         Caption         =   "Jumlah Beli"
         Height          =   375
         Left            =   2760
         TabIndex        =   10
         Top             =   2520
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Jenis Barang"
         Height          =   375
         Left            =   2760
         TabIndex        =   7
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Harga Barang"
         Height          =   255
         Left            =   2760
         TabIndex        =   5
         Top             =   1680
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Jenis Barang"
         Height          =   735
         Left            =   0
         TabIndex        =   4
         Top             =   120
         Width           =   615
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
List1.Clear
If Combo1.Text = "Rokok" Then
    List1.AddItem "Dji Sam Soe"
    List1.AddItem "Djarum Super"
    List1.AddItem "Starmild"
    List1.AddItem "Neomild"
    List1.AddItem "Signature"
ElseIf Combo1.Text = "Mie" Then
    List1.AddItem "Indomie"
    List1.AddItem "Supermie"
    List1.AddItem "Popmie"
    List1.AddItem "Sarimie"
    List1.AddItem "Mie Sedap"
End If
End Sub
Private Sub Command2_Click()
Label5.Caption = Val(Text1.Text) * Val(Text3.Text)
End Sub

Private Sub Command3_Click()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
End Sub
Private Sub Command4_Click()
End
End Sub
Private Sub Command5_Click()
Dim total As String
total = Label5.Caption
If Text4.Text = "" Then
    Text4.Text = total
ElseIf Text5.Text = "" Then
    Text5.Text = total
ElseIf Text6.Text = "" Then
    Text6.Text = total
ElseIf Text7.Text = "" Then
    Text7.Text = total
ElseIf Text8.Text = "" Then
    Text8.Text = total
ElseIf Text9.Text = "" Then
    Text9.Text = total
Else
    MsgBox "Data Bayar Sudah Penuh!"
End If
End Sub

Private Sub Command6_Click()
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Text8.Text = ""
Text9.Text = ""
End Sub
Private Sub Command7_Click()
Dim total_beli As Double
total_beli = Val(Text4.Text) + Val(Text5.Text) + Val(Text6.Text) + Val(Text7.Text) + Val(Text8.Text) + Val(Text9.Text)
Label6.Caption = total_beli
Label7.Caption = "Terimakasih Telah Berbelanja disini "
End Sub
Private Sub List1_Click()
Dim harga As Double
Dim jenis As String
If Combo1.Text = "Rokok" Then
jenis = "Rokok"
Select Case List1.Text
Case "Dji Sam Soe"
harga = 12000
Case "Djarum Super"
harga = 10000
Case "Starmild"
harga = 11000
Case "Neomild"
harga = 10500
Case "Signature"
harga = 14000
End Select
ElseIf Combo1.Text = "Mie" Then
jenis = "Mie"
Select Case List1.Text
Case "Indomie"
harga = 1500
Case "Supermie"
harga = 1400
Case "Popmie"
harga = 6000
Case "Sarimie"
harga = 1300
Case "Mie Sedap"
harga = 1200
End Select
End If
Text1.Text = harga
Text2.Text = jenis
End Sub
Private Sub Timer1_Timer()
Label8.Caption = Format(Now, "d mmmm yyyy")
Label9.Caption = Format(Now, "hh : mm : ss")
End Sub

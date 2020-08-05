VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000007&
   Caption         =   "Membuat Efek Animasi pada Form"
   ClientHeight    =   3555
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5595
   LinkTopic       =   "Form1"
   ScaleHeight     =   3555
   ScaleWidth      =   5595
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   2040
      TabIndex        =   0
      Top             =   2520
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
  'Ganti '500' di bawah dengan kecepatan dari efek
  'ledakan form.
  Call ImplodeForm(Me, 50)
  End
  Set Form1 = Nothing
End Sub

Private Sub Form_Load()
  Call ExplodeForm(Me, 50)  'ledakan form
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  Call ImplodeForm(Me, 50)  'pengembalian form
End Sub


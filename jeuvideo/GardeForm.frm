VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} GardeForm 
   Caption         =   "UserForm2"
   ClientHeight    =   9315.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9345.001
   OleObjectBlob   =   "GardeForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "GardeForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub fermerBouton_Click()
Unload Me


End Sub

Private Sub okLabel_Click()

End Sub

Private Sub stopBouton_Click()
On Error GoTo erreur
Unload Me
Dim i As Integer, macaron, carnetdebord, badge
i = 1
Dim mycond As Boolean
mycond = False
While Worksheets("vehicles").Range("A" & i) <> "" And mycond = False
If Worksheets("vehicles").Range("D" & i) = UserForm1.mypic.Name Then
UserForm1.Controls.Remove ("grade" & img.Name)
UserForm1.Controls.Remove (img.Name)

Worksheets("vehicles").Rows(i).EntireRow.Delete
mycond = True
nbvehiculearrest = nbvehiculearrest + 1
UserForm1.Controls.Remove ("arme" & x.Name)

End If
Wend
erreur:
Debug.Print "erreur stop bouton"
Unload Me

End Sub

Private Sub UserForm_Initialize()
nbvehiculearrest = 0
Dim i As Integer, macaron, carnetdebord, badge
i = 1
Dim mycond As Boolean
mycond = False
While Worksheets("vehicles").Range("A" & i) <> "" And mycond = False
If Worksheets("vehicles").Range("D" & i) = UserForm1.mypic.Name Then
macaron = Worksheets("vehicles").Range("G" & i)
carnetdebord = Worksheets("vehicles").Range("H" & i)
badge = Worksheets("vehicles").Range("F" & i)
If macaron = "1" Or badge = "1" Or carnetdebord = "1" Then
stopBouton.Visible = True
fermerBouton.Visible = False
If macaron = "1" Then
MacaronOptionButton.Visible = True

End If
If badge = "1" Then
BadgeOptionButton.Visible = True
End If
If carnetdebord = "1" Then
CarnetDeBordOptionButton.Visible = True
End If


Else
stopBouton.Visible = False
fermerBouton.Visible = True
okLabel.Visible = True
MacaronOptionButton.Visible = False

BadgeOptionButton.Visible = False
CarnetDeBordOptionButton.Visible = False
End If

mycond = True
End If
Wend

End Sub

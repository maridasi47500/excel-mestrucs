Attribute VB_Name = "soldat"
Public form1 As UserForm
Sub autreshelicos()
UserForm1.sethelicos = True

End Sub
Sub nouvelhelico()
Dim nom_str
Dim x As Integer
Dim y As Integer
Dim imgheight As Integer
imgheight = 150
y = 1
Dim badge, carnetdebord, macaron
Dim lowerbound, upperbound
Dim myheliconom(1 To 100) As String
Dim myhelicoimage(1 To 100) As String
Dim chemabs
chemabs = Sheets("le_cheminabsolu").Range("I10")


'x = Int((upperbound - lowerbound + 1) * Rnd + lowerbound)
Debug.Print x & "= nombre entre 23 et 96"
Dim img As MSForms.Image
Dim imstr As String


Dim gradeimg As MSForms.Image
Set img = UserForm1.Controls.Add("forms.image.1")

Dim lowerb, upperb
lowerb = 23
upperb = 96

Debug.Print imstr & " = image"
For x = lowerb To upperb
imstr = Worksheets("grades").Range("A" & x)
img_str = Replace(Replace(Replace(Replace(imstr, "%20", " "), "%C3%A0", "à"), "-de-petite-capacite", ""), "%C3%A8", "è")
nom_str = Worksheets("grades").Range("B" & x)
If InStr(nom_str, "hélico") Or InStr(nom_str, "Hélico") Or InStr(nom_str, "systeme-mini-drone") Or InStr(nom_str, "Cougar") Or InStr(nom_str, "CARACAL") Or InStr(nom_str, "SDT") Or InStr(nom_str, "Pilatus") Or InStr(nom_str, "GAZELLE") Or InStr(nom_str, "mini-drone") Then
myheliconom(y) = nom_str
myhelicoimage(y) = img_str
y = y + 1
End If
Next
lowerbound = 1
upperbound = y - 1
Randomize
x = Int((upperbound - lowerbound + 1) * Rnd + lowerbound)

img_str = myhelicoimage(x)
nom_str = myheliconom(x)
Set img = UserForm1.Controls.Add("forms.image.1")
With img
.PictureSizeMode = fmPictureSizeModeStretch
.Width = 200
.Height = imgheight
.Picture = LoadPicture(chemabs & img_str)
.ControlTipText = nom_str
.Top = 0
.Tag = "helico"
End With

End Sub
Sub vehicules()
On Error GoTo erreur
Dim chemabs
chemabs = Sheets("le_cheminabsolu").Range("I10")
Dim section(1 To 9)   As String
Dim MyValue As Integer
section(1) = "E"
section(2) = "F"
section(3) = "G"
section(4) = "H"
section(5) = "I"
section(6) = "J"
section(7) = "K"
section(8) = "L"
section(9) = "M"
Dim nom_str As String
Dim img_str As String
Dim grade As Integer
Randomize
grade = Int((16 * Rnd) + 1)

Dim sectionid As Integer
Randomize
sectionid = Int((9 * Rnd) + 1)
'''
While Worksheets("grades").Range(section(sectionid) & grade) = "0"
Randomize
grade = Int((16 * Rnd) + 1)
Randomize
sectionid = Int((9 * Rnd) + 1)
Wend
'''
Dim x As Integer
Dim badge, carnetdebord, macaron
Dim lowerbound, upperbound
lowerbound = 23
upperbound = 96
Randomize

x = Int((upperbound - lowerbound + 1) * Rnd + lowerbound)
Debug.Print x & "= nombre entre 23 et 96"
Dim img As MSForms.Image
Dim imstr As String

imstr = Worksheets("grades").Range("A" & x)
Dim gradeimg As MSForms.Image
Set img = UserForm1.Controls.Add("forms.image.1")



Debug.Print imstr & " = image"

img_str = Replace(Replace(Replace(Replace(imstr, "%20", " "), "%C3%A0", "à"), "-de-petite-capacite", ""), "%C3%A8", "è")
nom_str = Worksheets("grades").Range("B" & x)
Debug.Print nom_str & "= nom"
Debug.Print chemabs & img_str
Dim imgheight As Integer
imgheight = 150
Debug.Print imgheight & " = image height"
Debug.Print "picture size"
Debug.Print "width"
Debug.Print "height"
Debug.Print "picture"
With img
.PictureSizeMode = fmPictureSizeModeStretch
.Width = 200
.Height = imgheight
.Picture = LoadPicture(chemabs & img_str)
.ControlTipText = nom_str
.Top = UserForm1.Height - imgheight
.Tag = "vehicule"
End With
Debug.Print "vehicule 1 vehicule"
Debug.Print img.Name


Debug.Print "control tip text"
Debug.Print "tag"
Debug.Print "user form1"
Debug.Print "top"
Debug.Print "userform 2"

Debug.Print img.Name & " = nom image"
UserForm1.MYEventsrsma.Add vehicule1
Debug.Print " = evenemen ajoute"
UserForm1.myid = UserForm1.myid + 1
Debug.Print " = id + 1"
Set UserForm1.mypic = UserForm1.Controls(img.Name)
Debug.Print " = my pic set"
Debug.Print img.Name & ": nom image"

'''
If InStr(nom_str, "hélico") Or InStr(nom_str, "Hélico") Or InStr(nom_str, "systeme-mini-drone") Or InStr(nom_str, "Cougar") Or InStr(nom_str, "CARACAL") Or InStr(nom_str, "SDT") Or InStr(nom_str, "Pilatus") Or InStr(nom_str, "GAZELLE") Or InStr(nom_str, "mini-drone") Then
With img
.Top = 0
.Tag = "helico"
End With
End If
Set gradeimg = UserForm1.Controls.Add("forms.image.1")
Debug.Print Worksheets("grades").Range("B" & grade)

With gradeimg
.Picture = LoadPicture(Worksheets("grades").Range("B" & grade))
.Left = img.Left + img.Width - 30
.Top = img.Top
.Width = 30
.Height = 30
.PictureSizeMode = fmPictureSizeModeStretch
.Name = "grade" & img.Name
.Tag = "grade"
End With
Debug.Print ("ok grade image")
Randomize
MyValue = Int((2 * Rnd) + 4)

Dim z As Integer
z = 1
While Worksheets("vehicles").Range("A" & z) <> ""
z = z + 1
Wend
Worksheets("vehicles").Range("A" & z) = grade
Worksheets("vehicles").Range("B" & z) = sectionid
Worksheets("vehicles").Range("C" & z) = x
Worksheets("vehicles").Range("D" & z) = img.Name
Worksheets("vehicles").Range("E" & z) = gradeimg.Name
lowerbound = 0
upperbound = 1
Randomize
Worksheets("vehicles").Range("F" & z) = Int((upperbound - lowerbound + 1) * Rnd + lowerbound)
Randomize
Worksheets("vehicles").Range("G" & z) = Int((upperbound - lowerbound + 1) * Rnd + lowerbound)
Randomize
Worksheets("vehicles").Range("H" & z) = Int((upperbound - lowerbound + 1) * Rnd + lowerbound)
Application.OnTime Now + TimeValue("0:0:" & MyValue), "vehicules"
Exit Sub


erreur:
Debug.Print "my string"
Debug.Print Err.Description
MsgBox ("ERREUR : " & Err.Description)
'Randomize
'MyValue = Int((2 * Rnd) + 6)
'Dim mycond As Boolean
'mycond = False
'Application.OnTime Now + TimeValue("0:0:" & MyValue), "vehicules"

'z = 1
'While Worksheets("vehicles").Range("A" & z) <> "" And mycond = False
'If Worksheets("vehicles").Range("D" & z) = img.Name Then
'Worksheets("vehicles").Rows(z).EntireRow.Delete
'UserForm1.Controls.Remove (img.Name)
'mycond = True
'End If
'z = z + 1
'Wend


End Sub

Sub shoot()
Dim x As Integer, y As Integer
Dim muniton As MSForms.Image
Dim chemabs
chemabs = Sheets("le_cheminabsolu").Range("I10")
Dim Vehicle(1 To 100) As MSForms.Image, helico(1 To 100) As Control, a As Control
Dim upperbound, lowerbound, value
upperbound = 20
lowerbound = 10
value = Int((upperbound - lowerbound + 1) * Rnd + lowerbound)
x = 1
y = 1
For Each a In UserForm1.Controls
If TypeOf a Is Image And a.Tag = "vehicule" Then
Set Vehicle(x) = a
Debug.Print "nom du véhciule : " & a.Name

x = x + 1
End If

Next
Dim z As Integer
Randomize

lowerbound = 1
upperbound = x - 1
Randomize

z = Int((upperbound - lowerbound + 1) * Rnd + lowerbound)
Dim monarme As MSForms.Image
Dim monvehicule As MSForms.Image
Dim imstr As String
Set monvehicule = Vehicle(z)
Dim j As MSForms.Image

Dim w As Integer
If x > 1 Then
Randomize
w = Int((4 * Rnd) + 1)
Debug.Print Worksheets("armes").Range("B" & w)
Dim vehiculename As String
Dim findmunition As Image
Debug.Print monvehicule.Left - 40
imstr = Worksheets("armes").Range("B" & w)
vehiculename = Replace(Replace(Replace(Replace(imstr, "%20", " "), "%C3%A0", "à"), "-de-petite-capacite", ""), "%C3%A8", "è")
Set j = UserForm1.Controls.Add("forms.image.1")
With j
.Picture = LoadPicture(chemabs & vehiculename)
.PictureSizeMode = fmPictureSizeModeStretch
.Left = monvehicule.Left - 40
.Width = 80
.Top = monvehicule.Top + monvehicule.Height / 3
.Height = monvehicule.Height / 3
End With
Dim nbmunition As Integer
nbmunition = 0

For Each Control In UserForm1.Controls
If Control.Name = "munition" Then
nbmunition = nbmunition + 1
End If

Next


If nbmunition = 0 Then

Set munition = UserForm1.Controls.Add("forms.image.1")
With munition
.Picture = LoadPicture(chemabs & "munition.jpg")
.PictureSizeMode = fmPictureSizeModeStretch

End With
With munition
.Left = j.Left - munition.Width
.Top = j.Top
.Tag = "munition"
.Name = "munition"
End With

End If

End If

Application.OnTime Now + TimeValue("0:0:" & value), "shoot"

End Sub

Sub helico1()
Dim chemabs
chemabs = Sheets("le_cheminabsolu").Range("I10")
Dim x As Integer, y As Integer
Dim muniton As MSForms.Image
Dim Vehicle(1 To 100) As MSForms.Image, helico(1 To 100) As Control, a As Control
Dim upperbound, lowerbound, value
upperbound = 30
lowerbound = 20
Application.OnTime "helico1", Now + TimeValue("0:0:" & value)
If Int(UserForm1.nbvehiculearrestLabel) < 3 Then
Exit Sub
End If

value = Int((upperbound - lowerbound + 1) * Rnd + lowerbound)
x = 1
y = 1
For Each a In UserForm1.Controls
If TypeOf a Is Image And a.Tag = "helico" Then
Set Vehicle(x) = a
Debug.Print "nom du véhciule : " & a.Name

x = x + 1
End If

Next
Dim z As Integer
Randomize
Dim lowerbound, upperbound
lowerbound = 1
upperbound = x - 1
Randomize

z = Int((upperbound - lowerbound + 1) * Rnd + lowerbound)
Dim monarme As MSForms.Image
Dim monvehicule As MSForms.Image
Dim imstr As String
Set monvehicule = Vehicle(z)
Dim j As MSForms.Image

Dim w As Integer

Randomize
w = Int((4 * Rnd) + 1)
Debug.Print Worksheets("armes").Range("B" & w)
Dim vehiculename As String
imstr = Worksheets("armes").Range("B" & w)
vehiculename = Replace(Replace(Replace(Replace(imstr, "%20", " "), "%C3%A0", "à"), "-de-petite-capacite", ""), "%C3%A8", "è")
Set j = UserForm1.Controls.Add("forms.image.1")
With j
.Picture = LoadPicture(chemabs & vehiculename)
.PictureSizeMode = fmPictureSizeModeStretch
.Left = monvehicule.Left - 40
.Width = 80
.Top = monvehicule.Top + monvehicule.Height / 3
.Height = monvehicule.Height / 3
End With
Set munition = UserForm1.Controls.Add("forms.image.1")
With munition
.Picture = LoadPicture(chemabs & "parachute.jpg")
.PictureSizeMode = fmPictureSizeModeStretch

End With
With munition
.Left = j.Left + 20
.Top = j.Top + j.Height
.Tag = "parachute"

End With


End Sub

Sub avance()
On Error Resume Next
UserForm1.nbsec = UserForm1.nbsec + 1

Dim x As Control
Dim z As Integer
Dim nbmun As Integer
nbmun = 0
Dim yy As Integer
Dim mycond1 As Boolean
Dim imagevehicule As String
mycond1 = False
yy = 1
z = 1


Dim mycond As Boolean
Dim xx As MSForms.Image

mycond = False
For Each x In UserForm1.Controls 'pour chaque objet de formulaire
imagevehicule = x.Name
If x.Tag = "parachute" Then 'si l'image est un parachuet
If x.Top > UserForm1.Height Then
UserForm1.Controls.Remove (x.Name)
Else
x.Top = x.Top + 100
End If
ElseIf x.Tag = "munition" Then ' lsi l'image est une munition

If x.Left < 0 Then  ' lsi l'image est une munition et on l voit plus
UserForm1.Controls.Remove (imagevehicule)
Else  ' lsi l'image est une munition visible
x.Left = x.Left - 100
nbmun = nbmun + 1
 End If ' fin si x plus petit que 0( munition )
ElseIf TypeOf x Is Image Then 'si l'image est un vehicule militaire
If x.Left > UserForm1.Width Then  'si l'image est un vehicule militaire et on ne le voit plus


While Worksheets("vehicles").Range("A" & z) <> "" And mycond = False 'chercher ce vehicule mlitaire sur excel
If Worksheets("vehicles").Range("D" & z) = imagevehicule Then
Worksheets("vehicles").Rows(z).EntireRow.Delete 'supprimer ce vehicule militaire dans excel

mycond = True
End If
z = z + 1
Wend 'fin chercher ce vehicule mlitaire sur excel

UserForm1.Controls.Remove (imagevehicule)
UserForm1.Controls.Remove ("grade" & imagevehicule)
UserForm1.Controls.Remove ("arme" & imagevehicule)
Else  'si l'image est un vehicule militaire et on le voit toujours
x.Left = x.Left + 100
End If 'fin si l'image est un vehicule militaire et on ne le voit plus/on le voit toujours
End If 'fin si l'image est un vehicule militaire/parachut/munition





Next 'fin pour chaque objet de formulaire

'''

If nbmun > 0 Then ' si il y a des munitions
Debug.Print "si il y a des munitions"

For Each Control In UserForm1.Controls ' pour chaque objet de formulaire si il y a une munition
Debug.Print "pour chaque objet de formulaire si il y a une munition"
If Control.Name <> "menu" And Control.Tag <> "helico" And Control.Name <> "tag" And Control.Tag <> "parachute" And Control.Tag <> "grade" Then 'si c'est un véhicule militaire
''''----
Debug.Print "si c'est un véhicule militaire"
Set x = UserForm1.Controls("munition") 'x est la munition
Set xx = Control 'xx est le véhicule militaire
imagevehicule = xx.Name
If x.Left > xx.Left And x.Left < (xx.Left + xx.Width) Then ' if projectile touche vehicule militaire
Debug.Print "projectile touche vehicule militaire"
UserForm1.Controls.Remove (imagevehicule)
While Worksheets("vehicles").Range("D" & yy) <> "" And mycond1 = False 'chercher le vehicule
Debug.Print "chercher le vehicule"
If Worksheets("vehicles").Range("D" & yy) = imagevehicule Then 'si le vehicule est trouvé
Debug.Print "véhicule est trouvé"
Dim mycellrow As Integer
Dim mycellcol As Integer
mycellrow = Int(Worksheets("vehicles").Range("A" & yy))
mycellcol = Int(Worksheets("vehicles").Range("B" & yy))
If (Int(Worksheets("grades").Cells(mycellrow, mycellcol)) - 1) <> Int(0) Then 'si la personne de ce grade et de ce service est trouvée
Debug.Print "personne de ce service est trouvee"
Worksheets("grades").Cells(mycellrow, mycellcol) = Int(Worksheets("grades").Cells(mycellrow, mycellcol)) - 1 'retirer une personne de ce grade dans le tableua excel

End If 'fin si la personne de ce grade et de ce service est trouvée
Worksheets("vehicles").Rows(yy).EntireRow.Delete

Debug.Print "ligne excel supprimee"
UserForm1.Controls.Remove ("grade" & imagevehicule)
Debug.Print "grade supprime du formulaire"

mycond1 = True
Debug.Print "arme"
UserForm1.Controls.Remove ("arme" & imagevehicule)
Debug.Print "arme supprimee"
Debug.Print "ok vehicule trouve"
End If 'fin si le vehicule est trouvé



yy = yy + 1
Wend 'fin chercher le vehicule

End If 'fin if projectile touche vehicule militaire


''''---
End If 'fin si c'est un véhicule militaire


Next ' fin pour chaque objet de formulaire si il y a une munition

End If 'fin si il y a des munitions


 

'''


Application.OnTime Now + TimeValue("0:0:1"), "avance"
Exit Sub

erreur:

MsgBox (Err.Description)
End Sub

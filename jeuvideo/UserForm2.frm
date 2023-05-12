VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm2 
   Caption         =   "UserForm2"
   ClientHeight    =   6840
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4470
   OleObjectBlob   =   "UserForm2.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
On Error Resume Next    ' Hide the Userform and set cancelled to true
   
    Dim g As Integer
    For g = 1 To 20
    
      Application.OnTime Now + TimeValue("0:0:" & g), "avance", , False
  Application.OnTime Now + TimeValue("0:0:" & g), "vehicules", , False
 Application.OnTime Now + TimeValue("0:0:" & g), "shoot", , False
    Next
     Hide
    m_Cancelled = True
    End
End Sub

Private Sub CommandButton3_Click()
Unload Me
End Sub

Private Sub CommandButton4_Click()
UserForm1.sethelicos = False
Application.OnTime Now + TimeValue("0:0:1"), "nouvelhelico"
Application.OnTime Now + TimeValue("0:0:10"), "nouvelhelico"
Application.OnTime Now + TimeValue("0:0:15"), "autreshelicos"
'Application.OnTime Now + TimeValue("0:0:15"), "nouvelhelico"
'Application.OnTime Now + TimeValue("0:0:20"), "nouvelhelico"
'Application.OnTime Now + TimeValue("0:0:25"), "nouvelhelico"
'Application.OnTime Now + TimeValue("0:0:30"), "autreshelicos"
Unload Me
End Sub

Private Sub UserForm_Activate()
If UserForm1.gethelicos = True Then
CommandButton4.Visible = True
renfortLabel.Visible = False
Else
CommandButton4.Visible = False
renfortLabel.Visible = True
End If
End Sub

Private Sub UserForm_Click()

End Sub

Private Sub UserForm_initialize()
If UserForm1.gethelicos = True Then
CommandButton4.Visible = True
renfortLabel.Visible = False
Else
CommandButton4.Visible = False
renfortLabel.Visible = True
End If
End Sub
Private Sub CommandButton2_Click()
'e17 a e19 et e20
Dim section(1 To 9)   As String
Dim sectionid As Integer
Dim grade As Integer
section(1) = "E"
section(2) = "F"
section(3) = "G"
section(4) = "H"
section(5) = "I"
section(6) = "J"
section(7) = "K"
section(8) = "L"
section(9) = "M"
Dim appel As String
appel = "compagnie garde à vous par ordre des sections présentes faites et rendez l'appel compagnie repos "
Dim x As Integer
x = 1
Dim mdr, sousofficier, off, servicenom, sum
While Worksheets("grades").Range(section(x) & 17) <> ""

mdr = Worksheets("grades").Range(section(x) & 17)
sousofficier = Worksheets("grades").Range(section(x) & 18)
off = Worksheets("grades").Range(section(x) & 19)
sum = Int(mdr) + Int(sousofficier) + Int(off)

servicenom = Worksheets("grades").Range(section(x) & 20)
If sum > 0 Then
appel = appel & " " & servicenom & " garde à vous. effectif réalisé " & mdr & " " & sousofficier & " " & off & " effectif sur les rangs " & mdr & " " & sousofficier & " " & off & " appel rendu section repos"
Else
appel = appel & " " & sectionnom & " garde à vous l'appel sera rendu à l'issue du rapport repos."
End If
x = x + 1
Wend
appel = appel & " compagnie garde à vous à disposition des chefs de section"
MsgBox (appel)
End Sub

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MyVehicle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public WithEvents vehicule1 As MSForms.Image
Attribute vehicule1.VB_VarHelpID = -1
Public WithEvents mypicform1 As MSForms.Image
Attribute mypicform1.VB_VarHelpID = -1
Public WithEvents form2 As MSForms.UserForm
Attribute form2.VB_VarHelpID = -1




Private Sub vehicule1_Click()
 Debug.Print vehicule1.Name & " = me vehicule"
 


Set VBA.UserForms("UserForm1").mypic = vehicule1
Debug.Print vehicule1.Name

VBA.UserForms("MyForm").Show
Debug.Print frm.Name
End Sub

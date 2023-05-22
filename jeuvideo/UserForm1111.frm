VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   9510.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14220
   OleObjectBlob   =   "UserForm1111.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public mypic As MSForms.Image

Public MYEventsrsma As Collection
Public myid As Integer
Public nbvehiculearrest As Integer
Public nbsec As Integer
Private Sub UserForm_initialize()
nbsec = 0
Set MYEventsrsma = New Collection
myid = 1

Worksheets("vehicles").Cells.Clear
Worksheets("vehicles").Range("A1") = "grade"
Worksheets("vehicles").Range("B1") = "section"
Worksheets("vehicles").Range("C1") = "ligne vehicule"
Worksheets("vehicles").Range("D1") = "image vehicule"
Worksheets("vehicles").Range("E1") = "image grade"
Worksheets("vehicles").Range("F1") = "badge"
Worksheets("vehicles").Range("G1") = "macaron"
Worksheets("vehicles").Range("H1") = "carnet de bord"
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
For grade = 1 To 16
For sectionid = 1 To 9
Randomize
Worksheets("grades").Range(section(sectionid) & grade) = Int((6 * Rnd))
Next
Next
''
Dim img As String
 Dim nom As String
 Dim ligne As Integer
 ligne = 23
 
 img = "https://www.defense.gouv.fr/sites/default/files/styles/16_9_sm/public/terre/"
 nom = "fr-h5 fr-m-0 fr-pb-1w"
Dim myFile As String, text As String, textline As String, posLat As Integer, posLong As Integer
Dim cola As Integer
cola = 1
text = ""
Dim cheminabs
cheminabs = Sheets("le_cheminabsolu").Range("I10")
myFile = cheminabs & "mesvehicules.txt.txt"
Dim textFileNum, rowNum, colNum As Integer
Dim textFileLocation, textDelimiter, textData As String
Dim tArray() As String
Dim sArray() As String
Dim arrSplitStrings1() As String
Dim arrSplitStrings2() As String
Dim strSingleString1 As String
Dim strSingleString2 As String
Dim strSingleString3 As String

Dim x As String
Dim i As Long
textFileLocation = myFile
textDelimiter = "img"
textFileNum = FreeFile
Open textFileLocation For Input As textFileNum
textData = Input(LOF(textFileNum), textFileNum)
Close textFileNum

 ligne = 23
cola = 1
 Dim DateiName As String
    Dim ReplacePrep As String
    Dim LineFromFile As String
    Dim LineItems As Variant
    Dim row_number As Long

    Dim objStream As Object

    DateiName = myFile

    Set objStream = CreateObject("ADODB.Stream")
    objStream.Charset = "utf-8"
    objStream.Open
    objStream.LoadFromFile (DateiName)

   row_number = 10

    Do Until objStream.EOS                       'Go through the entire text document
        LineFromFile = objStream.ReadText(-2)    'Read line from source file
        
If (InStr(LineFromFile, img) Or InStr(LineFromFile, nom)) And Not InStr(LineFromFile, "capacites-terre") Then

If cola = 2 Then
Debug.Print LineFromFile
arrSplitStrings1 = Split(LineFromFile, ">")

arrSplitStrings2 = Split(arrSplitStrings1(1), "<")
strSingleString1 = arrSplitStrings2(0)
Worksheets("grades").Cells(ligne, colNum + cola) = strSingleString1
cola = 1
ligne = ligne + 1

Else
Debug.Print LineFromFile
arrSplitStrings1 = Split(LineFromFile, "/")
arrSplitStrings2 = Split(arrSplitStrings1(UBound(arrSplitStrings1, 1)), "?")
strSingleString1 = arrSplitStrings2(0)
Worksheets("grades").Cells(ligne, colNum + cola) = Replace(strSingleString1, "png", "jpg")
cola = 2

End If


End If


   Loop
    Set objStream = Nothing

MsgBox "Data Imported Successfully", vbInformation
  Application.OnTime Now + TimeValue("0:0:1"), "avance"
  Application.OnTime Now + TimeValue("0:0:1"), "vehicules"
 Application.OnTime Now + TimeValue("0:0:5"), "shoot"
  Application.OnTime Now + TimeValue("0:1:0"), "nbvehicule"
End Sub

Sub ImportABunch()

Dim strTemp As String
Dim strPath As String
Dim strFileSpec As String
Dim oSld As Slide
Dim oPic As Shape
Dim slideIndex As Integer
Dim slidePosition As Integer
Dim incrementLeft As Single
Dim incrementTop As Single

' Edit these to suit:
'strPath = "C:\Documents and Settings\conrad\Desktop\"
'strPath = Replace(WScript.ScriptFullName, WScript.ScriptName, "")
'turns out if strpath is blank it will get the current directory.
strFileSpec = "*.jpg"

strTemp = Dir(strPath & strFileSpec)

slideIndex = 2 ' 1=cover, 2,3,4... slides to add pictures to
slidePosition = 0 ' 0=top, 1=mid, 2=bot
incrementLeft = 49

Do While strTemp <> ""
    
    
    Set oSld = ActivePresentation.Slides(slideIndex)
    Set oPic = oSld.Shapes.AddPicture(FileName:=strPath & strTemp, _
    LinkToFile:=msoFalse, _
    SaveWithDocument:=msoTrue, _
    Left:=0, _
    Top:=0, _
    Width:=100, _
    Height:=100)

    Select Case slidePosition
    Case 0
        incrementTop = 127
    Case 1
        incrementTop = 343
    Case 2
        incrementTop = 559
    End Select

    ' Reset size to the size we want, only works if 1200 is max dimension and proper scale.
    With oPic
        .ScaleHeight 0.315, msoTrue
        .ScaleWidth 0.315, msoTrue
        .Width = 286.87
        .Height = 214.87
        
    End With
    With oPic
        .incrementLeft incrementLeft
        .incrementTop incrementTop
    End With


    ' Get the next file that meets the spec and go round again
    strTemp = Dir
    
    slidePosition = slidePosition + 1
    
    If (slidePosition = 3) Then
        slidePosition = 0
        slideIndex = slideIndex + 1
    End If
    
Loop

End Sub

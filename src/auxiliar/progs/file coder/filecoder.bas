Attribute VB_Name = "Module1"
Option Explicit

Public Function Cifrar(Nombre As String) As String
Dim n As Integer
Cifrar = ""
For n = 1 To Len(Nombre)
    Cifrar = Cifrar & Chr(255 - Asc(Mid$(Nombre, n, 1)))
    DoEvents
Next
End Function

Public Function APath(ByVal Nombre As String) As String
Dim xxx As Integer, xx As Integer
If Nombre = "" Then Exit Function
For xxx = Len(Nombre) To 1 Step -1
    xx = InStr(xxx, Nombre, "\")
    If xx <> 0 Then Exit For
Next
APath = Mid$(Nombre, 1, xx - 1)
End Function

Public Function NombreR(ByVal Nombre As String) As String
Dim xxx As Integer, xx As Integer
If Nombre = "" Then Exit Function
For xxx = Len(Nombre) To 1 Step -1
    xx = InStr(xxx, Nombre, ".")
    If xx <> 0 Then Exit For
Next
NombreR = Mid$(Nombre, xx + 1, Len(Nombre) - xx + 1)
End Function

Public Function NombreArchivo(ByVal Nombre As String) As String
Dim xxx As Integer, xx As Integer
If Nombre = "" Then Exit Function
For xxx = Len(Nombre) To 1 Step -1
    xx = InStr(xxx, Nombre, "\")
    If xx <> 0 Then Exit For
Next
NombreArchivo = Mid$(Nombre, xx + 1, Len(Nombre) - xx + 1)
End Function



Public Function ExtractFile(ByVal File As String, ByVal Dir As String) As Boolean
On Local Error GoTo Out
Dim Path As String
Path = Dir
If Right(Path, 1) <> "\" Then Path = Path & "\"
Dim ff As Integer, ff2 As Integer, x As Long
Dim cad As String, ln As Long, max As Long, num As Long, cad2 As String
ff = FreeFile
Open File For Binary As #ff
    cad = Space(14)
    Get #ff, , cad
    Get #ff, , num
    
    For x = 1 To num
        Get #ff, , ln
        cad = Space(ln)
        Get #ff, , cad
        cad = Code(cad)
        ff2 = FreeFile
        Open Path & "\" & cad For Binary As #ff2
            Get #ff, , ln
            If ln <= 1127 Then
                cad = Space(ln)
                Get #ff, , cad
                cad = Code(cad)
                Put #ff2, , cad
            Else
                cad2 = Space(ln - 1127)
                Get #ff, , cad2
                cad = Space(1127)
                Get #ff, , cad
                cad = Code(cad)
                Put #ff2, , cad
                Put #ff2, , cad2
            End If
        Close #ff2
    Next
Close #ff
ExtractFile = True
Out:
End Function

Public Function Code(var As String) As String
Dim n As Integer
Code = ""
For n = 1 To Len(Code)
    Code = Code & Chr(255 - Asc(Mid$(var, n, 1)))
Next
End Function


Public Sub DecodeFile(ByVal File As String, ByVal File2 As String)
On Local Error Resume Next
Kill File2
On Local Error GoTo Out:
Dim ff As Integer, ff2 As Integer
Dim tam As Long, x As Long
Dim by As Byte, by2 As Long
ff = FreeFile
Open File For Binary As #ff
ff2 = FreeFile
Open File2 For Binary As #ff2
tam = FileLen(File)
For x = 1 To tam
    If (x Mod 2) = 1 Then
        Get #ff, , by
        by = 255 - by
        Put #ff2, , by
    Else
        Get #ff, , by
        by2 = by
        by2 = by2 - 121
        If by2 < 0 Then by2 = by2 + 256
        by = by2
        Put #ff2, , by
    End If
Next
Close #ff
Close #ff2
Out:
End Sub

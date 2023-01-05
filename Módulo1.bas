Attribute VB_Name = "Módulo1"
Option Explicit
Dim Fs As New Scripting.FileSystemObject
Public Ar As String
Public Ps As String

Sub GerarTelefones()

    Dim QtdDoadores As Long
    Dim i As Long
    
    QtdDoadores = InputBox("Digite a quantidade de doadores à gerar: ")
    
    Planilha1.Range("A1").CurrentRegion.ClearContents

    
    For i = 1 To QtdDoadores
    
        Planilha1.Cells(i, 1).Value = 14
        Planilha1.Cells(i, 2).Value = WorksheetFunction.RandBetween(911111111, 999999999)
    
    Next

End Sub

Public Function verfica_arq()
    
 Ps = ThisWorkbook.Path & "\csv"
 Ar = Ps & "\" & "exportacao - " & Format(Date, "dd-mm-yyyy") & ".csv"
 
 If Not Fs.FolderExists(Ps) Then
 
    Fs.CreateFolder (Ps)
 
 End If
 
 If Not Fs.FileExists(Ar) Then
 
    Fs.CreateTextFile (Ar)
    
 End If

End Function

Sub GeraCSV()


Call verfica_arq

Dim Linha As Long

Linha = 1

With Planilha1

    Open Ar For Output As #1
        
        Do Until .Cells(Linha, 1) = ""
            
            Print #1, .Cells(Linha, 1).Value & ", " & .Cells(Linha, 2)
            
            Linha = Linha + 1
            
        Loop
        
        MsgBox "Exportado com sucesso", vbInformation, "Htec"
        
    Close #1
    
End With

End Sub

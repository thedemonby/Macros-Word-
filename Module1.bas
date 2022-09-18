Attribute VB_Name = "Module1"
Public Sub Autoopen()

    Dim Num_str As Integer
    Dim Num_rl As Integer
    
    Num_str = 0
    Num_rl = 0
    
    Num_str = InputBox("Введите количество страниц")
    Num_rl = InputBox("Введите номер РЛ")
    
    For i = 1 To Num_str - 1
        Selection.EndKey Unit:=wdStory
        Selection.InsertBreak Type:=wdPageBreak
    Next i
       
    ActiveDocument.Variables("Num_rl_p").Delete

    ActiveDocument.Variables.Add Name:="Num_rl_p", Value:=Num_rl
    
    ActiveDocument.Sections(1).Footers(1).Range.Fields.Update

End Sub

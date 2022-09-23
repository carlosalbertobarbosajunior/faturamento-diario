Attribute VB_Name = "Módulo1"
Sub exportpic()


Dim WS As Worksheet, Inpt As Worksheet
Dim rgExp As Range
Dim CH As ChartObject





Set Inpt = Sheets("Faturamento")
Set rgExp = Inpt.Range("A1:Q28")


    For Each WS In ThisWorkbook.Sheets
        If WS.Name = "Faturamento" Then
            WS.Range("A1:Q28").CopyPicture Appearance:=xlScreen, Format:=xlBitmap
            Set CH = WS.ChartObjects.Add(Left:=rgExp.Left, Top:=rgExp.Top, Width:=rgExp.Width, Height:=rgExp.Height)
            CH.Chart.ChartArea.Select
            CH.Chart.Paste
            CH.Chart.Export "C:\Users\carlos.junior\Desktop\FaturamentoDiario\" & WS.Name & ".jpg"
            CH.Delete


        End If
    Next WS


End Sub

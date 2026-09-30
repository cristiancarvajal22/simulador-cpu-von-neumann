Attribute VB_Name = "SetupProyecto"
Sub TransformarSimulador()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Simulador")
    ws.Unprotect
    On Error Resume Next
    ThisWorkbook.Sheets("Log").Name = "Registros"
    Application.DisplayAlerts = False
    ThisWorkbook.Sheets("Teoria").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    ws.Shapes("PANEL_UC").Fill.ForeColor.RGB = RGB(225, 240, 255)
    ws.Shapes("PANEL_UAL").Fill.ForeColor.RGB = RGB(255, 235, 225)
    ws.Shapes("PANEL_MEM").Fill.ForeColor.RGB = RGB(230, 255, 230)
    ws.Shapes("PRIM_BOX").Fill.ForeColor.RGB = RGB(230, 255, 230)
    ws.Shapes("BTN_LOAD").Fill.ForeColor.RGB = RGB(70, 130, 180)
    ws.Shapes("BTN_STEP").Fill.ForeColor.RGB = RGB(60, 179, 113)
    ws.Shapes("BTN_INSTR").Fill.ForeColor.RGB = RGB(60, 179, 113)
    ws.Shapes("BTN_RUN").Fill.ForeColor.RGB = RGB(147, 112, 219)
    ws.Shapes("BTN_PAUSE").Fill.ForeColor.RGB = RGB(244, 164, 96)
    ws.Shapes("BTN_RESET").Fill.ForeColor.RGB = RGB(205, 92, 92)
    Dim shp As Shape
    For Each shp In ws.Shapes
        shp.Placement = 3
    Next shp
    ws.Columns("I:X").Cut
    ws.Columns("B:B").Insert Shift:=xlToRight
    ws.Columns("R:T").Cut
    ws.Columns("Z:Z").Insert Shift:=xlToRight
    ws.Columns("Z:Z").ColumnWidth = 7.67
    ws.Columns("AA:AA").ColumnWidth = 39.22
    ws.Columns("AB:AB").ColumnWidth = 11.11
    ws.Shapes("BTN_EJEMPLO").Left = ws.Columns("AA:AA").Left + 10
    ws.Shapes("LBL_MEMTIT").Left = ws.Columns("I:I").Left
    ws.Shapes("MEMW_TD").Left = ws.Columns("J:J").Left
    ws.Shapes("MEMW_TH").Left = ws.Columns("L:L").Left
    ws.Shapes("MEMW_TB").Left = ws.Columns("M:M").Left
End Sub

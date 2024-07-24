Attribute VB_Name = "DoAll"
Sub Main()

Call Levels.LevelsByIndex
Call Calculation.Main
Call Consolidation.Main
Call Products.Main
ThisWorkbook.Worksheets("Расшифровка").Activate

End Sub

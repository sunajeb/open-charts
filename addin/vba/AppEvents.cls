Attribute VB_Name = "AppEvents"
' PowerPoint application events for the open-charts add-in. A selection change is when the add-in
' notices a chart the user resized (and lays it out again at the new size) and refreshes the
' ribbon buttons that need a selected chart.
Option Explicit

Public WithEvents App As Application

Private Sub App_WindowSelectionChange(ByVal Sel As Selection)
    OpenCharts.OnSelectionChange
End Sub

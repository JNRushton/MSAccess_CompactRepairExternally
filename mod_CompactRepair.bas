Attribute VB_Name = "mod_CompactRepair"
Option Compare Database
Option Explicit

Public Sub compactRepairMSAccessDB(ByVal filePath As String)
    On Error GoTo errorHandler
    Dim openedApp As Access.Application

    Set openedApp = New Access.Application
    openedApp.OpenCurrentDatabase filePath, False
    openedApp.Visible = True
    compactDatabase openedApp
exitProcedure:
    If Not openedApp Is Nothing Then openedApp.Quit
    Set openedApp = Nothing
    Exit Sub
errorHandler:
    MsgBox Err.Number & " : " & Err.Description, vbCritical, "Error Test"
    Resume exitProcedure
End Sub

Private Sub compactDatabase(ByRef acApp As Access.Application)
    Dim target As String: target = acApp.CurrentProject.FullName
    Dim destination As String: destination = Replace(target, ".accdb", "") & "(TEMP).accdb"

    acApp.CloseCurrentDatabase
    CompactRepair target, destination
    Kill target
    Name destination As target
    acApp.OpenCurrentDatabase target
    acApp.Visible = True
End Sub

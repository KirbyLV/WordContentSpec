
Sub CallForm()
Dim oFrm As Form1
Dim oVars As Word.Variables
Dim strTemp As String
Dim oRng As Word.Range
Dim i As Long
Dim strMultiSel As String
    Set oVars = ActiveDocument.Variables
    Set oFrm = New Form1
    With oFrm
        .Show
        oVars("Show_Name") = .TB_ShowName
        oVars("Show_Date") = .TB_ShowDate
        oVars("Venue_Name") = .TB_VenueName
        oVars("Venue_City") = .TB_VenueCity
        oVars("Frame_Rate") = .TB_FrameRate
        oVars("DocsLink") = .TB_Link
    End With

lbl_Exit:
    myUpdateFields
    Exit Sub
End Sub

Sub SetValues()
Dim ShowName As String
   ShowName = Form1.TB_ShowName.Value
Dim ShowDate As String
    ShowDate = Form1.TB_ShowDate.Value
Dim VenueName As String
    VenueName = Form1.TB_VenueName.Value
Dim VenueCity As String
    VenueCity = Form1.TB_VenueCity.Value
Dim fps As String
    fps = Form1.TB_FrameRate.Value
Dim Link As String
    Link = Form1.TB_Link.Value

ActiveDocument.Variables.Item("Show_Name").Value = ShowName

ActiveDocument.Variables.Item("Show_Date").Value = ShowDate
ActiveDocument.Variables.Item("Venue_Name").Value = VenueName
ActiveDocument.Variables.Item("Venue_City").Value = VenueCity
ActiveDocument.Variables.Item("Frame_Rate").Value = fps
ActiveDocument.Variables.Item("DocsLink").Value = Link

End Sub

Sub myUpdateFields()
Dim oStyRng As Word.Range
Dim iLink As Long
    iLink = ActiveDocument.Sections(1).Headers(1).Range.StoryType
    For Each oStyRng In ActiveDocument.StoryRanges
        Do
            oStyRng.Fields.Update
            Set oStyRng = oStyRng.NextStoryRange
        Loop Until oStyRng Is Nothing
    Next
 
End Sub

Sub Create_Reset_Variables()
    With ActiveDocument.Variables
        .Item("ShowName").Value = "SHOW NAME"
        .Item("ShowDate").Value = "SHOW DATE"
        .Item("VenueName").Value = "VENUE NAME"
        .Item("VenueCity").Value = "VENUE CITY"
        .Item("FrameRate").Value = "FPS"
        .Item("DocsLink").Value = "DOCUMENTLINK"
    End With
    myUpdateFields
lbl_Exit:
    Exit Sub
End Sub

Sub AutoNew()
    Create_Reset_Variables
    CallForm
lbl_Exit:
    Exit Sub
End Sub

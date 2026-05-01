Attribute VB_Name = "modSetting"
Public Sub SelectLogo()

    Dim ws As Worksheet
    Dim fd As FileDialog
    Dim shp As Shape
    
    Set ws = ThisWorkbook.Worksheets("Setup_UI")
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    
    With fd
        .Title = "Select Company Logo"
        .Filters.Clear
        .Filters.Add "Images", "*.jpg; *.jpeg; *.png; *.bmp"
        
        If .Show <> -1 Then Exit Sub
    End With
    
    ' Delete old logo
    For Each shp In ws.Shapes
        If shp.name = "CompanyLogo" Then shp.Delete
    Next shp
    
    ' Insert new logo
    Set shp = ws.Shapes.AddPicture( _
        Filename:=fd.SelectedItems(1), _
        LinkToFile:=msoFalse, _
        SaveWithDocument:=msoTrue, _
        Left:=ws.Range("A10").Left, _
        Top:=ws.Range("A10").Top, _
        Width:=-1, _
        Height:=-1)
    
    shp.name = "CompanyLogo"
    shp.LockAspectRatio = msoTrue
    shp.Placement = xlMoveAndSize
    
    ' Limit max size
    If shp.Width > 160 Then shp.Width = 160
    If shp.Height > 80 Then shp.Height = 80

End Sub


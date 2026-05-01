Attribute VB_Name = "modProgress"
Option Explicit

Public Sub ProgressStart(ByVal titleText As String, Optional ByVal detailText As String = "Preparing...")

    On Error Resume Next

    Load frmProgress

    With frmProgress
        .caption = "SimpleERP Processing"
        .lblTitle.caption = titleText
        .lblDetail.caption = detailText
        .lblBar.Width = 0
        .Show vbModeless
        .Repaint
    End With

    Application.StatusBar = titleText & " - " & detailText
    DoEvents

End Sub

Public Sub ProgressUpdate( _
    ByVal actionText As String, _
    ByVal currentRow As Long, _
    ByVal totalRows As Long, _
    Optional ByVal itemText As String = "")

    Dim pct As Double
    Dim detailText As String

    If totalRows <= 0 Then Exit Sub

    If currentRow Mod 10 <> 0 And currentRow <> totalRows Then Exit Sub

    pct = currentRow / totalRows
    If pct < 0 Then pct = 0
    If pct > 1 Then pct = 1

    detailText = actionText & " " & currentRow & " / " & totalRows & _
                 " (" & Format(pct, "0%") & ")"

    If itemText <> "" Then
        detailText = detailText & " | " & itemText
    End If

    On Error Resume Next

    With frmProgress
        .lblDetail.caption = detailText
        .lblBar.Width = .lblBarBack.Width * pct
        .Repaint
    End With

    Application.StatusBar = Replace(detailText, vbCrLf, " - ")
    DoEvents

End Sub

Public Sub ProgressStep(ByVal detailText As String)

    On Error Resume Next

    frmProgress.lblDetail.caption = detailText
    frmProgress.Repaint

    Application.StatusBar = detailText
    DoEvents

End Sub

Public Sub ProgressEnd()

    On Error Resume Next

    Unload frmProgress
    Application.StatusBar = False
    DoEvents

End Sub


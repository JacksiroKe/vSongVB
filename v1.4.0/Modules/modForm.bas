Attribute VB_Name = "modForm"
Option Explicit

Sub ShowOrHide(tab_mode)
    If tab_mode = "0" Then
        frmCcProject.Visible = False
        frmCcProject.cmdAdd.Visible = False
        frmCcProject.cmdMinus.Visible = False
        frmCcProject.cmdFont.Visible = False
        frmCcProject.cmdKala.Visible = False
        frmCcProject.cmdClose.Visible = False
        frmCcProject.cmdPlay.Visible = False
    Else
        frmCcProject.Visible = True
        frmCcProject.cmdAdd.Visible = True
        frmCcProject.cmdMinus.Visible = True
        frmCcProject.cmdFont.Visible = True
        frmCcProject.cmdKala.Visible = True
        frmCcProject.cmdClose.Visible = True
        frmCcProject.cmdPlay.Visible = True
    End If
End Sub

Sub SetProjectionTheme(prj_theme)
    If prj_theme = "1" Then
        frmCcProject.BackColor = &H0&
        frmCcProject.lblSongTitle.ForeColor = &HFFFFFF
        frmCcProject.lblUserName.ForeColor = &HFFFFFF
        frmCcProject.lblSongText.ForeColor = &HFFFFFF
        frmCcProject.lineTop.BorderColor = &HFFFFFF
        frmCcProject.lineDown.BorderColor = &HFFFFFF
        frmCcProject.cmdLine.BackColor = &H0&
        frmCcProject.cmdLine.ForeColor = &H0&
    ElseIf prj_theme = "2" Then
        frmCcProject.BackColor = &HFFFFFF
        frmCcProject.lblSongTitle.ForeColor = &H0&
        frmCcProject.lblUserName.ForeColor = &H0&
        frmCcProject.lblSongText.ForeColor = &H0&
        frmCcProject.lineTop.BorderColor = &H0&
        frmCcProject.lineDown.BorderColor = &H0&
        frmCcProject.cmdLine.BackColor = &HFFFFFF
        frmCcProject.cmdLine.ForeColor = &HFFFFFF
    ElseIf prj_theme = 3 Then
        frmCcProject.BackColor = &HC00000
        frmCcProject.lblSongTitle.ForeColor = &HFFFFFF
        frmCcProject.lblUserName.ForeColor = &HFFFFFF
        frmCcProject.lblSongText.ForeColor = &HFFFFFF
        frmCcProject.lineTop.BorderColor = &HFFFFFF
        frmCcProject.lineDown.BorderColor = &HFFFFFF
        frmCcProject.cmdLine.BackColor = &HC00000
        frmCcProject.cmdLine.ForeColor = &HC00000
    ElseIf prj_theme = "4" Then
        frmCcProject.BackColor = &HFFFFFF
        frmCcProject.lblSongTitle.ForeColor = &HC00000
        frmCcProject.lblUserName.ForeColor = &HC00000
        frmCcProject.lblSongText.ForeColor = &HC00000
        frmCcProject.lblSongText.BackColor = &HC00000
        frmCcProject.lineTop.BorderColor = &HC00000
        frmCcProject.lineDown.BorderColor = &HC00000
        frmCcProject.cmdLine.BackColor = &HFFFFFF
        frmCcProject.cmdLine.ForeColor = &HFFFFFF
    ElseIf prj_theme = "5" Then
        frmCcProject.BackColor = &H8000&
        frmCcProject.lblSongTitle.ForeColor = &HFFFFFF
        frmCcProject.lblUserName.ForeColor = &HFFFFFF
        frmCcProject.lblSongText.ForeColor = &HFFFFFF
        frmCcProject.lineTop.BorderColor = &HFFFFFF
        frmCcProject.lineDown.BorderColor = &HFFFFFF
        frmCcProject.cmdLine.BackColor = &H8000&
        frmCcProject.cmdLine.ForeColor = &H8000&
    ElseIf prj_theme = "6" Then
        frmCcProject.BackColor = &HFFFFFF
        frmCcProject.lblSongTitle.ForeColor = &H8000&
        frmCcProject.lblUserName.ForeColor = &H8000&
        frmCcProject.lblSongText.ForeColor = &H8000&
        frmCcProject.lineTop.BorderColor = &H8000&
        frmCcProject.lineDown.BorderColor = &H8000&
        frmCcProject.cmdLine.BackColor = &HFFFFFF
        frmCcProject.cmdLine.ForeColor = &HFFFFFF
    ElseIf prj_theme = "7" Then
        frmCcProject.BackColor = &H40C0&
        frmCcProject.lblSongTitle.ForeColor = &HFFFFFF
        frmCcProject.lblUserName.ForeColor = &HFFFFFF
        frmCcProject.lblSongText.ForeColor = &HFFFFFF
        frmCcProject.lineTop.BorderColor = &HFFFFFF
        frmCcProject.lineDown.BorderColor = &HFFFFFF
        frmCcProject.cmdLine.BackColor = &H40C0&
        frmCcProject.cmdLine.ForeColor = &H40C0&
    ElseIf prj_theme = "8" Then
        frmCcProject.BackColor = &HFFFFFF
        frmCcProject.lblSongTitle.ForeColor = &H40C0&
        frmCcProject.lblUserName.ForeColor = &H40C0&
        frmCcProject.lblSongText.ForeColor = &H40C0&
        frmCcProject.lineTop.BorderColor = &H40C0&
        frmCcProject.lineDown.BorderColor = &H40C0&
        frmCcProject.cmdLine.BackColor = &HFFFFFF
        frmCcProject.cmdLine.ForeColor = &HFFFFFF
    End If
End Sub

Sub Projection_Form_Resize()

        frmCcProject.lineTop.X2 = frmCcProject.Width - 360
        frmCcProject.lineDown.X2 = frmCcProject.Width - 360
        frmCcProject.lineDown.Y1 = frmCcProject.Height - 900
        frmCcProject.lineDown.Y2 = frmCcProject.Height - 900
        frmCcProject.lblSongText.Width = frmCcProject.Width - 1440
        frmCcProject.lblSongText.Height = frmCcProject.Height - 2200
        frmCcProject.lblSongTitle.Width = frmCcProject.Width - 500
        frmCcProject.lblUserName.Top = frmCcProject.Height - 800
    
        frmCcProject.lblUserName.Width = frmCcProject.Width - 7000
    
        frmCcProject.cmdMinus.Top = frmCcProject.Height - 800
        frmCcProject.cmdAdd.Top = frmCcProject.Height - 800
        frmCcProject.cmdFont.Top = frmCcProject.Height - 800
        frmCcProject.cmdKala.Top = frmCcProject.Height - 800
        frmCcProject.cmdPrev.Top = frmCcProject.Height - 800
        frmCcProject.cmdNext.Top = frmCcProject.Height - 800
    
        frmCcProject.cmdMinus.Left = frmCcProject.Width - 4400
        frmCcProject.cmdAdd.Left = frmCcProject.Width - 3760
        frmCcProject.cmdFont.Left = frmCcProject.Width - 3120
        frmCcProject.cmdKala.Left = frmCcProject.Width - 2480
        frmCcProject.cmdPrev.Left = frmCcProject.Width - 1840
        frmCcProject.cmdNext.Left = frmCcProject.Width - 1200
    
        frmCcProject.cmdClose.Left = frmCcProject.Width - 700
 
End Sub


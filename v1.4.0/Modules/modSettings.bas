Attribute VB_Name = "modSettings"

Sub cmdTabOne()
    frmEeOptions.cmdTab1.Top = 50
    frmEeOptions.cmdTab2.Top = 120
    frmEeOptions.cmdTab3.Top = 120
    frmEeOptions.fraSettings1.Visible = True
    frmEeOptions.fraSettings2.Visible = False
    frmEeOptions.fraSettings3.Visible = False
End Sub

Sub cmdTabTwo()
    frmEeOptions.cmdTab1.Top = 120
    frmEeOptions.cmdTab2.Top = 50
    frmEeOptions.cmdTab3.Top = 120
    frmEeOptions.fraSettings1.Visible = False
    frmEeOptions.fraSettings2.Visible = True
    frmEeOptions.fraSettings3.Visible = False
End Sub

Sub cmdTabThree()
    frmEeOptions.cmdTab1.Top = 120
    frmEeOptions.cmdTab2.Top = 120
    frmEeOptions.cmdTab3.Top = 50
    frmEeOptions.fraSettings1.Visible = False
    frmEeOptions.fraSettings2.Visible = False
    frmEeOptions.fraSettings3.Visible = True
End Sub

Sub SavedVsbTheme(pj_theme)
    If pj_theme = "1" Then
        frmEeOptions.themeOne.Top = 300
        frmEeOptions.themeOne.Height = 2295
        frmEeOptions.shapeOne.BorderColor = &HFF&
    ElseIf pj_theme = "2" Then
        frmEeOptions.themeTwo.Top = 300
        frmEeOptions.themeTwo.Height = 2295
        frmEeOptions.shapeTwo.BorderColor = &HFF&
    ElseIf pj_theme = 3 Then
        frmEeOptions.themeThree.Top = 300
        frmEeOptions.themeThree.Height = 2295
        frmEeOptions.shapeThree.BorderColor = &HFF&
    ElseIf pj_theme = "4" Then
        frmEeOptions.themeFour.Top = 300
        frmEeOptions.themeFour.Height = 2295
        frmEeOptions.shapeFour.BorderColor = &HFF&
    ElseIf pj_theme = "5" Then
        frmEeOptions.themeFive.Top = 3500
        frmEeOptions.themeFive.Height = 2295
        frmEeOptions.shapeFive.BorderColor = &HFF&
    ElseIf pj_theme = "6" Then
        frmEeOptions.themeSix.Top = 3500
        frmEeOptions.themeSix.Height = 2295
        frmEeOptions.shapeSix.BorderColor = &HFF&
    ElseIf pj_theme = "7" Then
        frmEeOptions.themeSeven.Top = 3500
        frmEeOptions.themeSeven.Height = 2295
        frmEeOptions.shapeSeven.BorderColor = &HFF&
    ElseIf pj_theme = "8" Then
        frmEeOptions.themeEight.Top = 3500
        frmEeOptions.themeEight.Height = 2295
        frmEeOptions.shapeEight.BorderColor = &HFF&
    End If
End Sub

Function themeOne_Value() As String
    frmEeOptions.themeOne.Top = 300
    frmEeOptions.themeTwo.Top = 900
    frmEeOptions.themeThree.Top = 900
    frmEeOptions.themeFour.Top = 900
    frmEeOptions.themeFive.Top = 4000
    frmEeOptions.themeSix.Top = 4000
    frmEeOptions.themeSeven.Top = 4000
    frmEeOptions.themeEight.Top = 4000
    
    frmEeOptions.themeOne.Height = 2295
    frmEeOptions.themeTwo.Height = 1700
    frmEeOptions.themeThree.Height = 1700
    frmEeOptions.themeFour.Height = 1700
    frmEeOptions.themeFive.Height = 1700
    frmEeOptions.themeSix.Height = 1700
    frmEeOptions.themeSeven.Height = 1700
    frmEeOptions.themeEight.Height = 1700
    
    frmEeOptions.shapeOne.BorderColor = &HFF&
    frmEeOptions.shapeTwo.BorderColor = &H0&
    frmEeOptions.shapeThree.BorderColor = &H0&
    frmEeOptions.shapeFour.BorderColor = &H0&
    frmEeOptions.shapeFive.BorderColor = &H0&
    frmEeOptions.shapeSix.BorderColor = &H0&
    frmEeOptions.shapeSeven.BorderColor = &H0&
    frmEeOptions.shapeEight.BorderColor = &H0&
        
    themeOne_Value = "1"
End Function

Function themeTwo_Value() As String
    frmEeOptions.themeOne.Top = 900
    frmEeOptions.themeTwo.Top = 300
    frmEeOptions.themeThree.Top = 900
    frmEeOptions.themeFour.Top = 900
    frmEeOptions.themeFive.Top = 4000
    frmEeOptions.themeSix.Top = 4000
    frmEeOptions.themeSeven.Top = 4000
    frmEeOptions.themeEight.Top = 4000
    
    frmEeOptions.themeOne.Height = 1700
    frmEeOptions.themeTwo.Height = 2295
    frmEeOptions.themeThree.Height = 1700
    frmEeOptions.themeFour.Height = 1700
    frmEeOptions.themeFive.Height = 1700
    frmEeOptions.themeSix.Height = 1700
    frmEeOptions.themeSeven.Height = 1700
    frmEeOptions.themeEight.Height = 1700
    
    frmEeOptions.shapeOne.BorderColor = &H0&
    frmEeOptions.shapeTwo.BorderColor = &HFF&
    frmEeOptions.shapeThree.BorderColor = &H0&
    frmEeOptions.shapeFour.BorderColor = &H0&
    frmEeOptions.shapeFive.BorderColor = &H0&
    frmEeOptions.shapeSix.BorderColor = &H0&
    frmEeOptions.shapeSeven.BorderColor = &H0&
    frmEeOptions.shapeEight.BorderColor = &H0&
    
    themeTwo_Value = "2"
End Function

Function themeThree_Value() As String
    frmEeOptions.themeOne.Top = 900
    frmEeOptions.themeTwo.Top = 900
    frmEeOptions.themeThree.Top = 300
    frmEeOptions.themeFour.Top = 900
    frmEeOptions.themeFive.Top = 4000
    frmEeOptions.themeSix.Top = 4000
    frmEeOptions.themeSeven.Top = 4000
    frmEeOptions.themeEight.Top = 4000
    
    frmEeOptions.themeOne.Height = 1700
    frmEeOptions.themeTwo.Height = 1700
    frmEeOptions.themeThree.Height = 2295
    frmEeOptions.themeFour.Height = 1700
    frmEeOptions.themeFive.Height = 1700
    frmEeOptions.themeSix.Height = 1700
    frmEeOptions.themeSeven.Height = 1700
    frmEeOptions.themeEight.Height = 1700
    
    frmEeOptions.shapeOne.BorderColor = &H0&
    frmEeOptions.shapeTwo.BorderColor = &H0&
    frmEeOptions.shapeThree.BorderColor = &HFF&
    frmEeOptions.shapeFour.BorderColor = &H0&
    frmEeOptions.shapeFive.BorderColor = &H0&
    frmEeOptions.shapeSix.BorderColor = &H0&
    frmEeOptions.shapeSeven.BorderColor = &H0&
    frmEeOptions.shapeEight.BorderColor = &H0&
    
    themeThree_Value = "3"
End Function

Function themeFour_Value() As String
    frmEeOptions.themeOne.Top = 900
    frmEeOptions.themeTwo.Top = 900
    frmEeOptions.themeThree.Top = 900
    frmEeOptions.themeFour.Top = 300
    frmEeOptions.themeFive.Top = 4000
    frmEeOptions.themeSix.Top = 4000
    frmEeOptions.themeSeven.Top = 4000
    frmEeOptions.themeEight.Top = 4000
    
    frmEeOptions.themeOne.Height = 1700
    frmEeOptions.themeTwo.Height = 1700
    frmEeOptions.themeThree.Height = 1700
    frmEeOptions.themeFour.Height = 2295
    frmEeOptions.themeFive.Height = 1700
    frmEeOptions.themeSix.Height = 1700
    frmEeOptions.themeSeven.Height = 1700
    frmEeOptions.themeEight.Height = 1700
    
    frmEeOptions.shapeOne.BorderColor = &H0&
    frmEeOptions.shapeTwo.BorderColor = &H0&
    frmEeOptions.shapeThree.BorderColor = &H0&
    frmEeOptions.shapeFour.BorderColor = &HFF&
    frmEeOptions.shapeFive.BorderColor = &H0&
    frmEeOptions.shapeSix.BorderColor = &H0&
    frmEeOptions.shapeSeven.BorderColor = &H0&
    frmEeOptions.shapeEight.BorderColor = &H0&
    
    themeFour_Value = "4"
End Function

Function themeFive_Value() As String
    frmEeOptions.themeOne.Top = 900
    frmEeOptions.themeTwo.Top = 900
    frmEeOptions.themeThree.Top = 900
    frmEeOptions.themeFour.Top = 900
    frmEeOptions.themeFive.Top = 3500
    frmEeOptions.themeSix.Top = 4000
    frmEeOptions.themeSeven.Top = 4000
    frmEeOptions.themeEight.Top = 4000
    
    frmEeOptions.themeOne.Height = 1700
    frmEeOptions.themeTwo.Height = 1700
    frmEeOptions.themeThree.Height = 1700
    frmEeOptions.themeFour.Height = 1700
    frmEeOptions.themeFive.Height = 2295
    frmEeOptions.themeSix.Height = 1700
    frmEeOptions.themeSeven.Height = 1700
    frmEeOptions.themeEight.Height = 1700
    
    frmEeOptions.shapeOne.BorderColor = &H0&
    frmEeOptions.shapeTwo.BorderColor = &H0&
    frmEeOptions.shapeThree.BorderColor = &H0&
    frmEeOptions.shapeFour.BorderColor = &H0&
    frmEeOptions.shapeFive.BorderColor = &HFF&
    frmEeOptions.shapeSix.BorderColor = &H0&
    frmEeOptions.shapeSeven.BorderColor = &H0&
    frmEeOptions.shapeEight.BorderColor = &H0&
    
    themeFive_Value = "5"
End Function

Function themeSix_Value() As String
    frmEeOptions.themeOne.Top = 900
    frmEeOptions.themeTwo.Top = 900
    frmEeOptions.themeThree.Top = 900
    frmEeOptions.themeFour.Top = 900
    frmEeOptions.themeFive.Top = 4000
    frmEeOptions.themeSix.Top = 3500
    frmEeOptions.themeSeven.Top = 4000
    frmEeOptions.themeEight.Top = 4000
    
    frmEeOptions.themeOne.Height = 1700
    frmEeOptions.themeTwo.Height = 1700
    frmEeOptions.themeThree.Height = 1700
    frmEeOptions.themeFour.Height = 1700
    frmEeOptions.themeFive.Height = 1700
    frmEeOptions.themeSix.Height = 2295
    frmEeOptions.themeSeven.Height = 1700
    frmEeOptions.themeEight.Height = 1700
    
    frmEeOptions.shapeOne.BorderColor = &H0&
    frmEeOptions.shapeTwo.BorderColor = &H0&
    frmEeOptions.shapeThree.BorderColor = &H0&
    frmEeOptions.shapeFour.BorderColor = &H0&
    frmEeOptions.shapeFive.BorderColor = &H0&
    frmEeOptions.shapeSix.BorderColor = &HFF&
    frmEeOptions.shapeSeven.BorderColor = &H0&
    frmEeOptions.shapeEight.BorderColor = &H0&
    
    themeSix_Value = "6"
End Function

Function themeSeven_Value() As String
    frmEeOptions.themeOne.Top = 900
    frmEeOptions.themeTwo.Top = 900
    frmEeOptions.themeThree.Top = 900
    frmEeOptions.themeFour.Top = 900
    frmEeOptions.themeFive.Top = 4000
    frmEeOptions.themeSix.Top = 4000
    frmEeOptions.themeSeven.Top = 3500
    frmEeOptions.themeEight.Top = 4000
    
    frmEeOptions.themeOne.Height = 1700
    frmEeOptions.themeTwo.Height = 1700
    frmEeOptions.themeThree.Height = 1700
    frmEeOptions.themeFour.Height = 1700
    frmEeOptions.themeFive.Height = 1700
    frmEeOptions.themeSix.Height = 1700
    frmEeOptions.themeSeven.Height = 2295
    frmEeOptions.themeEight.Height = 1700
    
    frmEeOptions.shapeOne.BorderColor = &H0&
    frmEeOptions.shapeTwo.BorderColor = &H0&
    frmEeOptions.shapeThree.BorderColor = &H0&
    frmEeOptions.shapeFour.BorderColor = &H0&
    frmEeOptions.shapeFive.BorderColor = &H0&
    frmEeOptions.shapeSix.BorderColor = &H0&
    frmEeOptions.shapeSeven.BorderColor = &HFF&
    frmEeOptions.shapeEight.BorderColor = &H0&
    
    themeSeven_Value = "7"
End Function

Function themeEight_Value() As String
    frmEeOptions.themeOne.Top = 900
    frmEeOptions.themeTwo.Top = 900
    frmEeOptions.themeThree.Top = 900
    frmEeOptions.themeFour.Top = 900
    frmEeOptions.themeFive.Top = 4000
    frmEeOptions.themeSix.Top = 4000
    frmEeOptions.themeSeven.Top = 4000
    frmEeOptions.themeEight.Top = 3500
    
    frmEeOptions.themeOne.Height = 1700
    frmEeOptions.themeTwo.Height = 1700
    frmEeOptions.themeThree.Height = 1700
    frmEeOptions.themeFour.Height = 1700
    frmEeOptions.themeFive.Height = 1700
    frmEeOptions.themeSix.Height = 1700
    frmEeOptions.themeSeven.Height = 1700
    frmEeOptions.themeEight.Height = 2295
    
    frmEeOptions.shapeOne.BorderColor = &H0&
    frmEeOptions.shapeTwo.BorderColor = &H0&
    frmEeOptions.shapeThree.BorderColor = &H0&
    frmEeOptions.shapeFour.BorderColor = &H0&
    frmEeOptions.shapeFive.BorderColor = &H0&
    frmEeOptions.shapeSix.BorderColor = &H0&
    frmEeOptions.shapeSeven.BorderColor = &H0&
    frmEeOptions.shapeEight.BorderColor = &HFF&
    
    themeEight_Value = "8"
End Function


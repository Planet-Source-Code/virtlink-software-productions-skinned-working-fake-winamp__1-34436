Attribute VB_Name = "mdlSkin"
Public intCButtons(1 To 6, 1 To 2, 1 To 4) As Integer
Public intTitlebar(1 To 18, 1 To 2, 1 To 4) As Integer
Public intShufRep(1 To 4, 1 To 4, 1 To 4) As Integer
Public intPosBar(1 To 2, 1 To 2, 1 To 4) As Integer
Public intVolume(1 To 30, 1 To 1, 1 To 4) As Integer
Public intBalance(1 To 30, 1 To 1, 1 To 4) As Integer
Public intPlayPaus(1 To 6, 1 To 1, 1 To 4) As Integer
Public intNumbers(0 To 10, 1 To 1, 1 To 4) As Integer
Public intMonoster(1 To 2, 1 To 2, 1 To 4) As Integer

Public Enum enmObjType
    obtCButtons = 1
    obtTitlebar = 2
    obtShufRep = 3
    obtPosBar = 4
    obtVolume = 5
    obtBalance = 6
    obtPlayPaus = 7
    obtNumers = 8
    obtMonoSter = 9
End Enum

Public Enum enmObjStat
    obsDefault
    obsPressed
    obsOn
    obsOnPressed
End Enum

Public Sub LoadSkin(strPath As String)
    Dim objIni As New APIINISystem, q As Integer
    
    '*** Perform general operations ***'
    frmMain.MoveControls
    InitClipVars
    
    '*** Load main skin ***'
    frmMain.Picture = LoadPicture(strPath & "\main.bmp")
    frmMain.Mask.Picture = LoadPicture(strPath & "\main.bmp")
    frmMain.imgMove.Move 0, 0, frmMain.Mask.ScaleWidth, frmMain.Mask.ScaleHeight
    '*** Set tranparant sections ***'
    objIni.INIFile = strPath & "\region.txt"
    frmMain.ChangeMask RGB(CInt(objIni.INI_Read("Normal", "TransColorR", "255")), CInt(objIni.INI_Read("Normal", "TransColorG", "0")), CInt(objIni.INI_Read("Normal", "TransColorB", "255")))
    
    '*** Set play-buttons ***'
    frmMain.clipCButtons.Picture = LoadPicture(strPath & "\cbuttons.bmp")
    For q = 1 To 6
        GetClip obtCButtons, q, obsDefault
    Next q
    
    '*** Set titlebar ***'
    frmMain.clipTitlebar.Picture = LoadPicture(strPath & "\titlebar.bmp")
    GetClip obtTitlebar, 3, obsDefault
    GetClip obtTitlebar, 4, obsDefault
    GetClip obtTitlebar, 5, obsDefault
    GetClip obtTitlebar, 6, obsDefault
    GetClip obtTitlebar, 7, obsDefault
    
    '*** Set buttons ***'
    frmMain.clipShufRep.Picture = LoadPicture(strPath & "\shufrep.bmp")
    GetClip obtShufRep, 1, obsDefault
    GetClip obtShufRep, 2, obsDefault
    GetClip obtShufRep, 3, obsDefault
    GetClip obtShufRep, 4, obsDefault
    
    '*** Set pos slider ***'
    frmMain.clipPosBar.Picture = LoadPicture(strPath & "\posbar.bmp")
    GetClip obtPosBar, 1, obsDefault
    GetClip obtPosBar, 2, obsDefault
    
    '*** Set volume control ***'
    frmMain.clipVolume.Picture = LoadPicture(strPath & "\volume.bmp")
    GetClip obtVolume, 14, obsDefault
    GetClip obtVolume, 29, obsDefault
    
    '*** Set balance control ***'
    frmMain.clipBalance.Picture = LoadPicture(strPath & "\balance.bmp")
    GetClip obtBalance, 1, obsDefault
    GetClip obtBalance, 29, obsDefault
    
    '*** Set playpaus control ***'
    frmMain.clipPlayPaus.Picture = LoadPicture(strPath & "\playpaus.bmp")
    GetClip obtPlayPaus, 1, obsDefault
    GetClip obtPlayPaus, 6, obsDefault
    
    '*** Set numbers control ***'
    frmMain.clipNumbers.Picture = LoadPicture(strPath & "\numbers.bmp")
    GetClip obtNumers, 0, 0
    GetClip obtNumers, 0, 1
    GetClip obtNumers, 0, 2
    GetClip obtNumers, 0, 3
    
    '*** Set monoster control ***'
    frmMain.clipMonoster.Picture = LoadPicture(strPath & "\monoster.bmp")
    GetClip obtMonoSter, 1, obsDefault
    GetClip obtMonoSter, 2, obsDefault
End Sub

Public Sub InitClipVars()
    Dim q As Integer, r As Integer
    
    '*** Initialize clip coordinates ***'
    
    '*** Adjust: "cbuttons.bmp" ***'
    For q = 1 To 5
        intCButtons(q, 1, 2) = 0
        intCButtons(q, 1, 3) = 23
        intCButtons(q, 1, 4) = 18
        
        intCButtons(q, 2, 2) = 18
        intCButtons(q, 2, 3) = 23
        intCButtons(q, 2, 4) = 18
    Next q
    intCButtons(5, 1, 3) = 22
    intCButtons(5, 2, 3) = 22
    For q = 1 To 2
        intCButtons(1, q, 1) = 0
        intCButtons(2, q, 1) = 23
        intCButtons(3, q, 1) = 46
        intCButtons(4, q, 1) = 69
        intCButtons(5, q, 1) = 92
        intCButtons(6, q, 1) = 114
        
        intCButtons(6, q, 3) = 22
        intCButtons(6, q, 4) = 16
    Next q
    intCButtons(6, 1, 2) = 0
    intCButtons(6, 2, 2) = 16
    
    '*** Adjust: "titlebar.bmp" ***'
    'The real titlebar:
    For q = 1 To 3
        For r = 1 To 2
            intTitlebar(q, r, 1) = 27
            intTitlebar(q, r, 3) = 274
        Next r
    Next q
    intTitlebar(1, 1, 2) = 0
    intTitlebar(1, 2, 2) = 15
    intTitlebar(2, 1, 2) = 29
    intTitlebar(2, 2, 2) = 42
    intTitlebar(3, 1, 2) = 57
    intTitlebar(3, 2, 2) = 72
    
    intTitlebar(1, 1, 4) = 12
    intTitlebar(1, 2, 4) = 12
    intTitlebar(2, 1, 4) = 13
    intTitlebar(2, 2, 4) = 13
    intTitlebar(3, 1, 4) = 12
    intTitlebar(3, 2, 4) = 12
    
    'The four top menu buttons:
    For q = 4 To 7
        For r = 1 To 2
            intTitlebar(q, r, 3) = 9
            intTitlebar(q, r, 4) = 9
        Next r
    Next q
    For q = 4 To 6
        intTitlebar(q, 1, 2) = 0
        intTitlebar(q, 2, 2) = 9
    Next q
    intTitlebar(4, 1, 1) = 0
    intTitlebar(4, 2, 1) = 0
    
    intTitlebar(5, 1, 1) = 9
    intTitlebar(5, 2, 1) = 9
    
    intTitlebar(6, 1, 1) = 18
    intTitlebar(6, 2, 1) = 18
    
    intTitlebar(7, 1, 1) = 0
    intTitlebar(7, 1, 2) = 18
    intTitlebar(7, 2, 1) = 9
    intTitlebar(7, 2, 2) = 18
    
    '*** Adjust: "shufrep.bmp" ***'
    'The shuffle/repeat/EQ/PL-buttons:
    For q = 1 To 4
        intShufRep(1, q, 1) = 0
        intShufRep(1, q, 3) = 28
        intShufRep(1, q, 4) = 15
        
        intShufRep(2, q, 1) = 28
        intShufRep(2, q, 3) = 47
        intShufRep(2, q, 4) = 15
    Next q
    For q = 1 To 2
        intShufRep(q, 1, 2) = 0
        intShufRep(q, 2, 2) = 15
        intShufRep(q, 3, 2) = 30
        intShufRep(q, 4, 2) = 45
    Next q
    
    For q = 3 To 4
        For r = 1 To 4
            intShufRep(q, r, 3) = 23
            intShufRep(q, r, 4) = 12
        Next r
    Next q
    For q = 3 To 4
        For r = 1 To 2
            intShufRep(q, r, 2) = 61
            intShufRep(q, r + 2, 2) = 73
        Next r
    Next q
    intShufRep(3, 1, 1) = 0
    intShufRep(3, 3, 1) = 0
    intShufRep(4, 1, 1) = 23
    intShufRep(4, 3, 1) = 23
    intShufRep(3, 2, 1) = 46
    intShufRep(3, 4, 1) = 46
    intShufRep(4, 2, 1) = 69
    intShufRep(4, 4, 1) = 69
    
    '*** Adjust: "posbar.bmp" ***'
    intPosBar(1, 1, 1) = 0
    intPosBar(1, 1, 2) = 0
    intPosBar(1, 1, 3) = 248
    intPosBar(1, 1, 4) = 9
    
    intPosBar(2, 1, 1) = 248
    intPosBar(2, 2, 1) = 278
    For q = 1 To 2
        intPosBar(2, q, 2) = 0
        intPosBar(2, q, 3) = 29
        intPosBar(2, q, 4) = 10
    Next q
    
    '*** Adjust: "volume.bmp" ***'
    For q = 1 To 28
        intVolume(q, 1, 1) = 0
        intVolume(q, 1, 2) = (q - 1) * 15
        intVolume(q, 1, 3) = 68
        intVolume(q, 1, 4) = 13
    Next q
    For q = 29 To 30
        intVolume(q, 1, 2) = 422
        intVolume(q, 1, 3) = 14
        intVolume(q, 1, 4) = 11
    Next q
    intVolume(29, 1, 1) = 15
    intVolume(30, 1, 1) = 0
    
    '*** Adjust: "balance.bmp" ***'
    For q = 1 To 28
        intBalance(q, 1, 1) = 9
        intBalance(q, 1, 2) = (q - 1) * 15
        intBalance(q, 1, 3) = 38
        intBalance(q, 1, 4) = 13
    Next q
    For q = 29 To 30
        intBalance(q, 1, 2) = 422
        intBalance(q, 1, 3) = 14
        intBalance(q, 1, 4) = 11
    Next q
    intBalance(29, 1, 1) = 15
    intBalance(30, 1, 1) = 0
    
    '*** Adjust: "playpaus.bmp" ***'
    For q = 1 To 6
        intPlayPaus(q, 1, 2) = 0
        intPlayPaus(q, 1, 3) = 9
        intPlayPaus(q, 1, 4) = 9
    Next q
    intPlayPaus(5, 1, 3) = 3
    intPlayPaus(6, 1, 3) = 3
    
    intPlayPaus(1, 1, 1) = 0
    intPlayPaus(2, 1, 1) = 9
    intPlayPaus(3, 1, 1) = 18
    intPlayPaus(4, 1, 1) = 27
    intPlayPaus(5, 1, 1) = 36
    intPlayPaus(6, 1, 1) = 39
    
    '*** Adjust: "numbers.bmp" ***'
    For q = 0 To 10
        intNumbers(q, 1, 1) = q * 9
        intNumbers(q, 1, 2) = 0
        intNumbers(q, 1, 3) = 9
        intNumbers(q, 1, 4) = 13
    Next q
    
    '*** Adjust: "monoster.bmp" ***'
    For q = 1 To 2
        For r = 1 To 2
            intMonoster(q, r, 3) = 29
            intMonoster(q, r, 4) = 12
        Next
        intMonoster(q, 2, 2) = 0
        intMonoster(q, 1, 2) = 12
        
        intMonoster(1, q, 1) = 0
        intMonoster(2, q, 1) = 29
    Next
End Sub
Public Function GetClip(intObjType As enmObjType, intNumber As Integer, intStatus As enmObjStat)
    On Error GoTo 0
    On Error Resume Next
    
    Select Case intObjType
        Case obtCButtons
            With frmMain.clipCButtons
                .ClipX = intCButtons(intNumber, intStatus + 1, 1)
                .ClipY = intCButtons(intNumber, intStatus + 1, 2)
                .ClipWidth = intCButtons(intNumber, intStatus + 1, 3)
                .ClipHeight = intCButtons(intNumber, intStatus + 1, 4)
                frmMain.imgPlayBut(intNumber - 1).Picture = .Clip
            End With
        Case obtTitlebar
            With frmMain.clipTitlebar
                .ClipX = intTitlebar(intNumber, intStatus + 1, 1)
                .ClipY = intTitlebar(intNumber, intStatus + 1, 2)
                .ClipWidth = intTitlebar(intNumber, intStatus + 1, 3)
                .ClipHeight = intTitlebar(intNumber, intStatus + 1, 4)
                Select Case intNumber
                    Case 1 To 3
                        frmMain.imgTitlebar.Picture = .Clip
                    Case 4 To 7
                        frmMain.imgMenu(intNumber - 4).Picture = .Clip
                End Select
            End With
        Case obtShufRep
            With frmMain.clipShufRep
                .ClipX = intShufRep(intNumber, intStatus + 1, 1)
                .ClipY = intShufRep(intNumber, intStatus + 1, 2)
                .ClipWidth = intShufRep(intNumber, intStatus + 1, 3)
                .ClipHeight = intShufRep(intNumber, intStatus + 1, 4)
'                Select Case intNumber
'                    Case 1, 2
                frmMain.imgSufRep(intNumber - 1).Picture = .Clip
'                    Case 3, 4
'                        'frmMain.imgMenu(intNumber - 4).Picture = .Clip
'                End Select
            End With
        Case obtPosBar
            With frmMain.clipPosBar
                .ClipX = intPosBar(intNumber, intStatus + 1, 1)
                .ClipY = intPosBar(intNumber, intStatus + 1, 2)
                .ClipWidth = intPosBar(intNumber, intStatus + 1, 3)
                .ClipHeight = intPosBar(intNumber, intStatus + 1, 4)
                Select Case intNumber
                    Case 1
                        frmMain.imgPosBar.Picture = .Clip
                    Case 2
                        frmMain.imgPosSlide(0).Picture = .Clip
                End Select
            End With
        Case obtVolume
            With frmMain.clipVolume
                .ClipX = intVolume(intNumber, intStatus + 1, 1)
                .ClipY = intVolume(intNumber, intStatus + 1, 2)
                .ClipWidth = intVolume(intNumber, intStatus + 1, 3)
                .ClipHeight = intVolume(intNumber, intStatus + 1, 4)
                Select Case intNumber
                    Case 1 To 28
                        frmMain.imgVolume.Picture = .Clip
                    Case 29, 30
                        frmMain.imgVolumeBut(0).Picture = .Clip
                End Select
            End With
        Case obtBalance
            With frmMain.clipBalance
                .ClipX = intBalance(intNumber, intStatus + 1, 1)
                .ClipY = intBalance(intNumber, intStatus + 1, 2)
                .ClipWidth = intBalance(intNumber, intStatus + 1, 3)
                .ClipHeight = intBalance(intNumber, intStatus + 1, 4)
                Select Case intNumber
                    Case 1 To 28
                        frmMain.imgBalance.Picture = .Clip
                    Case 29, 30
                        frmMain.imgBalanceBut(0).Picture = .Clip
                End Select
            End With
        Case obtPlayPaus
            With frmMain.clipPlayPaus
                .ClipX = intPlayPaus(intNumber, intStatus + 1, 1)
                .ClipY = intPlayPaus(intNumber, intStatus + 1, 2)
                .ClipWidth = intPlayPaus(intNumber, intStatus + 1, 3)
                .ClipHeight = intPlayPaus(intNumber, intStatus + 1, 4)
                Select Case intNumber
                    Case 1 To 4
                        frmMain.imgPlayPaus(1).Picture = .Clip
                    Case 5, 6
                        frmMain.imgPlayPaus(0).Picture = .Clip
                End Select
            End With
        Case obtNumers
            With frmMain.clipNumbers
                .ClipX = intNumbers(intNumber, 1, 1)
                .ClipY = intNumbers(intNumber, 1, 2)
                .ClipWidth = intNumbers(intNumber, 1, 3)
                .ClipHeight = intNumbers(intNumber, 1, 4)
                frmMain.imgNumber(intStatus).Picture = .Clip
            End With
        Case obtMonoSter
            With frmMain.clipMonoster
                .ClipX = intMonoster(intNumber, intStatus + 1, 1)
                .ClipY = intMonoster(intNumber, intStatus + 1, 2)
                .ClipWidth = intMonoster(intNumber, intStatus + 1, 3)
                .ClipHeight = intMonoster(intNumber, intStatus + 1, 4)
                Select Case intNumber
                    Case 1
                        frmMain.imgMonoSter(0).Picture = .Clip
                    Case 2
                        frmMain.imgMonoSter(1).Picture = .Clip
                End Select
            End With
    End Select
End Function

Public Sub Recycled()
'    With frmMain.clipCButtons
'        .ClipX = 0
'        .ClipY = 0
'        .ClipWidth = 22
'        .ClipHeight = 18
'        frmMain.imgPlayBut(0).Picture = .Clip
'
'        .ClipX = 23
'        .ClipY = 0
'        .ClipWidth = 22
'        .ClipHeight = 18
'        frmMain.imgPlayBut(1).Picture = .Clip
'
'        .ClipX = 46
'        .ClipY = 0
'        .ClipWidth = 22
'        .ClipHeight = 18
'        frmMain.imgPlayBut(2).Picture = .Clip
'
'        .ClipX = 69
'        .ClipY = 0
'        .ClipWidth = 22
'        .ClipHeight = 18
'        frmMain.imgPlayBut(3).Picture = .Clip
'
'        .ClipX = 92
'        .ClipY = 0
'        .ClipWidth = 22
'        .ClipHeight = 18
'        frmMain.imgPlayBut(4).Picture = .Clip
'
'        .ClipX = 114
'        .ClipY = 0
'        .ClipWidth = 21
'        .ClipHeight = 15
'        frmMain.imgPlayBut(5).Picture = .Clip
'    End With
End Sub


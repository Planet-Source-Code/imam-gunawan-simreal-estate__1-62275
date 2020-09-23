Attribute VB_Name = "DxMod"
Option Explicit

'The Main Object of DirectX Components
Public DX7 As DirectX7
Public DxD As DirectDraw7

Public Display As DirectDrawSurface7
Public Layar As DirectDrawSurface7          'Layar tampilan untuk penggulungan
Public Primary As DirectDrawSurface7        'The Screen We See
Public BackBuffer As DirectDrawSurface7     'The Backbuffer we need in Flipping Mode
Public XSpot As DirectDrawSurface7          'Layar Spot

'Data DirectDrawSurface7 Sprite
Public GrassTex As DirectDrawSurface7       'grass texture
Public RoadTex As DirectDrawSurface7        'road texture
Public ToolbarTex As DirectDrawSurface7     'toolbar texture
Public MouseTex As DirectDrawSurface7       'mouse texture
Public RoadGUI As DirectDrawSurface7        'road toolbar
Public BuildingTex As DirectDrawSurface7    'builing toolbar
Public MiniMapTex As DirectDrawSurface7     'minimap texture
Public PosTex As DirectDrawSurface7         'Pos texture
Public GerejaTex As DirectDrawSurface7      'Gereja texture
Public PohonTex As DirectDrawSurface7       'Pohon Texture
Public ListrikTex As DirectDrawSurface7     'Listrik Texture
Public MsgBoxTex As DirectDrawSurface7      'Msgbox texture
Public HelpBox As DirectDrawSurface7        'HelpBox texture

'DirectDrawSurface7 untuk simulasi
Public ParkTex As DirectDrawSurface7        'park texture
Public GedungA As DirectDrawSurface7        'gedung A texture
Public GedungB As DirectDrawSurface7        'gedung B texture
Public GedungC As DirectDrawSurface7        'gedung C texture

'Untuk pengontrolan Gamma
Public mobjGammaControler As DirectDrawGammaControl    'The object that gets/sets gamma ramps
Public mudtGammaRamp As DDGAMMARAMP                    'The gamma ramp we'll use to alter the screen state
Public mudtOriginalRamp As DDGAMMARAMP                 'The gamma ramp we'll use to store the original screen state
Public mintRedVal As Integer                        'Store the currend red value w.r.t. original
Public mintGreenVal As Integer                      'Store the currend green value w.r.t. original
Public mintBlueVal As Integer                       'Store the currend blue value w.r.t. original
'Public mblnGamma As Boolean                         'Do we have gamma support?
'Public mblnFadeIn As Boolean                        'Should we fade back in?

Public StillRunning As Boolean

'kelas musik dan efek suara
Public SFXMusik As New SoundCls

'Program flow variables
Dim mlngFrameTime As Long                   'How long since last frame?
Dim mlngTimer As Long                       'How long since last FPS count update?
Dim mintFPSCounter As Integer               'Our FPS counter
Public mintFPS As Integer                   'Our FPS storage variable

Public ddschar As DDSURFACEDESC2
Public ddsmap As DDSURFACEDESC2

Public CursorX As Long
Public CursorY As Long
Public Mouse_Button0 As Boolean
Public Mouse_Button1 As Boolean
Public Mouse_Button2 As Boolean
Public Mouse_Button3 As Boolean

Public Const MAX_FPS = 70
Public Const MOUSESPEED = 2

'Variabel Perhitungan Permainan
Private Rmh1 As Integer, Rmh2 As Integer, Rmh3 As Integer
Private RoadCnt As Integer, ElectricCnt As Byte, TreesCnt As Integer
Private ChurchCnt As Integer, PosCnt As Integer, ParkCnt As Integer
Private Polusi As Single, Infrastruktur As Single, Aman As Single, Investasi As Single
Private Lingkungan As Single, TotalRmh As Integer
Private Ket1 As String, Ket2 As String, Ket3 As String, Ket4 As String, Ket5 As String


Private DimPilih As Byte

Public Function PosisiPosDekat(ArXDunia As Integer, ArYDunia As Integer) As Boolean
    PosisiPosDekat = False
    Dim X As Integer
    Dim Y As Integer
    
    For X = 1 To 4
        If (ArXDunia + X) < 60 Then
            If ArGenap(ArXDunia + X, ArYDunia).Tipe = POS Then
                PosisiPosDekat = True
                Exit Function
            End If
        Else
            PosisiPosDekat = True
        End If
        If (ArXDunia - X) > 2 Then
            If ArGenap(ArXDunia - X, ArYDunia).Tipe = POS Then
                PosisiPosDekat = True
                Exit Function
            End If
        Else
            PosisiPosDekat = True
        End If
    Next X
    
    For Y = 1 To 4
        If (ArYDunia + Y) < 30 Then
            If ArGenap(ArXDunia, ArYDunia + Y).Tipe = POS Then
                PosisiPosDekat = True
                Exit Function
            End If
        Else
            PosisiPosDekat = True
        End If
        If (ArYDunia - Y) > 2 Then
            If ArGenap(ArXDunia, ArYDunia - Y).Tipe = POS Then
                PosisiPosDekat = True
                Exit Function
            End If
        Else
            PosisiPosDekat = True
        End If
    Next Y
    
    For X = 1 To 4
    For Y = 1 To 4
        If (ArYDunia + Y) < 30 And (ArXDunia + X) < 60 Then
            If ArGenap(ArXDunia + X, ArYDunia + Y).Tipe = POS Then
                PosisiPosDekat = True
                Exit Function
            End If
        Else
            PosisiPosDekat = True
        End If
        If (ArYDunia - Y) > 2 And (ArXDunia - X) > 2 Then
            If ArGenap(ArXDunia - X, ArYDunia - Y).Tipe = POS Then
                PosisiPosDekat = True
                Exit Function
            End If
        Else
            PosisiPosDekat = True
        End If
        If (ArYDunia - Y) > 2 And (ArXDunia + X) < 60 Then
            If ArGenap(ArXDunia + X, ArYDunia - Y).Tipe = POS Then
                PosisiPosDekat = True
                Exit Function
            End If
        Else
            PosisiPosDekat = True
        End If
        
        If (ArYDunia + Y) < 30 And (ArXDunia - X) > 2 Then
            If ArGenap(ArXDunia - X, ArYDunia + Y).Tipe = POS Then
                PosisiPosDekat = True
                Exit Function
            End If
        Else
            PosisiPosDekat = True
        End If
    Next Y
    Next X
End Function
Public Sub SetGamma(intRed As Integer, intGreen As Integer, intBlue As Integer)

Dim i As Integer

    'Alter the gamma ramp to the percent given by comparing to original state
    'A value of zero ("0") for intRed, intGreen, or intBlue will result in the
    'gamma level being set back to the original levels. Anything ABOVE zero will
    'fade towards FULL colour, anything below zero will fade towards NO colour
    For i = 0 To 255
        If intRed < 0 Then mudtGammaRamp.red(i) = ConvToSignedValue(ConvToUnSignedValue(mudtOriginalRamp.red(i)) * (100 - Abs(intRed)) / 100)
        If intRed = 0 Then mudtGammaRamp.red(i) = mudtOriginalRamp.red(i)
        If intRed > 0 Then mudtGammaRamp.red(i) = ConvToSignedValue(65535 - ((65535 - ConvToUnSignedValue(mudtOriginalRamp.red(i))) * (100 - intRed) / 100))
        If intGreen < 0 Then mudtGammaRamp.green(i) = ConvToSignedValue(ConvToUnSignedValue(mudtOriginalRamp.green(i)) * (100 - Abs(intGreen)) / 100)
        If intGreen = 0 Then mudtGammaRamp.green(i) = mudtOriginalRamp.green(i)
        If intGreen > 0 Then mudtGammaRamp.green(i) = ConvToSignedValue(65535 - ((65535 - ConvToUnSignedValue(mudtOriginalRamp.green(i))) * (100 - intGreen) / 100))
        If intBlue < 0 Then mudtGammaRamp.blue(i) = ConvToSignedValue(ConvToUnSignedValue(mudtOriginalRamp.blue(i)) * (100 - Abs(intBlue)) / 100)
        If intBlue = 0 Then mudtGammaRamp.blue(i) = mudtOriginalRamp.blue(i)
        If intBlue > 0 Then mudtGammaRamp.blue(i) = ConvToSignedValue(65535 - ((65535 - ConvToUnSignedValue(mudtOriginalRamp.blue(i))) * (100 - intBlue) / 100))
    Next
    
    mobjGammaControler.SetGammaRamp DDSGR_DEFAULT, mudtGammaRamp

End Sub

Public Function ConvToSignedValue(lngValue As Long) As Integer

    'Cheezy method for converting to signed integer
    If lngValue <= 32767 Then
        ConvToSignedValue = CInt(lngValue)
        Exit Function
    End If
    
    ConvToSignedValue = CInt(lngValue - 65535)

End Function

Public Function ConvToUnSignedValue(intValue As Integer) As Long

    'Cheezy method for converting to unsigned integer
    If intValue >= 0 Then
        ConvToUnSignedValue = intValue
        Exit Function
    End If
    
    ConvToUnSignedValue = intValue + 65535

End Function


Sub FPS()
    'Delay until specified FPS achieved
    Do While mlngFrameTime + (1000 \ MAX_FPS) > DX7.TickCount
        DoEvents
    Loop
    mlngFrameTime = DX7.TickCount

    'Count FPS
    If mlngTimer + 1000 <= DX7.TickCount Then
        mlngTimer = DX7.TickCount
        mintFPS = mintFPSCounter + 1
        mintFPSCounter = 0
    Else
        mintFPSCounter = mintFPSCounter + 1
    End If
End Sub

Function Init() As Boolean
    On Error Resume Next
    
    Set DX7 = New DirectX7
    If DX7 Is Nothing Then
        MsgBox "Error Creating DirectX7 !", vbExclamation
        Exit Function
    End If
    
    Set DxD = DX7.DirectDrawCreate("")
    'set the cooperative level
    
    DxD.SetCooperativeLevel frmMain.hWnd, DDSCL_ALLOWMODEX Or DDSCL_FULLSCREEN Or DDSCL_EXCLUSIVE
    DxD.SetDisplayMode 1024, 768, 16, 0, DDSDM_DEFAULT
    
    'set the primary and backbuffer for flipping chain
    ddsmap.lFlags = DDSD_BACKBUFFERCOUNT Or DDSD_CAPS
    ddsmap.ddsCaps.lCaps = DDSCAPS_COMPLEX Or DDSCAPS_PRIMARYSURFACE Or DDSCAPS_FLIP Or DDSCAPS_VIDEOMEMORY
    ddsmap.lBackBufferCount = 1
    Set Primary = DxD.CreateSurface(ddsmap)
    Dim DD As DDSCAPS2
    DD.lCaps = DDSCAPS_BACKBUFFER
    Set BackBuffer = Primary.GetAttachedSurface(DD)
    
    BackBuffer.GetSurfaceDesc ddsmap
    
    Dim ddsd2 As DDSURFACEDESC2
    ddsd2.lFlags = DDSD_CAPS Or DDSD_WIDTH Or DDSD_HEIGHT
    ddsd2.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_VIDEOMEMORY
    ddsd2.lWidth = 1890
    ddsd2.lHeight = 930
    Set Display = DxD.CreateSurface(ddsd2)
    Set Layar = DxD.CreateSurface(ddsd2)
    
    ddsd2.lFlags = DDSD_CAPS Or DDSD_WIDTH Or DDSD_HEIGHT
    ddsd2.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_VIDEOMEMORY
    ddsd2.lWidth = 1024
    ddsd2.lHeight = 768
    Set XSpot = DxD.CreateSurface(ddsd2)
    
    XSpot.BltColorFill BoxRect(0, 0, 1024, 768), 0
    Display.BltColorFill BoxRect(0, 0, 1890, 930), 0
    
    'Make a new gamma controler
    Set mobjGammaControler = Primary.GetDirectDrawGammaControl

    'Fill out the original gamma ramps
    mobjGammaControler.GetGammaRamp DDSGR_DEFAULT, mudtOriginalRamp
    
    'Set our initial colour values to zero
    mintRedVal = 0
    mintGreenVal = 0
    mintBlueVal = 0
    
    frmMain.Show
    
End Function

Public Sub InitVariables()
    Dim X As Byte
    Dim Y As Byte
    For X = 1 To 60
        For Y = 1 To 30
            ArGenap(X, Y).Tipe = GRASS
        Next Y
    Next X
    
    'variabel awal
    AdaSoundCard = False
    StillRunning = True
    Scroll.ScrollX = GRASS_WIDTH
    Scroll.ScrollY = GRASS_HEIGHT \ 2
    
    OnToolbar = False
    ShowRoadGUI = False
    ShowBuildingGUI = False
    ShowMiniMap = True
    Delayment = 5
    EachTick = 0
    
    With Game
        .Budget = 300000
        .Tanggal = #1/1/2000#
    End With
End Sub


Public Function IsEven(Value As Integer) As Boolean
    IsEven = False
    If Value Mod 2 = 0 Then
        IsEven = True
    End If
End Function

Public Function LoadBitmap() As Boolean
    LoadBitmap = False
    'loading semua bitmap yang diperlukan
    If Not LoadSprite(App.Path & "\Graphics\Grass.Bmp", 0, GrassTex) Then
        MsgBox "Data Sprite [Grass.Bmp] Tidak Ditemukan !!", vbExclamation Or vbOKOnly
        Exit Function
    End If
    If Not LoadSprite(App.Path & "\Graphics\Roads.Bmp", 0, RoadTex) Then
        MsgBox "Data Sprite [Roads.Bmp] Tidak Ditemukan !!", vbExclamation Or vbOKOnly
        Exit Function
    End If
    If Not LoadSprite(App.Path & "\Graphics\Toolbar.Bmp", 0, ToolbarTex) Then
        MsgBox "Data Sprite [ToolBar.Bmp] Tidak Ditemukan !!", vbExclamation Or vbOKOnly
        Exit Function
    End If
    If Not LoadSprite(App.Path & "\Graphics\Pointers.Bmp", 0, MouseTex) Then
        MsgBox "Data Sprite [Pointers.Bmp] Tidak Ditemukan !!", vbExclamation Or vbOKOnly
        Exit Function
    End If
    If Not LoadSprite(App.Path & "\Graphics\RoadGUI.Bmp", 0, RoadGUI) Then
        MsgBox "Data Sprite [RoadGUI.Bmp] Tidak Ditemukan !!", vbExclamation Or vbOKOnly
        Exit Function
    End If
    If Not LoadSprite(App.Path & "\Graphics\Building.Bmp", 0, BuildingTex) Then
        MsgBox "Data Sprite [RoadGUI.Bmp] Tidak Ditemukan !!", vbExclamation Or vbOKOnly
        Exit Function
    End If
    If Not LoadSprite(App.Path & "\Graphics\Minimap.Bmp", 0, MiniMapTex) Then
        MsgBox "Data Sprite [Minimap.Bmp] Tidak Ditemukan !!", vbExclamation Or vbOKOnly
        Exit Function
    End If
    If Not LoadSprite(App.Path & "\Graphics\Park.Bmp", 0, ParkTex) Then
        MsgBox "Data Sprite [Park.Bmp] Tidak Ditemukan !!", vbExclamation Or vbOKOnly
        Exit Function
    End If
    If Not LoadSprite(App.Path & "\Graphics\GedungA.Bmp", 0, GedungA) Then
        MsgBox "Data Sprite [GedungA.Bmp] Tidak Ditemukan !!", vbExclamation Or vbOKOnly
        Exit Function
    End If
    If Not LoadSprite(App.Path & "\Graphics\GedungB.Bmp", 0, GedungB) Then
        MsgBox "Data Sprite [GedungB.Bmp] Tidak Ditemukan !!", vbExclamation Or vbOKOnly
        Exit Function
    End If
    If Not LoadSprite(App.Path & "\Graphics\GedungC.Bmp", 0, GedungC) Then
        MsgBox "Data Sprite [GedungC.Bmp] Tidak Ditemukan !!", vbExclamation Or vbOKOnly
        Exit Function
    End If
    If Not LoadSprite(App.Path & "\Graphics\PosTex.Bmp", 0, PosTex) Then
        MsgBox "Data Sprite [PosTex.Bmp] Tidak Ditemukan !!", vbExclamation Or vbOKOnly
        Exit Function
    End If
    If Not LoadSprite(App.Path & "\Graphics\GerejaTex.Bmp", 0, GerejaTex) Then
        MsgBox "Data Sprite [GerejaTex.Bmp] Tidak Ditemukan !!", vbExclamation Or vbOKOnly
        Exit Function
    End If
    If Not LoadSprite(App.Path & "\Graphics\PohonTex.Bmp", 0, PohonTex) Then
        MsgBox "Data Sprite [PohonTex.Bmp] Tidak Ditemukan !!", vbExclamation Or vbOKOnly
        Exit Function
    End If
    If Not LoadSprite(App.Path & "\Graphics\Listriktex.Bmp", 0, ListrikTex) Then
        MsgBox "Data Sprite [ListrikTex.Bmp] Tidak Ditemukan !!", vbExclamation Or vbOKOnly
        Exit Function
    End If
    If Not LoadSprite(App.Path & "\Graphics\Msgbox.Bmp", 0, MsgBoxTex) Then
        MsgBox "Data Sprite [Msgbox.Bmp] Tidak Ditemukan !!", vbExclamation Or vbOKOnly
        Exit Function
    End If
    If Not LoadSprite(App.Path & "\Graphics\HelpBox.Bmp", 0, HelpBox) Then
        MsgBox "Data Sprite [HelpBox.Bmp] Tidak Ditemukan !!", vbExclamation Or vbOKOnly
        Exit Function
    End If
    
    'atur tulisan di dalam building
    StFont.Bold = False
    StFont.Name = "Tahoma"
    StFont.Size = 8
    BuildingTex.SetFont StFont
    BuildingTex.SetForeColor RGB(0, 0, 0)
    BuildingTex.DrawText 214, 29, "Rumah Sederhana", False
    BuildingTex.DrawText 214, 42, "Rumah Menengah", False
    BuildingTex.DrawText 214, 55, "Rumah Kelas Atas", False
    BuildingTex.DrawText 214, 68, "Taman", False
    BuildingTex.DrawText 214, 81, "Pos Penjagaan", False
    BuildingTex.DrawText 214, 94, "Tempat Ibadah", False
    BuildingTex.DrawText 214, 107, "Pembangkit Listrik", False
    BuildingTex.DrawText 214, 120, "Pohon-pohon", False
    
    LoadBitmap = True
End Function


Sub Main()
    Call Init
    If Not LoadBitmap Then
        MsgBox "Silakan Install Ulang Program !!", vbOKOnly Or vbInformation
        Call EndAll
    End If
    
    Call InitVariables
    
    SFXMusik.aHwnd = frmMain.hWnd
    'mulai jalankan musik dan suara
    If SFXMusik.InitDM Then
        If SFXMusik.LoadMusic(App.Path & "\Music\starlight.Mid") Then
            SFXMusik.PlayMusic
        End If
    End If
    
    If SFXMusik.InitDS Then
        AdaSoundCard = True
        SFXON = True
    End If
    
    Call Render
End Sub


Sub Render()
    Call CreateTerrain
    'Call RefreshTerrain
    Call Screen.AmbilSkala
    
    BackBuffer.SetForeColor RGB(255, 255, 255)
    StFont.Bold = False
    StFont.Name = "Tahoma"
    StFont.Size = 8
    BackBuffer.SetFont StFont
    
    'SetGamma -50, -50, -50
    Do While StillRunning
        'XSpot.BltFast 0, 0, Display, BoxRect(Scroll.ScrollX, Scroll.ScrollY, Scroll.ScrollX + 1024, Scroll.ScrollY + 768), DDBLTFAST_WAIT
        BackBuffer.BltFast 0, 0, Display, BoxRect(Scroll.ScrollX, Scroll.ScrollY, Scroll.ScrollX + 1024, Scroll.ScrollY + 768), DDBLTFAST_WAIT
        'BackBuffer.BltFast 0, 0, XSpot, BoxRect(0, 0, 1024, 768), DDBLTFAST_WAIT
        
        If ShowRoadGUI Then
            BackBuffer.BltFast 0, 33, RoadGUI, BoxRect(0, 0, 254, 39), DDBLTFAST_WAIT
        End If
        If ShowBuildingGUI Then
            BackBuffer.BltFast 0, 73, BuildingTex, BoxRect(0, 0, 331, 217), DDBLTFAST_WAIT
        End If
        If ShowMiniMap Then
            BackBuffer.BltFast 802, 0, MiniMapTex, BoxRect(0, 0, 221, 158), DDBLTFAST_WAIT
            Call Screen.DrawMiniMap
        End If
        If ShowGoMap Then
            Call ShowGo
        End If
        If ShowHelp Then
            BackBuffer.BltFast 323, 300, HelpBox, BoxRect(0, 0, 378, 150), DDBLTFAST_WAIT
            BackBuffer.DrawText 343, 330, "Nama : Gunawan ", False
            BackBuffer.DrawText 343, 342, "NIM : 00 xxxx", False
            BackBuffer.DrawText 343, 354, "STMIK Mikroskil 2005 (July)", False
        End If
        
        'dapatkan koordinat kursor dunia
        CursorXDunia = CursorX + (Scroll.ScrollX - 63)
        CursorYDunia = CursorY + (Scroll.ScrollY - 15)
        ArXDunia = Int(CursorXDunia / (GRASS_WIDTH / 2) + 4)
        ArYDunia = Int((CursorYDunia / GRASS_HEIGHT) + 1)
        If ArYDunia > 0 And ArXDunia > 0 Then
            Select Case ArGenap(ArXDunia, ArYDunia).Tipe
            Case HOUSE
                BackBuffer.DrawText 5, 700, "TIPE : RUMAH", False
                Select Case ArGenap(ArXDunia, ArYDunia).HouseStyle
                Case 1
                    BackBuffer.DrawText 5, 714, "JENIS : RUMAH SEDERHANA", False
                Case 2
                    BackBuffer.DrawText 5, 714, "JENIS : RUMAH MENENGAH", False
                Case 3
                    BackBuffer.DrawText 5, 714, "JENIS : RUMAH MEWAH", False
                End Select
            Case ROAD
                BackBuffer.DrawText 5, 700, "TIPE : JALAN", False
            Case POS
                BackBuffer.DrawText 5, 700, "TIPE : POS PENJAGA", False
            Case CHURCH
                BackBuffer.DrawText 5, 700, "TIPE : RUMAH IBADAH", False
            Case TREES
                BackBuffer.DrawText 5, 700, "TIPE : POHON-POHON", False
            Case PARK
                BackBuffer.DrawText 5, 700, "TIPE : TAMAN", False
            Case ELECTRIC
                BackBuffer.DrawText 5, 700, "TIPE : PEMBANGKIT", False
            End Select
        End If
        
        Call HandleMouse
        Call Screen.CheckScroll
        Call CreateGUI
        Call ShowCursor
        Call AturCuacadanWaktu
        
        'BackBuffer.BltColorFill BoxRect(0, 0, 1024, 768), RGB(128, 128, 128)
        BackBuffer.DrawText 500, 10, "Frame :" & mintFPS, False
        'BackBuffer.DrawText 0, 340, Game.Tanggal, False
        'BackBuffer.DrawText 50, 0, (CursorX \ 63) & "," & CursorX & "," & CursorX / 63, False
        'BackBuffer.DrawText 0, 510, Scroll.ScrollX & "," & Scroll.ScrollY, False
        'BackBuffer.DrawText 0, 200, CursorX & "," & CursorY, False
        'BackBuffer.DrawText 0, 220, CursorXDunia & "," & CursorYDunia, False
        'BackBuffer.BltFast CursorX, CursorY, GrassTex, BoxRect(64, 0, 128, 31), DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
        'BackBuffer.DrawText 0, 100, IsEven(Int(CursorX + (Scroll.ScrollX / 63) / (GRASS_WIDTH / 2))), False
        'BackBuffer.DrawText 0, 120, Int((((CursorX - Scroll.ScrollX) / (GRASS_WIDTH / 2)) + 1) * (GRASS_WIDTH / 2)), False
        'BackBuffer.DrawText 0, 300, Int((((CursorX) / (GRASS_WIDTH / 2)) + 1)) + 4 + Int((Scroll.ScrollX \ (63 / 2))) & "," & Int((((CursorY)) / GRASS_HEIGHT) + 1) + 1, False
        
        'BackBuffer.DrawText 0, 320, ArXDunia & "," & ArYDunia, False
        
        If Not OnToolbar Then
            'ArXLocal = ArXDunia - (Scroll.ScrollX \ (GRASS_WIDTH \ 2)) - 3
            'ArYLocal = ArYDunia - (Scroll.ScrollY \ (GRASS_HEIGHT))
            'BackBuffer.DrawText 0, 340, ArXLocal & "," & ArYLocal, False
            
            'CursorXDunia = CursorX + Scroll.ScrollX
            'CursorYDunia = CursorY + Scroll.ScrollY
            If IsEven(ArXDunia) Then
                'Display.BltFast 10 * (GRASS_WIDTH \ 2), 10 * (GRASS_HEIGHT), GrassTex, BoxRect(64, 0, 128, 31), DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
                'BackBuffer.DrawText 300, 300, ArXLocal * (GRASS_WIDTH \ 2) & "," & ArYLocal * GRASS_HEIGHT, False
                'BackBuffer.BltFast (ArXLocal * GRASS_WIDTH \ 2) + (Scroll.ScrollX Mod GRASS_WIDTH) + 31, ArYLocal * GRASS_HEIGHT + (Scroll.ScrollY Mod (GRASS_HEIGHT \ 2)), GrassTex, BoxRect(64, 0, 128, 31), DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
                'BackBuffer.BltFast (ArXLocal * (GRASS_WIDTH \ 2)) + ((Scroll.ScrollX Mod GRASS_WIDTH) \ 2), ArYLocal * (GRASS_HEIGHT) - GRASS_HEIGHT, GrassTex, BoxRect(64, 0, 128, 31), DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
                'BackBuffer.BltFast Int((((CursorX - Scroll.ScrollX) / (GRASS_WIDTH / 2)) + 1)) * (GRASS_WIDTH / 2), Int((((CursorY - Scroll.ScrollY)) / GRASS_HEIGHT) + 1) * GRASS_HEIGHT, GrassTex, BoxRect(64, 0, 128, 31), DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
                'Display.BltFast Int((((CursorXDunia / (GRASS_WIDTH / 2)) + 1))) * (GRASS_WIDTH / 2), Int(((CursorYDunia) / GRASS_HEIGHT) + 1) * GRASS_HEIGHT - (GRASS_HEIGHT / 2), GrassTex, BoxRect(64, 0, 128, 31), DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
                'Display.BltFast Int((((CursorXDunia / (GRASS_WIDTH / 2)) + 1))) * (GRASS_WIDTH / 2), Int(((CursorY) / GRASS_HEIGHT) + 1) * GRASS_HEIGHT, GrassTex, BoxRect(64, 0, 128, 31), DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
                'Display.BltFast Int(((((CursorX + (Scroll.ScrollX Mod GRASS_WIDTH)) / (GRASS_WIDTH / 2)) * (GRASS_WIDTH / 2)))), Int((((CursorY + (Scroll.ScrollY Mod GRASS_HEIGHT))) / GRASS_HEIGHT) + 1) * GRASS_HEIGHT, GrassTex, BoxRect(64, 0, 128, 31), DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
            Else
                'BackBuffer.DrawText 300, 300, ArXLocal * (GRASS_WIDTH \ 2) & "," & ArYLocal * GRASS_HEIGHT + (GRASS_HEIGHT \ 2), False
                'BackBuffer.BltFast (ArXLocal * (GRASS_WIDTH \ 2)) + (Scroll.ScrollX Mod (GRASS_WIDTH \ 2)) + 31, ArYLocal * GRASS_HEIGHT + (GRASS_HEIGHT \ 2) + (Scroll.ScrollY Mod 15), GrassTex, BoxRect(64, 0, 128, 31), DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
                'A = CursorXDunia
                'BackBuffer.BltFast ((CursorX \ (GRASS_WIDTH \ 2)) + A) * GRASS_WIDTH, Int(((CursorYDunia) / GRASS_HEIGHT) + 1) * GRASS_HEIGHT, GrassTex, BoxRect(64, 0, 128, 31), DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
                'Display.BltFast Int((((CursorXDunia / (GRASS_WIDTH / 2)) + 1))) * (GRASS_WIDTH / 2), Int(((CursorYDunia) / GRASS_HEIGHT) + 1) * GRASS_HEIGHT, GrassTex, BoxRect(64, 0, 128, 31), DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
                
                'BackBuffer.BltFast Int((((CursorX - Scroll.ScrollX) / (GRASS_WIDTH / 2)) + 1)) * (GRASS_WIDTH / 2), Int(((CursorY - Scroll.ScrollY) / GRASS_HEIGHT) + 1) * GRASS_HEIGHT - (GRASS_HEIGHT / 2), GrassTex, BoxRect(64, 0, 128, 31), DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
                'BackBuffer.BltFast Int((((CursorX) / (GRASS_WIDTH / 2)) + 1)) * (GRASS_WIDTH / 2), Int(((CursorY) / GRASS_HEIGHT) + 1) * GRASS_HEIGHT - (GRASS_HEIGHT / 2), GrassTex, BoxRect(64, 0, 128, 31), DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
            End If
        End If
        
        'Display.BltFast Scroll.ScrollX, Scroll.ScrollY, XSpot, BoxRect(0, 0, 1024, 768), DDBLTFAST_WAIT
        
        'bersihkan daerah yang dispot
        'Display.BltFast CursorX, CursorY, XSpot, BoxRect(0, 0, 100, 100), DDBLTFAST_WAIT
        
        Call ShowStatus
        Primary.Flip Nothing, DDFLIP_WAIT
        FPS
        DoEvents
    Loop
End Sub
Public Function LoadSprite(StrFileName As String, nColor As Long, ByRef Tex As DirectDrawSurface7) As Boolean
    On Error GoTo Keluar
    'loading file bitmap apakah berhasil atau tidak
    Dim ddsspr As DDSURFACEDESC2
    ddsspr.lFlags = DDSD_CAPS
    ddsspr.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
    Set Tex = DxD.CreateSurfaceFromFile(StrFileName, ddsspr)
    
    'set the key color for the transparency thing
    Dim nColorKey As DDCOLORKEY
    nColorKey.high = nColor       'black
    nColorKey.low = nColor        'black
    Tex.SetColorKey DDCKEY_SRCBLT, nColorKey
    
    LoadSprite = True
    Exit Function
Keluar:
    MsgBox "Data Sprite tidak bisa diloading !", vbExclamation
    LoadSprite = False
End Function
Sub EndAll()
    On Error GoTo Keluar
    
    SFXMusik.StopMusic
    SFXMusik.CloseDM
    SFXMusik.CloseDS
    Set SFXMusik = Nothing
    
    StillRunning = False
    End
    Set DX7 = Nothing
    Set DxD = Nothing
Keluar:
    End
End Sub


Function BoxRect(X, Y, X1, Y1) As RECT
    BoxRect.Top = Y
    BoxRect.Left = X
    BoxRect.Right = X1
    BoxRect.Bottom = Y1
End Function

Public Sub ShowCursor()
    If (CursorX > 0 And CursorX < 300) And (CursorY > 0 And CursorY < 32) Then
        OnToolbar = True
        'highlight toolbar and change mouse pointer
        
        BackBuffer.SetForeColor RGB(128, 255, 128)
        Select Case CursorX
        Case 8 To 27
            BackBuffer.DrawBox 8, 7, 28, 27
            If Mouse_Button0 Then
                Mouse_Button0 = False
                Call EndAll
            End If
            Call ShowKata("Keluar Permainan")
        Case 32 To 49
            BackBuffer.DrawBox 32, 7, 50, 27        'Buka permainan
            If Mouse_Button0 Then
                Mouse_Button0 = False
                Call LoadGame
                Call RefreshTerrain
                Call Screen.RefreshMiniMap
            End If
            Call ShowKata("Buka Permainan")
        Case 56 To 74
            BackBuffer.DrawBox 56, 7, 76, 27        'simpan permainan
            If Mouse_Button0 Then
                Mouse_Button0 = False
                Call SaveGame
            End If
            Call ShowKata("Simpan Permainan")
        Case 89 To 106
            BackBuffer.DrawBox 89, 7, 106, 27       'World
            Call ShowKata("Mini Map")
            If Mouse_Button0 Then
                Mouse_Button0 = False
                'tampilkan toolbar jalan
                ShowMiniMap = Not ShowMiniMap
            End If
        Case 120 To 139
            BackBuffer.DrawBox 120, 7, 139, 27      'rumah
            Call ShowKata("Peralatan Perumahan")
            If Mouse_Button0 Then
                Mouse_Button0 = False
                'tampilkan toolbar jalan
                ShowBuildingGUI = Not ShowBuildingGUI
                If ShowBuildingGUI = False Then
                    SelectedBuild = False
                    SelectedChoice = 0
                End If
            End If
        Case 144 To 163
            BackBuffer.DrawBox 144, 7, 163, 27      'road/jalan
            Call ShowKata("Peralatan Jalan")
            If Mouse_Button0 Then
                Mouse_Button0 = False
                'tampilkan toolbar jalan
                ShowRoadGUI = Not ShowRoadGUI
            End If
        Case 169 To 187
            BackBuffer.DrawBox 169, 7, 187, 27      'hancur
            Call ShowKata("Peralatan Hancur")
            If Mouse_Button0 Then
                Mouse_Button0 = False
                RoadSelected = False
                SelectedBuild = False
                SelectedDestroy = True
            End If
        Case 192 To 211
            BackBuffer.DrawBox 192, 7, 211, 27
            If Mouse_Button0 Then
                Mouse_Button0 = False
                ShowGoMap = Not ShowGoMap
                RefreshGo = False
            End If
            Call ShowKata("Status Permainan")
        Case 223 To 243
            BackBuffer.DrawBox 223, 7, 243, 27      'Musik dan Efek Suara
            If Mouse_Button0 Then
                Mouse_Button0 = False
                'matikan tombol sfx
                SFXON = Not SFXON
            End If
            Call ShowKata("Musik dan Efek Suara")
        Case 245 To 266
            BackBuffer.DrawBox 245, 7, 266, 27      'Setting Permainan
            Call ShowKata("Setting Permainan")
        Case 271 To 291
            BackBuffer.DrawBox 271, 7, 291, 27      'help
            If Mouse_Button0 Then
                Mouse_Button0 = False
                'matikan tombol sfx
                ShowHelp = Not ShowHelp
            End If
            Call ShowKata("Panduan Permainan")
        End Select
        
        BackBuffer.BltFast CursorX, CursorY, MouseTex, BoxRect(31, 0, 46, 15), DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
        
    ElseIf (CursorX > 0 And CursorX < 254) And (CursorY > 33 And CursorY < 72) And ShowRoadGUI Then  'toolbar jalan
        'highlight road toolbar and change mouse pointer
        OnToolbar = True
        BackBuffer.SetForeColor RGB(128, 255, 128)
        
        Select Case CursorX
        Case 1 To 23
            BackBuffer.DrawBox 1, 48, 23, 70
            Call ShowKata("Jalan NE")
            If Mouse_Button0 Then
                Mouse_Button0 = False
                RoadType = ROAD_UP: RoadSelected = True
            End If
        Case 23 To 46
            BackBuffer.DrawBox 24, 48, 46, 70
            If Mouse_Button0 Then
                Mouse_Button0 = False
                RoadType = ROAD_DOWN: RoadSelected = True
            End If
        Case 47 To 69
            BackBuffer.DrawBox 47, 48, 69, 70
            'Call ShowKata("Simpan Permainan")
            If Mouse_Button0 Then
                Mouse_Button0 = False
                RoadType = ROAD_4WAY: RoadSelected = True
            End If
        Case 70 To 92
            BackBuffer.DrawBox 70, 48, 92, 70
            'Call ShowKata("Mini Map")
            If Mouse_Button0 Then
                Mouse_Button0 = False
                RoadType = ROAD_BENDUP: RoadSelected = True
            End If
        Case 93 To 115
            BackBuffer.DrawBox 93, 48, 115, 70
            'Call ShowKata("Peralatan Perumahan")
            If Mouse_Button0 Then
                Mouse_Button0 = False
                RoadType = ROAD_BENDRIGHT: RoadSelected = True
            End If
        Case 116 To 138
            BackBuffer.DrawBox 116, 48, 138, 70
            'Call ShowKata("Peralatan Jalan")
            If Mouse_Button0 Then
                Mouse_Button0 = False
                RoadType = ROAD_BENDDOWN: RoadSelected = True
            End If
        Case 139 To 161
            BackBuffer.DrawBox 139, 48, 161, 70
            'Call ShowKata("Peralatan Hancur")
            If Mouse_Button0 Then
                Mouse_Button0 = False
                RoadType = ROAD_BENDLEFT: RoadSelected = True
            End If
        Case 162 To 184
            BackBuffer.DrawBox 162, 48, 184, 70
            If Mouse_Button0 Then
                Mouse_Button0 = False
                RoadType = ROADT_UPRIGHT: RoadSelected = True
            End If
        Case 185 To 207
            BackBuffer.DrawBox 185, 48, 207, 70
            If Mouse_Button0 Then
                Mouse_Button0 = False
                RoadType = ROADT_UPLEFT: RoadSelected = True
            End If
        Case 208 To 230
            BackBuffer.DrawBox 208, 48, 230, 70
            If Mouse_Button0 Then
                Mouse_Button0 = False
                RoadType = ROADT_DOWNRIGHT: RoadSelected = True
            End If
        Case 231 To 253
            BackBuffer.DrawBox 231, 48, 253, 70
            If Mouse_Button0 Then
                Mouse_Button0 = False
                RoadType = ROADT_DOWNLEFT: RoadSelected = True
            End If
        End Select
        
        BackBuffer.BltFast CursorX, CursorY, MouseTex, BoxRect(31, 0, 46, 15), DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY

    ElseIf (CursorX > 802 And CursorX < 1024) And (CursorY > 0 And CursorY < 158) And ShowMiniMap Then  'minimap
        OnToolbar = True
        BackBuffer.BltFast CursorX, CursorY, MouseTex, BoxRect(0, 0, 15, 15), DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
        
    ElseIf (CursorX > 0 And CursorX < 331) And (CursorY > 73 And CursorY < 290) And ShowBuildingGUI Then  'building toolbar
        OnToolbar = True
        BackBuffer.BltFast CursorX, CursorY, MouseTex, BoxRect(31, 0, 46, 15), DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
        
        Select Case CursorX
        Case 214 To 320
            BuildingTex.SetForeColor RGB(255, 255, 255)
            BuildingTex.BltColorFill BoxRect(8, 25, 200, 129), 0
            BuildingTex.BltColorFill BoxRect(49, 137, 203, 183), 0
            BackBuffer.SetForeColor QBColor(5)
            Select Case CursorY
            Case 102 To 114
                BackBuffer.DrawLine 214, 114, 300, 114  'rumah kelas bawah, tampilkan keterangan
                Call ShowInformation(1)
                DimPilih = 1
            Case 115 To 127
                BackBuffer.DrawLine 214, 127, 300, 127  'rumah kelas menengah, tampilkan keterangan
                Call ShowInformation(2)
                DimPilih = 2
            Case 128 To 140
                BackBuffer.DrawLine 214, 140, 300, 140  'rumah kelas atas, tampilkan keterangan
                Call ShowInformation(3)
                DimPilih = 3
            Case 141 To 153
                BackBuffer.DrawLine 214, 153, 300, 153  'taman,tampilkan keterangan dan harga
                Call ShowInformation(4)
                DimPilih = PARK
            Case 154 To 166
                BackBuffer.DrawLine 214, 166, 300, 166  'pos penjagaan, tampilkan keterangan
                Call ShowInformation(5)
                DimPilih = POS
            Case 167 To 179
                BackBuffer.DrawLine 214, 179, 300, 179  'gEREJA , tampilkan keterangan
                Call ShowInformation(6)
                DimPilih = CHURCH
            Case 180 To 192
                BackBuffer.DrawLine 214, 192, 300, 192  'pembangkit listrik, tampilkan keterangan
                Call ShowInformation(7)
                DimPilih = ELECTRIC
            Case 193 To 205
                BackBuffer.DrawLine 214, 205, 300, 205  'pohon-pohon, tampilkan keterangan
                Call ShowInformation(8)
                DimPilih = TREES
            End Select
        End Select
        If Mouse_Button0 Then
            Mouse_Button0 = False
            'dipilih gedung sesuai dengan dimpilih
            SelectedDestroy = False
            RoadSelected = False
            SelectedBuild = True
            SelectedChoice = DimPilih
        End If
        
    ElseIf Scroll.blScroll = True Then
        Select Case Scroll.ScrollWay
        Case SCROLL_LEFT
            BackBuffer.BltFast CursorX, CursorY, MouseTex, BoxRect(75, 0, 90, 15), DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
        Case SCROLL_RIGHT
            BackBuffer.BltFast CursorX, CursorY, MouseTex, BoxRect(105, 0, 120, 15), DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
        Case SCROLL_UP
            BackBuffer.BltFast CursorX, CursorY, MouseTex, BoxRect(90, 0, 105, 15), DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
        Case SCROLL_DOWN
            BackBuffer.BltFast CursorX, CursorY, MouseTex, BoxRect(60, 0, 75, 15), DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
        End Select
    Else
        OnToolbar = False
        BackBuffer.BltFast CursorX, CursorY, MouseTex, BoxRect(0, 0, 15, 15), DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
        If Mouse_Button0 Then
            'BackBuffer.DrawText 0, 320, Int(CursorXDunia / (GRASS_WIDTH / 2) + 6) & "," & Int(CursorYDunia / GRASS_HEIGHT) + 2, False
            'ArXDunia = Int(CursorXDunia / (GRASS_WIDTH / 2) + 6)
            'ArYDunia = Int(CursorYDunia / GRASS_HEIGHT) + 2
            Mouse_Button0 = False
            If RoadSelected Then
                If AdaSoundCard And SFXON Then
                    If SFXMusik.LoadWave(App.Path & "\Music\construct.wav") Then
                        SFXMusik.PlayWave
                    End If
                End If
                'Debug.Print Int((((CursorX) / (GRASS_WIDTH / 2)) + 1)) + 5 & "," & Int((((CursorY)) / GRASS_HEIGHT) + 1) + 1
                Game.Budget = Game.Budget - 50
                Game.Pengeluaran = Game.Pengeluaran + 50
                ArGenap(ArXDunia, ArYDunia).Tipe = ROAD
                ArGenap(ArXDunia, ArYDunia).RoadStyle = RoadType
                Call RefreshTerrain: Call ShowStatus
                If ShowMiniMap Then Call Screen.RefreshMiniMap
            ElseIf SelectedDestroy Then
                If AdaSoundCard And SFXON Then
                    If SFXMusik.LoadWave(App.Path & "\Music\blast.wav") Then
                        SFXMusik.PlayWave
                    End If
                End If
                ArGenap(ArXDunia, ArYDunia).Tipe = GRASS
                'ArGenap(Int((((CursorX) / (GRASS_WIDTH / 2)) + 1) + 5), Int((((CursorY)) / GRASS_HEIGHT) + 1) + 1).RoadStyle = RoadType
                Call RefreshTerrain
                If ShowMiniMap Then Call Screen.RefreshMiniMap
            ElseIf SelectedBuild Then
                If AdaSoundCard And SFXON Then
                    If SFXMusik.LoadWave(App.Path & "\Music\construct.wav") Then
                        SFXMusik.PlayWave
                    End If
                End If
                
                Select Case SelectedChoice
                Case 1 To 3
                    ArGenap(ArXDunia, ArYDunia).Tipe = HOUSE
                    ArGenap(ArXDunia, ArYDunia).HouseStyle = SelectedChoice
                    Dim P As Byte
                    If SelectedChoice = 1 Then
                        P = Int(Rnd * 3) + 1
                    ElseIf SelectedChoice = 2 Then
                        P = Int(Rnd * 4) + 1
                    ElseIf SelectedChoice = 3 Then
                        P = Int(Rnd * 5) + 1
                    End If
                    ArGenap(ArXDunia, ArYDunia).Placed = IIf(P <= 2, True, False)
                    Game.Budget = Game.Budget - (SelectedChoice * 1000)
                    Game.Pengeluaran = Game.Pengeluaran + (SelectedChoice * 1000)
                    Call ShowStatus
                Case PARK
                    ArGenap(ArXDunia, ArYDunia).Tipe = PARK
                    'ArGenap(Int((((CursorX) / (GRASS_WIDTH / 2)) + 1) + 5), Int((((CursorY)) / GRASS_HEIGHT) + 1) + 1).HouseStyle = SelectedChoice
                    Game.Budget = Game.Budget - 1000
                    Game.Pengeluaran = Game.Pengeluaran + 1000
                    Call ShowStatus
                Case POS
                    'periksa posisi pos supaya jangan bersebelahan
                    If PosisiPosDekat(ArXDunia, ArYDunia) = False Then
                        ArGenap(ArXDunia, ArYDunia).Tipe = POS
                        'ArGenap(Int((((CursorX) / (GRASS_WIDTH / 2)) + 1) + 5), Int((((CursorY)) / GRASS_HEIGHT) + 1) + 1).HouseStyle = SelectedChoice
                        Game.Budget = Game.Budget - 4000
                        Game.Pengeluaran = Game.Pengeluaran + 4000
                        Call ShowStatus
                    End If
                Case CHURCH
                    ArGenap(ArXDunia, ArYDunia).Tipe = CHURCH
                    'ArGenap(Int((((CursorX) / (GRASS_WIDTH / 2)) + 1) + 5), Int((((CursorY)) / GRASS_HEIGHT) + 1) + 1).HouseStyle = SelectedChoice
                    Game.Budget = Game.Budget - 3500
                    Game.Pengeluaran = Game.Pengeluaran + 3500
                    Call ShowStatus
                Case TREES
                    ArGenap(ArXDunia, ArYDunia).Tipe = TREES
                    'ArGenap(Int((((CursorX) / (GRASS_WIDTH / 2)) + 1) + 5), Int((((CursorY)) / GRASS_HEIGHT) + 1) + 1).HouseStyle = SelectedChoice
                    Game.Budget = Game.Budget - 500
                    Game.Pengeluaran = Game.Pengeluaran + 500
                    Call ShowStatus
                Case ELECTRIC
                    ArGenap(ArXDunia, ArYDunia).Tipe = ELECTRIC
                    Game.Budget = Game.Budget - 30000
                    Game.Pengeluaran = Game.Pengeluaran + 30000
                    Call ShowStatus
                End Select
                Call RefreshTerrain
                If ShowMiniMap Then Call Screen.RefreshMiniMap
            End If
        ElseIf Mouse_Button1 Then
            Mouse_Button1 = False
            If AdaSoundCard And SFXON Then
                If SFXMusik.LoadWave(App.Path & "\Music\blast.wav") Then
                    SFXMusik.PlayWave
                End If
            End If
            ArGenap(ArXDunia, ArYDunia).Tipe = GRASS
            'ArGenap(Int((((CursorX) / (GRASS_WIDTH / 2)) + 1) + 5), Int((((CursorY)) / GRASS_HEIGHT) + 1) + 1).RoadStyle = RoadType
            Call RefreshTerrain
            If ShowMiniMap Then Call Screen.RefreshMiniMap
        End If
    End If
End Sub

Public Sub ShowGo()
    If RefreshGo = False Then
        RefreshGo = True
        'mulai perhitungan
        Dim X As Integer, Y As Integer
        Rmh1 = 0: Rmh2 = 0: Rmh3 = 0
        For X = 1 To 60
            For Y = 1 To 30
                If ArGenap(X, Y).Tipe = HOUSE Then
                    Select Case ArGenap(X, Y).HouseStyle
                    Case 1
                        Rmh1 = Rmh1 + 1
                    Case 2
                        Rmh2 = Rmh2 + 1
                    Case 3
                        Rmh3 = Rmh3 + 1
                    End Select
                ElseIf ArGenap(X, Y).Tipe = ROAD Then
                    RoadCnt = RoadCnt + 1
                ElseIf ArGenap(X, Y).Tipe = CHURCH Then
                    ChurchCnt = ChurchCnt + 1
                ElseIf ArGenap(X, Y).Tipe = POS Then
                    PosCnt = PosCnt + 1
                ElseIf ArGenap(X, Y).Tipe = TREES Then
                    TreesCnt = TreesCnt + 1
                ElseIf ArGenap(X, Y).Tipe = ELECTRIC Then
                    ElectricCnt = ElectricCnt + 1
                ElseIf ArGenap(X, Y).Tipe = PARK Then
                    ParkCnt = ParkCnt + 1
                End If
            Next Y
        Next X
        'tampilkan hasil dan perhitungannya,berapa jumlah setiap objek yang dibangun
        TotalRmh = Rmh1 + Rmh2 + Rmh3
        Polusi = ((TotalRmh * 150) + (ElectricCnt * 1000) - ((ParkCnt * 50) + (TreesCnt * 10))) / 100
        Infrastruktur = (((TotalRmh * 50) + (PosCnt * 50) + (ChurchCnt * 50) + (ParkCnt * 50)) - RoadCnt * 20) / 100
        Lingkungan = (((TotalRmh * 50) + (PosCnt * 50) + (ChurchCnt * 150) + (ParkCnt * 50) + (TreesCnt * 50)) - (ElectricCnt * 200)) / 100
        Aman = ((TotalRmh * 50) + (ChurchCnt * 50) - (PosCnt * 100)) / 100
        Investasi = (Infrastruktur + Lingkungan + Aman - Polusi) / 100
    End If
    
    'rutin untuk menampilkan informasi permainan sejauh ini
    BackBuffer.SetForeColor RGB(255, 10, 255)
    BackBuffer.BltFast 207, 131, MsgBoxTex, BoxRect(0, 0, 609, 506), DDBLTFAST_WAIT
    With BackBuffer
        .DrawText 225, 161, "Jumlah Bangunan/Objek", False
        .DrawText 225, 168, "------------------------------", False
        
        .SetForeColor RGB(255, 255, 255)
        .DrawText 225, 178, "Rumah Sederhana :" & Format(Rmh1, "#,##0"), False
        .DrawText 225, 190, "Rumah Menengah :" & Format(Rmh2, "#,##0"), False
        .DrawText 225, 202, "Rumah Kelas Atas :" & Format(Rmh3, "#,##0"), False
        .DrawText 225, 214, "Taman :" & Format(ParkCnt, "#,##0"), False
        .DrawText 225, 226, "Pos Penjagaan :" & Format(PosCnt, "#,##0"), False
        .DrawText 225, 238, "Rumah Ibadah :" & Format(ChurchCnt, "#,##0"), False
        .DrawText 225, 250, "Generator :" & Format(ElectricCnt, "#,##0"), False
        .DrawText 225, 262, "Jalan :" & Format(RoadCnt, "#,##0"), False
        
        'mulai kalkulasi perhitungan pantas dan tidak
        .SetForeColor RGB(255, 10, 255)
        .DrawText 225, 290, "Status dan Kondisi terakhir", False
        .DrawText 225, 297, "---------------------------------", False
        .SetForeColor RGB(255, 255, 255)
        .DrawText 225, 307, "Infrastruktur Jalan :" & Format(Infrastruktur, "#,##0.00"), False
        .DrawText 225, 319, "Lingkungan :" & Format(Lingkungan, "#,##0.00"), False
        .DrawText 225, 331, "Keamanan :" & Format(Aman, "#,##0.00"), False
        .DrawText 225, 343, "Polusi :" & Format(Polusi, "#,##0.00"), False
        .DrawText 225, 355, "Nilai Investasi :" & Format(Investasi, "#,##0.00"), False
                
        If ElectricCnt > 0 Then
        
        If (TotalRmh \ ElectricCnt) >= 50 Then
            Ket1 = "Jumlah Generator sudah memadai"
        Else
            Ket1 = "Perlu ditambahkan generator untuk listrik"
        End If
        
        Else
            Ket1 = "Belum ada generator untuk listrik"
        End If
        
        If RoadCnt > 0 Then
        If ((ParkCnt + TotalRmh + PosCnt + ChurchCnt) \ RoadCnt) >= 1 Then
            Ket2 = "Jumlah jalan sudah memadai"
        Else
            Ket2 = "Masih kurang jalan sebagai sarana transportasi"
        End If
        Else
            Ket2 = "Belum dibangun jalan"
        End If
        
        If RoadCnt > 0 Then
        If ((ParkCnt + TotalRmh + PosCnt + ChurchCnt) \ RoadCnt) >= 1 Then
            Ket2 = "Jumlah jalan sudah memadai"
        Else
            Ket2 = "Masih kurang jalan sebagai sarana transportasi"
        End If
        Else
            Ket2 = "Belum dibangun jalan"
        End If
        
        If ParkCnt > 0 Then
            If ((TotalRmh + (ElectricCnt * 50)) \ ParkCnt) >= 10 Then
                Ket3 = "Jumlah Taman sudah memenuhi syarat"
            Else
                Ket3 = "Taman masih kurang"
            End If
        Else
            Ket3 = "Belum dibangun taman"
        End If
        
        If ChurchCnt > 0 Then
            If ChurchCnt > 3 Then
                Ket4 = "Jumlah rumah ibadah sudah cukup"
            Else
                Ket4 = "Masih kurang dari jumlah yang sesuai"
            End If
        Else
            Ket4 = "Belum dibangun rumah ibadah"
        End If
        
        If PosCnt > 0 Then
            If (TotalRmh + ParkCnt) \ PosCnt >= 10 Then
                Ket5 = "Jumlah Pos Penjaga sudah cukup"
            Else
                Ket5 = "Masih banyak kejahatan, bangun pos lebih banyak"
            End If
        Else
            Ket5 = "Belum dibangun pos penjaga"
        End If
        
        'keterangan dari kalkulasi
        .SetForeColor RGB(255, 10, 255)
        .DrawText 225, 383, "Keterangan/Masukan", False
        .DrawText 225, 390, "-----------------------", False
        .SetForeColor RGB(255, 255, 255)
        .DrawText 225, 400, "Listrik :" & Ket1, False
        .DrawText 225, 412, "Taman :" & Ket3, False
        .DrawText 225, 424, "Jalan :" & Ket2, False
        .DrawText 225, 436, "Pos Penjaga :" & Ket5, False
        .DrawText 225, 448, "Rumah Ibadah :" & Ket4, False
        .DrawText 225, 460, "Keseluruhan : masih butuh latihan, walau sudah bagus", False
        
        'selama permainan
        .SetForeColor RGB(255, 10, 255)
        .DrawText 225, 488, "Hasil", False
        .DrawText 225, 495, "-------", False
        .SetForeColor RGB(255, 255, 255)
        .DrawText 225, 505, "Pendapatan :" & Format(Game.Pendapatan, "#,##0.00"), False
        .DrawText 225, 517, "Pengeluran :" & Format(Game.Pengeluaran, "#,##0.00"), False
    End With
End Sub
Public Sub ShowKata(Str As String)
    BackBuffer.BltColorFill BoxRect(221, 751, 219, 766), 13
    BackBuffer.SetForeColor RGB(255, 255, 255)
    BackBuffer.DrawText 224, 752, Str, False
End Sub

Public Sub ShowStatus()
    With Game
        BackBuffer.SetForeColor RGB(255, 255, 255)
        BackBuffer.BltColorFill BoxRect(10, 751, 100, 766), 13
        BackBuffer.BltColorFill BoxRect(135, 751, 190, 766), 13
        BackBuffer.DrawText 13, 753, Format(.Tanggal, "dd/mmmm/yyyy"), False
        BackBuffer.DrawText 140, 753, Format(.Budget, "#,##0.00"), False
    End With
End Sub





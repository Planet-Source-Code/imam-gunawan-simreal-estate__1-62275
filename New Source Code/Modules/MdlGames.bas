Attribute VB_Name = "MdlGames"
Option Explicit

'NIlai Konstanta di dalam Game
Public Const GRASS_WIDTH = 63
Public Const GRASS_HEIGHT = 31

'Scroll Constanta
Public Const SCROLL_NONE = 0
Public Const SCROLL_LEFT = 1
Public Const SCROLL_RIGHT = 2
Public Const SCROLL_UP = 3
Public Const SCROLL_DOWN = 4

'Road Constanta
Public Const ROAD_4WAY = 0
Public Const ROADT_UPRIGHT = 1
Public Const ROADT_UPLEFT = 2
Public Const ROADT_DOWNRIGHT = 3
Public Const ROADT_DOWNLEFT = 4
Public Const ROAD_BENDDOWN = 5
Public Const ROAD_BENDLEFT = 6
Public Const ROAD_BENDRIGHT = 7
Public Const ROAD_BENDUP = 8
Public Const ROAD_UP = 9
Public Const ROAD_DOWN = 10

Public Const GRASS = 1
Public Const ROAD = 2
Public Const HOUSE = 3
Public Const PARK = 4
Public Const ELECTRIC = 5
Public Const POS = 6
Public Const CHURCH = 7
Public Const TREES = 8

'scrolling concept
Type udtScroll
    blScroll As Boolean
    ScrollWay As Byte
    ScrollX As Long
    ScrollY As Long
End Type

'game status
Type udtGame
    Pendapatan As Long
    Pengeluaran As Long
    Budget As Long
    Tanggal As Date
    JumlahRumah As Integer
    JumlahJalan As Integer
    JumlahPohon As Integer
    JumlahListrik As Integer
    JumlahPos As Integer
    JumlahIbadah As Integer
End Type

Type udtArray
    Tipe As Byte
    Placed As Boolean
    RoadStyle As Byte
    HouseStyle As Byte
End Type

'Font variable
Public StFont As New StdFont

'Main Variables
Public ArGenap(1 To 60, 1 To 30) As udtArray

'Variable used in the game
Public Scroll As udtScroll
Public Screen As New clsScreen
Public ShowRoadGUI As Boolean
Public ShowBuildingGUI As Boolean
Public ShowMiniMap As Boolean
Public ShowGoMap As Boolean
Public RefreshGo As Boolean
Public ShowHelp As Boolean

'Road variables
Public RoadSelected As Boolean
Public RoadType As Byte
Public A As Integer
Public B As Integer
Public Selesai As Boolean
Public DuaKali As Byte
   
Public Game As udtGame
Public OnToolbar As Boolean
Public SelectedBuild As Boolean
Public SelectedChoice As Byte
Public SelectedDestroy As Boolean
Public CursorXDunia As Long
Public CursorYDunia As Long
Public ArXDunia As Integer
Public ArYDunia As Integer
Public ArXLocal As Integer
Public ArYLocal As Integer

Public AdaSoundCard As Boolean
Public SFXON As Boolean

Public EachTick As Integer
Public Delayment As Integer
Public Sub AturCuacadanWaktu()
    If EachTick > MAX_FPS * Delayment Then
        Game.Tanggal = Game.Tanggal + 1
        Call ShowStatus
        EachTick = 0
        'perhitungan cuaca
        Dim CuacaMendung As Byte
        CuacaMendung = Int(Rnd * 2) + 1
        If CuacaMendung = 1 Then    'hujan
            SetGamma -15, -15, -15
        Else
            SetGamma 0, 0, 0
        End If
        'perhitungan pendapatan
        Dim X As Integer, Y As Integer
        For X = 1 To 60
            For Y = 1 To 30
            
                If ArGenap(X, Y).Tipe = HOUSE And (ArGenap(X, Y).Placed = True) Then
                
                Select Case ArGenap(X, Y).HouseStyle
                Case 1
                    Game.Budget = Game.Budget + 10
                    Game.Pendapatan = Game.Pendapatan + 10
                Case 2
                    Game.Budget = Game.Budget + 15
                    Game.Pendapatan = Game.Pendapatan + 15
                Case 3
                    Game.Budget = Game.Budget + 17.5
                    Game.Pendapatan = Game.Pendapatan + 17.5
                End Select
                
                End If
                
            Next Y
        Next X
    Else
        EachTick = EachTick + 1
    End If
End Sub

Public Sub CreateGUI()
    BackBuffer.SetForeColor RGB(0, 0, 255)
    BackBuffer.DrawBox 0, 750, 1024, 768
    BackBuffer.BltColorFill BoxRect(1, 751, 1023, 767), 13
    BackBuffer.DrawLine 120, 750, 120, 767
    BackBuffer.DrawLine 220, 750, 220, 767
    BackBuffer.BltFast 0, 0, ToolbarTex, BoxRect(0, 0, 300, 32), DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
    BackBuffer.SetForeColor RGB(255, 255, 255)
End Sub


Public Sub CreateRoad()
    Dim X As Byte, Y As Byte
    For X = 5 To 10
        For Y = 10 To 10
            'Display.BltFast GRASS_WIDTH * (X - 1), GRASS_HEIGHT * (Y - 1), RoadTex, BoxRect(ROAD_DOWN * 63, 0, (ROAD_DOWN * 63) + 64, 31), DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
            Display.BltFast GRASS_WIDTH * (X - 1), GRASS_HEIGHT * (Y - 1), RoadTex, BoxRect(ROAD_DOWN * 64, 0, (ROAD_DOWN * 64) + 64, 31), DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
        Next Y
    Next X
End Sub

Public Sub CreateTerrain()
    Dim Y As Integer
    Dim X As Integer
    Display.BltColorFill BoxRect(0, 0, 1024, 768), 0
    For X = 1 To 30
        For Y = 1 To 30
            Display.BltFast GRASS_WIDTH * (X - 1), GRASS_HEIGHT * (Y - 1), GrassTex, BoxRect(0, 0, 63, 31), DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
        Next Y
    Next X
    
    'now fill the blank
    For X = 1 To 30
        For Y = 1 To 30
            Display.BltFast GRASS_WIDTH * (X - 1) + (GRASS_WIDTH \ 2), GRASS_HEIGHT * (Y - 1) + (GRASS_HEIGHT \ 2), GrassTex, BoxRect(0, 0, 63, 31), DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
        Next Y
    Next X
End Sub

Public Sub HandleMouse()
    'check the position of mouse and place the spot
    
End Sub


Public Sub LoadGame()
    'rutin untuk membuka permainan yang telah disimpan
    'rutin ini akan melakukan penyimpanan terhadap game
    Open App.Path & "\SaveGame\Game.Gun" For Binary Access Read As #1
    
    'array permainan
    Get #1, 1, ArGenap
    
    'Variabel Game
    Get #1, , Game
    
    Get #1, , Scroll.ScrollX
    Get #1, , Scroll.ScrollY
    
    Close #1
End Sub

Public Sub RefreshTerrain()
    Dim Y As Integer
    Dim X As Integer
    
    Display.BltColorFill BoxRect(0, 0, 1890, 930), 0
    Display.SetFont StFont
    
    For X = 1 To 60
        For Y = 1 To 30
            If X Mod 2 = 1 Then
                If ArGenap(X, Y).Tipe = GRASS Or ArGenap(X, Y).Tipe = HOUSE Or ArGenap(X, Y).Tipe = TREES Or ArGenap(X, Y).Tipe = CHURCH Or ArGenap(X, Y).Tipe = ELECTRIC Then
                    Display.BltFast GRASS_WIDTH * ((X \ 2) - 1), GRASS_HEIGHT * (Y - 1), GrassTex, BoxRect(0, 0, 63, 31), DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
                    'Display.DrawText GRASS_WIDTH * ((X \ 2) - 1) + 20, GRASS_HEIGHT * (Y - 1), X & "," & Y, False
                ElseIf ArGenap(X, Y).Tipe = ROAD Then
                    Display.BltFast GRASS_WIDTH * ((X \ 2) - 1), GRASS_HEIGHT * (Y - 1), RoadTex, BoxRect(ArGenap(X, Y).RoadStyle * 64, 0, (ArGenap(X, Y).RoadStyle * 64) + 63, 31), DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
                    'Display.DrawText GRASS_WIDTH * ((X \ 2) - 1) + 20, GRASS_HEIGHT * (Y - 1), X & "," & Y, False
                End If
            Else
                If ArGenap(X, Y).Tipe = GRASS Or ArGenap(X, Y).Tipe = HOUSE Or ArGenap(X, Y).Tipe = TREES Or ArGenap(X, Y).Tipe = CHURCH Or ArGenap(X, Y).Tipe = ELECTRIC Then
                    Display.BltFast GRASS_WIDTH * (((X - 1) \ 2) - 1) + (GRASS_WIDTH \ 2), GRASS_HEIGHT * (Y - 1) + (GRASS_HEIGHT \ 2), GrassTex, BoxRect(0, 0, 63, 31), DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
                    'Display.DrawText GRASS_WIDTH * (((X - 1) \ 2) - 1) + (GRASS_WIDTH \ 2) + 20, GRASS_HEIGHT * (Y - 1) + (GRASS_HEIGHT \ 2), X & "," & Y, False
                ElseIf ArGenap(X, Y).Tipe = ROAD Then
                    Display.BltFast GRASS_WIDTH * (((X - 1) \ 2) - 1) + (GRASS_WIDTH \ 2), GRASS_HEIGHT * (Y - 1) + (GRASS_HEIGHT \ 2), RoadTex, BoxRect(ArGenap(X, Y).RoadStyle * 64, 0, (ArGenap(X, Y).RoadStyle * 64) + 63, 31), DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
                    'Display.DrawText GRASS_WIDTH * (((X - 1) \ 2) - 1) + (GRASS_WIDTH \ 2) + 20, GRASS_HEIGHT * (Y - 1) + (GRASS_HEIGHT \ 2), X & "," & Y, False
                End If
            End If
        Next Y
    Next X
    
    'draw routine gedung
    For X = 60 To 1 Step -1
        A = X: B = 1: Selesai = False:
        If A Mod 2 = 0 Then
            DuaKali = 1
        Else
            DuaKali = 2
        End If
        Do While A < 60 And Not Selesai
            A = A + 1
            If A Mod 2 = 1 Then
                If ArGenap(A, B).Tipe = HOUSE Then
                    Select Case ArGenap(A, B).HouseStyle
                    Case 1
                        Display.BltFast GRASS_WIDTH * ((A \ 2) - 1), GRASS_HEIGHT * (B - 1) - 8, GedungA, BoxRect(0, 0, 64, 39), DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
                    Case 2
                        Display.BltFast GRASS_WIDTH * ((A \ 2) - 1), GRASS_HEIGHT * (B - 1) - 11, GedungB, BoxRect(0, 0, 64, 42), DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
                    Case 3
                        Display.BltFast GRASS_WIDTH * ((A \ 2) - 1), GRASS_HEIGHT * (B - 1) - 20, GedungC, BoxRect(0, 0, 64, 51), DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
                    End Select
                ElseIf ArGenap(A, B).Tipe = TREES Then
                    Display.BltFast GRASS_WIDTH * ((A \ 2) - 1), GRASS_HEIGHT * (B - 1) - 27, PohonTex, BoxRect(0, 0, 64, 58), DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
                ElseIf ArGenap(A, B).Tipe = CHURCH Then
                    Display.BltFast GRASS_WIDTH * ((A \ 2) - 1), GRASS_HEIGHT * (B - 1) - 10, GerejaTex, BoxRect(0, 0, 64, 41), DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
                ElseIf ArGenap(A, B).Tipe = POS Then
                    Display.BltFast GRASS_WIDTH * ((A \ 2) - 1), GRASS_HEIGHT * (B - 1) - 6, PosTex, BoxRect(0, 0, 64, 37), DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
                ElseIf ArGenap(A, B).Tipe = ELECTRIC Then
                    Display.BltFast GRASS_WIDTH * ((A \ 2) - 1), GRASS_HEIGHT * (B - 1) - 65, ListrikTex, BoxRect(0, 0, 64, 96), DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
                ElseIf ArGenap(A, B).Tipe = PARK Then
                    Display.BltFast GRASS_WIDTH * ((A \ 2) - 1), GRASS_HEIGHT * (B - 1) - 9, ParkTex, BoxRect(0, 0, 64, 40), DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
                End If
            Else
                If ArGenap(A, B).Tipe = HOUSE Then
                    Select Case ArGenap(A, B).HouseStyle
                    Case 1
                        Display.BltFast GRASS_WIDTH * (((A - 1) \ 2) - 1) + (GRASS_WIDTH \ 2), GRASS_HEIGHT * (B - 1) + (GRASS_HEIGHT \ 2) - 8, GedungA, BoxRect(0, 0, 64, 39), DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
                    Case 2
                        Display.BltFast GRASS_WIDTH * (((A - 1) \ 2) - 1) + (GRASS_WIDTH \ 2), GRASS_HEIGHT * (B - 1) + (GRASS_HEIGHT \ 2) - 11, GedungB, BoxRect(0, 0, 64, 42), DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
                    Case 3
                        Display.BltFast GRASS_WIDTH * (((A - 1) \ 2) - 1) + (GRASS_WIDTH \ 2), GRASS_HEIGHT * (B - 1) + (GRASS_HEIGHT \ 2) - 20, GedungC, BoxRect(0, 0, 64, 51), DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
                    End Select
                ElseIf ArGenap(A, B).Tipe = TREES Then
                    Display.BltFast GRASS_WIDTH * (((A - 1) \ 2) - 1) + (GRASS_WIDTH \ 2), GRASS_HEIGHT * (B - 1) + (GRASS_HEIGHT \ 2) - 27, PohonTex, BoxRect(0, 0, 64, 58), DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
                ElseIf ArGenap(A, B).Tipe = CHURCH Then
                    Display.BltFast GRASS_WIDTH * (((A - 1) \ 2) - 1) + (GRASS_WIDTH \ 2), GRASS_HEIGHT * (B - 1) + (GRASS_HEIGHT \ 2) - 10, GerejaTex, BoxRect(0, 0, 64, 41), DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
                ElseIf ArGenap(A, B).Tipe = POS Then
                    Display.BltFast GRASS_WIDTH * (((A - 1) \ 2) - 1) + (GRASS_WIDTH \ 2), GRASS_HEIGHT * (B - 1) + (GRASS_HEIGHT \ 2) - 6, PosTex, BoxRect(0, 0, 64, 37), DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
                ElseIf ArGenap(A, B).Tipe = ELECTRIC Then
                    Display.BltFast GRASS_WIDTH * (((A - 1) \ 2) - 1) + (GRASS_WIDTH \ 2), GRASS_HEIGHT * (B - 1) + (GRASS_HEIGHT \ 2) - 65, ListrikTex, BoxRect(0, 0, 64, 96), DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
                ElseIf ArGenap(A, B).Tipe = PARK Then
                    Display.BltFast GRASS_WIDTH * (((A - 1) \ 2) - 1) + (GRASS_WIDTH \ 2), GRASS_HEIGHT * (B - 1) + (GRASS_HEIGHT \ 2) - 9, ParkTex, BoxRect(0, 0, 64, 40), DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
                End If
            End If
            DuaKali = DuaKali + 1
            If DuaKali > 2 Then
                B = B + 1
                DuaKali = 1
            End If
            If B > 30 Then
                Selesai = True
            End If
        Loop
    Next X
    
    'bagian ketiga, membersihkan petak terakhir
    For Y = 1 To 30
        A = Y: B = 0: Selesai = False:
        If B Mod 2 = 0 Then
            DuaKali = 1
        Else
            DuaKali = 2
        End If
        Do While B < 30 And Not Selesai
            B = B + 1
            If B Mod 2 = 1 Then
                If ArGenap(B, A).Tipe = HOUSE Then
                    Select Case ArGenap(B, A).HouseStyle
                    Case 1
                        Display.BltFast GRASS_WIDTH * ((B \ 2) - 1), GRASS_HEIGHT * (A - 1) - 8, GedungA, BoxRect(0, 0, 64, 39), DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
                    Case 2
                        Display.BltFast GRASS_WIDTH * ((B \ 2) - 1), GRASS_HEIGHT * (A - 1) - 11, GedungB, BoxRect(0, 0, 64, 42), DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
                    Case 3
                        Display.BltFast GRASS_WIDTH * ((B \ 2) - 1), GRASS_HEIGHT * (A - 1) - 20, GedungC, BoxRect(0, 0, 64, 51), DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
                    End Select
                ElseIf ArGenap(B, A).Tipe = TREES Then
                    Display.BltFast GRASS_WIDTH * ((B \ 2) - 1), GRASS_HEIGHT * (A - 1) - 27, PohonTex, BoxRect(0, 0, 64, 58), DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
                ElseIf ArGenap(B, A).Tipe = CHURCH Then
                    Display.BltFast GRASS_WIDTH * ((B \ 2) - 1), GRASS_HEIGHT * (A - 1) - 10, GerejaTex, BoxRect(0, 0, 64, 41), DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
                ElseIf ArGenap(B, A).Tipe = POS Then
                    Display.BltFast GRASS_WIDTH * ((B \ 2) - 1), GRASS_HEIGHT * (A - 1) - 6, PosTex, BoxRect(0, 0, 64, 37), DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
                ElseIf ArGenap(B, A).Tipe = ELECTRIC Then
                    Display.BltFast GRASS_WIDTH * ((B \ 2) - 1), GRASS_HEIGHT * (A - 1) - 65, ListrikTex, BoxRect(0, 0, 64, 96), DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
                ElseIf ArGenap(B, A).Tipe = PARK Then
                    Display.BltFast GRASS_WIDTH * ((B \ 2) - 1), GRASS_HEIGHT * (A - 1) - 9, ParkTex, BoxRect(0, 0, 64, 40), DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
                End If
            Else
                If ArGenap(B, A).Tipe = HOUSE Then
                    Select Case ArGenap(B, A).HouseStyle
                    Case 1
                        Display.BltFast GRASS_WIDTH * (((B - 1) \ 2) - 1) + (GRASS_WIDTH \ 2), GRASS_HEIGHT * (A - 1) + (GRASS_HEIGHT \ 2) - 8, GedungA, BoxRect(0, 0, 64, 39), DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
                    Case 2
                        Display.BltFast GRASS_WIDTH * (((B - 1) \ 2) - 1) + (GRASS_WIDTH \ 2), GRASS_HEIGHT * (A - 1) + (GRASS_HEIGHT \ 2) - 11, GedungB, BoxRect(0, 0, 64, 42), DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
                    Case 3
                        Display.BltFast GRASS_WIDTH * (((B - 1) \ 2) - 1) + (GRASS_WIDTH \ 2), GRASS_HEIGHT * (A - 1) + (GRASS_HEIGHT \ 2) - 20, GedungC, BoxRect(0, 0, 64, 51), DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
                    End Select
                ElseIf ArGenap(B, A).Tipe = TREES Then
                    Display.BltFast GRASS_WIDTH * (((B - 1) \ 2) - 1) + (GRASS_WIDTH \ 2), GRASS_HEIGHT * (A - 1) + (GRASS_HEIGHT \ 2) - 27, PohonTex, BoxRect(0, 0, 64, 58), DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
                ElseIf ArGenap(B, A).Tipe = CHURCH Then
                    Display.BltFast GRASS_WIDTH * (((B - 1) \ 2) - 1) + (GRASS_WIDTH \ 2), GRASS_HEIGHT * (A - 1) + (GRASS_HEIGHT \ 2) - 10, GerejaTex, BoxRect(0, 0, 64, 41), DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
                ElseIf ArGenap(B, A).Tipe = POS Then
                    Display.BltFast GRASS_WIDTH * (((B - 1) \ 2) - 1) + (GRASS_WIDTH \ 2), GRASS_HEIGHT * (A - 1) + (GRASS_HEIGHT \ 2) - 6, PosTex, BoxRect(0, 0, 64, 37), DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
                ElseIf ArGenap(B, A).Tipe = ELECTRIC Then
                    Display.BltFast GRASS_WIDTH * (((B - 1) \ 2) - 1) + (GRASS_WIDTH \ 2), GRASS_HEIGHT * (A - 1) + (GRASS_HEIGHT \ 2) - 65, ListrikTex, BoxRect(0, 0, 64, 96), DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
                ElseIf ArGenap(B, A).Tipe = PARK Then
                    Display.BltFast GRASS_WIDTH * (((B - 1) \ 2) - 1) + (GRASS_WIDTH \ 2), GRASS_HEIGHT * (A - 1) + (GRASS_HEIGHT \ 2) - 9, ParkTex, BoxRect(0, 0, 64, 40), DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
                End If
            End If
            DuaKali = DuaKali + 1
            If DuaKali > 2 Then
                A = A + 1
                DuaKali = 1
            End If
            If A > 30 Then
                Selesai = True
            End If
        Loop
    Next Y
End Sub

Public Sub SaveGame()
    'rutin ini akan melakukan penyimpanan terhadap game
    Open App.Path & "\SaveGame\Game.Gun" For Binary Access Write As #1
    
    'array permainan
    Put #1, 1, ArGenap
    Put #1, , Game
    Put #1, , Scroll.ScrollX
    Put #1, , Scroll.ScrollY
    
    Close #1
End Sub

Public Sub ShowInformation(POS As Byte)
    Select Case POS
    Case 1
        BuildingTex.BltFast 76, 61, GedungA, BoxRect(0, 0, 64, 39), DDBLTFAST_WAIT
        BuildingTex.DrawText 54, 139, "Rumah Sederhana ", False
        BuildingTex.DrawText 54, 154, "1,000", False
        BuildingTex.DrawText 54, 170, "Rumah", False
    Case 2
        BuildingTex.BltFast 76, 60, GedungB, BoxRect(0, 0, 64, 42), DDBLTFAST_WAIT
        BuildingTex.DrawText 54, 139, "Rumah Menengah", False
        BuildingTex.DrawText 54, 154, "2,000", False
        BuildingTex.DrawText 54, 170, "Rumah", False
    Case 3
        BuildingTex.BltFast 76, 60, GedungC, BoxRect(0, 0, 64, 51), DDBLTFAST_WAIT
        BuildingTex.DrawText 54, 139, "Rumah Kelas Atas", False
        BuildingTex.DrawText 54, 154, "3,000", False
        BuildingTex.DrawText 54, 170, "Rumah", False
    Case 4
        BuildingTex.BltFast 76, 61, ParkTex, BoxRect(0, 0, 64, 40), DDBLTFAST_WAIT
        BuildingTex.DrawText 54, 139, "Taman Bunga", False
        BuildingTex.DrawText 54, 154, "1,000", False
        BuildingTex.DrawText 54, 170, "Taman", False
    Case 5
        BuildingTex.BltFast 76, 61, PosTex, BoxRect(0, 0, 64, 37), DDBLTFAST_WAIT
        BuildingTex.DrawText 54, 139, "Pos Penjagaan", False
        BuildingTex.DrawText 54, 154, "4,000", False
        BuildingTex.DrawText 54, 170, "Taman", False
    Case 6
        BuildingTex.BltFast 76, 61, GerejaTex, BoxRect(0, 0, 64, 41), DDBLTFAST_WAIT
        BuildingTex.DrawText 54, 139, "Tempat Ibadah", False
        BuildingTex.DrawText 54, 154, "3,500", False
        BuildingTex.DrawText 54, 170, "Taman", False
    Case 7
        BuildingTex.BltFast 76, 25, ListrikTex, BoxRect(0, 0, 64, 96), DDBLTFAST_WAIT
        BuildingTex.DrawText 54, 139, "Pembangkit Listrik", False
        BuildingTex.DrawText 54, 154, "30,000", False
        BuildingTex.DrawText 54, 170, "Taman", False
    Case 8
        BuildingTex.BltFast 76, 61, PohonTex, BoxRect(0, 0, 64, 58), DDBLTFAST_WAIT
        BuildingTex.DrawText 54, 139, "Pohon-Pohon", False
        BuildingTex.DrawText 54, 154, "150", False
        BuildingTex.DrawText 54, 170, "Taman", False
    End Select
End Sub


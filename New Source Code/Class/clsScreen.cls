VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsScreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public XScroll As Integer
Public YScroll As Integer
Public Redraw As Boolean
Public SkalaMap As Single
Public WidthMap As Single
Public HeightMap As Single
Public JarakX As Single
Public JarakY As Single
Public RedrawMiniMap As Boolean
Public Sub AmbilSkala()
    SkalaMap = 60 / 202
    WidthMap = 32 / SkalaMap
    HeightMap = 24 / SkalaMap
End Sub

Public Sub CheckScroll()
    'mouse digunakan untuk melakukan kontrol terhadap permainan
    'baik itu penggulungan dan sebagainya, dengan nilai cursorX/Y
    With Scroll
        .blScroll = False
        .ScrollWay = SCROLL_NONE
        'MousePointer = MOUSE_DEFAULT
    
        If (CursorX <= 0 Or CursorX >= 950) Or (CursorY <= 0 Or CursorY >= 740) Then
        
        Call UpdateMinimap
        .blScroll = True
        'ganti kursor mouse
    
        'then check for which way the screen should scroll
        Select Case CursorX
        Case Is <= 0
            .ScrollWay = SCROLL_LEFT
            If .ScrollX > GRASS_WIDTH Then
                .ScrollX = .ScrollX - 5
            ElseIf .ScrollX < GRASS_WIDTH Then
                .ScrollX = GRASS_WIDTH
            End If
            
        Case Is >= 775
            .ScrollWay = SCROLL_RIGHT
            If .ScrollX < 820 Then .ScrollX = .ScrollX + 5
        End Select
        
        Select Case CursorY
        Case Is >= 740
            .ScrollWay = SCROLL_DOWN
            If .ScrollY < 150 Then .ScrollY = .ScrollY + 5
        Case Is <= 0
            .ScrollWay = SCROLL_UP
            If .ScrollY > (GRASS_HEIGHT \ 2) Then
                .ScrollY = .ScrollY - 5
            ElseIf .ScrollY < (GRASS_HEIGHT \ 2) Then
                .ScrollY = (GRASS_HEIGHT \ 2)
            End If
        End Select
        End If
    End With
End Sub
Public Sub DrawMiniMap()
    BackBuffer.SetForeColor QBColor(14)
    BackBuffer.DrawBox 812 + CInt(JarakX), 27 + CInt(JarakY), 812 + CInt(JarakX + WidthMap), 27 + CInt(JarakY + HeightMap)
End Sub

Public Sub RefreshMiniMap()
    'rutin ini hanya dijalankan apabila Minimap akan direfresh
    'dengan data baru
    Dim X As Byte
    Dim Y As Byte
    'If RedrawMiniMap Then
    
    MiniMapTex.BltColorFill BoxRect(10, 31, 210, 122), RGB(0, 100, 0)
    
    'melakukan penggambaran terhadap minimap dengan pengisian setiap pixel dengan nilai array
    For X = 1 To 60
        For Y = 1 To 30
            If ArGenap(Int(X * SkalaMap) + 1, Int(Y * SkalaMap) + 1).Tipe = ROAD Then 'titik jalan
                MiniMapTex.BltFast X + 10, Y + 33, RoadTex, BoxRect(58, 0, 59, 1), DDBLTFAST_WAIT
            ElseIf ArGenap(Int(X * SkalaMap) + 1, Int(Y * SkalaMap) + 1).Tipe = HOUSE Then
                MiniMapTex.BltFast X + 10, Y + 33, RoadTex, BoxRect(6, 15, 7, 16), DDBLTFAST_WAIT
            Else
            
            End If
        Next Y
    Next X
    
    RedrawMiniMap = False
    'End If
    
    'transfer ke backbuffer
    BackBuffer.BltFast 802, 0, MiniMapTex, BoxRect(0, 0, 221, 158), DDBLTFAST_WAIT
End Sub


Public Sub UpdateMinimap()
    Dim TmpX As Single
    Dim TmpY As Single
    
    'rutin untuk melakukan update terhadap layar kecil
    TmpX = Scroll.ScrollX / GRASS_WIDTH
    TmpY = Scroll.ScrollY / GRASS_HEIGHT
    
    JarakX = TmpX / SkalaMap
    JarakY = TmpY / SkalaMap
    
    BackBuffer.SetForeColor QBColor(14)
    BackBuffer.DrawBox 812 + CInt(JarakX), 27 + CInt(JarakY), 812 + CInt(JarakX + WidthMap), 27 + CInt(JarakY + HeightMap)
End Sub



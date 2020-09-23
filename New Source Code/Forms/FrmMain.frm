VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Form1"
   ClientHeight    =   3660
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5355
   LinkTopic       =   "Form1"
   ScaleHeight     =   244
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   357
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Implements DirectXEvent

'The DirectX declarations
Dim DX As New DirectX7 'Root object
'Dim DXEvent As DirectXEvent
Dim Di As DirectInput 'DirectInput root object
Dim DiDev As DirectInputDevice 'Represents our Mouse

'How big our buffer of events will be.
'The bigger it gets the slower it gets. 10 is fine.
Const BufferSize = 10

'Button Flags for Mouse
Dim Button_0 As Boolean
Dim Button_1 As Boolean
Dim Button_2 As Boolean
Dim Button_3 As Boolean

'EventHandle holds the number representing
'the DirectXEvent.
Dim EventHandle As Long
Dim NotActive As Boolean

''''''''''API declarations
'GetCursorPos: We must know where the mouse is when we start
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
'SetCursorPos: When we lose the mouse; we tell windows where it is
Private Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
'ScreenToClient: Because the CursorPos is related to the screen (0,0 is the top of the screen)
'                      We use this to convert it so that 0,0 is the top corner of our form
Private Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
'ClientToScreen: Same as the previous one, but the other way around.
Private Declare Function ClientToScreen Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long

'A simple custom type that's used by
'some of the API calls.
Private Type POINTAPI
        X As Long
        Y As Long
End Type


Private Sub DirectXEvent_DXCallback(ByVal eventid As Long)
'This procedure is automatically called by DirectX
'when something has changed. Because we know that it's
'changed we'll check everything and update things accordingly

'NB: as the sensitivity/speed increases the movement becomes less
'accurate and more jerky. Anything above 4/5 will cause this.

  Dim DIDeviceData(1 To BufferSize) As DIDEVICEOBJECTDATA
  Dim NumItems As Integer
  Dim i As Integer
  Static OldSequence As Long
  
  ' Get data
  On Error GoTo INPUTLOST
  NumItems = DiDev.GetDeviceData(DIDeviceData, 0)
  On Error GoTo 0
  
  ' Process data
  For i = 1 To NumItems
    Select Case DIDeviceData(i).lOfs
      Case DIMOFS_X
        CursorX = CursorX + DIDeviceData(i).lData * MOUSESPEED
           
        ' We don't want to update the cursor in response to
        ' separate axis movements, or we will get a staircase instead of diagonal movement.
        ' A diagonal movement of the mouse results in two events with the same sequence number.
        ' In order to avoid postponing the last event till the mouse moves again, we always
        ' reset OldSequence after it's been tested once.
          
        If OldSequence <> DIDeviceData(i).lSequence Then
          UpdateCursor
          OldSequence = DIDeviceData(i).lSequence
        Else
          OldSequence = 0
        End If
         
      Case DIMOFS_Y
        CursorY = CursorY + DIDeviceData(i).lData * MOUSESPEED
        If OldSequence <> DIDeviceData(i).lSequence Then
          UpdateCursor
          OldSequence = DIDeviceData(i).lSequence
        Else
          OldSequence = 0
        End If
      'Case DIMOFS_Z
                'If you want to use the Z axis uncomment this
        
      'Check the mouse buttons
      Case DIMOFS_BUTTON0
                Button_0 = True
                Mouse_Button0 = True
                If DIDeviceData(i).lData = 0 Then  ' button up
                    Button_0 = False
                    Mouse_Button0 = False
                End If
      Case DIMOFS_BUTTON1
                Button_1 = True
                Mouse_Button1 = True
                If DIDeviceData(i).lData = 0 Then  ' button up
                    Button_1 = False
                    Mouse_Button1 = True
                End If
      Case DIMOFS_BUTTON2
                Button_2 = True
                Mouse_Button2 = True
                If DIDeviceData(i).lData = 0 Then  ' button up
                    Button_2 = False
                    Mouse_Button2 = True
                End If
      Case DIMOFS_BUTTON3
                Button_3 = True
                Mouse_Button3 = True
                If DIDeviceData(i).lData = 0 Then  ' button up
                    Button_3 = False
                    Mouse_Button3 = True
                End If
    End Select
  Next i
  
  Exit Sub
  
INPUTLOST:
  ' Windows stole the mouse from us. DIERR_INPUTLOST is raised if the user switched to
  ' another app, but DIERR_NOTACQUIRED is raised if the Windows key was pressed.
  If (Err.Number = DIERR_INPUTLOST) Or (Err.Number = DIERR_NOTACQUIRED) Then
    'We must clear up after ourselves; or things
    'may go pear shaped :)
    CleanUp
    Exit Sub
  End If
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        DiDev.Unacquire
        Unload Me
        End
    End If
End Sub

Sub UpdateCursor()
'We'll keep the mouse within our form area; this way it
'is easy to draw the mouse cursor.

'When DirectInput tells us what the mouse is doing it
'tells us NOT where it is; it tells us what it's just done.
'ie, x+1, Y-1 etc... We then translate this so that it +1 or -1
'of the current coordinates. This way it is easy to keep it within
'the form.
  If CursorX < 0 Then CursorX = 0
  If CursorX >= Me.ScaleWidth - 15 Then CursorX = Me.ScaleWidth - 16
  
  If CursorY < 0 Then CursorY = 0
  If CursorY >= Me.ScaleHeight + 15 Then CursorY = Me.ScaleHeight + 14
  
  'Because we are only operating within windows; we'll use an
  'image control to represent the mouse. In DirectDraw this could
  '(and probably should) be replaced by Blitting a cursor to the correct
  'coordinates.
  'If you use DirectInput in normal windows operations bare in mind
  'that an image control may well go behind other controls - so you'll
  'have to bare in mind the Z-Order of controls.
End Sub


Sub CleanUp()
'We'll call this when the user has finished *playing*
'It basically gets rid of DirectInput's control (but DOESN'T close it down)
'and it tells windows where it's cursor is.
  Dim m_point As POINTAPI
  
  'Unlink ourselves from the mouse -
  'give up control of it.
  DiDev.Unacquire
  
  'Copy the variables
  m_point.X = CursorX
  m_point.Y = CursorY
  'Convert the coordinates from our
  'local variables to screen variables
  Call ClientToScreen(hWnd, m_point)
  'Conversion done. Update the cursor properly.
  Call SetCursorPos(m_point.X, m_point.Y)
End Sub


Sub Initialise()
''''''''UI Stuff
CursorX = Me.ScaleWidth \ 2
CursorY = Me.ScaleHeight \ 2

Set Di = DX.DirectInputCreate
Set DiDev = Di.CreateDevice("GUID_SYSMOUSE")
Call DiDev.SetCommonDataFormat(DIFORMAT_MOUSE)
Call DiDev.SetCooperativeLevel(Me.hWnd, DISCL_FOREGROUND Or DISCL_EXCLUSIVE)

'''''''''We require a buffer of events to stop
'''''''''us from missing any. Set the constant
'''''''''in the declarations section to change the size
'''''''''although, 10 should be fine.
Dim Property As DIPROPLONG
Property.lHow = DIPH_DEVICE
Property.lObj = 0
Property.lData = BufferSize
Property.lSize = Len(Property)
Call DiDev.SetProperty("DIPROP_BUFFERSIZE", Property)

'''''''''Create our Events. These events tell us when something
''''''''has happened (position change for example)
EventHandle = DX.CreateEvent(Me)
Call DiDev.SetEventNotification(EventHandle)

AquireMouse
End Sub

Sub AquireMouse()
'We create a seperate prcedure for this
'as it can be called from different places in the code.

  Dim CursorCoord As POINTAPI
  
  'These are the API calls
  'First we get where the mouse currently is.
  'As soon as we acquire the mouse we wont be able
  'to ask windows.
  Call GetCursorPos(CursorCoord)
  'Convert the information that we just got to
  'local coordinates - 0,0 now = the top left of our window.
  Call ScreenToClient(hWnd, CursorCoord)
  
  On Error GoTo AQUIREERROR
  'Simple! - we now aquire the mouse; at the same
  'time Windows loses it.
  DiDev.Acquire
  'Copy the mouse's coordinates to our internal variables.
  CursorX = CursorCoord.X
  CursorY = CursorCoord.Y
  
  'Update the cursor position
  UpdateCursor
  On Error GoTo 0
  Exit Sub

AQUIREERROR:
  Exit Sub
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        EndAll
    End If
End Sub

Private Sub Form_Load()
    DoEvents
    Call Initialise
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim didevstate As DIMOUSESTATE
  
  On Error GoTo NOTYETACQUIRED
  Call DiDev.GetDeviceStateMouse(didevstate)
  On Error GoTo 0
  Exit Sub
  
NOTYETACQUIRED:
  Call AquireMouse
End Sub


Private Sub Form_Unload(Cancel As Integer)
    If EventHandle <> 0 Then DX.DestroyEvent EventHandle
End Sub



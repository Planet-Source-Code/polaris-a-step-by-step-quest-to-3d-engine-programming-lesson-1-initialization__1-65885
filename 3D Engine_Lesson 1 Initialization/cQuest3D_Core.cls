VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cQuest3D_Core"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'==============================================================================================================
'
'       THIS IS THE CORE CLASS OF THE ENGINE
'
'  - we start 3D and close 3D here
'  - We acces the GFX here
'  - we change all GFX state here
'================================================================================================


'Lesson 1
'
'We write
'  - an initialization method
'  - Begin3D to start 3d rendering
'  - End3D to close 3D drawing and flip all draw stuff to the screen
'  - Free to close engine
'  - GetKeyPressed() to get the current pressed key


'=====================================================
'Method:Init
' Create 3D device and all objects needed
'
'
'
'Return: Boolean
' return true if the engine creation suceed
' return false if there was an error
'=====================================================
Function Init(ByVal FormHandle As Long)



    Set obj_D3D = obj_DX.Direct3DCreate() 'we create the D3D object
    
    Set obj_D3DX = New D3DX8
    
    
     ' Get The current Display Mode format
    Dim mode As D3DDISPLAYMODE
    obj_D3D.GetAdapterDisplayMode D3DADAPTER_DEFAULT, mode
    
    
    'we set the parameters for the creation of the Device
    
    Data.WinParam.Windowed = 1
    Data.WinParam.SwapEffect = D3DSWAPEFFECT_COPY_VSYNC
    Data.WinParam.BackBufferFormat = mode.Format

    'create the device
      'D3DADAPTER_DEFAULT means we use default Graphic card VGA is the first
      'D3DCREATE_SOFTWARE_VERTEXPROCESSING we use directx driver software vertex processing
    
    Set obj_Device = obj_D3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, FormHandle, D3DCREATE_SOFTWARE_VERTEXPROCESSING, Data.WinParam)

    'if the device is nothing we fail else Great WE DID IT nice!!
    Data.Is_engineActive = Not (obj_Device Is Nothing)
    
    'we prepare informations
    If Data.Is_engineActive Then
      DoDefault
    End If
    
    Init = Data.Is_engineActive

End Function

Private Sub DoDefault()
'here we set default parameters

'we set the color to blue
 Data.BackBuff_ClearColor = D3DColorXRGB(0, 0, 255)
 
 Me.Set_Font "arial", 12

End Sub

Sub Freeengine()
  
    'we deallocate memory and free objects
    Set obj_DX = Nothing
    Set obj_D3D = Nothing
    Set obj_Device = Nothing
    
    Data.Is_engineActive = False

End Sub



'=======================================
'We tell D3D that we Begin 3D rendering
'
'======================================
Sub Begin3D()

   ComputeFrames

   If Data.Is_engineActive = False Then Exit Sub

  'we clear previous drawn
  obj_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, Data.BackBuff_ClearColor, 1#, 0
  'We tell the device we are ready to render triangles
  obj_Device.BeginScene
  
End Sub


'==========================================================
'Here we compute frame per second
'
'==========================================================
Private Sub ComputeFrames()
    Dim Elapsed As Single
    
    'we get the current time in second
    Data.Fps_TimePassed = timeGetTime * 0.001 - Data.Fps_CurrentTime
    Data.Fps_CurrentTime = timeGetTime * 0.001
    'we count Frames
    Data.Fps_FrameCounter = Data.Fps_FrameCounter + 1
    
    'we compute time passed during two rendering scene
    Elapsed = Data.Fps_CurrentTime - Data.Fps_LastTime

'if we are over 1 sec
If Elapsed > 1 Then
  'we reset the last time to Now
  Data.Fps_LastTime = Data.Fps_CurrentTime
  
  'we check that we have 1 second,1.004=85 is not true so we do (85 x 1)/1.004
  If Elapsed <> 0 Then _
  Data.Fps_FramePerSecond = Data.Fps_FrameCounter / (Elapsed)
  'we reset the counter
  Data.Fps_FrameCounter = 0

End If


End Sub

Function Get_FramesPerSeconde() As Single
        Get_FramesPerSeconde = Data.Fps_FramePerSecond
End Function

Function Get_TimePassed() As Single
        Get_TimePassed = Data.Fps_TimePassed
End Function

'=========================================
'We tell D3D that we Stop 3D rendering
'=========================================
Sub End3D()

   If Data.Is_engineActive = False Then Exit Sub

  obj_Device.EndScene 'We tell the device we'are now finished
  obj_Device.Present ByVal 0, ByVal 0, 0, ByVal 0 'we flip all drawn stuffs over the monitor' screen

End Sub


Function Get_KeyPressed(ByVal Key As KeyCodeConstants) As Boolean
   Get_KeyPressed = (GetAsyncKeyState(Key) <> 0)
End Function


'============================================================
'Engine State
'
'============================================================
Sub Set_BackbufferClearColor(ByVal ClearColor As Long)
    Data.BackBuff_ClearColor = ClearColor
End Sub

Private Sub Class_Initialize()
  Randomize Timer
End Sub

Sub Set_Font(ByVal FontName As String, ByVal PoliceSize As Integer)

    Set Data.StFont = New StdFont

    Data.StFont.Name = FontName
    Data.StFont.Size = PoliceSize
    Set Data.FontDesc = Data.StFont

    Set Data.MainFont = obj_D3DX.CreateFont(obj_Device, Data.FontDesc.hFont)

End Sub


Sub Draw_Text(ByVal St As String, ByVal x As Long, ByVal y As Long, Optional ByVal color As Long = &HFFFFCC00)

    Dim TextRect As RECT

    TextRect.Top = y

    TextRect.Left = x

   
    obj_D3DX.DrawText Data.MainFont, color, St, TextRect, DT_TOP Or DT_LEFT
   

End Sub

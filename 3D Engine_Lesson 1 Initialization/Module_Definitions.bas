Attribute VB_Name = "Module_Definitions"
Option Explicit
'==============================================================================================================
'
'       In this module we define all objects and types
'       we will use in the Engine like,3D Device object,iformations
'       like matricies ect....
'
'
'================================================================================================




'global objects

'for accessing to all functions provided by Directx Lib
  Global obj_DX As New DirectX8
  
'this object is an interface that provide functions and methods.
'this routines allow to check if the real 3D device has some required
'capabilities for a 3D engine
  Global obj_D3D As Direct3D8
  
  
'this engine is an interface that communicate
'directly with the 3D GFX Card
  Global obj_Device As Direct3DDevice8
  
  
  Global obj_D3DX As D3DX8



'=======================================================================
' here we define all type that will be required
'
'
'=======================================================================

Public Type NemoCFG
    
    'actual  screen width
    Buffer_Width As Integer
    'actual screen_height
    Buffer_Height As Integer
    'screen Rectangle (left,right,top,bottom values)
    Buffer_Rect As RECT
    'are we in windowed mode
    Is_Windowed As Boolean
    'dephtbit size
    Bpp As Integer
    'the engine is active
    Is_engineActive As Boolean
    'color for the back buffer
    BackBuff_ClearColor As Long
    
   
    'for font
    MainFont As D3DXFont
    StFont As StdFont
    FontDesc As IFont

  
  
    'handle of the form or the windows interface
    Hwindow As Long
    
    
    'device creation parameters
    WinParam As D3DPRESENT_PARAMETERS
    
    
    'for frame counter
    Fps_CurrentTime As Single
    Fps_LastTime As Single
    Fps_FrameCounter As Single
    Fps_FramePerSecond As Single
    Fps_TimePassed As Single
   
End Type

Public Data As NemoCFG

'some apis
'to retrieve keyboard state
Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
'to retrieve time
Declare Function timeGetTime Lib "winmm.dll" () As Long

'some types



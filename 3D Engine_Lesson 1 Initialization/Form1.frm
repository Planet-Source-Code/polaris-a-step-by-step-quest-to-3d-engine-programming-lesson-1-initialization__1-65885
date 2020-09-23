VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "3D Engine_Lesson 1 Initialization"
   ClientHeight    =   6765
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7935
   LinkTopic       =   "Form1"
   ScaleHeight     =   6765
   ScaleWidth      =   7935
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'==================================================================================
'WELCOME Engine_Lesson 1 Initialization
'_____________________________________________________________________________________
'-------------------------------------------------------------------------------------
'
'===================================================================================
'Welcome to this Step by step Quest to a 3D Engine programming
'this tutorial will show you how to design a simple 3D
'engine, Next tutorials will show how to add other engine objet
'like Camera,Mesh and Object Polygon
'
'This tutorial 1:3D Engine_Lesson 1 Initialization
'
'It shows you how to initialize a 3D device
'  - how to compute engine Frames per second
'  - change engine state like backbuffer color,Drawing font
'  - close the engine
'  - for any bug se polaris johna_pop@yahoo.fr
'
'
'How to read this code
'   - Form1: is the engine code in action
'   - Module_definitions will hold all engine objets definitions and types
'   - cQuest3D_Core is our first object, it defines Main entry of the engine
'
'
'Good coding
'
'Vote if you want the sequel!!
'
'==================================================================================







'we use the engine here
'we declare an objet
Dim QUEST As cQuest3D_Core


Private Sub Form_Load()

        'we allocate memory here
        Set QUEST = New cQuest3D_Core
        
        Me.Refresh
        Me.Show
        
        'we initialize the engine
        If QUEST.Init(Me.hWnd) = False Then
         MsgBox "Sorry there was an error"
         End
        End If
        
        'we call game loop
        GameLoop

End Sub

Sub GameLoop()

    Do
          'change the clear color randomely
          If QUEST.Get_KeyPressed(vbKeySpace) Then QUEST.Set_BackbufferClearColor (D3DColorXRGB(Rnd * 255, Rnd * 255, Rnd * 255))
          'we begin 3D
          QUEST.Begin3D
          '
          '
          'Drawing code will be added in Next tutorials
        
          
          'draw FPS
          QUEST.Draw_Text "FPS=" + CStr(QUEST.Get_FramesPerSeconde), 1, 10, &HFFFFFFFF
          'we close 3D
          QUEST.End3D
          DoEvents
          
         If QUEST.Get_KeyPressed(vbKeyEscape) Then Call CloseGame
    Loop

End Sub

'we quit game here
Sub CloseGame()
  QUEST.Freeengine
  End
End Sub

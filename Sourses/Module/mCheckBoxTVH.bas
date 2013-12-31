Attribute VB_Name = "mCheckBoxTVH"
'///////////////////////////////////////// Truong Van Hieu ////////////////////////////////////////
'////////////////////////////////// tvhhh2003@yahoo.com /////////////////////////////////////
'//////////////////////////////////// Special for Vietnamese /////////////////////////////////////
Option Explicit

Public Enum E_MouseEvent
    eMouseLeaving = 0
    eMouseLeavingClicking = 1
    eMouseMoving = 2
    eMouseMovingClicking = 3
End Enum

#If False Then 'Trick preserves Case of Enums when typing in IDE

    Private eMouseLeaving, eMouseLeavingClicking, eMouseMoving, eMouseMovingClicking
#End If

Public Enum E_AlignmentCheckBox
    ecbLeft = 0
    ecbRight = 1
    ecbTop = 2
    ecbBottom = 3
End Enum

#If False Then 'Trick preserves Case of Enums when typing in IDE

    Private ecbLeft, ecbRight, ecbTop, ecbBottom
#End If

Attribute VB_Name = "Module1"
'***********************************************************
'**   NMEA - AIS Parse Sentence Version 1.0  Feb-2008     **
'**                                                       **
'**           Joe Petrix                                  **
'**                                                       **
'**                                                       **
'**                                                       **
'***********************************************************

Public Type MSG_1 ' Position report
    ID_MSG As Byte
    Repeat_indicator As Byte
    MMSI As Long
    Navigation_Status As Byte ' 0 - OK Motore, 1 ancora - 2 non controllato -
    Rate_of_turn As Byte '°/min
    Speed_Over_Ground As Single
    Position_accuracy As Boolean
    varLongitude As Single
    varLatitude As Single
    Course_Over_Ground As Single
    True_Heading As Long
    time_from_report As Byte
End Type

Public Type MSG_5 ' Static & Voyage related Data
    strCourse1 As String
    strReference1 As String
    strCourse2 As String
    strReference2 As String
    strSpeed1 As String
    strSpeed2 As String
    strSpeedUnit1 As String
    strSpeedUnit2 As String
End Type








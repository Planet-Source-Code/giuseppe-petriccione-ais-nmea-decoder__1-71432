VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6270
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9375
   LinkTopic       =   "Form1"
   ScaleHeight     =   6270
   ScaleWidth      =   9375
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Clear"
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   1440
      Width           =   1095
   End
   Begin VB.ListBox List1 
      Height          =   5130
      Left            =   1320
      TabIndex        =   2
      Top             =   840
      Width           =   4815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Decode"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   405
      Left            =   120
      TabIndex        =   0
      Text            =   "..¥lpá.€£fñö..E..Yb»..<..¿..ôõ..ôó.¼:˜UÁ÷ûÒ›iDP...y‡..!AIVDM,1,1,,A,13=FEH002A15b8nDt1v`anrv0`C2,0*2C.."
      Top             =   240
      Width           =   8895
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit

Dim objMSG_1 As MSG_1
Dim objMSG_5 As MSG_5
Public AISString As String


   Private Function BinToDec(ByVal sIn As String) As Long
   Dim x As Integer
   BinToDec = 0
   For x = 1 To Len(sIn)
     BinToDec = BinToDec + (CInt(Mid(sIn, x, 1)) * (2 ^ (Len(sIn) - x)))
   Next x
End Function

Private Sub Command1_Click()

    Dim MSG_1 As MSG_1
    InBuff = Text1.Text

    Dim x, Z, st, fn, intst As Integer
    Dim strSentence, strTemp, strDummy As String
    Dim i, ch
    Dim JP_the_parser As String

    ' aggiunta reference ms script runtime -- utilizza dictionary al posto di Collection
    Dim table                                              'As Dictionary      ' è la tabella Ascii 6 bit usata in AIS.. ma proprio quì la devo caricare ?
    Set table = CreateObject("Scripting.Dictionary")
    table.Add "0", "000000"
    table.Add "1", "000001"
    table.Add "2", "000010"
    table.Add "3", "000011"
    table.Add "4", "000100"
    table.Add "5", "000101"
    table.Add "6", "000110"
    table.Add "7", "000111"
    table.Add "8", "001000"
    table.Add "9", "001001"
    table.Add ":", "001010"
    table.Add ";", "001011"
    table.Add "<", "001100"
    table.Add "=", "001101"
    table.Add ">", "001110"
    table.Add "?", "001111"
    table.Add "@", "010000"
    table.Add "A", "010001"
    table.Add "B", "010010"
    table.Add "C", "010011"
    table.Add "D", "010100"
    table.Add "E", "010101"
    table.Add "F", "010110"
    table.Add "G", "010111"
    table.Add "H", "011000"
    table.Add "I", "011001"
    table.Add "J", "011010"
    table.Add "K", "011011"
    table.Add "L", "011100"
    table.Add "M", "011101"
    table.Add "N", "011110"
    table.Add "O", "011111"
    table.Add "P", "100000"
    table.Add "Q", "100001"
    table.Add "R", "100010"
    table.Add "S", "100011"
    table.Add "T", "100100"
    table.Add "U", "100101"
    table.Add "V", "100110"
    table.Add "W", "100111"
    table.Add "`", "101000"
    table.Add "a", "101001"
    table.Add "b", "101010"
    table.Add "c", "101011"
    table.Add "d", "101100"
    table.Add "e", "101101"
    table.Add "f", "101110"
    table.Add "g", "101111"
    table.Add "h", "110000"
    table.Add "i", "110001"
    table.Add "j", "110010"
    table.Add "k", "110011"
    table.Add "l", "110100"
    table.Add "m", "110101"
    table.Add "n", "110110"
    table.Add "o", "110111"
    table.Add "p", "111000"
    table.Add "q", "111001"
    table.Add "r", "111010"
    table.Add "s", "111011"
    table.Add "t", "111100"
    table.Add "u", "111101"
    table.Add "v", "111110"
    table.Add "w", "111111"
    'Table_Ascii_6bit = table.Items                         'Get the items

    '  For i = 1 To 64
    '   Debug.Print i; "   "; table("'")
    ' Next

    '** trova il carattere ! o la stringa !AIVDM....
    For x = 1 To Len(InBuff)
        If Mid(InBuff, x, 6) = "!AIVDM" Then
            st = x                                         ' memorizza posizione
            Exit For
        End If
    Next

    fn = 0
    '** verifica se la fine della stringa è un Carriage Return e memorizza
    For x = st To Len(InBuff)
        If Mid(InBuff, x, 1) = vbCr Then
            fn = x
            Exit For
        End If
    Next
    If fn = 0 Then fn = Len(InBuff) + 1                    ' nel caso non trovi CR ma voglio andare avanti

    '** Separa la parte d'interesse cioè p.e. da !AIVDM
    strSentence = Mid(InBuff, st, fn - st)

    '** calcolo lunghezza stringa AIS e conteggio virgole
    For x = 1 To Len(strSentence)
        If Mid(strSentence, x, 1) = "," Then
            intcomma = intcomma + 1
            If intcomma = 5 Then STR_bit_init = x + 1
            If intcomma = 6 Then STR_bit_lenght = x - STR_bit_init
        End If
    Next

    ' carico la parte ITU 6bit per decodifica in formato VISUAL
    strSentence_6bit = Mid(strSentence, STR_bit_init, STR_bit_lenght)
    strSentence_6bit_lenght = Len(strSentence_6bit)

    '************************************* the JP_the_parser************************************************
    '*******************************************************************************************************



    JP_the_parser = String$(6 * strSentence_6bit_lenght, 48)    ' riempe tutto di ascii_48 ossia simbolo 0
    For i = 1 To strSentence_6bit_lenght
        ' get the character
        ch = Mid(strSentence_6bit, i, 1)
        ' add to JP_the_parser the converted bits
        avnz = i * 6 - 5
        Mid$(JP_the_parser, i * 6 - 5, 6) = table(ch)      ' sostituisce il bit zero con ipad di table
        ' Debug.Print i; ch; "   "; JP_the_parser
    Next
    'Debug.Print "DATA JP_the_parser"
    'Debug.Print strSentence_6bit; " = "; JP_the_parser

    ' Chop stream in segmenti
    bits = Array(6, 2, 30, 4, 8, 10, 1, 28, 27, 12, 9, 6, 4, 1, 1, 19, 168)    ' Ref page 35 Algo

    seg = Left(JP_the_parser, bits(0))                     ' leggo i bit del dato
    MSG_1.ID_MSG = seg
    List1.AddItem "MSG ID                               " & MSG_1.ID_MSG

    JP_the_parser = Mid(JP_the_parser, bits(0) + 1)        ' li rimuovo
    seg = Left(JP_the_parser, bits(1))                     ' leggo i bit del dato
    MSG_1.Repeat_indicator = seg
    List1.AddItem "Rep Ind                              " & MSG_1.Repeat_indicator

    JP_the_parser = Mid(JP_the_parser, bits(1) + 1)        ' li rimuovo
    seg = Left(JP_the_parser, bits(2))                     ' leggo i bit del dato
    MSG_1.MMSI = BinToDec(seg)
    List1.AddItem "MMSI                                 " & MSG_1.MMSI

    JP_the_parser = Mid(JP_the_parser, bits(2) + 1)        ' li rimuovo
    seg = Left(JP_the_parser, bits(3))                     ' leggo i bit del dato
    MSG_1.Navigation_Status = seg
    List1.AddItem "Nav Status                           " & MSG_1.Navigation_Status

    JP_the_parser = Mid(JP_the_parser, bits(3) + 1)        ' li rimuovo
    seg = Left(JP_the_parser, bits(4))                     ' leggo i bit del dato
    MSG_1.Rate_of_turn = seg
    List1.AddItem "Rate_of_turn                         " & MSG_1.Rate_of_turn

    JP_the_parser = Mid(JP_the_parser, bits(4) + 1)        ' li rimuovo
    seg = Left(JP_the_parser, bits(5))                     ' leggo i bit del dato
    MSG_1.Speed_Over_Ground = BinToDec(seg) / 10
    List1.AddItem "Speed_Over_Ground                    " & MSG_1.Speed_Over_Ground

    JP_the_parser = Mid(JP_the_parser, bits(5) + 1)        ' li rimuovo
    seg = Left(JP_the_parser, bits(6))                     ' leggo i bit del dato
    MSG_1.Position_accuracy = seg
    List1.AddItem "Position_accuracy                    " & MSG_1.Position_accuracy

    JP_the_parser = Mid(JP_the_parser, bits(6) + 1)        ' li rimuovo
    seg = Left(JP_the_parser, bits(7))                     ' leggo i bit del dato
    MSG_1.varLongitude = BinToDec(seg) / 10000 / 60        ' 1/10'000 minuti
    If MSG_1.varLongitude < 0 Then Parallelo$ = "O"
    List1.AddItem "Longitude                            " & MSG_1.varLongitude

    JP_the_parser = Mid(JP_the_parser, bits(7) + 1)        ' li rimuovo
    seg = Left(JP_the_parser, bits(8))                     ' leggo i bit del dato
    MSG_1.varLatitude = BinToDec(seg) / 10000 / 60
    If MSG_1.varLatitude < 0 Then Meridiano$ = "S"
    List1.AddItem "Latitude                             " & MSG_1.varLatitude

    JP_the_parser = Mid(JP_the_parser, bits(8) + 1)        ' li rimuovo
    seg = Left(JP_the_parser, bits(9))                     ' leggo i bit del dato
    MSG_1.Course_Over_Ground = BinToDec(seg) / 10
    List1.AddItem "Course_Over_Ground                   " & MSG_1.Course_Over_Ground

    JP_the_parser = Mid(JP_the_parser, bits(9) + 1)        ' li rimuovo
    seg = Left(JP_the_parser, bits(10))                    ' leggo i bit del dato
    MSG_1.True_Heading = BinToDec(seg)
    List1.AddItem "True_Heading                         " & MSG_1.True_Heading

    JP_the_parser = Mid(JP_the_parser, bits(10) + 1)       ' li rimuovo
    seg = Left(JP_the_parser, bits(11))                    ' leggo i bit del dato
    MSG_1.time_from_report = BinToDec(seg)
    List1.AddItem "time_from_report                     " & MSG_1.time_from_report
    List1.AddItem "***************************************"

End Sub

Private Sub Command2_Click()
List1.Clear
End Sub


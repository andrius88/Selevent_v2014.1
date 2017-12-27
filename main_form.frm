VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form main_form 
   Caption         =   "Selevent ver.2014.1"
   ClientHeight    =   5685
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   8505
   LinkTopic       =   "Form1"
   ScaleHeight     =   5685
   ScaleWidth      =   8505
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Output file type"
      Enabled         =   0   'False
      Height          =   855
      Left            =   240
      TabIndex        =   11
      Top             =   1440
      Width           =   3255
      Begin VB.OptionButton Option1 
         Caption         =   "Compact"
         Height          =   375
         Left            =   240
         TabIndex        =   13
         Top             =   360
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton Option2 
         Caption         =   "S-files"
         Height          =   375
         Left            =   1800
         TabIndex        =   12
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   240
      TabIndex        =   10
      Text            =   "14.0"
      Top             =   3360
      Width           =   855
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   2400
      TabIndex        =   9
      Text            =   "33.0"
      Top             =   3360
      Width           =   855
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   1320
      TabIndex        =   8
      Text            =   "50.0"
      Top             =   3960
      Width           =   855
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1320
      TabIndex        =   7
      Text            =   "60.0"
      Top             =   2760
      Width           =   855
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   186
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   240
      ScrollBars      =   1  'Horizontal
      TabIndex        =   2
      Top             =   5160
      Width           =   7935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Select+Convert"
      Enabled         =   0   'False
      Height          =   975
      Left            =   5160
      TabIndex        =   0
      Top             =   3000
      Width           =   2175
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   7680
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label9 
      Caption         =   "The last event processed :"
      Height          =   255
      Left            =   240
      TabIndex        =   16
      Top             =   4800
      Width           =   2175
   End
   Begin VB.Label Label2 
      Caption         =   "Events scaned:"
      Height          =   255
      Left            =   4080
      TabIndex        =   15
      Top             =   1560
      Width           =   2655
   End
   Begin VB.Label Label7 
      Caption         =   "Events found:"
      Height          =   255
      Left            =   4080
      TabIndex        =   14
      Top             =   2160
      Width           =   2655
   End
   Begin VB.Label Label6 
      Caption         =   "Long Max"
      Height          =   255
      Left            =   2520
      TabIndex        =   6
      Top             =   3120
      Width           =   855
   End
   Begin VB.Label Label5 
      Caption         =   "Long Min"
      Height          =   255
      Left            =   360
      TabIndex        =   5
      Top             =   3120
      Width           =   735
   End
   Begin VB.Label Label4 
      Caption         =   "Lat Max"
      Height          =   255
      Left            =   1440
      TabIndex        =   4
      Top             =   2520
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "Lat Min"
      Height          =   255
      Left            =   1440
      TabIndex        =   3
      Top             =   3720
      Width           =   735
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H000000FF&
      BorderColor     =   &H80000006&
      FillColor       =   &H0000C000&
      Height          =   375
      Left            =   6960
      Shape           =   4  'Rounded Rectangle
      Top             =   2040
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "The first html file :"
      Height          =   975
      Left            =   360
      TabIndex        =   1
      Top             =   240
      Width           =   6855
   End
   Begin VB.Menu menFile 
      Caption         =   "&File"
      Begin VB.Menu menOpen 
         Caption         =   "&Open"
      End
      Begin VB.Menu menClose 
         Caption         =   "&Close"
         Enabled         =   0   'False
      End
      Begin VB.Menu menExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu menHelp 
      Caption         =   "&Help"
      Begin VB.Menu menAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "main_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Sub Command1_Click()
 
 Label2.Caption = "Events scaned: "   'resets counter left from previous task
 Label7.Caption = "Events found: "    'resets counter of found events
 
 latmax = Val(Text2.Text)
 latmin = Val(Text3.Text)
 longmax = Val(Text4.Text)
 longmin = Val(Text5.Text)
 
 main_form.MousePointer = 11
 
 FailoVardas$ = Right$(CommonDialog1.FileName, 13)
 failokelias$ = Left$(CommonDialog1.FileName, Len(CommonDialog1.FileName) - 13)
 failodiena% = Val(Mid$(FailoVardas$, 6, 3))
 
 eilnr = 0    ' sveikas kint. pasako kelintas ivykis apdorojamas
 seinr = 0    ' sveikas kint. skaiciuoja kiek seisminiu ivykiu buvo surasta
 
 If Option1.Value = True Then
    naujasFailoVardas$ = "GBF" + Mid$(FailoVardas$, 4, 2) + "_sei.txt"
    Open (failokelias$ + naujasFailoVardas$) For Output As #2
 ElseIf Option2.Value = True Then
    naujasFailoVardas$ = "GBF" + Mid$(FailoVardas$, 4, 2) + "_sei_phase_data.txt"
    Open (failokelias$ + naujasFailoVardas$) For Output As #2
 End If
 
 
 '''' Open (failokelias$ + "GBF" + Mid$(FailoVardas$, 4, 2) + "_sei.txt") For Output As #2
 
 Do       ' ciklas skirtingiems failams skaityti
  
   On Error GoTo 20   ' apsisaugoma tuo atveju, jei direktorijoje nebus visu metu dienu failu
30 GoTo 40
20 failodiena% = failodiena% + 1
   If failodiena% > 366 Then
    Exit Do
   End If
   Resume 50
40 Rem' continue
 
50 curentfile$ = failokelias$ + Left$(FailoVardas$, 5) + Format$(failodiena%, "000") + Right$(FailoVardas, 5)
   Rem formuojamas sekantis failo vardas
 
   Open curentfile$ For Input As #1
  
   Label1.Caption = "The current (last) html file : " & vbCrLf & curentfile$ & vbCrLf & vbCrLf & _
                    "The output 'txt' file : " & vbCrLf & failokelias$ + naujasFailoVardas$
 
   eilute$ = ""  ' char kintamasis savyje laikantis nuskaityta eilute
   
   eilute$ = ReadLine$()   '''' eilute$ = ReadLine$(eilnr)
 
   Do Until EOF(1)
    
     If Len(eilute$) > 5 Then      ' if line is not empty
      
       Call IsGbfHeadLine(eilute$, IsHeadLine)  ' tikrina ar eilute yra GBF Head line ir is naujo tipo (<2014) eilutes padaro seno tipo
    
       If IsHeadLine = True Then    ' IsGbfHeadLine(eilute$) - funkcija, kuri tikrina ar eilute yra GBF head eilute
   
         eilnr = eilnr + 1
         
         latitude = Val(Mid$(eilute$, 25, 5))
         longnitude = Val(Mid$(eilute$, 33, 5))
     
           If latitude >= latmin And latitude <= latmax And longnitude >= longmin And longnitude <= longmax Then
       
             Shape1.BackStyle = 1
             seinr = seinr + 1
       
             yyyy% = Val(Mid$(eilute$, 3, 4))    ' nuskaito metus
             doy% = Val(Mid$(eilute$, 8, 3))     ' nuskaito doy
             
             Call doyToyyyymmdd(doy%, yyyy%, mm%, dd%)  ' is yyyy ir doy gaumani: yyyy, mm, dd
             
             Call makeNordicHeadLine(eilute$, yyyy%, mm%, dd%, seiline$)    ' konstruojama HEAD NORDIC eilute
       
             Print #2, seiline$     ' irasoma Nordic Head eilute
       
             Text1.Text = seiline$
       
             If Option2.Value = True And Not EOF(1) Then   ' jeigu kuriamas pilnas S-failas
         
                eilute$ = ReadLine$()      ' nuskaito/praleidzia tuscia eilute
                eilute$ = ReadLine$()      ' nuskaito/praleidzia anaotaciju eilute
                eilute$ = ReadLine$()      ' skaito pirma parametrine eilute
         
                phaseline$ = "                                                                               3"     ' raso komentaro eilute
                Print #2, phaseline$
                phaseline$ = " STAT SP IPHASW D HRMM SECON CODA AMPLIT PERI AZIMU VELO SNR AR TRES W  DIS CAZ7"     ' raso anotaciju eilute
                Print #2, phaseline$
         
                Do While Len(eilute$) > 5
                
                  Call makeNordicParametricLine(eilute$, phaseline$)    ' konstruojama nauja NORSAR parametrine eilute
                       
                  Print #2, phaseline$
         
                  eilute$ = ReadLine$()
         
                Loop
                  
              phaseline$ = ""       ' kintamasis istrinamas
              Print #2, phaseline$
                  
            End If
       
       Label7.Caption = "Events found: " + Format$(seinr, "00000000")
     
     End If

  End If
  
  End If
    
  Label2.Caption = "Events scaned: " + Format$(eilnr, "00000000")
 
  eilute$ = ReadLine$()
     
 Loop
 
 Close #1
 failodiena% = failodiena% + 1
 'Shape1.BackStyle = 0
 
 Loop While failodiena% <= 366
 
 Close #2
 main_form.MousePointer = 0
End Sub
Private Sub makeNordicHeadLine(eilute$, yyyy%, mm%, dd%, seiline$)

     Rem konstruojama HEAD NORDIC eilute
     seiline$ = " " + Mid$(eilute$, 3, 4)  'metai
     seiline$ = seiline$ + " " + Format$(mm%, "00") ' menuo
     seiline$ = seiline$ + Format$(dd%, "00")  'diena
     seiline$ = seiline$ + " " + Mid$(eilute$, 12, 2) 'val
     seiline$ = seiline$ + Mid$(eilute$, 15, 2) 'min
     seiline$ = seiline$ + " " + Mid$(eilute$, 18, 4) 'sek
     seiline$ = seiline$ + " L "
     seiline$ = seiline$ + " " + Mid$(eilute$, 25, 5) + "0" 'latitude
     seiline$ = seiline$ + " " + Mid$(eilute$, 33, 6) + "0" ' longnitude
     seiline$ = seiline$ + "       "  ' gylis nevertintas
     seiline$ = seiline$ + "GBF" ' raportavusi agentura
     seiline$ = seiline$ + Mid$(eilute$, 76, 3)  'stociu sk.
     seiline$ = seiline$ + Mid$(eilute$, 49, 4)  'rezidiulai
    
       If Mid$(eilute$, 82, 1) = " " Then
         seiline$ = seiline$ + "    "     ' kai magnitudes nera
       Else
         seiline$ = seiline$ + Mid$(eilute$, 81, 4) ' kai magnitude yra
       End If
    
    seiline$ = seiline$ + "N"    ' magnitudes tipas
    seiline$ = seiline$ + "GBF"  ' agentura
    seiline$ = seiline$ + "        " ' II magnitude paliekama tuscia
    seiline$ = seiline$ + "        1"  ' III magnitude paliekama tuscia ir NORDIC eilutes pabaiga

End Sub
Private Sub makeNordicParametricLine(eilute$, phaseline$)

    ' toliau konstruojama nauja NORSAR parametrine eilute
    phaseline$ = " " + Mid$(eilute$, 3, 4)
    phaseline$ = phaseline$ + " SZ"
    phaseline$ = phaseline$ + " " + Mid$(eilute$, 23, 3) + "  "
    phaseline$ = phaseline$ + " "  'weight
    phaseline$ = phaseline$ + "A"  ' automatic pick
    phaseline$ = phaseline$ + "  "  ' first motion
    phaseline$ = phaseline$ + Mid$(eilute$, 27, 2) ' hour
    phaseline$ = phaseline$ + Mid$(eilute$, 30, 2) ' min
    phaseline$ = phaseline$ + " "
    phaseline$ = phaseline$ + Mid$(eilute$, 33, 4) ' sec
    phaseline$ = phaseline$ + " "  ' free
    phaseline$ = phaseline$ + "     "  ' coda
    phaseline$ = phaseline$ + " " + Mid$(eilute$, 78, 6) ' amplitude

    period = 1 / Val(Mid$(eilute$, 86, 5))
    phaseline$ = phaseline$ + " " + Format$(period, "0.00")
    phaseline$ = phaseline$ + " "
    phaseline$ = phaseline$ + Mid$(eilute$, 46, 5)  ' azimut
    phaseline$ = phaseline$ + "     "   ' velocity
    phaseline$ = phaseline$ + Mid$(eilute$, 70, 4) ' siganal to noise
    Ares% = Abs(Val(Mid$(eilute$, 54, 4)))
    phaseline$ = phaseline$ + " "
    phaseline$ = phaseline$ + Format$(Ares%, "00")  ' azimut res
    phaseline$ = phaseline$ + Mid$(eilute$, 39, 5)
    phaseline$ = phaseline$ + "  "  ' actual weight
    phaseline$ = phaseline$ + Mid$(eilute$, 7, 5) ' distance"
    phaseline$ = phaseline$ + " "
    phaseline$ = phaseline$ + "    "

End Sub

Private Sub menAbout_Click()
 Load Info_form
 info_string = "This application converts NORSAR GBF seismic catalogues (http://www.norsardata.no/NDC/bulletins/) " & _
               "to NORDIC (SEISAN) format." & vbCrLf & vbCrLf & _
               "The NORSAR GBF catalogue in form of html files has to be downloaded and saved in some local directory " & _
               "before the application can be run." & vbCrLf & vbCrLf & _
               "Open menu 'File' and select the first html file you want to start from. Then click button 'Select+Convert' " & _
               "and the application will read all files in that directory one by one, convert to NORSAR format and " & _
               "write to one 'txt' file." & vbCrLf & vbCrLf & _
               "Seismic events will be transfered to the NORDIC catalogue just in case the coordinates of the event get " & _
               "inside the area of interest defined by four coordinates presented in the main window." & vbCrLf & vbCrLf & _
               "The application 'Selevent' was written by A.Pacea for his personal needs back in distant 2000 and slightly " & _
               "modified in 2014 then NORSAR catalogue format had been changed."
               
 Info_form.Label_info.Font.Size = 10
 Info_form.Label_info.Caption = info_string
 Info_form.Show
 
End Sub

Private Sub menClose_Click()
  CommonDialog1.FileName = ""
  Label1.Caption = "The first html file : " & vbCrLf & CommonDialog1.FileName
  Command1.Enabled = False
  menClose.Enabled = False
  Frame1.Enabled = False
  Shape1.BackStyle = 0
End Sub

Private Sub menExit_Click()
 End
End Sub

Private Sub menOpen_Click()
 Shape1.BackStyle = 0
 CommonDialog1.MaxFileSize = 1024
 CommonDialog1.ShowOpen
 If CommonDialog1.FileName <> "" Then
    main_form.MousePointer = 11
    Label1.Caption = "The first html file : " & vbCrLf & CommonDialog1.FileName
    menClose.Enabled = True
    Command1.Enabled = True
    Frame1.Enabled = True
    
 Else: menClose.Enabled = False
    Command1.Enabled = False
    Frame1.Enabled = False
 End If
 
 main_form.MousePointer = 0
 
End Sub

Sub doyToyyyymmdd(doy%, yyyy%, mm%, dd%)
 
 jan% = 31
 feb% = 28
 mar% = 31
 apr% = 30
 may% = 31
 jun% = 30
 jul% = 31
 aug% = 31
 sep% = 30
 Oct% = 31
 nov% = 30
 dec% = 31
 
 If yyyy% = 1992 Or yyyy% = 1996 Or yyyy% = 2000 Or _
    yyyy% = 2004 Or yyyy% = 2008 Or yyyy% = 2012 Or _
    yyyy% = 2016 Or yyyy% = 2020 Or yyyy% = 2024 Or _
    yyyy% = 2024 Or yyyy% = 2028 Or yyyy% = 2032 Then
     feb% = 29
 Else
     feb% = 28
 End If
 
 If (doy% - jan%) <= 0 Then
  mm% = 1
  dd% = doy%
 ElseIf (doy% - jan% - feb%) <= 0 Then
  mm% = 2
  dd% = doy% - jan%
 ElseIf (doy% - jan% - feb% - mar%) <= 0 Then
  mm% = 3
  dd% = doy% - jan% - feb%
 ElseIf (doy% - jan% - feb% - mar% - apr%) <= 0 Then
  mm% = 4
  dd% = doy% - jan% - feb% - mar%
 ElseIf (doy% - jan% - feb% - mar% - apr% _
         - may%) <= 0 Then
  mm% = 5
  dd% = doy% - jan% - feb% - mar% - apr%
 ElseIf (doy% - jan% - feb% - mar% - apr% _
       - may% - jun%) <= 0 Then
  mm% = 6
  dd% = doy% - jan% - feb% - mar% - apr% - may%
 ElseIf (doy% - jan% - feb% - mar% - apr% _
       - may% - jun% - jul%) <= 0 Then
  mm% = 7
  dd% = doy% - jan% - feb% - mar% - apr% - may% - jun%
 ElseIf (doy% - jan% - feb% - mar% - apr% _
              - may% - jun% - jul% - aug%) <= 0 Then
  mm% = 8
  dd% = doy% - jan% - feb% - mar% - apr% - may% - jun% _
             - jul%
 ElseIf (doy% - jan% - feb% - mar% - apr% _
              - may% - jun% - jul% - aug% _
              - sep%) <= 0 Then
  mm% = 9
  dd% = doy% - jan% - feb% - mar% - apr% - may% - jun% _
             - jul% - aug%
 ElseIf (doy% - jan% - feb% - mar% - apr% _
              - may% - jun% - jul% - aug% _
              - sep% - Oct%) <= 0 Then
  mm% = 10
  dd% = doy% - jan% - feb% - mar% - apr% - may% - jun% _
             - jul% - aug% - sep%
 ElseIf (doy% - jan% - feb% - mar% - apr% _
              - may% - jun% - jul% - aug% _
              - sep% - Oct% - nov%) <= 0 Then
  mm% = 11
  dd% = doy% - jan% - feb% - mar% - apr% - may% - jun% _
             - jul% - aug% - sep% - Oct%
 ElseIf (doy% - jan% - feb% - mar% - apr% _
              - may% - jun% - jul% - aug% _
              - sep% - Oct% - nov% - dec%) <= 0 Then
  mm% = 12
  dd% = doy% - jan% - feb% - mar% - apr% - may% - jun% _
             - jul% - aug% - sep% - Oct% - nov%
 Else
  Text1.Text = "Problemos su menesiu"
 End If
 
End Sub

Private Function ReadLine$()   ' funkcija vienai tekstinio failo eilutei perskaityti
 
 PartialLine$ = ""
 temp$ = ""
 
 Do
   PartialLine$ = PartialLine$ + temp$
   temp$ = Input(1, #1)
 Loop While temp$ <> Chr$(10) And Not EOF(1)
 
 ReadLine$ = PartialLine$
 
End Function
 
Private Sub IsGbfHeadLine(GbfHeadLine$, IsHeadLine)
Rem tikrina ar yra GBF Head eilute ir jeigu naujo tipo, tai konvertuoja i seno tipo

 If Mid$(GbfHeadLine$, 3, 3) = "199" Or Mid$(GbfHeadLine$, 3, 3) = "200" Then

    IsHeadLine = True
    GbfHeadLine$ = GbfHeadLine$ ' seno stiliaus head eile
 
 ElseIf Mid$(GbfHeadLine$, 1, 8) = "<A NAME=" And Mid$(GbfHeadLine$, 33, 2) = "20" Then
        
    IsHeadLine = True
    GbfHeadLine$ = "  " + Mid$(GbfHeadLine$, 33, Len(GbfHeadLine$) - 33) ' nupjaunama pradzia <A NAME=...
    
 Else
 
    IsHeadLine = False
 
 End If
 
End Sub
   

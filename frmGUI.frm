VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmGUI 
   Caption         =   "Einstellungen speichern"
   ClientHeight    =   2625
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   2640
   OleObjectBlob   =   "frmGUI.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "frmGUI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Beschreibung
'------------
    'Diese Demo zeigt, wie der Zustand von Steuerelementen eines UserForms
    '(hier am Beispiel von Kontrollkästchen und Optionsfeldern, also ob diese
    'jeweils angehakt sind oder nicht) beim Schließen in eine Textdatei geschrieben
    'und - sofern die Datei vorhanden ist - beim Öffnen des UserForms wieder
    'ausgelesen und im UserForm gesetzt werden.
    'Dazu werden die beiden UserForm-Ereignisse "Initialize" und "Terminate" genutzt.
    '
    'Für ein Debugging kann am Anfang der beiden Prozeduren jeweils ein Breakpoint
    'gesetzt und der Code im Einzelschrittmodus (F8) ausgeführt werden.
    'Das erstmalige Anlegen der Textdatei (also beim ersten Schließen des UserForms)
    'kann in diesem Beispiel im Windows-Ordner "C:\Users\USERNAME\AppData\Roaming\Microsoft\AddIns"
    'mitverfolgt werden. Jedes weitere Schließen überschreibt die Datei mit den aktuellen
    'Status der Steuerelemente ohne Rückfrage.
    '
    'Die Textdatei mit den Einstellungen (hier "demo.settings") kann mit einem
    'einfachen Texteditor (z.B. Windows Editor oder Notepad++) geöffnet und
    'angesehen werden.

'Code
'----
'Variablen für die Werte der Tool-Einstellungen als String definieren,
'weil diese aus einer einfachen Textdatei ausgelesen werden und erst dann
'in den richtigen Datentyp für Excel konvertiert werden.
Dim option1 As String, option2 As String, option3 As String, option4 As String

'Die "Terminate"-Prozedur wird automatisch beim Schließen des UserForms gestartet
Private Sub UserForm_Terminate()
    
    'Ausgangskanal öffnen (hier #1)
    Open Application.UserLibraryPath & "demo.settings" For Output As #1
        'Werte der Steuerelemente in der GUI jeweils in einen Integer-Wert konvertieren
        '(TRUE wird -1 und FALSE wird 0) und zeilenweise in die Textdatei schreiben
        Print #1, CInt(CheckBox1.Value) 'Kontrollkästchen 1
        Print #1, CInt(CheckBox2.Value) 'Kontrollkästchen 2
        Print #1, CInt(OptionButton1.Value) 'Optionsfeld 1
        Print #1, CInt(OptionButton2.Value) 'Optionsfeld 2
    'Ausgangskanal schließen
    Close #1
    
End Sub

'Die "Initialize"-Prozedur wird automatisch beim Aufruf des UserForms gestartet
Private Sub UserForm_Initialize()
    'Gespeicherte Tool-Einstellungen laden sofern die Datei vorhanden ist,
    'in diesem Beispiel eine einfach Text-Datei mit dem Namen "demo.settings".
    'Hier wird im Standardordner für Office-Add-Ins gesucht, der
    'unter Windows unter C:\Users\USERNAME\AppData\Roaming\Microsoft\AddIns
    'zu finden ist und z.B. durch Eingabe von %APPDATA%\Microsoft\AddIns im
    'Windows Explorer geöffnet werden kann.
    
    If Dir(Application.UserLibraryPath & "demo.settings") <> "" Then 'nur wenn es die Datei gibt
        
        'Eingangskanal öffnen (hier #1)
        Open Application.UserLibraryPath & "demo.settings" For Input As #1
            Line Input #1, option1 'Zeile 1 der Textdatei auslesen und in Variable "option1" speichern
            Line Input #1, option2 'Zeile 2 der Textdatei auslesen und in Variable "option2" speichern
            Line Input #1, option3 'Zeile 3 der Textdatei auslesen und in Variable "option3" speichern
            Line Input #1, option4 'Zeile 4 der Textdatei auslesen und in Variable "option4" speichern
        'Eingangskanal schließen
        Close #1
        
        'Aus der Textdatei ausgelesene Werte in boolesche Werte konvertieren
        'und den Steuerelementen in der GUI zuweisen (-1 wird TRUE, 0 wird FALSE)
        CheckBox1.Value = CBool(option1)
        CheckBox2.Value = CBool(option2)
        OptionButton1.Value = CBool(option3)
        OptionButton2.Value = CBool(option4)
    
    End If
    
End Sub

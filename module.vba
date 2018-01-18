Declare PtrSafe Function SetTimer Lib "user32" (ByVal hwnd As LongPtr, ByVal nIDEvent As LongPtr, ByVal uElapse As LongPtr, ByVal lpTimerfunc As LongPtr) As LongPtr

Declare PtrSafe Function KillTimer Lib "user32" (ByVal hwnd As LongPtr, ByVal nIDEvent As LongPtr) As LongPtr

 

Public TimerID As LongPtr 'Need a timer ID to eventually turn off the timer. If the timer ID <> 0 then the timer is running

 

Public Sub TriggerTimer(ByVal hwnd As Long, ByVal uMsg As Long, ByVal idevent As Long, ByVal Systime As Long)

Call SendMail

End Sub

 

Public Sub DeactivateTimer()

Dim lSuccess As LongPtr              '<~ Corrected here

  lSuccess = KillTimer(0, TimerID)

  If lSuccess = 0 Then

    MsgBox "The timer failed to deactivate."

  Else

    TimerID = 0

  End If

End Sub

 

Public Sub ActivateTimer(ByVal nMinutes As Long)

  nMinutes = nMinutes * 1000 * 60 'The SetTimer call accepts milliseconds, so convert to minutes

  If TimerID <> 0 Then Call DeactivateTimer 'Check to see if timer is running before call to SetTimer

  TimerID = SetTimer(0, 0, nMinutes, AddressOf TriggerTimer)

  If TimerID = 0 Then

    MsgBox "The timer failed to activate."

  End If

End Sub

 

 

Public Sub Start()

Call ActivateTimer(1)

 

End Sub

 

 

Public Sub CreateNewMessage()

Dim objMsg As MailItem

 

Set objMsg = Application.CreateItem(olMailItem)

 

With objMsg

  .To = "Alias@domain.com"

  .CC = "Alias2@domain.com"

  .BCC = "Alias3@domain.com"

  .Subject = "This is the subject"

  .Categories = "Test"

  .VotingOptions = "Yes;No;Maybe;"

  .BodyFormat = olFormatPlain ' send plain text message

  .Importance = olImportanceHigh

  .Sensitivity = olConfidential

  .Attachments.Add ("path-to-file.docx")

 

' Calculate a date using DateAdd or enter an explicit date

  .ExpiryTime = DateAdd("m", 6, Now) '6 months from now

  .DeferredDeliveryTime = #8/1/2012 6:00:00 PM#

 

  .Display

End With

 

Set objMsg = Nothing

End Sub

 

 

 

Option Explicit

Sub SendMail()

    

    Dim olApp As Outlook.Application

    Dim olMail As Outlook.MailItem

    Dim blRunning As Boolean

    

     'get application

    blRunning = True

    On Error Resume Next

    Set olApp = GetObject(, "Outlook.Application")

    If olApp Is Nothing Then

       Set olApp = New Outlook.Application

        blRunning = False

    End If

    On Error GoTo 0

    

    Set olMail = olApp.CreateItem(olMailItem)

    With olMail

         'Specify the email subject

        .Subject = "My email with attachment"

         'Specify who it should be sent to

         'Repeat this line to add further recipients

        .Recipients.Add "angelo.trozzo@gmail.com"

         'specify the file to attach

         'repeat this line to add further attachments

        .Attachments.Add "C:\Users\trozza\mycalndar.xlsx"

         'specify the text to appear in the email

        .Body = "Here is an email"

         'Choose which of the following 2 lines to have commented out

         .Send ' This will send the message straight away

    End With

    

    If Not blRunning Then olApp.Quit

    

    Set olApp = Nothing

    Set olMail = Nothing

    

End Sub


Sub ExportAppointmentsToExcel()

    Const SCRIPT_NAME = "Export Appointments to Excel"

    Dim olkFld As Object, _

        olkLst As Object, _

        olkApt As Object, _

        excApp As Object, _

        excWkb As Object, _

        excWks As Object, _

        lngRow As Long, _

        intCnt As Integer

    Set olkFld = Application.ActiveExplorer.CurrentFolder

    If olkFld.DefaultItemType = olAppointmentItem Then

        strFilename = InputBox("Enter a filename (including path) to save the exported appointments to.", SCRIPT_NAME)

        If strFilename <> "" Then

            Set excApp = CreateObject("Excel.Application")

            Set excWkb = excApp.Workbooks.Add()

            Set excWks = excWkb.Worksheets(1)

            'Write Excel Column Headers

            With excWks

                .Cells(1, 1) = "Organizer"

                .Cells(1, 2) = "Created"

                .Cells(1, 3) = "Subject"

                .Cells(1, 4) = "Start"

                .Cells(1, 5) = "Required"

                .Cells(1, 6) = "Optional"

            End With

            lngRow = 2

            Set olkLst = olkFld.Items

            olkLst.Sort "[Start]"

            olkLst.IncludeRecurrences = True

            'Write appointments to spreadsheet

            For Each olkApt In Application.ActiveExplorer.CurrentFolder.Items

                'Only export appointments

                If olkApt.Class = olAppointment Then

                    'Add a row for each field in the message you want to export

                    excWks.Cells(lngRow, 1) = olkApt.Organizer

                    excWks.Cells(lngRow, 2) = olkApt.CreationTime

                    excWks.Cells(lngRow, 3) = olkApt.Subject

                    excWks.Cells(lngRow, 4) = olkApt.Start

                    excWks.Cells(lngRow, 5) = olkApt.RequiredAttendees

                    excWks.Cells(lngRow, 6) = olkApt.OptionalAttendees

                    lngRow = lngRow + 1

                    intCnt = intCnt + 1

                End If

            Next

           excWks.Columns("A:F").AutoFit

            excWkb.SaveAs strFilename

            excWkb.Close

            MsgBox "Process complete.  A total of " & intCnt & " appointments were exported.", vbInformation + vbOKOnly, SCRIPT_NAME

        End If

    Else

        MsgBox "Operation cancelled.  The selected folder is not a calendar.  You must select a calendar for this macro to work.", vbCritical + vbOKOnly, SCRIPT_NAME

    End If

    Set excWks = Nothing

    Set excWkb = Nothing

    Set excApp = Nothing

    Set olkApt = Nothing

    Set olkLst = Nothing

    Set olkFld = Nothing

End Sub
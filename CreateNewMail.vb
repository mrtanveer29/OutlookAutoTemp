Sub CreateNewMail()
    Dim obApp As Object
    Dim NewMail As MailItem
    Dim HTMLBody As String
    Dim style As String
    Dim endhour As Integer
    Dim endminute As Integer
    Dim timediff As Integer
     Dim phour As Integer
      Dim pmin As Integer
       Dim currTime As String
       Dim signature As String
    
 
    Set obApp = Outlook.Application
    Set NewMail = obApp.CreateItem(0)
    ' Set Body = Application.CreateItemFromTemplate("C:\Users\Mahmuda Hasan\AppData\Roaming\Microsoft\Templates\Work log.oft")
    currTime = Format(Now(), "HH:mm")
  
    endhour = Hour(currTime)
    endminute = Minute(currTime)
   
    timediff = (endhour * 60 + endminute) - 510
    phour = timediff / 60
    pmin = timediff Mod 60
    
    
    
    style = "<style>table{font-family:arial,sans-serif;border-collapse:collapse;width:100%;}td,th{border:1px solid #dddddd;text-align:left;padding:8px;}tr:nth-child(even){background-color:#dddddd;}</style>"
    HTMLBody = style & "<table>"
    HTMLBody = HTMLBody & "   <tbody>  "
    HTMLBody = HTMLBody & "     <tr>  "
    HTMLBody = HTMLBody & "       <td>Start Time</td>  "
    HTMLBody = HTMLBody & "       <td>8:30 AM</td>  "
    HTMLBody = HTMLBody & "     </tr>  "
    HTMLBody = HTMLBody & "     <tr>  "
    HTMLBody = HTMLBody & "       <td>End Time</td>  "
    HTMLBody = HTMLBody & "       <td>" & Format(Now(), "HH:mm AM/PM") & "</td>  "
    HTMLBody = HTMLBody & "     </tr>  "
    HTMLBody = HTMLBody & "     <tr>  "
    HTMLBody = HTMLBody & "       <td>Various Breaks</td>  "
    HTMLBody = HTMLBody & "       <td></td>  "
    HTMLBody = HTMLBody & "     </tr>  "
    HTMLBody = HTMLBody & "     <tr>  "
    HTMLBody = HTMLBody & "       <td>Productive Working Hours</td>  "
    HTMLBody = HTMLBody & "       <td>" & phour & ":" & pmin & " h </td>  "
    HTMLBody = HTMLBody & "     </tr>  "
    HTMLBody = HTMLBody & "     <tr>  "
    HTMLBody = HTMLBody & "  <td>Internet Condition for Today</td>  "
    HTMLBody = HTMLBody & "       <td>Good</td>  "
    HTMLBody = HTMLBody & "     </tr>  "
    HTMLBody = HTMLBody & "     <tr>  "
    HTMLBody = HTMLBody & "       <td>Work Interrupted for Power Failure</td>  "
    HTMLBody = HTMLBody & "       <td>None</td>  "
    HTMLBody = HTMLBody & "     </tr>  "
    HTMLBody = HTMLBody & "     <tr>  "
    HTMLBody = HTMLBody & "       <td>Overall Work Atmosphere at Home</td>  "
    HTMLBody = HTMLBody & "       <td>Good</td>  "
    HTMLBody = HTMLBody & "     </tr>  "
    HTMLBody = HTMLBody & "     <tr>  "
    HTMLBody = HTMLBody & "       <td>Additional Comment</td>  "
    HTMLBody = HTMLBody & "       <td>None</td>  "
    HTMLBody = HTMLBody & "     </tr>  "
    HTMLBody = HTMLBody & "     <tr>  "
    HTMLBody = HTMLBody & "       <td>Comment from Office</td>  "
    HTMLBody = HTMLBody & "       <td></td>  "
    HTMLBody = HTMLBody & "     </tr>  "
    HTMLBody = HTMLBody & "     </tbody>  "
    HTMLBody = HTMLBody & "     </table>  <br/> "
     HTMLBody = HTMLBody & "    <h4>Work Log:</h4> "
     HTMLBody = HTMLBody & "    <ol><li></li></ol> "
    'You can change the concrete info as per your needs
    With NewMail
        .Display
    End With

signature = NewMail.HTMLBody
    With NewMail
         .Subject = "Work From Home - [Name: Tanveer Hasan]  - " & Date & " (Work - Full) "
         .To = "workfromhome@bitmascot.com"
         .CC = "ghosh@bitmascot.com"
         .BCC = "tasnia@webalive.com.au"
         .HTMLBody = HTMLBody & signature
         
         .Importance = olImportanceNormal
         
         .Display
    End With
 
    Set obApp = Nothing
    Set NewMail = Nothing
End Sub


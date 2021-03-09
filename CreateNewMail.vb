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
       Dim currTime() As String
       Dim exitTime As String
       Dim signature As String
    Dim entry As String
    Dim break As String
      Dim breakhour As Integer
      Dim breakmin As Integer
    Dim LArray() As String
    
    entry = InputBox("Give me Entry time")
    break = InputBox("Break time in minute")
    LArray = Split(entry, ":")
    
    entryminute = LArray(0) * 60 + LArray(1)
    
    Set obApp = Outlook.Application
    Set NewMail = obApp.CreateItem(0)
    ' Set Body = Application.CreateItemFromTemplate("C:\Users\Mahmuda Hasan\AppData\Roaming\Microsoft\Templates\Work log.oft")
    exitTime = InputBox("Give me Exit time")
    currTime = Split(exitTime, ":")
  
  breakhour = Int((break / 60))
    breakmin = break - (breakhour * 60)
    
    endhour = currTime(0) + 12
    endminute = currTime(1)
   
    timediff = (endhour * 60 + endminute) - (entryminute + break)
    phour = Int(timediff / 60)
    pmin = timediff - (phour * 60)
    
    
    
    style = "<style>table{font-family:arial,sans-serif;border-collapse:collapse;width:100%;}td,th{border:1px solid #dddddd;text-align:left;padding:8px;}tr:nth-child(even){background-color:#dddddd;}</style>"
     
     HTMLBody = "    <h3 style='color:#DAA520'>Tanveer Hasan's Work Log for " & Format(Now(), "dd/MM/yyyy") & ":</h3> "
     
      HTMLBody = HTMLBody & style & "<table>"
    HTMLBody = HTMLBody & "   <tbody>  "
    HTMLBody = HTMLBody & "     <tr>  "
   
    HTMLBody = HTMLBody & "       <td> Task</td>  "
     HTMLBody = HTMLBody & "       <td>Description</td>  "
      HTMLBody = HTMLBody & "       <td>Time</td>  "
    HTMLBody = HTMLBody & "     </tr>  "
    HTMLBody = HTMLBody & "     <tr>  "

    HTMLBody = HTMLBody & "       <td> Meeting</td>  "
     HTMLBody = HTMLBody & "       <td> </td>  "
      HTMLBody = HTMLBody & "       <td></td>  "
    HTMLBody = HTMLBody & "     </tr>  "
        HTMLBody = HTMLBody & "     <tr>  "
   
    HTMLBody = HTMLBody & "       <td> Estimation</td>  "
     HTMLBody = HTMLBody & "       <td> </td>  "
      HTMLBody = HTMLBody & "       <td></td>  "
    HTMLBody = HTMLBody & "     </tr>  "
        HTMLBody = HTMLBody & "     <tr>  "
   
    HTMLBody = HTMLBody & "       <td> R&D</td>  "
     HTMLBody = HTMLBody & "       <td> <table><tr><td style='width:80%'>.</td><td></td></tr></table></td>  "
      HTMLBody = HTMLBody & "       <td></td>  "
    HTMLBody = HTMLBody & "     </tr>  "
        HTMLBody = HTMLBody & "     <tr>  "
   
    HTMLBody = HTMLBody & "       <td> Requirement Analysis</td>  "
     HTMLBody = HTMLBody & "       <td> </td>  "
      HTMLBody = HTMLBody & "       <td></td>  "
    HTMLBody = HTMLBody & "     </tr>  "
        HTMLBody = HTMLBody & "     <tr>  "
   
    HTMLBody = HTMLBody & "       <td> Code Review</td>  "
     HTMLBody = HTMLBody & "       <td> </td>  "
      HTMLBody = HTMLBody & "       <td></td>  "
    HTMLBody = HTMLBody & "     </tr>  "
        HTMLBody = HTMLBody & "     <tr>  "
   
    HTMLBody = HTMLBody & "       <td> Issue</td>  "
     HTMLBody = HTMLBody & "       <td> </td>  "
      HTMLBody = HTMLBody & "       <td></td>  "
    HTMLBody = HTMLBody & "     </tr>  "
            HTMLBody = HTMLBody & "     <tr>  "

    HTMLBody = HTMLBody & "       <td> Feature</td>  "
     HTMLBody = HTMLBody & "       <td> </td>  "
      HTMLBody = HTMLBody & "       <td></td>  "
    HTMLBody = HTMLBody & "     </tr>  "
    HTMLBody = HTMLBody & "     </tbody>  "
    HTMLBody = HTMLBody & "     </table> "
        HTMLBody = HTMLBody & "   <h3 style='color:#DAA520'>Planning</h3>  "
     HTMLBody = HTMLBody & style & "<table>"
    HTMLBody = HTMLBody & "   <tbody>  "
    HTMLBody = HTMLBody & "     <tr>  "
    HTMLBody = HTMLBody & "       <td>Target</td>  "
    HTMLBody = HTMLBody & "       <td> Due on</td>  "
     HTMLBody = HTMLBody & "       <td>Status</td>  "
    HTMLBody = HTMLBody & "     </tr>  "
        HTMLBody = HTMLBody & "     </tbody>  "
    HTMLBody = HTMLBody & "     </table> "
    
    HTMLBody = HTMLBody & "   <h3 style='color:#DAA520'>Summary</h3>  "
    HTMLBody = HTMLBody & style & "<table>"
    HTMLBody = HTMLBody & "   <tbody>  "
    HTMLBody = HTMLBody & "     <tr>  "
    HTMLBody = HTMLBody & "       <td>Start Time</td>  "
    HTMLBody = HTMLBody & "       <td>" & entry & " AM</td>  "
    HTMLBody = HTMLBody & "     </tr>  "
    HTMLBody = HTMLBody & "     <tr>  "
    HTMLBody = HTMLBody & "       <td>End Time</td>  "
    HTMLBody = HTMLBody & "       <td>" & exitTime & " PM</td>  "
    HTMLBody = HTMLBody & "     </tr>  "
    HTMLBody = HTMLBody & "     <tr>  "
    HTMLBody = HTMLBody & "       <td>Various Breaks</td>  "
    HTMLBody = HTMLBody & "       <td>" & breakhour & " hour(s) " & breakmin & " minute(s) </td>  "
    HTMLBody = HTMLBody & "     </tr>  "
    HTMLBody = HTMLBody & "     <tr>  "
    HTMLBody = HTMLBody & "       <td>Productive Working Hours</td>  "
    HTMLBody = HTMLBody & "       <td>" & phour & " hour(s) " & pmin & " minute(s)  </td>  "
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
    HTMLBody = HTMLBody & "     </table> "
    'You can change the concrete info as per your needs
    With NewMail
        .Display
    End With


signature = NewMail.HTMLBody
    With NewMail
         .Subject = "Work From Home - [HPDS2 (Chamera & BP)] -[Name: Tanveer Hasan]  - " & Format(Now(), "dd/MM/yyyy") & " (Work - Full) "
         .To = "anick@bitmascot.com;"
         .CC = "hpds2-team@bitmascot.com;liakat@bitmascot.com"
         .BCC = ""
         .HTMLBody = HTMLBody & signature
         
         .Importance = olImportanceNormal
         
         .Display
    End With
 
    Set obApp = Nothing
    Set NewMail = Nothing
End Sub


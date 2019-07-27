' 3 h�napn�l r�gebbi Be�rkez�s �s T�voz�s bejegyz�seket t�rl� script
' Ferenczy Andr�s - 2005.01.12
' Cs�ti Attila	- 2006.11.13

    Dim myOlApp
    Dim myAppt
    Dim myNS
    Dim myAppts
    Dim strTheDay
    Dim strToday
    Dim strMsg

    dDate = cDate( Month(Date()) & "/13/" & Year(Date()) )

' Run only one day monthly (13th)
    if cDate(dDate) = Date() then 

' Calculate the date until events will be deleted (3 months before)
    strTheDay = DateAdd("m",-3,Date())
    strToday = "[Start] < '" & strTheDay & "'"
    Set myOlApp = CreateObject("Outlook.Application")
    Set myNS = myOlApp.GetNamespace("MAPI")
    Set myAppts = myNS.GetDefaultFolder(9).Items
' Sort the collection (required by IncludeRecurrences).
    myAppts.Sort "[Start]"
' Make sure recurring appointments are included.
    myAppts.IncludeRecurrences = True
' Filter the collection to include only the day's appointments.
    Set myAppts = myAppts.Restrict(strToday)
' Sort it again to put recurring appointments in correct order.
    myAppts.Sort "[Start]"
' Loop through collection and get "Be�rkez�s" and "T�voz�s" events.
' Remove "Be�rkezett." and "Elt�vozott." items before 3 months
    Set myAppt = myAppts.Find("[Subject] = """ & "Be�rkezett." & """ or [Subject] = """ & "T�vozott." & """ ")
'Remove all items
     'Set myAppt = myAppts.Find("[Subject] <> """ & "Be�rkezett." & """ ")
	Do While TypeName(myAppt) <> "Nothing"
' If this is a recurring item, it was created by user -> do not delete, find next
        if myAppt.RecurrenceState <> 0 then
            Set myAppt = myAppts.FindNext
	  else
'        strMsg = strMsg & vbLf & myAppt.Subject
'        strMsg = strMsg & " at " & FormatDateTime(myAppt.Start, vbshortdate)
         myAppt.delete
       Set myAppt = myAppts.FindNext
	  end if
    Loop
' Display the information.
'    MsgBox "Ezeket a bejegyz�seket t�r�ln� a script:" & vbLf & strMsg

     End if
 
    Set myOlApp = Nothing
    Set myAppt = Nothing
    Set myNS = Nothing
    Set myAppts = Nothing

    

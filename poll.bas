Attribute VB_Name = "poll"
'**************************************************************************
    'THE CGI32.BAS FILE IS NOT WRITTEN BY THE AUTHORS OF POLL.EXE
'Please read the introduction to cgi32.bas in the file to understand the
' concepts behind CGI.
' this program uses only one of the many fuctions that can be easily written
' in WinCGI.
'
'Poll.exe ie the WinCGI file generates the results of a Poll.
' many a times you might have seen polls on websites and liked it. So, theres the
'source code for it
'
'You need to run this from the HTML page itself..which is enclosed in the zip file Poll.zip
' Note the follwing files should be present in the poll.zip
'   1. poll.html
'   2. poll.bas
'   3. cgi32.bas
'   4. readme.txt
'   5. poll.exe
'   6. some extra VB files...
'
' In WinCGI, writing a CGI script is very easy. All you have to do is to include
'the cgi32.bas file with your project. then write another module which
'will contain your source code. the poll.bas shall clear most of the
' doubts regarding the basics of WinCGI
' Note to run WinCGI you need a server that runs on Win9x or NT, IIS4
' You cannot run these scripts on an UNIX server like tripod.com ( where I have
' my unofficial home page and my scripts donot run. I have to abandon the site, and build a new one!!)
'
'In case you want to test the program you can use a free server like Sambar server( that I use).
' To run and install sambar (www.sambar.com) please read the help file provided along the program).
' Incase of any doubts...or if your programs are not working fine then
'contact me at krishna_rath@hotmail.com. or krath@rediffmail.com....Incase my hotmail account is hacked and you donot recieve an reply by a week.
' This script can run on any win server ....but I tested it on Sambar cause it is free!!
'
'************************************************************************
'
' If you have used WinCGI then skip this part
' How to run this file
' First, if you make any changes to the file then complie it to an EXE file.
' 2. place the HTML page( poll.html) on your server folder ( if it is sambar server, then place it in the "path\sambar\docs" folder, where path is the "c:\" or "d:\"
' 3. place the Poll.exe and the Poll.txt file in "path\sambar\cgi-win\" folder, if cgi-bin does not exists then create it
' 4. If your server is not running then start it.
' 5. Open your browser, type in 127.0.0.1\poll.html. The poll.html file should appear. If it does not recheck your configuration
' 6. Click on the Poll button and see the results!!! If you donot see anything or see a "page not found" or a "Error no:500" it means that the poll.exe is not placed correctly.
' 7. If everything goes right please mail me at krishna_rath@hotmail.com or krath@rediffmail.com, ...Incase my hotmail account is hacked and you donot recieve an reply by a week.



Dim V1, V2, V3 'the 3 radio buttons
Dim Op  'opinion
Dim Uyes, Uno, Ucant, Utotal ' U stands for User...ie User says Yes is Uyes
Dim Pyes, Pno, Pcant 'percentages... Percentage of Yes


Sub inter_main()
' if you open the file as in the Windows Explorer then a message box is displayed
' saying that it is not an ordinary program!!
' I.e if you "double click" or "single click" the file then the following message
' can be displayed

MsgBox "Rath-India poll", vbOKOnly, "About" 'You can change this


End Sub

Sub CGI_Main()
' this is the sub where the program "gets" the information from the HTML page
'
If CGI_RequestMethod = "POST" Then
   
    On Error GoTo errhan

    R1 = GetSmallField("R1")
    
    f1 = FreeFile
    Open App.Path & "\poll.txt" For Input As f1 'poll.txt conatins info.
    Input #f1, Uyes, Uno, Ucant
    Close #f1
    
    
    If R1 = "V1" Then Uyes = Uyes + 1
    If R1 = "V2" Then Uno = Uno + 1
    If R1 = "V3" Then Ucant = Ucant + 1
    
    'calculating averages % tec
    Utotal = Uyes + Uno + Ucant: If Utotal = 0 Then Utotal = 1 'division by zero
    Pyes = (Uyes / Utotal) * 100: Pyes = Int(Pyes)
    Pno = (Uno / Utotal) * 100: Pno = Int(Pno)
    Pcant = (Ucant / Utotal) * 100: Pcant = Int(Pcant)
    ' note: the averages calcualted above donot sum up to 100
    ' can you reason out why?
    ' just think!
    
    f2 = FreeFile
    Open App.Path & "\poll.txt" For Output As f2    ' output the latest results
    Print #f2, Uyes, Uno, Ucant
    Close #f2
    
    
    Gen_HTML
    
    ' incase, of any error like the poll.txt not found in the folder , or it has been tampered with then
    ' an error will be displayed.
    ' a neat error handler showing the exact error can be created here
    
errhan:

    If Err.Number > 0 Then MsgBox "An error has taken place...check your configurations", vbCritical, "ERROR!!!!!!"
       

End If
End Sub

Sub Gen_HTML()
' Generate the HTML output
' use the Send procedure given in the CGI32.BAS

Send ("<html>")
Send ("<head>")
Send ("<title>Rich Poor POLL RESULTS</title></head>")

Send ("<body bgcolor=""#CCB3D0"">")

Send ("<h1 align=""center""><font color=""#FF0000"">THE RICH or POOR POLL RESULTS</font></h1>")

Send ("<p align=""center""><font color=""#FF0000""></font>&nbsp;</p>")

Send ("<table border=""0"" cellpadding=""2"" cellspacing=""4"" width=""100%"">")
Send ("<tr><td colspan=""3"" bgcolor=""#FFB871""><p align=""center""><font color=""#0000FF""><strong>Thank you for taking part in the polls. The results of the poll are as follows</strong></font></p></td></tr>")
Send ("<tr><td width=""20"" bgcolor=""#800000""><p align=""center""><font color=""#FFFF00""><em><strong>Opinion</strong></em></font></p></td>")
Send ("<td bgcolor=""#800000""><p align=""center""><font color=""#FFFF00""><em><strong>Votes</strong></em></font></p></td>")
Send ("<td bgcolor=""#800000""><p align=""center""><font color=""#FFFF00""><em><strong>Percentage</strong></em></font></p></td></tr>")
Send ("<tr><td width=""20%""><p align=""center""><strong>Yes</strong></td><td width=""33%""><p align=""center"">" & Uyes & "</td><td width=""34%""><p align=""center"">" & Pyes & "</p></td></tr>")
Send ("<tr><td width=""20%""><p align=""center""><strong>No</strong></td><td width=""33%""><p align=""center"">" & Uno & "</td><td width=""34%""><p align=""center"">" & Pno & "</p></td></tr>")
Send ("<tr><td width=""20%""><p align=""center""><strong>Cant say</strong></td><td width=""33%""><p align=""center"">" & Ucant & "</td><td width=""34%""><p align=""center"">" & Pcant & "</p></td></tr></table>")

Send ("<p>Please take part in the opinion poll and express your views.</p>")
Send ("<p>Thank you</p><p>&nbsp;</p><p><font color=""#008000"" size=""2""><em><a href=""mailto:krishna_rath@hotmail.com"">(mail to Vishnoo Rath and Krishna Rath )</a></em></font></p></body></html>")

End Sub


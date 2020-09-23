Attribute VB_Name = "cgi32"
'----------------------------------------------------------------------
'       *************
'       * CGI32.BAS *
'       *************
'
' VERSION: 1.9  (July 17, 1997)
'
' AUTHORS: Robert B. Denny <rdenny@dc3.com>
'          Christopher J. Duke <duke@ora.com>
'
' Common routines needed to establish a VB environment for
' Windows CGI programs that run behind the WebSite Server.
'
' INTRODUCTION
'
' The Common Gateway Interface (CGI) version 1.1 specifies a minimal
' set of data that is made available to the back-end application by
' an HTTP (Web) server. It also specifies the details for passing this
' information to the back-end. The latter part of the CGI spec is
' specific to Unix-like environments. The NCSA httpd for Windows does
' supply the data items (and more) specified by CGI/1.1, however it
' uses a different method for passing the data to the back-end.
'
' DEVELOPMENT
'
' WebSite requires any Windows back-end program to be an
' executable image. This means that you must convert your VB
' application into an executable (.EXE) before it can be tested
' with the server.
'
' ENVIRONMENT
'
' The WebSite server executes script requests by doing a
' CreateProcess with a command line in the following form:
'
'   prog-name cgi-profile
'
' THE CGI PROFILE FILE
'
' The Unix CGI passes data to the back end by defining environment
' variables which can be used by shell scripts. The WebSite
' server passes data to its back end via the profile file. The
' format of the profile is that of a Windows ".INI" file. The keyword
' names have been changed cosmetically.
'
' There are 7 sections in a CGI profile file, [CGI], [Accept],
' [System], [Extra Headers], and [Form Literal], [Form External],
' and [Form huge]. They are described below:
'
' [CGI]                <== The standard CGI variables
' CGI Version=         The version of CGI spoken by the server
' Request Protocol=    The server's info protocol (e.g. HTTP/1.0)
' Request Method=      The method specified in the request (e.g., "GET")
' Request Keep-Alive=  If the client requested connection re-use (Yes/No)
' Executable Path=     Physical pathname of the back-end (this program)
' Logical Path=        Extra path info in logical space
' Physical Path=       Extra path info in local physical space
' Query String=        String following the "?" in the request URL
' Content Type=        MIME content type of info supplied with request
' Content Length=      Length, bytes, of info supplied with request
' Request Range=       Byte-range specfication received with request
' Server Software=     Version/revision of the info (HTTP) server
' Server Name=         Server's network hostname (or alias from config)
' Server Port=         Server's network port number
' Server Admin=        E-Mail address of server's admin. (config)
' Referer=             URL of referring document
' From=                E-Mail of client user (rarely seen)
' User Agent=          String describing client/browser software/version
' Remote Host=         Remote client's network hostname
' Remote Address=      Remote client's network address
' Authenticated Username=Username if present in request
' Authenticated Password=Password if present in request
' Authentication Method=Method used for authentication (e.g., "Basic")
' Authentication Realm=Name of realm for users/groups
'
' [Accept]             <== What the client says it can take
' The MIME types found in the request header as
'    Accept: xxx/yyy; zzzz...
' are entered in this section as
'    xxx/yyy=zzzz...
' If only the MIME type appears, the form is
'    xxx/yyy=Yes
'
' [System]             <== Windows interface specifics
' GMT Offset=          Offset of local timezone from GMT, seconds (LONG!)
' Output File=         Pathname of file to receive results
' Content File=        Pathname of file containing raw request content
' Debug Mode=          If server's CGI debug flag is set (Yes/No)
'
' [Extra Headers]
' Any "extra" headers found in the request that activated this
' program. They are listed in "key=value" form. Usually, you'll see
' at least the name of the browser here as "User-agent".
'
' [Form Literal]
' If the request was a POST from a Mosaic form (with content type of
' "application/x-www-form-urlencoded"), the server will decode the
' form data. Raw form input is of the form "key=value&key=value&...",
' with the value parts "URL-encoded". The server splits the key=value
' pairs at the '&', then spilts the key and value at the '=',
' URL-decodes the value string and puts the result into key=value
' (decoded) form in the [Form Literal] section of the INI.
'
' [Form External]
' If the decoded value string is more than 254 characters long,
' or if the decoded value string contains any control characters
' or quote marks the server puts the decoded value into an external
' tempfile and lists the field in this section as:
'    key=<pathname> <length>
' where <pathname> is the path and name of the tempfile containing
' the decoded value string, and <length> is the length in bytes
' of the decoded value string.
'
' NOTE: BE SURE TO OPEN THIS FILE IN BINARY MODE UNLESS YOU ARE
'       CERTAIN THAT THE FORM DATA IS TEXT!
'
' [Form File]
' If the form data contained any uploaded files, they are described in
' this section as:
'    key=[<pathname>] <length> <type> <encoding> [<name>]
' where <pathname> is the path and name of the tempfile contining the
' uploaded file, <length> is the length in bytes of the uploaded file,
' <type> is the content type of the uploaded file as sent by the browser,
' <encoding> is the content-transfer encoding of the uploaded file, and
' <name> is the original file name of the uploaded file.
'
' [Form Huge]
' If the raw value string is more than 65,536 bytes long, the server
' does no decoding. In this case, the server lists the field in this
' section as:
'    key=<offset> <length>
' where <offset> is the offset from the beginning of the Content File
' at which the raw value string for this key is located, and <length>
' is the length in bytes of the raw value string. You can use the
' <offset> to perform a "Seek" to the start of the raw value string,
' and use the length to know when you have read the entire raw string
' into your decoder. Note that VB has a limit of 64K for strings, so
'
' Examples:
'
'    [Form Literal]
'    smallfield=123 Main St. #122
'
'    [Form External]
'    field300chars=c:\website\cgi-tmp\1a7fws.000 300
'    fieldwithlinebreaks=c:\website\cgi-tmp\1a7fws.001 43
'
'    [Form Huge]
'    field230K=c:\website\cgi-tmp\1a7fws.002 276920
'
' =====
' USAGE
' =====
' Include CGI32.BAS in your VB4 or VB5 project. Set the project options for
' "Sub Main" startup. The Main() procedure is in this module, and it
' handles all of the setup of the VB CGI environment, as described
' above. Once all of this is done, the Main() calls YOUR main procedure
' which must be called CGI_Main(). The output file is open, use Send()
' to write to it. The input file is NOT open, and "huge" form fields
' have not been decoded.
'
' NOTE: If your program is started without command-line args,
' the code assumes you want to run it interactively. This is useful
' for providing a setup screen, etc. Instead of calling CGI_Main(),
' it calls Inter_Main(). Your module must also implement this
' function. If you don't need an interactive mode, just create
' Inter_Main() and put a 1-line call to MsgBox alerting the
' user that the program is not meant to be run interactively.
' The samples furnished with the server do this.
'
' If a Visual Basic runtime error occurs, it will be trapped and result
' in an HTTP error response being sent to the client. Check out the
' Error Handler() sub. When your program finishes, be sure to RETURN
' TO MAIN(). Don't just do an "End".
'
' Have a look at the stuff below to see what's what.
'
'----------------------------------------------------------------------
' Author:   Robert B. Denny <rdenny@netcom.com>
'           April 15, 1995
'
' Revision History:
'   15-Apr-95 rbd   Initial release (ref VB3 CGI.BAS 1.7)
'   02-Aug-95 rbd   Changed to take input and output files from profile
'                   Server no longer produces long command line.
'   24-Aug-95 rbd   Make call to GetPrivateProfileString conditional
'                   so 16-bit and 32-bit versions supported. Fix
'                   computation of CGI_GMTOffset for offset=0 (GMT)
'                   case. Add FieldPresent() routine for checkbox
'                   handling. Clean up comments.
'   29-Oct-95 rbd   Added PlusToSpace() and Unescape() functions for
'                   decoding query strings, etc.
'   16-Nov-95 rbd   Add keep-alive variable, file uploading description
'                   in comments, and upload display.
'   20-Nov-95 rbd   Fencepost error in ParseFileValue()
'   23-Nov-95 rbd   Remove On Error Resume Next from error handler
'   03-Dec-95 rbd   User-Agent is now a variable, real HTTP header
'                   Add Request-Range as http header as well.
'   30-Dec-96 rbd   Fix "endless" loop if call FieldPresent() with
'                   zero form fields.
'   23-Feb-97 rbd   Per MS Tech Support, do not do Exit Sub in sub Main
'                   as this can corrupt VB40032.DLL. Use End instead.
'   17-Jul-97 cjd   1.9: Removed MAX_FORM_TUPLES (was previously a
'                   constant set to 100) and replaced it with a
'                   varying-sized array in GetFormTuples().
'----------------------------------------------------------------------
Option Explicit
'
' ==================
' Manifest Constants
' ==================
'
Const MAX_CMDARGS = 8       ' Max # of command line args
Const ENUM_BUF_SIZE = 4096  ' Key enumeration buffer, see GetProfile()
' These are the limits in the server
Const MAX_XHDR = 100        ' Max # of "extra" request headers
Const MAX_ACCTYPE = 100     ' Max # of Accept: types in request
Const MAX_HUGE_TUPLES = 16  ' Max # "huge" form fields
Const MAX_FILE_TUPLES = 16  ' Max # of uploaded file tuples
'
'
' =====
' Types
' =====
'
Type Tuple                  ' Used for Accept: and "extra" headers
    key As String           ' and for holding POST form key=value pairs
    value As String
End Type

Type FileTuple              ' Used for form-based file uploads
    key As String           ' Form field name
    file As String          ' Local tempfile containing uploaded file
    length As Long          ' Length in bytes of uploaded file
    type As String          ' Content type of uploaded file
    encoding As String      ' Content-transfer encoding of uploaded file
    name As String          ' Original name of uploaded file
End Type

Type HugeTuple              ' Used for "huge" form fields
    key As String           ' Keyword (decoded)
    offset As Long          ' Byte offset into Content File of value
    length As Long          ' Length of value, bytes
End Type
'
'
' ================
' Global Constants
' ================
'
' -----------
' Error Codes
' -----------
'
Global Const ERR_ARGCOUNT = 32767
Global Const ERR_BAD_REQUEST = 32766        ' HTTP 400
Global Const ERR_UNAUTHORIZED = 32765       ' HTTP 401
Global Const ERR_PAYMENT_REQUIRED = 32764   ' HTTP 402
Global Const ERR_FORBIDDEN = 32763          ' HTTP 403
Global Const ERR_NOT_FOUND = 32762          ' HTTP 404
Global Const ERR_INTERNAL_ERROR = 32761     ' HTTP 500
Global Const ERR_NOT_IMPLEMENTED = 32760    ' HTTP 501
Global Const ERR_TOO_BUSY = 32758           ' HTTP 503 (experimental)
Global Const ERR_NO_FIELD = 32757           ' GetxxxField "no field"
Global Const CGI_ERR_START = 32757          ' Start of our errors

' ====================
' CGI Global Variables
' ====================
'
' ----------------------
' Standard CGI variables
' ----------------------
'
Global CGI_ServerSoftware As String
Global CGI_ServerName As String
Global CGI_ServerPort As Integer
Global CGI_RequestProtocol As String
Global CGI_ServerAdmin As String
Global CGI_Version As String
Global CGI_RequestMethod As String
Global CGI_RequestKeepAlive As Integer
Global CGI_LogicalPath As String
Global CGI_PhysicalPath As String
Global CGI_ExecutablePath As String
Global CGI_QueryString As String
Global CGI_RequestRange As String
Global CGI_Referer As String
Global CGI_From As String
Global CGI_UserAgent As String
Global CGI_RemoteHost As String
Global CGI_RemoteAddr As String
Global CGI_AuthUser As String
Global CGI_AuthPass As String
Global CGI_AuthType As String
Global CGI_AuthRealm As String
Global CGI_ContentType As String
Global CGI_ContentLength As Long
'
' ------------------
' HTTP Header Arrays
' ------------------
'
Global CGI_AcceptTypes(MAX_ACCTYPE) As Tuple    ' Accept: types
Global CGI_NumAcceptTypes As Integer            ' # of live entries in array
Global CGI_ExtraHeaders(MAX_XHDR) As Tuple      ' "Extra" headers
Global CGI_NumExtraHeaders As Integer           ' # of live entries in array
'
' --------------
' POST Form Data
' --------------
'
Global CGI_FormTuples() As Tuple                ' Declare dynamic array for POST form key=value pairs
Global CGI_NumFormTuples As Integer             ' # of live entries in array
Global CGI_HugeTuples(MAX_HUGE_TUPLES) As HugeTuple ' Form "huge tuples
Global CGI_NumHugeTuples As Integer             ' # of live entries in array
Global CGI_FileTuples(MAX_FILE_TUPLES) As FileTuple ' File upload tuples
Global CGI_NumFileTuples As Integer             ' # of live entries in array
'
' ----------------
' System Variables
' ----------------
'
Global CGI_GMTOffset As Variant         ' GMT offset (time serial)
Global CGI_ContentFile As String        ' Content/Input file pathname
Global CGI_OutputFile As String         ' Output file pathname
Global CGI_DebugMode As Integer         ' Script Tracing flag from server
'
'
' ========================
' Windows API Declarations
' ========================
'
' NOTE: Declaration of GetPrivateProfileString is specially done to
' permit enumeration of keys by passing NULL key value. See GetProfile().
' Both the 16-bit and 32-bit flavors are given below. We DO NOT
' recommend using 16-bit VB4 with WebSite!
'
#If Win32 Then
Declare Function GetPrivateProfileString Lib "kernel32" _
    Alias "GetPrivateProfileStringA" _
   (ByVal lpApplicationName As String, _
    ByVal lpKeyName As Any, _
    ByVal lpDefault As String, _
    ByVal lpReturnedString As String, _
    ByVal nSize As Long, _
    ByVal lpFileName As String) As Long
#Else
Declare Function GetPrivateProfileString Lib "Kernel" _
   (ByVal lpSection As String, _
    ByVal lpKeyName As Any, _
    ByVal lpDefault As String, _
    ByVal lpReturnedString As String, _
    ByVal nSize As Integer, _
    ByVal lpFileName As String) As Integer
#End If
'
'
' ===============
' Local Variables
' ===============
'
Dim CGI_ProfileFile As String           ' Profile file pathname
Dim CGI_OutputFN As Integer             ' Output file number
Dim ErrorString As String
'---------------------------------------------------------------------------
'
'   ErrorHandler() - Global error handler
'
' If a VB runtime error occurs dusing execution of the program, this
' procedure generates an HTTP/1.0 HTML-formatted error message into
' the output file, then exits the program.
'
' This should be armed immediately on entry to the program's main()
' procedure. Any errors that occur in the program are caught, and
' an HTTP/1.0 error messsage is generated into the output file. The
' presence of the HTTP/1.0 on the first line of the output file causes
' NCSA httpd for WIndows to send the output file to the client with no
' interpretation or other header parsing.
'---------------------------------------------------------------------------
Sub ErrorHandler(code As Integer)

    Seek #CGI_OutputFN, 1    ' Rewind output file just in case
    Send ("HTTP/1.0 500 Internal Error")
    Send ("Server: " + CGI_ServerSoftware)
    Send ("Date: " + WebDate(Now))
    Send ("Content-type: text/html")
    Send ("")
    Send ("<HTML><HEAD>")
    Send ("<TITLE>Error in " + CGI_ExecutablePath + "</TITLE>")
    Send ("</HEAD><BODY>")
    Send ("<H1>Error in " + CGI_ExecutablePath + "</H1>")
    Send ("An internal Visual Basic error has occurred in " + CGI_ExecutablePath + ".")
    Send ("<PRE>" + ErrorString + "</PRE>")
    Send ("<I>Please</I> note what you were doing when this problem occurred,")
    Send ("so we can identify and correct it. Write down the Web page you were using,")
    Send ("any data you may have entered into a form or search box, and")
    Send ("anything else that may help us duplicate the problem. Then contact the")
    Send ("administrator of this service: ")
    Send ("<A HREF=""mailto:" & CGI_ServerAdmin & """>")
    Send ("<ADDRESS>&lt;" + CGI_ServerAdmin + "&gt;</ADDRESS>")
    Send ("</A></BODY></HTML>")

    Close #CGI_OutputFN

    '======
     End            ' Terminate the program
    '======
End Sub
'---------------------------------------------------------------------------
'
'   GetAcceptTypes() - Create the array of accept type structs
'
' Enumerate the keys in the [Accept] section of the profile file,
' then get the value for each of the keys.
'---------------------------------------------------------------------------
Private Sub GetAcceptTypes()
    Dim sList As String
    Dim i As Integer, j As Integer, l As Integer, n As Integer

    sList = GetProfile("Accept", "") ' Get key list
    l = Len(sList)                          ' Length incl. trailing null
    i = 1                                   ' Start at 1st character
    n = 0                                   ' Index in array
    Do While ((i < l) And (n < MAX_ACCTYPE)) ' Safety stop here
        j = InStr(i, sList, Chr$(0))        ' J -> next null
        CGI_AcceptTypes(n).key = Mid$(sList, i, j - i) ' Get Key, then value
        CGI_AcceptTypes(n).value = GetProfile("Accept", CGI_AcceptTypes(n).key)
        i = j + 1                           ' Bump pointer
        n = n + 1                           ' Bump array index
    Loop
    CGI_NumAcceptTypes = n                  ' Fill in global count

End Sub
'---------------------------------------------------------------------------
'
'   GetArgs() - Parse the command line
'
' Chop up the command line, fill in the argument vector, return the
' argument count (similar to the Unix/C argc/argv handling)
'---------------------------------------------------------------------------
Private Function GetArgs(argv() As String) As Integer
    Dim buf As String
    Dim i As Integer, j As Integer, l As Integer, n As Integer

    buf = Trim$(Command$)                   ' Get command line

    l = Len(buf)                            ' Length of command line
    If l = 0 Then                           ' If empty
        GetArgs = 0                         ' Return argc = 0
        Exit Function
    End If

    i = 1                                   ' Start at 1st character
    n = 0                                   ' Index in argvec
    Do While ((i < l) And (n < MAX_CMDARGS)) ' Safety stop here
        j = InStr(i, buf, " ")              ' J -> next space
        If j = 0 Then Exit Do               ' Exit loop on last arg
        argv(n) = Trim$(Mid$(buf, i, j - i)) ' Get this token, trim it
        i = j + 1                           ' Skip that blank
        Do While Mid$(buf, i, 1) = " "      ' Skip any additional whitespace
            i = i + 1
        Loop
        n = n + 1                           ' Bump array index
    Loop

    argv(n) = Trim$(Mid$(buf, i, (l - i + 1))) ' Get last arg
    GetArgs = n + 1                         ' Return arg count

End Function
'---------------------------------------------------------------------------
'
'   GetExtraHeaders() - Create the array of extra header structs
'
' Enumerate the keys in the [Extra Headers] section of the profile file,
' then get the value for each of the keys.
'---------------------------------------------------------------------------
Private Sub GetExtraHeaders()
    Dim sList As String
    Dim i As Integer, j As Integer, l As Integer, n As Integer

    sList = GetProfile("Extra Headers", "") ' Get key list
    l = Len(sList)                          ' Length incl. trailing null
    i = 1                                   ' Start at 1st character
    n = 0                                   ' Index in array
    Do While ((i < l) And (n < MAX_XHDR))   ' Safety stop here
        j = InStr(i, sList, Chr$(0))        ' J -> next null
        CGI_ExtraHeaders(n).key = Mid$(sList, i, j - i) ' Get Key, then value
        CGI_ExtraHeaders(n).value = GetProfile("Extra Headers", CGI_ExtraHeaders(n).key)
        i = j + 1                           ' Bump pointer
        n = n + 1                           ' Bump array index
    Loop
    CGI_NumExtraHeaders = n                 ' Fill in global count

End Sub
'---------------------------------------------------------------------------
'
'   GetFormTuples() - Create the array of POST form input key=value pairs
'
'---------------------------------------------------------------------------
Private Sub GetFormTuples()
    Dim sList As String
    Dim i As Integer, j As Integer, k As Integer
    Dim l As Integer, m As Integer, n As Integer
    Dim s As Long
    Dim buf As String
    Dim extName As String
    Dim extFile As Integer
    Dim extlen As Long

    n = 0                                     ' Index in array
    ReDim Preserve CGI_FormTuples(n) As Tuple ' Increase array size

    '
    ' Do the easy one first: [Form Literal]
    '
    sList = GetProfile("Form Literal", "")  ' Get key list
    l = Len(sList)                          ' Length incl. trailing null
    i = 1                                   ' Start at 1st character
    Do While i < l                          ' Safety stop here
        j = InStr(i, sList, Chr$(0))        ' J -> next null
        CGI_FormTuples(n).key = Mid$(sList, i, j - i) ' Get Key, then value
        CGI_FormTuples(n).value = GetProfile("Form Literal", CGI_FormTuples(n).key)
        i = j + 1                           ' Bump pointer
        n = n + 1                           ' Bump array index
        ReDim Preserve CGI_FormTuples(n) As Tuple ' Increase array size
    Loop
    '
    ' Now do the external ones: [Form External]
    '
    sList = GetProfile("Form External", "") ' Get key list
    l = Len(sList)                          ' Length incl. trailing null
    i = 1                                   ' Start at 1st character
    extFile = FreeFile
    Do While i < l                          ' Safety stop here
        j = InStr(i, sList, Chr$(0))        ' J -> next null
        CGI_FormTuples(n).key = Mid$(sList, i, j - i) ' Get Key, then pathname
        buf = GetProfile("Form External", CGI_FormTuples(n).key)
        k = InStr(buf, " ")                 ' Split file & length
        extName = Mid$(buf, 1, k - 1)           ' Pathname
        k = k + 1
        extlen = CLng(Mid$(buf, k, Len(buf) - k + 1)) ' Length
        '
        ' Use feature of GET to read content in one call
        '
        Open extName For Binary Access Read As #extFile
        CGI_FormTuples(n).value = String$(extlen, " ") ' Breathe in...
        Get #extFile, , CGI_FormTuples(n).value 'GULP!
        Close #extFile
        i = j + 1                           ' Bump pointer
        n = n + 1                           ' Bump array index
        ReDim Preserve CGI_FormTuples(n) As Tuple ' Increase array size
    Loop

    CGI_NumFormTuples = n                   ' Number of fields decoded
    n = 0                                   ' Reset counter
    '
    ' Next, the [Form Huge] section. Will this ever get executed?
    '
    sList = GetProfile("Form Huge", "")     ' Get key list
    l = Len(sList)                          ' Length incl. trailing null
    i = 1                                   ' Start at 1st character
    Do While i < l                          ' Safety stop here
        j = InStr(i, sList, Chr$(0))        ' J -> next null
        CGI_HugeTuples(n).key = Mid$(sList, i, j - i) ' Get Key
        buf = GetProfile("Form Huge", CGI_HugeTuples(n).key) ' "offset length"
        k = InStr(buf, " ")                 ' Delimiter
        CGI_HugeTuples(n).offset = CLng(Mid$(buf, 1, (k - 1)))
        CGI_HugeTuples(n).length = CLng(Mid$(buf, k, (Len(buf) - k + 1)))
        i = j + 1                           ' Bump pointer
        n = n + 1                           ' Bump array index
        ReDim Preserve CGI_FormTuples(n) As Tuple ' Increase array size
    Loop
    
    CGI_NumHugeTuples = n                   ' Fill in global count

    n = 0                                   ' Reset counter
    '
    ' Finally, the [Form File] section.
    '
    sList = GetProfile("Form File", "")     ' Get key list
    l = Len(sList)                          ' Length incl. trailing null
    i = 1                                   ' Start at 1st character
    Do While ((i < l) And (n < MAX_FILE_TUPLES)) ' Safety stop here
        j = InStr(i, sList, Chr$(0))        ' J -> next null
        CGI_FileTuples(n).key = Mid$(sList, i, j - i) ' Get Key
        buf = GetProfile("Form File", CGI_FileTuples(n).key)
        ParseFileValue buf, CGI_FileTuples(n)  ' Complicated, use Sub
        i = j + 1                           ' Bump pointer
        n = n + 1                           ' Bump array index
    Loop
    
    CGI_NumFileTuples = n                   ' Fill in global count

End Sub
'---------------------------------------------------------------------------
'
'   GetProfile() - Get a value or enumerate keys in CGI_Profile file
'
' Get a value given the section and key, or enumerate keys given the
' section name and "" for the key. If enumerating, the list of keys for
' the given section is returned as a null-separated string, with a
' double null at the end.
'
' VB handles this with flair! I couldn't believe my eyes when I tried this.
'---------------------------------------------------------------------------
Private Function GetProfile(sSection As String, sKey As String) As String
    Dim retLen As Long
    Dim buf As String * ENUM_BUF_SIZE

    If sKey <> "" Then
        retLen = GetPrivateProfileString(sSection, sKey, "", buf, ENUM_BUF_SIZE, CGI_ProfileFile)
    Else
        retLen = GetPrivateProfileString(sSection, 0&, "", buf, ENUM_BUF_SIZE, CGI_ProfileFile)
    End If
    If retLen = 0 Then
        GetProfile = ""
    Else
        GetProfile = Left$(buf, retLen)
    End If

End Function
'----------------------------------------------------------------------
'
' Get the value of a "small" form field given the key
'
' Signals an error if field does not exist
'
'----------------------------------------------------------------------
Function GetSmallField(key As String) As String
    Dim i As Integer

    For i = 0 To (CGI_NumFormTuples - 1)
        If CGI_FormTuples(i).key = key Then
            GetSmallField = Trim$(CGI_FormTuples(i).value)
            Exit Function           ' ** DONE **
        End If
    Next i
    '
    ' Field does not exist
    '
    Error ERR_NO_FIELD
End Function
'---------------------------------------------------------------------------
'
'   InitializeCGI() - Fill in all of the CGI variables, etc.
'
' Read the profile file name from the command line, then fill in
' the CGI globals, the Accept type list and the Extra headers list.
' Then open the input and output files.
'
' Returns True if OK, False if some sort of error. See ReturnError()
' for info on how errors are handled.
'
' NOTE: Assumes that the CGI error handler has been armed with On Error
'---------------------------------------------------------------------------
Sub InitializeCGI()
    Dim sect As String
    Dim argc As Integer
    Static argv(MAX_CMDARGS) As String
    Dim buf As String

    CGI_DebugMode = True    ' Initialization errors are very bad

    '
    ' Parse the command line. We need the profile file name (duh!)
    ' and the output file name NOW, so we can return any errors we
    ' trap. The error handler writes to the output file.
    '
    argc = GetArgs(argv())
    CGI_ProfileFile = argv(0)

    sect = "CGI"
    CGI_ServerSoftware = GetProfile(sect, "Server Software")
    CGI_ServerName = GetProfile(sect, "Server Name")
    CGI_RequestProtocol = GetProfile(sect, "Request Protocol")
    CGI_ServerAdmin = GetProfile(sect, "Server Admin")
    CGI_Version = GetProfile(sect, "CGI Version")
    CGI_RequestMethod = GetProfile(sect, "Request Method")
    buf = GetProfile(sect, "Request Keep-Alive")    ' Y or N
    If (Left$(buf, 1) = "Y") Then                   ' Must start with Y
        CGI_RequestKeepAlive = True
    Else
        CGI_RequestKeepAlive = False
    End If
    CGI_LogicalPath = GetProfile(sect, "Logical Path")
    CGI_PhysicalPath = GetProfile(sect, "Physical Path")
    CGI_ExecutablePath = GetProfile(sect, "Executable Path")
    CGI_QueryString = GetProfile(sect, "Query String")
    CGI_RemoteHost = GetProfile(sect, "Remote Host")
    CGI_RemoteAddr = GetProfile(sect, "Remote Address")
    CGI_RequestRange = GetProfile(sect, "Request Range")
    CGI_Referer = GetProfile(sect, "Referer")
    CGI_From = GetProfile(sect, "From")
    CGI_UserAgent = GetProfile(sect, "User Agent")
    CGI_AuthUser = GetProfile(sect, "Authenticated Username")
    CGI_AuthPass = GetProfile(sect, "Authenticated Password")
    CGI_AuthRealm = GetProfile(sect, "Authentication Realm")
    CGI_AuthType = GetProfile(sect, "Authentication Method")
    CGI_ContentType = GetProfile(sect, "Content Type")
    buf = GetProfile(sect, "Content Length")
    If buf = "" Then
        CGI_ContentLength = 0
    Else
        CGI_ContentLength = CLng(buf)
    End If
    buf = GetProfile(sect, "Server Port")
    If buf = "" Then
        CGI_ServerPort = -1
    Else
        CGI_ServerPort = CInt(buf)
    End If

    sect = "System"
    CGI_ContentFile = GetProfile(sect, "Content File")
    CGI_OutputFile = GetProfile(sect, "Output File")
    CGI_OutputFN = FreeFile
    Open CGI_OutputFile For Output Access Write As #CGI_OutputFN
    buf = GetProfile(sect, "GMT Offset")
    If buf <> "" Then                             ' Protect against errors
        CGI_GMTOffset = CVDate(Val(buf) / 86400#) ' Timeserial GMT offset
    Else
        CGI_GMTOffset = 0
    End If
    buf = GetProfile(sect, "Debug Mode")    ' Y or N
    If (Left$(buf, 1) = "Y") Then           ' Must start with Y
        CGI_DebugMode = True
    Else
        CGI_DebugMode = False
    End If

    GetAcceptTypes          ' Enumerate Accept: types into tuples
    GetExtraHeaders         ' Enumerate extra headers into tuples
    GetFormTuples           ' Decode any POST form input into tuples

End Sub
'----------------------------------------------------------------------
'
'   main() - CGI script back-end main procedure
'
' This is the main() for the VB back end. Note carefully how the error
' handling is set up, and how program cleanup is done. If no command
' line args are present, call Inter_Main() and exit.
'----------------------------------------------------------------------
Sub Main()
    On Error GoTo ErrorHandler

    If Trim$(Command$) = "" Then    ' Interactive start
        inter_main                  ' Call interactive main
        End                         ' Exit the program
    End If

    InitializeCGI       ' Create the CGI environment

    '===========
    cgi_main            ' Execute the actual "script"
    '===========

Cleanup:
    Close #CGI_OutputFN

    End                             ' End the program
'------------
ErrorHandler:
    Select Case Err                 ' Decode our "user defined" errors
        Case ERR_NO_FIELD:
            ErrorString = "Unknown form field"
        Case Else:
            ErrorString = Error$    ' Must be VB error
    End Select

    ErrorString = ErrorString & " (error #" & Err & ")"
    On Error GoTo 0                 ' Prevent recursion
    ErrorHandler (Err)              ' Generate HTTP error result
    Resume Cleanup
'------------
End Sub
'----------------------------------------------------------------------
'
'  Send() - Shortcut for writing to output file
'
'----------------------------------------------------------------------
Sub Send(s As String)
    Print #CGI_OutputFN, s
End Sub
'---------------------------------------------------------------------------
'
'   SendNoOp() - Tell browser to do nothing.
'
' Most browsers will do nothing. Netscape 1.0N leaves hourglass
' cursor until the mouse is waved around. Enhanced Mosaic 2.0
' oputs up an alert saying "URL leads nowhere". Your results may
' vary...
'
'---------------------------------------------------------------------------
Sub SendNoOp()

    Send ("HTTP/1.0 204 No Response")
    Send ("Server: " + CGI_ServerSoftware)
    Send ("")

End Sub
'---------------------------------------------------------------------------
'
'   WebDate - Return an HTTP/1.0 compliant date/time string
'
' Inputs:   t = Local time as VB Variant (e.g., returned by Now())
' Returns:  Properly formatted HTTP/1.0 date/time in GMT
'---------------------------------------------------------------------------
Function WebDate(dt As Variant) As String
    Dim t As Variant
    
    t = CVDate(dt - CGI_GMTOffset)      ' Convert time to GMT
    WebDate = Format$(t, "ddd dd mmm yyyy hh:mm:ss") & " GMT"

End Function

'----------------------------------------------------------------------
'
' Return True/False depending on whether a form field is present.
' Typically used to detect if a checkbox in a form is checked or
' not. Unchecked checkboxes are omitted from the form content.
'
'----------------------------------------------------------------------
Function FieldPresent(key As String) As Integer
    Dim i As Integer

    FieldPresent = False            ' Assume failure
    
    If (CGI_NumFormTuples = 0) Then Exit Function   ' Stop endless loop
    
    For i = 0 To (CGI_NumFormTuples - 1)
        If CGI_FormTuples(i).key = key Then
            FieldPresent = True     ' Found it
            Exit Function           ' ** DONE **
        End If
    Next i
                                    ' Exit with FieldPresent still False
End Function

'----------------------------------------------------------------------
'
' PlusToSpace() - Remove plus-delimiters from HTTP-encoded string
'
'----------------------------------------------------------------------
Public Sub PlusToSpace(s As String)
    Dim i As Integer
    
    i = 1
    Do While True
        i = InStr(i, s, "+")
        If i = 0 Then Exit Do
        Mid$(s, i) = " "
    Loop

End Sub
'----------------------------------------------------------------------
'
' Unescape() - Convert HTTP-escaped string to normal form
'
'----------------------------------------------------------------------
Public Function Unescape(s As String)
    Dim i As Integer, l As Integer
    Dim c As String
    
    If InStr(s, "%") = 0 Then               ' Catch simple case
        Unescape = s
        Exit Function
    End If
    
    l = Len(s)
    Unescape = ""
    For i = 1 To l
        c = Mid$(s, i, 1)                   ' Next character
        If c = "%" Then
            If Mid$(s, i + 1, 1) = "%" Then
                c = "%"
                i = i + 1                   ' Loop increments too
            Else
                c = x2c(Mid$(s, i + 1, 2))
                i = i + 2                   ' Loop increments too
            End If
        End If
        Unescape = Unescape & c
    Next i

End Function
'----------------------------------------------------------------------
'
' x2c() - Convert hex-escaped character to ASCII
'
'----------------------------------------------------------------------
Private Function x2c(s As String) As String
    Dim t As String
    
    t = "&H" & s
    x2c = Chr$(CInt(t))

End Function
Private Sub ParseFileValue(buf As String, ByRef t As FileTuple)
    Dim i, j, k, l As Integer
    
    l = Len(buf)
    
    i = InStr(buf, " ")                     ' First delimiter
    t.file = Mid$(buf, 1, (i - 1))          ' [file]
    t.file = Mid$(t.file, 2, Len(t.file) - 2)  ' file
    
    j = InStr((i + 1), buf, " ")            ' Next delimiter
    t.length = CLng(Mid$(buf, (i + 1), (j - i - 1)))
    i = j
    
    j = InStr((i + 1), buf, " ")            ' Next delimiter
    t.type = Mid$(buf, (i + 1), (j - i - 1))
    i = j
    
    j = InStr((i + 1), buf, " ")            ' Next delimiter
    t.encoding = Mid$(buf, (i + 1), (j - i - 1))
    i = j
    
    t.name = Mid$(buf, (i + 1), (l - i - 1))  ' [name]
    t.name = Mid$(t.name, 2, Len(t.name) - 1) ' name

End Sub
'---------------------------------------------------------------------------
'
'   FindExtraHeader() - Get the text from an "extra" header
'
' Given the extra header's name, return the stuff after the ":"
' or an empty string if not there.
'---------------------------------------------------------------------------
Public Function FindExtraHeader(key As String) As String
    Dim i As Integer

    For i = 0 To (CGI_NumExtraHeaders - 1)
        If CGI_ExtraHeaders(i).key = key Then
            FindExtraHeader = Trim$(CGI_ExtraHeaders(i).value)
            Exit Function           ' ** DONE **
        End If
    Next i
    '
    ' Not present, return empty string
    '
    FindExtraHeader = ""
End Function

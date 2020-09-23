VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_wrd_report_manager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'******************************************************************
' Name:         This Class object provides the rules needed to
'               produce a word mail merge from by passing the
'               template file and a valid recordset.
'******************************************************************
' Description:  Alls a word mail merge document to be created from
'               an ADO recordset and a template file.
'******************************************************************
' (c) Illuminaries Ltd 2002 . All rights
'     reserved.
'******************************************************************
'   Written By : Jonathan Adams (MCP)
'   web address : www.illuminaries.co.uk
'******************************************************************
'   Change the code as you please but please leave the
'   copyright information
'******************************************************************
' implements the word routines class
'******************************************************************
'Usage :-
'
'Dim  iobj_word as cls_wrd_report_manager
'
'  Set iobj_word = New cls_wrd_report_manager
'
'   rs_data =the recordset to output
'   s_report_path :=Patha and filename of template document
'   b_Print := True or Fales (Print or Show)
'
'   Call iobj_word.produce_mail_merge(rs_data, s_report_path, b_Print)
'******************************************************************
Private obj_wrd_routines As i_word_routines



 







Private Function pf_write_recordset_to_HTML(ByVal obj_recordset As ADODB.Recordset, ByRef s_header_record As String) As String
    On Error Resume Next
    Dim rsField As ADODB.Field
    Dim s_output As String
    Dim s_ref As String
    ' convert the recordset information into an HTML Table
    pf_write_recordset_to_HTML = ""
    
    s_header_record = ""
    obj_recordset.MoveFirst
    
    s_output = s_output & "<table><tr>"
    For Each rsField In obj_recordset.Fields
        s_ref = rsField.Name
        s_output = s_output & "<th>" & s_ref & "</th>"
        s_header_record = s_header_record & s_ref & ", "
    Next rsField
   s_output = s_output & "</tr>"
    Do Until obj_recordset.EOF
        s_output = s_output & "<tr>"
        For Each rsField In obj_recordset.Fields
                s_ref = rsField.Name
                If Not IsNull(rsField.Value) And Not IsEmpty(rsField.Value) Then
                    s_output = s_output & "<td>" & rsField.Value & "</td>"
                Else
                     s_output = s_output & "<td></td>"
                End If
        Next rsField
        s_output = s_output & "</tr>"
       
        obj_recordset.MoveNext
    Loop
    s_output = s_output & "</table>"
    ' return the HTML text to the caller
    pf_write_recordset_to_HTML = s_output
End Function




Public Function Get_Application_Pointer() As Object
    On Error GoTo get_word_application_pointer_errorhandler
    Dim wrd_app         As Object
    Dim b_got_error     As Boolean
    ' attempt to attached to an exsiting instance of word
    Set wrd_app = GetObject(, "Word.Application")
    
    If b_got_error Then
        ' if not create a new one.
         Set wrd_app = CreateObject("Word.Application")
    End If
    ' return the object reference
    Set Get_Application_Pointer = wrd_app
    

get_word_application_pointer_exit:

Exit Function
get_word_application_pointer_errorhandler:
    b_got_error = True
    Resume Next

Create_mailmerge_data_file_Exit:
Exit Function
Create_mailmerge_data_file_Errorhandler:
    Err.Raise Err.Number, "cls_wrd_report_manager.Get_Application_Pointer", Err.Description
    Resume
    Resume Create_mailmerge_data_file_Exit
End Function


Private Function pf_create_merge_data(ByVal rs_merge_data As ADODB.Recordset, ByRef s_header_record As String) As String
    '---------------------------------------------------------------------------
    '   Routine will create the mail merge document data file from the
    '   passed recordset object and passback the filename to the
    '   calling function
    '---------------------------------------------------------------------------
    On Error GoTo pf_create_merge_data_errorhandler
    Dim s_data_filename     As String
    Dim s_HTML_String       As String
    Dim l_filehandle        As Long
    Dim obj_word            As Object
    Dim obj_document           As Object
    pf_create_merge_data = False
    s_data_filename = pf_get_tempory_filename()
    
    s_HTML_String = pf_write_recordset_to_HTML(rs_merge_data, s_header_record)

    If Dir(s_data_filename) <> "" Then
        Kill s_data_filename
    End If

    l_filehandle = FreeFile
    
    ' create the HTML file version
    Open s_data_filename For Output As #l_filehandle
        Print #l_filehandle, s_HTML_String
    Close #l_filehandle
    ' create new application object
    Set obj_word = Get_Application_Pointer
    ' retrieve a document pointer
    Set obj_document = obj_word.Documents.Add
    
    If Not obj_document Is Nothing Then
        If Dir(Left(s_data_filename, Len(s_data_filename) - 3) & "doc") <> "" Then
            Kill Left(s_data_filename, Len(s_data_filename) - 3) & "doc"
        End If
        ' open word
        'obj_word.Application.Visible = True
        ' insert the HTML document into the word document
        obj_document.Range.InsertFile s_data_filename
        obj_document.SaveAs Left(s_data_filename, Len(s_data_filename) - 3) & "doc"
        obj_document.Close
    End If
    
    ' clear the word object reference
    Set obj_document = Nothing
    Set obj_word = Nothing
    
    pf_create_merge_data = Left(s_data_filename, Len(s_data_filename) - 3) & "doc"



pf_create_merge_data_exit:

Exit Function
pf_create_merge_data_errorhandler:
    Err.Raise Err.Number, "cls_wrd_report_manager.pf_create_merge_data", Err.Source
    GoTo pf_create_merge_data_exit
End Function

Private Function pf_get_tempory_filename() As String
    On Error GoTo pf_get_tempory_filename_errorhandler
    Dim s_file_name  As String
    
    pf_get_tempory_filename = ""
    ' call standard routine to return a file without
    ' header information
    s_file_name = get_tmp_file_name()
    
    If s_file_name <> "" Then
        ' add the relvant file extension to the dataflie
        s_file_name = s_file_name & ".htm"
    
    End If
    
    pf_get_tempory_filename = s_file_name

pf_get_tempory_filename_exit:


Exit Function
pf_get_tempory_filename_errorhandler:
    Err.Raise Err.Number, "cls_wrd_report_manager.pf_get_tempory_filename", Err.Source
    GoTo pf_get_tempory_filename_exit
End Function


Public Function produce_mail_merge(ByVal rs_merge_recordset As ADODB.Recordset, ByVal s_word_template_file_name As String, ByVal b_Print As Boolean) As Boolean
    On Error GoTo produce_mail_merge_errorhandler
    Dim obj_app As Object
    Dim wrd_doc As Object
    Dim s_header_record As String
    Dim s_message As String
    Dim s_datafile As String
    ' create a new pointer to a the word document
    Set obj_app = Me.Get_Application_Pointer()
    


    ' firstly, create a data file to hold the
    ' mailmerge data.
    s_datafile = pf_create_merge_data(rs_merge_recordset, s_header_record)
     
    
    ' Check the required word template is present on the
    ' system.
    If Dir(s_word_template_file_name) <> "" Then
        ' Open the file to insert data
        Set wrd_doc = obj_app.Documents.Open(s_word_template_file_name)
    Else
        s_message = "Template : " & s_word_template_file_name & String(2, vbCrLf) & _
            "The required word template file could not be found within your system. " & _
            "Quest will now create a blank record with the mail merge data you require."
        MsgBox s_message, vbInformation
        Set wrd_doc = obj_app.Documents.Add
    End If
    ' Take a copy of the document
    ' save to a tempory file name
    wrd_doc.SaveAs Left(s_datafile, Len(s_datafile) - 4) & "_output.doc"
    
    ' bind the data to the document
    wrd_doc.MailMerge.OpenDataSource Name:=s_datafile
    
    If Not b_Print Then
        ' show the new word document
        obj_app.Visible = True
    Else
        'TODO : Print the mail merge and close
        wrd_doc.MailMerge.Destination = 1
        wrd_doc.MailMerge.Execute
        ' Close the word session down.
        wrd_doc.Close
        obj_app.Visible = False
        Set wrd_doc = Nothing
        Set obj_app = Nothing
    End If
    ' clear the word objects
    

produce_mail_merge_exit:

Exit Function
produce_mail_merge_errorhandler:
    Err.Raise Err.Number, "cls_wrd_report_manager.produce_mail_merge", Err.Description
    Resume produce_mail_merge_exit
End Function


Private Sub Class_Initialize()
    Set obj_wrd_routines = New i_word_routines
End Sub


Private Sub Class_Terminate()
    Set obj_wrd_routines = Nothing
End Sub



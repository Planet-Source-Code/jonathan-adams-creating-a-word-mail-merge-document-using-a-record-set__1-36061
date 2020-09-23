VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "i_word_routines"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
    Option Explicit

Const CLASS_NAME = "i_word_routines"



  
Public Function Create_mailmerge_data_file(ByVal wrd_doc As Object, ByVal rs_merge_data As ADODB.Recordset, ByVal s_file_name As String, ByVal bJustHeader As Boolean, ByVal l_doc_style As Long, ByVal s_field_not_to_include As String) As Boolean
    On Error GoTo Create_mailmerge_data_file_Errorhandler
    Dim wrdDataDoc As Object
    Dim col_item  As Object
    Dim lCount As Long
    Dim s_header_record As String
    Dim rs_field As ADODB.Field
    Dim bExport_Complete As Boolean
    Const PROCEDURE_NAME = "Create_mailmerge_data_file"
  
    Create_mailmerge_data_file = False
     
    '*****************************************************
     bJustHeader = False
    '*****************************************************
    Dim TestFile As String
    
    ' if the file exists then delete it.
    If Dir(s_file_name) <> "" Then
        TestFile = s_file_name
        s_file_name = Left(s_file_name, InStr(1, s_file_name, ".dat") - 1) & Format(Now(), "yyyymmddhhnn") & ".dat"
    
    End If
    
 '   pf_create_datasource s_file_name, rs_merge_data, s_field_not_to_include
    

    s_header_record = ""
    For Each rs_field In rs_merge_data.Fields
        If Not rs_field Is Nothing Then
            If InStr(s_field_not_to_include, rs_field.Name) = 0 Then
                If s_header_record = "" Then
                    s_header_record = rs_field.Name
                Else
                    s_header_record = s_header_record & "," & rs_field.Name
                End If
            End If
        End If
    Next rs_field
    
    If l_doc_style > -1 Then
        wrd_doc.MailMerge.MainDocumentType = l_doc_style
    End If
    
  '  Call iobj_word.produce_mail_merge(rs_merge_data, s_report_path, b_Print)
    
    wrd_doc.MailMerge.CreateDataSource Name:=s_file_name, _
             HeaderRecord:=s_header_record


    ' Open the file to insert data
    Set wrdDataDoc = wrd_doc.Parent.Documents.Open(s_file_name)
    
 
    wrdDataDoc.Save
    wrdDataDoc.Close False
    If TestFile <> "" Then
        Kill TestFile
    End If
    Create_mailmerge_data_file = True

Create_mailmerge_data_file_Exit:

Exit Function
Create_mailmerge_data_file_Errorhandler:
    Err.Raise Err.Number, CLASS_NAME & "." & PROCEDURE_NAME, Err.Description
    GoTo Create_mailmerge_data_file_Exit
End Function


Public Function Create_word_document(ByVal s_template_name As String, _
                                        ByVal s_document_name As String, _
                                        ByVal rs_results As ADODB.Recordset, _
                                        ByVal s_data_filename As String, ByVal l_output_option As Long, _
                                        ByVal pb_Close_word As Boolean, _
                                        ByVal s_field_not_to_include As String) _
                                        As Boolean
    
    On Error GoTo Create_word_document_errorhandler
    
    Const PROCEDURE_NAME = "Create_word_document"
    
    Dim wrd_app         As Object
    Dim wrd_document    As Object
    Dim s_tmp_file_name As String
    Dim frmprog         As frm_progress
    
    Create_word_document = False
    If Dir(s_template_name) = "" Then
        GoTo Create_word_document_exit
    End If
    
    s_tmp_file_name = Left(s_document_name, Len(s_document_name) - 4)
    s_tmp_file_name = s_tmp_file_name & "_data.dat"
    
    Set frmprog = New frm_progress
    Load frmprog
    frmprog.HideProgress = True
    frmprog.Update_Progresss "Establishing a Link with Microsoft Word", 0
    frmprog.Show
    DoEvents
    
        
    ' Attempt to open a word application connection
    Set wrd_app = get_word_application_pointer()
    ' check for success
    If Not wrd_app Is Nothing Then
    
        frmprog.Update_Progresss "Creating a new document", 0
        
        ' create a new word document
        Set wrd_document = wrd_app.Documents.Add(Template:=s_template_name)
        
        ' save the template information
        wrd_document.SaveAs FileName:=s_document_name
        
        frmprog.Update_Progresss "Creating mail merge data file", 0
        If Create_mailmerge_data_file(wrd_document, rs_results, s_tmp_file_name, False, -1, s_field_not_to_include) Then
            
            Select Case l_output_option
                Case 0 ' Open Not Print
                    wrd_document.SaveAs FileName:=s_document_name
                    wrd_app.Visible = True
                
                Case 1 ' Print but don't open
                    wrd_app.Visible = True
                    wrd_document.MailMerge.Destination = 1
                    wrd_document.MailMerge.Execute
                    wrd_document.SaveAs FileName:=s_document_name
                    If pb_Close_word Then
                        wrd_app.Visible = False
                    End If
            End Select
        End If
    End If

    Set wrd_app = Nothing
    
    Unload frmprog
    Set frmprog = Nothing
    
    Create_word_document = True

Create_word_document_exit:



Exit Function
Create_word_document_errorhandler:
    Unload frm_progress
    Err.Raise Err.Number, CLASS_NAME & "." & PROCEDURE_NAME, Err.Description
    GoTo Create_word_document_exit
End Function

Public Function Insert_Table_Row(wrd_document As Object, ByVal l_rownumber As Long, rs_data As ADODB.Recordset, ByVal s_field_not_to_include As String) As Boolean
    On Error GoTo Insert_Table_Row_Errorhandler
    Const PROCEDURE_NAME = "Insert_Table_Row"
    Dim rs_field As ADODB.Field
    Dim col_counter As Long
    col_counter = 0
    Insert_Table_Row = False
    
    If Not rs_data.EOF Then
        With wrd_document.Tables(1)
            ' Insert the data into the specific cell
            For Each rs_field In rs_data.Fields
                If InStr(s_field_not_to_include, rs_field.Name) = 0 Then
                    col_counter = col_counter + 1
                    If Not rs_field Is Nothing Then
                        If Not IsNull(rs_field.Value) Then
                            .Cell(l_rownumber, col_counter).Range.InsertAfter rs_field.Value
                        Else
                            .Cell(l_rownumber, col_counter).Range.InsertAfter " "
                        End If
                    End If
                End If
            Next rs_field
            Insert_Table_Row = True
        End With
    End If
    
Insert_Table_Row_Exit:

Exit Function
Insert_Table_Row_Errorhandler:
    Err.Raise Err.Number, CLASS_NAME & "." & PROCEDURE_NAME, Err.Description
    GoTo Insert_Table_Row_Exit
End Function


Public Function Create_word_template(ByVal s_template_name As String, _
                                        ByVal rs_results As ADODB.Recordset, _
                                        ByVal l_doc_style As Long, _
                                        ByVal b_insert_fields As Boolean, _
                                        ByVal s_field_not_to_include As String) As Boolean
    On Error GoTo Create_word_template_errorhandler
    
    Const PROCEDURE_NAME = "Create_word_template"
    
    Dim wrd_app         As Object

    Dim wrd_document    As Object
    Dim s_tmp_file_name As String
    Dim frmprog         As frm_progress
    Set frmprog = Nothing
    Set frmprog = New frm_progress
    Load frmprog
    frmprog.HideProgress = True
    frmprog.Update_Progresss "Establishing a Link with Microsoft Word", 0
    frmprog.Show
    DoEvents
    
    Create_word_template = False
    
    s_tmp_file_name = Left(s_template_name, Len(s_template_name) - 4)
    s_tmp_file_name = s_tmp_file_name & "_data.dat"
   
    ' Attempt to open a word application connection
    Set wrd_app = get_word_application_pointer()
    ' check for success
    If Not wrd_app Is Nothing Then
        
        frmprog.Update_Progresss "Creating a new document", 0
        ' create a new word document
        Set wrd_document = wrd_app.Documents.Add()
        
        ' save the template information
        wrd_document.SaveAs FileName:=s_template_name
        frmprog.Update_Progresss "Creating mail merge data file", 0
        If Create_mailmerge_data_file(wrd_document, rs_results, s_tmp_file_name, True, l_doc_style, s_field_not_to_include) Then
            If b_insert_fields Then
                pf_insert_merge_fields wrd_document, rs_results, s_field_not_to_include
            End If
            wrd_document.SaveAs s_template_name
            wrd_app.Visible = True
        End If
    
    End If

    Set wrd_app = Nothing
    
    Unload frmprog
    Set frmprog = Nothing
    
    Create_word_template = True

Create_word_template_exit:



Exit Function
Create_word_template_errorhandler:
    Err.Raise Err.Number, CLASS_NAME & "." & PROCEDURE_NAME, Err.Description
 
    GoTo Create_word_template_exit
End Function

 
Public Function view_word_document(ByVal s_template_name As String, ByVal rs_results As ADODB.Recordset, _
                                        ByVal l_doc_style As Long, _
                                        ByVal b_insert_fields As Boolean, _
                                        ByVal s_field_not_to_include As String) As Boolean
    On Error GoTo view_word_document_errorhandler
    
    Const PROCEDURE_NAME = "View_word_document"
    
    Dim wrd_app         As Object
    Dim wrd_document    As Object
    Dim s_tmp_file_name As String
    
    view_word_document = False
    
    If Dir(s_template_name) = "" Then
        GoTo view_word_document_exit
    End If
    
    s_tmp_file_name = Left(s_template_name, Len(s_template_name) - 4)
    s_tmp_file_name = s_tmp_file_name & "_data.dat"
    
      'If Dir(s_tmp_file_name) <> "" Then
   '     Kill s_tmp_file_name
    'End If
  
    ' Attempt to open a word application connection
    Set wrd_app = get_word_application_pointer()
    
    ' check for success
    If Not wrd_app Is Nothing Then
        ' create a new word document
        Set wrd_document = wrd_app.Documents.Open(s_template_name)
        If Dir(s_template_name) <> "" Then
            s_tmp_file_name = wrd_document.MailMerge.DataSource.Name
        End If
        If Create_mailmerge_data_file(wrd_document, rs_results, s_tmp_file_name, True, l_doc_style, s_field_not_to_include) Then
            If b_insert_fields Then
                pf_insert_merge_fields wrd_document, rs_results, s_field_not_to_include
            End If
            wrd_document.SaveAs s_template_name
            wrd_app.Visible = True
        End If

        wrd_app.Visible = True
    End If

    Set wrd_app = Nothing
    
    
    view_word_document = True

view_word_document_exit:



Exit Function
view_word_document_errorhandler:
    Resume
    Err.Raise Err.Number, CLASS_NAME & "." & PROCEDURE_NAME, Err.Description
    GoTo view_word_document_exit
End Function





Public Function get_word_application_pointer() As Object
    On Error GoTo get_word_application_pointer_errorhandler
    Dim wrd_app         As Object
    Dim b_got_error     As Boolean
    
    Set wrd_app = GetObject(, "Word.Application")
    
    If b_got_error Then
         Set wrd_app = CreateObject("Word.Application")
    End If

    Set get_word_application_pointer = wrd_app
    

get_word_application_pointer_exit:

Exit Function
get_word_application_pointer_errorhandler:
    b_got_error = True
    Resume Next
End Function


Private Function pf_insert_merge_fields(ByVal obj_document As Object, ByVal rs_data As ADODB.Recordset, ByVal s_field_not_to_include As String) As Boolean
    On Error GoTo pf_insert_merge_fields_errorhandler
    Dim wrdSelection As Object
    Dim wrdMailMerge As Object
    Dim wrdMergeFields As Object
    Dim rs_field            As ADODB.Field
    
    Const PROCEDURE_NAME = "pf_insert_merge_fields"
    pf_insert_merge_fields = False
    
    Set wrdSelection = obj_document.Application.Selection
    Set wrdMailMerge = obj_document.MailMerge

    wrdSelection.ParagraphFormat.Alignment = 0 ' align Paragragh Left
    Set wrdMergeFields = wrdMailMerge.Fields
    For Each rs_field In rs_data.Fields
        If Not rs_field Is Nothing Then
            If InStr(s_field_not_to_include, rs_field.Name) = 0 Then
                wrdMergeFields.Add wrdSelection.Range, rs_field.Name
                wrdSelection.TypeText " "
                wrdSelection.TypeText vbCrLf
            End If
        End If
    Next rs_field

    pf_insert_merge_fields = True
    
pf_insert_merge_fields_exit:

Exit Function
pf_insert_merge_fields_errorhandler:
    Err.Raise Err.Number, CLASS_NAME & "." & PROCEDURE_NAME, Err.Description
    GoTo pf_insert_merge_fields_exit
End Function



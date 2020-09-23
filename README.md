<div align="center">

## Creating a Word Mail merge document using a  Record set


</div>

### Description

The code will allow you to pass a template name and a recordset to ONE routine and this will then create you a Word Mail Merge document based upon the selected template. If no template is found then a blank file is created

Now upated to include ALL code required to run ;)
 
### More Info
 
Usage :-

Dim iobj_word as cls_wrd_report_manager

Set iobj_word = New cls_wrd_report_manager

rs_data =the recordset to output

s_report_path :=Patha and filename of template document

b_Print := True or Fales (Print or Show)

Call iobj_word.produce_mail_merge(rs_data, s_report_path, b_Print)

A Boolean value for success or failure


<span>             |<span>
---                |---
**Submitted On**   |2000-10-25 14:34:48
**By**             |[Jonathan Adams](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/jonathan-adams.md)
**Level**          |Intermediate
**User Rating**    |4.2 (21 globes from 5 users)
**Compatibility**  |VB 6\.0
**Category**       |[Microsoft Office Apps/VBA](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/microsoft-office-apps-vba__1-42.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[Creating\_a1557793112003\.zip](https://github.com/Planet-Source-Code/jonathan-adams-creating-a-word-mail-merge-document-using-a-record-set__1-36061/archive/master.zip)









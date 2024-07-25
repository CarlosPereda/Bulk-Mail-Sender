; ======================================================================================================================
; Author:           CarlosPereda
; Repository:       https://github.com/CarlosPereda/Bulk-Mail-Sender
; Class:            GUI_SendChunkMails()
; Execution-Notes:  Execute directly
; Tested with:      AHK 2.0.11 (x64)
; Tested on:        Win 10 Home (x64)
; License:          Feel free to modify this software and in case of publishing mention the author.
; ======================================================================================================================
; This software is provided 'as-is', without any express or implied warranty.
; In no event will the author be held liable for any damages arising from the use of this software.
; Please consider the law of your region related to spam and avoid missuse of the following program. 
; ======================================================================================================================

#Requires AutoHotkey v2.0
#Include "%A_LineFile%"
#Include Include/Functions.ahk

GUI_SendChunkMails(){
    MyGui := Gui(, "Send Chunk Emails")
    
    MyGui.add("Text", "section", "From:")
    MyGui.add("Edit", "w170 Lowercase vFrom", "") ; vFrom

    MyGui.add("Text", "ys", "cc:")
    MyGui.add("Edit", "w170 vCc")

    MyGui.add("Text", "xs", "Subject:")
    MyGui.add("Edit", "w350 vSubject")
    
    MyGui.add("Text",, "Excel Path:")
    Edit_ExcelPath := MyGui.add("Edit", "section w290 vExcelPath")
    Button_OpenExcel := MyGui.add("Button", "w50 ys", "Open")
    Button_OpenExcel.OnEvent("click", OpenExcelDataBase)

    MyGui.add("Text", "xs", "HTML path:")
    Edit_HTMLPath := MyGui.add("Edit", "section w290 vHTMLPath")
    Button_OpenHTML := MyGui.add("Button", "w50 ys", "Open")
    Button_OpenHTML.OnEvent("click", OpenHTMLFile)

    MyGui.add("Text", "section xs y+12", "Sheet Name:")
    MyGui.add("Edit", "ys-3 vSheetName", "Sheet1")

    MyGui.add("Text", "xs", "Header Row:")
    MyGui.add("Edit", "yp w43 x82 Number vHeaderRow", "1")

    MyGui.add("Text", "section ys x+20", "First Data row:")
    MyGui.add("Edit", "ys-3 w40 Number vFirstDataRow", "2")

    MyGui.add("Text", "section xs", "Last Data row:")
    MyGui.add("Edit", "ys w40 Number vLastDataRow", "11") 

    Button_SendMails := MyGui.add("Button", "ys-30 x+20 w80 h50", "Send Mails:")
    Button_SendMails.OnEvent("click", SendMails)
    MyGui.show()


    OpenExcelDataBase(*){
        ExcelPath := FileSelect(3, A_WorkingDir,, "(*.xlsx)")
        Edit_ExcelPath.value := ExcelPath
    }

    OpenHTMLFile(*){
        HTMLPath := FileSelect(3, A_WorkingDir,, "(*.txt; *.doc; *.html)")
        Edit_HTMLPath.value := HTMLPath
    }

    SendMails(*){
        Saved := MyGui.Submit(0)

        if !Saved.From or !Saved.Subject or !Saved.SheetName or 
            !Saved.HeaderRow or !Saved.FirstDataRow or !Saved.LastDataRow
            return MsgBox("All fields need to be filled.", "Send Chunk Emails", "32")
        else if !FileExist(Saved.ExcelPath)
            return MsgBox("Select an excel file")
        else if !FileExist(Saved.HTMLPath)
            return MsgBox("Select a txt file")
        
        
        run_python_script("PyFiles/SendEmails.py",      ; Relative path to SendEmails.py 
                            Saved.From,         ; sys.argv[1]
                            Saved.cc,           ; sys.argv[2]
                            Saved.Subject,      ; sys.argv[3]
                            Saved.ExcelPath,    ; sys.argv[4]
                            Saved.HTMLPath,     ; sys.argv[5]
                            Saved.SheetName,    ; sys.argv[6]
                            Saved.HeaderRow,    ; sys.argv[7]
                            Saved.FirstDataRow, ; sys.argv[8]
                            Saved.LastDataRow)  ; sys.argv[9]
    }

}

If (A_ScriptFullPath == A_LineFile){
    GUI_SendChunkMails()
}
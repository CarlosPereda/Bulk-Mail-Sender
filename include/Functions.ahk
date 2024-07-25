double_quote_wrap(str){
    return "`"" . str . "`""
}

run_python_script(script, parameters*){
    str := ""
    for item in parameters{
        str := str . double_quote_wrap(item) " "
    }
    run('cmd.exe /k python' " " double_quote_wrap(script) " " str)
}

get_highlighted_text(){
	ClipboardOld := A_Clipboard 
    A_Clipboard  := ""

    Send("{Ctrl down}c{Ctrl up}")
    if !ClipWait(0.08)
        selected := ""
    else
        selected := A_Clipboard

    A_Clipboard := ClipboardOld
	return selected
}
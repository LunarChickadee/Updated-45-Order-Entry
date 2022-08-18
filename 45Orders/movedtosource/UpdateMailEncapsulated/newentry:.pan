 //OG Newentry
 
 ; beep
           ; WindowBox "459 437 858 1094"
            loop
                mailaddress=Con+¶+MAd+¶+City+" "+St+" "+pattern(Zip,"#####")
        superalert "Do you want to enter this one?"+¶+¶+mailaddress,{height=300 width=250 font="Helvetica"
                size=18 color="red" bgcolor="lightgoldenrodyellow" buttons="Yes:60;No:60" }
                if info("dialogtrigger") contains "Yes"
                    «M?»=?(«M?» contains "E" or «M?» contains "U" or «M?» contains "R", "", «M?»)
                    call "enter/e"
                else
                    Next
                    stoploopif info("found")=0
                endif
             while forever


//somewhat updated newentry
select MAd contains place And Zip=vzip
    selectedAddressArray=""
    arrayselectedbuild selectedAddressArray, ¶,"",str(«C#»)+¬+«Con»+¬+«Group»+¬+«MAd»+¬+«City»+¬+«St»
    selectall

find Zip=vzip
    superchoicedialog selectedAddressArray, chosenAddress, 
        {title="Choose the Correct Customer/Address" caption="Click -Other Search- to Search by something else" buttons="ok;other search;cancel"}
    if info("dialogtrigger") contains "ok"
        find exportline() contains chosenAddress
        if info("found")=-1
            message "Found!"
            call "enter/e"
        else
            message "error, repeating search..."
            farcall "45orders","NewSearch/`"
        endif
    endif
    if info("dialogtrigger") contains "search"
        farcall "45orders","NewSearch/`"
    endif
    if info("dialogtrigger") contains "cancel"
    stop
    endif
Num=«C#»

If Num=0
    call "numberNeeded"
endif

window (thisFYear+"orders")
    «C#»=Num
    «C#Text»=str(Num)
    case info("formname") = "ogsinput" and OrderNo>320000 and OrderNo<400000
        if Con≠""
            stop
        endif
    case info("formname") = "seedsinput" and OrderNo>710000
        if Con≠""
            stop
        endif
    case info("formname") = "treesinput" and OrderNo>420000 and OrderNo<500000
        if Con≠""
            stop
        endif
    case info("formname") = "mtinput" and OrderNo>620000 and OrderNo<700000
        if Con≠""
            stop
        endif
    case info("formname") = "bulbsinput" and  OrderNo>520000 and OrderNo<600000
        if Con≠""
            stop
        endif
    endcase

    call ".customerfill"
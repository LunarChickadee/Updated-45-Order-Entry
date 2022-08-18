Num=«C#»
If Num=0
;Message "You must create a Customer#"
                if ono>300000 and ono<400000
                    call "ogsity/ø"
                endif
                if ono>500000 and ono<600000
                    call "bulbous/∫"
                endif
                if ono>400000 and ono<500000
                    call "treed/†"
                endif
                 if ono>700000
                    call "seedy/ß"
                endif
                if ono>600000 and ono<700000
                    call "moosed/µ"
                endif
endif
case waswindow contains "bulbs"
Bf=?(Bf=0,1,Bf)
case waswindow contains "tree"
T=?(T=0,1,T)
case waswindow contains "seed"
S=?(S=0,1,S)
case waswindow contains "ogs"
S=?(S=0,1,S)
case waswindow contains "mt"
S=?(S=0,1,S)
endcase

window (thisyear+"orders")
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
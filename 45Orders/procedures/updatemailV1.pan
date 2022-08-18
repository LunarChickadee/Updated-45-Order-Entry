global selectedAddressArray, chosenAddress, chosenAddress1, FindWindow, WinNumber


waswindow=info("windowname")

;call ".dropship"
if info("formname")="treesinput" and (OrderNo<410000 or OrderNo>420000)
    call ".pool"
endif
EntryDate=today()
ono=OrderNo
vzip=Zip
vd=«C#»
conArray=Con[1," "][1,-2]+" "+Con["- ",-1][2,-1]
place=MAd["0-9",-1][1,2]
field «1stPayment»
;editcell

WinNumber=arraysearch(info("windows"), thisFYear+" mailing list", 1,¶)
if WinNumber=0
    openfile newyear+" mailing list"
endif
window newyear+" mailing list"
if vd>0
    find «C#»=vd
    if info("found")=-1
        YesNo "Enter this one?"
        if clipboard() contains "Yes"
            call "enter/e"
            stop
        endif
    endif

    if info("found")=0
        goto newzip
    endif
    window newyear+" mailing list"
endif

//*********************************************************//
//Searches by Zip
newzip:

find Zip=vzip

//Zip that's not in Mailaing List
if info("found")=0
    find Con contains extract(conArray," ",1)and Con contains extract(conArray," ",2)
    if info("Found")=-1
        goto newentry
    else
        addmail:

        insertrecord
        Con=grabdata(newyear+"orders",Con) 
        Group=grabdata(newyear+"orders",Group) 
        MAd=grabdata(newyear+"orders",MAd) 
        City=grabdata(newyear+"orders",City) 
        St=grabdata(newyear+"orders",St)
        Zip=grabdata(newyear+"orders",Zip) 
        adc=lookup("newadc","Zip3",pattern(Zip,"#####")[1,3],"adc",0,0)
        SAd=grabdata(newyear+"orders",SAd) 
        Cit=grabdata(newyear+"orders",Cit) 
        Sta=grabdata(newyear+"orders",Sta) 
        Z=grabdata(newyear+"orders",Z) 
        phone=grabdata(newyear+"orders",Telephone) 
        email=grabdata(newyear+"orders", Email)
        adc=lookup("fcmadc","Zip3",pattern(Zip,"#####")[1,3],"adc",0,0)
        if ono>300000 and ono<400000
            call "ogsity/ø"
        endif
        if ono>500000 and ono<600000
            call "bulbous/∫"
        endif
        if ono>400000 and ono<500000
            call "treed/†"
        endif
        if ono>700000 and ono < 1000000
            call "seedy/ß"
        endif
        if ono>600000 and ono<700000
            call "moosed/µ"
        endif
        window waswindow
    endif
endif


//Zip is in Mailing List
if info("found")=-1
    if place ≠ ""
        find MAd contains place And Zip=vzip
        if info("found")=0
            find Zip=vzip and Con contains extract(conArray," ",2)
            if info("found")=-1
                goto newentry
            endif
                insertrecord
                Con=grabdata(newyear+"orders",Con) 
                Group=grabdata(newyear+"orders",Group) 
                MAd=grabdata(newyear+"orders",MAd) 
                City=grabdata(newyear+"orders",City) 
                St=grabdata(newyear+"orders",St)
                Zip=grabdata(newyear+"orders",Zip) 
                adc=lookup("newadc","Zip3",pattern(Zip,"#####")[1,3],"adc",0,0)
                SAd=grabdata(newyear+"orders",SAd) 
                Cit=grabdata(newyear+"orders",Cit) 
                Sta=grabdata(newyear+"orders",Sta) 
                Z=grabdata(newyear+"orders",Z) 
                phone=grabdata(newyear+"orders",Telephone) 
                email=grabdata(newyear+"orders", Email)
                if ono>300000 and ono<400000
                call "ogsity/ø"
                endif
                if ono>500000 and ono<600000
                call "bulbous/∫"
                endif
                if ono>400000 and ono<500000
                call "treed/†"
                endif
            if ono>700000 and ono < 1000000
                call "seedy/ß"
                endif
                if ono>600000 and ono<700000
                call "moosed/µ"
                endif
        endif
        if info("found")=-1
           newentry:
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
             goto addmail
        endif
    endif
endif
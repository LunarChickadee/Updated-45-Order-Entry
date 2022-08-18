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
        /*
            OG:
            loops through addresses and tries to find the custopmer

            V2:
            tried to add more data to the search

            V3:
            currently called NewSearch
            */
    else
        call addmail:
        /*
        adds a record and fills it
        then calls filler
        */
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
                call addmail:
            /*
            adds a record and fills it
            then calls filler
            */
        endif
        if info("found")=-1
           call newentry:
            /*
            OG:
            loops through addresses and tries to find the custopmer

            V2:
            tried to add more data

            V3:
            currently called NewSearch
            */
             goto addmail
        endif
    endif
endif
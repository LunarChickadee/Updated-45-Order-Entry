/*
This Function is supposed to either 
1. find the customer already in the list
    1a. Call Enter

or

2. add a new line and fill it with this new data

*/



global vChoice,vFirstInitial,vLastName,vExtracted,vExtracted2,vFirstIntLastName,vCustNum,
vEmail,vPhoneNum, chooseCustomerArray,chooseCustChoice,rayj,MLWildCard,MLWildCard2,orderWildCard,amperArray,thisCustomerLine

//gets first name and last name
//kept for possible dependencies, 
//but made conArray to replace it
//so I know when other things are calling for the Con
rayj=Con[1," "][1,-2]+" "+Con["- ",-1][2,-1] 
conArray=Con[1," "][1,-2]+" "+Con["- ",-1][2,-1]

waswindow=info("windowname")

window thisFYear+"orders"

if info("formname")="treesinput" and (OrderNo<410000 or OrderNo>420000)
    call ".pool"
    endif

/*This fills the variables for the search*/
EntryDate=today()
ono=«OrderNo»
vzip=Zip
vd=«C#»  // #Davids old code, leaving in case there are dependencies
vCustNum=«C#»
vEmail=Email
vPhoneNum=Telephone
vChoice=0
place=MAd  //was originally just the number to the end of the address, but PO boxes broke it
vCustNum=«C#»
vEmail=Email
vPhoneNum=Telephone
vChoice=0
MLWildCard=""
MLWildCard2=""
orderWildCard=""
thisCustomerLine=""


//builds a WildCard of the name (see MATCH in the Pan Ref)
    vExtracted=""
    vExtracted=extract(conArray," ",1)
    vFirstInitial=vExtracted[1,1]
    vExtracted2=""
    vLastName=extract(conArray," ",2)
    orderWildCard=str(vFirstInitial+"*"+vLastName)





///___make sure mailing list is open_____
WinNumber=arraysearch(info("windows"), thisFYear+" mailing list", 1,¶)
if WinNumber=0
    openfile newyear+" mailing list"
endif

///____can we find them just with the C# they have on the order?
window newyear+" mailing list"
selectall
if vd>0
    find «C#»=vd
    if info("found")=-1
        
        thisCustomerLine=arrayrange(exportline(),1,7,¬)+" "+array(exportline(),16,¬)
        YesNo "Enter this one?"+¶+thisCustomerLine
        if clipboard() contains "Yes"
            call "enter/e"
            stop
            endif
    else
        ////____If we can't let's do a smart search_____///
        goto Choose
        endif
    endif


Choose:

global ChoiceList, Choice, Options
Choice=""

ChoiceList="First Initial & Last Name
By Email
Phone Number
Mailing Address
My Own Search 
Add New Customer Instead
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Please tell Lunar if you want another option here"
superchoicedialog ChoiceList, Choice, {caption="How Would You Like to Search?
You can click cancel to stop this process" captionstyle=bold captionheight="2" height="250" }
if info("dialogtrigger") contains "cancel"
Stop
endif
;message Choice

vChoice=arraysearch(ChoiceList, Choice, 1, ¶)

case vChoice=1
    window newyear+" mailing list"
    //try to find the first initial and last name
    //example J*Forester would be the wildcard, and would match names like Jane Forester
    select Con match orderWildCard
        //if it doesn't find it, check the & names
        if info("empty")
            select Con contains "&"
                selectwithin strip(array(Con,1,"&")) match orderWildCard // name left of & matches (Jane Forester & Kathrine Kidd)
                    or strip(array(Con,2,"&")) match orderWildCard // name right of & matches (Katherine Kidd & Jane Forester)
                    or strip(array(Con,1,"&"))+" "+extract(strip(array(Con,2,"&"))," ",2)["- ",-1] match orderWildCard 
                    //name to the left plus the last name to the right matches (Jane & Katherine Forester)
                        if info("empty")
                            message "Unable to find anyone with that name. Please, choose another option"
                                goto Choose
                                    endif
                                        endif

case vChoice=2
    window newyear+" mailing list"
    select email=vEmail
    if info("empty")
        message "No email match found."
        goto Choose
    endif
case vChoice=3
    window newyear+" mailing list"
    select phone=vPhoneNum
        if info("empty")
            message "No phone match found."
            goto Choose
        endif
case vChoice=4
    window newyear+" mailing list"
    select MAd=place and Zip=vzip
        if info("empty")
            message "No address match found."
            goto Choose
        endif
case vChoice=5
    window newyear+" mailing list"
    findselect
    if info("dialogtrigger") contains "cancel"
    goto Choose
        endif
    else 
        stop
case vChoice=6
    window newyear+" mailing list"
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
endcase


window thisFYear+" mailing list"

arrayselectedbuild chooseCustomerArray,¶,"",exportline()
        superchoicedialog chooseCustomerArray,chooseCustChoice, {
        captionstyle=bold
        caption="Please choose the appropriate customer or click other search to try something else.
        You can click cancel to look through the selection at any time"
        buttons=OK;OtherSearch:100;Cancel
        captionheight="2" 
        height="600" width="800"}
            if info("dialogtrigger") contains "cancel"
                stop
                    endif
            if info("dialogtrigger") contains "other"
                goto Choose
                    endif
            if info("dialogtrigger") contains "ok"
                find exportline() contains chooseCustChoice
                if info("found")=-1
                    thisCustomerLine=arrayrange(exportline(),1,7,¬)+" "+array(exportline(),16,¬)
                    YesNo "Is this is the customer you want?"+¶+thisCustomerLine
                    if clipboard() contains "Yes"
                        call "enter/e"
                        stop
                        endif
                    endif
                endif


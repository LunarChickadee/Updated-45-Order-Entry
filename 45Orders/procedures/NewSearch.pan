global vChoice,vFirstInitial,vLastName,vExtracted,vExtracted2,vFirstIntLastName,vCustNum,
vEmail,vPhoneNum, chooseCustomerArray,chooseCustChoice,rayj

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
rayj=Con[1," "][1,-2]+" "+Con["- ",-1][2,-1] //gets first name and last name
//kept rayj for dependencies, but made a new array named clearer for my use
conArray=Con[1," "][1,-2]+" "+Con["- ",-1][2,-1]
;place=MAd["0-9",-1][1,-1]
place=MAd
vExtracted=""
vExtracted=extract(rayj," ",1)
vFirstInitial=vExtracted[1,1]
vExtracted2=""
vLastName=extract(rayj," ",2)
vCustNum=«C#»
vEmail=Email
vPhoneNum=Telephone
vChoice=0
debug
WinNumber=arraysearch(info("windows"), thisFYear+" mailing list", 1,¶)
if WinNumber=0
    openfile newyear+" mailing list"
endif

window newyear+" mailing list"
selectall
if vd>0
    find «C#»=vd
    if info("found")=-1
        YesNo "Enter this one?"
        if clipboard() contains "Yes"
            call "enter/e"
            stop
            endif
    else
        goto Choose
        endif
    endif
;selectall

/*

supergettext vChoice,“caption="Please Choose How You'd like to Search:
1=Email 2=Phone Number 3=Mailing Address 4=Add New Customer 
Use the Command + Tilde '~' key to run this again!" captionheight=3”
*/
Choose:
global ChoiceList, Choice, Options
Choice=""

ChoiceList="By Email
Phone Number
Mailing Address
My Own Search
Add New Customer Instead"
superchoicedialog ChoiceList, Choice, {caption="How Would You Like to Search?" captionstyle=bold}
if info("dialogtrigger") contains "cancel"
Stop
endif
;message Choice

vChoice=arraysearch(ChoiceList, Choice, 1, ¶)

case vChoice=1
    window newyear+" mailing list"
    select email=vEmail
    if info("empty")
        message "No email match found."
        farcall (thisFYear+"orders"),"NewSearch/`"
    endif
case vChoice=2
    window newyear+" mailing list"
    select phone=vPhoneNum
        if info("empty")
            message "No phone match found."
            farcall (thisFYear+"orders"),"NewSearch/`"
        endif
case vChoice=3
    window newyear+" mailing list"
    select MAd=place and Zip=vzip
        if info("empty")
            message "No address match found."
            farcall (thisFYear+"orders"),"NewSearch/`"
        endif
case vChoice=4
    window newyear+" mailing list"
    findselect
    if info("dialogtrigger") contains "cancel"
    farcall (thisFYear+"orders"),"NewSearch/`"
        endif
case vChoice=5
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
        superchoicedialog chooseCustomerArray,chooseCustChoice, {caption="Please choose the appropriate customer or click other search to try sometning else."
        buttons=OK;OtherSearch:100;Cancel height="600" width="800"}
            if info("dialogtrigger") contains "cancel"
                stop
                    endif
            if info("dialogtrigger") contains "other"
                farcall (thisFYear+"orders"),"NewSearch/`"
                    endif
            if info("dialogtrigger") contains "ok"
                find exportline() contains chooseCustChoice
                    endif


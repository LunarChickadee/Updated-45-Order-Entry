global WinCheck, intOrder1, fromBranch

intOrder1=int(OrderNo)
fromBranch=""
WinCheck=info("windows")


/*
added this because I don't see a a reason to not ensure the user is
on the right entry form
also holds onto what branch this order is to check against for later
-L 8/22
*/

//_____CheckForm_________//
case intOrder1 ≥ 700000
    openform "seedsinput"
    fromBranch="Seeds"
case intOrder1 ≥ 600000 and intOrder1 < 700000
    openform "mtinput"
    fromBranch="OGS;POE"
case intOrder1 ≥ 500000 and intOrder1 < 600000
    openform "bulbsinput"
    fromBranch="OGS;Bulbs"
case intOrder1 ≥ 400000 and intOrder1 < 500000
    openform "treesinput"
    fromBranch="Trees"
case intOrder1 ≥ 300000 and intOrder1 <400000
    openform "ogsinput"
    fromBranch="OGS"
endcase
//____EndCheckForm______//

groupArray=?(Group≠"", Group, Con)
;rayc=«C#Text»
if OrderNo=int(OrderNo)
    rayg=?(MemDisc=round(Subtotal*.01,.01),"Y","")
endif
;if OrderNo ≠val(str(OrderNo)[1,5])
;MemDisc=?(rayg="Y",.01,MemDisc)
;downrecord
;stop
;endif
;if rayg=""
;message "rayg not set"
;stop
;endif
if info("formname")="treesinput"
    rayi=Pool
endif
ono=OrderNo
downrecord


oono=0
if info("formname")="addresschecker" 
    stop
endif
field «C#»
if OrderNo≠int(OrderNo)
    Group=groupArray
    ;«C#Text»=rayc
    MemDisc=?(rayg="Y",.01,MemDisc)
    if info("formname")="treesinput"
        Pool=rayi
    endif
    field Con
    ;call ".dropship"
endif
case info("windowname")=newyear+"orders:seedsinput"
    If OrderNo<700000 or OrderNo>1000000
        beep
        message "Order not in system yet!"
        stop
    endif
case info("windowname")=newyear+"orders:ogsinput"
    If OrderNo<300000 or OrderNo > 400000
        beep
        message "Order not in system yet!"
        stop
    endif
    ;;call .rollingdiscount

case info("windowname")=newyear+"orders:treesinput"
    If OrderNo<400000 or OrderNo>500000
        message "Order not in system yet!"
        beep
        stop
    endif
case info("windowname")=newyear+"orders:bulbsinput"
    If OrderNo<500000 or OrderNo≥600000
        message "Order not in system yet!"
        beep
        stop
    endif
case info("windowname")=newyear+"orders:mtinput"
    If OrderNo<600000 or OrderNo≥700000
        message "Order not in system yet!"
        beep
        stop
    endif
endcase

if str(OrderNo) contains "."
    EntryDate = today()
    stop
else
    EntryDate=today()
endif

;; selecting just internet orders

if OrderNo=int(OrderNo)
    if (OrderNo >= 320000 and OrderNo < 400000)
    or (OrderNo >= 420000 and OrderNo < 500000)
    or (OrderNo >= 520000 and OrderNo < 600000)
    or (OrderNo >= 620000 and OrderNo < 700000)
    or (OrderNo >= 710000 and OrderNo < 1000000)
            call "updatemail/7"
    endif
endif

//!!!!!! This is missing an argument. Why is it here?
drawobjects
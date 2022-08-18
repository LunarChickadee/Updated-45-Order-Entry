if «C#»≠0
    message "This Customer already has a Number"
    stop 
endif

//gives them an First Year and Branch of Order code
case fromBranch contains "Seeds"
Code = "I"+thisYear+"s"
case fromBranch contains "OGS"
code = "I"+thisYear+"o"
case fromBranch contains "Trees"
Code = "I"+thisYear+"t"

//gets a new number
openfile "Customer#"
call "newnumber"

//_______________________//
window "45 mailing list"
Field «C#»
    Paste

//why?
SpareText2=str(«C#»)

If inqcode=""
    Field inqcode
    inqcode=?(inqcode contains "17", inqcode[3,-1], inqcode)
    EditCell
    field «C#»
EndIf

call "filler/¬"

If inqcode=""
field inqcode
editcell
endif
endif

//_______________________________//
window "customer_history:secret"
opensheet
insertbelow
«C#»=grabdata("45 mailing list", «C#»)
Group=grabdata("45 mailing list", Group)
Con=grabdata("45 mailing list", Con)
MAd=grabdata("45 mailing list", MAd)
City=grabdata("45 mailing list", City)
St=grabdata("45 mailing list", St)
Zip=grabdata("45 mailing list", Zip)
Email=grabdata("45 mailing list", email)
SpareText2=grabdata("45 mailing list", SpareText2)
;CloseWindow

//______________________________//
window "45 mailing list"
Call "enter/e"
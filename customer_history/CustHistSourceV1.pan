___ PROCEDURE findname/5 _______________________________________________________
local firstname, lastname, clist, mlist
hide
noshow
firstrecord
loop
firstname=""
lastname=""
clist=0
mlist=0
firstname=Con[1," "][1,-2]
lastname=Con[" ",-1][2,-1]
clist=«C#»
window "37 mailing list"
find Con contains firstname and Con contains lastname
if info("found")=0
mlist=0
else
mlist=«C#»
endif
window "searchlist"
insertbelow
mailinglist=mlist
custhistory=clist
Con=firstname+" "+lastname
window "customer_history"
downrecord
until info("stopped")
show
___ ENDPROCEDURE findname/5 ____________________________________________________

___ PROCEDURE ccrider/ç ________________________________________________________
waswindow=info("windowname")
serverlookup "off"
NoUndo
GetScrap "enter the customer number"
Find «C#» = val(clipboard())
if info("found")=0
beep
endif
if info("files") contains "customer_history"
window "customer_history:secret"
Find «C#» = val(clipboard())
window waswindow
endif
;field «C#»
;field MAd
serverlookup "on"
___ ENDPROCEDURE ccrider/ç _____________________________________________________

___ PROCEDURE mailinglistlookup/1 ______________________________________________
custno=«C#»
window "41 mailing list"
find «C#»=custno
___ ENDPROCEDURE mailinglistlookup/1 ___________________________________________

___ PROCEDURE loopdown/2 _______________________________________________________
field «36Total»
Hide
loop
;copycell
;pastecell
downrecord
until 1000
Show
___ ENDPROCEDURE loopdown/2 ____________________________________________________

___ PROCEDURE fix money/3 ______________________________________________________
Field S31
loop
total
lastrecord
copy
find «C#»=cusno1
paste
right
stoploopif val(info("fieldname")[-2,-1])=val(str(«C#»)[1,2])-1
until forever
Field Bf31
loop
total
lastrecord
copy
find «C#»=cusno1
paste
right
stoploopif val(info("fieldname")[-2,-1])=val(str(«C#»)[1,2])-1
until forever
Field M31
loop
total
lastrecord
copy
find «C#»=cusno1
paste
right
stoploopif val(info("fieldname")[-2,-1])=val(str(«C#»)[1,2])-1
until forever
Field OGS31
loop
total
lastrecord
copy
find «C#»=cusno1
paste
right
stoploopif val(info("fieldname")[-2,-1])=val(str(«C#»)[1,2])-1
until forever
Field T31
loop
total
lastrecord
copy
find «C#»=cusno1
paste
right
stoploopif val(info("fieldname")[-2,-1])=val(str(«C#»)[1,2])-1
until forever
RemoveSummaries 7
find «C#»≠cusno1
cusno2=«C#»
deleterecord
window "33 mailing list"
find «C#»=cusno2
deleterecord
___ ENDPROCEDURE fix money/3 ___________________________________________________

___ PROCEDURE findorder/4 ______________________________________________________
getscrap "What order? (Please use all 5 digits)"
openfile "37orders"
ono=val(clipboard())
find OrderNo=ono
case (ono≥10000 and ono<30000) or (ono>60000 and ono<100000) or (ono>800000 and ono<1000000)
GoForm "seedsinput"
case (ono>30000 and ono<40000) or (ono>300000 and ono<400000)
goform "ogsinput"
case ono>40000 and ono<50000
goform "treesinput"
case ono>50000 and ono<60000
goform "bulbsinput"
case ono>70000 and ono<80000
goform "mtinput"
endcase

___ ENDPROCEDURE findorder/4 ___________________________________________________

___ PROCEDURE forceunlock ______________________________________________________
forceunlockrecord
___ ENDPROCEDURE forceunlock ___________________________________________________

___ PROCEDURE .Initialize ______________________________________________________
global custno
custno=0
;GoSheet
;Field "MAd"
;SortUp
;Field "City"
;SortUp
;Field "St"
;SortUp
;Field "Zip"
;SortUp
;windowtoback "customer_history"


___ ENDPROCEDURE .Initialize ___________________________________________________

___ PROCEDURE filladdress ______________________________________________________
select MAd=""
if info("selected")=info("records")
message "all uptodate"
stop
endif
field MAd
formulafill lookup("38 mailing list","C#",«C#»,"MAd","",0)
field City
formulafill lookup("38 mailing list","C#",«C#»,"City","",0)
field St
formulafill lookup("38 mailing list","C#",«C#»,"St","",0)
field Zip
formulafill lookup("38 mailing list","C#",«C#»,"Zip",0,0)
___ ENDPROCEDURE filladdress ___________________________________________________

___ PROCEDURE .find ____________________________________________________________
custno=«C#»
window "37 mailing list"
find «C#»=custno
___ ENDPROCEDURE .find _________________________________________________________

___ PROCEDURE sortup ___________________________________________________________
field MAd
sortup
field City
sortup
field St
sortup
field Zip
sortup
___ ENDPROCEDURE sortup ________________________________________________________

___ PROCEDURE consolidate ______________________________________________________
local mlist, clist
global dialogPause
window "searchlist"
loop
mlist=mailinglist
clist=custhistory
window "customer_history:customer history"
select «C#»=mlist
selectadditional «C#»=clist
find «C#»=clist
cancelok "Delete This record"
if clipboard()="OK"
deleterecord
endif
window "searchlist"
downrecord
until info("stopped")
___ ENDPROCEDURE consolidate ___________________________________________________

___ PROCEDURE delete ___________________________________________________________
lastrecord
loop
deleterecord
until info("selected")=1
field «C#»
copy
selectall
find «C#»=clipboard()
___ ENDPROCEDURE delete ________________________________________________________

___ PROCEDURE DeleteRecord _____________________________________________________
deleterecord
window "searchlist"
downrecord
call "consolidate"

___ ENDPROCEDURE DeleteRecord __________________________________________________

___ PROCEDURE check address ____________________________________________________
field MAd
select MAd notmatch  lookup("37 mailing list","C#",«C#»,"MAd","",0)
___ ENDPROCEDURE check address _________________________________________________

___ PROCEDURE fill info ________________________________________________________
forcesynchronize
window "45 mailing list"
call "forcesynchronize"
window "customer_history"
select Zip=0 and length(St)=2
if info("selected")=info("records")
beep
stop
endif
field Zip
formulafill lookup("45 mailing list", "C#",«C#», "Zip",0,0)
select MAd=""
field MAd
formulafill lookup("45 mailing list", "C#",«C#», "MAd","",0)
select City=""
field City
formulafill lookup("45 mailing list", "C#",«C#», "City","",0)
select St=""
field St
formulafill lookup("45 mailing list", "C#",«C#», "St","",0)
selectall
call "sortup"
select MAd=""
___ ENDPROCEDURE fill info _____________________________________________________

___ PROCEDURE forcesynchronize _________________________________________________
forcesynchronize
call "sortup"
___ ENDPROCEDURE forcesynchronize ______________________________________________

___ PROCEDURE checkit __________________________________________________________
local checktotal
checktotal=0
firstrecord
loop
checktotal=checktotal+val(«Gets Check»)
downrecord
until info("stopped")
message str(checktotal)
___ ENDPROCEDURE checkit _______________________________________________________

___ PROCEDURE fixaddress _______________________________________________________
field MAd
select MAd≠lookup("37 mailing list", "C#",«C#», "MAd","",0)
formulafill lookup("37 mailing list", "C#",«C#», "MAd","",0)
field City
formulafill lookup("37 mailing list", "C#",«C#», "City","",0)
field St
formulafill lookup("37 mailing list", "C#",«C#», "St","",0)
field Zip
formulafill lookup("37 mailing list", "C#",«C#», "Zip",0,0)
selectall
call "sortup"
___ ENDPROCEDURE fixaddress ____________________________________________________

___ PROCEDURE fix zipcode ______________________________________________________
select Zip=0 and length(St)=2
field Zip
formulafill lookup("37 mailing list","C#",«C#»,"Zip",0,0)
___ ENDPROCEDURE fix zipcode ___________________________________________________

___ PROCEDURE selectcustomers __________________________________________________
getscrap "Which division"
case clipboard() contains "bulbs"
select «Bf33» >0 or «Bf34» >0 or «Bf35» >0 or «Bf36» >0 
selectwithin «C#»>0
case clipboard() contains "trees"
select «T33» >0 or «T34» >0 or «T35» >0  or «T36» >0
selectwithin «C#»>0
case clipboard() contains "seeds"
select «S33» >0 or «S34» >0 or «S35» >0 or «S36» >0
selectadditional «MT33» >0 or «MT34» >0 or «MT35» >0 or «MT36» >0
selectadditional «OGS33» >0 or «OGS34» >0 or «OGS35» >0 or «OGS36» >0
selectwithin «C#»>0
endcase
window "37 mailing list"
field «C#»
select «C#»=lookupselected("customer_history","C#",«C#»,"C#",0,0) and «C#»>0
___ ENDPROCEDURE selectcustomers _______________________________________________

___ PROCEDURE close window _____________________________________________________
SelectAll
save
CloseFile
___ ENDPROCEDURE close window __________________________________________________

___ PROCEDURE huh ______________________________________________________________
local hasBranchInfo
/* 
added 8/22 by Lunar
*/



window thisFYear+" mailing list"

if S+T+Bf=0 and RedFlag=""
    yesno "- Customer has no catalogs requested"+¶+"- Customer has no RedFlag(s)"+¶+¶+"Autofill catalog requests by Zip/Order?"
    if clipboard()="Yes"
        Case Zip < 19000  And Zip>1000
            S=1
            «M?»=?(«M?» notcontains "X","X"+«M?»,«M?»)
            T=1
            «M?»=?(«M?» notcontains "W","W"+«M?»,«M?»)
            Bf=1
            «M?»=?(«M?» notcontains "Z","Z"+«M?»,«M?»)
        Case (Zip > 43000 And Zip < 46000) 
        or (Zip > 48000 And Zip < 50000) 
        or (Zip > 53000 And Zip < 57000) 
        or Zip>97000
            S=1
            «M?»=?(«M?» notcontains "X","X"+«M?»,«M?»)
            T=1
            «M?»=?(«M?» notcontains "W","W"+«M?»,«M?»)
            ///add an if ono or fromBranch=bulbs change this
            Bf=0
            «M?»=?(«M?» contains "Z",replace(«M?»,"Z",""),«M?»)
        DefaultCase
            S=1
            «M?»=?(«M?» notcontains "X","X"+«M?»,«M?»)
            T=0
            //same for trees and bulbs here
            «M?»=?(«M?» contains "W",replace(«M?»,"W",""),«M?»)
            Bf=0
            «M?»=?(«M?» contains "Z",replace(«M?»,"Z",""),«M?»)
        endcase     
    endif 
else 
    case RedFlag≠""
        message "Customer has a RedFlag."+¶+"Catalog requests will be set to zero"
            S=0
            T=0
            Bf=0
            «M?»=""
    defaultcase 
    noyes "Update Catalog Requests?"
    +¶+
    "Currently, Customer is set to receive"
    +¶+
    "Seeds:"+str(S)+" Bulbs:"+str(Bf)+" Trees:"+str(T)
    
    //make this smart enough to only say whaty they're getting?
        if clipboard()="Yes"

        ///this loop is from .UpdateCats
            loop
                rundialog
                “Form="CatalogRequest"
                    Movable=yes
                    okbutton=Update
                    Menus=normal
                    WindowTitle={CatalogRequest}
                    Height=264 Width=190
                    AutoEdit="Text Editor"
                    Variable:"val(«dS»)=val(«S»)"
                    Variable:"val(«dBf»)=val(«Bf»)"
                    Variable:"val(«dT»)=val(«T»)"”
                stoploopif info("trigger")="Dialog.Close"
            while forever 
              message "Customer is now set to receive"
                        +¶+
                        "Seeds:"+str(S)+" Bulbs:"+str(Bf)+" Trees:"+str(T)
                if S≥1 and «M?» notcontains "X"
                    «M?»="X"+«M?»
                else 
                    if S=0
                    «M?»=?(«M?» contains "X",replace(«M?»,"X",""),«M?»)
                    endif
                endif

                if T≥1 and «M?» notcontains "W"
                    «M?»="W"+«M?»
                else 
                    if T=0
                    «M?»=?(«M?» contains "W",replace(«M?»,"W",""),«M?»)
                    endif
                endif

                if Bf≥1 and «M?» notcontains "Z"
                    «M?»="Z"+«M?»
                else 
                    if Bf=0
                    «M?»=?(«M?» contains "Z",replace(«M?»,"Z",""),«M?»)
                    endif
                endif
        endif
    endcase
endif 


___ ENDPROCEDURE huh ___________________________________________________________

___ PROCEDURE (CommonFunctions) ________________________________________________

___ ENDPROCEDURE (CommonFunctions) _____________________________________________

___ PROCEDURE ExportMacros _____________________________________________________
local Dictionary1, ProcedureList
//this saves your procedures into a variable
exportallprocedures "", Dictionary1
clipboard()=Dictionary1

message "Macros are saved to your clipboard!"
___ ENDPROCEDURE ExportMacros __________________________________________________

___ PROCEDURE ImportMacros _____________________________________________________
local Dictionary1,Dictionary2, ProcedureList
Dictionary1=""
Dictionary1=clipboard()
yesno "Press yes to import all macros from clipboard"
if clipboard()="No"
stop
endif
//step one
importdictprocedures Dictionary1, Dictionary2
//changes the easy to read macros into a panorama readable file

 
//step 2
//this lets you load your changes back in from an editor and put them in
//copy your changed full procedure list back to your clipboard
//now comment out from step one to step 2
//run the procedure one step at a time to load the new list on your clipboard back in
//Dictionary2=clipboard()
loadallprocedures Dictionary2,ProcedureList
message ProcedureList //messages which procedures got changed

___ ENDPROCEDURE ImportMacros __________________________________________________

___ PROCEDURE Symbol Reference _________________________________________________
bigmessage "Option+7= ¶  [in some functions use chr(13)
Option+= ≠ [not equal to]
Option+\= « || Option+Shift+\= » [chevron]
Option+L= ¬ [tab]
Option+Z= Ω [lineitem or Omega]
Option+V= √ [checkmark]
Option+M= µ [nano]
Option+<or>= ≤or≥ [than or equal to]"


___ ENDPROCEDURE Symbol Reference ______________________________________________

___ PROCEDURE GetDBInfo ________________________________________________________
local DBChoice, vAnswer1, vClipHold

Message "This Procedure will give you the names of Fields, procedures, etc in the Database"
//The spaces are to make it look nicer on the text box
DBChoice="fields
forms
procedures
permanent
folder
level
autosave
fileglobals
filevariables
fieldtypes
records
selected
changes"
superchoicedialog DBChoice,vAnswer1,“caption="What Info Would You Like?"
captionheight=1”


vClipHold=dbinfo(vAnswer1,"")
bigmessage "Your clipboard now has the name(s) of "+str(vAnswer1)+"(s)"+¶+
"Preview: "+¶+str(vClipHold)
Clipboard()=vClipHold

___ ENDPROCEDURE GetDBInfo _____________________________________________________

___ PROCEDURE .linearray _______________________________________________________
displaydata info("trigger")
___ ENDPROCEDURE .linearray ____________________________________________________

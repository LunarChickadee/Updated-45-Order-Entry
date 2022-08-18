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
            Bf=?(fromBranch contains "OGS",1,0)
            «M?»=?(«M?» contains "Z",replace(«M?»,"Z",""),«M?»)
        DefaultCase
            S=1
            «M?»=?(«M?» notcontains "X","X"+«M?»,«M?»)
            T=?(fromBranch contains "Trees",1,0)
            //same for trees and bulbs here
            «M?»=?(«M?» contains "W",replace(«M?»,"W",""),«M?»)
            Bf=?(fromBranch contains "OGS",1,0)
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
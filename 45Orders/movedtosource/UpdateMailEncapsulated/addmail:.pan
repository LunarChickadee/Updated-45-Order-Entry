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

//Filler?______
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

window thisFYear+"orders"
import os
import xlrd
import sys

loc = 'STEEL2020.xls'
wb = xlrd.open_workbook(loc)
sheet = wb.sheet_by_index(0)
print('Choices are ALUM, STAINLESS STEEL,GALV.STEEL,STEEL,STEEL TUBING,STEEL CHANNEL,STEEL ANGLES,FLAT BAR,PIPE STEEL,SOLID STOCK,MISCELLANEOUS ')
question = input('What type of metal? ' )
question = question.lower()
#question = question.strip()
##### MAKE IT TO WHERE IT JUST ASK ALUM , GALV AND STAIN STEEL,
if  question == 'alum':
    alum = question

    def Aluminum(alum):
        ALUM = {}
        for i in range(3,62):
    #print(sheet.cell_value(i,0))
            ALUM[str(sheet.cell_value(i,0)+' '+sheet.cell_value(i,1))] = sheet.cell_value(i,4)
        for i in ALUM:
            print(i)
#print(ALUM.keys())
        print('')
        key = input('What type of Aluminum? ')
        print('')
 #   print('This metal costs ' + str(ALUM[key]))
        print('')
        return print('This metal costs ' + str(ALUM[key]))
    Aluminum(alum)
#print(ALUM['.063 ALUM DIAMOND PLATE 4X8'])
if question == 'stainless steel':
    stainlesssteel= question
    def Stainlesssteel(stainlesssteel):
        Stain = {}
        expsteel = {}
        for i in range(63,97):#
            if str(sheet.cell_value(i,0)) == 'EXPANDED STEEL':
                for i in range(97,105):
                    expsteel[str(sheet.cell_value(i,0)+' '+sheet.cell_value(i,1))] = sheet.cell_value(i,4)
            else:
                Stain[str(sheet.cell_value(i,0)+' '+sheet.cell_value(i,1))] = sheet.cell_value(i,4)
        aa = input('Is the metal Expanded Steel? Answer: yes or no: ')
        if aa == 'yes':
            for i in expsteel:
                print(i)
            print('')
            key = input('Copy paste selection '+ 'What type of expanded Steel: ' )
            print('')
            ans = expsteel[key]
        else:
            for i in Stain:
                print(i)
            print('')
            key = input('Copy paste selection '+ 'What type of Stainless Steel: ' )
            print('')
            ans = Stain[key]

       # print(expsteel)

        return print('The cost is ' + str(ans))
    Stainlesssteel(stainlesssteel)
if question == 'galv.steel':
    galv = question
    def GALV(galv):
        galv = {}
        for i in range(105,115):
    #print(sheet.cell_value(i,0))
            galv[str(sheet.cell_value(i,0)+' '+sheet.cell_value(i,1))] = sheet.cell_value(i,4)
        for i in galv:
            print(i)
#print(ALUM.keys())
        print('')
        key = input('What type of galv steel? ')
        print('')
        ans = galv[key]
 #   print('This metal costs ' + str(ALUM[key]))
        print('')
        return print('The cost is ' + str(ans))
    GALV(galv)
    if question == 'steel':
        steel = question

        def Steel(steel):
            steel = {}
            for i in range(119,172):
                steel[str(sheet.cell_value(i,0)+' '+sheet.cell_value(i,1))] = sheet.cell_value(i,4)
            for i in steel:
                print(i)
            print('')
            key = input('What type of steel? ')
            print('')
            ans = steel[key]
 #   print('This metal costs ' + str(ALUM[key]))
            print('')
            return print('The cost is ' + str(ans))
        Steel(steel)

if question == 'steel tubing':
    tub = question

    def Steeltubing(tub):
        tub = {}
        for i in range(174,251):
            print(sheet.cell_value(i,0))
            tub[str(sheet.cell_value(i,0)+' '+sheet.cell_value(i,1))] = sheet.cell_value(i,4)
        for i in tub:
            print(i)
        print('')
        key = input('What type of tubing? ')
        print('')
        ans = tub[key]
 
        print('')

        return print('The cost is ' + str(ans))
    Steeltubing(tub)

if question == 'steel channel':
    chan = question
    def Steelchannel(chan):
        chan = {}
        beam = {}
        for i in range(252,273):
            chan[str(sheet.cell_value(i,0)+' '+sheet.cell_value(i,1))] = sheet.cell_value(i,4)
            if str(sheet.cell_value(i,0)) == 'STEEL I-BEAM, H-BEAM':
                for i in range(273,296):
                    beam[str(sheet.cell_value(i,0)+' '+sheet.cell_value(i,1))] = sheet.cell_value(i,4)
        ll = input('Is the steal channel STEEL I-BEAM, H-BEAM? ')
        if ll == 'yes':
            for i in beam:
                print(i)
            print('')
            key = input('What type of Steel I-Beam or H-beam? ')
            print('')
            ans = beam[key]
        else:
            for i in chan:
                print(i)
            print(' ')
            key = input('What type of channel? ')     
            ans =   chan[key]



    
        return print('The cost is ' + str(ans))    
    Steelchannel(chan)

if question == 'steel angles':
    ang = question
    def steelangles(ang):
        ang = {}
        for i in range(298,333):
            ang[str(sheet.cell_value(i,0)+' '+sheet.cell_value(i,1))] = sheet.cell_value(i,4)
        for i in ang:
            print(i)
        print('')
        key = input('What type of Steel Angle? ')
        print('')
        ans = ang[key]
 
        print('')  

        return print('The cost is ' + str(ans))  
    steelangles(ang)

if question == 'flat bar':
    flat = question
    def flatbar(flat):
        flat = {}
        for i in range(333,402):
            flat[str(sheet.cell_value(i,0)+' '+sheet.cell_value(i,1))] = sheet.cell_value(i,4)
        for i in flat:
            print(i)
        print('')
        key = input('What type of Steel Angle? ')
        print('')
        ans = flat[key]
 
        print('')  


        return print('The cost is ' + str(ans))  
    flatbar(flat)

if question == 'pipe steel':
    pipe = question
    def pipesteel(pipe):
        pipe = {}
        for i in range(403,422):
            pipe[str(sheet.cell_value(i,0)+' '+sheet.cell_value(i,1))] = sheet.cell_value(i,4)
        for i in pipe:
            print(i)
        print('')
        key = input('What type of Pipe Steel? ')
        print('')
        ans = pipe[key]
 
        print('')  


        return print('The cost is ' + str(ans))  
    pipesteel(pipe)   

if question == 'solid stock':
    x = input('is it a cold roll: answer yes or no ')
    x = x.lower()
    if x == 'yes':
        roll = x
        def Coldroll(roll):
            cold = {}
            for i in range(424,438):
                cold[str(sheet.cell_value(i,0)+' '+sheet.cell_value(i,1))] = sheet.cell_value(i,4)
            for i in cold:
                print(i)
            print('')
            key = input('What type of Cold Roll? ')
            print('')
            ans = cold[key]



            return print('The cost is ' + str(ans)) 
        Coldroll(roll)
    else:
        roll = x
        def Hotroll(roll):
            hot = {}
            for i in range(439,460):
                hot[str(sheet.cell_value(i,0)+' '+sheet.cell_value(i,1))] = sheet.cell_value(i,4)
            for i in hot:
                print(i)
            print('')
            key = input('What type of Cold Roll? ')
            print('')
            ans = hot[key]



            return print('The cost is ' + str(ans))   
        Hotroll(roll)     

if question == 'miscellaneous':
    mis = question
    def Mis(mis):
        mis = {}
        for i in range(461,497):
            mis[str(sheet.cell_value(i,0)+' '+sheet.cell_value(i,1))] = sheet.cell_value(i,4)
        for i in mis:
            print(i)
        print('')
        key = input('What item? ')
        print('')
        ans = mis[key]
        return print('The cost is ' + str(ans))  
    Mis(mis) 








    










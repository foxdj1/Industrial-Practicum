
for j in range(3):
    for i in range(5):
        if i == 3:
            break
        print(i)


        if count == 2:
            if z==1:
                continue
            for row in range(2, wrkbk3[sheet].max_row + 1):
                row1 = comp.max_row + 1
                comp.cell(row=row1, column=1, value=wrkbk3[sheet].cell(row=row, column=1).value)
                comp.cell(row=row1, column=2, value=wrkbk3[sheet].cell(row=row, column=2).value)
                comp.cell(row=row1, column=3, value=wrkbk3[sheet].cell(row=row, column=3).value)
                wrkbk3.save(path + 'Target_' + on + '_Avoid_' + off + '.xlsx')
                z=1
                continue




                for sheet in wrkbk3.sheetnames:
                    if wrkbk3[sheet].cell(row=2, column=1).value != "":
                        len1 = len1 + 1
                for row in range(2, comp.max_row + 1):
                    comp.cell(row=row, column=2).value = comp.cell(row=row, column=2).value / len1
                wrkbk3.save(path + 'Target_' + on + '_Avoid_' + off + '.xlsx')








    #make a list of all the drugs that bind to the intended target
    good_druglist = []
    for sheet in wrkbk3.sheetnames:
        drug = sheet.split('_', 3)
        drug = drug[4]
        drug = drug.strip('.xlsx')
        good_druglist = good_druglist.append(drug)

    # make a list of all the drugs that bind to the unintended target
    bad_druglist = []
    for sheet in wrkbk5.sheetnames:
        drug = sheet.split('_', 3)
        drug = drug[4]
        drug = drug.strip('.xlsx')
        bad_druglist = bad_druglist.append(drug)

    #make a list of the drugs that bind to both the inteneded and unintended targets
    elim_drugs = []
    for x,y in good_druglist, bad_druglist:
        if x==y:
            elim_drugs.append(x)

    #remove drugs that bind to both the intended and unintended targets from the intended sheet
    for sheet in wrkbk3.sheetnames:
        for x in elim_drugs:
            if elim_drugs[x] in sheet:
                wrkbk3.remove_sheet(sheet)\

    #remove drugs that bind to both the intended and unintended targets from the unintended sheet
    for sheet in wrkbk5.sheetnames:
        for x in elim_drugs:
            if elim_drugs[x] in sheet:
                wrkbk5.remove_sheet(sheet)





  #divide the frequency value for each interaction by the number of drugs
    for row in range(2, comp.max_row + 1):
        comp.cell(row=row, column=2).value = comp.cell(row=row, column=2).value / len1
    wrkbk3.save(path + 'Target_' + on + '_Avoid_' + off + '.xlsx')

    # divide the frequency value for each interaction by the number of drugs
    for row in range(2, comp.max_row + 1):
        comp.cell(row=row, column=2).value = comp.cell(row=row, column=2).value / len1
    wrkbk5.save(path + 'Avoid_' + off + '_Target_' + on + '.xlsx')













        important_r.append(wb.cell(row=row, column=2).value)
        apolar.append(wb.cell(row=row, column=3).value)

    if len(important_r) >= 3:
        nonp = nsmallest(3, important_r).index
    elif len(important_r) == 2:
        nonp = nsmallest(2, important_r).index
    else:
        nonp = important_r.index

    print(nonp)
    if len(apolar) >= 3:
        p = nsmallest(3, apolar).index
    elif len(apolar) == 2:
        p = nsmallest(2, apolar).index
    else:
        p = apolar.index

    nonp_drug = []
    p_drug = []
    for x in nonp:
        row = x+2
        nonp_drug.append(wb.cell(row=row, column=1).value)

    for x in p:
        row = x+2
        p_drug.append(wb.cell(row=row, column=1).value)

    final_drugs = []
    for x in nonp_drug:
        for y in p_drug:
            if x==y:
                final_drugs.append(x)





excel_data= []
    for row in wb.iter_rows(min_row=2, values_only=True):
        excel_data.append(row)
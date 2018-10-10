import xlrd                                                   
                                                              
start = ['08:00', '09:50', '11:40', '14:00', '15:45', '17:30']
end = ['09:35', '11:25', '13:15', '15:35', '17:20', '19:05']  
                                                              
book = xlrd.open_workbook('imi2018.xls')                      
it4 = book.sheet_by_index(8)                                  
                                                              
for i in range(3,39):                                         
    if it4.cell(i, 8).value == "":                            
        continue                                              
    para = it4.cell(i, 8).value                               
    l_pr = it4.cell(i, 9).value                               
    room = it4.cell(i, 10).value                              
    print(start[(i-3)%6], end[(i-3)%6],    para, l_pr, room)  
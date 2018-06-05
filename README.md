# Excel-comparing    
这个程式是为了方便比对客户的BOM先后版本的差异格式为Excel    
因为客户BOM有coverpage，所以实际比对的是第二个Sheet，可以在下边的语句中调整

        sheet1=wb1.get_sheet_by_name(sheets1[1])
        sheet2=wb2.get_sheet_by_name(sheets2[0])
        sheet3=wb3.get_sheet_by_name(sheets3[1])
        sheet4=wb4.get_sheet_by_name(sheets4[0])

使用Tinker做UI,作为初学者和公司环境的特殊性，暂时只能用这个    

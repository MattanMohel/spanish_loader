from docx import Document

ex = """"""

MONTHS = [
     ('enero',      31), 
     ('febrero',    28), 
     ('marzo',      31), 
     ('abril',      30), 
     ('mayo',       31), 
     ('junio',      30), 
     ('julio',      31), 
     ('agosto',     31), 
     ('septiembre', 30), 
     ('octubre',    31), 
     ('noviembre',  30), 
     ('diciembre',  31)
    ]

# accepts a date formatted by 'month/day'
def date_to_id(date_str):
    date = date_str.split('/')
    
    try:
        month = int(date[0])
        day   = int(date[1])
    except:
        return None, 'incorrect input format'
    
    if month > 12 or month < 1:
        return None, 'incorrect month input'
    
    info = MONTHS[month - 1]
    
    if day > info[1] or day < 1:
        return None, f'{info[0]} dates restricted to 1-{info[1]}'
    
    # return month-name/day
    return info[0], day


def search_calendar(month, day):  
    # load calendar stored in .docx format
    doc = Document('docx/calendar.docx')
    
    cur_month = None
    cur_day   = None
    
    for table in doc.tables:   
        for row in table.rows:
            for cell in row.cells:          
                # assign current month if on month cell
                if cell.text in (month[0] for month in MONTHS):
                    cur_month = cell.text
                
                num = "" 
                for ch in cell.text:
                    if ch.isdigit():
                        num += ch
                    else:
                        break
                
                if num.isnumeric():
                    cur_day = int(num)  
                    
                    if cur_day == day and cur_month == month:
                        return month, day, (par.text for par in cell.paragraphs) 
                    
    return 'cell not found'                                                        

def get_date_info(date):
    month_id, day_id = date_to_id(date)
    month, day, cell = search_calendar(month_id, day_id)  
    
    res = []
    for txt in cell:
        res.append(f"\t{txt}")
        
    return f"{month} {day}:\n", res
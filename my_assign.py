import pandas as pd
from tabulate import tabulate
from fpdf import FPDF

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle
document = SimpleDocTemplate("table.pdf", pagesize=A4)

data = pd.read_excel("Dummy Data .xlsx")
df = pd.DataFrame(data,columns=['Registration Number'])
sets = set()
ll = 0
mm=[]
for i in range(1,len(data)):
    sets.add(data.loc[i][5])
    if(len(sets)!=ll):
        ll+=1
        mm.append(data.loc[i][5])
    
#print(sets)
carts=[]
sets = list(sets)
sets=mm
total_scores=[]
print(sets)
#print(len(sets))
for i in range(len(sets)):
    #head = ["Question No.","what you marked","Correct Answer","Outcome","Score if correct","Your score"]
    res_no = sets[i]
    cart = [["Question No.","what you marked","Correct Answer","Outcome","Score if correct","Your score"]]
    total_score = 0
    for j in range(1,len(data)):
        
        if res_no == data.loc[j][5]:#df.loc[j]['Registration Number']:
           # print(data.loc[j]['First Name '])
            ss = []
            ss.append(data.loc[j][13])
            if data.loc[j][14] != data.loc[j][14]:
                ss.append('')
            else:
                ss.append(data.loc[j][14])
            ss.append(data.loc[j][15])
          #  print(data.loc[j]['What you marked'])
            if data.loc[j][14] ==data.loc[j][15]:
                ss.append("Correct")
                ss.append(str(data.loc[j][17]))
                ss.append(str(data.loc[j][17]))
                total_score += data.loc[j][17]
            elif data.loc[j][14] != data.loc[j][14]:
                ss.append("Unattemp")
                ss.append(str(data.loc[j][17]))
                ss.append(str(0))
            else:
                ss.append("Incorrect")
                ss.append(str(data.loc[j][17]))
                ss.append(str(0))
            cart.append(ss)
    carts.append(cart)
    total_scores.append(total_score)
   # print(tabulate(cart))
   # print(total_score)


#from fpdf import FPDF


def create_table(table_data, title='', data_size = 10, title_size=12, align_data='L', align_header='L', cell_width='even', x_start='x_default',emphasize_data=[], emphasize_style=None, emphasize_color=(0,0,0)):
    """
    table_data: 
                list of lists with first element being list of headers
    title: 
                (Optional) title of table (optional)
    data_size: 
                the font size of table data
    title_size: 
                the font size fo the title of the table
    align_data: 
                align table data
                L = left align
                C = center align
                R = right align
    align_header: 
                align table data
                L = left align
                C = center align
                R = right align
    cell_width: 
                even: evenly distribute cell/column width
                uneven: base cell size on lenght of cell/column items
                int: int value for width of each cell/column
                list of ints: list equal to number of columns with the widht of each cell / column
    x_start: 
                where the left edge of table should start
    emphasize_data:  
                which data elements are to be emphasized - pass as list 
                emphasize_style: the font style you want emphaized data to take
                emphasize_color: emphasize color (if other than black) 
    
    """
    default_style = pdf.font_style
    if emphasize_style == None:
        emphasize_style = default_style
    # default_font = pdf.font_family
    # default_size = pdf.font_size_pt
    # default_style = pdf.font_style
    # default_color = pdf.color # This does not work

    # Get Width of Columns
    def get_col_widths():
        col_width = cell_width
        if col_width == 'even':
            col_width = pdf.epw / len(data[0]) - 1  # distribute content evenly   # epw = effective page width (width of page not including margins)
        elif col_width == 'uneven':
            col_widths = []

            # searching through columns for largest sized cell (not rows but cols)
            for col in range(len(table_data[0])): # for every row
                longest = 0 
                for row in range(len(table_data)):
                    cell_value = str(table_data[row][col])
                    value_length = pdf.get_string_width(cell_value)
                    if value_length > longest:
                        longest = value_length
                col_widths.append(longest + 4) # add 4 for padding
            col_width = col_widths



                    ### compare columns 

        elif isinstance(cell_width, list):
            col_width = cell_width  # TODO: convert all items in list to int        
        else:
            # TODO: Add try catch
            col_width = int(col_width)
        return col_width

    # Convert dict to lol
    # Why? because i built it with lol first and added dict func after
    # Is there performance differences?
    if isinstance(table_data, dict):
        header = [key for key in table_data]
        data = []
        for key in table_data:
            value = table_data[key]
            data.append(value)
        # need to zip so data is in correct format (first, second, third --> not first, first, first)
        data = [list(a) for a in zip(*data)]

    else:
        header = table_data[0]
        data = table_data[1:]

    line_height = pdf.font_size * 2.5

    col_width = get_col_widths()
    pdf.set_font(size=title_size)

    # Get starting position of x
    # Determin width of table to get x starting point for centred table
    if x_start == 'C':
        table_width = 0
        if isinstance(col_width, list):
            for width in col_width:
                table_width += width
        else: # need to multiply cell width by number of cells to get table width 
            table_width = col_width * len(table_data[0])
        # Get x start by subtracting table width from pdf width and divide by 2 (margins)
        margin_width = pdf.w - table_width
        # TODO: Check if table_width is larger than pdf width

        center_table = margin_width / 2 # only want width of left margin not both
        x_start = center_table
        pdf.set_x(x_start)
    elif isinstance(x_start, int):
        pdf.set_x(x_start)
    elif x_start == 'x_default':
        x_start = pdf.set_x(pdf.l_margin)


    # TABLE CREATION #

    # add title
    if title != '':
        pdf.multi_cell(0, line_height, title, border=0, align='j', ln=3, max_line_height=pdf.font_size)
        pdf.ln(line_height) # move cursor back to the left margin

    pdf.set_font(size=data_size)
    # add header
    y1 = pdf.get_y()
    if x_start:
        x_left = x_start
    else:
        x_left = pdf.get_x()
    x_right = pdf.epw + x_left
    if  not isinstance(col_width, list):
        if x_start:
            pdf.set_x(x_start)
        for datum in header:
            pdf.multi_cell(col_width, line_height, datum, border=0, align=align_header, ln=3, max_line_height=pdf.font_size)
            x_right = pdf.get_x()
        pdf.ln(line_height) # move cursor back to the left margin
        y2 = pdf.get_y()
        pdf.line(x_left,y1,x_right,y1)
        pdf.line(x_left,y2,x_right,y2)

        for row in data:
            if x_start: # not sure if I need this
                pdf.set_x(x_start)
            for datum in row:
                if datum in emphasize_data:
                    pdf.set_text_color(*emphasize_color)
                    pdf.set_font(style=emphasize_style)
                    pdf.multi_cell(col_width, line_height, datum, border=0, align=align_data, ln=3, max_line_height=pdf.font_size)
                    pdf.set_text_color(0,0,0)
                    pdf.set_font(style=default_style)
                else:
                    pdf.multi_cell(col_width, line_height, datum, border=0, align=align_data, ln=3, max_line_height=pdf.font_size) # ln = 3 - move cursor to right with same vertical offset # this uses an object named pdf
            pdf.ln(line_height) # move cursor back to the left margin
    
    else:
        if x_start:
            pdf.set_x(x_start)
        for i in range(len(header)):
            datum = header[i]
            pdf.multi_cell(col_width[i], line_height, datum, border=0, align=align_header, ln=3, max_line_height=pdf.font_size)
            x_right = pdf.get_x()
        pdf.ln(line_height) # move cursor back to the left margin
        y2 = pdf.get_y()
        pdf.line(x_left,y1,x_right,y1)
        pdf.line(x_left,y2,x_right,y2)


        for i in range(len(data)):
            if x_start:
                pdf.set_x(x_start)
            row = data[i]
            for i in range(len(row)):
                datum = row[i]
                if not isinstance(datum, str):
                    datum = str(datum)
                adjusted_col_width = col_width[i]
                if datum in emphasize_data:
                    pdf.set_text_color(*emphasize_color)
                    pdf.set_font(style=emphasize_style)
                    pdf.multi_cell(adjusted_col_width, line_height, datum, border=0, align=align_data, ln=3, max_line_height=pdf.font_size)
                    pdf.set_text_color(0,0,0)
                    pdf.set_font(style=default_style)
                else:
                    pdf.multi_cell(adjusted_col_width, line_height, datum, border=0, align=align_data, ln=3, max_line_height=pdf.font_size) # ln = 3 - move cursor to right with same vertical offset # this uses an object named pdf
            pdf.ln(line_height) # move cursor back to the left margin
    y3 = pdf.get_y()
    pdf.line(x_left,y3,x_right,y3)
#print(data)
#print(df)
#print(carts)
print(len(sets))
pdf = FPDF()
pdf.add_page()
pdf.set_font("Times", size=10)
for i in range(len(sets)):
    res_no = sets[i]
    for j in range(1,len(data)):
        if res_no == data.loc[j][5]:
            gender=data.loc[j][8]
            name=data.loc[j][4]
            city=data.loc[j][10]+','+data.loc[j][12]
            school_name=data.loc[j][7]
            test_date=data.loc[j][11]
            final_result = data.loc[j][19]
            rounds = data.loc[j][1]
    pdf.image("Pics for assignment/"+name+".png",x=None,y=None,w=20,h=20,type='PNG')
    
    pdf.image("images.png",x=180,y=10,w=20,h=20,type='PNG')
    pdf.cell(80,10,txt = "Registration Number"+str(res_no),ln=0)
    pdf.cell(80,10,txt = "Gender: "+gender,ln=1)
    pdf.cell(80,10,txt = "Full Name: "+name,ln=0)
    pdf.cell(80,10,txt = "City of Residence: "+city,ln=1)
    pdf.cell(80,10,txt = "Name of School : "+school_name,ln=0)
    pdf.cell(80,10,txt = "Date and time of test: "+test_date,ln=0)
    
    pdf.cell(50,10,ln=1)
    create_table(table_data = carts[i],title='PROGRESS REPORT OF QUALITY EXAM', cell_width='even')
    pdf.ln()
    pdf.cell(50,10,txt="Round : "+str(rounds),ln=1)
    pdf.cell(50,10,txt = "Your Total Score : "+str(total_scores[i]),ln=1)
    pdf.cell(150,10,txt="Final result: "+final_result,ln=1)
    pdf.cell(100,200,ln=1)



# create_table(table_data = data_as_dict,align_header='R', align_data='R', cell_width=[15,15,10,45,], x_start='C') 


pdf.output('table_function.pdf')



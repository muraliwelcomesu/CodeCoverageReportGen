import os,shutil,openpyxl,cx_Oracle,sys,traceback
from urllib.request import urlopen
from urllib.error import HTTPError
from urllib.error import URLError
import urllib.request
from bs4 import BeautifulSoup
from datetime import datetime
import CodeCoverageUtils as Globals
import schedule
import time
cx_Oracle.init_oracle_client(lib_dir=r"C:\OracleInstantClient\instantclient_19_9")
def get_curr_time():
    now = datetime.now()
    dt_string = now.strftime("%d/%m/%Y %H:%M:%S")
    return dt_string

def get_Code_Coverge_Details():
    print('Start of get_Code_Coverge_Details')
    l_dict = {}
    l_retlist = []
    l_total_percent = 0
    #dir_name = os.getcwd()
    #all_subdirs = [d for d in os.listdir('.') if os.path.isdir(d)]
    #latest_subdir = max(all_subdirs, key=os.path.getmtime)
    #print(latest_subdir)
    #os.chdir(os.path.join(dir_name,latest_subdir))
    #os.listdir()
    #fp = open("C:\\ChakraTeam-Share\\Murali\\file_name.html",'a+')
    for root_dir  in os.listdir():
        dir_name = os.path.join(os.getcwd(),root_dir)
        for file_name in os.listdir(dir_name):
            #print('{} -{}'.format(dir_name,file_name))
            if file_name == 'index.html':
                #new_file_name = root_dir + '_index.html'
                #print(new_file_name)
                #shutil.copy(os.path.join(dir_name,file_name),os.path.join('C:\\ChakraTeam-Share\\Murali',new_file_name)) 
                f = open(os.path.join(dir_name,file_name)).read()
                bsobj = BeautifulSoup(f,"html.parser")
                bsTbl = bsobj.findAll("table",{"class":"coverage"})
                for table in bsTbl:
                    table_rows = table.find_all('tr')
                    for tr in table_rows:
                        td = tr.find_all('td')
                        row = [i.text for i in td]
                        if row[0] == 'Total':
                            l_dict['dir_name'] = root_dir
                            l_dict['Item'] = row[0]
                            l_dict['Value'] = row[2]
                            try:
                                l_total_percent = l_total_percent + int(row[2].strip('%'))
                            except:
                                print('Some err in arriving total {} {} '.format(l_total_percent,row[2]))
                            
                            l_retlist.append(l_dict)
                            l_dict = {}
                '''for result in bsTbl:
                    fp.write(str(result))
                fp.write('<br> <br>')'''        
    #fp.close()

    if len(l_retlist) > 0:
        l_avg = round(l_total_percent/(len(l_retlist)),2)
        l_dict['dir_name'] = 'Total Execution %'
        l_dict['Item'] = 'Total'
        l_dict['Value'] = l_total_percent
        l_retlist.append(l_dict)
        l_dict = {}
        l_dict['dir_name'] = 'Average Execution %'
        l_dict['Item'] = 'Total'
        l_dict['Value'] = l_avg
        l_retlist.append(l_dict)
    print('Completed get_Code_Coverge_Details')               
    return l_retlist
    
    
def write_to_excel(l_list_coverage):
    print('Start of write_to_excel')
    OutputName = Globals.OutputExcelName
    if  os.path.isfile(OutputName):
        os.unlink(OutputName)
        
    wb1 = openpyxl.Workbook()
    sheet = wb1.create_sheet()
    sheet.title = Globals.Output_SheetName
    sheet['A1'] = 'Service'
    sheet['B1'] = 'Percentage of Code Covered'
    l_rownum = 1
    for item  in l_list_coverage:
        l_rownum += 1
        sheet['A' + str(l_rownum)].value = item['dir_name']
        sheet['B' + str(l_rownum)].value = item['Value']
    wb1.remove(wb1['Sheet'])
    wb1.save(OutputName) 
    print('Return from write_to_excel')   

def build_header(p_dict_out,p_header):
    
    l_html_str = '<span style = "color:black">Hi,<br> {} <br> <br> \n'.format(p_header)
    p_dict_out['Header'] = l_html_str

def build_footer(p_dict_out,p_name):
    l_html_str = '<span style = "color:black">Thanks and Regards<br> {} <br><br>\n'.format(p_name)
    p_dict_out['Footer'] = l_html_str

def write_dict_htlmfile(p_dict):
    os.chdir(Globals.Work_dir_path)
    fp = open(Globals.HtmlFileName,'w')
    for key,value in p_dict.items():
        fp.write(value + '\n')
    fp.close()
 
def conv_Excel_Dict(ExcelName,SheetName):
    wb = openpyxl.load_workbook(ExcelName,data_only=True)
    sheet = wb[SheetName]
    lst_row = []
    dict_sheet = {}
    l_rownum = 0
    for row in sheet.rows:
        l_rownum = l_rownum + 1
        l_col_cnt = 0
        l_reqdcols = [1,2]
        for cell in row:
            l_col_cnt = l_col_cnt + 1
            if l_col_cnt in l_reqdcols:
                if cell.value is None:
                    cell.value = ' '
                lst_row.append(cell.value)
        dict_sheet[l_rownum] = lst_row
        lst_row = []
    return dict_sheet

def Conv_Dict_HTMLDict(p_dict,Title,p_dict_out):
    l_html_str = ''
    if Title in p_dict_out:
        l_html_str = l_html_str + p_dict_out[Title]
    else:
        l_html_str = l_html_str + '<span style = "color:black"><b>{}</b><br>\n'.format(Title)
    l_html_str = l_html_str + '<br>'
    l_html_str = l_html_str + '<html><table border = 1>\n'
    #----------------------------------------------------------------------------------
    for key,values in p_dict.items():
        if str(key)=='1':
            if Title not in p_dict_out:
                #print('first row')
                l_col_str = ''
                for value in values:
                    l_col_str = l_col_str + '<td><b><span style="color:black" >{}</b></td>\n'.format(value)
                #print(l_col_str)
                l_html_str = l_html_str + '<tr style="background-color:red">{}</tr>\n'.format(l_col_str)
        else:
            l_col_str = ''
            if values[-1] == 'Y':
                for value in values:
                    l_col_str = l_col_str + '<td align = "left"><b><span style = "color:black">{}</b></td>\n'.format(value)
            else:
                l_color = 'black'
                #if values[-1] == 'SUCCESS' and values[-2] == 'FAILED':
                if  values[-1] == 'FAILED':
                    l_color = 'red'
                else:
                    l_color = 'black'
                 
                for value in values:
                    l_col_str = l_col_str + '<td align = "left"><span style = "color:{}">{}</td>\n'.format(l_color,value)
            #print('else part ' + l_col_str)
            l_html_str = l_html_str +  '<tr style="background-color:white">{}</tr>\n'.format(l_col_str)
  #----------------------------------------------------------------------------------
    l_html_str = l_html_str + '</table></html><br>\n'
    p_dict_out[Title] = l_html_str
    
def conv_Excel_html_cons_all():
    print('starting preparing html file')
    l_html_dict = {}
    build_header(l_html_dict,'Please find the Code Coverage Report for TD Regression Execution')
    dict_sheet = conv_Excel_Dict(Globals.OutputExcelName,Globals.Output_SheetName)
    Conv_Dict_HTMLDict(dict_sheet,Globals.Output_SheetName,l_html_dict)
    build_footer(l_html_dict,'Muralidharan R')
    write_dict_htlmfile(l_html_dict)
    print('Completed preparing html file')

def pr_sendMail_Plsql(p_sender,dir_name,html_file,p_recipents,p_subject):
    try:
        procedure = Globals.proc_send_mails
        fp  = open(os.path.join(dir_name,html_file))
        p_html_body = fp.read()
        fp.close()
        print('Start of pr_sendMail_Plsql')
        print('p_recipents',p_recipents)
        print('p_subject',p_subject)
        print('p_html_body length ',str(len(p_html_body)))
        p_attchment_path = '*'
        l_conn_str = Globals.translateMessage(Globals.Key,Globals.conn_str,'D')
        connection1 = cx_Oracle.connect(l_conn_str)
        cur1 = connection1.cursor()
        cur1.callproc(procedure,(p_sender,p_recipents, p_subject,p_html_body, p_attchment_path))
        cur1.close()
        connection1.close()
        print('completed pr_sendMail_Plsql')
    except:
        print('********* Failed in pr_sendMail_Plsql  **********')
        print('Unexpected error : {0}'.format(sys.exc_info()[0]))
        traceback.print_exc()

def Generate_Report():
    print('Starting')
    dir_name = Globals.dir_name 
    os.chdir(dir_name)
    l_list_coverage = get_Code_Coverge_Details()
    write_to_excel(l_list_coverage)
    conv_Excel_html_cons_all()
    pr_sendMail_Plsql(Globals.From,Globals.Work_dir_path,Globals.HtmlFileName,Globals.To_List,Globals.subject)
    print('Completed')
    


def Schedule_Report():
    print('\n \n \n \n')
    print('Code Coverage Report Generation will start at 09:30 AM .......')
    print('Please check with Murali before closing this...')
    Generate_Report()
    #def job():
    #    Generate_Report()
    #schedule.every().day.at("09:30").do(job)

    #while True:
    #    schedule.run_pending()
    #    time.sleep(60)
        

if __name__ == "__main__":
    #Generate_Report()
    Schedule_Report()
    

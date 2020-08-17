import pyodbc
import xlrd
from datetime import date
import os


def fetch_products(file_path):
    wb = xlrd.open_workbook(file_path)
    sheet = wb.sheet_by_index(0)
    cols = sheet.ncols
    rows = sheet.nrows
    pillars = []
    for col in range(0, cols):
        for row in range(0, rows):
            cell_value = sheet.cell_value(row, col)
            if cell_value == 'X' or cell_value == 'x':
                key = sheet.cell_value(0, col)
                value = sheet.cell_value(row, 0)
                pillars.append((key, value))
    pillars_products = {}
    for pillar, product in pillars:
        if pillar not in pillars_products:
            pillars_products[pillar] = ["'"+product+"'"]
        else:
            pillars_products[pillar].append("'"+product+"'")
    return pillars_products


def write_sql_files(pillars_products):
    query=""
    with open("SimulatorQuery.sql","w+") as f1:
        with open("EmulatorQuery.sql",'w') as f:
            for category,products in pillars_products.items():
                query += "select Name from JobDefinitions where CategoryPillar={0} AND (CategoryLevel like ('L4%') " \
                         "OR CategoryLevel='Memleak') and CategoryPlatform='Simulator' " \
                         "and IncludeInMetrics=1 and BuildMachineType like ('%25s%') " \
                         "and CategoryProduct IN ({1}) \nUNION\n".format("'" + category + "'", ','.join(products))
            f1.write(query.rstrip("\nUNION\n"))
            f1.write('order by Name')
            f1.seek(0)
            out = f1.read()
            out_emu = out.replace('Simulator', 'Emulator')
            f.write(out_emu)


def execute_query():
    conn = pyodbc.connect("Driver={SQL Server};"
                          "Server=hubvmps.psr.rd.hpicorp.net;"
                          "Database=VMProvisioning;"
                          "UID=vmprovisioning_admin;"
                          "PWD=vmprovisioning_admin")
    with open("Emulator.txt", 'w') as emu:
        return_job_names(conn,'EmulatorQuery.sql',emu)
    split_files("Emulator.txt",3)
    with open("Simulator.txt", 'w') as sim:
        return_job_names(conn,'SimulatorQuery.sql',sim)
    split_files("Simulator.txt",3)


def split_files(split_file,number_of_files):
    with open(split_file) as infp:
        output_files = [open('%s%d.txt' % (split_file.split('.')[0], i), 'w') for i in range(1,number_of_files+1)]
        for i, line in enumerate(infp):
            output_files[i % number_of_files].write(line)
        for f in output_files:
            f.close()


def result_file():
    files = ['Simulator1.txt','Emulator1.txt','Simulator2.txt','Emulator2.txt','Simulator3.txt','Emulator3.txt']
    with open('result.bat', 'wb') as result:
        for file in files:
            with open(file) as f:
                out = f.read()
                result.write(out.encode('utf_8'))
            if file not in files[-1]:
                if "Simulator" in file:
                    result.write(("SLEEP 600"+"\n").encode('utf_8'))
                else:
                    result.write(("SLEEP 86400"+"\n").encode('utf_8'))


def return_job_names(conn,sql_file,file_object):
    skip_products_emu = ['Agate','Fairmont','Kobe','Onyx']
    with open(sql_file) as f:
        sql = f.read()
        cursor = conn.cursor()
        cursor.execute(sql)
        today = date.today()
        d1 = today.strftime("%d%b%Y")
        prefix = 'HP.Tools.VMProvisioning.TriggerJob.exe -job:'
        suffix = ' -note:MajorRun'+d1
        for result in cursor.fetchall():
            job_name = ''.join(result)
            product = job_name.split('-')[0]
            if "Emu" in sql_file and product not in skip_products_emu:
                print(job_name)
                file_object.write(prefix+job_name+suffix)
                file_object.write("\n")
            if "Sim" in sql_file:
                file_object.write(prefix + job_name + suffix)
                file_object.write("\n")

        cursor.close()


pill_prods = fetch_products('C:\\Users\\ChakkaRa\\Desktop\\Pillars_Products.xlsx')
if os.path.exists("./OutputFiles"):
    os.chdir("./OutputFiles")
else:
    os.mkdir("./OutputFiles")
    os.chdir("./OutputFiles")

write_sql_files(pill_prods)
execute_query()
result_file()








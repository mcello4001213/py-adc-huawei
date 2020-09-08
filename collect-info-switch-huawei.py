import os,csv,re,xlwt
from xlwt import easyxf

#========================================================#
CPU = re.compile(r'CPU Usage            : *')
MEMORY = re.compile(r' Memory Using Percentage Is: *')
UPTIME = re.compile(r'(.*)Routing Switch uptime is(.*)')
VERSION = re.compile(r'(.*)software, Version(.*)')
SN = re.compile(r'BarCode=*')
#========================================================#
wb = xlwt.Workbook()
style_percent = easyxf(num_format_str='0%')
worksheet = wb.add_sheet('output')
worksheet.write_merge(0, 0, 1, 1, 'AVG CPU')
worksheet.write_merge(0, 0, 2, 2, 'MAX CPU')
worksheet.write_merge(0, 0, 3, 3, 'MEMORY USAGE')
worksheet.write_merge(0, 0, 4, 4, 'UPTIME')
worksheet.write_merge(0, 0, 5, 5, 'VERSION')
worksheet.write_merge(0, 0, 6, 6, 'SN')
#========================================================#
i=1;
#========================================================#
for filename in os.listdir('Source'):
    j = 1
    worksheet.write_merge(i, i, 0, 0, filename.replace('.txt',''))
    path = os.path.join('Source', filename)
    # ========================================================#
    data_raw = list(open(path).readlines())
    data_line = open(path).read()

    #========================================================#
    CPU_ = list(filter(CPU.match, data_raw))
    MEMORY_ = list(filter(MEMORY.match, data_raw))
    VERSION_ = list(filter(VERSION.match, data_raw))
    SN_ = list(filter(SN.match, data_raw))
    #========================================================#
    CPU_ = [w.replace('CPU Usage            : ', '') for w in CPU_]
    cpu_str=CPU_[0].replace('%','').replace(' ','').strip()
    list_cpu = cpu_str.split('Max:')
    avg_cpu = float(list_cpu[0]) / 100
    max_cpu = float(list_cpu[1]) / 100
    # ========================================================#
    worksheet.write_merge(i, i, j, j, avg_cpu,style_percent)
    j=j+1
    worksheet.write_merge(i, i, j, j, max_cpu,style_percent)
    j=j+1
    # ========================================================#
    MEMORY_ = [w.replace(' Memory Using Percentage Is: ', '') for w in MEMORY_]
    memory_percentage_str=MEMORY_[0].replace('\n','').strip()
    memory_percentage_str=MEMORY_[0].replace('%','').strip()
    memory_percentage=float(memory_percentage_str)/100
    worksheet.write_merge(i, i, j, j, memory_percentage,style_percent)
    j = j + 1
    # ========================================================#
    UPTIME_ = UPTIME.search(data_line)
    uptime=UPTIME_.group(2).strip()
    worksheet.write_merge(i, i, j, j, uptime)
    j = j + 1
    # ========================================================#
    VERSION_= [w.replace('VRP (R) software, Version ', '') for w in VERSION_]
    version=VERSION_[0].strip()
    worksheet.write_merge(i, i, j, j, version)
    j = j + 1
    # ========================================================#
    SN_ = [w.replace('BarCode=', '') for w in SN_]
    serialnumber=SN_[0].strip()
    worksheet.write_merge(i, i, j, j, serialnumber)

    print(filename)
    print('avg cpu = ',avg_cpu)
    print('max cpu = ',max_cpu)
    print('memory usage = ',memory_percentage)
    print('uptime = '+uptime)
    print('version = '+version)
    print('sn = '+serialnumber)
    print('\n')
    i+=1
wb.save('result.xls')
### author : Agustinus Marcello ###

### hasil nya ==>> result.xls ###
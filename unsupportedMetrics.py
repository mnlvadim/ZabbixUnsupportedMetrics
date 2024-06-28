import xlsxwriter
import time
from pyzabbix import ZabbixAPI
from config import *
#zapi = ZabbixAPI(zabbix_server, user=zabbix_user, password=zabbix_password)
zapi = ZabbixAPI(server=zabbix_server)
zapi.login(zabbix_user,zabbix_password)
answer=zapi.do_request('apiinfo.version')
print ("Version:",answer['result'])
timefrom_v=int(time.time())
item_id = zapi.item.get(hostids=10084, search={'key_':'zabbix[items_unsupported]'}, output=['itemid'])[0]['itemid']
utilization_now = zapi.history.get(itemids=item_id, time_from= timefrom_v-180, output='extend', history=3,limit=1)[0]['value']
utilization_day_ago = zapi.history.get(itemids=item_id, time_from= timefrom_v-86400, output='extend', history=3,limit=1)[0]['value']
#get hosts
hosts = zapi.host.get(output=['hostid', 'name'], filter={'status': 0})
unsupported_metrics = []
#get unsupported metrics
for host in hosts:
    items = zapi.item.get(output=['name', 'key_', 'error'], hostids=host['hostid'], filter={'state': 1,'status':0})
    if items:
        for item in items:
            unsupported_metrics.append({
                'host': host['name'],
                'name': item['name'],
                'key': item['key_'],
                'error': item['error']
            })

#Gen excel
workbook = xlsxwriter.Workbook('unsupported_metrics.xlsx')
worksheet = workbook.add_worksheet()
#titles
worksheet.write(0, 0, 'Host')
worksheet.write(0, 1, 'Name')
worksheet.write(0, 2, 'Key')
worksheet.write(0, 3, 'Error')
yellow = workbook.add_format({'bg_color': '#FFFF00'})  # yellow bckgcolor
#titles
worksheet.write(0, 4, "Unsupported metrics count now",yellow)
worksheet.write(0, 5, utilization_now,yellow)
worksheet.write(1, 4, "Unsupported metrics count day ago",yellow)
worksheet.write(1, 5, utilization_day_ago,yellow)
#writing
for row, metric in enumerate(unsupported_metrics, start=1):
    worksheet.write(row, 0, metric['host'])
    worksheet.write(row, 1, metric['name'])
    worksheet.write(row, 2, metric['key'])
    worksheet.write(row, 3, metric['error'])
#width
worksheet.set_column(0, 0, 15)
worksheet.set_column(1, 1, 50)
worksheet.set_column(2, 2, 30)
worksheet.set_column(3, 3, 30)
worksheet.set_column(4, 6, 40)
workbook.close()

from fortigate_api import FortiGateAPI
import xlsxwriter
import getpass
import datetime

# region Подключение
User = str(input('Enter username: '))
# PASSWORD = str(input('Enter password: '))
PASSWORD = getpass.getpass('Enter password: ')

FG_Main = FortiGateAPI(host='nsk-fgt', port=4443, username=User, password=PASSWORD, logging_error=True)

Get_Policy = FG_Main.cmdb.firewall.policy.get()  # Get-запрос политик
Get_VIP = FG_Main.cmdb.firewall.vip.get()  # Get-запрос VIP'ов

FileName = str(input('Enter file name: '))
Directory = "C:\\users\\{}\\Desktop\\" + FileName + ".xlsx"
Main_Book = xlsxwriter.Workbook(Directory.format(getpass.getuser()))
Main_Sheet = Main_Book.add_worksheet()
# endregion

# region Форматирование для Excel
merge_format = Main_Book.add_format(
    {
        "border": 1,
        "align": "center",
        "valign": "vcenter",
    }
)

merge_format2 = Main_Book.add_format(
    {
        "bottom": 1,
    }
)

Main_Sheet.set_column(0, 9, 35)
Main_Sheet.write(0, 0, 'Policy ID', merge_format)
Main_Sheet.write(0, 1, 'Policy name', merge_format)
Main_Sheet.write(0, 2, 'Policy source interface', merge_format)
Main_Sheet.write(0, 3, 'Policy destination interface', merge_format)
Main_Sheet.write(0, 4, 'Policy service', merge_format)
Main_Sheet.write(0, 5, 'Policy source address', merge_format)
Main_Sheet.write(0, 6, 'Policy destination address', merge_format)
Main_Sheet.write(0, 7, 'VIP external IP', merge_format)
Main_Sheet.write(0, 8, 'VIP mapped IP', merge_format)
# Main_Sheet.write(0, 9, 'VIP source filter', merge_format)
# Main_Sheet.write(0, 10, 'VIP service filter', merge_format)
Main_Sheet.write(0, 9, 'Policy expiry date', merge_format)
# endregion

# region Основной цикл
Current_Date = datetime.datetime.now()
Current_Date_INT = int(Current_Date.timestamp())

Chk_Row = 1
Counter = 0
while Counter != len(Get_Policy):
    if 'publ_' in Get_Policy[Counter]['name'] or 'PUBL_' in Get_Policy[Counter]['name']:  # Поиск публикаций

        # CHECK EXPIRY POLICY
        if Get_Policy[Counter]['policy-expiry'] == 'enable':
            Policy_Expiry = datetime.datetime.fromisoformat(Get_Policy[Counter]['policy-expiry-date']).timestamp()
            if Policy_Expiry < Current_Date_INT:
                Counter += 1
                continue
            else:
                pass

        Chk_Col = 0
        Main_Sheet.write(Chk_Row, Chk_Col, Get_Policy[Counter]['policyid'])
        Chk_Col += 1
        Main_Sheet.write(Chk_Row, Chk_Col, Get_Policy[Counter]['name'])

        # SOURCE INTERFACE
        temp_counter = 0
        Chk_Col += 1
        Main_Sheet.write(Chk_Row, Chk_Col, Get_Policy[Counter]['srcintf'][temp_counter]['name'])

        if len(Get_Policy[Counter]['srcintf']) > temp_counter:  # На случай если srcintf > 1
            SRC_INTF = 0
            while temp_counter != len(Get_Policy[Counter]['srcintf']) - 1:
                temp_counter += 1
                Chk_Row += 1
                Main_Sheet.write(Chk_Row, Chk_Col, Get_Policy[Counter]['srcintf'][temp_counter]['name'])
                SRC_INTF += 1
        Chk_Row -= SRC_INTF

        # DESTINATION INTERFACE
        temp_counter = 0
        Chk_Col += 1
        Main_Sheet.write(Chk_Row, Chk_Col, Get_Policy[Counter]['dstintf'][temp_counter]['name'])

        if len(Get_Policy[Counter]['dstintf']) > temp_counter:  # На случай если dstintf > 1
            DST_INTF = 0
            while temp_counter != len(Get_Policy[Counter]['dstintf']) - 1:
                temp_counter += 1
                Chk_Row += 1
                Main_Sheet.write(Chk_Row, Chk_Col, Get_Policy[Counter]['dstintf'][temp_counter]['name'])
                DST_INTF += 1
        Chk_Row -= DST_INTF

        # SERVICES
        temp_counter = 0
        Chk_Col += 1
        Main_Sheet.write(Chk_Row, Chk_Col, Get_Policy[Counter]['service'][temp_counter]['name'])

        if len(Get_Policy[Counter]['service']) > temp_counter:  # На случай если service > 1
            Services = 0
            while temp_counter != len(Get_Policy[Counter]['service']) - 1:
                temp_counter += 1
                Chk_Row += 1
                Main_Sheet.write(Chk_Row, Chk_Col, Get_Policy[Counter]['service'][temp_counter]['name'])
                Services += 1
        Chk_Row -= Services

        # SOURCE ADDRESS
        if Get_Policy[Counter]['srcaddr'] == []:  # Если в SRC ADDRESS используется IS
            Chk_Col += 1
            temp_counter = 0
            Main_Sheet.write(Chk_Row, Chk_Col, Get_Policy[Counter]['internet-service-src-name'][temp_counter]['name'])
            if len(Get_Policy[Counter]['internet-service-src-name']) > temp_counter:
                IS_ADDR = 0
                while temp_counter != len(Get_Policy[Counter]['internet-service-src-name']) - 1:
                    temp_counter += 1
                    Chk_Row += 1
                    Main_Sheet.write(Chk_Row, Chk_Col, Get_Policy[Counter]['internet-service-src-name'][temp_counter]['name'])
                    IS_ADDR += 1
            Chk_Row -= IS_ADDR

        else:
            temp_counter = 0
            Chk_Col += 1
            Main_Sheet.write(Chk_Row, Chk_Col, Get_Policy[Counter]['srcaddr'][temp_counter]['name'])

            if len(Get_Policy[Counter]['srcaddr']) > temp_counter:  # На случай если srcaddr > 1
                SRC_ADDR = 0
                while temp_counter != len(Get_Policy[Counter]['srcaddr']) - 1:
                    temp_counter += 1
                    Chk_Row += 1
                    Main_Sheet.write(Chk_Row, Chk_Col, Get_Policy[Counter]['srcaddr'][temp_counter]['name'])
                    SRC_ADDR += 1
            Chk_Row -= SRC_ADDR

        # DESTINATION ADDRESS
        temp_counter = 0
        Chk_Col += 1
        Main_Sheet.write(Chk_Row, Chk_Col, Get_Policy[Counter]['dstaddr'][temp_counter]['name'])

        # VIP EXTERNAL IP
        if 'vip_' in Get_Policy[Counter]['dstaddr'][temp_counter]['name'] or 'vsrv_' in Get_Policy[Counter]['dstaddr'][temp_counter]['name']:
            VIP_Counter = 0
            while VIP_Counter != len(Get_VIP):
                if Get_VIP[VIP_Counter]['name'] == Get_Policy[Counter]['dstaddr'][temp_counter]['name']:
                   Main_Sheet.write(Chk_Row, Chk_Col + 1, Get_VIP[VIP_Counter]['extip'])
                VIP_Counter += 1

        # VIP MAPPED IP
        if 'vip_' in Get_Policy[Counter]['dstaddr'][temp_counter]['name']:
            VIP_Counter = 0
            while VIP_Counter != len(Get_VIP):
                if Get_VIP[VIP_Counter]['name'] == Get_Policy[Counter]['dstaddr'][temp_counter]['name']:
                    Main_Sheet.write(Chk_Row, Chk_Col + 2, Get_VIP[VIP_Counter]['mappedip'][0]['range'])
                VIP_Counter += 1

        # VSRV REAL SERVER
        if 'vsrv_' in Get_Policy[Counter]['dstaddr'][temp_counter]['name']:
            VIP_Counter = 0
            VSRV_REAL = 0
            while VIP_Counter != len(Get_VIP):
                if Get_VIP[VIP_Counter]['name'] == Get_Policy[Counter]['dstaddr'][temp_counter]['name']:
                    VSRV_Counter = 0
                    while VSRV_Counter != len(Get_VIP[VIP_Counter]['realservers']):
                        Main_Sheet.write(Chk_Row, Chk_Col + 2, Get_VIP[VIP_Counter]['realservers'][VSRV_Counter]['ip'])
                        Chk_Row += 1
                        VSRV_REAL += 1
                        VSRV_Counter += 1
                VIP_Counter += 1
            Chk_Row -= VSRV_REAL

        # VIP SOURCE FILTER
        # if 'vip_' in Get_Policy[Counter]['dstaddr'][temp_counter]['name']:
        #     VIP_Counter = 0
        #     VIP_SRC_FLT = 0
        #     while VIP_Counter != len(Get_VIP):
        #         if Get_VIP[VIP_Counter]['name'] == Get_Policy[Counter]['dstaddr'][temp_counter]['name']:
        #             VIP_Service_count = 0
        #             while VIP_Service_count != len(Get_VIP[VIP_Counter]['src-filter']):
        #                 Main_Sheet.write(Chk_Row, Chk_Col + 3,
        #                                  Get_VIP[VIP_Counter]['src-filter'][VIP_Service_count]['range'])
        #                 Chk_Row += 1
        #                 VIP_SRC_FLT += 1
        #                 VIP_Service_count += 1
        #         VIP_Counter += 1
        #     Chk_Row -= VIP_SRC_FLT

        # VIP SERVICE FILTER
        # if 'vip_' in Get_Policy[Counter]['dstaddr'][temp_counter]['name']:
        #     VIP_Counter = 0
        #     VIP_SRV_FLT = 0
        #     while VIP_Counter != len(Get_VIP):
        #         if Get_VIP[VIP_Counter]['name'] == Get_Policy[Counter]['dstaddr'][temp_counter]['name']:
        #             VIP_Service_count = 0
        #             while VIP_Service_count != len(Get_VIP[VIP_Counter]['service']):
        #                 Main_Sheet.write(Chk_Row, Chk_Col + 4,
        #                                  Get_VIP[VIP_Counter]['service'][VIP_Service_count]['name'])
        #                 Chk_Row += 1
        #                 VIP_SRV_FLT += 1
        #                 VIP_Service_count += 1
        #         VIP_Counter += 1
        #     Chk_Row -= VIP_SRV_FLT

        if len(Get_Policy[Counter]['dstaddr']) > temp_counter:  # На случай если dstaddr > 1
            DST_ADDR = 0
            while temp_counter != len(Get_Policy[Counter]['dstaddr']) - 1:
                temp_counter += 1
                Chk_Row += 1
                Main_Sheet.write(Chk_Row, Chk_Col, Get_Policy[Counter]['dstaddr'][temp_counter]['name'])

                # VIP EXTERNAL IP
                if 'vip_' in Get_Policy[Counter]['dstaddr'][temp_counter]['name'] or 'vsrv_' in Get_Policy[Counter]['dstaddr'][temp_counter]['name']:
                    VIP_Counter = 0
                    while VIP_Counter != len(Get_VIP):
                        if Get_VIP[VIP_Counter]['name'] == Get_Policy[Counter]['dstaddr'][temp_counter]['name']:
                            Main_Sheet.write(Chk_Row, Chk_Col + 1, Get_VIP[VIP_Counter]['extip'])
                        VIP_Counter += 1

                # VIP MAPPED IP
                if 'vip_' in Get_Policy[Counter]['dstaddr'][temp_counter]['name']:
                    VIP_Counter = 0
                    while VIP_Counter != len(Get_VIP):
                        if Get_VIP[VIP_Counter]['name'] == Get_Policy[Counter]['dstaddr'][temp_counter]['name']:
                            Main_Sheet.write(Chk_Row, Chk_Col + 2, Get_VIP[VIP_Counter]['mappedip'][0]['range'])
                        VIP_Counter += 1

                # VSRV REAL SERVERS
                if 'vsrv_' in Get_Policy[Counter]['dstaddr'][temp_counter]['name']:
                    VIP_Counter = 0
                    VSRV_REAL = 0
                    while VIP_Counter != len(Get_VIP):
                        if Get_VIP[VIP_Counter]['name'] == Get_Policy[Counter]['dstaddr'][temp_counter]['name']:
                            VSRV_Counter = 0
                            while VSRV_Counter != len(Get_VIP[VIP_Counter]['realservers']):
                                Main_Sheet.write(Chk_Row, Chk_Col + 2,
                                                 Get_VIP[VIP_Counter]['realservers'][VSRV_Counter]['ip'])
                                Chk_Row += 1
                                VSRV_REAL += 1
                                VSRV_Counter += 1
                        VIP_Counter += 1
                    Chk_Row -= VSRV_REAL

                DST_ADDR += 1
        Chk_Row -= DST_ADDR

        # POLICY EXPIRY DATE
        if Get_Policy[Counter]['policy-expiry'] == 'enable':
            Chk_Col += 3
            Main_Sheet.write(Chk_Row, Chk_Col, Get_Policy[Counter]['policy-expiry-date'])

        # OTHER MANIPULATION
        Chk_Row += max(SRC_ADDR, DST_ADDR, Services) + 1
        Main_Sheet.set_row(Chk_Row - 1, 15, merge_format2)
    Counter += 1

Main_Sheet.autofilter('A1:J1')
Main_Book.close()
# endregion --------------------------------------------------------------------------------------------------------

print('--------------------------------------------------------------------------------------------------')
print('Done! File has been saved - ' + Directory)

int(input())
import requests
import json
import openpyxl
import datetime
import os


TOP_LIST = 'https://api.live.bilibili.com/xlive/app-room/v1/guardTab/topList'
GET_ROOM_INFO = 'https://api.live.bilibili.com/room/v1/Room/get_info'
XLSX_FILE_NAME = 'captain_export'
guard_level_list = {
    1: '总督',
    2: '提督',
    3: '舰长'
}


def get_top_list(roomId, ruid, page):
    return requests.get(TOP_LIST, {
        'roomid': roomId,
        'ruid': ruid,
        'page': page,
        'page_size': 30
    })


def get_list(roomId):
    r_t = requests.get(GET_ROOM_INFO,{"id": roomId})
    j_t = json.loads(r_t.text)
    # print(j_t['data']['uid'])
    ruid = j_t['data']['uid']
    r = get_top_list(roomId, ruid, 1)
    # print(len(json.loads(r.text)['data']['list']))
    toplist_result = json.loads(r.text)
    return_list = toplist_result['data']['top3']
    return_list += toplist_result['data']['list']
    for page in range(2, toplist_result['data']['info']['page']+1):
        r_top_list_loop = get_top_list(roomId, ruid, page)
        return_list += json.loads(r_top_list_loop.text)['data']['list']
    return return_list


def write_xlsx(user_list):
    xlsx_file_name = XLSX_FILE_NAME + datetime.datetime.now().strftime('%Y%m%d_%H%M%S') + '.xlsx'
    book = openpyxl.Workbook()
    sheet = book.active
    header = ['uid', 'username', 'guard_level']
    sheet.append(header)
    for user in user_list:
        user_guard_level = guard_level_list[user['guard_level']]
        sheet.append((user['uid'], user['username'], user_guard_level))
    book.save(xlsx_file_name)
    return xlsx_file_name

def main():
    print("舰长自动获取工具 版本：0.0.1   By cxhwd")
    roomId = input("请输入直播间id号，支持靓号或原房间号 (默认 5275):")
    if roomId == "":
        roomId = "5275"
    user_list = get_list(roomId)
    gen_file_name = write_xlsx(user_list)
    print('生成成功，文件名：' + gen_file_name)
    if os.name == 'nt':
        os.system('pause')


if __name__ == '__main__':
    main()

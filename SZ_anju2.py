import requests
from lxml import etree
import xlwt

# 分析页面
def lxmldata(data):
    datas = etree.HTML(data)
    list1 = []
    date = datas.xpath("//div[@class='list-content']//div[@class='zu-itemmod']")
    for i, dates in enumerate(date):
        dict = {}
        # 价格
        price = dates.xpath(".//div[@class='zu-side']//p//strong//b[@class='strongbox']/text()")
        if price:
            price = price[0]
        else:
            price = 'N/A'

        # 面积
        size = dates.xpath(".//p[@class='details-item tag']//b[@class='strongbox'][3]/text()")
        if size:
            dict['面积'] = size[0]
        else:
            dict['面积'] = 'N/A'

        # 房屋结构
        room_structure = dates.xpath(".//p[@class='details-item tag']//b[@class='strongbox'][1]//text()")
        room_structure = '/'.join(room_structure).strip()
        room_structure += '/' + '/'.join(
            dates.xpath(".//p[@class='details-item tag']//b[@class='strongbox'][2]//text()")).strip()
        if room_structure:
            dict['房间结构'] = room_structure
        else:
            dict['房间结构'] = 'N/A'

        # 详细标题
        title = dates.xpath(".//h3/a/b[@class='strongbox']//text()")
        title = ''.join(title).strip()
        if title:
            dict['详细标题'] = title
        else:
            dict['详细标题'] = 'N/A'

        # 具体位置
        local = dates.xpath(".//address[@class='details-item']/text()")
        if local:
            local = local[-1].strip()
        else:
            local = 'N/A'

        # Extract decoration status information
        decoration_status = dates.xpath(".//p[@class='details-item bot-tag']//span[contains(@class,'cls')]/text()")
        rental, orientation, elevator = 'N/A', 'N/A', 'N/A'

        if decoration_status:
            for status in decoration_status:
                if '整租' in status:
                    rental = status
                elif '朝向' in status:
                    orientation = status

        # Split the "装修情况" column into "朝向" and "电梯"
        decoration_info = dates.xpath(".//p[@class='details-item bot-tag']//span[contains(@class,'cls')]/text()")
        decoration_info = ', '.join(decoration_info)

        split_decoration = decoration_info.split(',')
        if len(split_decoration) > 1:
            orientation = split_decoration[1].strip()
        if len(split_decoration) > 2:
            elevator = split_decoration[2].strip()

        dict['出租'] = rental
        dict['朝向'] = orientation
        dict['电梯'] = elevator
        dict['价格'] = price
        dict['面积'] = size
        dict["详细标题"] = title
        dict['名称'] = local
        dict['房间结构'] = room_structure
        list1.append(dict)
    return list1

def save(data_list):
    if not data_list:
        print("No data to save.")
        return

    filename = "SZ_Rent_Data.xls"
    book = xlwt.Workbook(encoding='utf-8')
    sheet1 = book.add_sheet("sheet1")

    # 根据数据字典的键动态生成header
    header = data_list[0].keys()

    for i, column in enumerate(header):
        sheet1.write(0, i, column)

    j = 1
    for i in data_list:
        for col, val in enumerate(header):
            sheet1.write(j, col, i[val])
        j = j + 1

    book.save(filename)
    print("数据写入成功")

if __name__ == '__main__':
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/116.0.0.0 Safari/537.36",
        "Cookie": "Your Cookie String Here"
    }

    dict2 = []

    for i in range(50):
        url = "https://sz.zu.anjuke.com/fangyuan/p{}/".format(i + 1)
        response = requests.get(url=url, headers=headers)
        response.encoding = 'utf-8'
        response_text = response.text
        data_list = lxmldata(response_text)
        dict2.extend(data_list)
        print("第" + str(i + 1) + "页数据完成")

    save(dict2)

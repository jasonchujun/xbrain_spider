from xb_spider.spiders import *


class SharePointSpider(scrapy.Spider):
    name = "sharepoint"

    def __init__(self):
        with open("sharepoint_config.json", 'r') as f:
            self.config = json.load(f)

    def start_requests(self):
        url_host = "https://ericsson.sharepoint.com/sites/PDURadioChengduFirmware/Shared%20Documents/Forms/AllItems.aspx"
        yield scrapy.Request(url_host, headers=self.config['headers'], cookies=self.config['cookies'], meta={'cookiejar': 1}, callback=self.parse)

    def parse(self, response):
        url = "https://ericsson.sharepoint.com/sites/PDURadioChengduFirmware/_api/web/GetListUsingPath(DecodedUrl=@a1)/RenderListDataAsStream?@a1='/sites/PDURadioChengduFirmware/Shared Documents'&View=1d04f153-2aa3-4553-b6e2-14b66655e6d4&TryNewExperienceSingle=TRUE"
        yield scrapy.Request(url=url, method='POST', meta={'cookiejar': response.meta['cookiejar']}, callback=self.parse_dir)

    def parse_dir(self, response):
        url_dir = "https://ericsson.sharepoint.com/sites/PDURadioChengduFirmware/_api/web/GetListUsingPath(DecodedUrl=@a1)/RenderListDataAsStream?@a1='/sites/PDURadioChengduFirmware/Shared Documents'&RootFolder=/sites/PDURadioChengduFirmware/Shared Documents/2018 TeamBuilding&View=1d04f153-2aa3-4553-b6e2-14b66655e6d4&TryNewExperienceSingle=TRUE"
        url_doc = "https://ericsson.sharepoint.com/sites/PDURadioChengduFirmware/_layouts/15/download.aspx?UniqueId=0cf748e8%2D85b9%2D4b8a%2D9b9e%2Dda6ecf75d039"
        res = json.loads(response.body)
        dir_list = []
        for item in res['Row'][0:5]:
            tmp = {}
            tmp['UniqueId'] = item['UniqueId'].replace('{', '').replace('}', '')
            tmp['type'] = 'doc' if item['File_x0020_Size'] != '' else 'dir'
            tmp['name'] = item['FileLeafRef']
            tmp['urlencode'] = item['FileRef']
            tmp['editor_name'] = item['Editor'][0]['title']
            tmp['editor_email'] = item['Editor'][0]['email']
            tmp['modified'] = item['Modified.']
            # tmp['created'] = item['Created_x0020_Date']
            # dir_list.append(tmp)

            if tmp['type'] == 'dir':
                yield scrapy.Request(url_dir.replace(re.findall("RootFolder=(.*?)&View", url_dir)[0], item['FileRef']), method='POST', meta={'cookiejar': response.meta['cookiejar']}, callback=self.parse_dir)
            elif tmp['type'] == 'doc':
                if tmp['name'].split('.')[1] in ['doc', 'docx', 'xls', 'xlsx', 'csv', 'ppt', 'pptx', 'pdf', 'txt']:
                    yield scrapy.Request(url_doc.replace(re.findall("UniqueId=(.*?)$", url_doc)[0], tmp['UniqueId']), method='POST', meta={'path': tmp['urlencode'][1:], 'cookiejar': response.meta['cookiejar']}, callback=self.parse_doc)

    def parse_doc(self, response):
        path = response.meta['path']
        os.makedirs(os.path.dirname(path), exist_ok=True)
        with open(path, "wb") as f:
            f.write(response.body)

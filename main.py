from datetime import datetime, date
from bs4 import BeautifulSoup
import asyncio
import aiohttp
import time

#current version: v0.1
#does it work?: Yes!
#dependancies: bs4, aiohttp,  #openpyxl, pandas, 

#TODO:
### before releasing as 1.0 ###
## def read_config()
## def fetch_target_patents()
## def to_excel() : openpyxl
## def to_img() : pandas or matplotlib
## add type hinting 

### features to consider ###
## English compatibility: record.documentTitleEng +@
## executable: pyinstaller
## GUI: ??

class DrugPatentTracker():
    def __init__(self):
        # default config                
        self.api_user_key="" 
        self.target_patents = [] # 13digit application numbers.
        self.date_of_interest = (date(1970,1,1), datetime.now().date()) # to filter out records
        self.result_array = []
        
        #TODO
        # if os.exists(config_path): self.read_config(config_path)

    #TODO: 
    def read_config(self):
        #comment out this test code
        from config import api_user_key
        self.api_user_key = api_user_key
    
    #TODO: 
    def fetch_target_patents(self):
        # comment out this test code
        from config import target_patents
        print(target_patents)
        self.target_patents = target_patents

    async def track_patents(self):
        """ Track patents in self.target_patents

        Simply operates by repeating self.track_patent()
        
        Returns a 3-dimensional list
        """
        
        # returned values of each task will be appended into an empty list and then returned.
        # -> 3-dimensional list.
        futures = [asyncio.ensure_future(self.track_patent(patent)) for patent in self.target_patents]
        results = await asyncio.gather(*futures) 

        ## this code will work synchronously -> compare with async
        # results = []
        # for patent in self.target_patents:
        #     results.append(await self.track_patent(patent))
        # print(results)
        self.result_array = results
        return results

    async def track_patent(self, patent):
        """ Requests information of a patent and filter_out unneccesary information.
        
        Returns a 2-dimensional list.
        """
        records = await self.request_and_parse_kipris_API(application_number = patent, api_user_key = self.api_user_key)
        result_table = await self.filter_records(records)
        return result_table

    async def request_and_parse_kipris_API(self, application_number, api_user_key):
        """ Request kipris REST API (asynchronously) and parse data using Beautifulsoup.

        soup.findall("relateddocsonfileInfo") will be returned.
        """

        url = 'http://plus.kipris.or.kr/openapi/rest/RelatedDocsonfilePatService/relatedDocsonfileInfo'
        query = f'?applicationNumber={application_number}&accessKey={api_user_key}'
        
        time1 = time.time()
        print(f"request for patent:{application_number} started.")
        
        ## request by requests and loop.run_in_executor
        # import requests
        # loop_ = asyncio.get_event_loop()
        # response = await loop_.run_in_executor(None, requests.get, url+query)
        # text= response.text
        
        # request by aiohttp
        async with aiohttp.ClientSession() as session: 
            async with session.get(url+query) as response:
                text = await response.text()
        
        time2 = time.time()
        print(f"request for patent:{application_number} finished. time:{time2-time1}")
        
        # parse
        soup = BeautifulSoup(text, "xml")
        records = soup.find_all('relateddocsonfileInfo')
        if records == []: 
            print("No records detected. Please check result message ")
            print(f"result_message: {soup.find('resultMsg').text}")
        return records

    async def filter_records(self, records):
        """ Filters out unnecessary records and fields.

        Returns a 2-dimensional list.
        """
        filtered_records = []
        time1 = time.time()
        for i, record in enumerate(records):
            ymd = record.documentDate.text
            record_date = date(int(ymd[:4]), int(ymd[4:6]), int(ymd[6:8]))
            if record_date < self.date_of_interest[0] or record_date > self.date_of_interest[1]:
                continue
            else:
                filtered_records.append([
                    i+1, #n-th record
                    record.documentTitle.text, #서류명
                    record.documentDate.text, #접수/발송 일자
                    record.status.text, #처리상태
                    record.registrationNumber.text #접수/발송 번호
                ])
        time2 = time.time()
        print(f"filtering records from a patent finished. time:{time2-time1}")
        return filtered_records

if __name__ == "__main__":
    tracker = DrugPatentTracker()
    tracker.read_config()
    tracker.fetch_target_patents()
    print(f"number of patents to track: {len(tracker.target_patents)} ")
    time1 = time.time()
    loop = asyncio.get_event_loop()
    loop.run_until_complete(tracker.track_patents())
    loop.close
    time2 = time.time()
    print(f"total time taken: {time2-time1}")
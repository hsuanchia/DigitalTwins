import xlwings as xw
from openai import OpenAI
import os
from rich.progress import Progress, TextColumn, BarColumn, TimeElapsedColumn, TimeRemainingColumn

client = OpenAI(api_key='Please use your own api key')
def ChatGPT_api(input_text):
    response = client.chat.completions.create(
        model="gpt-3.5-turbo-0125",
            messages=[
                {
                    "role": "user",
                    "content": input_text
                }
            ],
            temperature=0.8,
            max_tokens=256,
            top_p=0.8,
            frequency_penalty=1,
            presence_penalty=0
    )
    ans = response.choices[0].message.content

    return ans

def prediction():
    excel_wb = xw.Book("Excel File record your question and ChatGPT's response")                                     
    excel_ws = excel_wb.sheets['prediction']                                             
    length = excel_ws.range('A1:A10000').end("down").row                           
    with Progress(TextColumn("[progress.description]{task.description}"),
                BarColumn(),
                TextColumn("[progress.percentage]{task.percentage:>3.0f}%"),
                TimeRemainingColumn(),
                TimeElapsedColumn()) as progress:                                  
        data_tqdm = progress.add_task(description="Asking ChatGPT", total=length)  
        for i in range(0, length-1):                                               
            input_text = excel_ws.cells(str(i+2), 'A').value                
            ans = ChatGPT_api(input_text)                                   
            excel_ws.cells(str(i+2), 'B').value = ans                       
            excel_wb.save()                                                                                     
            progress.advance(data_tqdm, advance=1)       

if __name__ == '__main__':
    excel_wb = xw.Book("Excel File record your question and ChatGPT's response")    ### Excel 的檔名
    excel_ws = excel_wb.sheets['data']                                              ### 工作表名稱
    length = excel_ws.range('A1:A10000').end("down").row                            ### 知道一個column有幾個有數據(資料數量)
    with Progress(TextColumn("[progress.description]{task.description}"),
                BarColumn(),
                TextColumn("[progress.percentage]{task.percentage:>3.0f}%"),
                TimeRemainingColumn(),
                TimeElapsedColumn()) as progress:                                   ### 進度條設定
        data_tqdm = progress.add_task(description="Asking ChatGPT", total=length)   ### 進度條任務設定
        for i in range(0, length-1):                                                ### 會把這個A欄位的資料依序詢問ChatGPT並且將回答儲存在右方
            ### 檢查回答是否為(N/A)或是空白 -> 如果不為空, 則跳過詢問
            if excel_ws.cells(str(i+2), 'K').value == 'N/A' or excel_ws.cells(str(i+2), 'K').value == None:  
                try:
                    input_text = excel_ws.cells(str(i+2), 'J').value                ### 問題在excel表的哪一格, cells(row, column)
                    ans = ChatGPT_api(input_text)                                   ### 透過api詢問ChatGPT
                    excel_ws.cells(str(i+2), 'K').value = ans                       ### 將ChatGPT的回答儲存到哪一格, cells(row, column)
                    excel_wb.save()                                                 ### 將回答儲存到excel
                except:
                    print(f"{i} q Error")                                           ### 如果詢問過程有問題, 則印出Error以及q的編號
            progress.advance(data_tqdm, advance=1)                                  ### 進度條進度+1

    prediction()
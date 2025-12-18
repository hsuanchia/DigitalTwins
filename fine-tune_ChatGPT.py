#coding=utf-8
import xlwings as xw
import jsonlines, time
from openai import OpenAI
from rich.progress import Progress, TextColumn, BarColumn, TimeElapsedColumn, TimeRemainingColumn, track

def build_jsonl_prompt_completion():
    excel_wb = xw.Book("Excel File record your question and ChatGPT's response")                         
    excel_ws = excel_wb.sheets['data']                                              
    length = excel_ws.range('A1:A10000').end("down").row   
    w = jsonlines.open("Json file record your question and ChatGPT's response",'w')
    for i in track(range(0, length-1)):
        tmp_prompt = excel_ws.cells(str(i+2), 'J').value
        tmp_completion = excel_ws.cells(str(i+2), 'K').value
        tmp_jsonl = {"prompt" : tmp_prompt, "completion" : tmp_completion}
        w.write(tmp_jsonl)
        print(tmp_jsonl)

def build_jsonl_prompt_completion_val():
    excel_wb = xw.Book("Excel File record your question and ChatGPT's response")                         
    excel_ws = excel_wb.sheets['data']                                              
    length = excel_ws.range('A1:A10000').end("down").row   
    train_len = length * 0.8
    train_w = jsonlines.open("Json file record your question and ChatGPT's response",'w')
    val_w = jsonlines.open("Json file record your question and ChatGPT's response",'w')
    for i in track(range(0, length-1)):
        tmp_prompt = excel_ws.cells(str(i+2), 'J').value
        tmp_completion = excel_ws.cells(str(i+2), 'K').value
        tmp_jsonl = {"prompt" : tmp_prompt, "completion" : tmp_completion}
        if i < train_len:
            train_w.write(tmp_jsonl)
        else:
            val_w.write(tmp_jsonl)
        print(tmp_jsonl)

def build_jsonl_prompt_completion_new_qa():
    excel_wb = xw.Book("Excel File record your question and ChatGPT's response")                     
    excel_ws = excel_wb.sheets['data']                                              
    length = excel_ws.range('A1:A10000').end("down").row   
    w = jsonlines.open("Json file record your question and ChatGPT's response",'w')
    for i in track(range(0, length-1)):
        tmp_prompt = excel_ws.cells(str(i+2), 'J').value
        tmp_completion = excel_ws.cells(str(i+2), 'K').value
        tmp_jsonl = {"prompt" : tmp_prompt, "completion" : tmp_completion}
        w.write(tmp_jsonl)
        print(tmp_jsonl)

def build_jsonl_chat_completion():
    excel_wb = xw.Book("Excel File record your question and ChatGPT's response")                         
    excel_ws = excel_wb.sheets['data']                                              
    length = excel_ws.range('A1:A10000').end("down").row   
    w = jsonlines.open("Json file record your question and ChatGPT's response",'w')
    for i in track(range(0, length-1)):
        tmp_prompt = excel_ws.cells(str(i+2), 'J').value
        tmp_completion = excel_ws.cells(str(i+2), 'K').value
        tmp_jsonl = {"messages" : [{"role" : "user", "content" : tmp_prompt}, {"role" : "assistant", "content" : tmp_completion}]}
        w.write(tmp_jsonl)
        print(tmp_jsonl)

def build_jsonl_chat_completion_val():
    excel_wb = xw.Book("Excel File record your question and ChatGPT's response")                         
    excel_ws = excel_wb.sheets['data']                                              
    length = excel_ws.range('A1:A10000').end("down").row   
    train_len = length * 0.8
    train_w = jsonlines.open("Json file record your question and ChatGPT's response",'w')
    val_w = jsonlines.open("Json file record your question and ChatGPT's response",'w')
    for i in track(range(0, length-1)):
        tmp_prompt = excel_ws.cells(str(i+2), 'J').value
        tmp_completion = excel_ws.cells(str(i+2), 'K').value
        tmp_jsonl = {"messages" : [{"role" : "user", "content" : tmp_prompt}, {"role" : "assistant", "content" : tmp_completion}]}
        if i < train_len:
            train_w.write(tmp_jsonl)
        else:
            val_w.write(tmp_jsonl)
        print(tmp_jsonl)

def build_jsonl_chat_completion_new_qa():
    excel_wb = xw.Book("Excel File record your question and ChatGPT's response")                         
    excel_ws = excel_wb.sheets['data']                                              
    length = excel_ws.range('A1:A10000').end("down").row   
    w = jsonlines.open("Json file record your question and ChatGPT's response",'w')
    for i in track(range(0, length-1)):
        tmp_prompt = excel_ws.cells(str(i+2), 'J').value
        tmp_completion = excel_ws.cells(str(i+2), 'K').value
        tmp_jsonl = {"messages" : [{"role" : "user", "content" : tmp_prompt}, {"role" : "assistant", "content" : tmp_completion}]}
        w.write(tmp_jsonl)
        print(tmp_jsonl)

def fine_tune_prediction_all_gpt(client):
    excel_wb = xw.Book("Excel File record your question and ChatGPT's response")                         
    excel_ws = excel_wb.sheets['prediction']
    with open("Json file record your question and ChatGPT's response","r+",encoding='UTF-8') as f:
        ind = 2
        for l in track(jsonlines.Reader(f)):
            # print(l)
            q = l['messages'][0]
            a = l['messages'][1]
            completion = client.chat.completions.create(
                model="ft:gpt-3.5-turbo-0125:personal::9BkerdNC",
                messages=[q]
            )
            pred = completion.choices[0].message.content
            print(q)
            print(a['content'])
            print(pred)
            excel_ws.cells(str(ind), 'A').value = q['content']    
            excel_ws.cells(str(ind), 'B').value = pred
            excel_ws.cells(str(ind), 'C').value = a['content']
            excel_wb.save()  
            ind += 1

def fine_tune_prediction_all_gpt_new_qa(client):
    excel_wb = xw.Book("Excel File record your question and ChatGPT's response")                         
    excel_ws = excel_wb.sheets['prediction']

    with open("Json file record your question and ChatGPT's response","r+",encoding='UTF-8') as f:
        ind = 2
        for l in track(jsonlines.Reader(f)):
            # print(l)
            q = l['messages'][0]
            a = l['messages'][1]
            completion = client.chat.completions.create(
                model="ft:gpt-3.5-turbo-0125:personal::9BkerdNC",
                messages=[q]
            )
            pred = completion.choices[0].message.content
            print(q)
            print("Answer:    ",a['content'])
            print("Prediction:",pred)
            excel_ws.cells(str(ind), 'A').value = q['content']    
            excel_ws.cells(str(ind), 'B').value = pred
            excel_ws.cells(str(ind), 'C').value = a['content']
            excel_wb.save()  
            ind += 1

def fine_tune_prediction_all_gpt_new_qa_binary(client):
    excel_wb = xw.Book("Excel File record your question and ChatGPT's response")                         
    excel_ws = excel_wb.sheets['prediction']

    data_length = excel_ws.range('A1:A10000').end("down").row
    x = 0
    for i in track(range(data_length-1)):
        if x >= 20:
            break
        ind = i+2
        x += 1
        question = excel_ws.cells(str(ind), 'A').value
        input_q = {"role" : "user", "content" : question}
        print(input_q)
        answer = excel_ws.cells(str(ind), 'C').value
        completion = client.chat.completions.create(
            model="ft:gpt-3.5-turbo-0125:personal::97N9644P",
            messages=[input_q]
        )
        pred = completion.choices[0].message.content
        print(input_q)
        print("Answer:    ", answer)
        print("Prediction:", pred)
        excel_ws.cells(str(ind), 'B').value = pred
        excel_wb.save()  

def process_data_binary():
    excel_wb = xw.Book("Excel File record your question and ChatGPT's response")                         
    excel_ws = excel_wb.sheets['prediction']
    replace_prompt = '11.回答時僅能回答:支持、不支持，我們會將原始訓練資料類別為: 非常支持跟有點支持分類為支持；將原始訓練資料類別為: 非常不支持、有點不支持跟普通分類為不支持'
    # positive_example = '接下來我會給你一些受試者的人物設定，你必須扮演一個受試者回答問題：1.性別為:女性2.職業屬於:民意代表、行政和企業主管、經理人員及自營商3.教育程度為:高中、職4.臺灣政黨傾向為:中立（都不偏）5.假如我現在居住地的鄉、鎮、市、或區符合核廢料存放的科學條件，我會回答「非常不同意6.我「非常同意7.我「非常不同意8.我「不太同意9.我對於環保優先還是經濟優先的看法為:經濟跟環保一樣重要[不提示]。'
    # negative_example = '接下來我會給你一些受試者的人物設定，你必須扮演一個受試者回答問題：1.性別為:男性2.職業屬於:公務員、教師3.教育程度為:專科4.臺灣政黨傾向為:中立（都不偏）5.假如我現在居住地的鄉、鎮、市、或區符合核廢料存放的科學條件，我會回答「還算同意6.我「非常不同意7.我「還算同意8.我「非常不同意9.我對於環保優先還是經濟優先的看法為:經濟比環保重要。'
    # final_replace_prompt = replace_prompt + '，以下是支持的例子: '+ positive_example + '以下是反對的例子: ' + negative_example

    data_length = excel_ws.range('A1:A10000').end("down").row
    for i in track(range(data_length-1)):
        ind = i+2
        question = excel_ws.cells(str(ind), 'A').value
        tmp = question.split('\n')
        tmp[-3] = replace_prompt
        new_prompt = "\n".join(tmp)
        excel_ws.cells(str(ind), 'A').value = new_prompt

def process_data_binary_2():
    excel_wb = xw.Book("Excel File record your question and ChatGPT's response")                         
    excel_ws = excel_wb.sheets['prediction']
    replace_prompt = ['11.回答時僅能回答:支持、不支持', 
                      '12.我們會將原始訓練資料類別為: 非常支持跟有點支持分類為支持', 
                      '13.我們會將原始訓練資料類別為: 非常不支持、有點不支持跟普通分類為不支持',
                      '14.請將回答放在回應的最前方，並且在前方加上"回答:"',
                      '15.請將回答依據接在回答後面'
                      ]
    # positive_example = '接下來我會給你一些受試者的人物設定，你必須扮演一個受試者回答問題：1.性別為:女性2.職業屬於:民意代表、行政和企業主管、經理人員及自營商3.教育程度為:高中、職4.臺灣政黨傾向為:中立（都不偏）5.假如我現在居住地的鄉、鎮、市、或區符合核廢料存放的科學條件，我會回答「非常不同意6.我「非常同意7.我「非常不同意8.我「不太同意9.我對於環保優先還是經濟優先的看法為:經濟跟環保一樣重要[不提示]。'
    # negative_example = '接下來我會給你一些受試者的人物設定，你必須扮演一個受試者回答問題：1.性別為:男性2.職業屬於:公務員、教師3.教育程度為:專科4.臺灣政黨傾向為:中立（都不偏）5.假如我現在居住地的鄉、鎮、市、或區符合核廢料存放的科學條件，我會回答「還算同意6.我「非常不同意7.我「還算同意8.我「非常不同意9.我對於環保優先還是經濟優先的看法為:經濟比環保重要。'
    # final_replace_prompt = replace_prompt + '，以下是支持的例子: '+ positive_example + '以下是反對的例子: ' + negative_example

    data_length = excel_ws.range('A1:A10000').end("down").row
    for i in track(range(data_length-1)):
        ind = i+2
        question = excel_ws.cells(str(ind), 'A').value
        tmp = question.split('\n')
        new_tmp = tmp[:-3]
        new_tmp.extend(replace_prompt)
        new_prompt = "\n".join(new_tmp)
        # print(new_prompt)
        excel_ws.cells(str(ind), 'A').value = new_prompt

def fine_tune_prediction_all_babbage(client):
    excel_wb = xw.Book("Excel File record your question and ChatGPT's response")                         
    excel_ws = excel_wb.sheets['prediction']
    with open("Json file record your question and ChatGPT's response","r+",encoding='UTF-8') as f:
        ind = 2
        for l in track(jsonlines.Reader(f)):
            # print(l)
            q = l['prompt']
            a = l['completion']
            completion = client.completions.create(
                model="ft:babbage-002:personal::9BjS37XJ",
                prompt=[q]
            )
            # completion = client.Completion.create(
            #     model="ft:babbage-002:personal::9BkaHKIF",
            #     prompt=[q]
            # )
            pred = completion.choices[0].text
            print(q)
            print("Answer:    ",a)
            print("Prediction:",pred)
            excel_ws.cells(str(ind), 'A').value = q    
            excel_ws.cells(str(ind), 'B').value = pred
            excel_ws.cells(str(ind), 'C').value = a
            excel_wb.save()  
            ind += 1

def fine_tune_prediction_all_babbage_new_qa(client):
    excel_wb = xw.Book("Excel File record your question and ChatGPT's response")                         
    excel_ws = excel_wb.sheets['prediction']
    with open("Json file record your question and ChatGPT's response","r+",encoding='UTF-8') as f:
        ind = 2
        for l in track(jsonlines.Reader(f)):
            # print(l)
            q = l['prompt']
            a = l['completion']
            completion = client.completions.create(
                model="ft:babbage-002:personal::9BjS37XJ",
                prompt=[q]
            )
            pred = completion.choices[0].text
            print(q)
            print("Answer:    ",a)
            print("Prediction:",pred)
            excel_ws.cells(str(ind), 'A').value = q    
            excel_ws.cells(str(ind), 'B').value = pred
            excel_ws.cells(str(ind), 'C').value = a
            excel_wb.save()  
            ind += 1

if __name__ == '__main__' :
    # build_jsonl_prompt_completion()
    # build_jsonl_chat_completion()
    build_jsonl_prompt_completion_val()
    # build_jsonl_chat_completion_val()
    # build_jsonl_chat_completion_new_qa()
    # build_jsonl_prompt_completion_new_qa()
    # start_fine_tune(client) 
    # fine_tune_prediction(client) 
    # fine_tune_prediction_all_gpt(client)
    # fine_tune_prediction_all_babbage(client)
    # fine_tune_prediction_babbage(client)
    # fine_tune_prediction_all_gpt_new_qa(client)
    # fine_tune_prediction_all_babbage_new_qa(client)
    # fine_tune_prediction_all_gpt_new_qa_binary(client)
    # process_data_binary()
    # process_data_binary_2()

# ft:gpt-3.5-turbo-0125:personal::97N9644P -> cost 0.02 to  2.10 -> cost 2.08
# ft:gpt-3.5-turbo-0125:personal::9BkerdNC -> cost 2.10 to 3.74 -> cost 1.64




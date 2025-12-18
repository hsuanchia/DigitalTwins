import xlwings as xw
from rich.progress import track

def gpt_turbo_evaluation():
    excel_wb = xw.Book("Excel File record your question and ChatGPT's response")                         
    excel_ws = excel_wb.sheets['prediction']
    length = excel_ws.range('A1:A10000').end("down").row   
    correct, total = 0, 0
    oppose_right, oppose_wrong, support_right, support_wrong = 0, 0, 0, 0
    for i in track(range(length-1)):
        total += 1
        ind = i + 2
        pred = excel_ws.cells(str(ind), 'B').value
        answer = excel_ws.cells(str(ind), 'C').value
        if '不支持' in pred:
            if '不支持' in answer:
                correct += 1
                oppose_right += 1
            else:
                oppose_wrong += 1
        else:
            if '支持' in answer and '不支持' not in answer:
                correct += 1
                support_right += 1
            else:
                support_wrong += 1

    print([[oppose_right, oppose_wrong],[support_wrong, support_right]])
    print(f"{correct} / {total} = {correct / total}")

def babbage_evaluation():
    excel_wb = xw.Book("Excel File record your question and ChatGPT's response")                         
    excel_ws = excel_wb.sheets['prediction']
    length = excel_ws.range('A1:A10000').end("down").row   
    flex_correct, direct_correct, total = 0, 0, 0
    for i in track(range(length-1)):
        total += 1
        ind = i + 2
        pred = excel_ws.cells(str(ind), 'D').value
        answer = excel_ws.cells(str(ind), 'C').value
        # Flexible options
        if '不支持' in pred:
            if '不支持' in answer:
                flex_correct += 1
        else:
            if '支持' in answer:
                flex_correct += 1
        # Directly compare
        if '不支持' in pred:
            if str(pred) == str(answer):
                direct_correct += 1
        else:
            if ('有點支持' in pred and '有點支持' in answer) or ('非常支持' in pred and '非常支持' in answer):
                direct_correct += 1
            

    print(f"Directly compare: {direct_correct} / {total} = {direct_correct / total}")
    print(f"Flexible options: {flex_correct} / {total} = {flex_correct / total}")

def before_fine_tune_evaluation():
    excel_wb = xw.Book("Excel File record your question and ChatGPT's response")                         
    excel_ws = excel_wb.sheets['prediction']
    length = excel_ws.range('A1:A10000').end("down").row   
    flex_correct, direct_correct, total = 0, 0, 0
    for i in track(range(length-1)):
        total += 1
        ind = i + 2
        pred = excel_ws.cells(str(ind), 'D').value
        answer = excel_ws.cells(str(ind), 'C').value
        # Flexible options
        if '不支持' in pred:
            if '不支持' in answer:
                flex_correct += 1
        else:
            if '支持' in answer:
                flex_correct += 1
        # Directly compare
        if str(answer) in pred:
            direct_correct += 1
        # if '不支持' in pred:
        #     if str(answer) in pred:
        #         direct_correct += 1
        # else:
        #     if ('有點支持' in pred and '有點支持' in answer) or ('非常支持' in pred and '非常支持' in answer):
        #         direct_correct += 1
            

    print(f"Directly compare: {direct_correct} / {total} = {direct_correct / total}")
    print(f"Flexible options: {flex_correct} / {total} = {flex_correct / total}")

if __name__ == '__main__':
    gpt_turbo_evaluation()
    # babbage_evaluation()
    # before_fine_tune_evaluation()
    
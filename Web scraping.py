import pandas as pd
import xlwt
import time

from selenium.webdriver import Chrome
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException

driver = Chrome()
df = pd.read_excel('/Users/pepitaxie/Downloads/excel/shishi.xlsx')
wb = xlwt.Workbook()

# Add an excel sheet
ws = wb.add_sheet('test_excel')
i = 0
for idx, row in df.iterrows():
    target1 = row['target1']
    target2 = row['target2']
    year_excel = row['year']

    # print(target1, target2)
    print(f'The IPC Combination is：{target1}, {target2}')
    target1 = target1.replace('/', '%2F')
    target2 = target2.replace('/', '%2F')
    url = 'https://worldwide.espacenet.com/patent/search?q=' + target1 + '%20' + target2
    driver.get(url)
    driver.maximize_window() # For maximizing window
    driver.implicitly_wait(20)  # Gives an implicit wait for 20 seconds
    time.sleep(1)
    url = driver.current_url
    print(url)
    try:
        if 'family' in url:
            result = 1
        else:
            driver.find_element(By.XPATH, '//*[@id="mui-component-select-Sort by"]').click()
            driver.find_element(By.XPATH, '//*[@id="menu-Sort by"]/div[3]/ul/li[3]').click()
            txt = driver.find_element(By.XPATH, '//*[@id="result-list"]/div[4]/article[1]/section/div[2]').text
            print(txt)
            for j in range(len(txt)):
                number = ['0', '1', '2', '3', '4', '5', '6', '7', '8', '9'] 
                if txt[j] in number:
                    break
            year_search = int(txt[j:j+4])
            print(f'The earliest result is：{year_search}')
            if year_search == year_excel:
                result = 1
            elif year_search < year_excel:
                result = 0
            elif year_search > year_excel:
                result = "error"
        print(f'The result is: {result}')
        ws.write(i, 0, result)
        i = i + 1
    except NoSuchElementException as e:
        print(f'An error occurred: {e}')
        ws.write(i, 0, "error_404")
        i = i + 1
        pass
    finally:
        wb.save('./test.xls')
    print(f"Now is the {i}th observation")
    continue

# Save the results in Excel
wb.save('./test.xls')
print('End')

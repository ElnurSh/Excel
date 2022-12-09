from string import digits  # 0123456789
from itertools import product
import win32com.client as client

possible_symbols = str(digits)
opened_doc = client.Dispatch("Excel.Application")

stop_loop = 0


def start_loop():
    try:
        interval = input("Select interval between two numbers, e.g., 3-7: ")
        interval = [int(item) for item in interval.split("-")]
    except:
        print("Apparently, your input data is not correct. Try again.")
    for item in range(interval[0], interval[1]+1):
        for password in product(possible_symbols, repeat=item):
            password = "".join(password)
            global stop_loop
            if stop_loop == 1:
                break
            try:
                opened_doc.Workbooks.Open(r"C:\Users\User\Desktop\test.xlsx", Password=password)
                print(f"Excel workbook`s password is: {password}")
                stop_loop += 1
            except:
                print(f"Incorrect password: {password}")
                pass


if __name__ == '__main__':
    start_loop()

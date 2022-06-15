import win32com.client
from pathlib import Path
        

def generate_password(length=1):
    s = "0123456789"
    # s += "abcdefghijklmnopqrstuvwxyz"
    # s += "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    # s += "~!@#$%^&*()_+|\<>,.?/"
    if (length == 1):
        for x in s:
            yield x
    else:
        for x in s:
            for y in generate_password(length-1):
                yield x + y
                
                
def attempt(excel, filename, password):
    try:
        excel.Workbooks.Open(filename, UpdateLinks=False, ReadOnly=True, Format=None, Password=password)
        excel.Quit()
        print("========={}=========".format(password))
        return True
    except Exception as e:
        print("Wrong Password:", password)
        return False


if __name__ == '__main__':
    # filename = Path("D:\a.xlsx")
    filename = Path.cwd().joinpath("test.xlsx")
    min_password_length = 1
    max_password_length = 20
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    ok = False
    for i in range(min_password_length, max_password_length + 1):
        if not ok:
            for password in generate_password(length=i):
                ok = attempt(excel, filename=filename, password=password)
                if ok:
                    break

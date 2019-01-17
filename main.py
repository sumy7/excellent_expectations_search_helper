import win32con
import win32api
import win32clipboard as w
import win32com.client as comclt
import time


# https://docs.microsoft.com/en-us/previous-versions/windows/internet-explorer/ie-developer/windows-scripting/8c6yea83(v=vs.84)
def send_parse_key():
    wsh = comclt.Dispatch("WScript.Shell")
    wsh.SendKeys("{TAB}")  # Press TAB
    wsh.SendKeys("^v")  # Press CTRL+V
    wsh.SendKeys("{ENTER}")  # Press ENTER


def clipboard_get_text():
    w.OpenClipboard()
    d = w.GetClipboardData(win32con.CF_UNICODETEXT)
    w.CloseClipboard()
    return d


def clipboard_set_text(text):
    w.OpenClipboard()
    w.EmptyClipboard()
    w.SetClipboardData(win32con.CF_UNICODETEXT, text)
    w.CloseClipboard()


if __name__ == '__main__':
    ans = input("Hello, Which search do you want to use? y-Baidu, n-Google")
    wordbook = 'baidu.cvs' if ans == 'y' else 'google.cvs'
    print("OK, choose %s... Waiting 5 sec to work" % wordbook)
    time.sleep(5)
    keywords_file = open(wordbook, 'r', encoding="UTF-8")
    keywords = keywords_file.readlines()
    keywords_file.close()
    print("Total %d line keywords" % len(keywords))
    for index, line in enumerate(keywords):
        word = line.strip().split(',')[0]
        print("%d/%d %s" % (index + 1, len(keywords), word))
        if len(word) == 0:
            continue
        clipboard_set_text(word)
        send_parse_key()
        time.sleep(0.5)
    print("ALL DONE! Have Fun!")

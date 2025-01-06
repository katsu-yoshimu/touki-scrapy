from selenium import webdriver
from plyer import notification

def openWindow():
    options = webdriver.ChromeOptions()
    options.add_argument('--window-position=0,0')
    options.add_argument('--window-size=200,300')
    options.add_experimental_option("excludeSwitches", ['enable-automation'])
    driver = webdriver.Chrome(options=options)

    html='<html><body>'
    html+='<h1>不動産請求情報収集処理</h1><div style=\"text-align: left;\">'
    html+='<b id=\"message\"style=\"background: yellow;\"></b>'
    html+='番目のスクリプト稼働中です。<br/>しばらくお待ちください。</div>'
    html+='</body></html>'
    
    script = f"document.write('{html}');"
    driver.execute_script(script)

    return driver

class ProcessStatus():
    driver = ""
    IS_NOTIFY_WINDOWS = 1
    IS_NOTIFY_BROWSER = 1

    def __init__(self, setting=None):
        if setting != None:
            self.IS_NOTIFY_WINDOWS = setting["ProcessStatus.IS_NOTIFY_WINDOWS"]
            self.IS_NOTIFY_BROWSER = setting["ProcessStatus.IS_NOTIFY_BROWSER"]
        
        if self.IS_NOTIFY_BROWSER == 1:
            self.driver = openWindow()
    
    def showStatus(self, message):
        try:
            if self.IS_NOTIFY_WINDOWS == 1:
                notification.notify(
                        title = "不動産請求情報収集",
                        message = f"{message}番目のスクリプト稼働中です。\nしばらくお待ちください。",
                        app_name = "不動産請求情報収集"
                    )
            
            if self.IS_NOTIFY_BROWSER == 1:
                script = f"document.getElementById(\"message\").innerText='{message}';"
                self.driver.execute_script(script)

        except Exception as e:
            if self.IS_NOTIFY_BROWSER == 1:
                self.driver = openWindow()
                self.driver.execute_script(script)
        
    def close(self):
        try:
            if self.IS_NOTIFY_BROWSER == 1:
                self.driver.quit()
        except Exception as e:
            pass
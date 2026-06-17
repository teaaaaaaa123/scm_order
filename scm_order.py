from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
import time
from datetime import datetime, timedelta
import json
import os
import traceback
import pandas as pd

# 手动指定ChromeDriver路径（使用本地ChromeDriver）
CHROMEDRIVER_PATH = r'e:\111\scm_order\chromedriver-147\chromedriver.exe'

try:
    from webdriver_manager.chrome import ChromeDriverManager
    from selenium.webdriver.chrome.service import Service
    USE_WEBDRIVER_MANAGER = True
except ImportError:
    USE_WEBDRIVER_MANAGER = False

def sleep(ms):
    time.sleep(ms / 1000)

class TestFullProcess:
    def __init__(self):
        print("=" * 60)
        print("初始化浏览器...")
        print("=" * 60)
        
        chrome_options = Options()
        chrome_options.add_argument('--start-maximized')
        chrome_options.add_argument('--disable-blink-features=AutomationControlled')
        chrome_options.add_argument('--no-sandbox')
        chrome_options.add_argument('--disable-dev-shm-usage')
        
        print(f"\n使用系统默认Chrome浏览器")
        print(f"使用webdriver-manager: {USE_WEBDRIVER_MANAGER}")
        
        try:
            # 优先使用手动指定的ChromeDriver
            if CHROMEDRIVER_PATH and os.path.exists(CHROMEDRIVER_PATH):
                print(f"\n使用本地ChromeDriver: {CHROMEDRIVER_PATH}")
                service = Service(CHROMEDRIVER_PATH)
                self.driver = webdriver.Chrome(service=service, options=chrome_options)
            elif USE_WEBDRIVER_MANAGER:
                print("\n正在自动下载匹配Chrome版本的ChromeDriver...")
                service = Service(ChromeDriverManager().install())
                self.driver = webdriver.Chrome(service=service, options=chrome_options)
            else:
                print("\n使用系统ChromeDriver...")
                self.driver = webdriver.Chrome(options=chrome_options)
            
            print("✓ Chrome浏览器启动成功!")
            
        except Exception as e:
            print(f"\n✗ 浏览器启动失败!")
            print(f"错误信息: {e}")
            print("\n" + "=" * 60)
            print("可能的解决方案:")
            print("1. 安装webdriver-manager: pip install webdriver-manager")
            print("2. 检查Chrome浏览器版本是否为145")
            print("3. 检查网络连接")
            print("4. 检查防火墙/杀毒软件是否阻止")
            print("=" * 60)
            print("\n详细错误:")
            traceback.print_exc()
            input("\n按回车键退出...")
            raise
        
        self.wait = WebDriverWait(self.driver, 30)
        self.actions = ActionChains(self.driver)
        self.production_no = ''  # 用于存储生产单号
    
    def find_banxing_by_rules(self, source, rule1, rule2, rule3, rule4):
        """根据规则组合从MTM所有品类规则.xlsx中查找版型编码"""
        excel_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'MTM所有品类规则.xlsx')
        
        if not os.path.exists(excel_path):
            print(f'  ⚠ 未找到版型规则文件: {excel_path}')
            return None
        
        try:
            df = pd.read_excel(excel_path, sheet_name=0)
            
            # 按条件筛选
            result = df[
                (df['来源工作表'] == source) &
                (df['规则1'] == rule1) &
                (df['规则2'] == rule2) &
                (df['规则3'] == rule3) &
                (df['规则4'] == rule4)
            ]
            
            if len(result) == 0:
                print(f'  ⚠ 未找到匹配的版型: {source} {rule1} {rule2} {rule3} {rule4}')
                return None

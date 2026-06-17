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
            
            banxing_code = result['版型编码'].iloc[0]
            print(f'  ✓ 根据规则查询到版型: {banxing_code}')
            return banxing_code
        
        except Exception as e:
            print(f'  ⚠ 查询版型失败: {e}')
            return None
    
    def load_banxing_from_config(self):
        """从config.json读取版型配置"""
        # 使用脚本所在目录下的config.json
        import os
        script_dir = os.path.dirname(os.path.abspath(__file__))
        config_file = os.path.join(script_dir, 'config.json')
        default_banxing_list = [
            {'banxing': '1KN001', 'chima_size': '50', 'luocha': 'R', 'custom_options': {'手巾袋': '无'},
             'fabric_width': 74, 'fabric_no': 'AI下单修改', 'fabric_style': '平板'}
        ]
        
        try:
            with open(config_file, 'r', encoding='utf-8') as f:
                config = json.load(f)
            
            # 保存配置到实例变量供其他方法使用
            self.config = config
            
            banxing_list = []
            
            # 先处理 banxingRules 配置（根据规则查询版型）
            banxing_rules = config.get('banxingRules', [])
            for rule_config in banxing_rules:
                source = rule_config.get('source')
                rule1 = rule_config.get('rule1')
                rule2 = rule_config.get('rule2')
                rule3 = rule_config.get('rule3')
                rule4 = rule_config.get('rule4')
                
                if not all([source, rule1, rule2, rule3, rule4]):
                    print(f'  ⚠ 规则配置不完整，跳过')
                    continue
                
                # 根据规则查询版型编码
                banxing_code = self.find_banxing_by_rules(source, rule1, rule2, rule3, rule4)
                
                if banxing_code:
                    style_type = self.get_style_type(banxing_code)
                    banxing_list.append({
                        'banxing': banxing_code,
                        'chima_size': rule_config.get('chimaSize', config.get('chimaSize', '50')),
                        'luocha': rule_config.get('luocha', '-' if style_type == '西裤' or style_type == '马甲' else 'R'),
                        'custom_options': rule_config.get('customOptions', {}),
                        'fabric_width': int(config.get('fabricMenfu', 74)),
                        'fabric_no': config.get('fabric', 'AI下单修改'),
                        'fabric_style': config.get('fabricStyle', '平板'),
                        'liangti_data': rule_config.get('liangtiData', {})
                    })
            
            # 再处理 banxingConfigs 配置（支持每个版型独立配置）
            banxing_configs = config.get('banxingConfigs', [])
            
            if banxing_configs:
                # 使用新的配置格式
                for banxing_config in banxing_configs:
                    banxing_code = banxing_config.get('banxing')
                    if not banxing_code:
                        continue
                    
                    style_type = self.get_style_type(banxing_code)
                    
                    # 获取版型独立配置，若无则使用全局默认值
                    banxing_list.append({
                        'banxing': banxing_code,
                        'chima_size': banxing_config.get('chimaSize', config.get('chimaSize', '50')),
                        'luocha': banxing_config.get('luocha', '-' if style_type == '西裤' or style_type == '马甲' else 'R'),
                        'custom_options': banxing_config.get('customOptions', {}),
                        'fabric_width': int(banxing_config.get('fabricMenfu', config.get('fabricMenfu', 74))),
                        'fabric_no': banxing_config.get('fabric', config.get('fabric', 'AI下单修改')),
                        'fabric_style': banxing_config.get('fabricStyle', config.get('fabricStyle', '平板')),
                        'liangti_data': banxing_config.get('liangtiData', {})
                    })
                
                print(f'  ✓ 从配置文件读取 {len(banxing_list)} 个版型配置')
            else:
                # 兼容旧的配置格式
                banxing_codes = config.get('banxingList', [])
                default_chima_size = config.get('chimaSize', '50')
                
                for banxing_code in banxing_codes:
                    style_type = self.get_style_type(banxing_code)
                    if style_type == '西裤':
                        luocha = '-'
                        custom_options = config.get('customOptions', {}).get('西裤', {'脚口': '反撬', '脚口反撬': '2'})
                    elif style_type == '马甲':
                        luocha = '-'
                        custom_options = config.get('customOptions', {}).get('马甲', {'马甲后背': '里布', '马甲后叉': '中开叉', '手巾袋': '无', '下开袋': '无'})
                    elif style_type == '大衣':
                        luocha = 'R'
                        custom_options = config.get('customOptions', {}).get('大衣', {'工艺': '半麻衬', '大衣面袋': '无', '手巾袋': '无', '大衣袖扣': '一扣',
                                         '袖口锁眼': '一扣', '后背款式': 'A款', '门襟锁眼': '一明扣三暗扣',
                                         '半里': '锁边翘边', '猎装': '腰带款', '驳头锁眼': '艾伦眼',
                                         '艾伦眼颜色': '顺色', '米兰眼颜色': '顺色'})
                    else:
                        luocha = 'R'
                        custom_options = config.get('customOptions', {}).get('上衣', {})
                    
                    banxing_list.append({
                        'banxing': banxing_code,
                        'chima_size': default_chima_size,
                        'luocha': luocha,
                        'custom_options': custom_options,
                        'fabric_width': int(config.get('fabricMenfu', 74)),
                        'fabric_no': config.get('fabric', 'AI下单修改'),
                        'fabric_style': config.get('fabricStyle', '平板')
                    })
                
                print(f'  ✓ 从配置文件读取 {len(banxing_list)} 个版型: {banxing_codes}')
            
            return banxing_list
        
        except Exception as e:
            print(f'  ⚠ 读取配置文件失败，使用默认版型: {e}')
            return default_banxing_list
    
    def get_style_type(self, banxing):
        type_map = {
            '1': '上衣',
            '4': '大衣',
            '5': '马甲',
            '6': '西裤'
        }
        return type_map.get(banxing[0], '上衣')
    
    def login(self):
        print('=' * 50)
        print('【1】登录系统...')
        print('=' * 50)
        
        self.driver.get('https://scm.ceyadi.cn/')
        sleep(2000)
        
        all_inputs = self.driver.find_elements(By.TAG_NAME, 'input')
        
        if len(all_inputs) > 0:
            all_inputs[0].clear()
            all_inputs[0].send_keys('AI下单')
        
        if len(all_inputs) > 1:
            all_inputs[1].clear()
            all_inputs[1].send_keys('123456')
        
        login_buttons = self.driver.find_elements(By.CSS_SELECTOR, 'button, .el-button')
        for btn in login_buttons:
            btn_text = btn.text.strip()
            if '登录' in btn_text:
                btn.click()
                break
        
        for i in range(15):
            sleep(1000)
            if '/login' not in self.driver.current_url:
                break
        
        sleep(3000)
        print('✓ 登录成功\n')
        return True
    
    def enter_custom_order(self):
        print('【2】进入定制下单...')
        
        menu_items = self.driver.find_elements(By.CSS_SELECTOR, '.el-menu-item')
        for item in menu_items:
            text = item.text.strip()
            if '生产下单管理' in text:
                self.driver.execute_script('arguments[0].click();', item)
                print(f'  ✓ 已点击：{text}')
                break
        
        sleep(3000)
        
        buttons = self.driver.find_elements(By.CSS_SELECTOR, 'button, .el-button')
        for btn in buttons:
            if btn.text.strip() == '定制下单':
                btn.click()
                print('  ✓ 已点击定制下单')
                break
        
        sleep(3000)
        print('✓ 已进入定制下单\n')
        return True
    
    def fill_dropdown(self, field_name, value, max_retries=3):
        for retry in range(max_retries):
            try:
                # 先通过 placeholder 找客商输入框（客商特殊处理）
                if field_name == '客商':
                    input_list = self.driver.find_elements(By.CSS_SELECTOR, '.el-input__inner')
                    for input_element in input_list:
                        try:
                            placeholder = input_element.get_attribute('placeholder')
                            if placeholder and '请选择客商' in placeholder:
                                print(f'    找到客商输入框，placeholder: {placeholder}')
                                input_element.click()
                                sleep(500)
                                
                                input_element.send_keys(value)
                                sleep(2000)
                                
                                # 查找并选择选项
                                options = self.driver.find_elements(By.CSS_SELECTOR, '.el-select-dropdown__item')
                                print(f'    找到 {len(options)} 个选项')
                                for opt in options:
                                    try:
                                        # 使用JS获取innerText
                                        opt_text = self.driver.execute_script('return arguments[0].innerText;', opt)
                                        if opt_text:
                                            opt_text = opt_text.strip()
                                            # 只在匹配到值时打印选项，减少日志
                                            if value in opt_text:
                                                print(f'    匹配到选项: {opt_text}')
                                                opt.click()
                                                print(f'  ✓ {field_name}：{value}')
                                                sleep(500)
                                                return True
                                    except:
                                        pass
                                break
                        except:
                            pass
                
                # 其他字段：参照JS逻辑，点击suffix按钮打开下拉
                form_items = self.driver.find_elements(By.CSS_SELECTOR, '.el-form-item')
                for item in form_items:
                    try:
                        label_elem = item.find_element(By.CSS_SELECTOR, '.el-form-item__label')
                        label = self.driver.execute_script('return arguments[0].innerText;', label_elem)
                        if label and field_name in label.strip():
                            print(f'    正在处理 {field_name}，目标值: {value}')
                            
                            # 滚动到可见
                            self.driver.execute_script('arguments[0].scrollIntoView({behavior: "smooth", block: "center"});', item)
                            sleep(800)
                            
                            # 点击 suffix 按钮打开下拉
                            suffix = item.find_element(By.CSS_SELECTOR, '.el-input__suffix')
                            if suffix:
                                self.driver.execute_script('arguments[0].click();', suffix)
                                print(f'    已点击{field_name}下拉')
                                sleep(3000)
                                
                                # 版型号需要滚动加载更多选项
                                if field_name == '版型号':
                                    for i in range(5):
                                        try:
                                            self.driver.execute_script('''
                                                const wrap = document.querySelector('.el-select-dropdown__wrap');
                                                if (wrap) wrap.scrollTop += 500;
                                            ''')
                                            sleep(500)
                                        except:
                                            pass
                                
                                # 查找并选择选项
                                opts = self.driver.find_elements(By.CSS_SELECTOR, '.el-select-dropdown__item')
                                print(f'    找到 {len(opts)} 个选项')
                                
                                # 先找完全匹配的选项
                                matched_opt = None
                                for opt in opts:
                                    try:
                                        text = self.driver.execute_script('return arguments[0].innerText;', opt)
                                        if text and text.strip() == value:
                                            matched_opt = opt
                                            break
                                    except:
                                        pass
                                
                                # 如果没有完全匹配，再找包含匹配（仅对非简单值）
                                if matched_opt is None:
                                    # 对于简单值（如"无"、"有"等单字），不使用包含匹配，避免误匹配
                                    simple_values = {'无', '有', '是', '否', '-', 'R', 'C'}
                                    if value not in simple_values:
                                        for opt in opts:
                                            try:
                                                text = self.driver.execute_script('return arguments[0].innerText;', opt)
                                                if text and value in text.strip():
                                                    matched_opt = opt
                                                    break
                                            except:
                                                pass
                                
                                if matched_opt:
                                    text = self.driver.execute_script('return arguments[0].innerText;', matched_opt)
                                    print(f'    匹配到选项: {text.strip()}')
                                    self.driver.execute_script('arguments[0].click();', matched_opt)
                                    print(f'  ✓ {field_name}: {text.strip()}')
                                    sleep(500)
                                    return True
                            break
                    except Exception as e:
                        print(f'    处理 {field_name} 时出错: {e}')
                        pass
            except Exception as e:
                print(f'    ⚠ {field_name} 第{retry+1}次尝试失败: {e}')
                if retry < max_retries - 1:
                    sleep(2000)
        print(f'    ✗ {field_name} 未能选中: {value}')
        return False
    
    def fill_input_field(self, field_name, value, max_retries=3):
        for retry in range(max_retries):
            try:
                form_items = self.driver.find_elements(By.CSS_SELECTOR, '.el-form-item')
                for item in form_items:
                    try:
                        label = item.find_element(By.CSS_SELECTOR, '.el-form-item__label').text.strip()
                        if field_name in label:
                            # 滚动到元素可见
                            self.driver.execute_script("arguments[0].scrollIntoView({behavior: 'smooth', block: 'center'});", item)
                            sleep(300)
                            
                            inputs = item.find_elements(By.CSS_SELECTOR, 'input, .el-input__inner')
                            if inputs:
                                # 用 JS 直接设置值，避免 invalid element state 问题
                                self.driver.execute_script('arguments[0].value = arguments[1]; arguments[0].dispatchEvent(new Event("input")); arguments[0].dispatchEvent(new Event("change"));', inputs[0], str(value))
                                print(f'    ✓ {field_name} 填写成功: {value}')
                                return True
                    except Exception as e:
                        print(f'    ⚠ 处理 {field_name} 输入框时出错: {e}')
                        pass
            except Exception as e:
                print(f'    ⚠ 第{retry+1}次尝试填写 {field_name} 失败: {e}')
                pass
            sleep(500)
        print(f'    ✗ {field_name} 未能填写: {value}')
        return False
    
    def fill_order_info(self):
        print('【3】填写订单基本信息...')
        
        # 从配置文件读取客商信息
        if hasattr(self, 'config'):
            customer_name = self.config.get('customerName', 'AIET算料')
            client_name = self.config.get('clientName', 'AIET算料')
        else:
            customer_name = 'AIET算料'
            client_name = 'AIET算料'
        tomorrow = (datetime.now() + timedelta(days=1)).strftime('%Y-%m-%d')
        
        # 客商选择 - 这是关键步骤，必须确保选中成功
        print(f'  开始选择客商：{customer_name}')
        customer_selected = self.fill_dropdown('客商', customer_name)
        
        # 额外验证客商是否选中成功
        if customer_selected:
            # 检查客商输入框的值
            form_items = self.driver.find_elements(By.CSS_SELECTOR, '.el-form-item')
            for item in form_items:
                try:
                    label_elem = item.find_element(By.CSS_SELECTOR, '.el-form-item__label')
                    label = self.driver.execute_script('return arguments[0].innerText;', label_elem)
                    if label and '客商' in label.strip():
                        select_input = item.find_element(By.CSS_SELECTOR, '.el-input__inner')
                        current_value = select_input.get_attribute('value')
                        if current_value and customer_name in current_value:
                            print(f'  ✓ 客商选中成功: {current_value}')
                        else:
                            print(f'  ⚠ 客商值可能未正确选中: {current_value}')
                            # 尝试再次选择
                            print('    尝试重新选择客商...')
                            customer_selected = self.fill_dropdown('客商', customer_name)
                    break
                except:
                    pass
        else:
            print(f'  ✗ 客商选择失败')
        
        self.fill_input_field('客户姓名', client_name)
        print(f'  ✓ 已填写客户姓名：{client_name}')
        
        # 交货日期 - 像 JS 那样通过 placeholder 查找输入框
        input_list3 = self.driver.find_elements(By.CSS_SELECTOR, '.el-input__inner')
        for input_elem in input_list3:
            try:
                placeholder = input_elem.get_attribute('placeholder')
                if placeholder and '交货' in placeholder:
                    self.driver.execute_script('arguments[0].value = arguments[1]; arguments[0].dispatchEvent(new Event("input")); arguments[0].dispatchEvent(new Event("change"));', input_elem, tomorrow)
                    print(f'  ✓ 已填写交货日期：{tomorrow}')
                    break
            except:
                pass
        
        sleep(1000)
        print('✓ 订单基本信息填写完成\n')
        return customer_selected
    
    def save_order_main(self):
        print('【4】保存订单主表...')
        
        buttons = self.driver.find_elements(By.CSS_SELECTOR, 'button, .el-button')
        for btn in buttons:
            try:
                btn_text = self.driver.execute_script('return arguments[0].innerText;', btn)
                if btn_text and btn_text.strip() == '保存':
                    btn.click()
                    print('  ✓ 已点击保存')
                    break
            except:
                pass
        
        sleep(2000)
        
        print('  等待保存结果...')
        save_success = False
        has_error = False
        error_msg = ''
        
        for j in range(8):
            sleep(1000)
            try:
                if j in [3, 8, 12]:
                    screenshot_path = f'e:\\111\\Material\\debug_save_{j}.png'
                    self.driver.save_screenshot(screenshot_path)
                    print(f'  📸 已保存截图：{screenshot_path}')
                
                msgs = self.driver.find_elements(By.CSS_SELECTOR, '.el-message, .el-message-box, .el-notification')
                msg_texts = [msg.text.strip() for msg in msgs if msg.text.strip()]
                if j < 10:
                    print(f'    提示：{json.dumps(msg_texts, ensure_ascii=False)}')
                
                for msg_text in msg_texts:
                    if '成功' in msg_text and '失败' not in msg_text:
                        save_success = True
                        break
                    if '不存在' in msg_text or '错误' in msg_text or '失败' in msg_text:
                        has_error = True
                        error_msg = msg_text
                        print(f'  ❌ 发现错误：{error_msg}')
                
                if save_success or has_error:
                    break
            except Exception as e:
                print(f'    检查提示时出错：{e}')
        
        if save_success:
            print('✓ 订单主表保存成功\n')
            # 获取生产单号
            production_no = ''
            try:
                production_no = self.driver.execute_script("""
                    const text = document.body.innerText;
                    const match = text.match(/\\*202\\d{6}/);
                    return match ? match[0] : '';
                """)
                if production_no:
                    print(f'  ✓ 生产单号：{production_no}')
                    # 保存到实例变量中
                    self.production_no = production_no
                    print(f'  ✓ 生产单号已保存到内存')
                else:
                    print('  ⚠ 未找到生产单号')
            except Exception as e:
                print(f'  ⚠ 获取生产单号失败：{e}')
        elif has_error:
            print(f'✗ 保存失败：{error_msg}\n')
        else:
            print('⚠ 等待超时\n')
        
        screenshot_path = f'e:\\111\\Material\\debug_final.png'
        self.driver.save_screenshot(screenshot_path)
        print(f'  📸 最终状态已保存：{screenshot_path}')
        
        return save_success
    
    def add_detail(self, banxing, chima_size, luocha=None, fabric_width=None, fabric_no=None, fabric_style=None):
        print('=' * 50)
        print(f'【5】新增版型：{banxing}')
        print('=' * 50)
        
        # 等待页面稳定
        sleep(1500)
        
        before_dialog_count = len(self.driver.find_elements(By.CSS_SELECTOR, '.el-dialog, .el-drawer'))
        print(f'    当前对话框数量：{before_dialog_count}')
        
        # 重新查找新增按钮，避免 stale element
        for retry in range(3):
            try:
                buttons = self.driver.find_elements(By.CSS_SELECTOR, 'button, .el-button')
                for btn in buttons:
                    try:
                        btn_text = self.driver.execute_script('return arguments[0].innerText;', btn)
                        if btn_text and btn_text.strip() == '新增':
                            btn.click()
                            print('  ✓ 已点击新增明细')
                            break
                    except:
                        pass
                break
            except Exception as e:
                print(f'    查找新增按钮失败，重试: {e}')
                sleep(500)
        
        print('  等待明细编辑对话框准备好...')
        dialog_ready = False
        for j in range(10):
            sleep(1000)
            try:
                dialog_count = len(self.driver.find_elements(By.CSS_SELECTOR, '.el-dialog, .el-drawer'))
                print(f'    对话框数量：{dialog_count}')
                if dialog_count > before_dialog_count:
                    print('  ✓ 明细编辑对话框已弹出')
                    dialog_ready = True
                    break
            except Exception as e:
                print(f'    检查对话框失败: {e}')
        
        if not dialog_ready:
            print('  ✓ 对话框已就绪（复用模式）')
        
        sleep(2500)
        print('  ✓ 等待元素渲染完成\n')
        
        print('【6】填写明细信息...')
        
        style_type = self.get_style_type(banxing)
        # 优先使用传入的参数，若无则从配置文件读取，最后使用默认值
        if hasattr(self, 'config'):
            fabric_width = fabric_width if fabric_width is not None else int(self.config.get('fabricMenfu', 74))
            fabric_no = fabric_no if fabric_no is not None else self.config.get('fabric', 'AI下单修改')
            lining_no = self.config.get('lining', 'AI下单修改')
            color = self.config.get('colorName', '10-黑')
            fabric_style = fabric_style if fabric_style is not None else self.config.get('fabricStyle', '平板')
            fabric_supply = self.config.get('fabricSupply', '库存面料')
            sleeve_lining = self.config.get('trousersLining', 'AI下单修改')
            libu_supply = self.config.get('liningSupply', '库存面料')
            libu_chengfen = self.config.get('liningComposition', '涤粘')
            xiukuli_chengfen = self.config.get('isLiningType', '涤粘')
            daoshunmao = self.config.get('bristle', '-')
            fabric_brand = self.config.get('fabricMark', '')
            niukou = self.config.get('button', '')
            chengfen = self.config.get('composition', '')
            fabric_chandi = self.config.get('fabricOrigin', '')
            xianhao = self.config.get('wireNumber', '')
            zhanchen = self.config.get('adhesiveLining', '')
            lalian = self.config.get('zipper', '')
            lingni = self.config.get('collar', '')
            linggua = self.config.get('collarHanger', '')
            libu_gongzhen = self.config.get('liningneedle', '')
            menjin_gongzhen = self.config.get('placketNeedle', '')
            xiaodingkouxian = self.config.get('smallThread', '')
            dadingkouxian = self.config.get('bigThread', '')
            fabric_mishu = self.config.get('sampleFabric', '')
            shifou_shiyang = self.config.get('isSample', '')
            beizhu = self.config.get('remark', '')
            banxing_tupian = ''
            gongyishu = ''
            banxing_jiegou = ''
            guamian_xiuhua = ''
            color_code = self.config.get('colorCode', '')
            fabric_quantity = self.config.get('fabricQuantity', '')
            half_delivery_date = self.config.get('halfDeliveryDate', '')
        else:
            fabric_width = fabric_width if fabric_width is not None else 74
            fabric_no = fabric_no if fabric_no is not None else 'AI下单修改'
            lining_no = 'AI下单修改'
            color = '10-黑'
            fabric_style = fabric_style if fabric_style is not None else '平板'
            fabric_supply = '库存面料'
            sleeve_lining = 'AI下单修改'
            libu_supply = '库存面料'
            libu_chengfen = '涤粘'
            xiukuli_chengfen = '涤粘'
            daoshunmao = '-'
            fabric_brand = ''
            niukou = ''
            chengfen = ''
            fabric_chandi = ''
            xianhao = ''
            zhanchen = ''
            lalian = ''
            lingni = ''
            linggua = ''
            libu_gongzhen = ''
            menjin_gongzhen = ''
            xiaodingkouxian = ''
            dadingkouxian = ''
            fabric_mishu = ''
            shifou_shiyang = ''
            beizhu = ''
            banxing_tupian = ''
            gongyishu = ''
            banxing_jiegou = ''
            guamian_xiuhua = ''
            color_code = ''
            fabric_quantity = ''
            half_delivery_date = ''
        
        print('  开始填写各字段...')
        self.fill_dropdown('款式类型', style_type)
        print('  ✓ 款式类型已填写')
        self.fill_dropdown('版型号', banxing)
        print('  ✓ 版型号已填写')
        
        # 面料编号
        self.set_input_by_label('面料编号', fabric_no)
        
        self.fill_dropdown('面料风格', fabric_style)
        print('  ✓ 面料风格已填写')
        
        # 面料门幅
        self.set_input_by_label('面料门幅', str(fabric_width))
        
        # 里布
        self.set_input_by_label('里布', lining_no)
        
        self.fill_dropdown('颜色', color)
        print('  ✓ 颜色已填写')
        
        # 袖裤里
        self.set_input_by_label('袖裤里', sleeve_lining)
        
        self.fill_dropdown('面料供应', fabric_supply)
        print('  ✓ 面料供应已填写')
        
        # 新增下拉框字段
        # 里布供应：如果有值才填写
        if libu_supply:
            self.fill_dropdown('里布供应', libu_supply)
            print('  ✓ 里布供应已填写')
        else:
            print('  ⏭ 里布供应：跳过（未配置）')
        
        # 里布成分：如果有值才填写
        if libu_chengfen:
            self.fill_dropdown('里布成分', libu_chengfen)
            print('  ✓ 里布成分已填写')
        else:
            print('  ⏭ 里布成分：跳过（未配置）')
        
        # 袖裤里成分：如果有值才填写
        if xiukuli_chengfen:
            self.fill_dropdown('袖裤里成分', xiukuli_chengfen)
            print('  ✓ 袖裤里成分已填写')
        else:
            print('  ⏭ 袖裤里成分：跳过（未配置）')
        
        # 面料倒顺毛：如果有值才填写
        if daoshunmao:
            self.fill_dropdown('面料倒顺毛', daoshunmao)
            print('  ✓ 面料倒顺毛已填写')
        else:
            print('  ⏭ 面料倒顺毛：跳过（未配置）')
        
        # 新增输入框字段
        self.set_input_by_label('面料品牌', fabric_brand)
        self.set_input_by_label('纽扣', niukou)
        self.set_input_by_label('成分', chengfen)
        self.set_input_by_label('面料产地', fabric_chandi)
        self.set_input_by_label('线号', xianhao)
        self.set_input_by_label('粘衬', zhanchen)
        self.set_input_by_label('拉链', lalian)
        self.set_input_by_label('领呢', lingni)
        self.set_input_by_label('领挂', linggua)
        self.set_input_by_label('里布贡针', libu_gongzhen)
        self.set_input_by_label('门襟贡针', menjin_gongzhen)
        self.set_input_by_label('小钉扣线', xiaodingkouxian)
        self.set_input_by_label('大钉扣线', dadingkouxian)
        # 面料到货米数：如果有值才填写
        if fabric_mishu:
            self.set_input_by_label('面料到货米数', fabric_mishu)
        else:
            print('  ⏭ 面料到货米数：跳过（未配置）')
        
        # 是否试样：如果有值才填写
        if shifou_shiyang:
            self.set_input_by_label('是否试样', shifou_shiyang)
        else:
            print('  ⏭ 是否试样：跳过（未配置）')
        
        # 备注：如果有值才填写
        if beizhu:
            self.set_input_by_label('备注', beizhu)
        else:
            print('  ⏭ 备注：跳过（未配置）')
        
        # 版型图片：如果有值才填写
        if banxing_tupian:
            self.set_input_by_label('版型图片', banxing_tupian)
        else:
            print('  ⏭ 版型图片：跳过（未配置）')
        
        # 工艺书：如果有值才填写
        if gongyishu:
            self.set_input_by_label('工艺书', gongyishu)
        else:
            print('  ⏭ 工艺书：跳过（未配置）')
        
        # 版型结构：如果有值才填写
        if banxing_jiegou:
            self.set_input_by_label('版型结构', banxing_jiegou)
        else:
            print('  ⏭ 版型结构：跳过（未配置）')
        
        # 挂面绣花：如果有值才填写
        if guamian_xiuhua:
            self.set_input_by_label('挂面绣花', guamian_xiuhua)
        else:
            print('  ⏭ 挂面绣花：跳过（未配置）')
        
        # 颜色代码：如果有值才填写
        if color_code:
            self.set_input_by_label('颜色代码', color_code)
        else:
            print('  ⏭ 颜色代码：跳过（未配置）')
        
        # 面料数量：如果有值才填写
        if fabric_quantity:
            self.set_input_by_label('面料数量', fabric_quantity)
        else:
            print('  ⏭ 面料数量：跳过（未配置）')
        
        # 半交期：如果有值才填写
        if half_delivery_date:
            self.set_input_by_label('半交期', half_delivery_date)
        else:
            print('  ⏭ 半交期：跳过（未配置）')
        
        # 落差：必须填写
        self.set_input_by_label('落差', luocha)
        
        # 尺码：必须填写
        self.set_input_by_label('尺码', chima_size)
        
        # 面辅料价格
        self.set_input_by_label('面辅料价格', '0')
        
        # 加工费
        self.set_input_by_label('加工费', '0')
        
        sleep(1000)
        print('✓ 明细信息填写完成\n')
        return True
    
    def set_input_by_label(self, label_text, value):
        """通过标签文本设置输入框值"""
        inputs = self.driver.find_elements(By.CSS_SELECTOR, '.el-form-item .el-input__inner')
        for input_elem in inputs:
            try:
                label_elem = input_elem.find_element(By.XPATH, './ancestor::div[contains(@class, "el-form-item")]//*[contains(@class, "el-form-item__label")]')
                label = self.driver.execute_script('return arguments[0].innerText;', label_elem)
                if label and label.strip() == label_text:
                    self.driver.execute_script('arguments[0].value = arguments[1]; arguments[0].dispatchEvent(new Event("input")); arguments[0].dispatchEvent(new Event("change"));', input_elem, str(value))
                    print(f'  ✓ {label_text}：{value}')
                    return True
            except:
                pass
        print(f'  ⚠ 未找到 {label_text} 输入框')
        return False
    
    def click_save_in_correct_dialog(self, dialog_title=None, required_fields=None, excluded_fields=None, button_text=['保存']):
        """
        在正确的对话框中点击保存按钮
        :param dialog_title: 对话框标题（精确匹配）
        :param required_fields: 对话框中必须包含的字段标签列表
        :param excluded_fields: 对话框中绝对不能包含的字段标签列表（用于排除其他对话框）
        :param button_text: 保存按钮的文字（可以是列表）
        :return: True表示成功点击，False表示失败
        """
        print(f'  🔍 正在查找正确的对话框...')
        print(f'    目标标题: {dialog_title}')
        print(f'    目标字段: {required_fields}')
        print(f'    排除字段: {excluded_fields}')
        print(f'    目标按钮: {button_text}')
        
        if isinstance(button_text, str):
            button_text = [button_text]
        
        result = self.driver.execute_script("""
            const dialogs = document.querySelectorAll('.el-dialog, .el-drawer');
            const targetTitle = arguments[0];
            const requiredFields = arguments[1] || [];
            const excludedFields = arguments[2] || [];
            const btnTexts = arguments[3] || ['保存'];
            
            console.log('=== 所有可用对话框 ===');
            
            const dialogInfo = [];
            
            for (let i = 0; i < dialogs.length; i++) {
                const dialog = dialogs[i];
                
                // 检查对话框是否可见
                const style = window.getComputedStyle(dialog);
                if (style.display === 'none' || style.visibility === 'hidden') {
                    console.log(`对话框${i}: 隐藏跳过`);
                    continue;
                }
                
                // 获取对话框标题（尝试多种选择器）
                let title = '';
                const titleSelectors = [
                    '.el-dialog__title', 
                    '.el-drawer__title', 
                    '.el-dialog__header span', 
                    '.el-dialog__header .el-dialog__title'
                ];
                for (const selector of titleSelectors) {
                    const titleElem = dialog.querySelector(selector);
                    if (titleElem && titleElem.innerText.trim()) {
                        title = titleElem.innerText.trim();
                        break;
                    }
                }
                
                // 获取对话框内所有字段标签
                const labels = [];
                dialog.querySelectorAll('.el-form-item__label').forEach(l => {
                    labels.push(l.innerText.trim());
                });
                
                // 获取对话框内所有按钮文字
                const buttonLabels = [];
                dialog.querySelectorAll('button, .el-button').forEach(b => {
                    buttonLabels.push(b.innerText.trim());
                });
                
                console.log(`对话框${i}: 标题="${title}", 字段=${labels.slice(0, 5)}, 按钮=${buttonLabels}`);
                dialogInfo.push({ index: i, title, labels, buttons: buttonLabels });
            }
            
            // 先打印所有对话框信息（用于调试）
            for (const dlg of dialogInfo) {
                console.log(`  - 对话框${dlg.index}: 标题="${dlg.title}", 字段=${dlg.labels.slice(0, 8)}, 按钮=${dlg.buttons}`);
            }
            
            // 再按顺序检查匹配
            for (const dlg of dialogInfo) {
                const i = dlg.index;
                const dialog = dialogs[i];
                
                // 检查是否匹配标题
                let titleMatch = !targetTitle;  // 如果没指定标题，则匹配所有
                if (targetTitle) {
                    titleMatch = dlg.title === targetTitle;
                }
                
                // 检查是否包含必需字段
                let fieldsMatch = requiredFields.length === 0;  // 如果没指定字段，则匹配所有
                if (requiredFields.length > 0) {
                    if (requiredFields.length <= 2) {
                        // 字段较少时（比如量体信息），要求所有字段都存在
                        fieldsMatch = requiredFields.every(req => 
                            dlg.labels.some(l => l.includes(req))
                        );
                    } else {
                        // 字段较多时（比如定制选项），只要包含其中一个字段就行
                        fieldsMatch = requiredFields.some(req => 
                            dlg.labels.some(l => l.includes(req))
                        );
                    }
                }
                
                // 检查是否包含排除字段（如果包含任何排除字段，则不匹配）
                let excludeMatch = excludedFields.length === 0;  // 如果没指定排除字段，则不排除
                if (excludedFields.length > 0) {
                    // 如果对话框包含任何排除字段，则不匹配
                    const hasExcludedField = excludedFields.some(exc => 
                        dlg.labels.some(l => l.includes(exc))
                    );
                    excludeMatch = !hasExcludedField;
                }
                
                if (titleMatch && fieldsMatch && excludeMatch) {
                    console.log(`✅ 找到目标对话框${i}: ${dlg.title}`);
                    
                    // 在对话框内查找保存按钮（支持多个按钮文字）
                    const buttons = dialog.querySelectorAll('button, .el-button');
                    for (const btn of buttons) {
                        const btnTxt = btn.innerText.trim();
                        if (btnTexts.includes(btnTxt)) {
                            if (btn.disabled || btn.classList.contains('is-disabled')) {
                                console.log(`❌ ${btnTxt}按钮被禁用`);
                                continue;
                            }
                            
                            // 滚动到按钮并点击
                            btn.scrollIntoView({ block: 'center', behavior: 'smooth' });
                            btn.click();
                            console.log(`✅ 已点击对话框${i}中的"${btnTxt}"按钮`);
                            return { success: true, dialogIndex: i, dialogTitle: dlg.title, dialogInfo, matchedLabels: dlg.labels };
                        }
                    }
                    
                    console.log(`❌ 对话框${i}中未找到匹配按钮`);
                    return { success: false, error: `对话框${i}中未找到匹配按钮`, dialogIndex: i, dialogTitle: dlg.title, dialogInfo };
                }
            }
            
            // 如果用标题/字段匹配不到，尝试查找包含特定文本的简单提示弹窗（如"是否确认"）
            const alertModals = document.querySelectorAll('.el-message-box, .el-dialog');
            for (const modal of alertModals) {
                const style = window.getComputedStyle(modal);
                if (style.display === 'none' || style.visibility === 'hidden') continue;
                
                const modalText = modal.innerText || '';
                // 检查是否是确认弹窗（包含"确认"、"确定"、"是否"等关键词）
                if (modalText.includes('确认') || modalText.includes('确定') || modalText.includes('是否')) {
                    console.log(`✅ 找到提示弹窗，内容包含确认关键词`);
                    
                    const buttons = modal.querySelectorAll('button, .el-button');
                    for (const btn of buttons) {
                        const btnTxt = btn.innerText.trim();
                        if (btnTexts.includes(btnTxt)) {
                            if (btn.disabled || btn.classList.contains('is-disabled')) continue;
                            
                            btn.scrollIntoView({ block: 'center' });
                            btn.click();
                            console.log(`✅ 已点击提示弹窗中的"${btnTxt}"按钮`);
                            return { success: true, dialogTitle: '提示弹窗', dialogInfo };
                        }
                    }
                }
            }
            
            console.log('❌ 未找到匹配的对话框，尝试匹配最后一个可见对话框...');
            
            // 如果用标题/字段匹配不到，就取最后一个可见的对话框（最上层的）
            if (dialogInfo.length > 0) {
                const lastDialogIndex = dialogInfo.length - 1;
                const lastDialogData = dialogInfo[lastDialogIndex];
                const dialog = dialogs[lastDialogData.index];
                
                console.log(`✅ 回退匹配：使用最后一个可见对话框${lastDialogData.index}`);
                
                // 在最后一个对话框中查找保存按钮
                const buttons = dialog.querySelectorAll('button, .el-button');
                for (const btn of buttons) {
                    const btnTxt = btn.innerText.trim();
                    if (btnTexts.includes(btnTxt)) {
                        if (btn.disabled || btn.classList.contains('is-disabled')) {
                            console.log(`❌ ${btnTxt}按钮被禁用`);
                            continue;
                        }
                        
                        // 滚动到按钮并点击
                        btn.scrollIntoView({ block: 'center', behavior: 'smooth' });
                        btn.click();
                        console.log(`✅ 已点击最后一个可见对话框${lastDialogData.index}中的"${btnTxt}"按钮`);
                        return { success: true, dialogIndex: lastDialogData.index, dialogTitle: lastDialogData.title, dialogInfo, matchedLabels: lastDialogData.labels, fallbackMatch: true };
                    }
                }
            }
            
            return { success: false, error: '未找到匹配的对话框，也没有可用的最后一个对话框', dialogInfo };
        """, dialog_title, required_fields, excluded_fields, button_text)
        
        if result and result.get('success'):
            print(f'  ✅ 成功点击对话框{result.get("dialogIndex")}的按钮')
            print(f'    对话框标题: "{result.get("dialogTitle", "未知")}"')
            print(f'    对话框字段: {result.get("matchedLabels", [])[:10]}')
            return True
        else:
            print(f'  ❌ 保存失败: {result.get("error", "未知错误")}')
            # 打印所有找到的对话框信息帮助调试
            dialog_info = result.get('dialogInfo', [])
            if dialog_info:
                print(f'  📋 找到的所有对话框:')
                for dlg in dialog_info:
                    print(f'    - 对话框{dlg["index"]}: 标题="{dlg["title"]}", 字段={dlg["labels"][:10]}, 按钮={dlg["buttons"]}')
            return False
    
    def edit_custom_options(self, custom_options):
        """修改定制选项"""
        print('【6.5】修改定制选项...')
        
        if not custom_options:
            print('  - 没有定制选项需要修改')
            return
        
        # 滚动页面找到定制选项区域
        for i in range(5):
            self.driver.execute_script('window.scrollBy(0, 200);')
            sleep(300)
        
        # 查找定制选项字段的编辑按钮
        edit_button_found = False
        form_items = self.driver.find_elements(By.CSS_SELECTOR, '.el-form-item')
        
        for item in form_items:
            try:
                label_elem = item.find_element(By.CSS_SELECTOR, '.el-form-item__label')
                label = self.driver.execute_script('return arguments[0].innerText;', label_elem)
                if label and '定制选项' in label.strip():
                    # 找到定制选项字段，查找右侧的编辑按钮
                    edit_buttons = item.find_elements(By.CSS_SELECTOR, 'button, .el-button, .el-icon-edit, .edit-btn')
                    for btn in edit_buttons:
                        try:
                            btn_text = self.driver.execute_script('return arguments[0].innerText;', btn)
                            btn_class = btn.get_attribute('class')
                            # 判断是否是编辑按钮
                            if btn_text and ('编辑' in btn_text or 'Edit' in btn_text) or \
                               'edit' in btn_class.lower() or \
                               'el-icon-edit' in btn_class:
                                self.driver.execute_script('arguments[0].scrollIntoView(true);', btn)
                                sleep(500)
                                self.driver.execute_script('arguments[0].click();', btn)
                                print('  ✓ 已点击定制选项编辑按钮')
                                edit_button_found = True
                                break
                        except:
                            pass
                    if edit_button_found:
                        break
            except:
                pass
        
        if not edit_button_found:
            # 尝试另一种方式：查找所有包含"编辑"的按钮
            print('  尝试查找编辑按钮...')
            all_buttons = self.driver.find_elements(By.CSS_SELECTOR, 'button, .el-button')
            for btn in all_buttons:
                try:
                    btn_text = self.driver.execute_script('return arguments[0].innerText;', btn)
                    if btn_text and '编辑' in btn_text:
                        # 检查这个按钮附近是否有定制选项相关的标签
                        parent = btn.find_element(By.XPATH, './ancestor::div[contains(@class, "el-form-item")]')
                        label_elem = parent.find_element(By.CSS_SELECTOR, '.el-form-item__label')
                        label = self.driver.execute_script('return arguments[0].innerText;', label_elem)
                        if '定制选项' in label:
                            self.driver.execute_script('arguments[0].scrollIntoView(true);', btn)
                            sleep(500)
                            self.driver.execute_script('arguments[0].click();', btn)
                            print('  ✓ 已点击定制选项编辑按钮')
                            edit_button_found = True
                            break
                except:
                    pass
        
        if not edit_button_found:
            print('  ! 未找到定制选项编辑按钮，跳过定制选项修改')
            return
        
        # 等待定制选项对话框弹出
        print('  等待定制选项对话框弹出...')
        for i in range(10):
            sleep(1000)
            dialogs = self.driver.find_elements(By.CSS_SELECTOR, '.el-dialog, .el-drawer')
            if len(dialogs) > 0:
                print(f'  ✓ 定制选项对话框已弹出，共 {len(dialogs)} 个对话框')
                break
        
        # 根据定制选项配置进行修改
        print(f'  开始修改定制选项：{custom_options}')
        for option_name, option_value in custom_options.items():
            print(f'    正在设置 {option_name}：{option_value}')
            self.fill_dropdown(option_name, option_value)
        
        # 使用通用函数在"定制选项"对话框中点击保存
        print('  保存定制选项...')
        # 定制选项对话框可能包含多种字段（上衣、西裤、马甲、大衣、猎装等）
        # 只要包含其中一个字段就匹配
        save_success = self.click_save_in_correct_dialog(
            required_fields=['工艺', '上衣后叉', '手巾袋', '脚口', '腰款式', '表袋', '马甲后背', '马甲后叉', '大衣面袋', '大衣袖扣', '袖口锁眼', '后背款式', '门襟锁眼', '半里', '驳头锁眼', '猎装腰带款'],
            excluded_fields=['款式类型', '流水号', '版型号']  # 排除明细对话框
        )
        
        if not save_success:
            print('  ✗ 保存失败！程序暂停，请检查...')
            input('  按回车键继续（或关闭程序）...')
            return
        
        # 等待对话框关闭
        print('  等待定制选项对话框关闭...')
        dialog_closed = False
        for i in range(15):
            sleep(1000)
            # 检查定制选项对话框是否关闭（通过标题检查）
            dialog_closed = self.driver.execute_script("""
                const dialogs = document.querySelectorAll('.el-dialog, .el-drawer');
                for (const dialog of dialogs) {
                    const style = window.getComputedStyle(dialog);
                    if (style.display === 'none' || style.visibility === 'hidden') {
                        continue;
                    }
                    const titleElem = dialog.querySelector('.el-dialog__title, .el-drawer__title');
                    if (titleElem) {
                        const title = titleElem.innerText.trim();
                        if (title === '定制选项') {
                            return false; // 定制选项对话框还在
                        }
                    }
                }
                return true;
            """)
            if dialog_closed:
                print('  ✓ 定制选项对话框已关闭')
                break
            if i % 5 == 4:
                print(f'    已等待{i+1}秒...')
        
        if not dialog_closed:
            print('  ⚠ 警告：定制选项对话框可能未完全关闭')
        
        print('✓ 定制选项修改完成\n')
    
    def enter_liangti_info(self, chima_size, style_type, luocha=None, liangti_data=None):
        print('【7】填写量体信息...')
        
        # 确定尺码和落差的值 - 使用传入的参数
        target_chima = chima_size if chima_size else '50'
        target_luocha = 'R'
        
        # 如果传入了luocha参数，使用传入的值
        if luocha is not None:
            target_luocha = luocha
        elif '裤' in style_type:
            target_luocha = '-'
        
        print(f'  目标尺码：{target_chima}, 目标落差：{target_luocha}')
        
        # 如果有量体数据，打印出来
        if liangti_data:
            print(f'  量体数据：{liangti_data}')
        
        for i in range(5):
            self.driver.execute_script("""
                document.querySelectorAll('.el-dialog__body, .el-drawer__body').forEach(dialogBody => {
                    dialogBody.scrollTop = dialogBody.scrollHeight;
                });
                window.scrollTo(0, document.body.scrollHeight);
            """)
            sleep(600)
            print(f'    第{i+1}次滚动')
        
        sleep(2000)
        
        liangti_edit_clicked = False
        form_items = self.driver.find_elements(By.CSS_SELECTOR, '.el-form-item')
        for item in form_items:
            try:
                label_elem = item.find_element(By.CSS_SELECTOR, '.el-form-item__label')
                label = self.driver.execute_script('return arguments[0].innerText;', label_elem)
                if label and '量体信息' in label.strip():
                    edit_btns = item.find_elements(By.CSS_SELECTOR, 'button, .el-button')
                    for edit_btn in edit_btns:
                        try:
                            btn_text = self.driver.execute_script('return arguments[0].innerText;', edit_btn)
                            if btn_text and btn_text.strip() == '编辑':
                                self.driver.execute_script("arguments[0].click();", edit_btn)
                                liangti_edit_clicked = True
                                print('  ✓ 找到量体信息编辑按钮并点击')
                                break
                        except:
                            pass
                    break
            except:
                pass
        
        if not liangti_edit_clicked:
            all_buttons = self.driver.find_elements(By.CSS_SELECTOR, 'button, .el-button')
            for btn in all_buttons:
                try:
                    btn_text = self.driver.execute_script('return arguments[0].innerText;', btn)
                    if btn_text and btn_text.strip() == '编辑':
                        self.driver.execute_script("arguments[0].click();", btn)
                        liangti_edit_clicked = True
                        print('  ✓ 点击了编辑按钮')
                        break
                except:
                    pass
        
        sleep(3000)
        
        print('  选择尺码和落差...')
        
        # 选择落差 - 像JS那样
        print('  选择落差...')
        luocha_found = False
        
        # 先找到量体信息对话框
        liangti_dialog = None
        dialogs = self.driver.find_elements(By.CSS_SELECTOR, '.el-dialog, .el-drawer')
        for dialog in dialogs:
            try:
                labels = dialog.find_elements(By.CSS_SELECTOR, '.el-form-item__label')
                label_texts = [l.text.strip() for l in labels if l.text.strip()]
                if any('尺码' in l or '落差' in l or '量体' in l for l in label_texts):
                    liangti_dialog = dialog
                    print('  ✓ 找到量体信息对话框')
                    break
            except:
                pass
        
        # 查找落差字段
        luocha_options = []
        form_items2 = self.driver.find_elements(By.CSS_SELECTOR, '.el-form-item')
        for item in form_items2:
            try:
                label_elem = item.find_element(By.CSS_SELECTOR, '.el-form-item__label')
                label = self.driver.execute_script('return arguments[0].innerText;', label_elem)
                if label and label.strip() == '落差':
                    # 点击suffix打开下拉
                    suffix = item.find_element(By.CSS_SELECTOR, '.el-input__suffix')
                    if suffix:
                        self.driver.execute_script("arguments[0].click();", suffix)
                        sleep(2000)
                        # 选择落差选项 - 优先在量体对话框内查找
                        options = []
                        if liangti_dialog:
                            try:
                                options = liangti_dialog.find_elements(By.CSS_SELECTOR, '.el-select-dropdown__item')
                            except:
                                pass
                        # 如果在对话框内没找到，再全局查找
                        if not options:
                            options = self.driver.find_elements(By.CSS_SELECTOR, '.el-select-dropdown__item')
                        
                        print(f'    找到 {len(options)} 个落差选项')
                        
                        # 收集所有可用的落差选项
                        available_luocha = []
                        for opt in options:
                            try:
                                opt_text = self.driver.execute_script('return arguments[0].innerText;', opt)
                                if opt_text and opt_text.strip() in ['R', 'C', '-']:
                                    available_luocha.append({'elem': opt, 'text': opt_text.strip()})
                            except:
                                pass
                        
                        print(f'    可用落差选项：{[o["text"] for o in available_luocha]}')
                        
                        # 按优先级选择：先选目标，再按 R -> C -> - 顺序尝试
                        luocha_priority = ['R', 'C', '-']
                        
                        # 如果目标在优先级列表中，放到最前面
                        if target_luocha in luocha_priority:
                            luocha_priority.remove(target_luocha)
                            luocha_priority.insert(0, target_luocha)
                        
                        # 按优先级尝试选择
                        for priority_luocha in luocha_priority:
                            for opt in available_luocha:
                                if opt['text'] == priority_luocha:
                                    self.driver.execute_script("arguments[0].click();", opt['elem'])
                                    print(f'  ✓ 落差选择：{priority_luocha}')
                                    luocha_found = True
                                    break
                            if luocha_found:
                                break
                    break
            except:
                pass
        
        if not luocha_found:
            print(f'  ! 未成功选择落差 {target_luocha}，尝试选第一个可用选项')
            options = self.driver.find_elements(By.CSS_SELECTOR, '.el-select-dropdown__item')
            if len(options) > 0:
                first_opt_text = self.driver.execute_script('return arguments[0].innerText;', options[0])
                self.driver.execute_script("arguments[0].click();", options[0])
                print(f'  ✓ 已选择第一个落差选项：{first_opt_text.strip()}')
        
        sleep(2000)
        
        # 选择尺码
        print('  选择尺码...')
        chima_found = False
        # 查找尺码字段
        form_items3 = self.driver.find_elements(By.CSS_SELECTOR, '.el-form-item')
        for item in form_items3:
            try:
                label_elem = item.find_element(By.CSS_SELECTOR, '.el-form-item__label')
                label = self.driver.execute_script('return arguments[0].innerText;', label_elem)
                if label and label.strip() == '尺码':
                    # 点击suffix打开下拉
                    suffix = item.find_element(By.CSS_SELECTOR, '.el-input__suffix')
                    if suffix:
                        self.driver.execute_script("arguments[0].click();", suffix)
                        sleep(2000)
                        # 选择尺码选项 - 优先在量体对话框内查找
                        options = []
                        if liangti_dialog:
                            try:
                                options = liangti_dialog.find_elements(By.CSS_SELECTOR, '.el-select-dropdown__item')
                            except:
                                pass
                        # 如果在对话框内没找到，再全局查找
                        if not options:
                            options = self.driver.find_elements(By.CSS_SELECTOR, '.el-select-dropdown__item')
                        
                        print(f'    找到 {len(options)} 个尺码选项')
                        for opt in options:
                            try:
                                opt_text = self.driver.execute_script('return arguments[0].innerText;', opt)
                                if opt_text and opt_text.strip() == target_chima:
                                    self.driver.execute_script("arguments[0].click();", opt)
                                    print(f'  ✓ 尺码选择：{target_chima}')
                                    chima_found = True
                                    break
                            except:
                                pass
                    break
            except:
                pass
        
        sleep(2000)
        
        # 填写具体的量体数据
        if liangti_data and len(liangti_data) > 0:
            print('  填写具体量体数据...')
            # 加载量体部位映射表
            import os
            script_dir = os.path.dirname(os.path.abspath(__file__))
            mapping_file = os.path.join(script_dir, '映射配置.json')
            liangti_mapping = {}
            try:
                with open(mapping_file, 'r', encoding='utf-8-sig') as f:
                    mapping = json.load(f)
                    liangti_mapping = mapping.get('量体部位映射', {})
                print(f'  ✓ 量体部位映射表加载成功，共 {len(liangti_mapping)} 个部位')
            except Exception as e:
                print(f'  ! 量体部位映射表加载失败：{e}')
            
            print(f'  待处理的量体数据：{liangti_data}')
            
            # 遍历量体数据，填写每个字段
            for field_code, field_value in liangti_data.items():
                # 获取字段的中文名称
                field_aliases = {
                    'hips': ['hip', '臀围', '全臀围', '净臀围', 'Hip', 'Full Hip'],
                    'hip': ['hips', '臀围', '全臀围', '净臀围', 'Hip', 'Full Hip'],
                }
                mapped_info = liangti_mapping.get(field_code, {})
                field_name = mapped_info.get('名称', field_code)
                field_candidates = [field_name, field_code]
                english_name = mapped_info.get('英文')
                if english_name:
                    field_candidates.append(english_name.strip())
                for alias in field_aliases.get(field_code, []):
                    field_candidates.append(alias)
                if field_name.startswith('全') and len(field_name) > 1:
                    field_candidates.append(field_name[1:])
                if field_name.startswith('净') and len(field_name) > 1:
                    field_candidates.append(field_name[1:])
                field_candidates = [name for name in dict.fromkeys(field_candidates) if name]
                print(f'    处理量体字段：[{field_code}] -> [{field_name}] = {field_value}')
                
                # 查找量体字段 - 使用 ElementUI InputNumber 组件，点击 + - 按钮
                field_found = False
                try:
                    # 解析 value
                    # 如果值以"-"开头，直接作为调整值
                    # 否则作为目标成衣尺寸，需要计算调整值
                    field_value_str = str(field_value).strip()
                    is_direct_adjust = field_value_str.startswith('-')
                    if is_direct_adjust:
                        target_cheng_yi_size = float(field_value_str)
                    else:
                        target_cheng_yi_size = float(field_value_str.replace('+', ''))
                    
                    # 使用JS查找并点击 - ElementUI 表格的表头/表体通常是分离的，必须按列名定位
                    fill_result = self.driver.execute_script("""
                        const fieldName = arguments[0];
                        const targetChengYiSize = arguments[1];
                        const fieldCandidates = arguments[2] || [fieldName];
                        
                        console.log('========== 开始查找量体字段：', fieldName, fieldCandidates, '目标成衣尺寸:', targetChengYiSize, '==========');
                        
                        // 1. 先找到量体信息对话框。页面会同时保留多个弹窗 DOM，必须按量体关键词和层级选最像的那个。
                        let liangtiDialog = null;
                        const dialogs = document.querySelectorAll('.el-dialog, .el-drawer');
                        console.log('找到', dialogs.length, '个对话框');
                        const visibleDialogs = [];

                        for (let i = 0; i < dialogs.length; i++) {
                            const dialog = dialogs[i];
                            const dialogStyle = window.getComputedStyle(dialog);
                            if (dialogStyle.display === 'none' || dialogStyle.visibility === 'hidden') continue;
                            const rect = dialog.getBoundingClientRect();
                            if (rect.width <= 0 || rect.height <= 0) continue;

                            const title = dialog.querySelector('.el-dialog__title, .el-drawer__title');
                            const titleText = title ? title.textContent.trim() : '';
                            const dialogText = dialog.innerText || dialog.textContent || '';
                            const zIndex = parseInt(dialogStyle.zIndex || '0', 10) || 0;

                            let score = 0;
                            if (titleText.includes('录入量体数据')) score += 200;
                            if (titleText.includes('量体')) score += 120;
                            if (dialogText.includes('量体部位')) score += 100;
                            if (dialogText.includes('调整尺寸')) score += 80;
                            if (dialogText.includes('标准尺寸')) score += 40;
                            if (dialogText.includes('特体部位')) score += 40;
                            if (dialogText.includes('落差') && dialogText.includes('尺码')) score += 30;
                            if (dialogText.includes(fieldName)) score += 30;
                            if (dialogText.includes('驳头类型') || dialogText.includes('纽扣数量') || dialogText.includes('里布结构')) score -= 80;

                            visibleDialogs.push({ dialog, index: i, titleText, score, zIndex, textPreview: dialogText.slice(0, 120) });
                        }

                        visibleDialogs.sort((a, b) => (b.score - a.score) || (b.zIndex - a.zIndex) || (b.index - a.index));
                        for (const info of visibleDialogs) {
                            console.log('对话框候选[' + info.index + ']: 标题="' + info.titleText + '", score=' + info.score + ', zIndex=' + info.zIndex + ', 内容预览="' + info.textPreview + '"');
                        }
                        const bestDialog = visibleDialogs.find(info => info.score > 0);
                        if (bestDialog) {
                            liangtiDialog = bestDialog.dialog;
                            console.log('✓ 选择量体对话框:', bestDialog.index, bestDialog.titleText, 'score=', bestDialog.score);
                        }
                        
                        // 如果没找到对话框，尝试全局查找
                        const searchScope = liangtiDialog || document;
                        console.log('搜索范围:', liangtiDialog ? '量体对话框内' : '全局');
                        
                        const getText = (el) => (el ? (el.innerText || el.textContent || '').trim() : '');
                        const normalizeText = (text) => (text || '').replace(/\\s+/g, '').replace(/[()（）]/g, '').trim();
                        const normalizedCandidates = fieldCandidates.map(normalizeText).filter(Boolean);
                        const textMatches = (text) => {
                            const normalizedText = normalizeText(text);
                            if (!normalizedText) return false;
                            return normalizedCandidates.some(candidate =>
                                normalizedText === candidate ||
                                normalizedText.includes(candidate) ||
                                candidate.includes(normalizedText)
                            );
                        };
                        const isVisible = (el) => {
                            if (!el) return false;
                            const style = window.getComputedStyle(el);
                            const rect = el.getBoundingClientRect();
                            return style.display !== 'none' && style.visibility !== 'hidden' && rect.width > 0 && rect.height > 0;
                        };
                        const findColumnIndex = (headers, keywords) => {
                            for (let i = 0; i < headers.length; i++) {
                                const text = getText(headers[i]).replace(/\\s+/g, '');
                                console.log('表头:', i, text);
                                if (keywords.some(keyword => text.includes(keyword))) return i;
                            }
                            return -1;
                        };
                        
                        // 2. ElementUI el-table 的 header/body 是两个 table，先按 el-table 容器精确定位
                        let targetRow = null;
                        let targetCellIndex = -1;
                        let targetCell = null;
                        let taomaValue = 0;
                        const visiblePartNames = [];
                        
                        console.log('方法1: 按 ElementUI 表格列名查找...');
                        const elTables = Array.from(searchScope.querySelectorAll('.el-table')).filter(isVisible);
                        console.log('找到', elTables.length, '个可见 el-table');
                        
                        for (let table of elTables) {
                            let headers = table.querySelectorAll('.el-table__header-wrapper th');
                            if (!headers.length) {
                                headers = table.querySelectorAll('.el-table__fixed-header-wrapper th');
                            }
                            if (!headers.length) continue;
                            
                            const liangtiColumnIndex = findColumnIndex(headers, ['量体部位', '部位']);
                            const adjustColumnIndex = findColumnIndex(headers, ['调整尺寸', '调整']);
                            const taomaColumnIndex = findColumnIndex(headers, ['套码', '标准尺寸']);
                            
                            if (liangtiColumnIndex >= 0 && adjustColumnIndex >= 0) {
                                console.log('✓ 找到量体表格列，量体部位列:', liangtiColumnIndex, '调整尺寸列:', adjustColumnIndex, '套码列:', taomaColumnIndex);
                                
                                const wrappers = Array.from(table.querySelectorAll('.el-table__body-wrapper'));
                                if (!wrappers.length) {
                                    wrappers.push(...table.querySelectorAll('.el-table__fixed-body-wrapper'));
                                }
                                const scrollPositions = [0, 0.15, 0.3, 0.45, 0.6, 0.75, 0.9, 1];
                                for (let wrapper of wrappers) {
                                    for (let pos of scrollPositions) {
                                        if (wrapper.scrollHeight > wrapper.clientHeight) {
                                            wrapper.scrollTop = Math.floor((wrapper.scrollHeight - wrapper.clientHeight) * pos);
                                        }
                                        const rows = wrapper.querySelectorAll('tbody tr');
                                        for (let i = 0; i < rows.length; i++) {
                                            const cells = rows[i].querySelectorAll('td, th');
                                            if (cells.length > liangtiColumnIndex) {
                                                const cellText = getText(cells[liangtiColumnIndex]);
                                                const normalizedCellText = normalizeText(cellText);
                                                if (normalizedCellText) visiblePartNames.push(normalizedCellText);
                                                console.log('检查行:', i, normalizedCellText);
                                                if (textMatches(cellText)) {
                                                    targetRow = rows[i];
                                                    targetCellIndex = adjustColumnIndex;
                                                    targetCell = cells[adjustColumnIndex] || null;
                                                    if (taomaColumnIndex >= 0 && cells.length > taomaColumnIndex) {
                                                        const taomaText = getText(cells[taomaColumnIndex]);
                                                        taomaValue = parseFloat(taomaText) || 0;
                                                        console.log('✓ 获取到套码值:', taomaText, '->', taomaValue);
                                                    }
                                                    console.log('✓ 找到目标行:', fieldName, '实际文本:', cellText);
                                                    break;
                                                }
                                            }
                                        }
                                        if (targetRow) break;
                                    }
                                    if (targetRow) break;
                                }
                                if (targetRow) break;
                            }
                        }
                        
                        // 兼容普通 table：表头和表体在同一个 table 内
                        if (!targetRow) {
                            console.log('方法2: 查找普通 table...');
                            const allTables = searchScope.querySelectorAll('table');
                            for (let table of allTables) {
                                const rows = table.querySelectorAll('tr');
                                if (rows.length < 2) continue;
                                
                                const headerCells = rows[0].querySelectorAll('th, td');
                                const liangtiColumnIndex = findColumnIndex(headerCells, ['量体部位', '部位']);
                                const adjustColumnIndex = findColumnIndex(headerCells, ['调整尺寸', '调整']);
                                const taomaColumnIndex = findColumnIndex(headerCells, ['套码', '标准尺寸']);
                                if (liangtiColumnIndex < 0 || adjustColumnIndex < 0) continue;
                                
                                for (let i = 1; i < rows.length; i++) {
                                    const cells = rows[i].querySelectorAll('td, th');
                                    if (cells.length <= Math.max(liangtiColumnIndex, adjustColumnIndex)) continue;
                                    
                                    const cellText = getText(cells[liangtiColumnIndex]);
                                    const normalizedCellText = normalizeText(cellText);
                                    if (normalizedCellText) visiblePartNames.push(normalizedCellText);
                                    if (textMatches(cellText)) {
                                        targetRow = rows[i];
                                        targetCellIndex = adjustColumnIndex;
                                        if (taomaColumnIndex >= 0 && cells.length > taomaColumnIndex) {
                                            const taomaText = getText(cells[taomaColumnIndex]);
                                            taomaValue = parseFloat(taomaText) || 0;
                                            console.log('✓ 获取到套码值:', taomaText, '->', taomaValue);
                                        }
                                        targetCell = cells[adjustColumnIndex];
                                        console.log('✓ 在普通 table 中找到目标行');
                                        break;
                                    }
                                }
                                if (targetRow) break;
                            }
                        }
                        
                        // 最后再按单元格文本找行，但仍尽量使用该行最后一个 input-number，避免误填标准尺寸列
                        if (!targetRow) {
                            console.log('方法3: 直接查找包含量体部位名称的单元格...');
                            const cells = searchScope.querySelectorAll('td, th');
                            for (let cell of cells) {
                                const cellText = getText(cell);
                                const normalizedCellText = normalizeText(cellText);
                                if (normalizedCellText) visiblePartNames.push(normalizedCellText);
                                if (textMatches(cellText)) {
                                    targetRow = cell.parentElement;
                                    console.log('✓ 通过单元格找到目标行');
                                    break;
                                }
                            }
                        }

                        // 页面如“录入量体数据V2”左右分表时，最稳的是直接按行文本命中量体部位
                        if (!targetRow) {
                            console.log('方法4: 按整行文本查找量体部位...');
                            const rows = searchScope.querySelectorAll('tbody tr, tr');
                            for (let row of rows) {
                                const rowText = getText(row);
                                if (!rowText) continue;
                                const normalizedRowText = normalizeText(rowText);
                                if (normalizedRowText) visiblePartNames.push(normalizedRowText);
                                if (textMatches(rowText)) {
                                    targetRow = row;
                                    const cells = row.querySelectorAll('td, th');
                                    targetCell = Array.from(cells).find(cell => cell.querySelector('.el-input-number__increase, .el-input-number__decrease, input')) || null;
                                    console.log('✓ 通过整行文本找到目标行:', normalizedRowText);
                                    break;
                                }
                            }
                        }
                        
                        if (!targetRow) {
                            console.log('❌ 未找到目标行');
                            return {
                                success: false,
                                error: '未找到目标行',
                                candidates: normalizedCandidates,
                                visiblePartNames: Array.from(new Set(visiblePartNames)).slice(0, 80)
                            };
                        }
                        
                        // 3. 计算调整值
                        // 如果传入的值是负数，直接作为调整值
                        // 否则作为目标成衣尺寸，需要计算：调整尺寸 = 目标成衣尺寸 - 套码尺寸
                        let adjustValue;
                        if (targetChengYiSize < 0) {
                            adjustValue = targetChengYiSize;
                            console.log('✓ 直接使用调整值:', adjustValue);
                        } else {
                            adjustValue = targetChengYiSize - taomaValue;
                            console.log('✓ 计算调整值: 目标成衣尺寸(' + targetChengYiSize + ') - 套码(' + taomaValue + ') = ' + adjustValue);
                        }
                        
                        const isPlus = adjustValue >= 0;
                        const num = Math.abs(Math.round(adjustValue));
                        
                        // 在“调整尺寸”列找按钮
                        let btnSelector = isPlus ? '.el-input-number__increase' : '.el-input-number__decrease';
                        let btn = null;
                        let inputBox = null;
                        const expectedValue = String(Math.round(adjustValue));
                        const setInputValue = (input, value) => {
                            input.scrollIntoView({ block: 'center' });
                            input.focus();
                            input.click();
                            const setter = Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, 'value').set;
                            setter.call(input, value);
                            input.dispatchEvent(new InputEvent('input', { bubbles: true, inputType: 'insertText', data: value }));
                            input.dispatchEvent(new Event('change', { bubbles: true }));
                            input.dispatchEvent(new Event('blur', { bubbles: true }));
                        };
                        const getTargetInput = () => {
                            if (targetCell) {
                                const input = targetCell.querySelector('input');
                                if (input) return input;
                            }
                            const inputs = Array.from(targetRow.querySelectorAll('input'));
                            return inputs.length ? inputs[inputs.length - 1] : null;
                        };
                        
                        if (targetCell) {
                            btn = targetCell.querySelector(btnSelector);
                            if (btn) console.log('✓ 方式1: 在调整尺寸列找到按钮');
                        }
                        
                        if (!btn) {
                            const cells = targetRow.querySelectorAll('td, th');
                            if (targetCellIndex >= 0 && cells.length > targetCellIndex) {
                                btn = cells[targetCellIndex].querySelector(btnSelector);
                                if (btn) console.log('✓ 方式2: 通过列索引找到按钮');
                            }
                        }
                        
                        if (!btn) {
                            const inputNumbers = targetRow.querySelectorAll('.el-input-number');
                            if (inputNumbers.length > 0) {
                                const inputNumber = inputNumbers[inputNumbers.length - 1];
                                btn = inputNumber.querySelector(btnSelector);
                                if (btn) console.log('✓ 方式3: 回退到该行最后一个input-number按钮');
                            }
                        }
                        
                        if (!btn) {
                            btn = targetRow.querySelector(btnSelector);
                            if (btn) console.log('✓ 方式4: 在目标行内找到按钮');
                        }
                        
                        if (btn) {
                            console.log('✓ 找到', isPlus ? '+ 号按钮' : '- 号按钮', '需要点击', num, '次');
                            
                            // 点击 N 次
                            for (let i = 0; i < num; i++) {
                                btn.click();
                                console.log('第', i+1, '次点击');
                            }
                            
                            inputBox = getTargetInput();
                            const currentValue = inputBox ? String(inputBox.value || '').trim() : '';
                            console.log('点击后的输入框值:', currentValue, '期望:', expectedValue);
                            if (!inputBox) {
                                return {
                                    success: false,
                                    error: '点击按钮后未找到输入框，无法验证是否生效',
                                    candidates: normalizedCandidates,
                                    visiblePartNames: Array.from(new Set(visiblePartNames)).slice(0, 80)
                                };
                            }
                            if (currentValue !== expectedValue) {
                                console.log('点击未生效，改用原生 setter 直接写入...');
                                setInputValue(inputBox, expectedValue);
                            }
                            const finalValue = String(inputBox.value || '').trim();
                            console.log('最终输入框值:', finalValue);
                            return {
                                success: finalValue === expectedValue,
                                value: finalValue,
                                taomaValue: taomaValue,
                                adjustValue: adjustValue,
                                error: finalValue === expectedValue ? '' : '已找到输入框但前端值未生效'
                            };
                        } else {
                            console.log('❌ 未找到按钮');
                            
                            // 尝试直接设置输入框值作为备选方案
                            inputBox = getTargetInput();
                            if (inputBox) {
                                console.log('尝试直接设置输入框值...');
                                setInputValue(inputBox, expectedValue);
                                const finalValue = String(inputBox.value || '').trim();
                                console.log('✓ 已直接设置输入框值:', finalValue);
                                return {
                                    success: finalValue === expectedValue,
                                    value: finalValue,
                                    taomaValue: taomaValue,
                                    adjustValue: adjustValue,
                                    error: finalValue === expectedValue ? '' : '已找到输入框但前端值未生效'
                                };
                            }
                            
                            return {
                                success: false,
                                error: '找到目标行但未找到调整尺寸输入控件',
                                candidates: normalizedCandidates,
                                visiblePartNames: Array.from(new Set(visiblePartNames)).slice(0, 80)
                            };
                        }
                    """, field_name, target_cheng_yi_size, field_candidates)
                    field_found = bool(fill_result and fill_result.get('success')) if isinstance(fill_result, dict) else bool(fill_result)
                    
                    if field_found:
                        taoma_val = fill_result.get('taomaValue', 0) if isinstance(fill_result, dict) else 0
                        adjust_val = fill_result.get('adjustValue', 0) if isinstance(fill_result, dict) else 0
                        print(f'      ✓ 已填写 {field_name}: 目标成衣尺寸={field_value}, 套码={taoma_val}, 调整尺寸={adjust_val}')
                    else:
                        print(f'      ! 未找到量体字段：{field_name}')
                        if isinstance(fill_result, dict):
                            print(f'        匹配候选：{fill_result.get("candidates", [])}')
                            visible_names = fill_result.get('visiblePartNames', [])
                            if visible_names:
                                print(f'        当前可见量体部位：{visible_names[:40]}')
                            if fill_result.get('error'):
                                print(f'        原因：{fill_result.get("error")}')
                        
                except Exception as e:
                    print(f'      ! 填写量体数据时出错：{e}')
        
        sleep(1000)
        
        print('  【第1步】保存量体信息窗口...')
        # 获取保存前对话框数量
        dialog_count_before = len(self.driver.find_elements(By.CSS_SELECTOR, '.el-dialog, .el-drawer'))
        print(f'  保存前对话框数量：{dialog_count_before}')
        
        # 在量体对话框中点击保存
        # 排除明细字段，避免匹配到明细对话框
        save_success = self.click_save_in_correct_dialog(
            required_fields=['尺码', '落差'],
            excluded_fields=['款式类型', '流水号', '版型号']
        )
        if not save_success:
            print('  ⚠ 第1次保存量体信息失败')
        else:
            print('  ✓ 已点击量体信息保存（第1次）')
        
        sleep(5000)
        
        print('  【第2步】等待尺码调整规则窗口...')
        # 等待尺码调整规则窗口弹出
        for i in range(10):
            sleep(1000)
            try:
                dialog_count = len(self.driver.find_elements(By.CSS_SELECTOR, '.el-dialog, .el-drawer'))
                print(f'  对话框数量：{dialog_count}')
                if dialog_count > dialog_count_before:
                    print('  ✓ 尺码调整规则窗口已弹出')
                    break
            except:
                pass
        
        print('  保存尺码调整规则窗口...')
        # 保存尺码调整规则窗口 - 用标题精确匹配
        save_success = self.click_save_in_correct_dialog(
            dialog_title='尺码调整规则设置',
            button_text=['确定', '保存']
        )
        if save_success:
            print('  ✓ 已点击尺码调整规则保存')
        else:
            print('  ⚠ 尺码调整规则保存失败')
        
        sleep(5000)
        
        print('  等待尺码调整规则窗口关闭...')
        # 等待尺码调整规则窗口关闭
        window_closed = False
        for i in range(5):
            sleep(1000)
            try:
                dialog_count = len(self.driver.find_elements(By.CSS_SELECTOR, '.el-dialog, .el-drawer'))
                print(f'  对话框数量：{dialog_count}')
                if dialog_count <= 3:
                    print('  ✓ 尺码调整规则窗口已关闭')
                    window_closed = True
                    break
            except:
                window_closed = True
                break
        
        if not window_closed:
            print('  ⚠ 窗口未关闭，继续执行...')
        
        print('  【第3步】再次保存量体信息窗口...')
        # 再次保存量体信息窗口
        dialog_count = len(self.driver.find_elements(By.CSS_SELECTOR, '.el-dialog, .el-drawer'))
        print(f'  当前对话框数量：{dialog_count}')
        
        save_success = self.click_save_in_correct_dialog(
            required_fields=['尺码', '落差'],
            excluded_fields=['款式类型', '流水号', '版型号']
        )
        if save_success:
            print('  ✓ 已点击量体信息保存（第2次）')
        else:
            print('  ⚠ 第2次保存量体信息失败')
        
        sleep(5000)
        
        print('  等待量体窗口关闭...')
        # 等待量体窗口关闭
        for i in range(10):
            sleep(1000)
            try:
                dialog_count = len(self.driver.find_elements(By.CSS_SELECTOR, '.el-dialog, .el-drawer'))
                print(f'  对话框数量：{dialog_count}')
                if dialog_count <= 2:
                    print('  ✓ 量体窗口已关闭')
                    break
            except:
                pass
        
        print('✓ 量体信息填写完成\n')
        return True
    
    def save_detail(self):
        print('【8】保存明细...')
        
        # 滚动到对话框最下方，确保按钮可见
        for i in range(3):
            self.driver.execute_script("""
                document.querySelectorAll('.el-dialog__body, .el-drawer__body').forEach(dialogBody => {
                    dialogBody.scrollTop = dialogBody.scrollHeight;
                });
                document.querySelectorAll('.el-dialog__footer, .el-drawer__footer').forEach(footer => {
                    footer.scrollIntoView({block: "center"});
                });
                window.scrollTo(0, document.body.scrollHeight);
            """)
            sleep(600)
            print(f'  滚动第{i+1}次')
        
        sleep(1000)
        
        # 截图保存当前状态
        screenshot_path = f'e:\\111\\Material\\debug_save_detail_before.png'
        self.driver.save_screenshot(screenshot_path)
        print(f'  📸 已保存截图：{screenshot_path}')
        
        # 使用通用函数在明细对话框中点击保存
        save_clicked_success = self.click_save_in_correct_dialog(required_fields=['版型号', '面料编号'])
        
        if not save_clicked_success:
            print('  ✗ 无法点击保存按钮')
            print('⚠ 保存失败\n')
            sleep(2000)
            return False
        
        # 等待保存结果
        print('  等待保存结果...')
        result_status = None
        for j in range(20):
            sleep(1000)
            msgs = self.driver.execute_script("""
                const msgs = [];
                document.querySelectorAll('.el-message').forEach(el => {
                    const text = el.innerText?.trim();
                    if (text) msgs.push(text);
                });
                return msgs;
            """)
            if j < 10:
                print(f'    提示：{msgs}')
            
            if any('成功' in m for m in msgs):
                result_status = 'success'
                break
        
        if result_status == 'success':
            print('  ✓ 保存成功')
            print('✓ 明细保存成功\n')
            sleep(2000)
            return True
        else:
            print('  ✗ 保存超时')
            print('⚠ 保存失败\n')
            sleep(2000)
            return False
    
    def get_production_no(self):
        """从内存获取生产单号"""
        if self.production_no:
            print(f'  ✓ 从内存获取生产单号：{self.production_no}')
            return self.production_no
        else:
            print('  ⚠ 内存中没有生产单号')
            return ''
    
    def run_test(self):
        print('=== SCM 订单完整流程测试 ===\n')
        
        try:
            self.login()
            self.enter_custom_order()
            
            # 先加载配置，确保 self.config 可用
            banxing_list = self.load_banxing_from_config()
            
            customer_ok = self.fill_order_info()
            if not customer_ok:
                print('❌ 客商选择失败，终止流程')
                print('\n浏览器保持打开，请按回车键关闭...')
                input()
                return
            
            self.save_order_main()
            
            for idx, item in enumerate(banxing_list, 1):
                banxing = item['banxing']
                chima_size = item['chima_size']
                luocha = item['luocha']
                custom_options = item.get('custom_options', {})
                fabric_width = item.get('fabric_width')
                fabric_no = item.get('fabric_no')
                fabric_style = item.get('fabric_style')
                liangti_data = item.get('liangti_data', {})
                
                print(f'\n{"="*50}')
                print(f'【第 {idx}/{len(banxing_list)} 个版型】{banxing}')
                print(f'{"="*50}')
                print(f'  尺码：{chima_size}')
                print(f'  面料编号：{fabric_no}')
                print(f'  面料风格：{fabric_style}')
                print(f'  门幅：{fabric_width}')
                
                style_type = self.get_style_type(banxing)
                
                self.add_detail(banxing, chima_size, luocha, fabric_width, fabric_no, fabric_style)
                self.edit_custom_options(custom_options)
                self.enter_liangti_info(chima_size, style_type, luocha, liangti_data)
                save_success = self.save_detail()
                
                if not save_success:
                    print(f'❌ 第{idx}个版型 {banxing} 保存失败，终止流程')
                    print('\n浏览器保持打开，请按回车键关闭...')
                    input()
                    return
            
            print('\n✓ 所有版型处理完成！')
            
            print('\n【读取生产单号】')
            production_no = self.get_production_no()
            
            print('\n【保存主表订单】')
            self.save_order_main()
            
            print('\n' + '='*50)
            print(f'✓ 流程完成！生产单号：{production_no}')
            print('='*50)
            
            print('\n浏览器保持打开，请按回车键关闭...')
            input()
            
        except Exception as e:
            print(f'\n发生错误：{e}')
            import traceback
            traceback.print_exc()
            print('\n浏览器保持打开，请按回车键关闭...')
            input()
        finally:
            self.driver.quit()
            print('浏览器已关闭')
    
if __name__ == '__main__':
    print('=== SCM订单完整流程测试 ===\n')
    
    test = TestFullProcess()
    test.run_test()

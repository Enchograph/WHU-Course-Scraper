from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
import pandas as pd
import time
import os
from datetime import datetime

def setup_driver():
    """设置 Chrome 驱动"""
    chrome_options = Options()
    # 如果需要无头模式（不显示浏览器窗口），取消下面这行的注释
    # chrome_options.add_argument('--headless')
    chrome_options.add_argument('--disable-gpu')
    chrome_options.add_argument('--no-sandbox')
    chrome_options.add_argument('--disable-dev-shm-usage')
    
    # 设置驱动路径，根据实际情况修改
    driver_path = "chromedriver路径"  # 例如：'C:/webdrivers/chromedriver.exe'
    
    service = Service(executable_path=None)
    driver = webdriver.Chrome(service=service, options=chrome_options)
    
    return driver

def login_and_navigate(driver, url):
    """打开页面并等待用户登录"""
    driver.get(url)

    # 切换至统一身份认证登录方式
    try:
        # 使用显式等待，但设置较短的超时时间，避免长时间等待
        WebDriverWait(driver, 5).until(
            EC.element_to_be_clickable((By.ID, "tysfyzdl"))
        )
        login_button = driver.find_element(By.ID, "tysfyzdl")
        login_button.click()
        print("成功点击统一身份认证登录按钮")
    except Exception as e:
        print(f"未找到统一身份认证登录按钮或点击失败: {str(e)}")
        # 尝试使用其他选择器定位按钮
        try:
            # 使用CSS选择器尝试找到按钮
            login_button = driver.find_element(By.CSS_SELECTOR, ".btn.btn-primary.btn-block.rzdl")
            login_button.click()
            print("使用CSS选择器成功点击登录按钮")
        except:
            # 使用XPath尝试找到按钮
            try:
                login_button = driver.find_element(By.XPATH, "//button[contains(text(), '统一身份认证登录')]")
                login_button.click()
                print("使用XPath成功点击登录按钮")
            except:
                print("所有尝试都失败，继续执行后续操作")
    
    driver.maximize_window()

    return driver

def get_table_id(driver):
    """自动检测表格ID"""
    try:
        # 查找所有可能的表格元素
        tables = driver.find_elements(By.CSS_SELECTOR, "table.ui-jqgrid-btable")
        
        if tables:
            # 获取第一个表格的ID
            table_id = tables[0].get_attribute("id")
            print(f"检测到表格ID: {table_id}")
            return table_id
        
        # 尝试通过表格容器查找
        grid_boxes = driver.find_elements(By.CSS_SELECTOR, "div.ui-jqgrid")
        if grid_boxes:
            for grid_box in grid_boxes:
                # 尝试找到表格ID
                grid_id = grid_box.get_attribute("id")
                if grid_id and grid_id.startswith("gbox_"):
                    # 从gbox_XXX提取表格ID XXX
                    table_id = grid_id[5:]  # 移除前缀 "gbox_"
                    print(f"从表格容器检测到表格ID: {table_id}")
                    
                    # 验证表格ID存在
                    if driver.find_elements(By.ID, table_id):
                        return table_id
        
        # 查找所有带角色属性的表格
        role_tables = driver.find_elements(By.CSS_SELECTOR, "table[role='grid']")
        if role_tables:
            table_id = role_tables[0].get_attribute("id")
            print(f"从角色属性检测到表格ID: {table_id}")
            return table_id
            
        print("未找到表格元素")
        return None
    except Exception as e:
        print(f"检测表格ID时出错: {e}")
        return None

def check_pagination(driver, table_id):
    """检查是否有分页，返回总页数"""
    try:
        # 从表格ID构建分页器ID
        pager_id = f"pager_{table_id}"
        
        # 尝试使用JavaScript获取总页数
        try:
            js_code = """
            // 尝试方法1: 从记录信息计算
            var pagerRight = document.querySelector('[id$="_right"] .ui-paging-info');
            if (pagerRight) {
                var text = pagerRight.textContent;
                var match = text.match(/共\\s+(\\d+)\\s+条/);
                if (match) {
                    var totalRecords = parseInt(match[1]);
                    var perPage = parseInt(document.querySelector('.ui-pg-selbox').value);
                    return Math.ceil(totalRecords / perPage);
                }
            }
            
            // 尝试方法2: 直接获取页码信息
            var pageSpan = document.querySelector('span[id^="sp_1_"]');
            if (pageSpan) {
                return parseInt(pageSpan.textContent);
            }
            
            // 如果都失败，返回默认值
            return 1;
            """
            total_pages = driver.execute_script(js_code)
            if total_pages and total_pages > 0:
                print(f"JavaScript获取的总页数: {total_pages}")
                return total_pages
        except Exception as js_err:
            print(f"使用JavaScript获取总页数时出错: {js_err}")

        # 找不到分页信息，或者只有一页
        return 1
    
    except Exception as e:
        print(f"检查分页信息时出错: {e}")
        return 1  # 默认返回1页

def navigate_to_page(driver, table_id, page_number):
    """导航到指定页码"""
    try:
        # 使用JavaScript直接进行分页导航，避免依赖CSS选择器
        js_code = f"""
        // 查找所有可能的分页按钮
        var nextButtons = document.querySelectorAll('a[role="button"][id^="next_"]');
        var pageInputs = document.querySelectorAll('.ui-pg-input');
        
        // 优先使用页码输入框
        if (pageInputs.length > 0) {{
            // 找到第一个可用的页码输入框
            for (var i = 0; i < pageInputs.length; i++) {{
                var input = pageInputs[i];
                if (window.getComputedStyle(input).display !== 'none') {{
                    input.value = '{page_number}';
                    
                    // 创建回车事件触发页面跳转
                    var event = new KeyboardEvent('keypress', {{
                        key: 'Enter',
                        code: 'Enter',
                        keyCode: 13,
                        which: 13,
                        bubbles: true
                    }});
                    
                    input.dispatchEvent(event);
                    return true;
                }}
            }}
        }}
        
        // 如果找不到页码输入框，尝试使用下一页按钮
        if (nextButtons.length > 0) {{
            // 找到第一个可用的下一页按钮
            for (var i = 0; i < nextButtons.length; i++) {{
                var btn = nextButtons[i];
                if (window.getComputedStyle(btn).display !== 'none' && 
                    !btn.classList.contains('ui-state-disabled')) {{
                    
                    // 模拟点击多次下一页按钮
                    var currentPage = 1; // 假设当前在第1页
                    var clickCount = {page_number} - currentPage;
                    
                    for (var j = 0; j < clickCount; j++) {{
                        btn.click();
                        // 需要一个小延迟来确保页面加载
                    }}
                    
                    return true;
                }}
            }}
        }}
        
        // 尝试直接查找页码按钮
        var pageButtons = document.querySelectorAll('.ui-pg-button');
        for (var i = 0; i < pageButtons.length; i++) {{
            var btn = pageButtons[i];
            if (btn.textContent.trim() === '{page_number}') {{
                btn.click();
                return true;
            }}
        }}
        
        return false;
        """
        
        result = driver.execute_script(js_code)
        # 等待页面加载
        time.sleep(1)
        
        print(f"使用JavaScript导航到第 {page_number} 页")
        return True
    
    except Exception as e:
        print(f"导航到第 {page_number} 页时出错: {e}")
        return False

def extract_table_data_with_js(driver, table_id):
    """使用JavaScript直接从DOM中提取表格数据"""
    try:
        # 检查表格是否存在
        if not driver.find_elements(By.ID, table_id):
            print(f"未找到ID为{table_id}的表格")
            return [], []
        
        # 使用JavaScript提取表头
        headers_js = """
        var headers = [];
        var headerElements = document.querySelectorAll('.ui-jqgrid-htable th div.ui-jqgrid-sortable');
        for (var i = 0; i < headerElements.length; i++) {
            if (window.getComputedStyle(headerElements[i]).display !== 'none') {  // 检查元素是否可见
                var text = headerElements[i].textContent.trim();
                if (text) headers.push(text);
            }
        }
        return headers;
        """
        
        headers = driver.execute_script(headers_js)
        # print(f"JavaScript提取的表头: {headers}")
        
        # 使用JavaScript提取数据行
        rows_js = f"""
        var rows = [];
        var rowElements = document.querySelectorAll('#{table_id} tr.jqgrow');
        
        for (var i = 0; i < rowElements.length; i++) {{
            var row = [];
            var cells = rowElements[i].querySelectorAll('td');
            
            for (var j = 0; j < cells.length; j++) {{
                if (window.getComputedStyle(cells[j]).display !== 'none') {{  // 检查元素是否可见
                    var cellText = cells[j].getAttribute('title') || cells[j].textContent.trim();
                    row.push(cellText || '');
                }}
            }}
            
            // 检查行是否包含非空数据
            var hasData = false;
            for (var k = 0; k < row.length; k++) {{
                if (row[k].trim()) {{
                    hasData = true;
                    break;
                }}
            }}
            
            if (hasData) rows.push(row);
        }}
        
        return rows;
        """
        
        rows = driver.execute_script(rows_js)
        print(f"JavaScript提取的数据行数: {len(rows)}")
        
        # 过滤和标准化数据
        if headers and rows and len(rows) > 0:
            # 跳过第一列（复选框列）
            filtered_rows = [row[1:] for row in rows]
            filtered_headers = headers[1:] if len(headers) > 1 else headers
            
            # 确保数据一致性
            max_row_len = max(len(row) for row in filtered_rows) if filtered_rows else 0
            
            if max_row_len > 0:
                # 调整表头长度
                if len(filtered_headers) < max_row_len:
                    filtered_headers.extend([f"列{i+1}" for i in range(len(filtered_headers), max_row_len)])
                filtered_headers = filtered_headers[:max_row_len]
                
                # 标准化行数据
                standardized_rows = []
                for row in filtered_rows:
                    if len(row) < max_row_len:
                        row.extend(["" for _ in range(max_row_len - len(row))])
                    standardized_rows.append(row[:max_row_len])
                
                return filtered_headers, standardized_rows
        
        return headers, rows
        
    except Exception as e:
        print(f"使用JavaScript提取表格数据时出错: {e}")
        import traceback
        traceback.print_exc()
        return [], []

def scrape_all_pages(driver):
    """爬取所有页的数据"""
    # 获取表格ID
    table_id = get_table_id(driver)
    
    print(f"使用表格ID: {table_id}")
    print()

    # 获取总页数
    total_pages = check_pagination(driver, table_id)
    print()

    
    all_headers = []
    all_rows = []
    
    # 使用JavaScript提取
    headers, rows = extract_table_data_with_js(driver, table_id)

    if headers:
        all_headers = headers
    if rows:
        all_rows.extend(rows)
    
    # 如果第一页没有获取到数据，记录错误并返回空结果
    if not all_headers or not all_rows:
        print(f"无法从第1页获取数据，请检查网页结构是否发生变化（提示：您登录时只需点击一次登录按钮）")
        # with open(f"debug_page_source_{datetime.now().strftime('%Y%m%d_%H%M%S')}.html", "w", encoding="utf-8") as f:
        #     f.write(driver.page_source)
        # print("已保存页面源代码以便调试")
        return [], []
    
    # 爬取剩余页面
    if total_pages > 1:
        for page in range(2, total_pages + 1):
            success = navigate_to_page(driver, table_id, page)
            if success:
                # 等待足够长的时间确保页面加载
                time.sleep(3)
                
                # 使用JavaScript获取后续页面数据
                page_headers, page_rows = extract_table_data_with_js(driver, table_id)

                if page_rows:
                    all_rows.extend(page_rows)
                    print(f"成功获取第 {page} 页数据: {len(page_rows)} 行")
                else:
                    print(f"无法从第 {page} 页获取数据")
            else:
                print(f"无法导航到第 {page} 页，跳过")
            print()
    
    return all_headers, all_rows

def save_to_excel(headers, rows, filename=None):
    """将数据保存为Excel文件"""
    current_dir = os.path.dirname(os.path.abspath(__file__))
    
    if not filename:
        # 如果没有指定文件名，则使用当前时间创建文件名
        current_time = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"课表查询结果_{current_time}.xlsx"
    
    file_path = os.path.join(current_dir, filename)
    
    # 创建DataFrame
    df = pd.DataFrame(rows, columns=headers)
    
    # 保存为Excel
    df.to_excel(file_path, index=False)
    
    print(f"数据已保存为: {file_path}")
    return file_path


def main():
    """主函数"""
    # 课表查询系统URL
    url = "https://jwgl.whu.edu.cn/design/viewFunc_cxDesignFuncPageIndex.html?gnmkdm=N214599&layout=default"
    
    try:
        driver = setup_driver()
        
        driver = login_and_navigate(driver, url)
        
        print("请在浏览器中登录并选择筛选条件。数据加载完成后按 Enter 继续...")
        input()
        
        # 等待一段时间确保结果加载完毕
        time.sleep(3)
        
        # 爬取所有页的数据
        headers, rows = scrape_all_pages(driver)
        
        if headers and rows:
            # 保存为Excel
            save_to_excel(headers, rows)
            print(f"成功爬取数据: {len(headers)} 列, {len(rows)} 行")

    except Exception as e:
        print(f"执行过程中出错: {e}")
        import traceback
        traceback.print_exc()
    
    finally:
        if 'driver' in locals():
            driver.quit()
        
        print("程序执行完毕")

if __name__ == "__main__":
    main()


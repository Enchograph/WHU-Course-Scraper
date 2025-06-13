# 武汉大学教务系统课程表爬取

代码基于 Selenium 自动化框架，从武汉大学教务系统页面爬取课表数据，导之出为 Excel 文件。

## 安装与使用

1. 克隆项目仓库：

   ```bash
   git clone https://github.com/Enchograph/WHU-Course-Scraper.git
   ```

2. 安装依赖：

   ```bash
   pip install -r requirements.txt
   ```
  
3. 使用

   ```bash
   python scrape_course_table.py
   ```

4. 若程序无法运行：
  请下载与您的 Chrome 浏览器版本匹配的 ChromeDriver：
   - 地址：https://sites.google.com/chromium.org/driver/
   - 将 `chromedriver` 放入系统环境变量 PATH 中，或在代码中手动指定路径。


## 注意事项

- 代码写就于2025年6月13日。请您注意网站结构的时效性。

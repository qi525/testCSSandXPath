import pandas as pd
from playwright.sync_api import sync_playwright
import datetime
import os
import subprocess

def extract_img_data_with_xpath_levels(url, resolution_width, resolution_height):
    """
    使用 Playwright 监控模式打开指定网站，获取所有 img 标签的 class, alt, id, 相对 XPath, 完整 XPath 和 XPath 层级。
    将原始数据保存到一个工作表，并在同一个工作簿的另一个工作表进行 class 属性的重复次数分析。
    如果出现错误，则生成日志文件。
    """
    log_file_path = "img_data_xpath_levels_log.txt"
    output_xlsx_path = "img_data_xpath_levels_analysis.xlsx"

    try:
        with sync_playwright() as p:
            browser = p.chromium.launch(headless=False)  # 设置 headless=False 启用监控模式
            page = browser.new_page()
            page.set_viewport_size({"width": resolution_width, "height": resolution_height})

            print(f"正在打开网站: {url}")
            page.goto(url)
            print("等待10秒...")
            page.wait_for_timeout(10000)  # 等待10秒

            print("正在获取所有 img 标签的 class, alt, id, 相对 XPath, 完整 XPath 和 XPath 层级...")
            # 将 'button' 更改为 'img'
            img_locators = page.locator("img").all()
            img_data_raw = []

            # --- 合并后的 JavaScript 函数 ---
            xpath_js_functions = """
            (element) => {
                // Helper function to get index of sibling
                function getElementIndex(node) {
                    let i = 1;
                    let sibling = node.previousElementSibling;
                    while(sibling) {
                        if (sibling.nodeName === node.nodeName) {
                            i++;
                        }
                        sibling = sibling.previousElementSibling;
                    }
                    return i;
                }

                // Function to get Relative XPath (id优先)
                function getRelativeXPath(el) {
                    if (!el) return '';
                    if (el.id) {
                        return `//*[@id="${el.id}"]`;
                    }
                    const parts = [];
                    let current = el;
                    while (current && current.nodeType === Node.ELEMENT_NODE && current !== document.body) {
                        let tagName = current.tagName.toLowerCase();
                        let index = getElementIndex(current);
                        let selector = tagName;
                        if (index > 1) {
                            selector += `[${index}]`;
                        }
                        parts.unshift(selector);
                        current = current.parentNode;
                    }
                    if (current === document.body) {
                        parts.unshift('/html/body');
                    } else if (current && current.tagName === 'HTML') {
                        parts.unshift('/html');
                    }
                    return parts.join('/');
                }

                // Function to get Full XPath (从 /html/body 开始)
                function getFullXPath(el) {
                    if (!el) return '';
                    const parts = [];
                    let current = el;
                    while (current && current.nodeType === Node.ELEMENT_NODE) {
                        let tagName = current.tagName.toLowerCase();
                        let index = getElementIndex(current);
                        let selector = tagName;
                        if (current.parentNode && Array.from(current.parentNode.children).filter(s => s.nodeName === current.nodeName).length > 1) {
                            selector += `[${index}]`;
                        }
                        parts.unshift(selector);
                        if (current.tagName === 'HTML') break;
                        current = current.parentNode;
                    }
                    return '/' + parts.join('/');
                }

                return {
                    relative: getRelativeXPath(element),
                    full: getFullXPath(element)
                };
            }
            """

            for i, img_locator in enumerate(img_locators): # 循环变量也改为 img_locator
                class_attr = img_locator.get_attribute("class")
                alt_attr = img_locator.get_attribute("alt")
                id_attr = img_locator.get_attribute("id")

                element_handle = img_locator.element_handle()

                xpath_results = page.evaluate(xpath_js_functions, element_handle)
                
                relative_xpath = xpath_results['relative']
                full_xpath = xpath_results['full']

                # 计算 Full_XPath 的层级（斜杠数量）
                xpath_level = full_xpath.count('/') if full_xpath else 0

                img_data_raw.append({
                    "Image_Index": i + 1, # 列名更改为 Image_Index
                    "Class": class_attr if class_attr else "",
                    "Alt": alt_attr if alt_attr else "",
                    "ID": id_attr if id_attr else "",
                    "Relative_XPath": relative_xpath if relative_xpath else "",
                    "Full_XPath": full_xpath if full_xpath else "",
                    "XPath_Level": xpath_level
                })

            if img_data_raw: # 检查 img_data_raw
                df_raw = pd.DataFrame(img_data_raw)

                # --- Class 属性重复次数分析 ---
                # 注意：img 标签的 class 属性通常也可能包含多个类名，此分析逻辑依然适用
                df_class_analysis = df_raw.copy()
                df_class_analysis['Class_Single'] = df_class_analysis['Class'].str.split(' ')
                df_class_analysis = df_class_analysis.explode('Class_Single')
                df_class_analysis = df_class_analysis[df_class_analysis['Class_Single'] != ''].copy()

                class_counts = df_class_analysis['Class_Single'].value_counts().reset_index()
                class_counts.columns = ['Class_Name', 'Count']
                df_class_counts = class_counts.sort_values(by='Count', ascending=False)

                # --- 将数据写入同一个 XLSX 文件的不同工作表 ---
                with pd.ExcelWriter(output_xlsx_path, engine='xlsxwriter') as writer:
                    df_raw.to_excel(writer, sheet_name='Raw_Image_Data', index=False) # 工作表名更改
                    df_class_counts.to_excel(writer, sheet_name='Class_Analysis', index=False)

                print(f"图片属性、XPath (相对/完整) 和 XPath 层级、Class 分析结果已成功保存到: {output_xlsx_path}")
                subprocess.run(["start", output_xlsx_path], shell=True)
            else:
                print("未找到任何 img 标签或其属性。")

            browser.close()

    except Exception as e:
        error_message = f"发生错误: {e}\n时间: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
        print(error_message)
        with open(log_file_path, "w", encoding="utf-8") as f:
            f.write(error_message)
        print(f"错误日志已保存到: {log_file_path}")
        subprocess.run(["start", log_file_path], shell=True)

if __name__ == "__main__":
    # 打印样例 XPath 的斜杠数量
    sample_xpath = "/html/body/div[1]/div/div/div/div/div/main/div[2]/div/div/div/div[2]/div[1]/div[3]/div[1]/div[1]/a/img"
    slash_count = sample_xpath.count('/')
    print(f"样例 XPath '{sample_xpath}' 中的斜杠数量为: {slash_count}\n")

    target_url = "https://civitai.com/images?tags=4"
    res_width = 2560
    res_height = 1440
    extract_img_data_with_xpath_levels(target_url, res_width, res_height)
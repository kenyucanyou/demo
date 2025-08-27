# 这个是python版本的，实现的跟相同目录下的html文件的功能是一样的。

from flask import Flask, render_template_string, request, jsonify
import re

app = Flask(__name__)

# 读取HTML模板内容
HTML_TEMPLATE = '''
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Markdown与Excel双向转换工具</title>
    <script src="https://cdn.tailwindcss.com"></script>
    <link href="https://cdn.jsdelivr.net/npm/font-awesome@4.7.0/css/font-awesome.min.css" rel="stylesheet">
    <script>
        tailwind.config = {
            theme: {
                extend: {
                    colors: {
                        primary: '#3B82F6',
                        secondary: '#10B981',
                        neutral: '#64748B',
                    },
                    fontFamily: {
                        sans: ['Inter', 'system-ui', 'sans-serif'],
                    },
                }
            }
        }
    </script>
    <style type="text/tailwindcss">
        @layer utilities {
            .content-auto {
                content-visibility: auto;
            }
            .transition-height {
                transition: max-height 0.3s ease-in-out;
            }
        }
    </style>
</head>
<body class="bg-gray-50 min-h-screen">
    <div class="container mx-auto px-4 py-8 max-w-6xl">
        <!-- 标题区域 -->
        <header class="text-center mb-8">
            <h1 class="text-[clamp(1.8rem,4vw,2.5rem)] font-bold text-gray-800 mb-2">
                <i class="fa fa-exchange text-primary mr-2"></i>Markdown与Excel双向转换工具
            </h1>
            <p class="text-gray-600 max-w-2xl mx-auto">
                支持Markdown表格与Excel格式的相互转换，轻松处理表格数据
            </p>
        </header>

        <!-- 转换方向选择 -->
        <div class="flex justify-center mb-6">
            <div class="inline-flex p-1 bg-gray-100 rounded-lg">
                <button id="mdToExcelBtn" class="conversion-direction-btn px-6 py-2 rounded-md bg-primary text-white font-medium">
                    Markdown → Excel
                </button>
                <button id="excelToMdBtn" class="conversion-direction-btn px-6 py-2 rounded-md text-gray-600 font-medium">
                    Excel → Markdown
                </button>
            </div>
        </div>

        <!-- 主要内容区域 -->
        <main class="grid md:grid-cols-2 gap-6">
            <!-- 输入区域 -->
            <div class="bg-white rounded-lg shadow-md p-5 transform transition-all duration-300 hover:shadow-lg">
                <div class="flex items-center mb-3">
                    <i id="inputIcon" class="fa fa-file-text-o text-primary text-xl mr-2"></i>
                    <h2 id="inputTitle" class="text-lg font-semibold text-gray-700">Markdown表格输入</h2>
                </div>
                <p id="inputDescription" class="text-sm text-gray-500 mb-3">请粘贴您的Markdown表格内容：</p>
                <textarea 
                    id="sourceInput" 
                    class="w-full h-64 p-3 border border-gray-300 rounded-md focus:ring-2 focus:ring-primary/50 focus:border-primary transition-all resize-none text-sm"
                    placeholder="例如：
| 姓名 | 年龄 | 职业 |
|------|------|------|
| 张三 | 25   | 工程师 |
| 李四 | 30   | 设计师 |"
                >| 任务ID | 任务名称 | 负责人 | 状态 | 截止日期 |
| --- | --- | --- | --- | --- |
| 001    | 需求分析 | 张三   | 已完成 | 2023-06-10 |
| 002    | 系统设计 | 李四   | 进行中 | 2023-06-20 |
| 003    | 编码开发 | 王五   | 未开始 | 2023-07-15 |</textarea>
                <div class="mt-4 flex justify-end">
                    <button 
                        id="convertBtn" 
                        class="bg-primary hover:bg-primary/90 text-white px-4 py-2 rounded-md transition-all flex items-center"
                    >
                        <i class="fa fa-exchange mr-2"></i><span id="convertBtnText">转换为Excel格式</span>
                    </button>
                </div>
            </div>

            <!-- 输出区域 -->
            <div class="bg-white rounded-lg shadow-md p-5 transform transition-all duration-300 hover:shadow-lg">
                <div class="flex items-center mb-3">
                    <i id="outputIcon" class="fa fa-file-excel-o text-secondary text-xl mr-2"></i>
                    <h2 id="outputTitle" class="text-lg font-semibold text-gray-700">Excel格式输出</h2>
                </div>
                <p id="outputDescription" class="text-sm text-gray-500 mb-3">转换后的内容（可直接复制到Excel）：</p>
                
                <div id="resultContainer" class="relative">
                    <div id="emptyState" class="h-64 flex flex-col items-center justify-center text-gray-400 border border-dashed border-gray-300 rounded-md">
                        <i class="fa fa-clipboard text-4xl mb-3"></i>
                        <p>转换结果将显示在这里</p>
                        <p class="text-xs mt-1">转换后可直接使用</p>
                    </div>
                    
                    <div id="resultContent" class="hidden">
                        <div id="convertedOutput" class="w-full h-64 p-3 border border-gray-300 rounded-md overflow-auto text-sm whitespace-pre"></div>
                        <div class="mt-3 flex justify-between items-center">
                            <div id="conversionInfo" class="text-xs text-gray-500"></div>
                            <button 
                                id="copyBtn" 
                                class="bg-secondary hover:bg-secondary/90 text-white px-3 py-1 rounded-md transition-all text-sm flex items-center"
                            >
                                <i class="fa fa-copy mr-1"></i>复制到剪贴板
                            </button>
                        </div>
                    </div>
                </div>
            </div>
        </main>

        <!-- 表格预览区域 -->
        <div class="mt-8 bg-white rounded-lg shadow-md p-5">
            <div class="flex items-center justify-between mb-4">
                <div class="flex items-center">
                    <i class="fa fa-eye text-neutral text-xl mr-2"></i>
                    <h2 class="text-lg font-semibold text-gray-700">表格预览</h2>
                </div>
                <button id="togglePreviewBtn" class="text-sm text-primary hover:text-primary/80 transition-colors flex items-center">
                    <i class="fa fa-chevron-up mr-1" id="previewIcon"></i>
                    <span>隐藏预览</span>
                </button>
            </div>
            
            <div id="previewContainer" class="overflow-x-auto max-h-96 transition-height">
                <div id="previewEmptyState" class="py-10 text-center text-gray-400">
                    <p>转换后表格将在这里预览</p>
                </div>
                <table id="tablePreview" class="hidden w-full border-collapse">
                    <thead>
                        <tr id="tableHeader"></tr>
                    </thead>
                    <tbody id="tableBody"></tbody>
                </table>
            </div>
        </div>

        <!-- 使用说明 -->
        <div class="mt-8 bg-white rounded-lg shadow-md p-5">
            <h2 class="text-lg font-semibold text-gray-700 mb-3 flex items-center">
                <i class="fa fa-info-circle text-primary mr-2"></i>使用说明
            </h2>
            <div class="grid md:grid-cols-2 gap-6 text-sm">
                <div>
                    <h3 class="font-medium text-gray-800 mb-2">Markdown → Excel</h3>
                    <ul class="list-disc pl-5 text-gray-600 space-y-1">
                        <li>粘贴Markdown格式的表格内容（包含表头和分隔线）</li>
                        <li>点击转换按钮，得到制表符分隔的Excel格式内容</li>
                        <li>直接复制结果到Excel中即可保持表格结构</li>
                    </ul>
                </div>
                <div>
                    <h3 class="font-medium text-gray-800 mb-2">Excel → Markdown</h3>
                    <ul class="list-disc pl-5 text-gray-600 space-y-1">
                        <li>从Excel中复制表格内容（包含表头）</li>
                        <li>粘贴到输入框中，点击转换按钮</li>
                        <li>得到标准的Markdown表格格式，可直接用于文档</li>
                    </ul>
                </div>
            </div>
        </div>

        <!-- 页脚 -->
        <footer class="mt-10 text-center text-gray-500 text-sm">
            <p>Markdown与Excel双向转换工具 (Python版) &copy; 2023</p>
        </footer>
    </div>

    <!-- 通知提示 -->
    <div id="notification" class="fixed bottom-5 right-5 px-4 py-3 rounded-md shadow-lg transform translate-y-20 opacity-0 transition-all duration-300 flex items-center">
        <i id="notificationIcon" class="mr-2"></i>
        <span id="notificationText"></span>
    </div>

    <script>
        // 全局状态 - 当前转换方向 (mdToExcel 或 excelToMd)
        let conversionDirection = 'mdToExcel';
        
        // DOM元素
        const mdToExcelBtn = document.getElementById('mdToExcelBtn');
        const excelToMdBtn = document.getElementById('excelToMdBtn');
        const sourceInput = document.getElementById('sourceInput');
        const convertBtn = document.getElementById('convertBtn');
        const convertBtnText = document.getElementById('convertBtnText');
        const convertedOutput = document.getElementById('convertedOutput');
        const copyBtn = document.getElementById('copyBtn');
        const emptyState = document.getElementById('emptyState');
        const resultContent = document.getElementById('resultContent');
        const conversionInfo = document.getElementById('conversionInfo');
        const notification = document.getElementById('notification');
        const notificationIcon = document.getElementById('notificationIcon');
        const notificationText = document.getElementById('notificationText');
        const togglePreviewBtn = document.getElementById('togglePreviewBtn');
        const previewIcon = document.getElementById('previewIcon');
        const previewContainer = document.getElementById('previewContainer');
        const previewEmptyState = document.getElementById('previewEmptyState');
        const tablePreview = document.getElementById('tablePreview');
        const tableHeader = document.getElementById('tableHeader');
        const tableBody = document.getElementById('tableBody');
        const inputIcon = document.getElementById('inputIcon');
        const inputTitle = document.getElementById('inputTitle');
        const inputDescription = document.getElementById('inputDescription');
        const outputIcon = document.getElementById('outputIcon');
        const outputTitle = document.getElementById('outputTitle');
        const outputDescription = document.getElementById('outputDescription');

        // 显示通知
        function showNotification(message, isSuccess = true) {
            notificationText.textContent = message;
            notificationIcon.className = isSuccess ? 'fa fa-check-circle text-green-500 mr-2' : 'fa fa-exclamation-circle text-red-500 mr-2';
            notification.className = `fixed bottom-5 right-5 px-4 py-3 rounded-md shadow-lg transform translate-y-0 opacity-100 transition-all duration-300 flex items-center ${isSuccess ? 'bg-green-50 text-green-700' : 'bg-red-50 text-red-700'}`;
            
            setTimeout(() => {
                notification.className = 'fixed bottom-5 right-5 px-4 py-3 rounded-md shadow-lg transform translate-y-20 opacity-0 transition-all duration-300 flex items-center';
            }, 3000);
        }

        // 切换转换方向
        function switchConversionDirection(direction) {
            conversionDirection = direction;
            
            // 更新按钮样式
            if (direction === 'mdToExcel') {
                mdToExcelBtn.className = 'conversion-direction-btn px-6 py-2 rounded-md bg-primary text-white font-medium';
                excelToMdBtn.className = 'conversion-direction-btn px-6 py-2 rounded-md text-gray-600 font-medium';
                
                // 更新输入区域
                inputIcon.className = 'fa fa-file-text-o text-primary text-xl mr-2';
                inputTitle.textContent = 'Markdown表格输入';
                inputDescription.textContent = '请粘贴您的Markdown表格内容：';
                sourceInput.placeholder = `例如：
| 姓名 | 年龄 | 职业 |
|------|------|------|
| 张三 | 25   | 工程师 |
| 李四 | 30   | 设计师 |`;
                
                // 更新输出区域
                outputIcon.className = 'fa fa-file-excel-o text-secondary text-xl mr-2';
                outputTitle.textContent = 'Excel格式输出';
                outputDescription.textContent = '转换后的内容（可直接复制到Excel）：';
                convertBtnText.textContent = '转换为Excel格式';
                
                // 填充示例数据
                sourceInput.value = `| 任务ID | 任务名称 | 负责人 | 状态 | 截止日期 |
| --- | --- | --- | --- | --- |
| 001    | 需求分析 | 张三   | 已完成 | 2023-06-10 |
| 002    | 系统设计 | 李四   | 进行中 | 2023-06-20 |
| 003    | 编码开发 | 王五   | 未开始 | 2023-07-15 |`;
            } else {
                mdToExcelBtn.className = 'conversion-direction-btn px-6 py-2 rounded-md text-gray-600 font-medium';
                excelToMdBtn.className = 'conversion-direction-btn px-6 py-2 rounded-md bg-primary text-white font-medium';
                
                // 更新输入区域
                inputIcon.className = 'fa fa-file-excel-o text-secondary text-xl mr-2';
                inputTitle.textContent = 'Excel表格输入';
                inputDescription.textContent = '请粘贴从Excel复制的表格内容：';
                sourceInput.placeholder = `例如（从Excel复制后会自动保持这种格式）：
姓名	年龄	职业
张三	25	工程师
李四	30	设计师`;
                
                // 更新输出区域
                outputIcon.className = 'fa fa-file-text-o text-primary text-xl mr-2';
                outputTitle.textContent = 'Markdown格式输出';
                outputDescription.textContent = '转换后的Markdown表格内容：';
                convertBtnText.textContent = '转换为Markdown格式';
                
                // 填充示例数据
                sourceInput.value = `任务ID	任务名称	负责人	状态	截止日期
001	需求分析	张三	已完成	2023-06-10
002	系统设计	李四	进行中	2023-06-20
003	编码开发	王五	未开始	2023-07-15`;
            }
            
            // 清空输出
            emptyState.classList.remove('hidden');
            resultContent.classList.add('hidden');
            previewEmptyState.classList.remove('hidden');
            tablePreview.classList.add('hidden');
        }

        // 转换按钮点击事件
        convertBtn.addEventListener('click', async () => {
            const inputContent = sourceInput.value.trim();
            
            if (!inputContent) {
                showNotification('请输入表格内容', false);
                return;
            }
            
            try {
                // 调用后端Python API进行转换
                const response = await fetch('/convert', {
                    method: 'POST',
                    headers: {
                        'Content-Type': 'application/json',
                    },
                    body: JSON.stringify({
                        content: inputContent,
                        direction: conversionDirection
                    })
                });
                
                const result = await response.json();
                
                if (!response.ok) {
                    throw new Error(result.error || '转换失败');
                }
                
                // 更新输出
                convertedOutput.textContent = result.converted_content;
                emptyState.classList.add('hidden');
                resultContent.classList.remove('hidden');
                
                // 更新信息
                conversionInfo.textContent = `转换完成: ${result.columns} 列, ${result.rows} 行`;
                
                // 更新预览
                updateTablePreview(result.headers, result.data);
                
                showNotification('表格转换成功');
            } catch (error) {
                showNotification(error.message, false);
            }
        });

        // 更新表格预览
        function updateTablePreview(headers, data) {
            // 清空现有内容
            tableHeader.innerHTML = '';
            tableBody.innerHTML = '';
            
            // 创建表头
            headers.forEach(header => {
                const th = document.createElement('th');
                th.textContent = header;
                th.className = 'border border-gray-300 px-4 py-2 bg-gray-50 font-semibold text-gray-700';
                tableHeader.appendChild(th);
            });
            
            // 创建表格内容
            data.forEach((row, rowIndex) => {
                const tr = document.createElement('tr');
                tr.className = rowIndex % 2 === 0 ? 'bg-white' : 'bg-gray-50';
                
                row.forEach(cell => {
                    const td = document.createElement('td');
                    td.textContent = cell;
                    td.className = 'border border-gray-300 px-4 py-2 text-gray-700';
                    tr.appendChild(td);
                });
                
                tableBody.appendChild(tr);
            });
            
            // 显示表格预览，隐藏空状态
            previewEmptyState.classList.add('hidden');
            tablePreview.classList.remove('hidden');
        }

        // 复制按钮点击事件
        copyBtn.addEventListener('click', () => {
            const textToCopy = convertedOutput.textContent;
            
            navigator.clipboard.writeText(textToCopy)
                .then(() => {
                    showNotification('内容已复制到剪贴板');
                })
                .catch(err => {
                    showNotification('复制失败，请手动复制', false);
                    console.error('复制失败:', err);
                });
        });

        // 切换预览显示状态
        togglePreviewBtn.addEventListener('click', () => {
            const isExpanded = previewContainer.style.maxHeight !== '0px';
            
            if (isExpanded) {
                previewContainer.style.maxHeight = '0px';
                previewIcon.className = 'fa fa-chevron-down mr-1';
                togglePreviewBtn.querySelector('span').textContent = '显示预览';
            } else {
                previewContainer.style.maxHeight = '500px';
                previewIcon.className = 'fa fa-chevron-up mr-1';
                togglePreviewBtn.querySelector('span').textContent = '隐藏预览';
            }
        });

        // 转换方向按钮事件
        mdToExcelBtn.addEventListener('click', () => {
            switchConversionDirection('mdToExcel');
        });

        excelToMdBtn.addEventListener('click', () => {
            switchConversionDirection('excelToMd');
        });

        // 页面加载时初始化
        window.addEventListener('load', () => {
            switchConversionDirection('mdToExcel');
        });
    </script>
</body>
</html>
'''

# Python实现的转换功能
def parse_markdown_table(markdown):
    """解析Markdown表格，返回表头和数据"""
    lines = [line.strip() for line in markdown.strip().split('\n') if line.strip()]
    
    if len(lines) < 2:
        raise ValueError("表格格式不正确，至少需要表头和分隔线")
    
    # 解析表头
    header_line = lines[0]
    headers = [col.strip() for col in re.sub(r'^[\| ]+|[\| ]+$', '', header_line).split('|')]
    header_count = len(headers)
    
    # 寻找分隔线
    separator_index = -1
    for i, line in enumerate(lines[1:]):
        if line.startswith('|') and line.endswith('|'):
            columns = [col.strip() for col in re.sub(r'^[\| ]+|[\| ]+$', '', line).split('|')]
            if len(columns) == header_count and all('-' in col for col in columns):
                separator_index = i + 1  # 加上1是因为我们从lines[1:]开始循环
                break
    
    if separator_index == -1:
        raise ValueError("未找到有效的分隔线，请检查格式。分隔线应包含 | 和 - 字符，例如：| --- | --- |")
    
    # 解析数据行
    data = []
    for line in lines[separator_index + 1:]:
        row = [col.strip() for col in re.sub(r'^[\| ]+|[\| ]+$', '', line).split('|')]
        data.append(row)
    
    return headers, data

def convert_markdown_to_excel(markdown):
    """将Markdown表格转换为Excel格式（制表符分隔）"""
    headers, data = parse_markdown_table(markdown)
    
    # 构建Excel内容
    excel_lines = ['\t'.join(headers)]
    for row in data:
        excel_lines.append('\t'.join(row))
    
    return '\n'.join(excel_lines), headers, data

def parse_excel_table(excel_text):
    """解析Excel表格（制表符分隔），返回表头和数据"""
    lines = [line.strip() for line in excel_text.strip().split('\n') if line.strip()]
    
    if not lines:
        raise ValueError("未找到表格数据，请确保粘贴了有效的Excel表格内容")
    
    # 解析表头
    headers = [cell.strip() for cell in lines[0].split('\t')]
    
    # 解析数据行
    data = []
    for line in lines[1:]:
        row = [cell.strip() for cell in line.split('\t')]
        data.append(row)
    
    return headers, data

def convert_excel_to_markdown(excel_text):
    """将Excel格式（制表符分隔）转换为Markdown表格"""
    headers, data = parse_excel_table(excel_text)
    column_count = len(headers)
    
    # 计算每列的最大宽度
    column_widths = [len(header) for header in headers]
    for row in data:
        for i, cell in enumerate(row):
            if i < column_count:
                column_widths[i] = max(column_widths[i], len(cell))
    
    # 构建Markdown表格
    markdown_lines = []
    
    # 表头行
    header_line = '| '
    for i, header in enumerate(headers):
        header_line += header.ljust(column_widths[i]) + ' | '
    markdown_lines.append(header_line)
    
    # 分隔线行
    separator_line = '| '
    for width in column_widths:
        separator_line += '-' * width + ' | '
    markdown_lines.append(separator_line)
    
    # 数据行
    for row in data:
        data_line = '| '
        for i, cell in enumerate(row):
            if i < column_count:
                data_line += cell.ljust(column_widths[i]) + ' | '
        markdown_lines.append(data_line)
    
    return '\n'.join(markdown_lines), headers, data

# Flask路由
@app.route('/')
def index():
    """首页"""
    return render_template_string(HTML_TEMPLATE)

@app.route('/convert', methods=['POST'])
def convert():
    """处理转换请求"""
    data = request.get_json()
    content = data.get('content', '')
    direction = data.get('direction', 'mdToExcel')
    
    try:
        if direction == 'mdToExcel':
            converted_content, headers, table_data = convert_markdown_to_excel(content)
        else:
            converted_content, headers, table_data = convert_excel_to_markdown(content)
        
        return jsonify({
            'success': True,
            'converted_content': converted_content,
            'headers': headers,
            'data': table_data,
            'columns': len(headers),
            'rows': len(table_data)
        })
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 400

if __name__ == '__main__':
    app.run(debug=True)

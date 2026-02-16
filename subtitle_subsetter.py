import gradio as gr
import re
import os
from fontTools.ttLib import TTFont
from fontTools.subset import Subsetter

# 全局变量：存储系统中所有字体的信息
all_fonts = {}
# 标记是否已读取字体列表
fonts_loaded = False

version = "0.1.1"

# 解析带样式的文本
def parse_text_with_style(style: str, text: str) -> dict:
    R"""
    解析带 {\...} 控制序列的文本，支持 \fn 字体切换。
    
    - 默认使用传入的 style。
    - 遇到 {...} 时，尝试从中提取 \fn<font>。
    - 字体名从 \fn 后开始，直到遇到 \ 或 } 为止。
    - {...} 本身不输出字符，仅用于切换样式。
    - {...} 后的普通文本按当前样式拆分为字符列表存入 dict。
    """
    text = text.replace("\\n", "").replace("\\N", "").replace("\\h", "") # 去除换行和空格控制符
    result = {}
    current_style = style
    current_chars = []
    
    # 用于遍历：记录上一个处理位置
    pos = 0
    n = len(text)
    
    while pos < n:
        # 查找下一个 '{'
        brace_start = text.find('{', pos)
        if brace_start == -1:
            # 没有更多 {，剩余全是普通文本
            remaining = text[pos:]
            current_chars.extend(list(remaining))
            break
        
        # 先处理 '{' 之前的内容（普通文本）
        plain_text = text[pos:brace_start]
        if plain_text:
            current_chars.extend(list(plain_text))
        
        # 查找匹配的 '}'
        brace_end = text.find('}', brace_start)
        if brace_end == -1:
            # 没有闭合 }，将剩余全部视为普通文本（或报错，这里保守处理）
            remaining = text[brace_start:]
            current_chars.extend(list(remaining))
            break
        
        # 提取 {...} 内容
        control_block = text[brace_start + 1:brace_end]
        
        # 在 control_block 中查找 \fn
        # 使用正则：匹配 \fn 后接非 \ 和非 } 的字符序列
        fn_match = re.search(r'\\fn([^\\}]*)', control_block)
        if fn_match:
            font_name = fn_match.group(1).strip()
            if font_name:
                # 忽略前面的 @ 符号
                if font_name.startswith('@'):
                    font_name = font_name[1:]
                # 保存当前已收集的字符
                if current_chars:
                    result.setdefault(current_style, []).extend(current_chars)
                    current_chars = []
                # 切换样式
                current_style = font_name
        
        # 更新 pos 到 } 之后
        pos = brace_end + 1
    
    # 处理最后剩余的字符
    if current_chars:
        result.setdefault(current_style, []).extend(current_chars)
    
    return result

# 解析 ASS 字幕文件
def parse_ass_file(file_path):
    with open(file_path, "r", encoding="utf-8") as f:
        data = f.read()
    lines = data.splitlines()

    styles = {}
    dialogues = {}

    for line in lines:
        if line.startswith("Style:"):
            style_data = line[len("Style:"):].strip().split(",")
            if len(style_data) > 1:
                style_name = style_data[0].strip()
                font_name = style_data[1].strip()
                # 忽略前面的 @ 符号
                if font_name.startswith('@'):
                    font_name = font_name[1:]
                styles[style_name] = font_name
            continue
        
        if line.startswith("Dialogue:"):
            dialogue_data = line[len("Dialogue:"):].strip().split(",", 9)
            if len(dialogue_data) > 9:
                style_name = dialogue_data[3].strip()
                text = dialogue_data[9].strip()
                if style_name in styles:
                    parsed = parse_text_with_style(styles[style_name], text)
                    for style, chars in parsed.items():
                        dialogues.setdefault(style, []).extend(chars)
            continue

    # 去重
    for style in dialogues:
        dialogues[style] = list(set(dialogues[style]))
    
    return styles, dialogues

# 读取系统中所有字体的信息
def load_all_fonts():
    global all_fonts
    global fonts_loaded
    
    all_fonts = {}
    fonts_loaded = False
    
    try:
        # 字体目录列表
        font_dirs = []
        
        # 1. 系统字体目录
        system_fonts_dir = os.path.join(os.environ.get('WINDIR', 'C:\\Windows'), 'Fonts')
        if os.path.exists(system_fonts_dir):
            font_dirs.append(system_fonts_dir)
        
        # 2. 用户字体目录
        user_fonts_dir = os.path.join(os.environ.get('LOCALAPPDATA', 'C:\\Users\\Default\\AppData\\Local'), 'Microsoft\\Windows\\Fonts')
        if os.path.exists(user_fonts_dir):
            font_dirs.append(user_fonts_dir)
        
        total_font_files = 0
        processed_files = 0
        
        print(f"开始读取系统字体，共 {len(font_dirs)} 个字体目录...")
        
        # 统计所有字体文件
        all_font_files = []
        for fonts_dir in font_dirs:
            font_files = [os.path.join(fonts_dir, file) for file in os.listdir(fonts_dir) if file.endswith(('.ttf', '.otf', '.ttc', '.TTF', '.OTF', '.TTC'))]
            all_font_files.extend(font_files)
        
        total_font_files = len(all_font_files)
        print(f"发现 {total_font_files} 个字体文件...")
        
        # 处理所有字体文件
        for font_path in all_font_files:
            try:
                file = os.path.basename(font_path)
                # 处理 TTC 文件（TrueType Collection）
                if file.endswith(('.ttc', '.TTC')):
                    # 尝试读取 TTC 文件中的第一个字体
                    font = TTFont(font_path, fontNumber=0)
                    # 获取字体名称
                    for record in font['name'].names:
                        if record.nameID == 4:  # 字体全名
                            try:
                                font_full_name = record.toUnicode()
                                all_fonts[font_full_name] = font_path
                            except:
                                pass
                    # 尝试读取 TTC 文件中的第二个字体
                    try:
                        font = TTFont(font_path, fontNumber=1)
                        # 获取字体名称
                        for record in font['name'].names:
                            if record.nameID == 4:  # 字体全名
                                try:
                                    font_full_name = record.toUnicode()
                                    all_fonts[font_full_name] = font_path
                                except:
                                    pass
                    except:
                        pass
                else:
                    # 处理普通字体文件
                    font = TTFont(font_path)
                    # 获取字体名称
                    for record in font['name'].names:
                        if record.nameID == 4:  # 字体全名
                            try:
                                font_full_name = record.toUnicode()
                                all_fonts[font_full_name] = font_path
                            except:
                                pass
            except Exception as font_error:
                # 字体文件读取失败，跳过
                pass
            
            processed_files += 1
            if processed_files % 10 == 0:
                print(f"已处理 {processed_files}/{total_font_files} 个字体文件...")
        
        print(f"字体读取完成，共读取到 {len(all_fonts)} 个唯一字体")
        fonts_loaded = True
    except Exception as e:
        print(f"读取字体时出错: {e}")

# 检查字体是否安装
def check_font_installed(font_name):
    global all_fonts
    
    # 检查字体名称是否在字体字典中
    for font_full_name in all_fonts:
        if font_name in font_full_name:
            return True
    return False

# 根据字体名称查找字体路径
def find_font_path(font_name):
    global all_fonts
    
    # 查找字体路径
    for font_full_name in all_fonts:
        if font_name in font_full_name:
            return all_fonts[font_full_name]
    return None

# 生成8位随机字符串
def generate_random_name():
    import random
    import string
    return ''.join(random.choices(string.ascii_uppercase + string.digits, k=8))

# 子集化字体
def subset_font(font_path, chars, output_path, random_name):
    try:
        # 尝试处理 TTC 文件，使用不同的字体编号
        if font_path.endswith(('.ttc', '.TTC')):
            # 尝试不同的字体编号
            for font_number in range(5):  # 尝试前 5 个字体
                try:
                    font = TTFont(font_path, fontNumber=font_number)
                    subsetter = Subsetter()
                    subsetter.populate(text=''.join(chars))
                    subsetter.subset(font)

                    name_table = font["name"]
                    for record in name_table.names:
                        if record.nameID == 1: # Family Name
                            encoding = record.getEncoding()
                            record.string = random_name.encode(encoding) if encoding else random_name.encode("utf-8")
                            
                    font.save(output_path)
                    return True
                except Exception as e:
                    # 如果是字体编号错误，继续尝试下一个编号
                    if "specify a font number" in str(e):
                        continue
                    else:
                        print(f"子集化字体时出错: {e}")
                        return False
        else:
            # 处理普通字体文件
            font = TTFont(font_path)
            subsetter = Subsetter()
            subsetter.populate(text=''.join(chars))
            subsetter.subset(font)

            name_table = font["name"]
            for record in name_table.names:
                if record.nameID == 1: # Family Name
                    encoding = record.getEncoding()
                    record.string = random_name.encode(encoding) if encoding else random_name.encode("utf-8")
                    
            font.save(output_path)
            return True
    except Exception as e:
        print(f"子集化字体时出错: {e}")
        return False

# 主函数
def process_subtitle(file):
    # 解析字幕文件
    styles, dialogues = parse_ass_file(file.name)
    
    # 提取字体列表
    font_list = list(dialogues.keys())
    
    # 检查字体安装状态
    all_installed = True
    uninstalled_fonts = []
    font_paths = {}
    for font in font_list:
        installed = check_font_installed(font)
        if not installed:
            all_installed = False
            uninstalled_fonts.append(font)
        else:
            # 查找字体路径
            font_path = find_font_path(font)
            if font_path:
                font_paths[font] = font_path
    
    return font_list, all_installed, uninstalled_fonts, font_paths, dialogues

# 检查字体状态
def check_fonts(font_list):
    all_installed = True
    uninstalled_fonts = []
    for font in font_list:
        installed = check_font_installed(font)
        if not installed:
            all_installed = False
            uninstalled_fonts.append(font)
    return all_installed, uninstalled_fonts

# 子集化处理
def perform_subsetting(dialogues, font_paths, subtitle_file, output_dir):
    try:
        # 确保输出目录存在
        os.makedirs(output_dir, exist_ok=True)
        
        # 生成的字体文件映射：原字体名称 -> 子集化字体名称
        font_mapping = {}
        
        # 处理每个字体
        for font_name, font_path in font_paths.items():
            # 获取该字体使用的字符
            chars = dialogues.get(font_name, [])
            if not chars:
                continue
            
            # 生成8位随机文件名
            random_name = generate_random_name()
            # 保持原字体文件扩展名
            ext = os.path.splitext(font_path)[1]
            output_path = os.path.join(output_dir, f"{random_name}{ext}")
            
            # 执行子集化
            success = subset_font(font_path, chars, output_path, random_name)
            if success:
                # 去掉文件后缀名
                subset_font_name = os.path.basename(output_path)
                subset_font_name_no_ext = os.path.splitext(subset_font_name)[0]
                font_mapping[font_name] = subset_font_name_no_ext
                print(f"成功子集化字体: {font_name} -> {subset_font_name_no_ext}")
            else:
                print(f"子集化失败: {font_name}")
        
        if font_mapping:
            # 修改字幕文件
            modify_subtitle_file(subtitle_file, font_mapping, output_dir)
            return f"子集化完成！生成了 {len(font_mapping)} 个字体文件，并修改了字幕文件"
        else:
            return "子集化失败，未生成任何字体文件"
    except Exception as e:
        print(f"子集化处理时出错: {e}")
        return f"子集化失败: {str(e)}"

# 修改字幕文件
def modify_subtitle_file(subtitle_file, font_mapping, output_dir):
    try:
        # 读取字幕文件，使用utf-8-sig编码处理BOM
        with open(subtitle_file, 'r', encoding='utf-8-sig') as f:
            lines = f.readlines()
        
        # 查找[Script Info]部分
        script_info_index = -1
        for i, line in enumerate(lines):
            if line.strip().startswith('[Script Info]'):
                script_info_index = i
                break
        
        # 生成字体子集化信息
        import datetime
        current_time = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        font_subset_info = []
        font_subset_info.append('; ----- Font Subset Information -----\n')
        for original_font, subset_font in font_mapping.items():
            font_subset_info.append(f'; Font Subset: {subset_font} - {original_font}\n')
        font_subset_info.append('\n')
        font_subset_info.append(f'; Generated by Subsetter v{version} at {current_time}\n')
        font_subset_info.append('; ----- Font Subset Information End -----\n')
        
        # 在[Script Info]后插入字体子集化信息
        if script_info_index != -1:
            # 直接修改lines列表，在[Script Info]行之后插入
            new_lines = lines[:script_info_index+1] + font_subset_info + lines[script_info_index+1:]
            lines = new_lines
        else:
            # 如果没有找到[Script Info]，在文件开头添加
            lines = font_subset_info + lines
        
        # 分析字幕文件，找出所有style
        all_styles = {}
        style_section_start = -1
        style_section_end = -1
        
        # 找出[V4+ Styles]部分的开始和结束位置
        for i, line in enumerate(lines):
            if line.strip().startswith('[V4+ Styles]'):
                style_section_start = i
            elif style_section_start != -1 and line.strip().startswith('['):
                style_section_end = i
                break
        if style_section_end == -1 and style_section_start != -1:
            style_section_end = len(lines)
        
        # 提取所有style，跳过Comment行
        if style_section_start != -1:
            for line in lines[style_section_start+1:style_section_end]:
                if line.startswith('Style:'):
                    parts = line.split(',')
                    if len(parts) > 1:
                        font_name = parts[1].strip()
                        style_name = line.split(',')[0].split(':')[1].strip()
                        all_styles[style_name] = (font_name, line)
        
        # 过滤出使用的style：所有被子集化的字体对应的style
        used_style_lines = []
        if style_section_start != -1:
            used_style_lines.append(lines[style_section_start])  # 保留[V4+ Styles]行
            for style_name, (font_name, style_line) in all_styles.items():
                if font_name in font_mapping:
                    used_style_lines.append(style_line)
        
        # 重建字幕文件内容，跳过[Event]部分中的Comment行
        rebuilt_lines = []
        i = 0
        in_event_section = False
        while i < len(lines):
            line = lines[i]
            # 检查是否进入或退出[Event]部分
            if line.strip().startswith('[Events]'):
                in_event_section = True
                rebuilt_lines.append(line)
                i += 1
            elif in_event_section and line.strip().startswith('['):
                in_event_section = False
                rebuilt_lines.append(line)
                i += 1
            elif style_section_start != -1 and i == style_section_start:
                # 替换为过滤后的style
                rebuilt_lines.extend(used_style_lines)
                # 跳过原有的style部分
                i = style_section_end
            elif in_event_section and line.strip().startswith('Comment:'):
                # 跳过[Event]部分中的Comment行
                i += 1
            else:
                rebuilt_lines.append(line)
                i += 1
        
        # 更新行数，用于后续处理
        lines = rebuilt_lines
        
        # 替换字幕文件中的字体名称
        modified_lines = []
        for line in lines:
            modified_line = line
            for original_font, subset_font in font_mapping.items():
                # 替换Style部分的字体名称
                if line.startswith('Style:'):
                    parts = line.split(',')
                    if len(parts) > 1 and parts[1].strip() == original_font:
                        parts[1] = subset_font  # 去掉空格
                        modified_line = ','.join(parts)
                # 替换\fn后的字体名称
                modified_line = modified_line.replace(f'\\fn{original_font}', f'\\fn{subset_font}')
            modified_lines.append(modified_line)
        
        # 生成输出文件名：原文件名_subset.ass
        base_name = os.path.basename(subtitle_file)
        name_without_ext = os.path.splitext(base_name)[0]
        output_subtitle = os.path.join(output_dir, f"{name_without_ext}_subset.ass")
        with open(output_subtitle, 'w', encoding='utf-8') as f:
            f.writelines(modified_lines)
        
        # 计算删除的样式数量
        used_style_count = 0
        for style_name, (font_name, style_line) in all_styles.items():
            if font_name in font_mapping:
                used_style_count += 1
        deleted_style_count = len(all_styles) - used_style_count
        
        print(f"字幕文件已修改并保存到: {output_subtitle}")
        print(f"删除了 {deleted_style_count} 个未使用的style")
        print(f"删除了[Event]部分中的所有Comment行")
    except Exception as e:
        print(f"修改字幕文件时出错: {e}")
        raise

# 初始化函数：读取字体列表
def initialize_app():
    load_all_fonts()
    return fonts_loaded

# 重新读取字体列表
def reload_fonts():
    load_all_fonts()
    return f"字体重新读取完成，共 {len(all_fonts)} 个字体"

# 初始化应用并启用组件
def init_and_enable():
    success = initialize_app()
    if success:
        return "字体读取完成，可以上传字幕文件", gr.update(interactive=True), gr.update(interactive=True)
    else:
        return "字体读取失败", gr.update(interactive=False), gr.update(interactive=False)

# 创建 Gradio 界面
with gr.Blocks() as demo:
    gr.Markdown("# 字幕字体子集化工具")
    
    # 初始化状态
    initialization_status = gr.Textbox(label="初始化状态", value="正在读取系统字体...", interactive=False)
    
    # 上传组件（默认禁用，字体读取完成后启用）
    file_input = gr.File(label="上传 ASS 字幕文件", file_types=[".ass"], interactive=False)
    
    # 保存位置选择
    output_dir_input = gr.Textbox(label="子集化字体保存位置", value="Subset", interactive=True, placeholder="请输入保存文件夹路径")
    
    # 字体列表和状态
    font_list_output = gr.JSON(label="使用的字体列表")
    uninstalled_fonts_output = gr.JSON(label="未安装的字体")
    all_installed_output = gr.Textbox(label="安装状态", interactive=False)
    
    # 按钮
    subset_button = gr.Button("子集化", interactive=False)
    check_button = gr.Button("再次检查字体", interactive=False)
    reload_button = gr.Button("重新读取系统字体")
    reload_status = gr.Textbox(label="重新读取状态", interactive=False)
    
    # 输出
    result_output = gr.Textbox(label="结果", interactive=False)
    
    # 存储状态
    font_list_state = gr.State([])
    uninstalled_fonts_state = gr.State([])
    font_paths_state = gr.State({})
    dialogues_state = gr.State({})
    
    # 应用启动时执行初始化
    demo.load(
        fn=init_and_enable,
        inputs=[],
        outputs=[initialization_status, file_input, check_button]
    )
    
    # 上传文件后处理
    def on_file_upload(file):
        if file is None:
            return [], [], "请上传 ASS 字幕文件", [], [], {}, {}, gr.update(interactive=False), ""
        font_list, all_installed, uninstalled_fonts, font_paths, dialogues = process_subtitle(file)
        if all_installed:
            status_text = "所有字体均已安装"
        else:
            status_text = f"有 {len(uninstalled_fonts)} 个字体未安装"
        return font_list, uninstalled_fonts, status_text, font_list, uninstalled_fonts, font_paths, dialogues, gr.update(interactive=all_installed), ""
    
    # 绑定上传事件
    file_input.change(
        fn=on_file_upload,
        inputs=[file_input],
        outputs=[font_list_output, uninstalled_fonts_output, all_installed_output, font_list_state, uninstalled_fonts_state, font_paths_state, dialogues_state, subset_button, result_output]
    )
    
    # 再次检查字体
    def on_check_fonts(font_list):
        if not font_list:
            return [], "请先上传 ASS 字幕文件", gr.update(interactive=False)
        all_installed, uninstalled_fonts = check_fonts(font_list)
        status_text = "所有字体均已安装" if all_installed else f"有 {len(uninstalled_fonts)} 个字体未安装"
        return uninstalled_fonts, status_text, gr.update(interactive=all_installed)
    
    # 绑定检查按钮事件
    check_button.click(
        fn=on_check_fonts,
        inputs=[font_list_state],
        outputs=[uninstalled_fonts_output, all_installed_output, subset_button]
    )
    
    # 重新读取字体
    def on_reload_fonts():
        status = reload_fonts()
        return status
    
    # 绑定重新读取按钮事件
    reload_button.click(
        fn=on_reload_fonts,
        inputs=[],
        outputs=[reload_status]
    )
    
    # 子集化处理
    def on_subset(dialogues, font_paths, file, output_dir):
        if file is None:
            return "请先上传字幕文件"
        if not output_dir:
            return "请输入保存文件夹路径"
        result = perform_subsetting(dialogues, font_paths, file.name, output_dir)
        return result
    
    # 绑定子集化按钮事件
    subset_button.click(
        fn=on_subset,
        inputs=[dialogues_state, font_paths_state, file_input, output_dir_input],
        outputs=[result_output]
    )

# 运行界面
if __name__ == "__main__":
    demo.launch(share=False)

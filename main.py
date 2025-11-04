import json
import os
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
import html

class PPTXParser:
    def __init__(self, input_dir="input", output_dir="output"):
        self.input_dir = input_dir
        self.output_dir = output_dir
        self.output_img_dir = os.path.join(output_dir, "images")
        
        # 创建输出目录
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)
        if not os.path.exists(self.output_img_dir):
            os.makedirs(self.output_img_dir)

    def parse_all_pptx(self):
        """解析input文件夹中的所有PPTX文件"""
        # 获取input文件夹中所有pptx文件
        pptx_files = [f for f in os.listdir(self.input_dir) 
                     if f.lower().endswith(('.pptx', '.ppt'))]
        
        if not pptx_files:
            print(f"在 {self.input_dir} 文件夹中未找到PPT/PPTX文件")
            return
        
        for file_name in pptx_files:
            print(f"\n正在解析文件: {file_name}")
            file_path = os.path.join(self.input_dir, file_name)
            
            # 为每个文件创建解析器实例
            parser = SingleFileParser(file_path, self.output_img_dir)
            results = parser.parse_pptx()
            
            # 保存JSON结果到output文件夹
            output_json = os.path.join(self.output_dir, f"{parser.file_name}_parsed.json")
            parser.save_to_json(output_json)
            
            # 打印结果摘要
            print(f"文件 {file_name} 解析完成！")
            text_count = len([r for r in results if r['type'] == 'text'])
            table_count = len([r for r in results if r['type'] == 'table'])
            image_count = len([r for r in results if r['type'] == 'image'])
            
            print(f"总共解析出 {len(results)} 个元素")
            print(f"文本元素: {text_count} 个")
            print(f"表格元素: {table_count} 个")
            print(f"图片元素: {image_count} 个")

class SingleFileParser:
    def __init__(self, file_path, output_img_dir="images"):
        self.file_path = file_path
        self.output_img_dir = output_img_dir
        self.file_name = os.path.splitext(os.path.basename(file_path))[0]
        self.results = []
        self.img_counter = 0

    def parse_pptx(self):
        """解析 PPTX 文件"""
        prs = Presentation(self.file_path)
        
        for page_idx, slide in enumerate(prs.slides):
            print(f"正在解析第 {page_idx + 1} 页...")
            
            # 首先处理标题
            self._process_title(slide, page_idx)
            
            # 然后处理其他形状
            self._process_other_shapes(slide, page_idx)
        
        return self.results
    
    def _process_title(self, slide, page_idx):
        """处理标题"""
        if slide.shapes.title:
            title_text = slide.shapes.title.text.strip()
            if title_text:
                unique_id = len(self.results)
                self.results.append({
                    "type": "text",
                    "text": title_text,
                    "text_level": 1,  # 标题层级为1
                    "page_idx": page_idx,
                    "unique_id": unique_id,
                    "file_name": self.file_name,
                    "img_path": ""
                })
    
    def _process_other_shapes(self, slide, page_idx):
        """处理其他形状（正文、表格、图片）"""
        shape_id = 0
        for shape in slide.shapes:
            # 跳过标题形状，因为已经在 _process_title 中处理过了
            if shape == slide.shapes.title:
                continue
            
            unique_id = len(self.results)
            
            # 处理文本框（正文）
            if shape.has_text_frame:
                self._process_text_shape(shape, page_idx, unique_id, shape_id)
                shape_id += 1
            
            # 处理表格
            elif shape.shape_type == MSO_SHAPE_TYPE.TABLE:
                self._process_table_shape(shape, page_idx, unique_id)
                shape_id += 1
            
            # 处理图片
            elif shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                self._process_image_shape(shape, page_idx, unique_id)
                shape_id += 1
            
            # 处理组合形状（可能包含文本）
            elif shape.shape_type == MSO_SHAPE_TYPE.GROUP:
                self._process_group_shape(shape, page_idx, unique_id, shape_id)
                shape_id += 1
    
    def _process_text_shape(self, shape, page_idx, unique_id, shape_id):
        """处理文本形状（正文）"""
        text_frame = shape.text_frame
        text_content = text_frame.text.strip()
        
        if text_content:
            # 确定文本层级
            text_level = self._determine_text_level(text_frame)
            
            self.results.append({
                "type": "text",
                "text": text_content,
                "text_level": 9,
                "page_idx": page_idx,
                "unique_id": unique_id,
                "file_name": self.file_name,
                "img_path": ""
            })
    
    def _determine_text_level(self, text_frame):
        """确定文本层级"""
        # 如果有多个段落，使用第一个段落的级别
        if text_frame.paragraphs:
            first_paragraph = text_frame.paragraphs[0]
            # paragraph.level 返回缩进级别，0=顶级，1=一级缩进，以此类推
            # 我们将其转换为：1=标题，2=一级正文，3=二级正文...
            return first_paragraph.level + 2 if first_paragraph.level is not None else 2
        return 2  # 默认正文层级
    
    def _process_table_shape(self, shape, page_idx, unique_id):
        """处理表格形状"""
        try:
            table = shape.table
            html_table = self._table_to_html(table)
            
            # 保存表格为HTML文件
            table_filename = f"{self.file_name}_table_{page_idx}_{self.img_counter}.html"
            table_path = os.path.join(self.output_img_dir, table_filename)
            
            with open(table_path, 'w', encoding='utf-8') as f:
                f.write(html_table)
            
            self.results.append({
                "type": "table",
                "text": "",  # 表格内容在html文件中
                "text_level": 0,
                "page_idx": page_idx,
                "unique_id": unique_id,
                "file_name": self.file_name,
                "img_path": table_path
            })
            
            self.img_counter += 1
            
        except Exception as e:
            print(f"处理表格时出错: {e}")
    
    def _table_to_html(self, table):
        """将表格转换为HTML格式"""
        html_content = ['<table border="1">']
        
        for row in table.rows:
            html_content.append('<tr>')
            for cell in row.cells:
                cell_text = cell.text_frame.text.strip() if cell.text_frame else ""
                # 转义HTML特殊字符
                cell_text = html.escape(cell_text)
                html_content.append(f'<td>{cell_text}</td>')
            html_content.append('</tr>')
        
        html_content.append('</table>')
        return ''.join(html_content)
    
    def _process_image_shape(self, shape, page_idx, unique_id):
        """处理图片形状"""
        try:
            image = shape.image
            # 获取图片二进制数据
            image_data = image.blob
            image_extension = image.ext  # 图片扩展名
            
            # 生成图片文件名
            img_filename = f"{self.file_name}_img_{page_idx}_{self.img_counter}.{image_extension}"
            img_path = os.path.join(self.output_img_dir, img_filename)
            
            # 保存图片
            with open(img_path, 'wb') as f:
                f.write(image_data)
            
            self.results.append({
                "type": "image",
                "text": "",  # 图片没有文本内容
                "text_level": 0,
                "page_idx": page_idx,
                "unique_id": unique_id,
                "file_name": self.file_name,
                "img_path": img_path
            })
            
            self.img_counter += 1
            
        except Exception as e:
            print(f"处理图片时出错: {e}")
    
    def _process_group_shape(self, shape, page_idx, unique_id, shape_id):
        """处理组合形状"""
        try:
            # 递归处理组合形状中的子形状
            for sub_shape in shape.shapes:
                if sub_shape.has_text_frame:
                    self._process_text_shape(sub_shape, page_idx, unique_id, shape_id)
        except Exception as e:
            print(f"处理组合形状时出错: {e}")
    
    def save_to_json(self, output_file=None):
        """保存解析结果到JSON文件"""
        if not output_file:
            output_file = f"{self.file_name}_parsed.json"
        
        with open(output_file, 'w', encoding='utf-8') as f:
            json.dump(self.results, f, ensure_ascii=False, indent=2)
        
        print(f"解析完成！结果已保存到: {output_file}")
        return output_file

# 使用示例
def main():
    # 解析input文件夹中的所有PPTX文件
    parser = PPTXParser(input_dir="input", output_dir="output")
    parser.parse_all_pptx()

if __name__ == "__main__":
    main()
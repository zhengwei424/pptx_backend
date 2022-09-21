import pptx
import pptx.slide
import pptx.table
import pptx.text.text
import pptx.shapes.base
import pptx.chart.chart
from pptx.util import Pt, Cm, Emu
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.chart.data import ChartData
from pptx.enum.lang import MSO_LANGUAGE_ID
from pptx.shapes.placeholder import SlidePlaceholder
from pptx.enum.text import PP_PARAGRAPH_ALIGNMENT as PP_ALIGN
from pptx.enum.chart import XL_LEGEND_POSITION, XL_CHART_TYPE, XL_DATA_LABEL_POSITION


# presentation -> slide -> shapes -> placeholder,graphfrm -> chart(table -> cell) -> text_frame -> paragraphs -> font
#            |      +-> placeholder
#            |->slide_master -> slide_layout
# https://mhammond.github.io/pywin32/html/com/win32com/HTML/QuickStartClientCom.html

class PresentationBuilder(object):
    def __init__(self, prs_path):
        self.prs_path = prs_path
        self.prs = Presentation(self.prs_path)

    @property
    def xml_slides(self):
        return self.prs.slides._sldIdLst

    # 插入slide_layout到指定位置
    def insert_slide_by_slide_layout(self, index, slide_layout):
        # type -> pptx.oxml.presentation.CT_SlideIdList
        slideIdList = list(self.xml_slides)
        self.prs.slides.add_slide(slide_layout)
        pop_item = slideIdList.pop()
        self.xml_slides.insert(index, pop_item)

    # 复制一个slide到指定位置
    def copy_slide(self):
        pass

    # 删除指定位置的slide
    def delete_slide(self, index):
        slideIdList = list(self.xml_slides)
        self.xml_slides.remove(slideIdList[index])


# 先写内容，后改样式，否则可能出现样式不生效问题，如字体之类的？？？
def set_cell_format(cell: pptx.table._Cell,
                    cell_background_color=None,
                    margin_left=None,
                    margin_right=None,
                    margin_top=None,
                    margin_bottom=None):
    # cell中文字相对cell边框的距离，类似与html中的padding
    cell.margin_left = margin_left
    cell.margin_right = margin_right
    cell.margin_top = margin_top
    cell.margin_bottom = margin_bottom

    if cell_background_color:
        # 填充cell前景色
        # cell.fill.patterned()  # 图案填充（可以设置cell前景色和背景色）
        cell.fill.solid()  # 纯色填充（只能设置cell前景色，但是对cell中的text而言也是背景）
        # cell.fill.back_color.rgb = RGBColor.from_string(cell_background_color)
        cell.fill.fore_color.rgb = RGBColor.from_string(cell_background_color)
    # else:
    #     # cell颜色无填充
    #     cell.fill.background()


# 设置text_frame的paragraph文字格式
def set_text_frame_paragraph_format(
        tfp: pptx.text.text._Paragraph,
        font_name="Microsoft YaHei",
        font_size=8,
        font_bold=False,
        align=PP_ALIGN.LEFT,
        font_color="000000"):
    # 语言设置，NONE表示移除所有语言设置
    tfp.font.language_id = MSO_LANGUAGE_ID.NONE
    # 字体
    tfp.font.name = font_name
    # 字体大小
    tfp.font.size = Pt(int(font_size))
    # 是否加粗
    tfp.font.bold = font_bold
    # 水平对齐
    tfp.alignment = align

    # 用RGB表示字体颜色（两种方式）
    # tfp.font.color.rgb = RGBColor.from_string(fontColor)
    # 设置前景色或背景色之前需要执行
    # tfp.font.fill.patterned()  # 图案填充（可以设置前景色和背景色）
    tfp.font.fill.solid()  # 纯色填充（只能设置前景色）
    # 字体前景色(就是字体颜色）
    tfp.font.fill.fore_color.rgb = RGBColor.from_string(font_color)
    # 字体背景色（就是文字本身的背景，不是整个cell）
    # tfp.font.fill.back_color.rgb = RGBColor.from_string(cell_background_color)
    # 字体颜色透明度
    # tfp.font.color.brightness = -1  # 取值范围-1~1，暗->亮


class TableAttribute:
    def __init__(self, table):
        self.tb = table  # type: pptx.table.Table


class WeaklyReports(object):
    def __init__(self, prs):
        self._prs = prs

    # 1. 运维工作统计（次数）
    def slide_1(self, events_count: list):  #
        if len(events_count) != 6:
            print("events_count length is 6,Please check out.")
        slide = self._prs.slides[8]  # type:pptx.slide.Slide
        shape = slide.shapes[0]  # type: pptx.shapes.base.BaseShape
        index = 0
        for shape in slide.shapes:
            '''
            # 判断shape是否是图表
            if shape.has_chart:
                # https://python-pptx.readthedocs.io/en/latest/user/charts.html
                chart = sl1.shapes[index].chart  # type: pptx.chart.chart.Chart
                # 判断chart类型
                if chart.chart_type == XL_CHART_TYPE.COLUMN_CLUSTERED:  # 柱状图
                    chart_data = CategoryChartData()
                    chart_data.categories = ["a", "b", "c"]  # 类别
                    chart_data.add_series('Series 1', (19.2, 21.4, 16.7))  # 系列（即每个类别可以有多个系列）
    
                if chart.chart_type == XL_CHART_TYPE.PIE:  # 饼图
                    chart_data = CategoryChartData()
                    chart_data.categories = ['West', 'East', 'North', 'South', 'Other']  # 类别
                    chart_data.add_series('Series 1', (0.135, 0.324, 0.180, 0.235, 0.126))  # 系列（即每个类别可以有多个系列）
            # 判断shape是否包含text
            if shape.has_text_frame:
                tf = sl1.shapes[index].text_frame  # type: pptx.text.text.TextFrame
            '''

            # 判断shape是否是表格（即找到需要修改的表格）
            if shape.has_table:
                tb = shape.table  # type: pptx.table.Table
                tb.cell(3, 1).text = str(events_count[0])  # 变更
                tb.cell(3, 2).text = str(events_count[1])  # 资源权限管理
                tb.cell(3, 3).text = str(events_count[2])  # 配合操作
                tb.cell(3, 4).text = str(events_count[3])  # 支撑发版
                tb.cell(3, 5).text = str(events_count[4])  # 问题和告警处理
                tb.cell(3, 6).text = str(events_count[5])  # 故障处理
            index += 1

    # 2. 巡检
    def slide_2(self, weekly_inspect: list):
        if not weekly_inspect:
            weekly_inspect = ["11", "0", "6", "√", "√", "√", "√", "√", "√", "√", "√", "√", "√", "√", "√", "√", "√"]
        slide = self._prs.slides[9]  # type:pptx.slide.Slide
        shape = slide.shapes[0]  # type: pptx.shapes.base.BaseShape
        for shape in slide.shapes:
            if shape.has_table:
                tb = shape.table  # type: pptx.table.Table
                cells = [
                    tb.cell(3, 1),
                    tb.cell(3, 3),
                    tb.cell(3, 5),
                    tb.cell(4, 5),
                    tb.cell(3, 7),
                    tb.cell(4, 7),
                    tb.cell(3, 9),
                    tb.cell(4, 9),
                    tb.cell(3, 11),
                    tb.cell(4, 11),
                    tb.cell(3, 13),
                    tb.cell(4, 13),
                    tb.cell(3, 15),
                    tb.cell(4, 15),
                    tb.cell(3, 17),
                    tb.cell(4, 17),
                ]
                for index in range(len(cells)):
                    cells[index].text = weekly_inspect[index]
                    set_text_frame_paragraph_format(cells[index].text_frame.paragraphs[0])
                # tb.cell(3, 1).text = "11"  # 巡检次数
                # tb.cell(3, 2).text = "0"  # 异常次数
                # tb.cell(3, 3).text = "6"  # 报告提交次数
                # tb.cell(3, 5).text = "√"  # 周一上午
                # tb.cell(4, 5).text = "√"  # 周一下午
                # tb.cell(3, 7).text = "√"  # 周二上午
                # tb.cell(4, 7).text = "√"  # 周二下午
                # tb.cell(3, 9).text = "√"  # 周三上午
                # tb.cell(4, 9).text = "√"  # 周三下午
                # tb.cell(3, 11).text = "√"  # 周四上午
                # tb.cell(4, 11).text = "√"  # 周四下午
                # tb.cell(3, 13).text = "√"  # 周五上午
                # tb.cell(4, 13).text = "√"  # 周五下午
                # tb.cell(3, 15).text = "√"  # 周六上午
                # tb.cell(4, 15).text = "√"  # 周六下午
                # tb.cell(3, 17).text = "√"  # 周日上午
                # tb.cell(4, 17).text = "√"  # 周日下午

    # 3. 变更
    def slide_3(self, content: list):
        slide = self._prs.slides[10]  # type:pptx.slide.Slide
        shape = slide.shapes[0]  # type: pptx.shapes.base.BaseShape
        for shape in slide.shapes:
            if shape.has_table:
                tb = shape.table  # type: pptx.table.Table
                for row_idx in range(1, len(tb.rows)):
                    for col_idx in range(len(tb.columns)):
                        tb.cell(row_idx, col_idx).text = content[row_idx - 1][col_idx]

    # 4. 支撑发版
    def slide_4(self, content: list):
        slide = self._prs.slides[11]  # type:pptx.slide.Slide
        shape = slide.shapes[0]  # type: pptx.shapes.base.BaseShape
        for shape in slide.shapes:
            if shape.has_table:
                tb = shape.table  # type: pptx.table.Table
                for row_idx in range(1, len(tb.rows)):
                    for col_idx in range(len(tb.columns)):
                        tb.cell(row_idx, col_idx).text = content[row_idx - 1][col_idx]

    # 5. 资源权限管理
    def slide_5(self, content: list):
        slide = self._prs.slides[12]  # type:pptx.slide.Slide
        shape = slide.shapes[0]  # type: pptx.shapes.base.BaseShape
        for shape in slide.shapes:
            if shape.has_table:
                tb = shape.table  # type: pptx.table.Table
                for row_idx in range(1, len(tb.rows)):
                    for col_idx in range(len(tb.columns)):
                        tb.cell(row_idx, col_idx).text = content[row_idx - 1][col_idx]

    # 6. 配合操作及排障
    def slide_6(self, content: list):
        slide = self._prs.slides[13]  # type:pptx.slide.Slide
        shape = slide.shapes[0]  # type: pptx.shapes.base.BaseShape
        for shape in slide.shapes:
            if shape.has_table:
                tb = shape.table  # type: pptx.table.Table
                for row_idx in range(1, len(tb.rows)):
                    for col_idx in range(len(tb.columns)):
                        tb.cell(row_idx, col_idx).text = content[row_idx - 1][col_idx]

    # 7. 问题及告警
    def slide_7(self, content: list):
        slide = self._prs.slides[14]  # type:pptx.slide.Slide
        shape = slide.shapes[0]  # type: pptx.shapes.base.BaseShape
        for shape in slide.shapes:
            if shape.has_table:
                tb = shape.table  # type: pptx.table.Table
                for row_idx in range(1, len(tb.rows)):
                    for col_idx in range(len(tb.columns)):
                        tb.cell(row_idx, col_idx).text = content[row_idx - 1][col_idx]

    # 8. 运行情况分析
    def slide_8(self, cluster_pie_data, cluster_table_data):
        # 传入集群应用部署信息
        for cluster in cluster_table_data.keys():
            table_data = cluster_table_data[cluster]

            # 定义每个集群（运行情况分析）ppt页数
            # 每页最多存75(80)行数据
            table_pg_num = 0
            if cluster == "fcp":
                table_pg_num = 2
            else:
                table_pg_num = 1

            # 集群名称
            cluster_name = ""
            if cluster == "fcp":
                cluster_name = "FCP业务集群"
            elif cluster == "fcp-inner-microservice":
                cluster_name = "内网微服务区"
            elif cluster == "fcp-inner-backend":
                cluster_name = "内网后台区"
            elif cluster == "fcp-outer-microservice":
                cluster_name = "外网微服务区"
            elif cluster == "fcp-outer-backend":
                cluster_name = "外网后台区"
            # 集群合计信息
            total_count = table_data[-1]
            # 集群资源已分配数值
            cluster_allocated_cpu = cluster_pie_data.get(cluster).get("cpu").get("allocated")
            cluster_allocated_memory = cluster_pie_data.get(cluster).get("memory").get("allocated")
            # 集群资源总数值
            cluster_total_cpu = cluster_pie_data.get(cluster).get("cpu").get("total")
            cluster_total_memory = cluster_pie_data.get(cluster).get("memory").get("total")
            # 集群资源已分配比率
            cpu_allocated_rate = float("%.2f" % (cluster_allocated_cpu / cluster_total_cpu))
            memory_allocated_rate = float("%.2f" % (cluster_allocated_memory / cluster_total_memory))
            for page in range(1, table_pg_num + 1):
                # 添加一张幻灯片
                slide = self._prs.slides.add_slide(self._prs.slide_masters[0].slide_layouts[1])

                # 幻灯片标题
                slidePlaceholder = slide.shapes[0]  # type: pptx.shapes.placeholder.SlidePlaceholder
                slidePlaceholder.text = "2.11慧企运行情况分析–集群资源使用情况"

                # cpu 扇形图
                chart_cpu_data = ChartData()
                chart_cpu_data.categories = ['已分配', '未分配']
                # 可以设置多组数据填入扇形图编辑时的excel中，但只有第一组数据会展示
                chart_cpu_data.add_series('cpu', (cpu_allocated_rate, 1 - cpu_allocated_rate))
                # chart_cpu_data.add_series('xxx2', (0.35, 0.65))
                # chart_cpu_data.add_series('xxx3', (0.45, 0.55))
                shape_cpu = slide.shapes.add_chart(XL_CHART_TYPE.PIE, Cm(0.1), Cm(2.2), Cm(6), Cm(6), chart_cpu_data)
                chart_cpu = shape_cpu.chart

                # 设置图例说明（会在图中标识已分配、未分配的颜色说明）
                chart_cpu.has_legend = True
                chart_cpu.legend.position = XL_LEGEND_POSITION.BOTTOM
                chart_cpu.legend.include_in_layout = False
                chart_cpu.font.name = 'Microsoft YaHei'
                chart_cpu.font.size = Pt(10)

                # 设置标题（不设置，默认值是add_series中的列标题"cpu"）-> "ChartTitle" has no attribute "width"??
                chart_cpu.has_title = True
                chart_cpu.chart_title.has_text_frame = True
                chart_cpu.chart_title.text_frame.text = cluster_name + " CPU(%)"
                set_text_frame_paragraph_format(chart_cpu.chart_title.text_frame.paragraphs[0],
                                                font_size=10,
                                                font_name="Microsoft YaHei")

                chart_cpu.plots[0].has_data_labels = True
                chart_cpu_data_labels = chart_cpu.plots[0].data_labels
                chart_cpu_data_labels.number_format = '0%'
                chart_cpu_data_labels.position = XL_DATA_LABEL_POSITION.CENTER

                # memory 扇形图
                chart_mem_data = ChartData()
                chart_mem_data.categories = ['已分配', '未分配']
                chart_mem_data.add_series('memory', (memory_allocated_rate, 1 - memory_allocated_rate))
                chart_mem = slide.shapes.add_chart(XL_CHART_TYPE.PIE, Cm(0.1), Cm(8.3), Cm(6), Cm(6),
                                                   chart_mem_data).chart

                # 设置图例说明（会在图中标识已分配、未分配的颜色说明）
                chart_mem.has_legend = True
                chart_mem.legend.position = XL_LEGEND_POSITION.BOTTOM
                chart_mem.legend.include_in_layout = False
                chart_mem.font.name = 'Microsoft YaHei'
                chart_mem.font.size = Pt(10)

                # 设置标题（不设置，默认值是add_series中的列标题"cpu")
                chart_mem.has_title = True
                chart_mem.chart_title.has_text_frame = True
                chart_mem.chart_title.text_frame.text = cluster_name + " 内存(%)"
                set_text_frame_paragraph_format(chart_mem.chart_title.text_frame.paragraphs[0],
                                                font_size=10,
                                                font_name="Microsoft YaHei")

                chart_mem.plots[0].has_data_labels = True
                chart_mem_data_labels = chart_mem.plots[0].data_labels
                chart_mem_data_labels.number_format = '0%'
                chart_mem_data_labels.position = XL_DATA_LABEL_POSITION.CENTER
                # 添加table
                if page == 1 and table_pg_num == 1:
                    if len(table_data) <= 40:
                        # table_1
                        # 根据数据长度设置表格的行列数（同时通过行高*行数，计算表格高度，并与模板对比，做表格当前页分割，或换页分割）！！！！！
                        # x: 左边距，y: 上边距, cx: 单元格宽度, cy: 单元格高度
                        x, y, cx, cy = Cm(7.5), Cm(2.2), Cm(1), Cm(0.2)
                        table = slide.shapes.add_table(len(table_data), 3, x, y, cx, cy).table
                        table.columns[0].width = Cm(3.5)
                        table.columns[1].width = Cm(4)
                        table.columns[2].width = Cm(1)
                        index = 0
                        cells = list(table.iter_cells())
                        count = len(table_data)
                        for i in range(count):
                            for j in table_data.pop(0):
                                if index < len(cells):
                                    cells[index].text = str(j)
                                    set_cell_format(cells[index],
                                                    margin_left=0,
                                                    margin_right=0,
                                                    margin_top=0,
                                                    margin_bottom=0)
                                    set_text_frame_paragraph_format(cells[index].text_frame.paragraphs[0], font_size=7)
                                    index += 1

                        # 右下插入文本框
                        txbox = slide.shapes.add_textbox(Cm(16.5), Cm(12.5), Cm(8), Cm(1))
                        # 文本框应该有一个默认段落，直接获取，额外添加段落需要使用text_frame.add_paragraph()，清除段落用text_frame.clear()
                        tf = txbox.text_frame
                        # 自动换行
                        tf.word_wrap = True
                        tfp = tf.paragraphs[0]
                        tfp.text = cluster_name + "部署" + total_count[0] + "个应用项目，包含" + \
                                 total_count[1] + "个服务，运行实例" + total_count[2] + \
                                 "个。CPU已分配" + str(cluster_allocated_cpu) + "/" + str(cluster_total_cpu) + "，占比" + \
                                 "{:.0%}".format(cpu_allocated_rate) + "，内存已分配" + \
                                 str(cluster_allocated_memory) + "/" + str(cluster_total_memory) + \
                                 "，占比" + "{:.0%}".format(memory_allocated_rate)
                        set_text_frame_paragraph_format(tfp)
                    elif len(table_data) > 40:
                        # table_1
                        # x: 左边距，y: 上边距, cx: 单元格宽度, cy: 单元格高度
                        x, y, cx, cy = Cm(7.5), Cm(2.2), Cm(1), Cm(0.2)
                        table = slide.shapes.add_table(40, 3, x, y, cx, cy).table
                        table.columns[0].width = Cm(3.5)
                        table.columns[1].width = Cm(4)
                        table.columns[2].width = Cm(1)
                        index = 0
                        cells = list(table.iter_cells())
                        count = len(table_data)
                        for i in range(count):
                            if index == len(cells):
                                break
                            for j in table_data.pop(0):
                                if index < len(cells):
                                    cells[index].text = str(j)
                                    set_cell_format(cells[index],
                                                    margin_left=0,
                                                    margin_right=0,
                                                    margin_top=0,
                                                    margin_bottom=0)
                                    set_text_frame_paragraph_format(cells[index].text_frame.paragraphs[0], font_size=7)
                                    index += 1

                        # table_2
                        # x: 左边距，y: 上边距, cx: 单元格宽度, cy: 单元格高度
                        # x: 左边距，y: 上边距, cx: 单元格宽度, cy: 单元格高度
                        x, y, cx, cy = Cm(16.5), Cm(2.2), Cm(1), Cm(0.2)
                        table = slide.shapes.add_table(len(table_data), 3, x, y, cx, cy).table
                        table.columns[0].width = Cm(3.5)
                        table.columns[1].width = Cm(4)
                        table.columns[2].width = Cm(1)
                        # 取消表格第一行特殊格式
                        table.first_row = False
                        index = 0
                        cells = list(table.iter_cells())
                        count = len(table_data)
                        for i in range(count):
                            if index == len(cells):
                                break
                            for j in table_data.pop(0):
                                if index < len(cells):
                                    cells[index].text = str(j)
                                    set_cell_format(cells[index],
                                                    margin_left=0,
                                                    margin_right=0,
                                                    margin_top=0,
                                                    margin_bottom=0)
                                    set_text_frame_paragraph_format(cells[index].text_frame.paragraphs[0], font_size=7)
                                    index += 1
                        # 右下插入文本框
                        txbox = slide.shapes.add_textbox(Cm(16.5), Cm(12.5), Cm(8), Cm(1))
                        # 文本框应该有一个默认段落，直接获取，额外添加段落需要使用text_frame.add_paragraph()，清除段落用text_frame.clear()
                        tf = txbox.text_frame
                        # 自动换行
                        tf.word_wrap = True
                        tfp = tf.paragraphs[0]
                        tfp.text = cluster_name + "部署" + total_count[0] + "个应用项目，包含" + \
                                 total_count[1] + "个服务，运行实例" + total_count[2] + \
                                 "个。CPU已分配" + str(cluster_allocated_cpu) + "/" + str(cluster_total_cpu) + "，占比" + \
                                 "{:.0%}".format(cpu_allocated_rate) + "，内存已分配" + \
                                 str(cluster_allocated_memory) + "/" + str(cluster_total_memory) + \
                                 "，占比" + "{:.0%}".format(memory_allocated_rate)
                        set_text_frame_paragraph_format(tfp)
                if page == 1 and table_pg_num == 2:
                    # table_1
                    # x: 左边距，y: 上边距, cx: 单元格宽度, cy: 单元格高度
                    x, y, cx, cy = Cm(7.5), Cm(2.2), Cm(1), Cm(0.2)
                    table = slide.shapes.add_table(40, 3, x, y, cx, cy).table
                    table.columns[0].width = Cm(3.5)
                    table.columns[1].width = Cm(4)
                    table.columns[2].width = Cm(1)
                    index = 0
                    cells = list(table.iter_cells())
                    count = len(table_data)
                    for i in range(count):
                        if index == len(cells):
                            break
                        for j in table_data.pop(0):
                            if index < len(cells):
                                cells[index].text = str(j)
                                set_cell_format(cells[index],
                                                margin_left=0,
                                                margin_right=0,
                                                margin_top=0,
                                                margin_bottom=0)
                                set_text_frame_paragraph_format(cells[index].text_frame.paragraphs[0], font_size=7)
                                index += 1
                    # table_2
                    # x: 左边距，y: 上边距, cx: 单元格宽度, cy: 单元格高度
                    # x: 左边距，y: 上边距, cx: 单元格宽度, cy: 单元格高度
                    x, y, cx, cy = Cm(16.5), Cm(2.2), Cm(1), Cm(0.2)
                    table = slide.shapes.add_table(40, 3, x, y, cx, cy).table
                    table.columns[0].width = Cm(3.5)
                    table.columns[1].width = Cm(4)
                    table.columns[2].width = Cm(1)
                    # 取消表格第一行特殊格式
                    table.first_row = False
                    index = 0
                    cells = list(table.iter_cells())
                    count = len(table_data)
                    for i in range(count):
                        if index == len(cells):
                            break
                        for j in table_data.pop(0):
                            if index < len(cells):
                                cells[index].text = str(j)
                                set_cell_format(cells[index],
                                                margin_left=0,
                                                margin_right=0,
                                                margin_top=0,
                                                margin_bottom=0)
                                set_text_frame_paragraph_format(cells[index].text_frame.paragraphs[0], font_size=7)
                                index += 1
                if page == 2 and table_pg_num == 2:
                    if len(table_data) <= 40:
                        # table_1
                        # 根据数据长度设置表格的行列数（同时通过行高*行数，计算表格高度，并与模板对比，做表格当前页分割，或换页分割）！！！！！
                        # x: 左边距，y: 上边距, cx: 单元格宽度, cy: 单元格高度
                        x, y, cx, cy = Cm(7.5), Cm(2.2), Cm(1), Cm(0.2)
                        table = slide.shapes.add_table(len(table_data), 3, x, y, cx, cy).table
                        table.columns[0].width = Cm(3.5)
                        table.columns[1].width = Cm(4)
                        table.columns[2].width = Cm(1)
                        index = 0
                        cells = list(table.iter_cells())
                        count = len(table_data)
                        for i in range(count):
                            for j in table_data.pop(0):
                                if index < len(cells):
                                    cells[index].text = str(j)
                                    set_cell_format(cells[index],
                                                    margin_left=0,
                                                    margin_right=0,
                                                    margin_top=0,
                                                    margin_bottom=0)
                                    set_text_frame_paragraph_format(cells[index].text_frame.paragraphs[0], font_size=7)
                                    index += 1

                        # 右下插入文本框
                        txbox = slide.shapes.add_textbox(Cm(16.5), Cm(12.5), Cm(8), Cm(1))
                        # 文本框应该有一个默认段落，直接获取，额外添加段落需要使用text_frame.add_paragraph()，清除段落用text_frame.clear()
                        tf = txbox.text_frame
                        # 自动换行
                        tf.word_wrap = True
                        tfp = tf.paragraphs[0]
                        tfp.text = cluster_name + "部署" + total_count[0] + "个应用项目，包含" + \
                                 total_count[1] + "个服务，运行实例" + total_count[2] + \
                                 "个。CPU已分配" + str(cluster_allocated_cpu) + "/" + str(cluster_total_cpu) + "，占比" + \
                                 "{:.0%}".format(cpu_allocated_rate) + "，内存已分配" + \
                                 str(cluster_allocated_memory) + "/" + str(cluster_total_memory) + \
                                 "，占比" + "{:.0%}".format(memory_allocated_rate)
                        set_text_frame_paragraph_format(tfp)
                    elif len(table_data) > 40:
                        # table_1
                        # x: 左边距，y: 上边距, cx: 单元格宽度, cy: 单元格高度
                        x, y, cx, cy = Cm(7.5), Cm(2.2), Cm(1), Cm(0.2)
                        table = slide.shapes.add_table(40, 3, x, y, cx, cy).table
                        table.columns[0].width = Cm(3.5)
                        table.columns[1].width = Cm(4)
                        table.columns[2].width = Cm(1)
                        index = 0
                        cells = list(table.iter_cells())
                        count = len(table_data)
                        for i in range(count):
                            if index == len(cells):
                                break
                            for j in table_data.pop(0):
                                if index < len(cells):
                                    cells[index].text = str(j)
                                    set_cell_format(cells[index],
                                                    margin_left=0,
                                                    margin_right=0,
                                                    margin_top=0,
                                                    margin_bottom=0)
                                    set_text_frame_paragraph_format(cells[index].text_frame.paragraphs[0], font_size=7)
                                    index += 1

                        # table_2
                        # x: 左边距，y: 上边距, cx: 单元格宽度, cy: 单元格高度
                        # x: 左边距，y: 上边距, cx: 单元格宽度, cy: 单元格高度
                        x, y, cx, cy = Cm(16.5), Cm(2.2), Cm(1), Cm(0.2)
                        table = slide.shapes.add_table(len(table_data), 3, x, y, cx, cy).table
                        table.columns[0].width = Cm(3.5)
                        table.columns[1].width = Cm(4)
                        table.columns[2].width = Cm(1)
                        # 取消表格第一行特殊格式
                        table.first_row = False
                        index = 0
                        cells = list(table.iter_cells())
                        count = len(table_data)
                        for i in range(count):
                            if index == len(cells):
                                break
                            for j in table_data.pop(0):
                                if index < len(cells):
                                    cells[index].text = str(j)
                                    set_cell_format(cells[index],
                                                    margin_left=0,
                                                    margin_right=0,
                                                    margin_top=0,
                                                    margin_bottom=0)
                                    set_text_frame_paragraph_format(cells[index].text_frame.paragraphs[0], font_size=7)
                                    index += 1
                        # 右下插入文本框
                        txbox = slide.shapes.add_textbox(Cm(16.5), Cm(12.5), Cm(8), Cm(1))
                        # 文本框应该有一个默认段落，直接获取，额外添加段落需要使用text_frame.add_paragraph()，清除段落用text_frame.clear()
                        tf = txbox.text_frame
                        # 自动换行
                        tf.word_wrap = True
                        tfp = tf.paragraphs[0]
                        tfp.text = cluster_name + "部署" + total_count[0] + "个应用项目，包含" + \
                                 total_count[1] + "个服务，运行实例" + total_count[2] + \
                                 "个。CPU已分配" + str(cluster_allocated_cpu) + "/" + str(cluster_total_cpu) + "，占比" + \
                                 "{:.0%}".format(cpu_allocated_rate) + "，内存已分配" + \
                                 str(cluster_allocated_memory) + "/" + str(cluster_total_memory) + \
                                 "，占比" + "{:.0%}".format(memory_allocated_rate)
                        set_text_frame_paragraph_format(tfp)
                # 合并单元格(即this_cell.merge(other_cell))
                # cell.merge(other_cell)

    # 9. 下周工作计划
    def slide_9(self, content: list):
        slide = self._prs.slides[25]  # type: pptx.slide.Slide
        shape = slide.shapes[0]  # type: pptx.shapes.base.BaseShape
        for shape in slide.shapes:
            if shape.has_table:
                tb = shape.table  # type: pptx.table.Table
                for row_idx in range(1, len(tb.rows)):
                    for col_idx in range(len(tb.columns)):
                        tb.cell(row_idx, col_idx).text = content[row_idx - 1][col_idx]

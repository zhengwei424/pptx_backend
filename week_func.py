import pptx
import pptx.slide
import pptx.shapes.base
import pptx.table
import pptx.chart.chart
import pptx.text.text
from pptx import Presentation
from pptx.enum.dml import MSO_COLOR_TYPE
from pptx.dml.color import RGBColor
from pptx.util import Pt
from pptx.enum.lang import MSO_LANGUAGE_ID
from pptx.enum.dml import MSO_FILL_TYPE
from pptx.enum.chart import XL_CHART_TYPE
from pptx.chart.data import CategoryChartData
from pptx.util import Cm
from pptx.shapes.placeholder import SlidePlaceholder
from pptx.enum.chart import XL_LEGEND_POSITION, XL_CHART_TYPE, XL_DATA_LABEL_POSITION
from pptx.chart.data import ChartData


# presentation->slide->shapes->placeholder,graphfrm
#            |      +->placeholder
#            |->slide_master-slide_layout
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


class TableAttribute:
    def __init__(self, table):
        self.tb = table  # type: pptx.table.Table

    # 先写内容，后改样式，否则可能出现样式不生效问题，如字体之类的？？？
    def set_cell_font(self,
                      cell: pptx.table._Cell,
                      fontName="FangSong",
                      fontSize=8,
                      fontBold=False,
                      fontColor="000000",
                      cellbgColor=None):
        # 语言设置，NONE表示移除所有语言设置
        cell.text_frame.paragraphs[0].font.language_id = MSO_LANGUAGE_ID.NONE
        # 字体
        cell.text_frame.paragraphs[0].font.name = fontName
        # 字体大小
        cell.text_frame.paragraphs[0].font.size = Pt(int(fontSize))
        # 是否加粗
        cell.text_frame.paragraphs[0].font.bold = fontBold

        # 用RGB表示字体颜色（两种方式）
        # cell.text_frame.paragraphs[0].font.color.rgb = RGBColor.from_string(fontColor)
        # 设置前景色或背景色之前需要执行
        # cell.text_frame.paragraphs[0].font.fill.patterned()  # 图案填充（可以设置前景色和背景色）
        cell.text_frame.paragraphs[0].font.fill.solid()  # 纯色填充（只能设置前景色）
        # 字体前景色(就是字体颜色）
        cell.text_frame.paragraphs[0].font.fill.fore_color.rgb = RGBColor.from_string(fontColor)
        # 字体背景色（就是文字本身的背景，不是整个cell）
        # cell.text_frame.paragraphs[0].font.fill.back_color.rgb = RGBColor.from_string(cellbgColor)
        # 字体颜色透明度
        # cell.text_frame.paragraphs[0].font.color.brightness = -1  # 取值范围-1~1，暗->亮
        if cellbgColor:
            # 填充cell前景色
            # cell.fill.patterned()  # 图案填充（可以设置cell前景色和背景色）
            cell.fill.solid()  # 纯色填充（只能设置cell前景色，但是对cell中的text而言也是背景）
            # cell.fill.back_color.rgb = RGBColor.from_string(cellbgColor)
            cell.fill.fore_color.rgb = RGBColor.from_string(cellbgColor)
        else:
            # cell颜色无填充
            cell.fill.background()


def change_table_data():
    pass


def chart_render():
    pass


def change_chart_data():
    pass


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
        if len(weekly_inspect) != 17:
            print("weekly_inspect length is 17,Please check out.")
        slide = self._prs.slides[9]  # type:pptx.slide.Slide
        shape = slide.shapes[0]  # type: pptx.shapes.base.BaseShape
        for shape in slide.shapes:
            if shape.has_table:
                tb = shape.table  # type: pptx.table.Table
                test = TableAttribute(tb)
                tb.cell(3, 1).text = "11"  # 巡检次数
                test.set_cell_font(tb.cell(3, 1), fontBold=False, cellbgColor="CDC839", fontColor="3C6F6A")
                # test.set_cell_font(tb.cell(3, 1), fontBold=True, cellbgColor="", fontColor="3C6F6A")
                tb.cell(3, 2).text = "0"  # 异常次数
                tb.cell(3, 3).text = "6"  # 报告提交次数
                tb.cell(3, 5).text = "√"  # 周一上午
                tb.cell(4, 5).text = "√"  # 周一下午
                tb.cell(3, 7).text = "√"  # 周二上午
                tb.cell(4, 7).text = "√"  # 周二下午
                tb.cell(3, 9).text = "√"  # 周三上午
                tb.cell(4, 9).text = "√"  # 周三下午
                tb.cell(3, 11).text = "√"  # 周四上午
                tb.cell(4, 11).text = "√"  # 周四下午
                tb.cell(3, 13).text = "√"  # 周五上午
                tb.cell(4, 13).text = "√"  # 周五下午
                tb.cell(3, 15).text = "√"  # 周六上午
                tb.cell(4, 15).text = "√"  # 周六下午
                tb.cell(3, 17).text = "√"  # 周日上午
                tb.cell(4, 17).text = "√"  # 周日下午

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
    def slide_8(self):
        # slide = self._prs.slides[18]  # type:pptx.slide.Slide
        # shape = slide.shapes[0]  # type: pptx.shapes.base.BaseShape
        # for shape in slide.shapes:
        #     # if shape.has_table: # 获取现有shape的坐标及基本参数（宽度高度等）
        #     #     print("shape: {}, left: {}, top: {}, height: {}, width: {}".format(shape, shape.left/360000, shape.top/360000,
        #     #                                                                        shape.height/360000, shape.width/360000))
        #     if shape.has_chart:
        #         chart = shape.chart  # type: pptx.chart.chart.Chart
        #         print(chart.chart_title.text_frame.text)
        slide = self._prs.slides.add_slide(self._prs.slide_masters[0].slide_layouts[1])
        # 根据数据长度设置表格的行列数（同时通过行高*行数，计算表格高度，并与模板对比，做表格当前页分割，或换页分割）！！！！！
        x, y, cx, cy = Cm(7), Cm(2.2), Cm(5), Cm(2)
        # x: 左边距，y: 上边距, cx: 单元格宽度, cy: 单元格高度
        table = slide.shapes.add_table(3, 3, x, y, cx, cy).table
        # 访问cell
        table.cell(0, 0).text = "测试"
        # 合并单元格(即this_cell.merge(other_cell))
        # cell.merge(other_cell)
        slidePlaceholder = slide.shapes[0]  # type: pptx.shapes.placeholder.SlidePlaceholder
        slidePlaceholder.text = "hahahaha"

        # cpu 扇形图
        chart_cpu_data = ChartData()
        chart_cpu_data.categories = ['已分配', '未分配']
        chart_cpu_data.add_series('cpu', (0.25, 0.75))
        chart_cpu = slide.shapes.add_chart(XL_CHART_TYPE.PIE, Cm(0.1), Cm(2.2), Cm(6), Cm(6), chart_cpu_data).chart

        # 设置图例说明（会在图中标识已分配、未分配的颜色说明）
        chart_cpu.has_legend = True
        chart_cpu.legend.position = XL_LEGEND_POSITION.BOTTOM
        chart_cpu.legend.include_in_layout = False

        # 设置标题（不设置，默认值是add_series中的列标题"cpu"）
        chart_cpu.has_title = True
        chart_cpu.chart_title.has_text_frame = True
        chart_cpu.chart_title.text_frame.text = "外网微服务区CPU(核)"
        chart_cpu.chart_title.text_frame.fit_text(font_family='Microsoft YaHei', max_size='12', bold=False,
                                                  italic=False, font_file=None)

        chart_cpu.plots[0].has_data_labels = True
        chart_cpu_data_labels = chart_cpu.plots[0].data_labels
        chart_cpu_data_labels.number_format = '0%'
        chart_cpu_data_labels.position = XL_DATA_LABEL_POSITION.OUTSIDE_END

        print("chart_has_legend: ", chart_cpu.has_legend, chart_cpu.legend)
        print("chart_has_title: ", chart_cpu.has_title, chart_cpu.chart_title.has_text_frame)
        # memory 扇形图
        chart_mem_data = ChartData()
        chart_mem_data.categories = ['已分配', '未分配']
        chart_mem_data.add_series('mem', (0.25, 0.75))
        chart_mem = slide.shapes.add_chart(XL_CHART_TYPE.PIE, Cm(0.1), Cm(8.3), Cm(6), Cm(6), chart_mem_data).chart
        # 设置图例说明（会在图中标识已分配、未分配的颜色说明）
        chart_mem.has_legend = True
        chart_mem.legend.position = XL_LEGEND_POSITION.BOTTOM
        chart_mem.legend.include_in_layout = False

        chart_mem.plots[0].has_data_labels = True
        chart_mem_data_labels = chart_mem.plots[0].data_labels
        chart_mem_data_labels.number_format = '0%'
        chart_mem_data_labels.position = XL_DATA_LABEL_POSITION.OUTSIDE_END

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

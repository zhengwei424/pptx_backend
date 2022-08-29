import pptx
from pptx import Presentation
from pptx.presentation import Presentation as pptclass
from pptx.util import Cm, Emu
from week_func import PresentationBuilder, WeaklyReports

if __name__ == '__main__':
    # 这种打开方式适合ppt2007及最新，不适合ppt2003及以前。支持stringio/bytesio stream
    prs = Presentation("test.pptx")  # type: pptx.presentation.Presentation # 设置type，会有代码提示
    wr = WeaklyReports(prs)
    wr.slide_2()


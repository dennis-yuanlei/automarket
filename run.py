from analyze_excel import AutoMarket
from draw_echarts import DrawEcharts
from utils import get_dkem_share

if __name__ == "__main__":
    working_dir = './data/20250609'
    market = AutoMarket(working_dir)
    data = market.forward()
    market_echart = DrawEcharts(data)
    pie = market_echart.draw_pie(f'{working_dir}/res/市占率份额统计.html')
    get_dkem_share(data, f'{working_dir}/res/DKEM份额统计.xlsx')
import pandas as pd
import json
import os

from pyecharts import options as opts
from pyecharts.charts import Line

import ipdb

class CompareAnalyzer():
    def __init__(self, data_dir='./data'):
        self.dfs = []
        self.date_list = os.listdir(data_dir)
        if '.DS_Store' in self.date_list:
            self.date_list.remove('.DS_Store')
        self.date_list.sort()
        for date in self.date_list:
            excel_path = os.path.join(data_dir, date, 'res', 'DKEM市占率.xlsx')
            df = pd.read_excel(excel_path, index_col=0)
            self.dfs.append(df)

    def compare_last_now(self, thre=1, save_path='./data/20250616/'):
        # 对比本周相对于上周的变化
        data_diff = []
        print(f'\033[1;32m---- 正在对比{self.date_list[-1]}相对于{self.date_list[-2]}的百分比变化 ----\033[0m')
        df_now = self.dfs[-1]
        df_last = self.dfs[-2]
        not_in_df_now = [x for x in df_last.index if x not in df_now.index]
        not_in_df_last = [x for x in df_now.index if x not in df_last.index]
        print(f'❌ 本周客户{not_in_df_last} 不在上一周的客户列表中')
        print(f'❌ 上周客户{not_in_df_now} 不在本周的客户列表中')
        for cust in df_now.index:
            if cust not in df_last.index:
                continue
            zhengyin = float(df_now.loc[cust].loc['正面副栅占比'].strip('%'))-float(df_last.loc[cust].loc['正面副栅占比'].strip('%'))
            beiyin = float(df_now.loc[cust].loc['背面副栅占比'].strip('%'))-float(df_last.loc[cust].loc['背面副栅占比'].strip('%'))
            if zhengyin<thre and beiyin<thre:
                continue
            tmp = {'客户':cust, '正银':f'{round(zhengyin,2)}%', 
                                '背银':f'{round(beiyin,2)}%'}
            data_diff.append(tmp)

        df_diff = pd.DataFrame(data_diff)  
        df_diff.to_excel(f'{save_path}/res/DKEM市占率相对上周变动.xlsx', index=False)
        return df_diff  

    def compare_all_date(self, save_path='./data/20250616/'):
        # 绘制主客户的每周占比变动折线图
        self.custs = ['晶澳', '通威', '中润', '晶科', '爱旭', '正泰', '阿特斯', '捷泰', '英发', '天合', '隆基']
        colors = {
                "晶澳": "#F08080",
                "通威": "#FFD700",
                "中润": "#FFDEAD",
                "晶科": "#DA70D6",
                "爱旭": "#FF6347",
                "正泰": "#6495ED",
                "阿特斯": "#3CB371",
                "捷泰": "#FFA500",
                "英发": "#00FFFF",
                "天合": "#EE82EE",
                "隆基": "#40E0D0"
                }
        data = dict()
        for cust in self.custs:
            zhengyin = []
            beiyin = []
            for i in range(len(self.dfs)):
                df = self.dfs[i]
                zhengyin.append(float(df.loc[cust].loc['正面副栅占比'].strip('%')))
                beiyin.append(float(df.loc[cust].loc['背面副栅占比'].strip('%')))
            data[cust] = {'正面副栅占比':zhengyin, '背面副栅占比':beiyin}
        
        # draw line
        # 创建正面副栅占比折线图对象
        line = Line()
        # 添加 X 轴数据（时间轴）
        x_data = self.date_list
        # 添加多个主体的 Y 轴数据（示例数据）
        line.add_xaxis(x_data)
        for i in range(len(self.custs)):
            cust = self.custs[i]
            line.add_yaxis(
                series_name=cust,
                y_axis=data[cust]['正面副栅占比'],
                label_opts=opts.LabelOpts(is_show=False),
                linestyle_opts=opts.LineStyleOpts(width=2),
                color=colors[self.custs[i]]
            )
        # 设置全局配置项
        line.set_global_opts(
            title_opts=opts.TitleOpts(title="客户I正面副栅占比趋势",
                        pos_left="center",  # 主标题居中
                        title_textstyle_opts=opts.TextStyleOpts(font_size=48)),  # 字体放大
            xaxis_opts=opts.AxisOpts(name="date", type_="category"),
            yaxis_opts=opts.AxisOpts(name="percent", type_="value"),
            tooltip_opts=opts.TooltipOpts(trigger="axis"),
            toolbox_opts=opts.ToolboxOpts(is_show=True),
            legend_opts=opts.LegendOpts(pos_top="15%", pos_right="0%", orient='vertical')
        )
        # 渲染图表到 HTML 文件
        line.render(f"{save_path}/res/正面副栅占比趋势.html")

        #################################################
        # 创建背面面副栅占比折线图对象
        line = Line()
        # 添加 X 轴数据（时间轴）
        x_data = self.date_list
        # 添加多个主体的 Y 轴数据（示例数据）
        line.add_xaxis(x_data)
        for i in range(len(self.custs)):
            cust = self.custs[i]
            line.add_yaxis(
                series_name=cust,
                y_axis=data[cust]['背面副栅占比'],
                label_opts=opts.LabelOpts(is_show=False),
                linestyle_opts=opts.LineStyleOpts(width=2),
                color=colors[self.custs[i]]
            )
        # 设置全局配置项
        line.set_global_opts(
            title_opts=opts.TitleOpts(title="客户I背面副栅占比趋势",
                        pos_left="center",  # 主标题居中
                        title_textstyle_opts=opts.TextStyleOpts(font_size=48)),  # 字体放大
            xaxis_opts=opts.AxisOpts(name="date", type_="category"),
            yaxis_opts=opts.AxisOpts(name="percent", type_="value"),
            tooltip_opts=opts.TooltipOpts(trigger="axis"),
            toolbox_opts=opts.ToolboxOpts(is_show=True),
            legend_opts=opts.LegendOpts(pos_top="15%", pos_right="0%", orient='vertical')
        )
        # 渲染图表到 HTML 文件
        line.render(f"{save_path}/res/背面副栅占比趋势.html")
                
    
    

if __name__ == "__main__":
    # df_diff = compare_last_now()
    # df_diff.to_excel(f'./data/20250616/res/DKEM市占率相对上周变动.xlsx', index=False)
    test = CompareAnalyzer()
    test.compare_last_now()
    test.compare_all_date()
    
        


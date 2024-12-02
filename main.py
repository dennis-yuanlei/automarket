import os
import pandas as pd
from openpyxl import Workbook
from openpyxl.drawing.image import Image
import numpy as np
import colorsys
import matplotlib.pyplot as plt
import json
import random
import ipdb

random.seed(1234567)

class AutoMarket():
    def __init__(self, working_dir, dump_vis=True):
        self.working_dir = working_dir
        self.res_dir = f'{self.working_dir}/res'
        os.makedirs(self.res_dir, exist_ok=True)
        self.dump_vis = dump_vis
        self.header_prefix = ['客户', '基地', '负责人', '实际开线数', '日用量（KG/线)', '其他', '总数核对', '备注']
        self.products = ['正面副栅', '背面副栅', '正面主栅', '背面主栅']
        self.color_map = self.generate_random_colors(20)
        self.customers = []
        self.competitors = None

    def split_excel(self, df):
        '''
        把src_dataframe拆分为两个
        '''
        r,c = df.shape
        r_split = []
        i = 1
        while i<r:
            while i<r and (not pd.isna(df.iloc[i,1])):
                i += 1
            if i<r and pd.isna(df.iloc[i+1,1]):
                r_split.append(i-1)
            while i<r and pd.isna(df.iloc[i,1]):
                i+=1
            if i<r and (not pd.isna(df.iloc[i+1, 1])):
                r_split.append(i)
        # print(r_split)
        return df.iloc[0:r_split[0]+1], df.iloc[r_split[1]:r+1]

    def unmerge(self, df):
        '''
        填充因合并单元格导致的缺失label
        确保key唯一:
            将第一行和第二行的item合并为新item
            将客户 基地 负责人合并为新item
        '''
        # 填充表头(第一行)
        headers = df.columns.tolist()
        assert not pd.isna(headers[0]), '第一列的表头不能为空'
        i = 0
        while i<len(headers):
            while i<len(headers) and ('Unnamed:' not in headers[i]):
                i += 1
            while i<len(headers) and ('Unnamed:' in headers[i]):
                headers[i] = headers[i-1]
                i += 1
        df.columns = headers

        # 从第一行中提取competitors
        competitors = headers.copy()
        i = 0
        while i<len(competitors):
            if competitors[i] in self.header_prefix:
                competitors.pop(i)
            else:
                i += 1

        # 合并第一行和第二行为新索引
        self.company_product = headers
        product = df.iloc[0].tolist()
        assert len(product)==len(headers), '第一行与第二行列数不相等'
        for i in range(len(product)):
            if pd.isna(product[i]):
                continue
            self.company_product[i] = f'{headers[i]}_{product[i]}'
        # reset headers
        df.columns = self.company_product
        # 删除第二行
        df = df.iloc[1:]
        
        # self.competitors去重
        competitors = list(set(competitors))
        if not self.competitors:
            self.competitors = competitors
        assert self.competitors==competitors, 'competitor conflict'

        # 填充第一列
        customers = df['客户'].tolist()
        
        print(len(customers))
        i = 0
        while i<len(customers):
            while i<len(customers) and (not pd.isna(customers[i])):
                i += 1
            while i<len(customers) and i>0 and pd.isna(customers[i]):
                customers[i] = customers[i-1]
                i += 1
            i += 1
        df['客户'] = customers
        self.customers = list(set(customers+self.customers))

        # 重置第一列
        jidi = df['基地'].tolist()
        fuzeren = df['负责人'].tolist()
        assert len(jidi)==len(fuzeren), '基地和负责人数量对不齐'
        assert len(jidi)==len(customers), '基地和客户数量对不齐'
        idx_new = customers
        for i in range(len(idx_new)):
            idx_new[i] = f'{customers[i]}_{jidi[i]}_{fuzeren[i]}'
        df.columns = ['客户_基地_负责人']+df.columns.tolist()[1:]
        df['客户_基地_负责人'] = idx_new
        # 删除基地 负责人列
        del df['基地']
        del df['负责人']

        return df

    def analyze(self, df):
        # self.info 保存了各客户-各产品-各友商的线数
        self.cust_dfs = dict()
        self.info = dict()
        for cust in self.customers:
            self.info[cust] = dict()
            for prod in self.products:
                self.info[cust][prod] = dict()
                for compe in self.competitors:
                    company_product = f'{compe}_{prod}'
                    filtered_df = df[df.iloc[:, 0].astype(str).str.contains(cust, na=False)]
                    sum_company_product = filtered_df[company_product].sum()
                    self.info[cust][prod][compe] = sum_company_product
                if self.dump_vis:
                    self.draw_pie(self.info[cust][prod].keys(), self.info[cust][prod].values(), title=f'{cust}_{prod}')
        self.dump_analyze_results() 

        # 行业各prod友商share分析
        self.prod_share = dict()
        for prod in self.products:
            self.prod_share[prod] = dict()
            for compe in self.competitors:
                company_product = f'{compe}_{prod}'
                columns_with_comp_prod = [col for col in df.columns if company_product in col.lower()]
                filtered_df = df[columns_with_comp_prod]
                sum_company_product = filtered_df[company_product].sum()
                self.prod_share[prod][compe] = sum_company_product
            if self.dump_vis:
                self.draw_pie(self.info[cust][prod].keys(), self.info[cust][prod].values(), title=f'{prod}_大客户')
        df_prod_share = pd.DataFrame(self.prod_share)
        # ipdb.set_trace()
            


    def dump_analyze_results(self):
        # 将统计结果保存为json,方便后续回溯读取
        with open(f'{self.res_dir}/main.json', 'w', encoding='utf-8') as f:
            json.dump(self.info, f, ensure_ascii=False, indent=4) 

        if self.dump_vis:

            # dump友商在各个客户的占比情况并插入图片,输出最终用于汇报的excel
            with pd.ExcelWriter(f'{self.res_dir}/market_share_vis.xlsx', engine='xlsxwriter') as writer:
                start_row = 0
                for cust, cust_info in self.info.items():
                    cust_df = pd.DataFrame(cust_info)
                    cust_df.to_excel(writer, sheet_name='Sheet1', startrow=start_row+1, startcol=0, index=False)
                    
                    # 获取xlsxwriter的工作表对象
                    wb = writer.book
                    ws = writer.sheets['Sheet1']
                    # 在Excel文件中添加标题
                    ws.write(f'A{start_row+1}', cust)
                    # # 在每个dataframe的最后一列插入pie
                    # for i, prod in enumerate(self.products):
                    #     img_path = f'{self.working_dir}/res/pies/{cust}_{prod}.png'
                    #     img = Image(img_path)
                    #     max_column = len(cust_df.columns)+2
                    #     ws.add_image(img, f"{chr(65 + max_column)}{start_row + 2}")
                    [h,w] = cust_df.shape
                    start_row += (h+8)



    def generate_random_colors(self, num_colors):
        # 生成颜色的函数
        colors = []
        for _ in range(num_colors):
            # 生成随机的HSV值
            h = random.random()  # 色调
            s = random.uniform(0.5, 1)  # 饱和度，避免太低以确保颜色的鲜明度
            v = random.uniform(0.5, 1)  # 明度，避免太低以确保颜色的鲜明度
            # 将HSV转换为RGB
            r, g, b = colorsys.hsv_to_rgb(h, s, v)
            # 将RGB转换为十六进制颜色代码
            hex_color = '#{:02x}{:02x}{:02x}'.format(int(r * 255), int(g * 255), int(b * 255))
            colors.append(hex_color)
        return colors

    def draw_pie(self, labels, sizes, title):
        # 自定义autopct函数，如果百分比不为0，则显示
        def autopct_format(values):
            def my_autopct(pct):
                return ('%1.1f%%' % pct) if pct > 0 else ''
            return my_autopct
        
        # match color by competitor
        assert len(self.color_map)>=len(self.competitors), 'the number of competitor is more than colors'
        indexes = [self.competitors.index(label) for label in labels]
        colors = [self.color_map[i] for i in indexes]
        
        save_dir = f'{self.working_dir}/res/pies'
        os.makedirs(save_dir, exist_ok=True)
        plt.rcParams['font.sans-serif']=['SimHei'] #用来正常显示中文标签
        plt.rcParams['axes.unicode_minus'] = False #用来正常显示负号
        plt.clf()

        # 自定义labels，如果sizes为0，则不显示
        custom_labels = [label if size > 0 else '' for label, size in zip(labels, sizes)]
        plt.title(title)
        plt.pie(sizes, labels=custom_labels,
                colors=colors,
                autopct=autopct_format(sizes),  # 使用自定义函数
                )
        plt.axis('equal')
        plt.savefig(f'{save_dir}/{title}.png')

    def prepare(self, df):
        '''
        切分df为main与other
        填充第一行第一列的nan
        '''
        df_main, df_other = self.split_excel(df)
        df_main = self.unmerge(df_main)
        return df_main, df_other


    def forward(self):
        print('\033[1;32m---- prepare data ----\033[0m')
        # data prepare
        dfs = []
        for f in os.listdir(self.working_dir):
            if not f.endswith('.xlsx'):
                continue
            excel_path = os.path.join(self.working_dir, f)
            print(f'\033[1;32m---- process {excel_path} ----\033[0m')
            df = pd.read_excel(excel_path)
            self.excel_path = excel_path
            df_main, df_other = self.prepare(df)
            dfs.append(df_main)
        df_main = pd.concat(dfs, ignore_index=True)
        df_main.to_excel(f'{self.res_dir}/main.xlsx', index=False)
        print(self.competitors)
        print('\033[1;32m---- analyze data ----\033[0m')
        self.analyze(df_main)
        print('\033[1;32m---- finished ----\033[0m')







if __name__ == "__main__":
    working_dir = f'./data/20241125'
    market = AutoMarket(working_dir)
    market.forward()
            



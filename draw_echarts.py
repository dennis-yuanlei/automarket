import json
from pyecharts.charts import Pie, Grid, Page
from pyecharts.components import Table
from pyecharts import options as opts

# Function to load JSON data from a file path
def load_json(file_path):
    try:
        with open(file_path, 'r', encoding='utf-8') as file:
            data = json.load(file)
            if not isinstance(data, dict):
                raise ValueError("JSON data must be a dictionary")
            return data
    except FileNotFoundError:
        print(f"Error: File '{file_path}' does not exist.")
        exit(1)
    except json.JSONDecodeError:
        print(f"Error: File '{file_path}' is not a valid JSON format.")
        exit(1)
    except ValueError as e:
        print(f"Error: {e}")
        exit(1)

class DrawEcharts():
    def __init__(self, data):
        # self.data = load_json(file_path)
        self.data = data

    # Function to generate table
    def generate_table(self, data_cust):
        table = Table()
        rows = list()
        headers = ['厂商', '正面副栅', '背面副栅', '正面主栅', '背面主栅']
        
        supplies_dict = dict()
        for prod, supplies in data_cust.items():
            if prod == '实际开线数':
                xianshu_sum = int(supplies)
                continue
            supplies_dict.update(supplies)
        supplies_list = supplies_dict.keys()

        for ss in supplies_list:
            rows.append({'厂商':ss, '正面副栅':data_cust['正面副栅'][ss], 
                         '背面副栅':data_cust['背面副栅'][ss], '正面主栅':data_cust['正面主栅'][ss], 
                         '背面主栅':data_cust['背面主栅'][ss]})
        
        table.add(headers, rows)
            
    
    # Function to generate _pie charts with table
    def generate_pie_charts_with_table(self, data):
        page = Page(layout=Page.SimplePageLayout)
        for customer, products in data.items():
            grid = Grid(init_opts=opts.InitOpts(width="100%", height="400px"))
            position = 0
            for product, suppliers in products.items():
                if product=='实际开线数':
                    continue
                # Filter out suppliers with zero share
                supplier_data = [(k, v) for k, v in suppliers.items() if v > 0]
                if not supplier_data:
                    continue
                pie = (
                    Pie()
                    .add(
                        "",
                        supplier_data,
                        radius=["30%", "50%"],
                        center=[f"{position * 25 + 12.5}%", "50%"],
                        label_opts=opts.LabelOpts(is_show=True, position="outside")
                    )
                    .set_global_opts(
                        title_opts=opts.TitleOpts(title=f"{customer} - {product}", pos_left=f"{position * 25 + 8}%", pos_top="top"),
                        legend_opts=opts.LegendOpts(is_show=False)
                    )
                    .set_series_opts(
                        label_opts=opts.LabelOpts(formatter="{b}\n{d}%")
                    )
                )
                grid.add(pie, grid_opts=opts.GridOpts(pos_left=f"{position * 25}%", pos_right=f"{75 - position * 25}%"))
                position += 1
            if position > 0:
                page.add(grid)
            
            # # add table
            # grid = Grid(init_opts=opts.InitOpts(width="100%", height="400px"))
            # cust_table = self.generate_table(products)
            # page.add(grid)
        return page

    # Main function
    def draw_pie(self, output_html="pie_charts.html"):
        page = self.generate_pie_charts_with_table(self.data)
        page.render(output_html)
        print(f"✅ 已生成：Charts have been saved to {output_html}")

# Example usage
if __name__ == "__main__":
    file_path = "/Users/leiyuan/program/code/python/auto_market/data/20250602/res/main.json"
    dd = DrawEcharts(load_json(file_path))
    dd.draw_pie(file_path)
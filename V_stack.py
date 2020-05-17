# -*- coding: utf-8 -*-
# author: Kindle Hsieh time: 2020/5/16

from bokeh.plotting import figure, output_file, show

output_file("V_stack.html", title='stack')
fruits = ['Apples', 'Pears', 'Nectarines', 'Plums', 'Grapes', 'Strawberries']
years = ["2015", "2016", "2017"]
colors = ["#c9d9d3", "#718dbf", "#e84d60"]

data = {'fruits' : fruits,
        '2015'   : [2, 1, 4, 3, 2, 4],
        '2016'   : [5, 3, 4, 2, 4, 6],
        '2017'   : [3, 2, 4, 4, 5, 3]}

p = figure(x_range=fruits, plot_height=250, title="Fruit Counts by Year", toolbar_location='right')
p.vbar_stack(years, x='fruits', width=0.9, color=colors, source=data, legend_label=years)

p.y_range.start=0
p.x_range.range_padding=0.2
p.xgrid.grid_line_color = None
p.axis.minor_tick_line_color = None
p.outline_line_color = None

p.legend.location='top_left'
p.legend.orientation='horizontal'
p.legend.click_policy= "hide"
show(p)
# -*- coding: utf-8 -*-
# author: Kindle Hsieh time: 2020/5/17

from bokeh.plotting import figure, output_file, show
from bokeh.models import Legend
from bokeh.palettes import brewer


months = ['JAN', 'FEB', 'MAR']
data = {"month" : months,
        "cat1"  : [1, 4, 12],
        "cat2"  : [2, 5, 3],
        "cat3"  : [5, 6, 1],
        "cat4"  : [8, 2, 1],
        "cat5"  : [1, 1, 3]}
categories = list(data.keys())[1:]  # slice 除了第一筆資料。
colors = brewer['YlGnBu'][len(categories)]

p = figure(x_range=months, plot_height=250, title="Categories by month", toolbar_location="right")
v = p.vbar_stack(categories, x='month', width=0.9, color=colors, source=data
                 # , legend_label=categories
                 )
# for add annotation-legend outside of the figure.
# version one.
# legend = Legend(items=[
#     ("cat1",   [v[0]]),
#     ("cat2",   [v[1]]),
#     ("cat3",   [v[2]])],
#     location=(0, 0))
# version two.
legend = Legend(items=[(c, [v[i]]) for i, c in enumerate(categories)], location=(0, 0))
p.add_layout(legend, 'right')

# # Place legend inner.
# p.legend.location = 'top_center'
# p.legend.orientation = 'horizontal'

# let legend can interactive.
p.legend.click_policy = 'hide'
show(p)
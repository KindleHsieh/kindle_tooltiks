from bokeh.plotting import figure, output_file, show
from bokeh.models import ColumnDataSource
from bokeh.palettes import GnBu3, OrRd3, Spectral6
from bokeh.transform import factor_cmap

output_file("colormapped_bars.html")

fruits = ['Apples', 'Pears', 'Nectarines', 'Plums', 'Grapes', 'Strawberries']
counts = [5, 3, 4, 2, 4, 6]

source = ColumnDataSource(data=dict(fruits=fruits, counts=counts))

p = figure(x_range=fruits, plot_height=250, toolbar_location=None, title="Fruit Counts")
# p.vbar(x='fruits', top='counts', width=0.9, source=source, legend_field="fruits",
#        line_color='white', fill_color=factor_cmap('fruits', palette=Spectral6, factors=fruits))
for i in range(len(fruits)):
    p.vbar(x=[fruits[i]], top=counts[i], width=0.9, legend_label=fruits[i], line_color="white")
# p.vbar(x=[fruits[0]], top=counts[0], width=0.9, legend_label=fruits[0], line_color="white")

p.xgrid.grid_line_color = None
p.y_range.start = 0
p.y_range.end = 9
p.legend.orientation = "horizontal"
p.legend.location = "top_center"

p.legend.click_policy='hide'
show(p)
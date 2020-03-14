import plotly.graph_objects as go
import xlrd

workbook = xlrd.open_workbook("1672920091153934971.xls")
worksheet = workbook.sheet_by_index(0)

fig = go.Figure()

yearAxis = []
priceData = []

# Year   Row:4 Column: D - HA
# Ay    Row:5 Column: D - HA
# Range (2,5): Generates 2,3,4.

for x in range (3,209):
    year = worksheet.cell(3,x).value
    month = worksheet.cell(4,x).value
    yearAxis.append(str(month) + str(year))

# Datas are from Row 7 to Row 424 
for x in range (6,423):
    priceData.clear()

    for xx in range (3,209):
        val = worksheet.cell(x,xx).value
        if val == '':
            val = 0

        if xx < 27:
            priceData.append(val/1000000)  # Before transition from TL to YTL
        else:
            priceData.append(val)
    
    itemName = worksheet.cell(x,1).value

    fig.add_trace(go.Scatter(y=priceData,x=yearAxis,visible=True,name=itemName))


fig.write_html('TurkeyInflationData.html', auto_open=True)

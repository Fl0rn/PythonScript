import plotly
import pandas as pd
import plotly.express as px


def ConvToSmall(excel_file):
    df = pd.read_excel(excel_file, sheet_name='Test Results',  engine='openpyxl')
    last_date = df['Date'].max()
# Calculate the date three months ago from the last date
    three_months_ago = last_date - pd.DateOffset(months=3)

# Filter data for the last three months
    last_three_months_data = df[df['Date'] >= three_months_ago].sort_values(by='Passed', ascending=False)
    last_three_months_data.reset_index(drop=True, inplace=True)

    last_three_months_data['Month'] = last_three_months_data['Date'].dt.month

# Group by 'Month' and find the day with the maximum 'Passed' count for each month
    result = last_three_months_data.groupby('Month').apply(lambda x: x.loc[x['Passed'].idxmax()])

# Print the result
    result['Date'] = result['Date'].dt.date
    with pd.ExcelWriter(excel_file, engine='openpyxl', mode='a') as writer:
        # Write df2 to a new sheet named "Sheet2" (you can change the sheet name if needed)
        result.to_excel(writer, sheet_name='Test Conclusion', index=False)

    fig = px.bar(result, x="Month", y=["Passed", "Failed", "Missing"], title="Status teste")
# Define colors for each trace (Passed=green, Failed=red, Missing=yellow)
    colors = {'Passed': 'green', 'Failed': 'red', 'Missing': 'yellow'}

# Update the colors of each trace
    for trace_name, color in colors.items():
        fig.update_traces(marker=dict(color=color), selector=dict(name=trace_name))

    plotly.offline.plot(fig, filename="TestReport.html")



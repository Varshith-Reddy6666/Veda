
import pandas as pd
from flask import Flask, render_template, request, send_file
import io

df = pd.read_excel('Subnat.xlsx')
data = df.copy()
app = Flask(__name__)

grouped = None



@app.route('/', methods=['GET', 'POST'])
def upload():
    global grouped
    unique_values = {}  # To store unique values for dropdowns
    sample_data = None  # To display a sample of the filtered data
    selected_columns = []  # Columns selected by the user
    selected_metrics = []  # Metrics (numerical columns) selected by the user
    applied_filters = {}  # Filters applied by the user
    agge_columns = []
    selected_columns = ['Region', 'District', 'Territory', 'Target Tier', 'Speciality', 'Field Reach']
    columns = ['Region', 'District', 'Territory']
  
    # percentage = ['Digital Reach', 'Sample to NBRx Ratio', 'Early Copay Enrollment Rate',
    #               'Omnichannel Reach', 'Calls to Targets', 'Targets Received Sample', 'Targets Attending Speaker Program',
    #               'Calls with Lunch / Total Call', 'Target Calls with Lunch', 'Otezla NPT Share NT+Tyk2 (LAAD)',
    #               'Targets Received RTE', 'RTE Open Rate']

    # Get unique values for filtering (categorical columns)
    metrics = data.select_dtypes(include=['number']).columns.tolist()
    for col in data.select_dtypes(include=['object', 'category']).columns:
        unique_values[col] = data[col].dropna().unique().tolist()

    for key in columns:
        del unique_values[key]
    
    # Handle form submission
    if request.method == 'POST':
        # Get selected columns
        if request.form.getlist('aggregation_column'):
            agge_columns = request.form.getlist('aggregation_column')
        
        # Get selected metrics (numerical columns)
        if request.form.getlist('numbercol'):
            selected_metrics = request.form.getlist('numbercol')

        if 'Field Reach' in selected_metrics:
            selected_metrics.remove('Field Reach')
        
        # Get applied filters
        for col in unique_values.keys():
            if request.form.getlist(col):
                applied_filters[col] = request.form.getlist(col)

        # Create a copy of the dataset for filtering
        filtered_data = data.copy()

        # Apply column and metric selection
        # if selected_columns or selected_metrics:
        #     filtered_columns = selected_columns + selected_metrics
        #     filtered_data = filtered_data[filtered_columns]

        # Apply row filters
        for col, values in applied_filters.items():
            if col in filtered_data.columns:
                filtered_data = filtered_data[filtered_data[col].isin(values)]
        
        # Display a sample of the filtered data
       

            grouped = filtered_data.groupby(agge_columns).agg('sum').reset_index()
        
            grouped["45 Days NBRx Dispense Rate (SP)"]=(grouped['45 Days NBRx Referals']*100/grouped["Otezla New Patient Referrals (SP)"]).round(2).astype(str)+"%"

            grouped["Otezla NBRx Share NT+Tyk2 (XPO)"]=(grouped['Otezla NBRX (XPO)']*100/grouped["NBRx (XPO)"]).round(2).astype(str)+'%'

            grouped['Sample to NBRx Ratio']=grouped['Samples']/grouped["NBRx (XPO)"]
         
            grouped=grouped[agge_columns+selected_metrics].reset_index(inplace=False)
            

        sample_data = grouped.head(100)


    return render_template(
        'upload.html',
        unique_values=unique_values,
        selected_columns=selected_columns,
        applied_filters=applied_filters,
        metrics=metrics,
        agge_columns=agge_columns,
        
        sample_data=sample_data.to_html(classes='table table-striped', index=False) if sample_data is not None else None
    )

@app.route('/download', methods=['POST'])
def download_data():
    # Assuming 'filtered_data' is a DataFrame containing the filtered data
    filtered_data = grouped.copy()  # Replace this with your actual filtered data logic

    # Write DataFrame to an Excel file in memory
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        filtered_data.to_excel(writer, index=False, sheet_name='FilteredData')
    output.seek(0)  # Move to the beginning of the BytesIO buffer

    # Send the file as an Excel attachment
    return send_file(
        output,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True,
        download_name='filtered_data.xlsx'
    )

if __name__ == '__main__':
    app.run(debug=True)
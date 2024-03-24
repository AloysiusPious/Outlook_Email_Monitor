import win32com.client
import win32com.client as win32
import csv
import os
import re
from datetime import datetime
import pandas as pd
import matplotlib.pyplot as plt
import configparser
def clean_up():
    # Predefined list of directories to delete
    directories_to_delete = ['Charts', 'Reports', 'Critical_Alerts', 'tmp', 'Incidents']

    # Delete directories and their contents
    for directory in directories_to_delete:
        try:
            # Remove files in the directory
            for root, dirs, files in os.walk(directory):
                for file in files:
                    os.remove(os.path.join(root, file))
            # Remove the directory itself
            os.rmdir(directory)
            print(f'Cleaned Up directory and its contents: {directory}')
        except FileNotFoundError:
            print(f'Directory not found: {directory}')
        except PermissionError:
            print(f'Permission denied for directory: {directory}')
        except Exception as e:
            print(f'Error Cleaning Up directory: {directory}, {e}')
def get_the_incident_in_table():
    html_body = """
        <style>
            table {
                border-collapse: collapse;
                width: auto;
            }
            th, td {
                border: 1px solid #dddddd;
                text-align: left;
                padding: 8px;
            }
            th {
                background-color: #f2f2f2;
            }
        </style>

        <table>
            <tr>
                <th style="text-align: center;">No Of Incidents</th>
                <th style="text-align: center;"><font color="red"><b>Incidents Alerts</b></font></th>
            </tr>
    """
    # Read all CSV files from the directory and add them as rows in the HTML table
    for filename in os.listdir(incident_directory):
        #if filename.endswith('.csv'):
        file_path = os.path.join(incident_directory, filename)
        # Read the CSV file into a DataFrame
        df = pd.read_csv(file_path)
        # Convert the DataFrame to an HTML table row
        html_table = df.to_html(index=False)
        # Add the filename and HTML table row to the email body table
        html_body += f"""
        <tr>
            <td style="height: auto; padding: 0; margin: 0;">{filename}</td>
            <td style="height: auto; padding: 0; margin: 0;">{html_table}</td>
        </tr>
        """
    # Close the HTML table and complete the email body
    html_body += """
        </table>
    """
    if len(os.listdir(incident_directory)) > 0:
        return html_body
    else:
        return ""

def send_email(mail_to):
    # Create an instance of the Outlook application
    outlook = win32.Dispatch('outlook.application')
    # Create a new email message
    message = outlook.CreateItem(0)
    # Add recipients, subject, and other email fields
    message.Subject = f"{mail_folder}_{start_date.date()}_To_{end_date.date()}"
    message.To = mail_to
    # message.Body = 'This is the body of the email.'
    files_list = os.listdir(incident_directory)
    # Calculate the number of files in the directory
    num_files = len(files_list)
    # Specify the path to the subdirectory
    subdirectory_path = os.path.join(os.path.dirname(__file__), 'Charts')
    # Read all CSV files from the directory and add them as tables to the email body
    # Verify if the subdirectory path exists
    if not os.path.exists(subdirectory_path):
        print(f"Subdirectory path '{subdirectory_path}' does not exist.")
        # Handle the error or exit gracefully
    else:
        # Get a list of all files in the subdirectory_path
        all_files = os.listdir(subdirectory_path)

        # Loop through all files and attach them to the email
        for filename in all_files:
            file_path = os.path.join(subdirectory_path, filename)
            if os.path.isfile(file_path):
                # Attach the file to the email
                attachment = message.Attachments.Add(file_path)
                # Embed the attachment in the email body
                image_cid = 'image_' + filename  # Create a unique Content-ID for each image
                # Define the HTML string with the image resized to half of its original size
                html_content = f'''
                <img src="cid:{image_cid}" alt="Embedded Image" width="50%" height="50%">
                '''
                attachment.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001E",
                                                        image_cid)
                message.Body += html_content
        # Insert the email content at the beginning of the email body
        file_count = 'No' if num_files < 1 else num_files
        incident_table = get_the_incident_in_table()
        message.HTMLBody = f"<center><Font face='Verdana' size='+1' color='red'>{report_duration}</font></center><br>Dear ITOps Team,<br><br><Font face='Verdana' size='+1' color='blue'>{file_count}</font> Incidents were found from <b>{start_date.date()} to {end_date.date()}</b> for alerts received from <font color='blue'> monitoring@detasad.com.sa</font>. <br><br> Please find attached comprehensive charts and detailed files providing insights into the Alerts identified.<br><br><br>{incident_table}<br><hr>{message.Body}<br><br><br><hr>{email_note}"
    # Specify the paths to the 'dir1' and 'dir2' directories
    directories = ['Critical_Alerts','Reports']
    # Loop through each directory to attach files as attachments
    for directory in directories:
        directory_path = os.path.join(os.path.dirname(__file__), directory)
        if not os.path.exists(directory_path):
            print(f"Directory path '{directory_path}' does not exist.")
        else:
            # Loop through files in the directory and attach them to the email
            for filename in os.listdir(directory_path):
                file_path = os.path.join(directory_path, filename)
                message.Attachments.Add(file_path)

    # Send the email
    message.Send()

def create_incident():
    for csv_file in os.listdir(tmp_directory):
        df = pd.read_csv(tmp_directory + "/"+csv_file)
        # Combine 'Date' and 'Time' columns into a single datetime column
        df['DateTime'] = pd.to_datetime(df['Date'] + ' ' + df['Time'])
        df = df.drop(columns=['Date'])
        df = df.drop(columns=['Time'])
        # Group the data by 'Target', 'Category', and 30-minute intervals
        groups = df.groupby(['Equipment', 'Category', pd.Grouper(key='DateTime', freq=frequency_window)])
        # Initialize a counter to track repeated incidents
        incident_counter = {}
        # Iterate over the groups
        for name, group in groups:
            equipment, category, datetime = name
            # Check if the group has more than 5 entries
            if len(group) >= incident_threshold:
                # Increment the incident counter for the specific target and category
                incident_counter.setdefault((equipment, category), 0)
                incident_counter[(equipment, category)] += 1
                # Define the output filename
                output_filename = f"{incident_directory}/{equipment}_{datetime.strftime('%Y-%m-%d')}_{category}_Incident_{incident_counter[(equipment, category)]}"
                # Write the group to a CSV file
                group.to_csv(output_filename, index=False)
def fileter_critical_alerts_based_on_hostname():
    df = pd.read_csv(critical_alerts_path)
    # Get unique hostnames
    hostnames = df['Equipment'].unique()
    # Filter based on hostname and save to individual CSV files
    for hostname in hostnames:
        # Filter the data for the current hostname
        filtered_data = df[df['Equipment'] == hostname]
        # Generate the CSV filename
        filename = f'{tmp_directory}/{hostname}_filtered_data.csv'
        # Save the filtered data to CSV
        filtered_data.to_csv(filename, index=False)
        #print(f"Filtered data for {hostname} saved to {filename}")
# Define the categorization function
def categorize(Description):
    # Define keywords and their corresponding categories
    keywords_categories = {
        'CPU': 'CPU',
        'Memory': 'Memory',
        'down': 'Node Down',
        'packet': 'Packet Loss',
        # Add more keywords and categories as needed
    }
    # Function to search for keywords in the 'Because' column and return the corresponding category
    for keyword, category in keywords_categories.items():
        if keyword in Description:
            return category
    return 'Other'  # If no keyword is found, categorize as 'Other'
def re_process_csv():
    # Read the CSV file
    df = pd.read_csv(csv_file_path)

    # Check if 'Category' column already exists, if not, add it after the 'Because' column
    if 'Category' not in df.columns:
        because_index = df.columns.get_loc('Description')
        df.insert(because_index + 1, 'Category', '')

    # Apply the categorize function to the 'Because' column and fill the 'Category' column
    df['Category'] = df['Description'].apply(categorize)
    df.to_csv(csv_file_path, index=False)
    # Filter rows where "Alert Type" is equal to "Critical"
    df = df[df['Alarm Severity'] == 'Critical']
    # Save the filtered DataFrame to another sheet or file
    df.to_csv(critical_alerts_path, index=False)
def sort_csv_with_timestamp():
    # Read the CSV file
    df = pd.read_csv(csv_file_path)
    # Drop duplicate rows if any
    df = df.drop_duplicates()
    # Convert the 'Date' and 'Time' columns to datetime
    df['Date'] = pd.to_datetime(df['Date'])
    df['Time'] = pd.to_datetime(df['Time'], format='%H:%M:%S').dt.time
    # Combine 'Date' and 'Time' columns into a single datetime column
    df['DateTime'] = pd.to_datetime(df['Date'].astype(str) + ' ' + df['Time'].astype(str))
    # Sort the DataFrame by the new datetime column
    df = df.sort_values(by='DateTime')
    # Drop the 'DateTime' column
    df = df.drop(columns=['DateTime'])
    df = df.drop(columns=['Incident Date'])
    # Save the sorted DataFrame to a new CSV file
    df.to_csv(csv_file_path, index=False)
def read_outlook_write_to_csv():
    # Specify fields to extract
    fields_to_extract = ["S.No", "Equipment", "Alarm Severity", "Monitoring Point", "Incident Date", "Description"]
    # Create a report
    report = []
    # Iterate through emails
    for idx, email in enumerate(emails, start=1):
        if email.Subject:  # Check if the email has a subject
            email_data = {"S.No": idx}  # Add S.No
            for field in fields_to_extract[1:]:  # Exclude S.No
                # Check if the field exists in the email
                if field in email.Body:
                    # Extract the value of the field
                    value_start = email.Body.find(field) + len(field) + 1
                    value_end = email.Body.find("\n", value_start)
                    value = email.Body[value_start:value_end].strip()
                    # Extracting Date and Time
                    if field == "Incident Date":
                        date_time = datetime.strptime(value, "%A, %B %d, %Y %I:%M %p")
                        email_data["Date"] = date_time.strftime("%d-%B-%Y")
                        email_data["Time"] = date_time.strftime("%H:%M:%S")
                    else:
                        email_data[field] = value
            # Add the email data to the report
            report.append(email_data)

    # Write data to CSV file
    with open(csv_file_path, mode="w", newline="") as file:
        writer = csv.DictWriter(file, fieldnames=fields_to_extract + ["Date", "Time"])
        writer.writeheader()
        for entry in report:
            writer.writerow(entry)
    # Read the CSV file into a DataFrame
    df = pd.read_csv(csv_file_path)
    df.to_csv(csv_file_path, index=False)
    # Convert 'Date' to datetime format
    df['Date'] = pd.to_datetime(df['Date'], format='%d-%B-%Y')
    # Filter the DataFrame based on the date range
    filtered_df = df[(df['Date'] >= start_date) & (df['Date'] <= end_date)]
    # Rewrite the filtered data to the same CSV file
    filtered_df = filtered_df.drop(columns=['S.No'])
    filtered_df.to_csv(csv_file_path, index=False)
    #print(f"CSV file with today's date created: {csv_file_path}")
def create_pie_chart():
   # Read the CSV file
    data = pd.read_csv(csv_file_path)

    # Convert 'Date' column to datetime format
    data['Date'] = pd.to_datetime(data['Date'])

    # Filter data based on start and end dates
    if start_date and end_date:
        data = data[(data['Date'] >= start_date) & (data['Date'] <= end_date)]

    # Group data by alert type
    grouped_data = data.groupby('Alarm Severity').size()

    # Define colors for each alert type
    colors = {'Normal': 'green', 'Warning': 'orange', 'Critical': 'red'}

    # Plotting the pie chart
    fig, ax = plt.subplots()
    grouped_data.plot(kind='pie', autopct='%1.1f%%',
                      colors=[colors.get(alert_type, 'gray') for alert_type in grouped_data.index], ax=ax)

    # Add title
    ax.set_title(mail_folder+' ('+str(start_date.date())+'_To_'+str(end_date.date())+')')

    # Equal aspect ratio ensures that pie is drawn as a circle
    ax.set_aspect('equal')
    file_path = charts_directory + '/' + mail_folder + '_'+str(start_date.date())+'_To_'+str(end_date.date())+'_pie.png'
    # Save the chart
    fig.savefig(file_path)
    # Show the plot
    plt.tight_layout()
    #plt.show()
def create_bar_chart():
    # Read the CSV file
    data = pd.read_csv(csv_file_path)
    # Convert the 'Date' column to datetime format
    data['Date'] = pd.to_datetime(data['Date'])

    # Filter data based on start and end dates
    if start_date and end_date:
        data = data[(data['Date'] >= start_date) & (data['Date'] <= end_date)]

    # Group the data by Date and Alert Type and count occurrences
    grouped_data = data.groupby(['Date', 'Alarm Severity']).size().unstack(fill_value=0)

    # Plotting all data in one chart
    fig, ax = plt.subplots(figsize=(12, 8))

    # Define colors for each alert type
    colors = {'Normal': 'green', 'Warning': 'orange', 'Critical': 'red', 'Unknown': 'blue'}

    # Plot each alert type for each date
    bar_width = 0.2
    index = grouped_data.index
    dates = range(len(index))
    for i, alert_type in enumerate(grouped_data.columns):
        ax.bar([x + i * bar_width for x in dates], grouped_data[alert_type], bar_width, label=alert_type,
               color=colors.get(alert_type, 'gray'))

    # Add labels and title
    ax.set_xlabel('Date')
    ax.set_ylabel('Count')
    ax.set_title(mail_folder+' ('+str(start_date.date())+'_To_'+str(end_date.date())+')')
    ax.set_xticks([i + bar_width for i in dates])
    ax.set_xticklabels(index.strftime('%Y-%m-%d'), rotation=45, ha='right')

    # Add legend
    ax.legend()
    if not os.path.exists(charts_directory):
        os.makedirs(charts_directory)
    # Combine the directory path and the filename
    today_date = datetime.now().strftime("%Y-%m-%d")
    file_path = charts_directory+'/' + mail_folder+'_'+str(start_date.date())+'_To_'+str(end_date.date())+'_bar.png'

    # Save the chart
    fig.savefig(file_path)
    # Show the plot
    plt.tight_layout()
    #plt.show()

def create_pie_chart_for_alaram_category():
   # Read the CSV file
    data = pd.read_csv(csv_file_path)
    # Convert 'Date' column to datetime format
    data['Date'] = pd.to_datetime(data['Date'])
    category_counts = data['Category'].value_counts()

    # Plotting the pie chart
    plt.figure(figsize=(8, 8))
    plt.pie(category_counts, labels=category_counts.index, autopct='%1.1f%%', startangle=140)
    plt.title('\n'+mail_folder +'_Category_From ('+str(start_date.date())+'_To_'+str(end_date.date())+')\n')
    plt.axis('equal')  # Equal aspect ratio ensures that pie is drawn as a circle.
    global category_chart_file
    category_chart_file = mail_folder + '_Category_From_'+str(start_date.date())+'_To_'+str(end_date.date())+'_pie.png'

    # Save the pie chart to a .png file
    plt.savefig(charts_directory + '/' + category_chart_file)
    # Show the plot
    plt.tight_layout()
    #plt.show()
def create_bar_chart_for_critical_count():
    # Read the CSV file
    df = pd.read_csv(critical_alerts_path)
    # Filter the DataFrame for 'Critical' alerts
    critical_alerts = df[df['Alarm Severity'] == 'Critical']

    # Group the DataFrame by 'Target' and count the occurrences of 'Critical' alerts for each target
    target_counts = critical_alerts['Equipment'].value_counts()

    # Plotting the bar chart
    plt.figure(figsize=(10, 6))
    target_counts.plot(kind='bar', color='red')
    plt.title(mail_folder+' : Number of Critical Alerts by HostName ('+ str(start_date.date()) + '_To_' + str(
        end_date.date())+')' )
    plt.xlabel('Equipment')
    plt.ylabel('Number of Critical Alerts')
    plt.xticks(rotation=45, ha='right')
    plt.tight_layout()
    global critical_chart_file
    critical_chart_file = mail_folder + '_Critical_Alerts_Count_' + str(start_date.date()) + '_To_' + str(
        end_date.date()) + '_pie.png'
    # Save the chart as a PNG file
    plt.savefig(charts_directory + '/' + critical_chart_file)

    # Show the plot
    #plt.show()
def create_node_down_chart():
    # Read the CSV file
    global node_chart_file
    df = pd.read_csv(critical_alerts_path)

    # Filter the dataframe where "Monitoring Point" is "Node is down"
    node_down_df = df[df['Monitoring Point'] == 'Node is down']
    if not node_down_df.empty:
        # Create a new CSV file with "Node is down" list
        node_down_df.to_csv(critical_directory+f"/Node_down_list_{start_date.date()}_To_{end_date.date()}.csv", index=False)

        # Create a chart for "Node is down"
        plt.figure(figsize=(10, 6))
        # Plot the equipment against the date and time
        plt.plot(node_down_df['Date'] + ' ' + node_down_df['Time'], node_down_df['Equipment'], marker='o', linestyle='-')
        plt.xlabel('Date and Time')
        plt.ylabel('Equipment')
        plt.title('Node Down Chart')
        plt.xticks(rotation=45)  # Rotate x-axis labels for better visibility
        plt.grid(True)
        plt.tight_layout()  # Adjust layout to prevent clipping of labels
        node_chart_file = mail_folder + '_NodeDown_Alerts_Count_' + str(start_date.date()) + '_To_' + str(end_date.date()) + '_pie.png'
        # Save the chart as a PNG file
        plt.savefig(charts_directory + '/' + node_chart_file)
        #plt.show()
    else:
        node_chart_file = "empty"
        print("No data available for 'Node is down'.")
def create_node_down_chart_1():
    # Read the CSV file
    df = pd.read_csv(critical_alerts_path)

    # Filter the dataframe where "Monitoring Point" is "Node is down"
    node_down_df = df[df['Monitoring Point'] == 'Node is down']

    # Create a new CSV file with "Node is down" list
    node_down_df.to_csv(critical_directory+f"/Node_down_list_{start_date.date()}_To_{end_date.date()}.csv", index=False)

    # Check if there are any rows in the filtered dataframe
    if not node_down_df.empty:
        # Convert 'Date' and 'Time' columns to datetime format
        node_down_df['DateTime'] = pd.to_datetime(node_down_df['Date'] + ' ' + node_down_df['Time'])

        # Create a pivot table to count the number of node down events for each date and time
        node_down_pivot = node_down_df.pivot_table(index='Date', columns='Time', aggfunc='size', fill_value=0)

        # Create a heatmap
        plt.figure(figsize=(12, 8))
        plt.imshow(node_down_pivot, cmap='YlOrRd', aspect='auto', interpolation='nearest')
        plt.colorbar(label='Number of Node Down Events')
        plt.xlabel('Time')
        plt.ylabel('Date')
        plt.title('Node Down Events Over Time')
        plt.xticks(range(len(node_down_pivot.columns)), node_down_pivot.columns, rotation=45)
        plt.yticks(range(len(node_down_pivot.index)), node_down_pivot.index)
        plt.tight_layout()
        plt.show()
    else:
        print("No data available for 'Node is down'.")


##################################################################
# Create a ConfigParser object
config = configparser.ConfigParser()
# Read the cfg file
config.read('DetaSadMonitoring.cfg')
# Access sections and keys
start_d = config['form_to']['start_date']
end_d = config['form_to']['end_date']
mail_to = config['email']['mail_to']
##
start_date = datetime.strptime(start_d, '%Y-%m-%d')
end_date = datetime.strptime(end_d, '%Y-%m-%d')
report_duration = end_date - start_date
report_duration = int(report_duration.days)
#report_duration = "**Monthly Report**" if report_duration > 8 else "**Weekly Report**"
report_duration = "** " + start_date.strftime('%Y-%m-%d') + " -to- " + end_date.strftime('%Y-%m-%d') + " **"

####
incident_threshold = int(config['threshold']['incident_threshold'])
frequency_window = str(config['threshold']['frequency_window']) +str("min")
##################################################################
start_date = pd.to_datetime(start_d)  # Example start date
end_date = pd.to_datetime(end_d)    # Example end date

email_note=f"<b><u>Note:</u></b> <br> <font color='blue ' face = 'Verdana' size='-1'>Our monitoring system diligently scans alert emails with 100% Python automation, eliminating manual effort.</font><br> <font color='#3383FF' face = 'Comic sans MS' size='-1'>1. The Python code diligently monitors email alerts and automatically marks any repeated critical alerts occurring more than <font color='red'><b>{incident_threshold}</b></font> times in <font color='red'><b>{frequency_window}</b></font>  as Incidents. Should any incidents pertain to your team, please take appropriate action.<br><br> 2. Additionally, please review the chart illustrating the highest number of critical alerts received for individual hostnames. If necessary, take proactive steps based on the insights provided.</font>"

##################################################################
clean_up()
##################################################################
# Connect to Outlook
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
# Specify the folder name where emails are located
folder_name = "Detasad_Monitoring"
# Get the folder by name
folder = outlook.Folders.Item("pious.aloysius@.xyz.com").Folders.Item(folder_name)
# Get emails from the specified folder
emails = folder.Items
# Specify the directory to save the CSV file
reports_directory = "Reports"
charts_directory = "Charts"
critical_directory = "Critical_Alerts"
tmp_directory = "tmp"
incident_directory = "Incidents"
os.makedirs(reports_directory, exist_ok=True)
os.makedirs(charts_directory, exist_ok=True)
os.makedirs(critical_directory, exist_ok=True)
os.makedirs(tmp_directory, exist_ok=True)
os.makedirs(incident_directory, exist_ok=True)
# Create CSV file with today's date
mail_folder = "Detasad_Monitoring"
today_date = datetime.now().strftime("%Y-%m-%d")
csv_file_path = os.path.join(reports_directory, f"{mail_folder}_From_{start_date.date()}_To_{end_date.date()}.csv")
critical_alerts_path = os.path.join(critical_directory, f"{mail_folder}_Critical_{start_date.date()}_To_{end_date.date()}.csv")
# Specify the folder name where emails are located

#############################################
#############################################
read_outlook_write_to_csv()
alert_count_df = pd.read_csv(csv_file_path)
if not alert_count_df.empty:
    sort_csv_with_timestamp()
    re_process_csv()# Reprocess to add Category Field(CPU/Memory) and filter 'Critical'
    ###
    create_bar_chart_for_critical_count()
    create_pie_chart_for_alaram_category()
    #create_bar_chart()
    #create_pie_chart()
    ###
    fileter_critical_alerts_based_on_hostname()
    create_node_down_chart()
    #create_node_down_chart_1()
    create_incident()
    send_email(mail_to)

#############################################
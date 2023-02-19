import pandas as pd
import smtplib, ssl
import config


# Read in the Spreadsheet as Pandas
df = pd.read_excel('AutomationAssignment.xlsx')

# Filter for all Active Orders

activeOrders = df.loc[df['Status'] == 'Active']

# Will use the DataFrame to find a few important KPIs
# Top 5 Longest Active Orders
# Total Number of Active Orders
# Number of On Hold Orders
# Customer with Longest Total Incident Duration
# Number of New Active Orders in last hour
# Number of Active Orders By Team
# Number of Incidents by Priority 
# Number of Incidents by Incident Type

# Find the number of Active Orders

numActive = len(activeOrders.index)

# Find 5 Longest Active Orders

longestActive = activeOrders.sort_values(by='IncidentDurationInHours', ascending=False).head(5)
longestActive = longestActive['Incident ID'].reset_index(drop=True)

# Find the On Hold Orders

onHold = df.loc[df['Status'] == 'On Hold']

# Find the number of On Hold orders

numOnHold = len(onHold.index)

# Find the Customer with longest Total Incident Duration

customerInfo = activeOrders.groupby('Customer Account Name').sum('IncidentDurationInHours').round(2)
customerInfo.drop(columns = ['Incident ID', 'Priority', 'Time To Resolve'], inplace = True)
customerInfo = customerInfo.head(1)
customerInfo.reset_index(inplace=True)

# Find the number of new Active Orders

newActive = activeOrders.loc[activeOrders['IncidentDurationInHours'] <= 1]
numNewActive = len(newActive.index)

# Find the Number of Active Orders By Team

ordersByTeam = activeOrders.groupby('Owned By Team').count()
numOrdersByTeam = ordersByTeam['Incident ID']
numOrdersByTeam = numOrdersByTeam.reset_index()
numOrdersByTeam.values

# Find the Number of Incidents By Priority

activeOrders['Priority'].fillna(value=0, inplace=True)
priorityCount = activeOrders.groupby('Priority').count()
priorityCount.sort_values(by='Priority', inplace=True)

columns = ['Customer Account Name', 'Customer Contact Name', 'Call Source', 'Incident Type', 'Status', 'Created Date Time', 'MTTR Description', 'Time To Resolve', 'Owned By Team', 'Last Modified Date Time', 'IncidentDurationInHours']

priorityCount.drop(columns = columns, inplace=True)
priorityCount.rename(columns = {'Incident ID': 'Count of Priority'}, inplace=True)
priorityCount.reset_index(inplace=True)
priorityCount['Priority'] = priorityCount['Priority'].astype(int)


# Find the Number of Incidents by Incident Type

incidentByType = activeOrders.groupby('Incident Type').count()
columns = ['Customer Account Name', 'Customer Contact Name', 'Call Source', 'Status', 'Created Date Time', 'MTTR Description', 'Time To Resolve', 'Priority', 'Owned By Team', 'Last Modified Date Time', 'IncidentDurationInHours']
incidentByType.drop(columns=columns, inplace=True)
incidentByType.rename(columns = {'Incident ID': 'Count of Incident Type'}, inplace=True)
incidentByType.reset_index(inplace=True)

# Email section

smtp_server = 'smtp.gmail.com'
port = 587 
sender_email = config.sender_eml

receiver_email = 'nickferrar@gmail.com'
   
message = f'''From: Automatic Nick <autoreportferrari@gmail.com>
To: Nick Ferrari nickferrar@gmail.com>
Subject: Hourly Service Report

Total Number of Active Orders : {numActive}
Number of Active New Orders in Last Hour: {numNewActive}
Top 5 Longest Active Orders by Incident ID: {longestActive.values}
Highest Total Duration of Incidents by Customer {customerInfo.iat[0,0], customerInfo.iat[0,1]}
Number of Orders On Hold: {numOnHold}
Number of Incidents by Type
--------------------------------
{incidentByType.to_string()}

Number of Incidents by Priority
--------------------------------
{priorityCount.to_string()}

Number of Active Orders by Team
--------------------------------
'''
for team, num in numOrdersByTeam.values:
    message = message + f"{team} {',':^1} {num:>3}\n"

# Create a secure SSL context
context = ssl.create_default_context()

# Try to log in to server and send email
try:
    server = smtplib.SMTP(smtp_server, port)
    server.ehlo() 
    server.starttls(context = context)
    server.ehlo()
    server.login(sender_email, config.password)
    server.sendmail(sender_email, receiver_email, message)
    
except Exception as e:
    print(e)
finally:
    server.quit()

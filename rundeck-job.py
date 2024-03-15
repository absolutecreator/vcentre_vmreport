import requests

print("Added new project. Start new week. Reports will be backed up. Script being restarted.")
headers = {'Content-Type': 'text/xml'}
params = {'authtoken': 'AUTHTOKEN'}
response = requests.post('http://127.0.0.1:4440/api/40/job/job-id/run', params=params, headers=headers)
print(response)

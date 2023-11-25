import requests

# Constants
CLIENT_ID = 'your_client_id'
CLIENT_SECRET = 'your_client_secret'
TENANT_ID = 'your_tenant_id'
AUTHORITY_URL = 'https://login.microsoftonline.com/' + TENANT_ID
RESOURCE = 'https://graph.microsoft.com/'
API_VERSION = 'v1.0'





USER_TOKEN = "EwCAA8l6BAAUAOyDv0l6PcCVu89kmzvqZmkWABkAAaeHSRaZ5LGE6cnTUTkDPt6bC5eB4gYvRwhtahbXa7j5f/8eaVPRpYNvsqqWgRmYAJiOijcAcu6BYTgEE4+SJOai2sNR/n00j4RTxK6aTr6EIOipBAr0FFfIKclH2fKiBSc8dVOAv3BBPFMBIV3TfB4eeKzUqhffTntEHB0qVmaDSW7C/hRc+Ce2ZfvJfE1/6b0GtLvBcS447U5EnLDJFRw+qWpIDAKvVtPsSXltQJSNXbJTjJLlhZnjqcvI8Uo6oLVxCkdidcZU38ZHMdNL7zjZPYAQ+Dtd3bWddbUw4YSvGlkUbgiPfXQp3AmUg0aAXsVonqecLwhGE5jRlobwqUgDZgAACFMHT2sF2WuMUAL0wmFaT4sunPd9jZANY8kh0mRiNjC9ur0weg/P7rZxa7Uofrk7K58llqiaEyMLnyeGOVYvjhNYsvQZNXK8AYmO8/FFCi58m46PaMSnzLEocO9miN7DT0etbSQ31+gV1fb7pXs2mLxu5d0A6fn40A53cMN11WfLcp1DAIo1fSa2JIEV/MwL/x2tgFSCczkZjG/lbVSh/jI54AFglFHJvc/Sltij/ARF23RVE5rAIvdgjAh+JDmjn3MXSBszki+vhHYDGTVURHXAumbGTmnBCDPW0wyJ3jJZAhaJeSGgSEsjyDddHKdX1sAbdhvDc49b21b2aNxRkvlcmjaQhgMgcwDop4thcw+lc+iG4haelOh1d4c0UcR4zTSlRMmVYxPcBHXgldlCITsk0vzfxnGuP3uFX+4tCKkebq0QvXB1n2cWn+qX0RieF+4VsYNtmsHQrrT68ylQwZBv9ouzKnjZS5ZYCiLl0mw88hnzMVNY7fsun9tJQl3wp3Jypl7YM7jp0JIFx3eqUGVXCbFy7lLRSWnJZiJkIxm+zOc1ZvCX1ZvDzrr5a7sk2r6KeYEAKVh+d/dj33r4jpjzXfSwdj0lZ3JTmgjtiOxbwmqhqm5++UdSkqLCy11j4bzIWyrVUoZ/wNMMe1IXOzjMJ+ijHNTGPvSAmnQ5zTrvudSU4Uijjwn8hWcAfZ8C4V5WIbUx56smah1BGkjxkynhJXJ0iyqnOMZuAFgjiF+44KV7FyEj/NvB9SwgbmnlP9oGCxS1jjKA4opCnLWQHQoVig4xbvSgYBuyhwI="


# Function to get an access token
def get_access_token(client_id, client_secret, authority_url, resource):
    token_endpoint = f"{authority_url}/oauth2/v2.0/token"
    payload = {
        'grant_type': 'client_credentials',
        'client_id': client_id,
        'client_secret': client_secret,
        'scope': resource + '.default'
    }
    response = requests.post(token_endpoint, data=payload)
    access_token = response.json().get('access_token')
    return access_token

# Function to interact with Microsoft To Do API
def microsoft_todo_stats():
    access_token = get_access_token(CLIENT_ID, CLIENT_SECRET, AUTHORITY_URL, RESOURCE)

    headers = {
        'Authorization': 'Bearer ' + access_token,
        'Content-Type': 'application/json'
    }

    # Get task groups (lists)
    response = requests.get(RESOURCE + API_VERSION + '/me/todo/lists', headers=headers)
    task_groups = response.json().get('value', [])
    num_task_groups = len(task_groups)

    for group in task_groups:
        list_id = group['id']

        # Get tasks in each group
        response = requests.get(RESOURCE + API_VERSION + f'/me/todo/lists/{list_id}/tasks', headers=headers)
        tasks = response.json().get('value', [])

        # Count completed and open tasks
        completed_tasks = sum(1 for task in tasks if task['status'] == 'completed')
        open_tasks = len(tasks) - completed_tasks

        print(f"List '{group['displayName']}' has {len(tasks)} tasks: {completed_tasks} completed, {open_tasks} open")

    return num_task_groups

# Call the function
num_task_groups = microsoft_todo_stats()
print(f"Total number of task groups: {num_task_groups}")

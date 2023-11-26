---

# Microsoft Graph Python Project

This project demonstrates the use of the Microsoft Graph API in Python to interact with Microsoft To Do lists, manage emails, and access user information. It provides a command-line interface to view and manage tasks, read the inbox, send emails, and retrieve user details.

## Features

- **Task Management**: View task lists and individual tasks in Microsoft To Do.
- **Email Interaction**: List inbox emails, send emails.
- **User Information**: Display details about the logged-in user.

## Requirements

- Python 3.6 or later.
- `aiohttp` library for asynchronous HTTP requests.
- Microsoft Azure AD application with appropriate permissions (Tasks.Read, Tasks.ReadWrite, Mail.Read, Mail.Send, User.Read).

## Setup

1. Clone the repository:
   ```
   git clone [repository-url]
   ```
2. Install dependencies:
   ```
   pip install -r requirements.txt
   ```
3. Configure your Azure AD application details in `config.cfg`:
   ```
   [azure]
   client_id = YOUR_CLIENT_ID
   client_secret = YOUR_CLIENT_SECRET
   tenant_id = YOUR_TENANT_ID
   ```
4. Run the application:
   ```
   python main.py
   ```

## Usage

The application provides a menu-driven interface:

1. **List Task Groups**: Displays the task lists from Microsoft To Do.
2. **List Tasks in a Group**: Enter a list ID to display tasks in that list.
3. **List My Inbox**: Shows the recent emails in your inbox.
4. **Send Mail**: Send an email.
5. **Display Access Token**: Show the current access token.
6. **Make a Graph Call**: Perform a custom call to the Graph API.

## Example Output

```
List: Personal Tasks, Id: XXXXXXXX
List: Work, Id: XXXXXXXX
...
Please choose one of the following options:
0. Exit
1. Display access token
2. List my inbox
3. Send mail
4. Make a Graph call
5. Get Tasks
6. Get Tasks in List
```

## Contributing

Contributions to this project are welcome. Please ensure that your code adheres to the project's style and that all tests pass.

## License

This project is licensed under the MIT License.

---

**Note**: Replace `[repository-url]` with the actual URL of your GitHub repository. The README is a template and should be customized based on the actual structure and features of your project. Ensure that all instructions are accurate and reflect your project's setup and usage.

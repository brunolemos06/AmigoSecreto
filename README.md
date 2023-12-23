# Amigo Secreto Python Script

## Description

This Python script automates the process of organizing a Secret Santa (Amigo Secreto) event by randomly assigning participants to each other and sending out email notifications. The script reads participant information from an Excel file, performs the secret friend assignment, and optionally sends out emails to notify participants of their assigned secret friend.

## Prerequisites

- Python 3.x
- Pandas library
- Dotenv library
- Gmail account (for sending emails)

## Installation

1. **Clone the repository:**

   ```bash
   git clone https://github.com/your-username/amigo-secreto-script.git
    ```
2. **Install the required dependencies:**

   ```bash
   pip install -r requirements.txt
   ```
3. **Create a .env file in the root directory of the project and add the following environment variables:**

   ```bash
    USER=<your-gmail-username>
    TOKEN=<your-gmail-token>
    ```
    - The `USER` variable should be set to your Gmail username (e.g. `
    - The `TOKEN` OAuth2 token can be generated by following the instructions [here](https://developers.google.com/gmail/api/quickstart/python#step_3_set_up_the_sample).

    MAKE SURE TO KEEP THIS FILE PRIVATE AND DO NOT SHARE IT WITH ANYONE!

4. **Create an Excel file named `participants.xlsx` with the following columns:**
    - `Name`: Participant's name
    - `Email`: Participant's email address

    The Excel file should be saved in the root directory of the project.

## Usage

If you want to send out email notifications to participants, run the following command:

```bash
python amigo_secreto.py
```

If you want to just test the script without sending out emails, run the following command:

```bash
python amigo_secreto.py --test
```

## Participants Excel File

You can export the participants Excel file from Google Forms.
The Google Form should have the Name and Email fields, and the responses should be saved in a Google Sheet file and exported as an Excel file.


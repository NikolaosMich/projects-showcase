Here is a general overview of each project, the reason why they were created and the results. 

Auto Backup Control: 
    This project is essentially an automated email handler. It processes emails by marking them as read, moving them to a subfolder, and extracting data from them. The data is then imported into an Excel file, which is created and modified automatically when the script runs.
    The project went live and has saved us around 1.5 hours each day while also reducing human error.

Translator: 
    This project "listens" to the audio coming from your PC’s speakers, transcribes it using OpenAI's Whisper, and translates the text from German to English (or other languages as needed). While the transcription and translation were accurate, it ran slowly since Whisper was running locally and is quite resource-heavy.
    I didn’t get the chance to test it further or try it on a more powerful machine. This script was built to improve internal communication.

Ticket Reminder:
    This project pulls data from the web version of Dynamics 365 Business Central (NAV) using Selenium, processes that data, and sends an HTTP request to a Power Automate workflow. The workflow then emails the relevant user based on the data.
